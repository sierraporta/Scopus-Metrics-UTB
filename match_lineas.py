"""
match_lineas.py
===============
Asocia cada publicación Scopus de docentes UTB con la(s) línea(s) de
investigación del archivo Book_GIUTB.xlsx usando similitud semántica TF-IDF.

Salida: lineas_investigacion.xlsx
  - Una fila por publicación x docente UTB
  - Columnas: DOCENTE, ESCUELA, Año, Tipo, Título, Fuente, DOI,
              Línea 1..3, Grupo 1..3, Clasificación 1..3, Score 1..3

Uso:
    python match_lineas.py

Requiere: pandas, openpyxl, scikit-learn, numpy
"""

import re
import glob
import pathlib
import numpy as np
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────
BASE_DIR         = pathlib.Path(__file__).parent
FACULTY_XLSX     = BASE_DIR / "Base de Datos Scopus 2025.xlsx"
LINEAS_XLSX      = BASE_DIR / "Book_GIUTB.xlsx"
OUT_XLSX         = BASE_DIR / "lineas_investigacion.xlsx"

# Toma el export más reciente (mismo criterio que build_dashboard.py)
def _scopus_date(p):
    m = re.search(r"scopus_export_(\w+ \d+-\d+)", p.name)
    return m.group(1) if m else ""

_candidates = sorted(BASE_DIR.glob("scopus_export_*.csv"), key=_scopus_date)
SCOPUS_CSV = _candidates[-1] if _candidates else None

START_YEAR   = 2020    # publicaciones desde este año
TOP_LINES    = 3       # número de líneas a asignar por publicación
MIN_SCORE    = 0.02    # score mínimo para considerar una coincidencia

# ─── HELPERS ──────────────────────────────────────────────────────────────────
_LOWER_ES = {"de","del","la","el","las","los","y","e","a","con","en","o","por","para","al"}

def title_case(s):
    if not isinstance(s, str): return s
    words = s.strip().split()
    return " ".join(
        w.lower() if i > 0 and w.lower() in _LOWER_ES else w.capitalize()
        for i, w in enumerate(words)
    )

def _extract_authorid(url):
    if not isinstance(url, str): return None
    m = re.search(r"authorId=(\d+)", url)
    return m.group(1) if m else None

def _canon_id(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return None
    s = re.sub(r"[^0-9]", "", str(v))
    return s if s else None

def doc_type_bucket(t):
    if not isinstance(t, str): return "Other"
    t = t.lower()
    if "review"    in t: return "Review"
    if "conference" in t or "proceeding" in t: return "Conference"
    if "article"   in t or "data paper" in t or "short survey" in t: return "Article"
    return "Other"

def clean_text(s):
    """Limpia y normaliza texto para vectorización."""
    if not isinstance(s, str) or str(s).strip().lower() in ("nan", "", "none"):
        return ""
    # Quita caracteres especiales, reduce espacios
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

# ─── 1. CARGAR DOCENTES ───────────────────────────────────────────────────────
print("📂  Cargando docentes de planta...")
fac = pd.read_excel(FACULTY_XLSX, sheet_name="DOCENTES DE PLANTA 2025-2")
fac["_id_url"]   = fac.get("SCOPUS",    pd.Series(dtype=object)).apply(_extract_authorid)
fac["_id_cell"]  = fac.get("ID SCOPUS", pd.Series(dtype=object)).apply(_canon_id)
fac["author_id"] = fac["_id_cell"].where(fac["_id_cell"].notna(), fac["_id_url"])
fac = fac[fac["author_id"].notna()].copy()
fac["author_id"] = fac["author_id"].astype(str).str.strip()
fac = fac.drop_duplicates("author_id")
fac["DOCENTE"] = fac["DOCENTE"].apply(title_case)

faculty_ids  = set(fac["author_id"])
name_map     = fac.set_index("author_id")["DOCENTE"].to_dict()
school_map   = fac.set_index("author_id")["ESCUELA"].to_dict()
print(f"   {len(faculty_ids)} docentes con ID Scopus")

# ─── 2. CARGAR SCOPUS CSV ─────────────────────────────────────────────────────
if SCOPUS_CSV is None:
    raise FileNotFoundError("No se encontró ningún archivo scopus_export_*.csv en la carpeta.")
print(f"📂  Cargando export Scopus: {SCOPUS_CSV.name}")
df = pd.read_csv(SCOPUS_CSV, encoding="utf-8-sig")

# Filtrar por año
df["Year"] = pd.to_numeric(df.get("Year", pd.Series(dtype=float)), errors="coerce")
df = df[df["Year"] >= START_YEAR].copy()

# Normalizar EID
df["EID"] = df.get("EID", pd.Series(dtype=object)).astype(str).str.strip()

# ─── 3. IDENTIFICAR DOCENTES UTB EN CADA PAPER ────────────────────────────────
print("🔍  Identificando autores UTB...")
rows_out = []
for _, paper in df.iterrows():
    # Parsear "Author(s) ID" y "Author full names" en paralelo
    raw_ids   = str(paper.get("Author(s) ID", ""))
    raw_names = str(paper.get("Author full names", ""))
    ids   = [x.strip() for x in raw_ids.split(";")   if x.strip()]
    names = [x.strip() for x in raw_names.split(";") if x.strip()]

    # Extraer IDs numéricos de los nombres con formato "Apellido, Nombre (ID)"
    parsed_ids = []
    for entry in names:
        m = re.search(r"\((\d+)\)", entry)
        parsed_ids.append(m.group(1) if m else None)

    # Usar IDs del campo Author(s) ID como fallback
    combined_ids = []
    for i, pid in enumerate(parsed_ids):
        if pid:
            combined_ids.append(pid)
        elif i < len(ids):
            combined_ids.append(ids[i])
        else:
            combined_ids.append(None)

    # ¿Alguno pertenece a planta UTB?
    utb_authors = [aid for aid in combined_ids if aid and aid in faculty_ids]
    if not utb_authors:
        # Último recurso: buscar en ids directamente
        utb_authors = [aid for aid in ids if aid in faculty_ids]
    if not utb_authors:
        continue

    for aid in utb_authors:
        rows_out.append({
            "EID":          paper["EID"],
            "author_id":    aid,
            "DOCENTE":      name_map.get(aid, aid),
            "ESCUELA":      school_map.get(aid, ""),
            "Year":         int(paper["Year"]),
            "doc_type":     doc_type_bucket(paper.get("Document Type", "")),
            "Title":        str(paper.get("Title",        "") or ""),
            "Source":       str(paper.get("Source title", "") or ""),
            "DOI":          str(paper.get("DOI",          "") or ""),
            "Abstract":     str(paper.get("Abstract",     "") or ""),
            "AuthKw":       str(paper.get("Author Keywords", "") or ""),
            "IdxKw":        str(paper.get("Index Keywords",  "") or ""),
            "Cited_by":     paper.get("Cited by", 0) or 0,
        })

papers_df = pd.DataFrame(rows_out).drop_duplicates(subset=["EID", "author_id"]).reset_index(drop=True)
print(f"   {len(papers_df)} filas papel×docente ({papers_df['EID'].nunique()} papers únicos, "
      f"{papers_df['DOCENTE'].nunique()} docentes)")

# ─── 4. CARGAR LÍNEAS DE INVESTIGACIÓN ────────────────────────────────────────
print("📂  Cargando líneas de investigación...")
li = pd.read_excel(LINEAS_XLSX)
li.columns = ["Grupo", "Clasificacion", "Linea"]
li = li.dropna(subset=["Linea"]).copy()
li["Linea"]         = li["Linea"].astype(str).str.strip()
li["Grupo"]         = li["Grupo"].astype(str).str.strip()
li["Clasificacion"] = li["Clasificacion"].astype(str).str.strip()
print(f"   {len(li)} líneas de investigación · {li['Grupo'].nunique()} grupos")

# ─── 5. DICCIONARIO DE TRADUCCIONES (offline) ─────────────────────────────────
# Líneas en español → palabras clave en inglés para mejorar matching con abstracts
LINE_TRANSLATIONS = {
    # GIPC
    "Gestión de la cadena de suministro":
        "supply chain management logistics procurement distribution operations",
    "Gestión organizacional":
        "organizational management governance strategy performance operations",
    "Productividad y calidad":
        "productivity quality improvement lean six sigma performance",
    # FAPIS
    "Acústica y ultrasonido":
        "acoustics ultrasound vibration sound wave signal",
    "Educación en STEM: Fomento, igualdad y equidad en áreas STEM":
        "STEM education equity gender equality science technology engineering mathematics",
    "Interferometría y metrología óptica":
        "interferometry optical metrology measurement optics laser holography",
    "Procesamiento y análisis de imágenes y señales":
        "image processing signal analysis computer vision pattern recognition digital filtering",
    # EOLITO
    "Eficiencia energética y uso racional de la energía":
        "energy efficiency rational energy use consumption saving building thermal",
    "Energías alternativas":
        "alternative energy renewable solar wind biomass biofuel photovoltaic",
    "Mecánica de fluidos":
        "fluid mechanics hydraulics flow dynamics CFD turbulence",
    # GISAH
    "Control de la contaminación de los recursos (agua, aire y suelo)":
        "pollution control water quality air quality soil contamination environmental remediation",
    "Hidráulica urbana y costera":
        "urban coastal hydraulics water systems flood drainage",
    "Planeación urbana sostenible":
        "sustainable urban planning city infrastructure resilience",
    "Sostenibilidad ambiental aplicada":
        "applied environmental sustainability ecology ecosystem services",
    "Transformación de residuos biomásicos y del agro para aplicaciones ingenieriles":
        "biomass agricultural waste transformation biorefinery engineering bioenergy composites",
    # GRITAS
    "Analítica de Datos y Big Data":
        "data analytics big data machine learning data science visualization",
    "Aseguramiento de la información":
        "information assurance cybersecurity data security cryptography network security",
    "Ingeniería de software":
        "software engineering development architecture agile testing DevOps",
    "Inteligencia artificial":
        "artificial intelligence machine learning deep learning neural networks AI optimization",
    "Internet of everything":
        "internet of things IoT connected devices embedded systems smart sensor",
    "Sistemas distribuidos y redes":
        "distributed systems computer networks cloud computing parallel",
    "Tecnologías del lenguaje humano":
        "natural language processing NLP human language technology speech text",
    "TIC para el desarrollo":
        "ICT information communications technology development eHealth eGov",
    # GAICO
    "Sistemas de comunicaciones electrónicas":
        "electronic communications wireless telecommunications signal MIMO antenna",
    "Automatización y control de procesos industriales":
        "industrial automation process control robotics PLC SCADA manufacturing",
    "Energía eléctrica":
        "electric power energy electrical grid smart power electronics inverter",
    "Sistemas electrónicos":
        "electronic systems embedded circuits microcontrollers sensors FPGA hardware",
    # El Patio
    "Ecosistemas de innovación":
        "innovation ecosystem startup technology transfer entrepreneurship incubation",
    "Emprendimiento e intraemprendimiento (Entrepreneurship)":
        "entrepreneurship intrapreneurship startup business creation venture",
    "Innovación en el territorio":
        "regional innovation territory local development knowledge transfer",
    "Gestión de la innovación empresarial":
        "business innovation management corporate R&D open innovation",
    "Propiedad intelectual":
        "intellectual property patents trademark copyright licensing",
    # GIJUS
    "Debates actuales de\xa0derecho\xa0privado y comercial":
        "private commercial law legal regulation contracts",
    "Derechos humanos, interculturalidad y desarrollo":
        "human rights intercultural development social",
    "Políticas de desarrollo económico":
        "economic development policy growth",
    "Propiedad industrial y derechos de autor\xa0":
        "industrial property copyright intellectual property rights",
    "Educación para la Prevención de Violencia":
        "violence prevention education peace conflict",
    "Derecho Penal y penitenciario":
        "criminal law penitentiary prison justice",
    "Derecho Urbano, patrimonio y gestión urbana":
        "urban law heritage urban management policy",
    "Derecho y administración de Justicia":
        "justice administration law judiciary",
    "Seguridad ciudadana y resolución de conflictos":
        "citizen security conflict resolution peace",
    "Transformación jurídica digital: legal tech y legal design":
        "digital legal transformation legal technology design law",
    # Desarrollo Salud
    "Bienestar laboral":
        "workplace wellbeing occupational health ergonomics work stress",
    "Desarrollo y educación":
        "human development education child psychology",
    "Salud y sociedad":
        "health society public health community medicine epidemiology",
    # Estudios Químicos y Biológicos
    "Bioinformática aplicada":
        "bioinformatics computational biology genomics proteomics sequence",
    "Diseño y síntesis de compuestos con actividad biológica":
        "drug design synthesis biological activity compounds chemistry pharmacology antibacterial",
    "Emergencia de Patógenos microbianos":
        "microbial pathogens emergence infectious disease antibiotic resistance",
    "Microbiología y toxicología ambiental":
        "microbiology environmental toxicology pollutants bacteria fungi",
    "Microorganismos de ambientes extremos":
        "extremophile microorganisms extreme environments thermophile",
    "Origen y emergencia de patógenos microbianos":
        "microbial pathogens origin emergence virulence infectious bacteria",
    # GIMAT
    "Integridad estructural":
        "structural integrity fatigue fracture mechanics failure analysis",
    "Biopolímeros, Biocompuestos y Nanotecnología":
        "biopolymers biocomposites nanotechnology nanomaterials polymer",
    "Simulacion numerica en medios continuos":
        "numerical simulation continuum mechanics finite element FEM FEA",
    "Simulacion y diseño mecanico y mecatrónico de dispositivos para aprovechamiento energético alternativo":
        "mechanical mechatronic design simulation alternative energy harvesting device",
    # Gravitación y Matemática
    "Sistemas dinámicos":
        "dynamical systems chaos nonlinear differential equations stability",
    "Teoría de representación de álgebras de Lie":
        "Lie algebra representation theory abstract algebra mathematics",
    "Métodos estadísticos clásicos, Bayesianos y Ciencia de Datos":
        "statistical methods Bayesian data science statistics probability inference",
    "Clima espacial y rayos cósmicos ":
        "space weather cosmic rays astrophysics solar radiation",
    "Analítica de Datos":
        "data analytics statistics machine learning visualization prediction",
    "Astrofísica":
        "astrophysics astronomy cosmology gravitational waves dark matter",
    "Ciencia de Datos":
        "data science machine learning python statistics big data",
    "Machine Learning e Inteligencia Artificial":
        "machine learning artificial intelligence deep learning neural networks classification",
    # Educación e Innovación Educativa
    "Desarrollo sostenible":
        "sustainable development SDGs environmental sustainability goals",
    "Comunicación y educación\xa0":
        "communication education media pedagogy literacy",
    "Currículo y educación":
        "curriculum education school learning outcomes",
    "Didácticas específicas":
        "teaching methods pedagogy specific didactics STEM classroom",
    "Innovación educativa e inclusión":
        "educational innovation inclusion diversity accessibility learning",
    "Calidad de la educación":
        "education quality assessment accreditation standards",
    "Currículo y pedagogía":
        "curriculum pedagogy teaching learning instructional design",
    "Gerencia Educativa":
        "educational management administration school leadership",
    # Naval y Offshore
    "Diseño y Construcción de Sistemas Navales & Offshore":
        "naval offshore design construction ship systems marine",
    "Energías Sostenibles Offshore":
        "offshore sustainable energy wind marine renewable ocean",
    "Metocean Offshore & Oceanografía":
        "metocean offshore oceanography ocean waves marine environment",
    "Sistemas de Información Geográfica":
        "geographic information systems GIS spatial mapping remote sensing",
    "Materiales & Corrosión":
        "materials corrosion marine offshore protection coating",
    "Reología & Materiales Compuestos":
        "rheology composite materials mechanical properties viscoelastic",
    # IDEEAS
    "Desarrollo empresarial y organizacional":
        "business organizational development management strategy firm performance",
    "Desarrollo humano sostenible":
        "sustainable human development SDGs well-being poverty",
    "Desarrollo regional y local\xa0":
        "regional local development economics territory competitiveness",
    "Desarrollo, Cultura y Ambiente":
        "development culture environment sustainability society",
    "Energía renovable, cambio climático y desarrollo sostenible":
        "renewable energy climate change sustainable development carbon emissions",
    "Gestión contable y financiera":
        "financial accounting management finance reporting audit",
    "Negocios e internacionalización":
        "business internationalization trade FDI global market export",
    # GESH
    "Comunicación en Salud":
        "health communication public health media campaign",
    "Culturas periodísticas y recepción de audiencias":
        "journalism audience media communication culture reception",
    "Desarrollo humano, cambio social y buen vivir":
        "human development social change well-being community",
    "Género, inclusión, igualdad y diversidad":
        "gender inclusion equality diversity feminism social",
    "Memoria Colectiva, Aprendizaje Social y Construcción de Paz":
        "collective memory social learning peacebuilding reconciliation",
    # Centro de Estudios en Negocios
    "Educación empresarial, emprendimiento y aprendizaje experiencial":
        "business education entrepreneurship experiential learning case",
    "Estudios Contables y Financieros":
        "accounting financial studies finance investment audit",
    "Estudios de casos en Administración y Organizaciones":
        "case studies management organizations administration strategy",
    "Estudios en Internacionalización":
        "internationalization studies global trade foreign direct investment multinational",
    "Estudios para la inclusión social y los negocios verdes":
        "social inclusion green business sustainability circular economy",
    "Estudios sectoriales":
        "sectoral studies industry analysis economics market",
    # GIEC
    "I+D+i en diseño (Investigación, Desarrollo e Innovación en Diseño)":
        "design research development innovation R&D product design",
    "Patrimonio material y territorio":
        "material heritage cultural territory urban architecture",
    "Prácticas espaciales y procesos proyectuales contemporáneos":
        "spatial practices design processes contemporary architecture urban",
}

# ─── 5. VECTORIZACIÓN TF-IDF ──────────────────────────────────────────────────
print("🤖  Construyendo modelo TF-IDF bilingüe...")

# Texto de cada paper: título + abstract + keywords
def paper_text(row):
    parts = [row["Title"], row["Abstract"], row["AuthKw"], row["IdxKw"]]
    return " ".join(clean_text(p) for p in parts if clean_text(p))

# Texto de cada línea: español original + traducción inglesa + grupo abreviado
def line_text(row):
    grupo_short = re.sub(r"Grupo de [Ii]nvestigaci[oó]n\s*", "", row["Grupo"]).strip()
    grupo_short = re.sub(r"Grupo de investigaci[oó]n\s*", "", grupo_short).strip()
    es_text   = clean_text(row["Linea"])
    en_text   = clean_text(LINE_TRANSLATIONS.get(row["Linea"], ""))
    grp_text  = clean_text(grupo_short)
    return " ".join(filter(None, [es_text, en_text, grp_text]))

paper_texts = papers_df.apply(paper_text, axis=1).tolist()
line_texts  = li.apply(line_text, axis=1).tolist()

# Vectorizar corpus conjunto (papers + líneas) para un espacio TF-IDF compartido
vectorizer = TfidfVectorizer(
    analyzer="word",
    ngram_range=(1, 2),          # unigramas + bigramas
    min_df=1,
    sublinear_tf=True,
    strip_accents="unicode",
    token_pattern=r"(?u)\b\w+\b",
)
all_texts    = paper_texts + line_texts
tfidf_matrix = vectorizer.fit_transform(all_texts)

paper_vecs = tfidf_matrix[:len(paper_texts)]
line_vecs  = tfidf_matrix[len(paper_texts):]

print(f"   Vocabulario: {len(vectorizer.vocabulary_)} términos")

# ─── 6. CALCULAR SIMILITUD Y ASIGNAR LÍNEAS ───────────────────────────────────
print("📐  Calculando similitudes...")
sim_matrix = cosine_similarity(paper_vecs, line_vecs)  # (n_papers, n_lines)

# Para cada paper tomar los TOP_LINES con score > MIN_SCORE
result_rows = []
for i, (idx, row) in enumerate(papers_df.iterrows()):
    scores = sim_matrix[i]                          # array de longitud n_lines
    top_ix = np.argsort(scores)[::-1][:TOP_LINES]   # índices descendentes

    rec = {
        "DOCENTE":      row["DOCENTE"],
        "ESCUELA":      row["ESCUELA"],
        "Año":          row["Year"],
        "Tipo":         row["doc_type"],
        "Título":       row["Title"],
        "Fuente":       row["Source"],
        "DOI":          row["DOI"] if row["DOI"] not in ("nan","") else "",
        "Citaciones":   int(row["Cited_by"]) if str(row["Cited_by"]) not in ("nan","") else 0,
        "EID":          row["EID"],
    }

    assigned = 0
    for rank, li_idx in enumerate(top_ix, start=1):
        sc = float(scores[li_idx])
        if sc < MIN_SCORE:
            break
        li_row = li.iloc[li_idx]
        rec[f"Línea_{rank}"]           = li_row["Linea"]
        rec[f"Grupo_{rank}"]           = li_row["Grupo"]
        rec[f"Clasificación_{rank}"]   = li_row["Clasificacion"]
        rec[f"Score_{rank}"]           = round(sc, 4)
        assigned += 1

    # Rellenar columnas faltantes
    for rank in range(assigned + 1, TOP_LINES + 1):
        rec[f"Línea_{rank}"]         = ""
        rec[f"Grupo_{rank}"]         = ""
        rec[f"Clasificación_{rank}"] = ""
        rec[f"Score_{rank}"]         = ""

    result_rows.append(rec)

result_df = pd.DataFrame(result_rows).sort_values(["Año", "DOCENTE", "Título"])
print(f"   {len(result_df)} registros procesados")

# ─── 7. EXPORTAR A EXCEL ──────────────────────────────────────────────────────
print(f"💾  Exportando → {OUT_XLSX.name}")

from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter

wb = Workbook()

# ── Hoja 1: Publicaciones con líneas ──────────────────────────────────────────
ws = wb.active
ws.title = "Publicaciones × Líneas"

HEADER_FILL   = PatternFill("solid", fgColor="1C3448")
HEADER_FONT   = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
SUBH_FILLS    = [PatternFill("solid", fgColor="2F4858"),
                 PatternFill("solid", fgColor="33658A"),
                 PatternFill("solid", fgColor="F6AE2D")]
SUBH_FONT     = Font(bold=True, color="FFFFFF", size=10, name="Calibri")
SUBH_FONT_DK  = Font(bold=True, color="1C3448", size=10, name="Calibri")
ALT_FILL      = PatternFill("solid", fgColor="F0F4F9")
CENTER        = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT_WRAP     = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
LEFT          = Alignment(horizontal="left",   vertical="center", wrap_text=False)
THIN          = Side(style="thin",   color="CBD5E1")
THICK         = Side(style="medium", color="86BBD8")
BORDER        = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_R      = Border(left=THIN, right=THICK, top=THIN, bottom=THIN)

# Columnas base y columnas de líneas
base_cols = ["DOCENTE","ESCUELA","Año","Tipo","Título","Fuente","DOI","Citaciones","EID"]
line_cols = []
for r in range(1, TOP_LINES + 1):
    line_cols += [f"Línea_{r}", f"Grupo_{r}", f"Clasificación_{r}", f"Score_{r}"]

all_cols = base_cols + line_cols

# ── Fila 1: grupo de encabezados de sección ───────────────────────────────────
ws.row_dimensions[1].height = 20
ws.row_dimensions[2].height = 20
ws.freeze_panes = "A3"

# Merge base cols en fila 1
for ci, col in enumerate(base_cols, start=1):
    c = ws.cell(1, ci, "")
    c.fill = HEADER_FILL
ws.merge_cells(start_row=1, start_column=1,
               end_row=1,   end_column=len(base_cols))
ws.cell(1, 1).value = "Publicación"
ws.cell(1, 1).font  = HEADER_FONT
ws.cell(1, 1).alignment = CENTER

# Merge por cada bloque de 4 cols de línea
col_offset = len(base_cols) + 1
for r in range(1, TOP_LINES + 1):
    start = col_offset + (r - 1) * 4
    ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=start + 3)
    c = ws.cell(1, start)
    c.value = f"Línea de investigación #{r}"
    fill = SUBH_FILLS[(r - 1) % len(SUBH_FILLS)]
    c.fill = fill
    c.font = SUBH_FONT if r < 3 else Font(bold=True, color="1C3448", size=10, name="Calibri")
    c.alignment = CENTER

# ── Fila 2: nombres de columna ────────────────────────────────────────────────
LABELS = {
    "DOCENTE":"Docente","ESCUELA":"Escuela","Año":"Año","Tipo":"Tipo",
    "Título":"Título","Fuente":"Fuente","DOI":"DOI","Citaciones":"Citas","EID":"EID",
}
for rank in range(1, TOP_LINES + 1):
    LABELS[f"Línea_{rank}"]         = "Línea"
    LABELS[f"Grupo_{rank}"]         = "Grupo"
    LABELS[f"Clasificación_{rank}"] = "Clasif."
    LABELS[f"Score_{rank}"]         = "Score"

for ci, col in enumerate(all_cols, start=1):
    c = ws.cell(2, ci, LABELS.get(col, col))
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.alignment = CENTER
    c.border = BORDER

# ── Datos ─────────────────────────────────────────────────────────────────────
SCORE_FILL_HI  = PatternFill("solid", fgColor="D1FAE5")  # verde claro
SCORE_FILL_MID = PatternFill("solid", fgColor="FEF9C3")  # amarillo claro
SCORE_FILL_LO  = PatternFill("solid", fgColor="FEE2E2")  # rojo claro

for ri, (_, row) in enumerate(result_df.iterrows(), start=3):
    alt = (ri % 2 == 0)
    for ci, col in enumerate(all_cols, start=1):
        val = row.get(col, "")
        c   = ws.cell(ri, ci, val if str(val) != "nan" else "")
        c.border = BORDER
        c.font   = Font(size=10, name="Calibri")

        if alt:
            c.fill = ALT_FILL

        if col == "Título":
            c.alignment = LEFT_WRAP
        elif col in ("DOCENTE", "ESCUELA", "Fuente", "DOI") or col.startswith("Grupo") or col.startswith("Línea"):
            c.alignment = LEFT
        else:
            c.alignment = CENTER

        # Color score
        if col.startswith("Score_") and val != "":
            try:
                sv = float(val)
                if sv >= 0.20:
                    c.fill = SCORE_FILL_HI
                    c.font = Font(bold=True, size=10, name="Calibri", color="065F46")
                elif sv >= 0.10:
                    c.fill = SCORE_FILL_MID
                    c.font = Font(bold=True, size=10, name="Calibri", color="78350F")
                else:
                    c.fill = SCORE_FILL_LO
            except:
                pass

# ── Anchos de columna ─────────────────────────────────────────────────────────
COL_WIDTHS = {
    "DOCENTE":24, "ESCUELA":22, "Año":6, "Tipo":10, "Título":45,
    "Fuente":30, "DOI":22, "Citaciones":8, "EID":16,
}
for rank in range(1, TOP_LINES + 1):
    COL_WIDTHS[f"Línea_{rank}"]         = 32
    COL_WIDTHS[f"Grupo_{rank}"]         = 38
    COL_WIDTHS[f"Clasificación_{rank}"] = 8
    COL_WIDTHS[f"Score_{rank}"]         = 7

for ci, col in enumerate(all_cols, start=1):
    ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS.get(col, 14)

# ── Hoja 2: Resumen por docente ────────────────────────────────────────────────
ws2 = wb.create_sheet("Resumen por Docente")

# Calcular frecuencia de líneas por docente
rows_res = []
for docente, grp in result_df.groupby("DOCENTE"):
    escuela = grp["ESCUELA"].iloc[0]
    total   = grp["EID"].nunique()
    # Contar líneas asignadas (rank 1 = línea primaria)
    line_counts = {}
    for _, r in grp.iterrows():
        l1 = r.get("Línea_1", "")
        g1 = r.get("Grupo_1", "")
        c1 = r.get("Clasificación_1", "")
        if l1 and str(l1) != "nan":
            key = (l1, g1, c1)
            line_counts[key] = line_counts.get(key, 0) + 1
    # Top 3 líneas primarias
    top3 = sorted(line_counts.items(), key=lambda x: -x[1])[:3]
    rec  = {"Docente": docente, "Escuela": escuela, "Total publicaciones": total}
    for ri2, ((li2, gi2, ci2), cnt) in enumerate(top3, start=1):
        rec[f"Línea principal {ri2}"] = li2
        rec[f"Grupo {ri2}"]           = gi2
        rec[f"Clasif. {ri2}"]         = ci2
        rec[f"Freq. {ri2}"]           = cnt
    rows_res.append(rec)

res_df = pd.DataFrame(rows_res).sort_values("Total publicaciones", ascending=False)

# Escribir encabezado
res_cols = (["Docente","Escuela","Total publicaciones"] +
            [f"{x} {r}" for r in range(1,4) for x in ["Línea principal","Grupo","Clasif.","Freq."]])
# Reorganizar en orden fila
res_cols = ["Docente","Escuela","Total publicaciones"]
for r in range(1,4):
    res_cols += [f"Línea principal {r}", f"Grupo {r}", f"Clasif. {r}", f"Freq. {r}"]

ws2.row_dimensions[1].height = 22
for ci, col in enumerate(res_cols, start=1):
    c = ws2.cell(1, ci, col)
    c.fill  = HEADER_FILL
    c.font  = HEADER_FONT
    c.alignment = CENTER
    c.border = BORDER
ws2.freeze_panes = "A2"

for ri, (_, row) in enumerate(res_df.iterrows(), start=2):
    alt = (ri % 2 == 0)
    for ci, col in enumerate(res_cols, start=1):
        val = row.get(col, "")
        c   = ws2.cell(ri, ci, val if str(val) not in ("nan","") else "")
        c.border = BORDER
        c.font   = Font(size=10, name="Calibri")
        c.fill   = ALT_FILL if alt else PatternFill()
        c.alignment = CENTER if col in ("Total publicaciones","Clasif. 1","Clasif. 2","Clasif. 3",
                                         "Freq. 1","Freq. 2","Freq. 3","Escuela") else LEFT

RES_WIDTHS = {"Docente":26,"Escuela":24,"Total publicaciones":10,
              **{f"Línea principal {r}":34 for r in range(1,4)},
              **{f"Grupo {r}":38 for r in range(1,4)},
              **{f"Clasif. {r}":8 for r in range(1,4)},
              **{f"Freq. {r}":8 for r in range(1,4)}}
for ci, col in enumerate(res_cols, start=1):
    ws2.column_dimensions[get_column_letter(ci)].width = RES_WIDTHS.get(col, 14)

# ── Hoja 3: Resumen por línea ─────────────────────────────────────────────────
ws3 = wb.create_sheet("Frecuencia por Línea")

line_summary = {}
for _, row in result_df.iterrows():
    for rank in range(1, TOP_LINES + 1):
        li_name = row.get(f"Línea_{rank}", "")
        gr_name = row.get(f"Grupo_{rank}", "")
        cl_name = row.get(f"Clasificación_{rank}", "")
        if not li_name or str(li_name) == "nan":
            continue
        key = (li_name, gr_name, cl_name)
        if key not in line_summary:
            line_summary[key] = {"papers": set(), "docentes": set()}
        line_summary[key]["papers"].add(row["EID"])
        line_summary[key]["docentes"].add(row["DOCENTE"])

line_sum_rows = []
for (ln, gn, cn), d in sorted(line_summary.items(), key=lambda x: -len(x[1]["papers"])):
    line_sum_rows.append({
        "Línea":         ln,
        "Grupo":         gn,
        "Clasificación": cn,
        "# Papers":      len(d["papers"]),
        "# Docentes":    len(d["docentes"]),
        "Docentes":      ", ".join(sorted(d["docentes"])),
    })

sum_cols = ["Línea","Grupo","Clasificación","# Papers","# Docentes","Docentes"]
ws3.row_dimensions[1].height = 22
for ci, col in enumerate(sum_cols, start=1):
    c = ws3.cell(1, ci, col)
    c.fill = HEADER_FILL; c.font = HEADER_FONT
    c.alignment = CENTER; c.border = BORDER
ws3.freeze_panes = "A2"

for ri, rec in enumerate(line_sum_rows, start=2):
    alt = (ri % 2 == 0)
    for ci, col in enumerate(sum_cols, start=1):
        val = rec.get(col, "")
        c   = ws3.cell(ri, ci, val)
        c.border = BORDER
        c.font   = Font(size=10, name="Calibri")
        c.fill   = ALT_FILL if alt else PatternFill()
        c.alignment = LEFT if col in ("Línea","Grupo","Docentes") else CENTER

SUM_WIDTHS = {"Línea":34,"Grupo":40,"Clasificación":10,"# Papers":10,"# Docentes":10,"Docentes":60}
for ci, col in enumerate(sum_cols, start=1):
    ws3.column_dimensions[get_column_letter(ci)].width = SUM_WIDTHS.get(col, 14)

# ── Guardar ───────────────────────────────────────────────────────────────────
wb.save(OUT_XLSX)
print(f"\n✅  Archivo generado: {OUT_XLSX}")
print(f"   Hoja 1 — {len(result_df)} filas (publicaciones × docente)")
print(f"   Hoja 2 — {len(res_df)} docentes (resumen líneas primarias)")
print(f"   Hoja 3 — {len(line_sum_rows)} líneas con frecuencias")
print()
print("NOTA: El score TF-IDF va de 0 a 1.")
print("  ≥ 0.20 (verde)  → coincidencia fuerte")
print("  0.10–0.19 (amarillo) → coincidencia moderada")
print("  0.05–0.09 (rojo)     → coincidencia débil")
print("  Para mejorar la precisión puedes probar con sentence-transformers")
print("  (modelo multilingüe) instalando: pip install sentence-transformers")
