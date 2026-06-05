from io import BytesIO
from datetime import datetime
import html
import os
import re

import pandas as pd
import streamlit as st

from docx import Document
from docx.shared import RGBColor, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Strategic District Intelligence Brief",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)

BUILT_IN_WORKBOOK = "Texas Top 24 Research.xlsx"

SHEET_STRATEGIC = "Strategic Indicators"
SHEET_SCORECARD = "Master Scorecard"
SHEET_BASIC = "Basic District Info"
SHEET_CONTACTS = "District Contacts"
SHEET_LEADERSHIP = "Leadership and Governance"
SHEET_CSI = "CSI"
SHEET_TSI = "TSI"
DISTRICT_COL = "District Name"


# ============================================================
# CSS
# ============================================================

CUSTOM_CSS = """
<style>
:root {
    --navy:#102a43;
    --blue:#1d4ed8;
    --blue-soft:#eff6ff;
    --green:#166534;
    --green-soft:#f0fdf4;
    --red:#991b1b;
    --red-soft:#fef2f2;
    --amber:#92400e;
    --amber-soft:#fffbeb;
    --slate:#334155;
    --page:#f3f6fa;
    --surface:#ffffff;
    --text:#0f172a;
    --muted:#475569;
    --border:#cbd5e1;
}

.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stHeader"] {
    background: var(--page) !important;
    color: var(--text) !important;
}

html, body, p, li, span, div, label, section, article,
.stMarkdown, .stMarkdown p, .stMarkdown li, .stMarkdown span, .stMarkdown div {
    color: var(--text) !important;
}

.block-container {
    padding-top: 1.1rem;
    padding-bottom: 2.5rem;
    max-width: 1240px;
}

[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid var(--border) !important;
}

[data-testid="stSidebar"] * {
    color: var(--text) !important;
}

h1, h2, h3, h4, h5, h6 {
    color: var(--navy) !important;
}

.main-title {
    font-size: 2rem;
    font-weight: 850;
    color: var(--navy) !important;
    margin-bottom: .15rem;
    letter-spacing:-.02em;
}

.subtitle {
    color: var(--muted) !important;
    font-size: .98rem;
    margin-bottom: 1rem;
}

.section-title {
    color: var(--navy) !important;
    font-size: 1.22rem;
    font-weight: 850;
    margin: .6rem 0 .3rem;
}

.helper-text {
    color: var(--muted) !important;
    font-size: .9rem;
}

.stTabs [data-baseweb="tab-list"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 14px !important;
    padding: .25rem !important;
}

.stTabs [data-baseweb="tab"] {
    color: var(--slate) !important;
    background: transparent !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
}

.stTabs [aria-selected="true"] {
    background: var(--blue-soft) !important;
    color: var(--blue) !important;
}

[data-testid="stMetric"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 16px !important;
    padding: .85rem !important;
    box-shadow: 0 4px 14px rgba(15, 23, 42, .05) !important;
}

.stButton button,
.stDownloadButton button {
    background: var(--blue) !important;
    color: #ffffff !important;
    border: 1px solid var(--blue) !important;
    border-radius: 10px !important;
    font-weight: 750 !important;
}

.stButton button:hover,
.stDownloadButton button:hover {
    background: #1e40af !important;
    border-color: #1e40af !important;
    color: #ffffff !important;
}

[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid var(--border) !important;
    border-radius: 14px !important;
}

.intel-card {
    border:1px solid var(--border);
    border-radius:20px;
    padding:1.05rem;
    margin-bottom:1rem;
    background:var(--surface);
    color: var(--text) !important;
    box-shadow:0 8px 24px rgba(15,23,42,.08);
}

.intel-card h3 {
    color:var(--navy) !important;
    margin-top:0;
    margin-bottom:.25rem;
    font-size:1.35rem;
}

.meta {
    color:var(--muted) !important;
    font-size:.88rem;
    margin-bottom:.6rem;
    line-height:1.45;
}

.summary-grid {
    display:grid;
    grid-template-columns: repeat(3, minmax(0,1fr));
    gap:.75rem;
    margin:.75rem 0 .6rem;
}

.summary-box {
    background:#ffffff;
    border:1px solid var(--border);
    border-radius:14px;
    padding:.78rem;
    box-shadow:0 3px 12px rgba(15,23,42,.04);
}

.summary-label {
    color:var(--muted) !important;
    font-size:.72rem;
    font-weight:850;
    text-transform:uppercase;
    letter-spacing:.05em;
    margin-bottom:.2rem;
}

.summary-value {
    color:var(--text) !important;
    font-size:.92rem;
    line-height:1.35;
}

.brief-box {
    background: #ffffff;
    border:1px solid #bfdbfe;
    border-left:5px solid var(--blue);
    border-radius:16px;
    padding:.9rem;
    margin:.7rem 0 .8rem;
    color: var(--text) !important;
    box-shadow: 0 4px 16px rgba(15, 23, 42, .05);
}

.solution-card {
    background:#ffffff;
    border:1px solid var(--border);
    border-left:5px solid var(--green);
    border-radius:16px;
    padding:.85rem;
    margin:.55rem 0;
    box-shadow:0 4px 16px rgba(15,23,42,.05);
}

.pain-card {
    background:#ffffff;
    border:1px solid var(--border);
    border-left:5px solid var(--amber);
    border-radius:16px;
    padding:.85rem;
    margin:.55rem 0;
    box-shadow:0 4px 16px rgba(15,23,42,.05);
}

.hot-card {
    background:#ffffff;
    border:1px solid var(--border);
    border-left:5px solid var(--red);
    border-radius:16px;
    padding:.85rem;
    margin:.4rem 0;
    box-shadow:0 4px 16px rgba(15,23,42,.05);
}

.badge,
.chip {
    display:inline-block;
    border-radius:999px;
    padding:.25rem .58rem;
    margin:.14rem .16rem .14rem 0;
    font-size:.76rem;
    font-weight:750;
    line-height:1.2;
    border:1px solid transparent;
}

.chip {
    background:#f8fafc;
    color:#334155 !important;
    border-color:#e2e8f0;
}

.tag-math { background:#dbeafe; color:#1e3a8a !important; border-color:#bfdbfe; }
.tag-mtss { background:#ccfbf1; color:#134e4a !important; border-color:#99f6e4; }
.tag-spedell { background:#ede9fe; color:#4c1d95 !important; border-color:#ddd6fe; }
.tag-ccmr { background:#fef3c7; color:#78350f !important; border-color:#fde68a; }
.tag-teacher { background:#f1f5f9; color:#1e293b !important; border-color:#e2e8f0; }
.tag-hqim { background:#e0e7ff; color:#312e81 !important; border-color:#c7d2fe; }
.tag-funding { background:#fef9c3; color:#713f12 !important; border-color:#fde68a; }
.tag-relationship { background:#dcfce7; color:#14532d !important; border-color:#bbf7d0; }
.tag-literacy { background:#fae8ff; color:#701a75 !important; border-color:#f5d0fe; }
.tag-coaching { background:#ecfccb; color:#365314 !important; border-color:#d9f99d; }
.tag-data { background:#e0f2fe; color:#075985 !important; border-color:#bae6fd; }
.tag-default { background:#eef2ff; color:#312e81 !important; border-color:#e0e7ff; }

.priority {
    display:inline-block;
    border-radius:999px;
    padding:.28rem .62rem;
    font-size:.75rem;
    font-weight:850;
    vertical-align:middle;
}

.priority-very-high { background:var(--red-soft); color:var(--red) !important; border:1px solid #fecaca; }
.priority-high { background:var(--green-soft); color:var(--green) !important; border:1px solid #bbf7d0; }
.priority-medium-high { background:var(--amber-soft); color:var(--amber) !important; border:1px solid #fde68a; }
.priority-medium { background:#e0f2fe; color:#075985 !important; border:1px solid #bae6fd; }

.howto-box {
    background:#ffffff;
    color: var(--text) !important;
    border:1px solid var(--border);
    border-radius:18px;
    padding:1rem 1.1rem;
    margin:.75rem 0;
    box-shadow:0 4px 16px rgba(15,23,42,.05);
}

@media (max-width: 800px) {
    .summary-grid {
        grid-template-columns: 1fr;
    }
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def normalize_key(value):
    text = normalize_text(value).lower()
    text = text.replace("\n", " ").replace("\r", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def safe_html(value):
    return html.escape(str(value))


def district_key(value):
    return normalize_text(value).upper().strip()


def export_timestamp():
    return datetime.now().strftime("%B %d, %Y %I:%M %p")


def parse_float(value):
    try:
        text = normalize_text(value).replace(",", "")
        if not text:
            return None
        return float(text)
    except Exception:
        return None


def format_number(value, decimals=2):
    if normalize_text(value) == "":
        return ""
    try:
        return f"{float(value):.{decimals}f}"
    except Exception:
        return normalize_text(value)


def format_enrollment(value):
    if normalize_text(value) == "":
        return ""
    try:
        return f"{int(float(value)):,}"
    except Exception:
        return normalize_text(value)


def format_percent(value, decimals=0):
    if normalize_text(value) == "":
        return ""
    try:
        number = float(value)
        pct = number * 100 if abs(number) <= 1 else number
        if decimals == 0:
            return f"{pct:.0f}%"
        return f"{pct:.{decimals}f}%"
    except Exception:
        return normalize_text(value)


def truncate_text(text, max_chars=230):
    text = normalize_text(text)
    if len(text) <= max_chars:
        return text
    cut = text[:max_chars].rsplit(" ", 1)[0]
    return cut + "..."


def clean_bullet_text(text):
    text = normalize_text(text)
    text = text.replace("&amp;", "&")
    return text


def compact_items(items, limit=6, max_chars=230):
    output = []
    for item in items:
        item = clean_bullet_text(item)
        if not item:
            continue
        output.append(truncate_text(item, max_chars))
        if len(output) >= limit:
            break
    return output


def priority_rank(priority):
    return {
        "Very High": 4,
        "High": 3,
        "Medium-High": 2,
        "Medium": 1,
    }.get(priority, 0)


def score_sort_value(card):
    return parse_float(card.get("score")) or 0


def enrollment_sort_value(card):
    return parse_float(str(card.get("enrollment", "")).replace(",", "")) or 0


def make_unique_columns(columns):
    seen = {}
    output = []
    for col in columns:
        base = normalize_text(col) or "Unnamed"
        if base not in seen:
            seen[base] = 0
            output.append(base)
        else:
            seen[base] += 1
            output.append(f"{base}.{seen[base]}")
    return output


def clean_columns(df):
    if df.empty:
        return df

    df = df.copy()
    df.columns = make_unique_columns(
        [str(c).strip().replace("\n", " ").replace("\r", " ") for c in df.columns]
    )

    drop_cols = [
        c for c in df.columns
        if normalize_key(c).startswith("methods:")
        or normalize_key(c).startswith("unnamed")
    ]

    if drop_cols:
        df = df.drop(columns=drop_cols)

    return df


def find_sheet_name(xls, desired_name):
    lookup = {normalize_key(sheet): sheet for sheet in xls.sheet_names}
    return lookup.get(normalize_key(desired_name))


def read_table_sheet(xls, desired_sheet_name):
    actual_sheet = find_sheet_name(xls, desired_sheet_name)
    if not actual_sheet:
        return pd.DataFrame()

    raw = pd.read_excel(xls, sheet_name=actual_sheet, header=None, engine="openpyxl")
    if raw.empty:
        return pd.DataFrame()

    header_row = None
    for idx, row in raw.iterrows():
        values = [normalize_text(v) for v in row.tolist()]
        if DISTRICT_COL in values:
            header_row = idx
            break

    if header_row is None:
        return pd.DataFrame()

    headers = [normalize_text(v) for v in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1:].copy()
    data.columns = headers
    data = clean_columns(data)
    data = data.dropna(how="all")

    if DISTRICT_COL in data.columns:
        data = data[data[DISTRICT_COL].notna()]
        data = data[data[DISTRICT_COL].astype(str).str.strip() != ""]

    return data.reset_index(drop=True)


def find_col(df, exact_names=None, contains_all=None):
    if df.empty:
        return None

    exact_names = exact_names or []
    contains_all = contains_all or []
    normalized_map = {normalize_key(col): col for col in df.columns}

    for name in exact_names:
        key = normalize_key(name)
        if key in normalized_map:
            return normalized_map[key]

    if contains_all:
        for col in df.columns:
            col_key = normalize_key(col)
            if all(term.lower() in col_key for term in contains_all):
                return col

    return None


def get_value(row_or_dict, col_name, default=""):
    if row_or_dict is None:
        return default

    if isinstance(row_or_dict, dict):
        for key_name, value in row_or_dict.items():
            if normalize_key(key_name) == normalize_key(col_name):
                return value
        return default

    try:
        for key_name in row_or_dict.index:
            if normalize_key(key_name) == normalize_key(col_name):
                return row_or_dict.get(key_name, default)
    except Exception:
        return default

    return default


# ============================================================
# VISUAL TAGS
# ============================================================

def tag_class(tag):
    mapping = {
        "Math": "tag-math",
        "MTSS": "tag-mtss",
        "SPED/ELL": "tag-spedell",
        "CCMR": "tag-ccmr",
        "Teacher Capacity": "tag-teacher",
        "Curriculum / HQIM": "tag-hqim",
        "Funding / Grants": "tag-funding",
        "Existing Relationship": "tag-relationship",
        "Literacy / Dyslexia": "tag-literacy",
        "Coaching / Leadership": "tag-coaching",
        "Data Systems": "tag-data",
    }
    return mapping.get(tag, "tag-default")


def badge_html(tag):
    return f'<span class="badge {tag_class(tag)}">{safe_html(tag)}</span>'


def chip_html(item):
    return f'<span class="chip">{safe_html(item)}</span>'


# ============================================================
# PCG SERVICE LIBRARY
# ============================================================

PCG_SERVICE_LIBRARY = [
    {
        "family": "Instructional Math Resources",
        "offering": "Elevation Station Math Games",
        "pain_patterns": ["Math Fluency / Practice Gap", "Intervention System Inconsistency"],
        "tags": ["Math", "MTSS"],
        "keywords": [
            "math fluency", "numeracy", "staAR math", "algebra readiness", "practice",
            "small group", "tier 2", "intervention", "summer learning", "tutoring",
            "conceptual understanding", "procedural fluency"
        ],
        "buyer_personas": ["Math Lead", "Curriculum Leader", "Intervention Leader", "Chief Academic Officer"],
        "bd_motion": "Product-Supported Pilot Motion",
        "positioning": "Use as a classroom-facing math practice and fluency layer that complements core instruction and supports Tier 1, Tier 2, enrichment, tutoring, and extended learning.",
        "do_not_lead_with": "Do not frame as a replacement curriculum. Lead with practice, engagement, fluency, and ease of implementation.",
        "first_move": "Validate where students need more frequent math practice and where teachers need ready-to-use small-group or station-based supports.",
        "discovery": [
            "Where are students getting frequent opportunities to practice and explain mathematical thinking?",
            "How are teachers currently supporting fluency without relying only on rote practice?",
            "Where could small-group math resources reduce teacher planning burden?"
        ],
    },
    {
        "family": "K–5 Math Intervention",
        "offering": "Elevation Lessons",
        "pain_patterns": ["Math Fluency / Practice Gap", "Intervention System Inconsistency", "Instructional Execution Gap"],
        "tags": ["Math", "MTSS", "Curriculum / HQIM", "Teacher Capacity"],
        "keywords": [
            "k-5 math", "intervention", "remediation", "small group", "enrichment",
            "formative assessment", "problem solving", "reasoning", "discourse", "summer school"
        ],
        "buyer_personas": ["Elementary Leader", "Math Lead", "Intervention Leader", "Chief Academic Officer"],
        "bd_motion": "Targeted Instructional Support Motion",
        "positioning": "Position as K–5 supplemental lessons that give teachers ready-to-use routines for problem-solving, discourse, targeted intervention, and formative assessment.",
        "do_not_lead_with": "Do not lead with a product catalog. Lead with specific K–5 instructional gaps and teacher usability.",
        "first_move": "Ask where K–5 foundational math gaps show up most clearly and what teachers currently use for targeted reteaching.",
        "discovery": [
            "Where are K–5 students showing unfinished learning in foundational mathematics?",
            "How are teachers selecting lessons for small-group support?",
            "What evidence tells teachers when students are ready to move forward?"
        ],
    },
    {
        "family": "Career-Connected Math / CCMR",
        "offering": "RISE Career & Math Mini Lessons",
        "pain_patterns": ["Career Relevance / Readiness Gap", "Math Engagement Gap"],
        "tags": ["Math", "CCMR"],
        "keywords": [
            "career readiness", "ccmr", "career pathways", "cte", "industry-based certification",
            "real-world math", "applied math", "middle grades", "pathways"
        ],
        "buyer_personas": ["CCMR Leader", "CTE Leader", "Math Lead", "Middle Grades Leader"],
        "bd_motion": "Engagement / Readiness Motion",
        "positioning": "Use when a district wants students to connect grade-level math to career pathways, readiness, and real-world application.",
        "do_not_lead_with": "Do not present as a full CTE program. Lead with middle-grade relevance and applied math.",
        "first_move": "Validate whether the district is trying to make math feel more relevant to future pathways.",
        "discovery": [
            "Where do students first connect academic skills to future pathways?",
            "How are middle-grade students exposed to careers that use mathematics?",
            "What would make career-connected math useful without disrupting core pacing?"
        ],
    },
    {
        "family": "Professional Learning Infrastructure",
        "offering": "PCG Playbook Professional Learning Platform",
        "pain_patterns": ["Professional Learning Translation Gap", "Coaching Infrastructure Weakness", "Implementation Monitoring Gap"],
        "tags": ["Teacher Capacity", "Coaching / Leadership", "Data Systems", "SPED/ELL", "Curriculum / HQIM", "Literacy / Dyslexia"],
        "keywords": [
            "professional learning", "asynchronous", "coaching", "learning paths", "onboarding",
            "certification", "in-service", "pd management", "playbook", "teacher capacity",
            "dyslexia playbook", "special education playbook", "leadership development"
        ],
        "buyer_personas": ["Professional Learning Leader", "Chief Academic Officer", "HR / Talent Leader", "Special Education Leader", "IT / Data Leader"],
        "bd_motion": "Platform / Scale Motion",
        "positioning": "Position Playbook as professional learning infrastructure: role-based learning paths, coaching tools, certificates, PD management, reporting, and scalable implementation support.",
        "do_not_lead_with": "Do not lead with software features. Lead with the district’s need to scale, monitor, and sustain adult learning.",
        "first_move": "Ask how the district tracks whether professional learning is changing educator practice.",
        "discovery": [
            "How are professional learning, coaching, and completion data currently tracked?",
            "Where does the district need more role-specific or just-in-time support?",
            "What professional learning needs to scale without losing consistency?"
        ],
    },
    {
        "family": "Math Professional Learning",
        "offering": "Asynchronous Math Professional Learning + PLC Implementation Support",
        "pain_patterns": ["Professional Learning Translation Gap", "Instructional Execution Gap", "Math Fluency / Practice Gap"],
        "tags": ["Math", "Teacher Capacity", "Curriculum / HQIM", "Data Systems"],
        "keywords": [
            "math professional learning", "math pd", "plc", "diagnostics", "assessment",
            "survey", "implementation support", "middle grades math", "teacher training"
        ],
        "buyer_personas": ["Math Lead", "Professional Learning Leader", "Curriculum Leader", "Chief Academic Officer"],
        "bd_motion": "Adult Learning / Implementation Motion",
        "positioning": "Lead with teacher capacity, math PLCs, diagnostics, assessments, and sustained implementation routines; use Playbook when scale and tracking matter.",
        "do_not_lead_with": "Do not jump straight to student resources if the district’s problem is adult practice. Lead with teacher learning and PLC implementation.",
        "first_move": "Ask what math professional learning has changed in classroom practice and where support is still inconsistent.",
        "discovery": [
            "Where do math PLCs need clearer tools or common routines?",
            "How do leaders know whether math PD is changing classroom practice?",
            "Would district-owned math PL content support long-term sustainability?"
        ],
    },
    {
        "family": "MTSS / Intervention Systems",
        "offering": "PCG MTSS Consulting and Intervention Workflow Design",
        "pain_patterns": ["Intervention System Inconsistency", "Implementation Monitoring Gap", "Subgroup Access / Service Delivery Pressure"],
        "tags": ["MTSS", "Math", "SPED/ELL", "Data Systems"],
        "keywords": [
            "mtss", "tier 2", "tier 3", "intervention", "progress monitoring",
            "data cycles", "campus variation", "response to intervention", "rti",
            "early warning", "behavior supports"
        ],
        "buyer_personas": ["Chief Academic Officer", "Student Support Leader", "MTSS Leader", "Intervention Leader", "Special Education Leader"],
        "bd_motion": "Diagnostic / Systems Design Motion",
        "positioning": "Lead with intervention consistency, student identification, support assignment, progress monitoring, and data-to-action routines.",
        "do_not_lead_with": "Do not lead with a tool before understanding the district’s intervention workflow and MTSS maturity.",
        "first_move": "Offer to map the intervention workflow and identify where support slows down or varies across campuses.",
        "discovery": [
            "What slows down the response after students are identified for support?",
            "Where do Tier 2 and Tier 3 routines vary most across campuses?",
            "How do leaders know interventions are being implemented with fidelity?"
        ],
    },
    {
        "family": "Special Education Systems Improvement",
        "offering": "SPED Systems Redesign, Process Mapping, Compliance Visibility, and Implementation Coherence",
        "pain_patterns": ["Subgroup Access / Service Delivery Pressure", "Compliance / Visibility Risk", "Implementation Monitoring Gap"],
        "tags": ["SPED/ELL", "Teacher Capacity", "Data Systems", "Coaching / Leadership"],
        "keywords": [
            "special education", "sped", "exceptional student services", "compliance",
            "dispute resolution", "iep", "evaluation", "eligibility", "transition",
            "process mapping", "raci", "dashboard", "service delivery", "inclusive practices"
        ],
        "buyer_personas": ["Special Education Leader", "Student Support Leader", "Chief Academic Officer", "Compliance Leader", "Superintendent"],
        "bd_motion": "Systems Redesign / Diagnostic Motion",
        "positioning": "Frame PCG as a partner for moving SPED from disconnected compliance activity to coherent, accountable, student-centered implementation with process clarity, ownership, data visibility, and professional learning.",
        "do_not_lead_with": "Do not lead with compliance fear. Lead with access, clarity, service delivery, visibility, and proactive support.",
        "first_move": "Ask whether leaders can see where SPED processes are working and where workflow, ownership, or monitoring is breaking down.",
        "discovery": [
            "Where do special education workflows break down across schools?",
            "How are service delivery and compliance patterns monitored in real time?",
            "How clear are ownership and decision rights across central office and campuses?"
        ],
    },
    {
        "family": "Literacy / Dyslexia / Review Services",
        "offering": "Structured Literacy and Dyslexia Program Review",
        "pain_patterns": ["Literacy / Dyslexia Implementation Risk", "Instructional Execution Gap", "Compliance / Visibility Risk"],
        "tags": ["Literacy / Dyslexia", "SPED/ELL", "Teacher Capacity", "Curriculum / HQIM", "Data Systems"],
        "keywords": [
            "literacy", "reading", "science of reading", "structured literacy", "dyslexia",
            "reading disabilities", "foundational skills", "curriculum audit", "literacy audit",
            "evidence-aligned", "phonics", "fluency", "comprehension"
        ],
        "buyer_personas": ["Literacy Leader", "Chief Academic Officer", "Curriculum Leader", "Special Education Leader", "Professional Learning Leader"],
        "bd_motion": "Review / Implementation Coherence Motion",
        "positioning": "Position PCG around evidence-aligned literacy implementation, dyslexia identification/intervention, audit/review tools, professional learning, and actionable recommendations.",
        "do_not_lead_with": "Do not imply the district lacks literacy knowledge. Lead with implementation evidence, consistency, and support.",
        "first_move": "Ask whether a structured literacy or dyslexia implementation review would help clarify highest-leverage next steps.",
        "discovery": [
            "Where are literacy or dyslexia practices strongest today?",
            "Where does implementation still vary across campuses?",
            "What evidence helps leaders know if reading intervention is aligned and effective?"
        ],
    },
    {
        "family": "Coaching and Leadership Systems",
        "offering": "Coaching Infrastructure, Observation Tools, Feedback Routines, and Leadership Development",
        "pain_patterns": ["Coaching Infrastructure Weakness", "Professional Learning Translation Gap", "Instructional Execution Gap"],
        "tags": ["Coaching / Leadership", "Teacher Capacity", "Curriculum / HQIM", "Data Systems"],
        "keywords": [
            "coaching", "instructional coach", "coaching coordinator", "leadership",
            "observation", "feedback", "walkthrough", "coach selection", "onboarding",
            "rubric", "look-fors", "principal", "leadership pipeline"
        ],
        "buyer_personas": ["Professional Learning Leader", "Coaching Leader", "Chief Academic Officer", "School Improvement Leader", "HR / Talent Leader"],
        "bd_motion": "Coaching Infrastructure Motion",
        "positioning": "Position PCG as a partner that helps define effective coaching, build observation and feedback tools, support coach selection/onboarding, and create data-informed improvement routines.",
        "do_not_lead_with": "Do not use generic leadership-development language. Lead with coaching quality, implementation monitoring, and usable tools.",
        "first_move": "Ask how leaders currently know whether coaching is improving classroom practice.",
        "discovery": [
            "How does the district define effective coaching practice?",
            "What tools help leaders observe coaching quality and provide actionable feedback?",
            "How are coaches selected, onboarded, and supported?"
        ],
    },
    {
        "family": "Curriculum / HQIM Implementation",
        "offering": "Curriculum Implementation, PLC Routines, Walkthroughs, and Change Management",
        "pain_patterns": ["Curriculum / HQIM Implementation Gap", "Instructional Execution Gap", "Professional Learning Translation Gap"],
        "tags": ["Curriculum / HQIM", "Teacher Capacity", "Coaching / Leadership", "Data Systems"],
        "keywords": [
            "curriculum implementation", "hqim", "instructional materials", "adoption",
            "fidelity", "bluebonnet", "eureka", "carnegie", "curriculum audit",
            "plc", "instructional coherence", "walkthrough"
        ],
        "buyer_personas": ["Chief Academic Officer", "Curriculum Leader", "Professional Learning Leader", "School Improvement Leader"],
        "bd_motion": "Implementation Support Motion",
        "positioning": "Frame PCG as helping leaders make adopted materials usable, coherent, and consistently enacted across classrooms.",
        "do_not_lead_with": "Do not lead as a competing curriculum. Lead with implementation fidelity, teacher usability, PLC routines, and monitoring.",
        "first_move": "Ask where implementation is strongest and where variability remains.",
        "discovery": [
            "Where is implementation strongest and where does variability remain?",
            "What support do teachers need after initial training?",
            "How are leaders monitoring curriculum use and instructional quality?"
        ],
    },
    {
        "family": "Financial / Federal Programs",
        "offering": "Funding Alignment, Medicaid Revenue, and Program Administration Support",
        "pain_patterns": ["Funding / Sustainability Constraint", "Compliance / Visibility Risk"],
        "tags": ["Funding / Grants", "SPED/ELL", "Data Systems"],
        "keywords": [
            "funding", "grant", "title i", "federal programs", "medicaid",
            "revenue", "reimbursement", "compliance", "audit", "budget", "fiscal"
        ],
        "buyer_personas": ["Chief Financial Officer", "Federal Programs Leader", "Special Education Leader", "Operations Leader"],
        "bd_motion": "Funding Alignment Motion",
        "positioning": "Position PCG as a partner that helps align strategic priorities to sustainable funding, compliance, Medicaid reimbursement, and federal program operations.",
        "do_not_lead_with": "Do not lead with funding mechanics unless the district has already surfaced sustainability or compliance concerns.",
        "first_move": "Ask which strategic priorities have funding attached and which are difficult to sustain.",
        "discovery": [
            "Which priorities already have aligned funding and which do not?",
            "Where are documentation or compliance burdens slowing implementation?",
            "Are there reimbursement or federal program opportunities that could sustain this work?"
        ],
    },
]


# ============================================================
# PAIN PATTERN LIBRARY
# ============================================================

PAIN_PATTERN_LIBRARY = [
    {
        "name": "Instructional Execution Gap",
        "description": "The district has strategic priorities, but the likely challenge is translating those priorities into consistent classroom practice.",
        "tags": ["Teacher Capacity", "Curriculum / HQIM", "Coaching / Leadership", "Math", "Literacy / Dyslexia"],
        "keywords": ["implementation", "instructional consistency", "teacher capacity", "professional learning", "coaching", "classroom", "rigor", "plc", "walkthrough"],
        "bd_motion": "Implementation Support Motion",
        "best_first_conversation": "Talk about where district priorities are not yet showing up consistently in classrooms.",
        "leadership_angle": "Coherence, implementation fidelity, teacher usability, and measurable classroom change.",
    },
    {
        "name": "Intervention System Inconsistency",
        "description": "The district likely needs stronger routines for identifying students, assigning supports, monitoring progress, and adjusting intervention.",
        "tags": ["MTSS", "Math", "SPED/ELL", "Data Systems"],
        "keywords": ["mtss", "tier 2", "tier 3", "intervention", "progress monitoring", "data cycles", "rti", "early warning"],
        "bd_motion": "Diagnostic / Systems Design Motion",
        "best_first_conversation": "Talk about where the intervention workflow slows down or varies across campuses.",
        "leadership_angle": "Student identification, support assignment, fidelity, and data-to-action routines.",
    },
    {
        "name": "Subgroup Access / Service Delivery Pressure",
        "description": "The district shows pressure around SPED, emergent bilingual, multilingual learner, or other subgroup access and outcomes.",
        "tags": ["SPED/ELL", "MTSS", "Data Systems"],
        "keywords": ["special education", "sped", "english learner", "emergent bilingual", "multilingual", "students with disabilities", "access", "service delivery", "subgroup"],
        "bd_motion": "Systems Redesign / Support Motion",
        "best_first_conversation": "Talk about how consistently priority student groups access grade-level expectations and services across campuses.",
        "leadership_angle": "Access, service delivery, inclusive practice, subgroup progress monitoring, and implementation visibility.",
    },
    {
        "name": "Professional Learning Translation Gap",
        "description": "The district appears to be investing in professional learning, but the question is whether PD changes educator practice.",
        "tags": ["Teacher Capacity", "Coaching / Leadership", "Curriculum / HQIM", "Data Systems"],
        "keywords": ["professional learning", "pd", "training", "teacher capacity", "coaching", "onboarding", "learning path", "playbook", "certification"],
        "bd_motion": "Professional Learning / Platform Motion",
        "best_first_conversation": "Talk about how leaders know whether professional learning is translating into classroom execution.",
        "leadership_angle": "Scalable PL, role-based support, coaching, certificates, progress reporting, and sustainability.",
    },
    {
        "name": "Coaching Infrastructure Weakness",
        "description": "The district may have coaches or leadership routines, but lacks consistent tools for observing coaching quality and supporting implementation.",
        "tags": ["Coaching / Leadership", "Teacher Capacity", "Data Systems"],
        "keywords": ["coaching", "instructional coach", "walkthrough", "observation", "feedback", "coach selection", "onboarding", "leadership pipeline"],
        "bd_motion": "Coaching Infrastructure Motion",
        "best_first_conversation": "Talk about how the district defines effective coaching and monitors coaching quality.",
        "leadership_angle": "Observation tools, feedback routines, coach onboarding, coordinator support, and implementation monitoring.",
    },
    {
        "name": "Literacy / Dyslexia Implementation Risk",
        "description": "The district has literacy, reading, structured literacy, or dyslexia priorities that may require evidence review, implementation support, or professional learning.",
        "tags": ["Literacy / Dyslexia", "SPED/ELL", "Teacher Capacity", "Curriculum / HQIM"],
        "keywords": ["literacy", "reading", "science of reading", "structured literacy", "dyslexia", "foundational skills", "phonics", "fluency", "reading disabilities"],
        "bd_motion": "Review / Implementation Coherence Motion",
        "best_first_conversation": "Talk about where literacy or dyslexia implementation is strong and where it still varies.",
        "leadership_angle": "Evidence-aligned instruction, dyslexia identification/intervention, implementation coherence, and professional learning.",
    },
    {
        "name": "Curriculum / HQIM Implementation Gap",
        "description": "The district is adopting or implementing instructional materials, but needs support turning materials into consistent instruction.",
        "tags": ["Curriculum / HQIM", "Teacher Capacity", "Coaching / Leadership"],
        "keywords": ["curriculum", "hqim", "instructional materials", "adoption", "fidelity", "bluebonnet", "eureka", "carnegie", "implementation"],
        "bd_motion": "Implementation Support Motion",
        "best_first_conversation": "Talk about where adopted materials are not yet being used consistently or effectively.",
        "leadership_angle": "Teacher usability, PLC routines, coaching, walkthroughs, and implementation fidelity.",
    },
    {
        "name": "Career Relevance / Readiness Gap",
        "description": "The district is trying to improve CCMR, pathways, or student engagement by connecting academics to future opportunities.",
        "tags": ["CCMR", "Math"],
        "keywords": ["ccmr", "career", "pathways", "cte", "industry-based certification", "dual credit", "tsia", "college", "readiness"],
        "bd_motion": "Readiness / Relevance Motion",
        "best_first_conversation": "Talk about where students begin connecting academic skills to future pathways.",
        "leadership_angle": "Career-connected learning, pathway awareness, applied math, and middle-grade readiness.",
    },
    {
        "name": "Funding / Sustainability Constraint",
        "description": "The district may have strong priorities but needs help sustaining implementation through funding alignment or operational support.",
        "tags": ["Funding / Grants"],
        "keywords": ["funding", "grant", "title", "budget", "fiscal", "sustainability", "medicaid", "reimbursement"],
        "bd_motion": "Funding Alignment Motion",
        "best_first_conversation": "Talk about which priorities have funding attached and which are hard to sustain.",
        "leadership_angle": "Title, grants, Medicaid, federal programs, compliance, and sustainability.",
    },
]


# ============================================================
# BUYER PERSONA LOGIC
# ============================================================

def infer_persona(contact_text):
    t = normalize_key(contact_text)

    if any(x in t for x in ["superintendent", "chief of schools", "deputy superintendent"]):
        return "Superintendent / Executive Leader"

    if any(x in t for x in ["chief academic", "academics", "curriculum", "instruction"]):
        return "Chief Academic / Curriculum Leader"

    if any(x in t for x in ["math", "mathematics", "stem"]):
        return "Math / STEM Leader"

    if any(x in t for x in ["literacy", "reading", "ela", "dyslexia"]):
        return "Literacy / Dyslexia Leader"

    if any(x in t for x in ["special education", "sped", "exceptional", "student support", "pupil"]):
        return "Special Education / Student Support Leader"

    if any(x in t for x in ["career", "cte", "ccmr", "college"]):
        return "CCMR / CTE Leader"

    if any(x in t for x in ["professional learning", "talent", "human resources", "hr", "teacher development"]):
        return "Professional Learning / Talent Leader"

    if any(x in t for x in ["finance", "financial", "federal", "grant", "title", "budget", "medicaid"]):
        return "Finance / Federal Programs Leader"

    if any(x in t for x in ["technology", "data", "assessment", "analytics", "information"]):
        return "Data / Technology / Assessment Leader"

    if any(x in t for x in ["principal", "school improvement", "leadership"]):
        return "School Improvement / Leadership Leader"

    return "General District Influencer"


def persona_angle(persona, card):
    top_pains = card.get("pain_patterns", [])
    primary_pain = top_pains[0]["name"] if top_pains else "district implementation priorities"
    top_solution = card.get("solution_pathways", [{}])[0].get("offering", "PCG support") if card.get("solution_pathways") else "PCG support"

    mapping = {
        "Superintendent / Executive Leader": f"Focus on strategic coherence, measurable outcomes, and how PCG can help solve {primary_pain.lower()} through a focused first move.",
        "Chief Academic / Curriculum Leader": f"Focus on instruction, implementation consistency, curriculum use, and validating whether {top_solution} fits the district’s academic priorities.",
        "Math / STEM Leader": "Focus on math growth, fluency, small-group support, intervention, PLC routines, and teacher-ready math resources.",
        "Literacy / Dyslexia Leader": "Focus on structured literacy, dyslexia identification/intervention, reading implementation coherence, and evidence-aligned professional learning.",
        "Special Education / Student Support Leader": "Focus on access, service delivery, compliance visibility, inclusive practice, and progress monitoring.",
        "CCMR / CTE Leader": "Focus on applied readiness, pathways, career-connected math, and middle-grade relevance.",
        "Professional Learning / Talent Leader": "Focus on scalable professional learning, coaching, onboarding, Playbook, completion tracking, and evidence of classroom transfer.",
        "Finance / Federal Programs Leader": "Focus on sustainability, funding alignment, Medicaid/revenue opportunities, and compliance support.",
        "Data / Technology / Assessment Leader": "Focus on dashboards, progress monitoring, data-to-action routines, integration, and reporting.",
        "School Improvement / Leadership Leader": "Focus on walkthroughs, coaching quality, leadership routines, campus variation, and implementation monitoring.",
    }

    return mapping.get(persona, f"Focus on {primary_pain.lower()} and validate whether {top_solution} is worth exploring.")


# ============================================================
# EXCEL COLUMN DETECTION / LOOKUPS
# ============================================================

def detect_scorecard_columns(scorecard_df):
    return {
        "district": find_col(scorecard_df, exact_names=["District Name"]),
        "enrollment": find_col(scorecard_df, exact_names=["Enrollment"]),
        "strategic_score": find_col(
            scorecard_df,
            exact_names=["Strategic Weighted Score"],
            contains_all=["strategic", "weighted", "score"],
        ),
        "overall_score": find_col(
            scorecard_df,
            exact_names=[
                "Overall Weighted Score (Strategic+Contract+Relationship)",
                "Overall Weighted Score (Strategic+Relationship Weighted)",
                "Overall Weighted Score",
            ],
            contains_all=["overall", "weighted", "score"],
        ),
        "tier": find_col(scorecard_df, exact_names=["Tier"]),
        "existing_contracts": find_col(
            scorecard_df,
            exact_names=["Existing Contracts in Region (Yes/No)"],
            contains_all=["existing", "contracts"],
        ),
        "existing_relationships": find_col(
            scorecard_df,
            exact_names=["Existing Relationships"],
            contains_all=["existing", "relationships"],
        ),
    }


def is_substantive_strategic_row(row):
    fields_to_check = [
        "Strategic Plan Themes",
        "Math Improvement Mentioned",
        "Math Priority Strength",
        "Intervention Focus",
        "Intervention Focus Details",
        "Teacher Capacity/PD Focus",
        "Teacher Capacity Details",
        "Career Readiness Mentioned",
        "Career Readiness Details",
        "SPED/ELL Improvement Mentioned",
        "SPED/ELL Details",
        "MTSS/Tiered Support Mentioned",
        "MTSS Details",
        "Curriculum Review/Adoption Activity",
        "Curriculum Details",
        "Active Grants (Yes/No)",
        "Grants Details",
        "Sources",
        "Notes",
    ]
    return sum(1 for field in fields_to_check if normalize_text(get_value(row, field))) >= 2


def build_score_lookup(scorecard_df, score_cols):
    lookup = {}
    district_col = score_cols.get("district")

    if scorecard_df.empty or not district_col:
        return lookup

    for _, row in scorecard_df.iterrows():
        district = normalize_text(row.get(district_col, ""))
        if district:
            lookup[district_key(district)] = row.to_dict()

    return lookup


def build_basic_lookup(basic_df):
    lookup = {}

    if basic_df.empty or DISTRICT_COL not in basic_df.columns:
        return lookup

    for _, row in basic_df.iterrows():
        district = normalize_text(row.get(DISTRICT_COL, ""))
        if district:
            lookup[district_key(district)] = row.to_dict()

    return lookup


def build_csi_tsi_counts(csi_df, tsi_df):
    csi_counts = {}
    tsi_counts = {}

    if not csi_df.empty and DISTRICT_COL in csi_df.columns:
        csi_counts = (
            csi_df.groupby(csi_df[DISTRICT_COL].astype(str).str.upper().str.strip())
            .size()
            .to_dict()
        )

    if not tsi_df.empty and DISTRICT_COL in tsi_df.columns:
        tsi_counts = (
            tsi_df.groupby(tsi_df[DISTRICT_COL].astype(str).str.upper().str.strip())
            .size()
            .to_dict()
        )

    return csi_counts, tsi_counts


def determine_priority(tier, score):
    score_num = parse_float(score)

    if tier == "Tier 1":
        return "Very High"
    if tier == "Tier 2":
        return "High"
    if score_num is not None and score_num >= 3.5:
        return "Medium-High"
    return "Medium"


# ============================================================
# CONTACTS
# ============================================================

def format_contact_display(name="", title="", email="", position=""):
    name = normalize_text(name)
    title = normalize_text(title) or normalize_text(position)
    email = normalize_text(email)

    if name and title and email:
        return f"{name} — {title} | {email}"
    if name and title:
        return f"{name} — {title}"
    if name and email:
        return f"{name} | {email}"
    if title and email:
        return f"{title} | {email}"

    return name or title or email


def build_contacts_lookup_from_contacts(contacts_df):
    lookup = {}

    if contacts_df.empty or DISTRICT_COL not in contacts_df.columns:
        return lookup

    role_priority = [
        "SUPERINTENDENT",
        "CHIEF ACADEMIC OFFICER",
        "ASSISTANT SUPERINTENDENT",
        "CURRICULUM DIRECTOR",
        "MATH COORDINATOR",
        "ELA LITERACY COORDINATOR",
        "CTE COLLEGE CAREER READINESS DIRECTOR",
        "SPECIAL EDUCATION DIRECTOR",
        "ESL ELL COORDINATOR",
        "DIRECTOR ASSESSMENT DATA",
        "PROFESSIONAL LEARNING DIRECTOR",
    ]

    df = contacts_df.copy()

    if "Position" in df.columns:
        df["_priority"] = (
            df["Position"]
            .astype(str)
            .str.upper()
            .apply(lambda x: role_priority.index(x) if x in role_priority else 99)
        )
        df = df.sort_values([DISTRICT_COL, "_priority"])

    for _, row in df.iterrows():
        district = normalize_text(row.get(DISTRICT_COL, ""))
        if not district:
            continue

        first = normalize_text(row.get("First Name", ""))
        last = normalize_text(row.get("Last Name", ""))
        title = normalize_text(row.get("Title", ""))
        position = normalize_text(row.get("Position", ""))
        email = normalize_text(row.get("Email", ""))

        name = f"{first} {last}".strip()
        contact = format_contact_display(
            name=name,
            title=title,
            email=email,
            position=position,
        )

        if contact:
            key = district_key(district)
            lookup.setdefault(key, [])
            if contact not in lookup[key]:
                lookup[key].append(contact)

    return lookup


def build_contacts_lookup_from_leadership(leadership_df):
    lookup = {}

    if leadership_df.empty or DISTRICT_COL not in leadership_df.columns:
        return lookup

    for _, row in leadership_df.iterrows():
        district = normalize_text(row.get(DISTRICT_COL, ""))
        if not district:
            continue

        contacts = []

        superintendent = normalize_text(get_value(row, "Superintendent"))
        superintendent_email = normalize_text(get_value(row, "Email"))

        curriculum_lead = normalize_text(get_value(row, "Curriculum Lead"))
        curriculum_title = normalize_text(get_value(row, "Title")) or "Curriculum / Academic Lead"
        curriculum_email = normalize_text(get_value(row, "Email.1"))

        cte_lead = normalize_text(get_value(row, "CTE Lead"))
        cte_title = normalize_text(get_value(row, "Title.1")) or "CTE / Career Readiness Lead"
        cte_email = normalize_text(get_value(row, "Email.2"))

        math_lead = normalize_text(get_value(row, "Math Lead"))
        math_title = normalize_text(get_value(row, "Title.2")) or "Math Lead"
        math_email = normalize_text(get_value(row, "Email.3"))

        for contact in [
            format_contact_display(superintendent, "Superintendent", superintendent_email),
            format_contact_display(curriculum_lead, curriculum_title, curriculum_email),
            format_contact_display(cte_lead, cte_title, cte_email),
            format_contact_display(math_lead, math_title, math_email),
        ]:
            if contact:
                contact = contact.replace("[", "").replace("]", "")
                contact = contact.replace("(mailto:", " | ").replace(")", "")
                contacts.append(contact)

        lookup[district_key(district)] = contacts[:8]

    return lookup


def get_contacts(district_name, contacts_lookup, leadership_lookup):
    key_name = district_key(district_name)

    if key_name in contacts_lookup and contacts_lookup[key_name]:
        return contacts_lookup[key_name][:8]

    return leadership_lookup.get(key_name, [])[:8]


def build_contact_strategy(card):
    rows = []
    for contact in card.get("contacts", [])[:8]:
        persona = infer_persona(contact)
        rows.append({
            "contact": contact,
            "persona": persona,
            "angle": persona_angle(persona, card),
        })
    return rows


# ============================================================
# SIGNAL / TAG / CARD REASONING
# ============================================================

def infer_tags(strategic_row, score_row):
    tags = []
    combined_text = " ".join(normalize_text(v) for v in strategic_row.to_dict().values()).lower()

    if normalize_text(get_value(strategic_row, "Math Improvement Mentioned")) or normalize_text(get_value(strategic_row, "Math Priority Strength")) or "math" in combined_text:
        tags.append("Math")

    if normalize_text(get_value(strategic_row, "Intervention Focus")) or normalize_text(get_value(strategic_row, "MTSS/Tiered Support Mentioned")) or "mtss" in combined_text or "tier 2" in combined_text or "tier 3" in combined_text:
        tags.append("MTSS")

    if (
        "sped" in combined_text
        or "special education" in combined_text
        or "students with disabilities" in combined_text
        or "english learner" in combined_text
        or "emergent bilingual" in combined_text
        or "ell" in combined_text
        or "multilingual" in combined_text
    ):
        tags.append("SPED/ELL")

    if normalize_text(get_value(strategic_row, "Career Readiness Mentioned")) or normalize_text(get_value(strategic_row, "Career Readiness Details")) or "ccmr" in combined_text or "career" in combined_text:
        tags.append("CCMR")

    if normalize_text(get_value(strategic_row, "Teacher Capacity/PD Focus")) or normalize_text(get_value(strategic_row, "Teacher Capacity Details")) or "professional learning" in combined_text or "teacher capacity" in combined_text:
        tags.append("Teacher Capacity")

    if normalize_text(get_value(strategic_row, "Curriculum Review/Adoption Activity")) or normalize_text(get_value(strategic_row, "Curriculum Details")) or "curriculum" in combined_text or "hqim" in combined_text or "instructional materials" in combined_text:
        tags.append("Curriculum / HQIM")

    if normalize_text(get_value(strategic_row, "Active Grants (Yes/No)")) or normalize_text(get_value(strategic_row, "Grants Details")) or "grant" in combined_text or "title i" in combined_text or "funding" in combined_text:
        tags.append("Funding / Grants")

    if "literacy" in combined_text or "reading" in combined_text or "dyslexia" in combined_text or "science of reading" in combined_text or "structured literacy" in combined_text:
        tags.append("Literacy / Dyslexia")

    if "leadership" in combined_text or "walkthrough" in combined_text or "instructional coach" in combined_text or "coaching" in combined_text or "principal" in combined_text:
        tags.append("Coaching / Leadership")

    if "dashboard" in combined_text or "data system" in combined_text or "progress monitoring" in combined_text or "analytics" in combined_text or "monitoring" in combined_text:
        tags.append("Data Systems")

    if normalize_text(get_value(score_row, "Existing Relationships")).lower() == "yes":
        tags.append("Existing Relationship")

    if not tags:
        tags.append("Strategic Review")

    unique = []
    for tag in tags:
        if tag not in unique:
            unique.append(tag)

    return unique


def extract_signals(strategic_row, basic, csi_counts, tsi_counts, district_id, score_row):
    signals = []

    field_map = [
        ("Strategic themes", "Strategic Plan Themes"),
        ("Math signal", "Math Priority Strength"),
        ("Intervention signal", "Intervention Focus Details"),
        ("MTSS signal", "MTSS Details"),
        ("SPED/ELL signal", "SPED/ELL Details"),
        ("Teacher capacity signal", "Teacher Capacity Details"),
        ("CCMR / career readiness signal", "Career Readiness Details"),
        ("Curriculum / implementation signal", "Curriculum Details"),
        ("Funding / grants signal", "Grants Details"),
        ("Additional notes", "Notes"),
    ]

    for label, col in field_map:
        value = normalize_text(get_value(strategic_row, col))
        if value:
            signals.append(f"{label}: {value}")

    csi = normalize_text(get_value(basic, "CSI Schools")) or str(csi_counts.get(district_id, ""))
    tsi = normalize_text(get_value(basic, "TSI Schools")) or str(tsi_counts.get(district_id, ""))
    if csi or tsi:
        signals.append(f"Accountability pressure: {csi or '0'} CSI schools and {tsi or '0'} TSI schools.")

    context_parts = []
    context_fields = [
        ("schools", "Number of Schools", "text"),
        ("grade span", "Grade Span Served", "text"),
        ("setting", "Urban/Suburban/Rural", "text"),
        ("economically disadvantaged", "% Economically Disadvantaged", "percent"),
        ("English learner", "% English Learner", "percent"),
        ("special education", "% Special Education", "percent"),
        ("major student group", "Major Student Groups", "text"),
        ("student growth trend", "Student Growth Trend", "percent1"),
    ]

    for label, col, kind in context_fields:
        raw_value = get_value(basic, col)
        value = normalize_text(raw_value)

        if value:
            if kind == "percent":
                value = format_percent(raw_value, 0)
            elif kind == "percent1":
                value = format_percent(raw_value, 1)
            context_parts.append(f"{label}: {value}")

    if context_parts:
        signals.append("District context: " + "; ".join(context_parts) + ".")

    existing_relationships = normalize_text(get_value(score_row, "Existing Relationships"))
    existing_contracts = normalize_text(get_value(score_row, "Existing Contracts in Region (Yes/No)"))

    if existing_contracts and existing_relationships:
        signals.append(f"Relationship context: existing contracts in region = {existing_contracts}; existing relationships = {existing_relationships}.")
    elif existing_relationships:
        signals.append(f"Relationship context: existing relationships = {existing_relationships}.")
    elif existing_contracts:
        signals.append(f"Relationship context: existing contracts in region = {existing_contracts}.")

    return signals


def search_blob(card):
    values = [
        card.get("name", ""),
        card.get("tier", ""),
        str(card.get("score", "")),
        str(card.get("strategic_score", "")),
        card.get("enrollment", ""),
        card.get("priority", ""),
        card.get("primary_pain", ""),
        card.get("best_first_conversation", ""),
    ]

    for field in ["tags", "signals", "questions", "listen_for", "guidance", "next_steps"]:
        values.extend(card.get(field, []))

    for p in card.get("pain_patterns", []):
        values.extend([p.get("name", ""), p.get("description", ""), p.get("leadership_angle", "")])

    for s in card.get("solution_pathways", []):
        values.extend([s.get("family", ""), s.get("offering", ""), s.get("positioning", ""), s.get("first_move", "")])

    return " ".join(map(str, values)).lower()


def keyword_hits(blob, keywords):
    hits = []
    for kw in keywords:
        if normalize_key(kw) in blob:
            hits.append(kw)
    return hits


def diagnose_pain_patterns(card):
    blob = search_blob_for_reasoning(card)
    card_tags = set(card.get("tags", []))
    patterns = []

    for pattern in PAIN_PATTERN_LIBRARY:
        pattern_tags = set(pattern.get("tags", []))
        overlap = card_tags.intersection(pattern_tags)
        hits = keyword_hits(blob, pattern.get("keywords", []))

        score = 0
        score += len(overlap) * 1.6
        score += min(len(hits), 6) * 0.7

        if card.get("priority") in ["Very High", "High"]:
            score += 0.4

        if pattern["name"] == "Funding / Sustainability Constraint" and "Funding / Grants" not in card_tags:
            score -= 1.2

        if score > 0.9:
            confidence = "High" if score >= 5 else "Medium-High" if score >= 3 else "Selective"
            evidence = []

            if overlap:
                evidence.append("Matched tags: " + ", ".join(sorted(overlap)))
            if hits:
                evidence.append("Matched language: " + ", ".join(hits[:5]))

            patterns.append({
                "name": pattern["name"],
                "description": pattern["description"],
                "score": score,
                "confidence": confidence,
                "bd_motion": pattern["bd_motion"],
                "best_first_conversation": pattern["best_first_conversation"],
                "leadership_angle": pattern["leadership_angle"],
                "evidence": evidence,
            })

    patterns = sorted(patterns, key=lambda x: x["score"], reverse=True)
    return patterns[:5]


def search_blob_for_reasoning(card):
    values = []
    values.extend(card.get("tags", []))
    values.extend(card.get("signals", []))
    return " ".join(map(str, values)).lower()


def recommend_solution_pathways(card):
    blob = search_blob_for_reasoning(card)
    card_tags = set(card.get("tags", []))
    pain_names = [p["name"] for p in card.get("pain_patterns", [])]
    ranked = []

    for service in PCG_SERVICE_LIBRARY:
        score = 0
        service_tags = set(service.get("tags", []))
        tag_overlap = card_tags.intersection(service_tags)
        pain_overlap = set(pain_names).intersection(set(service.get("pain_patterns", [])))
        hits = keyword_hits(blob, service.get("keywords", []))

        score += len(tag_overlap) * 1.1
        score += len(pain_overlap) * 2.2
        score += min(len(hits), 6) * 0.45

        if "Existing Relationship" in card_tags:
            score += 0.25

        if card.get("priority") in ["Very High", "High"]:
            score += 0.25

        if score > 1.2:
            if score >= 6:
                fit = "High"
            elif score >= 4:
                fit = "Medium-High"
            else:
                fit = "Selective"

            reason = []
            if pain_overlap:
                reason.append("Pain pattern fit: " + ", ".join(sorted(pain_overlap)))
            if tag_overlap:
                reason.append("Tag fit: " + ", ".join(sorted(tag_overlap)))
            if hits:
                reason.append("Signal language: " + ", ".join(hits[:4]))

            ranked.append({
                **service,
                "score": score,
                "fit_level": fit,
                "match_reason": "; ".join(reason) if reason else "General strategic fit based on district profile.",
            })

    ranked = sorted(ranked, key=lambda x: x["score"], reverse=True)
    return ranked[:5]


def build_listen_for(tags):
    tag_map = {
        "Math": ["math growth", "early numeracy", "Algebra readiness", "STAAR math", "fluency", "student practice"],
        "MTSS": ["intervention fidelity", "Tier 2", "Tier 3", "progress monitoring", "campus variation"],
        "SPED/ELL": ["access to grade-level instruction", "service delivery", "emergent bilingual students", "students with disabilities", "subgroup gaps"],
        "CCMR": ["pathways", "industry-based certifications", "TSIA2", "dual credit", "career awareness"],
        "Teacher Capacity": ["teacher burden", "coaching", "professional learning", "PLC routines", "instructional consistency"],
        "Curriculum / HQIM": ["adoption fidelity", "HQIM", "curriculum implementation", "Eureka", "Bluebonnet", "instructional materials"],
        "Funding / Grants": ["Title funding", "grant alignment", "Medicaid", "implementation funding", "budget constraints"],
        "Literacy / Dyslexia": ["science of reading", "structured literacy", "dyslexia", "foundational skills", "reading intervention"],
        "Coaching / Leadership": ["walkthroughs", "observation tools", "feedback routines", "leadership pipeline", "coach onboarding"],
        "Data Systems": ["dashboards", "analytics", "data cycles", "implementation monitoring", "reporting"],
        "Existing Relationship": ["existing relationship", "current contract", "expansion", "trusted partner"],
    }

    listen_for = []
    for tag in tags:
        listen_for.extend(tag_map.get(tag, []))

    listen_for.extend(["implementation barriers", "capacity constraints", "pilot readiness"])

    unique = []
    for item in listen_for:
        if item not in unique:
            unique.append(item)

    return unique


def build_guidance(card):
    tags = card.get("tags", [])
    pathways = card.get("solution_pathways", [])
    guidance = []

    guidance.append("Do not lead with a product pitch. Start with the district’s stated priorities and implementation realities.")

    if pathways:
        guidance.append(pathways[0].get("do_not_lead_with", ""))

    if "Existing Relationship" in tags:
        guidance.append("Build from known PCG credibility and relationship context before introducing a new idea.")

    if "Curriculum / HQIM" in tags:
        guidance.append("If curriculum or HQIM is present, frame PCG around implementation support rather than replacement materials.")

    if "SPED/ELL" in tags:
        guidance.append("If SPED/ELL is present, connect the conversation to access, service delivery, progress monitoring, and implementation visibility.")

    return compact_items(guidance, limit=5, max_chars=240)


def build_conversation_guide(card):
    tags = card.get("tags", [])
    primary_pain = card.get("primary_pain", "district implementation priorities")
    top_solution = card.get("solution_pathways", [{}])[0].get("offering", "PCG support") if card.get("solution_pathways") else "PCG support"

    opener = (
        f"We noticed several signals around {primary_pain.lower()}. "
        f"I’m curious how those priorities are showing up at the campus level right now."
    )

    questions = []

    questions.append("Where is implementation strongest today, and where does variability remain?")
    questions.append("What is the cost of that variability for teachers, principals, students, or families?")

    if "Math" in tags:
        questions.append("What is keeping math growth from becoming more consistent across campuses or grade levels?")

    if "Literacy / Dyslexia" in tags:
        questions.append("Where are literacy or dyslexia practices strongest today, and where does implementation still vary?")

    if "MTSS" in tags:
        questions.append("When students are identified for support, where does the intervention process slow down or vary most?")

    if "SPED/ELL" in tags:
        questions.append("Where are students with disabilities or emergent bilingual students still not receiving consistent access to grade-level expectations?")

    if "Teacher Capacity" in tags or "Curriculum / HQIM" in tags:
        questions.append("Where is professional learning or curriculum implementation not yet translating into classroom execution?")

    if "Coaching / Leadership" in tags:
        questions.append("How are coaching quality, walkthroughs, and feedback routines monitored today?")

    if "CCMR" in tags:
        questions.append("Where do students begin connecting academic skills to future pathways and career readiness expectations?")

    questions.append(f"Would it be useful to explore whether {top_solution} could address one focused barrier before discussing a larger scope?")
    questions.append("What would make a small pilot or diagnostic credible enough for principals, teachers, and district leaders to support expansion?")

    return {
        "opener": opener,
        "questions": compact_items(questions, limit=8, max_chars=235),
    }


def build_next_steps(card):
    pathways = card.get("solution_pathways", [])
    contacts = card.get("contact_strategy", [])

    steps = []

    if pathways:
        steps.append(f"Validate whether the strongest initial pathway is {pathways[0].get('offering')}.")

    if contacts:
        steps.append("Use the contact strategy to identify the best first conversation owner and tailor the opening angle by role.")

    if card.get("pain_patterns"):
        steps.append(f"Frame the first conversation around {card['pain_patterns'][0]['best_first_conversation']}")

    steps.append("Prepare a focused diagnostic, pilot, or implementation-support concept tied to one measurable district priority.")
    steps.append("Avoid presenting the full PCG portfolio too early; validate the problem and buyer ownership first.")

    if "Funding / Grants" in card.get("tags", []):
        steps.append("Identify aligned funding sources, including Title, grant, Medicaid, or strategic funds, before proposing scope.")

    return compact_items(steps, limit=6, max_chars=250)


def build_score_explanation(card):
    tags = card.get("tags", [])
    explanations = []

    overall_num = parse_float(card.get("score"))
    strategic_num = parse_float(card.get("strategic_score"))

    if card.get("tier") == "Tier 1":
        explanations.append("Tier 1 placement signals strong overall fit and near-term prioritization.")
    elif card.get("tier") == "Tier 2":
        explanations.append("Tier 2 placement indicates meaningful opportunity with strategic alignment.")
    elif card.get("tier"):
        explanations.append(f"{card.get('tier')} placement indicates the district may be worth pursuing selectively.")

    if overall_num is not None:
        if overall_num >= 4.5:
            explanations.append("Very strong overall weighted score suggests unusually strong fit across multiple factors.")
        elif overall_num >= 4.0:
            explanations.append("Strong overall weighted score suggests the district is above average on strategic attractiveness.")
        elif overall_num >= 3.5:
            explanations.append("Moderate-to-strong score suggests selective but real opportunity.")

    if strategic_num is not None and strategic_num >= 4.0:
        explanations.append("Strategic score is especially strong, indicating substantive district priorities align to PCG positioning.")
    elif strategic_num is not None and strategic_num >= 3.5:
        explanations.append("Strategic score is solid enough to justify a guided discovery conversation.")

    if "Math" in tags:
        explanations.append("Math signal increases relevance for Emerald resources, math professional learning, intervention routines, or math implementation support.")
    if "MTSS" in tags:
        explanations.append("MTSS signal increases relevance for intervention workflow design and progress monitoring.")
    if "SPED/ELL" in tags:
        explanations.append("SPED/ELL signal suggests potential relevance for access, service delivery, inclusive practice, compliance visibility, or subgroup monitoring.")
    if "Literacy / Dyslexia" in tags:
        explanations.append("Literacy/dyslexia signal suggests relevance for structured literacy review, Science of Reading support, or reading implementation coherence.")
    if "Coaching / Leadership" in tags:
        explanations.append("Coaching/leadership signal suggests relevance for observation tools, feedback routines, coach onboarding, or leadership development.")

    unique = []
    for item in explanations:
        if item not in unique:
            unique.append(item)

    return unique[:7]


def score_band(score):
    n = parse_float(score)
    if n is None:
        return "Unknown"
    if n >= 4.5:
        return "Exceptional"
    if n >= 4.0:
        return "Strong"
    if n >= 3.5:
        return "Moderate-Strong"
    if n >= 3.0:
        return "Selective"
    return "Monitor"


def signal_density(tags, signals):
    meaningful = [t for t in tags if t != "Existing Relationship"]
    if len(meaningful) >= 6 or len(signals) >= 8:
        return "High"
    if len(meaningful) >= 3 or len(signals) >= 4:
        return "Medium"
    return "Low"


# ============================================================
# CARD BUILDING
# ============================================================

def build_card(strategic_row, score_row, basic_lookup, contacts_lookup, leadership_lookup, csi_counts, tsi_counts, score_cols):
    district_name = normalize_text(get_value(strategic_row, DISTRICT_COL))
    district_id = district_key(district_name)

    raw_score = score_row.get(score_cols.get("overall_score"), "")
    raw_strategic_score = score_row.get(score_cols.get("strategic_score"), "") if score_cols.get("strategic_score") else ""

    score = format_number(raw_score, 2)
    strategic_score = format_number(raw_strategic_score, 2)
    tier = normalize_text(score_row.get(score_cols.get("tier"), "")) if score_cols.get("tier") else ""
    enrollment = format_enrollment(score_row.get(score_cols.get("enrollment"), "")) if score_cols.get("enrollment") else ""

    tags = infer_tags(strategic_row, score_row)
    priority = determine_priority(tier, raw_score)
    basic = basic_lookup.get(district_id, {})
    signals = extract_signals(strategic_row, basic, csi_counts, tsi_counts, district_id, score_row)

    card = {
        "name": district_name,
        "tier": tier,
        "score": score,
        "strategic_score": strategic_score,
        "enrollment": enrollment,
        "priority": priority,
        "tags": tags,
        "contacts": get_contacts(district_name, contacts_lookup, leadership_lookup),
        "signals": signals,
        "score_band": score_band(score),
        "signal_density": signal_density(tags, signals),
    }

    card["pain_patterns"] = diagnose_pain_patterns(card)
    card["primary_pain"] = card["pain_patterns"][0]["name"] if card["pain_patterns"] else "Strategic fit requires validation"
    card["best_first_conversation"] = card["pain_patterns"][0]["best_first_conversation"] if card["pain_patterns"] else "Clarify district priorities and implementation barriers."
    card["solution_pathways"] = recommend_solution_pathways(card)
    card["contact_strategy"] = build_contact_strategy(card)
    card["listen_for"] = build_listen_for(tags)
    card["guidance"] = build_guidance(card)
    card["conversation_guide"] = build_conversation_guide(card)
    card["questions"] = card["conversation_guide"]["questions"]
    card["score_explanation"] = build_score_explanation(card)
    card["next_steps"] = build_next_steps(card)

    return card


# ============================================================
# WORKBOOK LOAD
# ============================================================

def get_workbook_source(uploaded_file):
    if uploaded_file is not None:
        return uploaded_file.getvalue(), "Uploaded workbook override", "bytes"

    if os.path.exists(BUILT_IN_WORKBOOK):
        return BUILT_IN_WORKBOOK, "Built-in workbook", "path"

    return None, "No workbook found", "none"


@st.cache_data(show_spinner=False)
def load_cards_from_workbook_cached(workbook_payload, source_type):
    if source_type == "bytes":
        xls = pd.ExcelFile(BytesIO(workbook_payload), engine="openpyxl")
    else:
        xls = pd.ExcelFile(workbook_payload, engine="openpyxl")

    strategic_df = read_table_sheet(xls, SHEET_STRATEGIC)
    scorecard_df = read_table_sheet(xls, SHEET_SCORECARD)
    basic_df = read_table_sheet(xls, SHEET_BASIC)
    contacts_df = read_table_sheet(xls, SHEET_CONTACTS)
    leadership_df = read_table_sheet(xls, SHEET_LEADERSHIP)
    csi_df = read_table_sheet(xls, SHEET_CSI)
    tsi_df = read_table_sheet(xls, SHEET_TSI)

    if strategic_df.empty or scorecard_df.empty:
        return (
            [],
            xls.sheet_names,
            strategic_df,
            scorecard_df,
            basic_df,
            contacts_df,
            leadership_df,
            csi_df,
            tsi_df,
            pd.DataFrame(),
            {},
        )

    score_cols = detect_scorecard_columns(scorecard_df)
    score_lookup = build_score_lookup(scorecard_df, score_cols)
    basic_lookup = build_basic_lookup(basic_df)
    contacts_lookup = build_contacts_lookup_from_contacts(contacts_df)
    leadership_lookup = build_contacts_lookup_from_leadership(leadership_df)
    csi_counts, tsi_counts = build_csi_tsi_counts(csi_df, tsi_df)

    cards = []
    audit_rows = []

    for _, strategic_row in strategic_df.iterrows():
        district_name = normalize_text(get_value(strategic_row, DISTRICT_COL))
        if not district_name:
            continue

        district_id = district_key(district_name)
        substantive = is_substantive_strategic_row(strategic_row)
        score_row = score_lookup.get(district_id, {})
        matched = bool(score_row)

        tier = normalize_text(score_row.get(score_cols.get("tier"), "")) if matched and score_cols.get("tier") else ""
        overall = score_row.get(score_cols.get("overall_score"), "") if matched and score_cols.get("overall_score") else ""

        eligible = substantive and matched and bool(tier) and bool(normalize_text(overall))

        audit_rows.append({
            "District": district_name,
            "Substantive Strategic Indicators": substantive,
            "Matched Master Scorecard": matched,
            "Detected Tier Column": score_cols.get("tier"),
            "Tier": tier,
            "Detected Overall Score Column": score_cols.get("overall_score"),
            "Overall Score": normalize_text(overall),
            "Eligible": eligible,
        })

        if eligible:
            cards.append(
                build_card(
                    strategic_row,
                    score_row,
                    basic_lookup,
                    contacts_lookup,
                    leadership_lookup,
                    csi_counts,
                    tsi_counts,
                    score_cols,
                )
            )

    audit_df = pd.DataFrame(audit_rows)

    return (
        cards,
        xls.sheet_names,
        strategic_df,
        scorecard_df,
        basic_df,
        contacts_df,
        leadership_df,
        csi_df,
        tsi_df,
        audit_df,
        score_cols,
    )


# ============================================================
# FILTER / SORT
# ============================================================

def filter_cards(cards, query, selected_tiers, selected_tags, shortlist, show_shortlist_only):
    query = query.strip().lower()
    filtered = []

    for card in cards:
        if selected_tiers and card.get("tier") not in selected_tiers:
            continue

        if selected_tags and not any(tag in card.get("tags", []) for tag in selected_tags):
            continue

        if show_shortlist_only and card.get("name") not in shortlist:
            continue

        if query and query not in search_blob(card):
            continue

        filtered.append(card)

    return filtered


def sort_cards(cards, sort_mode):
    if sort_mode == "Priority then Score":
        return sorted(cards, key=lambda c: (priority_rank(c.get("priority")), score_sort_value(c)), reverse=True)

    if sort_mode == "Overall Score":
        return sorted(cards, key=score_sort_value, reverse=True)

    if sort_mode == "Strategic Score":
        return sorted(cards, key=lambda c: parse_float(c.get("strategic_score")) or 0, reverse=True)

    if sort_mode == "Enrollment":
        return sorted(cards, key=enrollment_sort_value, reverse=True)

    if sort_mode == "Signal Density":
        density_rank = {"High": 3, "Medium": 2, "Low": 1}
        return sorted(cards, key=lambda c: (density_rank.get(c.get("signal_density"), 0), score_sort_value(c)), reverse=True)

    if sort_mode == "District A-Z":
        return sorted(cards, key=lambda c: c.get("name", ""))

    return cards


# ============================================================
# WORD EXPORT
# ============================================================

def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def build_overview(card):
    tag_text = ", ".join([t for t in card.get("tags", []) if t != "Existing Relationship"][:5])
    pieces = [
        f"{card.get('name')} is a {card.get('priority', '').lower()} priority opportunity with strategic signals around {tag_text}.",
    ]

    if card.get("enrollment"):
        pieces.append(f"The district serves approximately {card.get('enrollment')} students.")

    pieces.append(
        f"Current profile: {card.get('tier', '')}, overall score {card.get('score', '')}, strategic score {card.get('strategic_score', '')}."
    )

    if card.get("primary_pain"):
        pieces.append(f"The strongest diagnosed pain pattern is {card.get('primary_pain')}.")

    return " ".join([p for p in pieces if normalize_text(p)])


def build_objectives(card):
    objectives = []
    for signal in card.get("signals", []):
        s = clean_bullet_text(signal)

        if s.lower().startswith("strategic themes:"):
            objectives.append(s.replace("Strategic themes:", "").strip())
        elif "Math signal:" in s:
            objectives.append(s.replace("Math signal:", "Improve mathematics outcomes:").strip())
        elif "Intervention signal:" in s or "MTSS signal:" in s:
            cleaned = s.replace("Intervention signal:", "").replace("MTSS signal:", "").strip()
            objectives.append(f"Strengthen intervention systems and implementation consistency: {cleaned}")
        elif "SPED/ELL signal:" in s:
            objectives.append(s.replace("SPED/ELL signal:", "Improve access and outcomes for priority student groups:").strip())
        elif "Teacher capacity signal:" in s:
            objectives.append(s.replace("Teacher capacity signal:", "Build educator capacity and instructional consistency:").strip())
        elif "CCMR / career readiness signal:" in s:
            objectives.append(s.replace("CCMR / career readiness signal:", "Increase college, career, and military readiness:").strip())
        elif "Curriculum / implementation signal:" in s:
            objectives.append(s.replace("Curriculum / implementation signal:", "Strengthen curriculum implementation and instructional coherence:").strip())

    if not objectives and card.get("pain_patterns"):
        objectives.append(card["pain_patterns"][0]["description"])

    return compact_items(objectives, limit=7, max_chars=275)


def build_pain_points_for_doc(card):
    points = []
    for p in card.get("pain_patterns", [])[:5]:
        points.append(f"{p['name']}: {p['description']}")
    return compact_items(points, limit=6, max_chars=260)


def build_next_steps_for_doc(card):
    return compact_items(card.get("next_steps", []), limit=6, max_chars=260)


def add_bullets(doc, items):
    if not items:
        doc.add_paragraph("Validate through discovery.", style="List Bullet")
        return
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def build_docx(cards, data_source_label="Built-in workbook"):
    doc = Document()

    title = doc.add_heading("District Intelligence Brief", 0)
    for run in title.runs:
        run.font.color.rgb = RGBColor(23, 54, 93)

    doc.add_paragraph("Business development intelligence profile generated from the Strategic District Field Guide workbook.")
    doc.add_paragraph(f"Generated: {export_timestamp()}")
    doc.add_paragraph(f"Data source: {data_source_label}")

    for idx, card in enumerate(cards):
        doc.add_heading(card.get("name", "District Profile"), level=1)

        doc.add_paragraph(
            f"{card.get('tier', '')} | Overall Score {card.get('score', '')} | "
            f"Strategic Score {card.get('strategic_score', '')} | Enrollment {card.get('enrollment', '')} | "
            f"Priority {card.get('priority', '')}"
        )

        doc.add_heading("1. Executive BD Summary", level=2)
        doc.add_paragraph(f"Primary Pain Pattern: {card.get('primary_pain', '')}")
        doc.add_paragraph(f"Best First Conversation: {card.get('best_first_conversation', '')}")
        if card.get("solution_pathways"):
            doc.add_paragraph(f"Top PCG Pathway to Validate: {card['solution_pathways'][0].get('offering', '')}")

        doc.add_heading("2. District Overview", level=2)
        doc.add_paragraph(build_overview(card))

        doc.add_heading("3. What the District Appears to Be Trying to Solve", level=2)
        add_bullets(doc, build_objectives(card))

        doc.add_heading("4. Pain Pattern Diagnosis", level=2)
        add_bullets(doc, build_pain_points_for_doc(card))

        doc.add_heading("5. Recommended PCG Solution Pathways", level=2)
        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers = ["Pain / Need", "PCG Solution Pathway", "BD Motion", "How to Position"]
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            shade_cell(cell, "1F4E79")
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)
                    run.bold = True

        for s in card.get("solution_pathways", [])[:5]:
            row = table.add_row().cells
            row[0].text = ", ".join(s.get("pain_patterns", [])[:2])
            row[1].text = s.get("offering", "")
            row[2].text = s.get("bd_motion", "")
            row[3].text = s.get("positioning", "")
            for cell in row:
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

        doc.add_heading("6. Contact Strategy", level=2)
        if card.get("contact_strategy"):
            contact_table = doc.add_table(rows=1, cols=3)
            contact_table.style = "Table Grid"
            headers = ["Decision-Maker / Influencer", "Likely Persona", "Suggested Conversation Angle"]
            for i, header in enumerate(headers):
                cell = contact_table.rows[0].cells[i]
                cell.text = header
                shade_cell(cell, "1F4E79")
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.bold = True

            for c in card.get("contact_strategy", [])[:8]:
                row = contact_table.add_row().cells
                row[0].text = c.get("contact", "")
                row[1].text = c.get("persona", "")
                row[2].text = c.get("angle", "")
        else:
            doc.add_paragraph("No contacts were available in the workbook.", style="List Bullet")

        doc.add_heading("7. Recommended Leadership Approach with NEPQ Integration", level=2)
        doc.add_paragraph("Opening statement:")
        doc.add_paragraph(card.get("conversation_guide", {}).get("opener", ""))

        doc.add_paragraph("NEPQ-style discovery questions:")
        add_bullets(doc, card.get("conversation_guide", {}).get("questions", []))

        doc.add_heading("8. Suggested First Move", level=2)
        add_bullets(doc, build_next_steps_for_doc(card))

        doc.add_heading("9. Internal Notes / Validate Before Pitching", level=2)
        add_bullets(doc, card.get("guidance", []))

        if idx < len(cards) - 1:
            doc.add_page_break()

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def build_matrix_docx(cards, data_source_label="Built-in workbook"):
    doc = Document()

    section = doc.sections[0]
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)

    title = doc.add_heading("District Strategic Positioning Matrix", 0)
    for run in title.runs:
        run.font.color.rgb = RGBColor(31, 78, 121)

    doc.add_paragraph("Scope: Currently filtered districts from the Strategic District Intelligence Brief.")
    doc.add_paragraph(f"Generated: {export_timestamp()}")
    doc.add_paragraph(f"Data source: {data_source_label}")

    table = doc.add_table(rows=1, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    headers = ["District", "Primary Pain", "Top PCG Pathway", "BD Motion", "Top Contact Angle", "Next Step"]
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        shade_cell(cell, "1F4E79")
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.bold = True

    for card in cards:
        row = table.add_row().cells
        row[0].text = f"{card.get('name','')}\n{card.get('tier','')} | Score {card.get('score','')}"
        row[1].text = card.get("primary_pain", "")
        row[2].text = card.get("solution_pathways", [{}])[0].get("offering", "") if card.get("solution_pathways") else ""
        row[3].text = card.get("solution_pathways", [{}])[0].get("bd_motion", "") if card.get("solution_pathways") else ""
        row[4].text = card.get("contact_strategy", [{}])[0].get("angle", "") if card.get("contact_strategy") else ""
        row[5].text = card.get("next_steps", [""])[0] if card.get("next_steps") else ""
        for cell in row:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


# ============================================================
# RENDERING
# ============================================================

def render_briefing_mode(card):
    top_solution = card.get("solution_pathways", [{}])[0] if card.get("solution_pathways") else {}

    st.markdown(f"""
    <div class="brief-box">
        <strong>2-Minute Briefing</strong><br><br>
        <strong>Primary pain pattern:</strong> {safe_html(card.get("primary_pain", ""))}<br>
        <strong>Best first conversation:</strong> {safe_html(card.get("best_first_conversation", ""))}<br>
        <strong>Top PCG pathway:</strong> {safe_html(top_solution.get("offering", "Validate through discovery"))}<br>
        <strong>BD motion:</strong> {safe_html(top_solution.get("bd_motion", ""))}<br><br>
        <strong>Opening line:</strong> {safe_html(card.get("conversation_guide", {}).get("opener", ""))}
    </div>
    """, unsafe_allow_html=True)


def render_pain_patterns(card):
    st.markdown("**Pain Pattern Diagnosis**")
    if not card.get("pain_patterns"):
        st.info("No specific pain pattern was diagnosed. Use discovery to validate fit.")
        return

    for p in card.get("pain_patterns", [])[:5]:
        evidence = "; ".join(p.get("evidence", []))
        st.markdown(f"""
        <div class="pain-card">
            <strong>{safe_html(p.get("name", ""))}</strong>
            <span class="priority priority-medium-high">{safe_html(p.get("confidence", ""))}</span><br>
            {safe_html(p.get("description", ""))}<br><br>
            <strong>BD motion:</strong> {safe_html(p.get("bd_motion", ""))}<br>
            <strong>Best first conversation:</strong> {safe_html(p.get("best_first_conversation", ""))}<br>
            <strong>Evidence:</strong> {safe_html(evidence)}
        </div>
        """, unsafe_allow_html=True)


def render_solution_pathways(card):
    st.markdown("**Recommended PCG Solution Pathways**")
    if not card.get("solution_pathways"):
        st.info("No specific pathway identified. Use discovery to validate fit.")
        return

    for s in card.get("solution_pathways", [])[:5]:
        personas = ", ".join(s.get("buyer_personas", [])[:5])
        question = s.get("discovery", [""])[0] if s.get("discovery") else ""
        st.markdown(f"""
        <div class="solution-card">
            <strong>{safe_html(s.get("offering", ""))}</strong>
            <span class="priority priority-high">{safe_html(s.get("fit_level", ""))} Fit</span><br>
            <span class="helper-text">{safe_html(s.get("family", ""))}</span><br><br>
            <strong>Why it matches:</strong> {safe_html(s.get("match_reason", ""))}<br>
            <strong>BD motion:</strong> {safe_html(s.get("bd_motion", ""))}<br>
            <strong>How to position:</strong> {safe_html(s.get("positioning", ""))}<br>
            <strong>Do not lead with:</strong> {safe_html(s.get("do_not_lead_with", ""))}<br>
            <strong>Likely buyers:</strong> {safe_html(personas)}<br>
            <strong>First question:</strong> {safe_html(question)}
        </div>
        """, unsafe_allow_html=True)


def render_contact_strategy(card):
    st.markdown("**Contact Strategy**")
    if not card.get("contact_strategy"):
        st.markdown("_No contacts available in workbook._")
        return

    df = pd.DataFrame(card.get("contact_strategy", []))
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_conversation_guide(card):
    guide = card.get("conversation_guide", {})
    st.markdown("**Conversation Builder**")
    st.info(guide.get("opener", ""))

    st.markdown("**NEPQ-style questions**")
    for q in guide.get("questions", []):
        st.markdown(f"- {q}")

    st.markdown("**Listen for**")
    st.markdown("".join(chip_html(i) for i in card.get("listen_for", [])[:18]), unsafe_allow_html=True)


def render_evidence(card):
    st.markdown("**Strategic Evidence From Workbook**")
    for item in card.get("signals", []):
        st.markdown(f"- {item}")

    st.markdown("**Score Explanation**")
    for item in card.get("score_explanation", []):
        st.markdown(f"- {item}")

    st.markdown("**Guidance / Watchouts**")
    for item in card.get("guidance", []):
        st.markdown(f"- {item}")


def render_card(card, data_source_label, view_mode="Quick Brief"):
    priority_class = card.get("priority", "Medium").lower().replace(" ", "-")
    badges = "".join(badge_html(tag) for tag in card.get("tags", []))

    st.markdown(f"""
        <div class="intel-card">
            <h3>{safe_html(card['name'])} <span class="priority priority-{priority_class}">{safe_html(card.get('priority', ''))}</span></h3>
            <div class="meta">
                {safe_html(card.get('tier', ''))}
                | Overall Score {safe_html(card.get('score', ''))}
                | Strategic Score {safe_html(card.get('strategic_score', ''))}
                | Enrollment {safe_html(card.get('enrollment', ''))}
                | Score Band {safe_html(card.get('score_band', ''))}
                | Signal Density {safe_html(card.get('signal_density', ''))}
            </div>
            <div>{badges}</div>
        </div>
    """, unsafe_allow_html=True)

    render_briefing_mode(card)

    mini_tabs = st.tabs([
        "Brief",
        "Pain Diagnosis",
        "PCG Pathways",
        "Contacts",
        "Conversation Guide",
        "Evidence",
    ])

    with mini_tabs[0]:
        st.markdown("**What this district appears to be trying to solve**")
        for item in build_objectives(card)[:5]:
            st.markdown(f"- {item}")

        st.markdown("**Suggested first moves**")
        for item in card.get("next_steps", [])[:5]:
            st.markdown(f"- {item}")

    with mini_tabs[1]:
        render_pain_patterns(card)

    with mini_tabs[2]:
        render_solution_pathways(card)

    with mini_tabs[3]:
        render_contact_strategy(card)

    with mini_tabs[4]:
        render_conversation_guide(card)

    with mini_tabs[5]:
        render_evidence(card)

    quick_prep = (
        f"{card['name']} Intelligence Brief\n\n"
        f"Primary pain pattern: {card.get('primary_pain', '')}\n"
        f"Best first conversation: {card.get('best_first_conversation', '')}\n\n"
        f"Opening line:\n{card.get('conversation_guide', {}).get('opener', '')}\n\n"
        f"Top PCG pathways:\n"
        + "\n".join([f"- {s.get('offering','')} ({s.get('bd_motion','')})" for s in card.get("solution_pathways", [])[:3]])
        + "\n\n"
        f"NEPQ questions:\n"
        + "\n".join([f"- {q}" for q in card.get("conversation_guide", {}).get("questions", [])[:6]])
        + "\n\n"
        f"Next steps:\n"
        + "\n".join([f"- {n}" for n in card.get("next_steps", [])[:5]])
    )

    with st.expander("Copy Intelligence Brief", expanded=False):
        st.text_area("Copy/paste brief", quick_prep, height=320, key=f"brief_{card['name']}")

    st.download_button(
        "Download this district intelligence brief",
        data=build_docx([card], data_source_label=data_source_label),
        file_name=f"{card['name'].replace(' ', '_')}_District_Intelligence_Brief.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key=f"download_{card['name']}",
    )


def render_hot_accounts(cards):
    st.markdown('<div class="section-title">Hot Accounts</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="helper-text">Top accounts based on priority, score, pain-pattern clarity, and PCG pathway fit.</div>',
        unsafe_allow_html=True,
    )

    if not cards:
        st.info("No hot accounts available.")
        return

    hot = sorted(
        cards,
        key=lambda c: (
            priority_rank(c.get("priority")),
            score_sort_value(c),
            len(c.get("pain_patterns", [])),
            len(c.get("solution_pathways", [])),
            1 if "Existing Relationship" in c.get("tags", []) else 0,
        ),
        reverse=True,
    )[:8]

    for c in hot:
        top_solution = c.get("solution_pathways", [{}])[0].get("offering", "Validate through discovery") if c.get("solution_pathways") else "Validate through discovery"
        st.markdown(f"""
        <div class="hot-card">
            <strong>{safe_html(c.get("name", ""))}</strong> — {safe_html(c.get("priority", ""))}<br>
            <span class="helper-text">Score {safe_html(c.get("score", ""))} | {safe_html(c.get("tier", ""))}</span><br>
            <strong>Primary pain:</strong> {safe_html(c.get("primary_pain", ""))}<br>
            <strong>Top pathway:</strong> {safe_html(top_solution)}<br>
            <strong>First conversation:</strong> {safe_html(c.get("best_first_conversation", ""))}
        </div>
        """, unsafe_allow_html=True)


def render_compare_mode(cards):
    st.markdown('<div class="section-title">Compare Districts</div>', unsafe_allow_html=True)

    if len(cards) < 2:
        st.info("Select at least two visible districts to compare.")
        return

    options = [c["name"] for c in cards]

    selected = st.multiselect(
        "Choose districts to compare",
        options=options,
        default=options[: min(3, len(options))],
        max_selections=4,
    )

    selected_cards = [c for c in cards if c["name"] in selected]

    if not selected_cards:
        return

    compare_df = pd.DataFrame([
        {
            "District": c.get("name", ""),
            "Tier": c.get("tier", ""),
            "Priority": c.get("priority", ""),
            "Primary Pain": c.get("primary_pain", ""),
            "Top PCG Pathway": c.get("solution_pathways", [{}])[0].get("offering", "") if c.get("solution_pathways") else "",
            "BD Motion": c.get("solution_pathways", [{}])[0].get("bd_motion", "") if c.get("solution_pathways") else "",
            "Best First Conversation": c.get("best_first_conversation", ""),
        }
        for c in selected_cards
    ])

    st.dataframe(compare_df, use_container_width=True, hide_index=True)


def render_matrix(cards):
    st.markdown('<div class="section-title">Matrix View</div>', unsafe_allow_html=True)

    if not cards:
        st.info("No cards to show.")
        return

    df = pd.DataFrame([
        {
            "District": c.get("name", ""),
            "Tier": c.get("tier", ""),
            "Priority": c.get("priority", ""),
            "Score": c.get("score", ""),
            "Primary Pain": c.get("primary_pain", ""),
            "Top PCG Pathway": c.get("solution_pathways", [{}])[0].get("offering", "") if c.get("solution_pathways") else "",
            "BD Motion": c.get("solution_pathways", [{}])[0].get("bd_motion", "") if c.get("solution_pathways") else "",
            "First Conversation": c.get("best_first_conversation", ""),
            "Top Contact Angle": c.get("contact_strategy", [{}])[0].get("angle", "") if c.get("contact_strategy") else "",
        }
        for c in cards
    ])

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        "Download matrix as CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="district_intelligence_matrix.csv",
        mime="text/csv",
    )


def render_how_to(data_source_label):
    st.markdown('<div class="section-title">How-To Guide</div>', unsafe_allow_html=True)

    st.markdown(f"""
    <div class="howto-box">
    <strong>Current data source:</strong> {safe_html(data_source_label)}<br>
    This app uses the district workbook to diagnose pain patterns, recommend PCG solution pathways, infer contact strategy, and generate BD-ready intelligence briefs.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="howto-box">
    <strong>Recommended workflow</strong>
    <ol>
      <li>Search or filter to a target district.</li>
      <li>Read the 2-minute briefing.</li>
      <li>Review Pain Diagnosis before looking at products or services.</li>
      <li>Use PCG Pathways to decide which conversation lane to validate.</li>
      <li>Use Contacts to tailor the angle by role.</li>
      <li>Use Conversation Guide for NEPQ-style discovery.</li>
      <li>Download the District Intelligence Brief for follow-up planning.</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Useful searches")
    st.code("math\nMTSS\nSPED\nliteracy\ndyslexia\ncoaching\nPlaybook\ncareer\nfunding\nBluebonnet\nEureka", language="text")


def show_workbook_debug(sheet_names, strategic_df, scorecard_df, basic_df, contacts_df, leadership_df, csi_df, tsi_df, audit_df, score_cols, data_source_label):
    st.markdown('<div class="section-title">Workbook Diagnostics</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="helper-text">Current data source: <strong>{safe_html(data_source_label)}</strong></div>', unsafe_allow_html=True)

    st.markdown("### Sheets found")
    st.write(sheet_names)

    st.markdown("### Detected scorecard columns")
    st.write(score_cols)

    for label, df in [
        ("Strategic Indicators", strategic_df),
        ("Master Scorecard", scorecard_df),
        ("Basic District Info", basic_df),
        ("District Contacts", contacts_df),
        ("Leadership and Governance", leadership_df),
        ("CSI", csi_df),
        ("TSI", tsi_df),
    ]:
        with st.expander(label, expanded=False):
            st.write(f"Rows: {len(df)}")
            st.write("Columns:")
            st.write(list(df.columns))

            if DISTRICT_COL in df.columns:
                st.write("First district names found:")
                st.write(df[DISTRICT_COL].dropna().astype(str).head(15).tolist())

    st.markdown("### Eligibility audit")
    st.dataframe(audit_df, use_container_width=True)


# ============================================================
# APP
# ============================================================

st.markdown('<div class="main-title">Strategic District Intelligence Brief</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">BD-ready district intelligence: pain diagnosis, PCG solution pathways, contact strategy, and conversation guidance.</div>',
    unsafe_allow_html=True,
)

if "shortlist" not in st.session_state:
    st.session_state.shortlist = []

with st.sidebar:
    st.header("1. Data")
    uploaded_file = st.file_uploader("Optional: upload workbook override", type=["xlsx"])
    st.caption("If no workbook is uploaded, the app uses the built-in workbook from GitHub.")

workbook_payload, data_source_label, source_type = get_workbook_source(uploaded_file)

if workbook_payload is None:
    tab_field, tab_hot, tab_compare, tab_howto, tab_matrix, tab_diag = st.tabs(
        ["Field Guide", "Hot Accounts", "Compare Districts", "How-To Guide", "Matrix View", "Workbook Diagnostics"]
    )

    with tab_field:
        st.error(f"No workbook found. Add {BUILT_IN_WORKBOOK} to the GitHub repo or upload a workbook in the sidebar.")
    with tab_howto:
        render_how_to(data_source_label)
    with tab_matrix:
        st.info("Workbook required to populate the matrix.")
    with tab_diag:
        st.info("Workbook required to view diagnostics.")

    st.stop()

with st.spinner("Loading district intelligence..."):
    (
        cards,
        sheet_names,
        strategic_df,
        scorecard_df,
        basic_df,
        contacts_df,
        leadership_df,
        csi_df,
        tsi_df,
        audit_df,
        score_cols,
    ) = load_cards_from_workbook_cached(workbook_payload, source_type)

if not cards:
    tab_field, tab_hot, tab_compare, tab_howto, tab_matrix, tab_diag = st.tabs(
        ["Field Guide", "Hot Accounts", "Compare Districts", "How-To Guide", "Matrix View", "Workbook Diagnostics"]
    )

    with tab_field:
        st.warning("No eligible district cards were generated. Review the Workbook Diagnostics tab.")
    with tab_howto:
        render_how_to(data_source_label)
    with tab_matrix:
        st.info("No eligible cards available.")
    with tab_diag:
        show_workbook_debug(
            sheet_names,
            strategic_df,
            scorecard_df,
            basic_df,
            contacts_df,
            leadership_df,
            csi_df,
            tsi_df,
            audit_df,
            score_cols,
            data_source_label,
        )

    st.stop()

all_districts = [card.get("name", "") for card in cards]
all_tiers = sorted({card.get("tier", "") for card in cards if card.get("tier")})
all_tags = sorted({tag for card in cards for tag in card.get("tags", [])})

with st.sidebar:
    st.success(f"Using: {data_source_label}")

    st.header("2. Search")
    query = st.text_input("Search", placeholder="District, contact, signal, offering, pain...")

    st.header("3. Filters")
    selected_tiers = st.multiselect("Tier", options=all_tiers, default=[])
    selected_tags = st.multiselect("Strategic / solution tags", options=all_tags, default=[])

    st.header("4. Shortlist")
    shortlist_input = st.multiselect(
        "Shortlist districts",
        options=all_districts,
        default=[x for x in st.session_state.shortlist if x in all_districts],
    )
    st.session_state.shortlist = shortlist_input

    show_shortlist_only = st.checkbox("Show shortlisted only", value=False)

    if st.button("Clear shortlist"):
        st.session_state.shortlist = []
        st.rerun()

    st.header("5. Display")
    view_mode = st.radio("View mode", options=["Quick Brief", "Full Detail"], index=0)

    sort_mode = st.selectbox(
        "Sort by",
        options=[
            "Priority then Score",
            "Overall Score",
            "Strategic Score",
            "Enrollment",
            "Signal Density",
            "District A-Z",
        ],
        index=0,
    )

filtered_cards = filter_cards(
    cards,
    query,
    selected_tiers,
    selected_tags,
    st.session_state.shortlist,
    show_shortlist_only,
)

filtered_cards = sort_cards(filtered_cards, sort_mode)

tab_field, tab_hot, tab_compare, tab_howto, tab_matrix, tab_diag = st.tabs(
    ["Field Guide", "Hot Accounts", "Compare Districts", "How-To Guide", "Matrix View", "Workbook Diagnostics"]
)

with tab_field:
    metric1, metric2, metric3, metric4, metric5 = st.columns(5)

    metric1.metric("📍 Cards Shown", len(filtered_cards))
    metric2.metric("✅ Eligible", len(cards))
    metric3.metric("🔥 High Priority", sum(1 for card in filtered_cards if card.get("priority") in ["High", "Very High"]))
    metric4.metric("⭐ Tier 1", sum(1 for card in filtered_cards if card.get("tier") == "Tier 1"))
    metric5.metric("🧠 High Signal", sum(1 for card in filtered_cards if card.get("signal_density") == "High"))

    if query:
        st.markdown(f'<div class="helper-text">Showing results for: <strong>{safe_html(query)}</strong></div>', unsafe_allow_html=True)

    col_dl1, col_dl2, col_dl3 = st.columns([1, 1, 1])

    with col_dl1:
        st.download_button(
            "Download District Intelligence Brief",
            data=build_docx(filtered_cards, data_source_label=data_source_label),
            file_name="District_Intelligence_Brief.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            disabled=not filtered_cards,
        )

    with col_dl2:
        st.download_button(
            "Download Strategic Matrix",
            data=build_matrix_docx(filtered_cards, data_source_label=data_source_label),
            file_name="District_Strategic_Positioning_Matrix.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            disabled=not filtered_cards,
        )

    with col_dl3:
        selected_profile = st.selectbox(
            "One-district brief",
            options=[""] + [c["name"] for c in filtered_cards],
        )

        if selected_profile:
            selected_card = next(c for c in filtered_cards if c["name"] == selected_profile)
            st.download_button(
                "Download selected brief",
                data=build_docx([selected_card], data_source_label=data_source_label),
                file_name=f"{selected_profile.replace(' ', '_')}_District_Intelligence_Brief.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_selected_profile",
            )

    st.divider()

    if not filtered_cards:
        st.info("No district cards match the current search/filter.")
    else:
        for card in filtered_cards:
            left, right = st.columns([5, 1])

            with left:
                render_card(card, data_source_label=data_source_label, view_mode=view_mode)

            with right:
                district_name = card.get("name")

                if district_name in st.session_state.shortlist:
                    if st.button("Remove ⭐", key=f"remove_{district_name}"):
                        st.session_state.shortlist = [x for x in st.session_state.shortlist if x != district_name]
                        st.rerun()
                else:
                    if st.button("Shortlist ⭐", key=f"add_{district_name}"):
                        st.session_state.shortlist.append(district_name)
                        st.rerun()

            st.divider()

with tab_hot:
    render_hot_accounts(filtered_cards)

with tab_compare:
    render_compare_mode(filtered_cards)

with tab_howto:
    render_how_to(data_source_label)

with tab_matrix:
    render_matrix(filtered_cards)

    st.download_button(
        "Download matrix as Word document",
        data=build_matrix_docx(filtered_cards, data_source_label=data_source_label),
        file_name="District_Strategic_Positioning_Matrix.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=not filtered_cards,
        key="matrix_tab_download",
    )

with tab_diag:
    show_workbook_debug(
        sheet_names,
        strategic_df,
        scorecard_df,
        basic_df,
        contacts_df,
        leadership_df,
        csi_df,
        tsi_df,
        audit_df,
        score_cols,
        data_source_label,
    )
