import io
import re
import json
import unicodedata
from pathlib import Path
from collections import defaultdict, Counter
from zipfile import BadZipFile
import pandas as pd
from docx import Document
import streamlit as st

# ============================================================
# Constants (Rulebook)
# ============================================================

FIXED_EXPERIENCE_YEAR = 2026
BY_CATEGORY_SHEET = "By Category"

REQUIRED_COLUMNS = [
    "Name",
    "Years of Experience",
    "Language",
    "Cloud",
    "Databases",
    "OS",
    "Framework/Libraries",
    "DevOps",
    "Container/Orchestration",
    "Machine Learning/AI",
    "Networking",
    "Version Control",
    "Tools",
    "Other",
    "Certifications",
    "Degree/Associates",
    "Degree/Bachelors",
    "Degree/Masters",
    "Degree/Phds",
]

SKILL_CATEGORY_COLS = [
    "Language",
    "Cloud",
    "Databases",
    "OS",
    "Framework/Libraries",
    "DevOps",
    "Container/Orchestration",
    "Machine Learning/AI",
    "Networking",
    "Version Control",
    "Tools",
    "Other",
]

SKILLS_HEADERS = {
    "skills",
    "technical skills",
    "skills/tools",
    "skills/tools/technologies",
    "skills & tools",
}

EDU_HEADERS = {"education", "educations"}
CERT_HEADERS = {"certification", "certifications"}

NON_EXTRACTABLE_HEADERS = {
    "ASSOCIATION/HONORS:",
    "association/honors:",
    "summary",
    "clearance",
    "professional experience",
    "experience",
    "work experience",
    "employment history",
    "work history",
    "projects",
    "project experience",
    "publications",
    "awards",
    "organizations",
    "activities",
    "references",
    "Associations/Honors:",
    "Associations:",
    "Honors:"
}

_SPLIT = re.compile(r"[;,]")

_DASHES = str.maketrans({
    "‐": "-",
    "‑": "-",
    "‒": "-",
    "–": "-",
    "—": "-",
    "−": "-",
})

# ============================================================
# Rule 10: Compound splits (must run BEFORE normalization)
# ============================================================

COMPOUND_SPLITS = {
    "gitlab ci/cd": ["GitLab", "CI/CD"],
    "google earth/maps api": ["Google Earth", "Maps API"],
    "html/css": ["HTML", "CSS"],
    "oracle pl/sql": ["Oracle", "PL/SQL"],
    "scrum/kanban": ["Scrum", "Kanban"],
    "nifi gitlab": ["NiFI", "GitLab"],
    "mongodb spring jpa": ["MongoDB Spring", "Spring JPA"],
    "vmware java ee": ["VMWare", "Java EE"],
    "symantec. mcafee": ["Symantec", "McAfee"],
    "php visual basic": ["PHP", "Visual Basic"],
    "ida pro lod4j": ["IDA Pro", "Lod4j"],  
    "tomcat infinispan": ["Tomcat", "Infinispan"],
    "docker swarm docker compose": ["Docker Swarm", "Docker Compose"],
    "slick/scalaquery": ["SLICK", "ScalaQuery"],
    "xml/sqd" : ["XML","SQD"],
    "xml/xsd": ["XML", "XSD"]
}

# ============================================================
# Rule 11: Skill normalization (always applied)
# ============================================================

SKILL_NORMALIZATION = {
    "amazon aws": "AWS",
    "some aws": "AWS",
    "a ws": "AWS",
    "amazon": "AWS",
    "amazon ec2": "EC2",
    "ec2": "EC2",
    "arcgis tool": "ArcGIS",
    "arcgis tools": "ArcGIS",
    "bash scripting": "Bash",
    "extjs": "ExtJS",
    "ext.js": "ExtJS",
    "ida pro": "IDA Pro",
    "idapro": "IDA Pro",
    "jaws reader": "Jaws",
    "jupyter notebook": "Jupyter Notebooks",
    "jupyter": "Jupyter Notebooks",
    "red hat": "RedHat",
    "mips assembly": "Mips",
    "network protocol suites": "Network Protocol",
    "network protocols": "Network Protocol",
    "nodejs": "Node.js",
    "node.js": "Node.js",
    "objective c": "Objective‑C",
    "objective‑c": "Objective‑C",
    "onnx runtime": "ONNX",
    "python3": "Python",
    "r studio": "R",
    "and identity access management": "Identity Access Management",
    "groovy/grails tool suite": "Groovy/Grails",
    "groovy/grails tol suite": "Groovy/Grails",
    "groovy": "Groovy/Grails",
    "sql-lite": "SQLite",
    "vue.js": "Vue",
    "vue.js": "Vue",
    "visual studio code": "VSCode",
    "visual studio.": "VSCode",
    "vs code": "VSCode",
    "ms visual studio": "VSCode",
    "ajax": "AJAX",
    "randomforests": "Random Forests",
    "gaussian models.": "Gaussian Models",
    "nas and san storage arrays": "NAS/SAN Storage",
    "xlinx design suite": "Xilinx",
    "visio 2000": "Visio",
    "ms visio": "Visio",
    "microsoft office tools": "Microsoft Office",
    "xlinx and altera:": "FPGA Design",
    "all lsi and msi logic families": "LSI/MSI Logic",
    "spring boot deployments": "Spring Boot",
    "linux scripting": "LINUX",
    "linix": "LINUX",
    "linux servers": "LINUX",
    "cloudwatch": "CloudWatch",
    "res instances": "RES",
    "lambdas": "Lambda",
    "lambda": "Lambda",
    "ec2 instances": "EC2",
    "elasticSearch cvs": "Elasticsearch",
    "google earthmaps api": "Google Earth",
    "apache nifi": "Apache NiFi",
    "cvs.": "CVS",
    "shell": "Shell",
    "nifi": "NiFi",
    "gitlab": "GitLab",
    "openvpn.": "OpenVPN",
    "vmware": "VMWare",
    "xp": "Windows XP",
    "vista": "Windows Vista",
    "ubuntu.": "Ubuntu",
    "mash vm user.": "Mash VM",
    "pki": "PKIs",
    "removeview.": "RemoveView",
    "clustering": "Clustering",
    "cluster computing": "Clustering",
    "bash scripting": "Bash Scripting",
    "hpc": "High Performance Computing",
    "rhel": "RHEL",
    "windows server.": "Windows Server",
    "windows servers":"Windows Server",
    "window": "Windows",
    "remoteview.": "RemoteView",
    "microsoft windows xp": "Windows XP",
    "tomcat": "Tomcat",
    "matploblib": "Matplotlib",
    "sgu hardware": "SGI",
    "microsoft": "Microsoft Office",
    "ssh protocols": "SSH",
    "dns server config": "DNS",
    "scrum": "SCRUM",
    "swager": "Swagger",
    "cisco works:": "Cisco",
    "cisco prime.": "Cisco",
    "cisco event scripting": "Cisco",
    "and mgx 8800 series atm switches": "MGX 8800 Series ATM Switches",
    "tcl.": "TCL",
    "juniper srx series.": "Juniper SRX series",
    "and kiribati.": "Kiribati",
    "and kiribiti.": "Kiribati",
    "agile methodology": "Agile",
    "apache http server": "Apache HTTP",
    "consul and vault.":"Consultaiton and Vault",
    "java spring cloud": "Java Spring",
    "junit4/5": "Junit",
    "macosx": "MacOS X",
    "plsql": "PL/SQL",
    "spark.ml": "Spark",
    "sql developer": "SQL",
    "tensor analysis tool kit": "Tensor Analysis",
    "unix shell scripting": "UNIX"
}

 
# ============================================================
# Rule 12: Skill removal
# ============================================================

REMOVED_SKILLS = {"amazon management console eclipse"}

# ============================================================
# Rule 16: Certification normalization (CertificationFrequency ONLY)
# ============================================================

CERT_NORMALIZATION = {
    "aws solutions architect": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect - associate (aws certified)": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect associate": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect – associate": "AWS Certified Solutions Architect - Associate",
    "aws solutions certified - associate": "AWS Certified Solutions Architect - Associate",
    "aws certified developer": "AWS Certified Developer - Associate",
    "aws solutions architect (in process)": "AWS Certified Solutions Architect - Associate",
    "self-study (udemy courses): aws cloud practitioner": "AWS Certified Cloud Practitioner",
    "ccna (cisco certified network associate)": "CCNA",
    "ccna certification": "CCNA",
    "itil v3 foundation": "ITIL v3.0",
    "iril v3 foundation (cert# 894862)": "ITIL v3.0",
    "security+": "CompTIA Security+",
    "security +": "CompTIA Security+",
    "security+ (comptia)": "CompTIA Security+",
    "security+ ce": "CompTIA Security+",
    "comptia security+ ce": "CompTIA Security+",
    "comptia – security+": "CompTIA Security+",
    "comptia - security+": "CompTIA Security+",
    "comptia security+ ()":"CompTIA Security+",
    "security+ certification": "CompTIA Security+",
    "comptia networking +": "CompTIA Network+",
    "certified scrum master (csm)": "Certified Scrum Master",
    "certified scrum master": "Scrum Master",
    "certified scrum master - scrum alliance": "Scrum Master",
    "scrum alliance certified scrum master": "Scrum Master",
    "pmi agile certified practitioner (pmi-acp)":"PMI-ACP (Agile Certified Practitioner)",
    "pmp - project management institute (pmi)":"PMP (Project Management Professional)"
       
}

# ============================================================
# Rule 18: Degree normalization mapping (explicit only)
# ============================================================

def _degree_lookup_key(line: str) -> str:
    s = unicodedata.normalize("NFKC", (line or "")).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


# Your explicit entries (you said you manually lowercased long keys as needed)
DEGREE_NORMALIZATION = {
    # --- Your explicit entries ---
    "b.a. in computer science degree": "B.S., Computer Science",
    "b.a. in middle east studies / arabic": "B.S., Middle East Studies / Arabic",
    "b.a., mathematics/computer science": "B.S., Mathematics/Computer Science",
    "b.s. computer information systems, may 2011": "B.S., Computer Information Systems",
    "b.s. computer science": "B.S., Computer Science",
    "b.s. electrical engineering (1996)": "B.S., Electrical Engineering",
    "b.s. in computer networks & cybersecurity": "B.S., Computer Networks & Cybersecurity",
    "b.s., business and administration, 2018": "B.S., Business and Administration",
    "b.s., computer science (minor: business administration), towson university, 12/2010": "B.S., Computer Science",
    "b.s., computer science - university of maryland baltimore county (1999)": "B.S., Computer Science",
    "b.s., computer science, college of computer": "B.S., Computer Science",
    "b.s., electrical and computer engineering (minor in german)": "B.S., Electrical and Computer Engineering",
    "b.s., information systems management, university of maryland baltimore county (1988)": "B.S., Information Systems Management",
    "ba english, bryn mawr college": "B.S., English",
    "bachelor of engineering in computer science": "B.S., Computer Science",
    "bachelor of science degree, electronic media engineering": "B.S., Electronic Media Engineering",
    "bachelor of science electrical engineering": "B.S., Electrical Engineering",
    "bachelor of science in computer and information science": "B.S., Computer and Information Science",
    "bachelor of science in history": "B.S., History",
    "bachelor of science, electrical engineering, bucknell university, lewisburg pa (1984)": "B.S., Electrical Engineering",
    "bs in cybersecurity and computer science, mount st. mary’s university, december 2021": "B.S., Cybersecurity and Computer Science",
    "bs, business administration": "B.S., Business Administration",
    "bs, computer science": "B.S., Computer Science",
    "bs, computer science - cyber": "B.S., Computer Science - Cyber",
    "bs, computer science, university of baltimore, 1986": "B.S., Computer Science",
    "bs, cybersecurity": "B.S., Cybersecurity",
    "bs, economics": "B.S., Economics",
    "bs, information technology": "B.S., Information Technology",
    "bachelor in business admin": "B.S., Business and Administration",
    "2005: ba from the college of william and mary major in international studies and minor in computer science.": "B.S., International Studies and minor in Computer Science",
    "(in progress) bs in computer networks and security umuc": "B.S., Computer Networks and Security",
    "m.s.,data analytics": "M.S., Data Analytics",
    "master of business administration (mba)": "MBA",
    "master’s in business administration (mba)": "MBA",
    "mba, information technology": "MBA",
    "mba, university of baltimore, 1992": "MBA",
    "m.s.,computer science": "M.S., Computer Science",
    "master of science, computer science, johns hopkins university, baltimore md (1997)": "M.S., Computer Science",
    "master of science, electrical engineering": "M.S., Electrical Engineering",
    "master of science, electrical engineering, johns hopkins university, baltimore md (1989)": "M.S., Electrical Engineering"
    
}    

# Deterministic cleanup patterns
_degree_year_rx = re.compile(r"\b(19\d{2}|20\d{2})\b")
_degree_month_rx = re.compile(
    r"\b(?:jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\b",
    re.I
)
_degree_inst_rx = re.compile(r"\b(university|college|school|institute|campus)\b", re.I)
_degree_minor_rx = re.compile(r"\bminor\b", re.I)

def _degree_generic_cleanup(line: str) -> str:
    """
    Deterministic cleanup (Rule 18.1):
    - remove parentheticals containing year/month/minor
    - remove trailing comma segments containing year/month/institution
    - normalize prefixes: BA/BS/Bachelor->B.S., MS/Master->M.S.
    - enforce comma format after B.S./M.S.
    """
    s = unicodedata.normalize("NFKC", (line or "")).translate(_DASHES).strip()
    if not s:
        return ""

    # remove parentheses containing year/month/minor
    def paren_repl(m):
        inner = m.group(1)
        if _degree_year_rx.search(inner) or _degree_month_rx.search(inner) or _degree_minor_rx.search(inner):
            return ""
        return "(" + inner + ")"
    s = re.sub(r"\(([^)]*)\)", paren_repl, s)

    # drop trailing comma segments that look like institution/date/location metadata
    parts = [p.strip() for p in s.split(",")]
    kept = []
    for i, p in enumerate(parts):
        if i == 0:
            kept.append(p)
            continue
        if _degree_year_rx.search(p) or _degree_month_rx.search(p) or _degree_inst_rx.search(p):
            break
        kept.append(p)
    s = ", ".join([k for k in kept if k]).strip(" ,")

    # Normalize prefixes to prevent duplicates
    s = re.sub(r"^BA\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^B\.A\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^BS\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^B\.S\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of (Science|Arts)\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of Engineering\b", "B.S.", s, flags=re.I)

    s = re.sub(r"^MS\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^M\.S\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master of Science\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master in\b", "M.S.", s, flags=re.I)

    # ✅ FIX: remove punctuation-only segments like "." that create "B.S., ., X"
    segs = [seg.strip() for seg in s.split(",")]
    segs = [seg for seg in segs if seg and seg not in {".", "..", "..."}]
    # also remove segments that are only punctuation
    cleaned_segs = []
    for seg in segs:
        only_punct = re.fullmatch(r"[.\-–—_]+", seg) is not None
        if not only_punct:
            cleaned_segs.append(seg)
    s = ", ".join(cleaned_segs).strip()

    # re-normalize comma spacing again
    s = re.sub(r"\s*,\s*", ", ", s).strip(" ,")
    return s

# ============================================================
# Generalized “Label:” stripping inside SKILLS section (locked feature)
# ============================================================

_LABEL_ANYWHERE_RX = re.compile(r"(?i)(^|[\s,])([A-Za-z][A-Za-z0-9 /&\-\+]{0,60})\s*:\s*")

def _strip_skill_group_labels(line):
    s = (line or "").strip()
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    s = _LABEL_ANYWHERE_RX.sub(", ", s)
    s = re.sub(r"\s*,\s*", ", ", s)
    s = re.sub(r"(,\s*){2,}", ", ", s).strip(" ,")
    return s

def _skills_tokens_from_lines(skills_lines):
    tokens = []
    for line in skills_lines:
        cleaned = _strip_skill_group_labels(line)
        if not cleaned:
            continue
        tokens.extend([t.strip() for t in _SPLIT.split(cleaned) if t.strip()])
    return tokens


# ============================================================
# Certification parsing + cleaning (NEW FIX)
# ============================================================

_MONTHS_RX = r"(?:jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)"
_re_year = re.compile(r"\b(19\d{2}|20\d{2})\b")
_re_mmddyyyy = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
_re_month_year = re.compile(rf"\b{_MONTHS_RX}\b\s*\d{{0,2}}\s*,?\s*(19\d{{2}}|20\d{{2}})", re.I)
_re_paren = re.compile(r"\(([^)]*)\)")
_re_long_id = re.compile(r"\b[A-Z0-9]{8,}\b")
_re_code_like = re.compile(r"\b(?:COMP\d+|F\w{10,}|V\w{10,})\b", re.I)
_re_metadata_words = re.compile(r"\b(certification issued|Cert|Analyst# 10080|License|October|present|udemy courses|self-study|certification issues|issued|exp\.?|expires|Cert# 11073|expiration|taking test|attended Reinvent conference|202|in process)\b", re.I)


def _canon_text(s: str) -> str:
    s = unicodedata.normalize("NFKC", (s or "")).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _cert_lookup_key(s: str) -> str:
    return _canon_text(s).casefold()


def clean_cert_token(token: str) -> str:
    """Clean a single certification token deterministically (no inference)."""
    if token is None:
        return ""
    s = _canon_text(str(token))
    if not s:
        return ""

    # drop leading bullets
    s = re.sub(r"^[•\-\u2022\t\s]+", "", s).strip()

    # remove parenthetical metadata if it contains date/year or issuance words
    def paren_repl(m):
        inner = m.group(1)
        if _re_year.search(inner) or _re_mmddyyyy.search(inner) or _re_month_year.search(inner) or _re_metadata_words.search(inner):
            return ""
        return "(" + inner + ")"

    s = _re_paren.sub(paren_repl, s)

    # cut off at metadata word occurrence (e.g., 'certification issued')
    m = _re_metadata_words.search(s)
    if m:
        s = s[:m.start()].strip(" -;:,.\t")

    # remove date patterns
    s = _re_month_year.sub("", s)
    s = _re_mmddyyyy.sub("", s)
    # remove years
    s = _re_year.sub("", s)
    # remove IDs / codes
    s = _re_code_like.sub("", s)
    s = _re_long_id.sub("", s)

    # remove trailing keywords 'Certification' / 'Certificate'
    s = re.sub(r"\b(Certification|Certificate|certification|certificate|cert|Cert)\b\.?$", "", s, flags=re.I).strip()

    # normalize spaces and separators
    s = re.sub(r"\s+", " ", s).strip(" ;,-")

    # apply normalization mapping (explicit)
    key = _cert_lookup_key(s)
    if key in CERT_NORMALIZATION:
        s = CERT_NORMALIZATION[key]

    return s.strip()


def clean_certifications_from_lines(cert_lines):
    """
    Extract + clean certifications from Certification section lines.
    - Split only on commas/semicolons (Rule 15)
    - Treat line boundaries as separate entries (docx line-aware)
    - Remove dates/years/ids/issuance text
    - Output deterministic list
    """
    cleaned = []
    for line in cert_lines:
        line = _canon_text(line)
        if not line:
            continue
        # split only on commas/semicolons
        parts = [p.strip() for p in _SPLIT.split(line) if p.strip()]
        if not parts:
            continue
        for p in parts:
            c = clean_cert_token(p)
            if c:
                cleaned.append(c)

    # dedupe preserve order
    return list(dict.fromkeys(cleaned))


# ============================================================
# Utilities
# ============================================================

def ensure_by_category_columns(df):
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[REQUIRED_COLUMNS].copy()

def create_empty_by_category_df():
    return pd.DataFrame(columns=REQUIRED_COLUMNS)

def dedupe_preserve_order(items):
    seen = set()
    out = []
    for x in items:
        t = (x or "").strip()
        if not t:
            continue
        k = t.lower()
        if k not in seen:
            seen.add(k)
            out.append(t)
    return out

def split_on_commas_semicolons(text):
    if not text:
        return []
    return [p.strip() for p in _SPLIT.split(text) if p.strip()]

def _clean_header(s):
    return (s or "").strip().lower()

def _is_all_caps_header(line):
    t = (line or "").strip()
    if not t or len(t) > 60:
        return False
    letters = re.sub(r"[^A-Za-z]", "", t)
    if len(letters) < 3:
        return False
    return (t.upper() == t) and any(ch.isalpha() for ch in t)

def _is_section_boundary(line):
    h = _clean_header(line)
    return (
        h in SKILLS_HEADERS
        or h in EDU_HEADERS
        or h in CERT_HEADERS
        or h in NON_EXTRACTABLE_HEADERS
        or _is_all_caps_header(line)
    )

# ============================================================
# Rule 9–11 skill pipeline (canonical)
# ============================================================

def apply_compound_splitting(skill):
    k = re.sub(r"\s+", " ", (skill or "").strip().lower()).strip()
    return COMPOUND_SPLITS.get(k, [(skill or "").strip()])

def normalize_skill(skill):
    raw = (skill or "").strip()
    k = re.sub(r"\s+", " ", raw.lower()).strip()
    if k in REMOVED_SKILLS:
        return None
    return SKILL_NORMALIZATION.get(k, raw)

def process_skills(skills_raw):
    # Rule 9: dedupe within resume before further processing
    skills_raw = dedupe_preserve_order(skills_raw)

    # Rule 10: compound split
    expanded = []
    for s in skills_raw:
        expanded.extend(apply_compound_splitting(s))

    # Rule 11 + Rule 12
    normalized = []
    for s in expanded:
        ns = normalize_skill(s)
        if ns:
            normalized.append(ns)

    # final dedupe within resume
    return dedupe_preserve_order(normalized)

def parse_and_normalize_skills_from_cell(cell):
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    raw_tokens = [t.strip() for t in _SPLIT.split(str(cell)) if t.strip()]
    return process_skills(raw_tokens)

# ============================================================
# Persistent category map (JSON) — FIX: canonicalize keys via pipeline
# ============================================================

def load_category_map(path: Path):
    if path.exists():
        raw = json.loads(path.read_text(encoding="utf-8"))
    else:
        raw = {}

    canon = {}
    for key, cat in raw.items():
        # canonicalize mapping key(s) through Rule 9–11
        canon_keys = process_skills([key])
        for ck in canon_keys:
            canon[ck] = cat
    return canon

def save_category_map(path: Path, mapping: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(mapping, indent=2, ensure_ascii=False), encoding="utf-8")

# ============================================================
# DOCX parsing (safe + preserves internal line breaks)
# ============================================================

def _docx_lines(file_like):
    try:
        try:
            file_like.seek(0)
        except Exception:
            pass
        doc = Document(file_like)
        out = []
        for p in doc.paragraphs:
            raw = (p.text or "").strip()
            if not raw:
                continue
            parts = [x.strip() for x in raw.split("\n") if x.strip()]
            out.extend(parts)
        return out
    except (BadZipFile, Exception):
        return []

def peek_name_from_docx(file_like):
    for line in _docx_lines(file_like):
        if not _is_section_boundary(line):
            return line.strip()
    return ""

def _section_lines(lines, header_set):
    start = None
    for i, line in enumerate(lines):
        if _clean_header(line) in header_set:
            start = i + 1
            break
    if start is None:
        return []
    out = []
    for line in lines[start:]:
        if _is_section_boundary(line):
            break
        out.append(line)
    return out

def parse_resume_sections(file_like, name_override=None):
    file_like.seek(0)
    lines = _docx_lines(file_like)

    name = (name_override or "").strip()
    if not name:
        file_like.seek(0)
        name = peek_name_from_docx(file_like)

    skills_lines = _section_lines(lines, SKILLS_HEADERS)
    edu_lines = _section_lines(lines, EDU_HEADERS)
    cert_lines = _section_lines(lines, CERT_HEADERS)

    # Skills tokens are line-aware and label-stripped
    skills_raw = dedupe_preserve_order(_skills_tokens_from_lines(skills_lines))
    #certs_raw = dedupe_preserve_order(split_on_commas_semicolons(" ".join(cert_lines).strip()))
    certs_raw = clean_certifications_from_lines(cert_lines)


    return {
        "name": name,
        "skills_raw": skills_raw,
        "certs_raw": certs_raw,
        "education_lines": edu_lines,
        "education_text": " ".join(edu_lines).strip(),
    }

# ============================================================
# Mapping from existing By Category
# ============================================================

def build_category_map_from_by_category(by_cat):
    by_cat = ensure_by_category_columns(by_cat)
    mapping = {}
    conflicts = defaultdict(set)

    for _, row in by_cat.iterrows():
        for col in SKILL_CATEGORY_COLS:
            skills = parse_and_normalize_skills_from_cell(row.get(col, ""))
            for s in skills:
                if s in mapping and mapping[s] != col:
                    conflicts[s].update({mapping[s], col})
                else:
                    mapping[s] = col

    conflicts = {k: v for k, v in conflicts.items() if len(v) > 1}
    return mapping, conflicts

def resolve_conflicts_with_user(conflicts, category_map):
    st.error("Conflicts found: same skill appears under multiple columns.")
    for skill in sorted(conflicts.keys(), key=lambda x: x.lower()):
        options = sorted(list(conflicts[skill]))
        choice = st.selectbox(
            f"Resolve category for skill: {skill}",
            options=options,
            key=f"conflict_{skill}",
        )
        category_map[skill] = choice
    st.warning("Resolve all conflicts, then rerun.")
    st.stop()
    return category_map

def categorize_skills_with_user(skills, category_map, resume_label=""):
    # IMPORTANT: skills list here is already normalized by process_skills()
    unknown = [s for s in skills if s not in category_map]
    if unknown:
        st.warning(f"Uncategorized skills found in {resume_label}. Assign each to one category.")
        placeholder = "-- Select category --"
        options = [placeholder] + SKILL_CATEGORY_COLS

        incomplete = False
        for s in unknown:
            choice = st.selectbox(
                f"Select category for skill: {s}",
                options=options,
                index=0,
                key=f"cat_{resume_label}_{s}",
            )
            if choice == placeholder:
                incomplete = True
            else:
                category_map[s] = choice

        if incomplete:
            st.warning("Please select a category for every uncategorized skill.")
            st.stop()

    categorized = {c: [] for c in SKILL_CATEGORY_COLS}
    for s in skills:
        categorized[category_map[s]].append(s)

    for c in categorized:
        categorized[c] = sorted(categorized[c], key=lambda x: x.lower())

    return categorized, category_map

# ============================================================
# Education + Years of Experience (Rule 8/18)
# ============================================================

def extract_degree_lines(edu_lines):
    rx = re.compile(r"\b(Ph\.?D|PhD|Doctor|M\.?S|MBA|Master|B\.?S|B\.?A|BA|Bachelor|A\.?S|A\.?A|AS|Associate)\b", re.I)
    out = []
    for line in edu_lines:
        if rx.search(line):
            out.append(line.strip())
    return dedupe_preserve_order(out)

def classify_degree(deg):
    d = (deg or "").lower()
    if "ph.d" in d or "phd" in d or "doctor" in d:
        return "Degree/Phds"
    if "m.s" in d or "ms" in d or "mba" in d or "master" in d:
        return "Degree/Masters"
    if "b.s" in d or "b.a" in d or "bachelor" in d:
        return "Degree/Bachelors"
    if "a.s" in d or "a.a" in d or "associate" in d:
        return "Degree/Associates"
    return "Degree/Bachelors"

def apply_degrees(edu_lines, edu_text):
    degrees_by_col = {
        "Degree/Associates": [],
        "Degree/Bachelors": [],
        "Degree/Masters": [],
        "Degree/Phds": [],
    }

    for deg in extract_degree_lines(edu_lines):
        key = _degree_lookup_key(deg)
        norm = DEGREE_NORMALIZATION.get(key)

        if norm is None:
            cleaned = _degree_generic_cleanup(deg)
            key2 = _degree_lookup_key(cleaned)
            norm = DEGREE_NORMALIZATION.get(key2, cleaned)

        col = classify_degree(norm)
        degrees_by_col[col].append(norm)

    years = [int(y) for y in re.findall(r"\b(19\d{2}|20\d{2})\b", edu_text or "")]
    earliest = min(years) if years else None
    return degrees_by_col, earliest

def compute_years_experience(earliest_degree_year):
    if earliest_degree_year is None:
        return None
    return FIXED_EXPERIENCE_YEAR - earliest_degree_year

# ============================================================
# Upsert row (Rule 5)
# ============================================================

def upsert_candidate_row(by_cat, name, skills_by_category, certs_raw, degrees_by_col, earliest_degree_year, action):
    by_cat = ensure_by_category_columns(by_cat)
    name_norm = (name or "").strip()
    exists_mask = by_cat["Name"].astype(str).str.lower() == name_norm.lower()

    if action == "replace" and exists_mask.any():
        by_cat = by_cat.loc[~exists_mask].copy()

    row = {c: "" for c in REQUIRED_COLUMNS}
    row["Name"] = name_norm
    row["Years of Experience"] = compute_years_experience(earliest_degree_year)

    # ✅ Stored skills are always normalized tokens (fix)
    for col in SKILL_CATEGORY_COLS:
        row[col] = ", ".join(skills_by_category.get(col, []))

    row["Certifications"] = "; ".join(dedupe_preserve_order(certs_raw))

    for col in ["Degree/Associates", "Degree/Bachelors", "Degree/Masters", "Degree/Phds"]:
        vals = degrees_by_col.get(col, [])
        row[col] = "; ".join(dedupe_preserve_order(vals))

    return pd.concat([by_cat, pd.DataFrame([row])], ignore_index=True)

# ============================================================
# SkillFrequency (Rule 14) — rebuilt from normalized tokens
# ============================================================

def _skill_freq_key(s):
    s = unicodedata.normalize("NFKC", (s or "").strip()).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def rebuild_skill_frequency(by_cat):
    by_cat = ensure_by_category_columns(by_cat)
    counts = Counter()
    label_votes = defaultdict(Counter)

    for _, r in by_cat.iterrows():
        per_candidate = set()
        for col in SKILL_CATEGORY_COLS:
            for s in parse_and_normalize_skills_from_cell(r.get(col, "")):
                k = _skill_freq_key(s)
                if k:
                    per_candidate.add(k)
                    label_votes[k][s] += 1
        for k in per_candidate:
            counts[k] += 1

    key_to_label = {}
    for k, votes in label_votes.items():
        best = sorted(votes.items(), key=lambda kv: (-kv[1], kv[0].casefold(), kv[0]))[0][0]
        key_to_label[k] = best

    rows = []
    for k in sorted(counts.keys(), key=lambda x: key_to_label.get(x, x).casefold()):
        rows.append({"Skill": key_to_label.get(k, k), "Candidate Count": int(counts[k])})

    return pd.DataFrame(rows, columns=["Skill", "Candidate Count"])

# ============================================================
# CertificationFrequency (Rule 17 + Rule 16 + cosmetic merge)
# ============================================================

def _cert_key(s):
    s = unicodedata.normalize("NFKC", (s or "").strip()).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def _normalize_cert_for_frequency(cert):
    raw = (cert or "").strip()
    if not raw:
        return ""
    k = _cert_key(raw)
    return CERT_NORMALIZATION.get(k, CERT_NORMALIZATION.get(raw.lower().strip(), raw))

def rebuild_cert_frequency(by_cat):
    by_cat = ensure_by_category_columns(by_cat)
    counts = Counter()
    label_votes = defaultdict(Counter)

    for _, r in by_cat.iterrows():
        cell = r.get("Certifications", "")
        if cell is None or (isinstance(cell, float) and pd.isna(cell)) or not str(cell).strip():
            continue

        certs = [c.strip() for c in _SPLIT.split(str(cell)) if c.strip()]
        per_candidate = set()

        for c in certs:
            norm = _normalize_cert_for_frequency(c)
            if not norm:
                continue
            k = _cert_key(norm)
            per_candidate.add(k)
            label_votes[k][norm] += 1

        for k in per_candidate:
            counts[k] += 1

    key_to_label = {}
    for k, votes in label_votes.items():
        best = sorted(votes.items(), key=lambda kv: (-kv[1], kv[0].casefold(), kv[0]))[0][0]
        key_to_label[k] = best

    rows = []
    for k in sorted(counts.keys(), key=lambda x: key_to_label.get(x, x).casefold()):
        rows.append({"Certification": key_to_label.get(k, k), "Candidate Count": int(counts[k])})

    return pd.DataFrame(rows, columns=["Certification", "Candidate Count"])

# ============================================================
# DegreeFrequency (Rule 19)
# ============================================================

def rebuild_degree_frequency(by_cat, col):
    by_cat = ensure_by_category_columns(by_cat)
    counts = Counter()

    for _, r in by_cat.iterrows():
        cell = r.get(col, "")
        if cell is None or (isinstance(cell, float) and pd.isna(cell)) or not str(cell).strip():
            continue

        degrees = [d.strip() for d in str(cell).split(";") if d.strip()]
        for d in set(degrees):
            counts[d] += 1

    return pd.DataFrame(
        sorted(counts.items(), key=lambda x: x[0].lower()),
        columns=["Degree", "Candidate Count"]
    )

# ============================================================
# Excel I/O
# ============================================================

def write_excel_output(by_cat, skill_freq, cert_freq, deg_assoc, deg_bach, deg_mast, deg_phd):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        ensure_by_category_columns(by_cat).to_excel(writer, sheet_name="By Category", index=False)
        skill_freq.to_excel(writer, sheet_name="SkillFrequency", index=False)
        cert_freq.to_excel(writer, sheet_name="CertificationFrequency", index=False)
        deg_assoc.to_excel(writer, sheet_name="DegreeFrequency_Associates", index=False)
        deg_bach.to_excel(writer, sheet_name="DegreeFrequency_Bachelors", index=False)
        deg_mast.to_excel(writer, sheet_name="DegreeFrequency_Masters", index=False)
        deg_phd.to_excel(writer, sheet_name="DegreeFrequency_Phds", index=False)
    return out.getvalue()

def load_excel_by_category(excel_file):
    xls = pd.ExcelFile(excel_file, engine="openpyxl")
    df = xls.parse(BY_CATEGORY_SHEET)
    return ensure_by_category_columns(df)