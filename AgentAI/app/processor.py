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
    "Associations/Honors:"
    "Associations:",
    "Honors:"
    "Associations/Honors:"
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
    "certified scrum master (csm)": "Certified Scrum Master",
    "aws solutions architect associate": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect – associate": "AWS Certified Solutions Architect - Associate",
    "aws solutions certified - associate": "AWS Certified Solutions Architect - Associate",
    "ccna (cisco certified network associate)": "CCNA",
    "ccna certification": "CCNA",
    "itil v3 foundation": "ITIL v3.0",
    "security+": "CompTIA Security+",
    "security +": "CompTIA Security+",
    "security+ (comptia)": "CompTIA Security+",
    "security+ ce": "CompTIA Security+",
    "comptia security+ ce": "CompTIA Security+"
}

# ============================================================
# Rule 18: Degree normalization mapping (explicit only)
# ============================================================

DEGREE_NORMALIZATION = {
    "bachelor in business admin": "B.S., Business and Administration"
}

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
    certs_raw = dedupe_preserve_order(split_on_commas_semicolons(" ".join(cert_lines).strip()))

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
    rx = re.compile(r"\b(Ph\.?D|PhD|Doctor|M\.?S|MBA|Master|B\.?S|B\.?A|Bachelor|Associate)\b", re.I)
    out = []
    for line in edu_lines:
        if rx.search(line):
            out.append(line.strip())
    return dedupe_preserve_order(out)

def classify_degree(deg):
    d = deg.lower()
    if "ph.d" in d or "phd" in d or "doctor" in d:
        return "Degree/Phds"
    if "m.s" in d or "mba" in d or "master" in d:
        return "Degree/Masters"
    if "b.s" in d or "b.a" in d or "bachelor" in d:
        return "Degree/Bachelors"
    if "associate" in d or "a.s" in d:
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
        key = deg.strip().lower()
        norm = DEGREE_NORMALIZATION.get(key, deg.strip())
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