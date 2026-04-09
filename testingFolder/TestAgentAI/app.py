import io
import json
import re
import zipfile
from typing import Dict, List, Optional, Set

import pandas as pd
import streamlit as st
import requests
from docx import Document

# ======================================================
# CONFIG
# ======================================================
LM_STUDIO_URL = "http://127.0.0.1:1234/v1/chat/completions"
MODEL_NAME = "meta-llama_-_meta-llama-3-8b-instruct"
CURRENT_YEAR = 2026
MAPPING_FILE = "skill_category_mappings.json"

# ======================================================
# STRUCTURE
# ======================================================
CATEGORIES = [
    "Language", "Cloud", "Databases", "OS",
    "Framework/Libraries", "DevOps",
    "Container/Orchestration", "Machine Learning/AI",
    "Networking", "Version Control", "Tools", "Other",
]

BY_CATEGORY_COLUMNS = [
    "Name", "Years of Experience",
    *CATEGORIES,
    "Certifications",
    "Degree/Associates",
    "Degree/Bachelors",
    "Degree/Masters",
    "Degree/Phds",
]

# ======================================================
# SYSTEM PROMPT (EXTRACTION ONLY)
# ======================================================
SYSTEM_PROMPT = """
Return STRICT JSON only. No markdown, no commentary.

Keys:
- name
- skills_section_text
- education_section_text
- certifications_section_text
- experience_section_text

Extract ONLY from explicit section headers.
Return empty strings if a section does not exist.
""".strip()

# ======================================================
# SAFE HELPERS
# ======================================================
def split_items(text) -> List[str]:
    if text is None or isinstance(text, float):
        return []
    if not isinstance(text, str):
        text = str(text)
    return [x.strip() for x in re.split(r"[;,]", text) if x.strip()]

def dedupe(seq: List[str]) -> List[str]:
    seen = set()
    out = []
    for s in seq:
        if s not in seen:
            seen.add(s)
            out.append(s)
    return out

def normalize_skill(skill) -> Optional[str]:
    if not isinstance(skill, str):
        return None
    s = skill.strip()
    return s if s else None

# ======================================================
# EDUCATION HELPERS (NEW)
# ======================================================
def split_education_lines(text: str) -> List[str]:
    if not text or not isinstance(text, str):
        return []
    lines = []
    for line in text.splitlines():
        for part in line.split(";"):
            p = part.strip()
            if p:
                lines.append(p)
    return lines

def classify_degree_level(deg: str) -> Optional[str]:
    t = deg.lower()
    if any(k in t for k in ["ph.d", "phd", "doctor"]):
        return "Phds"
    if any(k in t for k in ["master", "m.s", "ms", "mba"]):
        return "Masters"
    if any(k in t for k in ["bachelor", "b.s", "bs", "b.a", "ba"]):
        return "Bachelors"
    if any(k in t for k in ["associate", "a.s", "as", "a.a"]):
        return "Associates"
    return None

# ======================================================
# SKILL MAP PERSISTENCE
# ======================================================
def load_skill_map() -> Dict[str, str]:
    try:
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_skill_map(m: Dict[str, str]) -> None:
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(m, f, indent=2, sort_keys=True)

# ======================================================
# EXCEL SKILL MAP (UNAMBIGUOUS ONLY)
# ======================================================
def build_excel_skill_map(df: pd.DataFrame) -> Dict[str, str]:
    found: Dict[str, Set[str]] = {}
    for cat in CATEGORIES:
        for cell in df[cat].dropna():
            for raw in split_items(cell):
                sk = normalize_skill(raw)
                if sk:
                    found.setdefault(sk, set()).add(cat)
    return {k: next(iter(v)) for k, v in found.items() if len(v) == 1}

# ======================================================
# LLM EXTRACTION
# ======================================================
def extract_resume(text: str) -> Dict[str, str]:
    payload = {
        "model": MODEL_NAME,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": text},
        ],
        "max_tokens": 1400,
    }
    r = requests.post(LM_STUDIO_URL, json=payload, timeout=120)
    r.raise_for_status()
    content = r.json()["choices"][0]["message"]["content"]
    content = re.sub(r"```(?:json)?", "", content, flags=re.IGNORECASE).strip()

    m = re.search(r"\{.*\}", content, flags=re.DOTALL)
    if not m:
        raise ValueError("LLM did not return JSON")

    obj = json.loads(m.group(0))
    required = {
        "name",
        "skills_section_text",
        "education_section_text",
        "certifications_section_text",
        "experience_section_text",
    }
    missing = required - obj.keys()
    if missing:
        raise ValueError(f"LLM JSON missing keys: {missing}")

    return obj

# ======================================================
# YEARS OF EXPERIENCE
# ======================================================
def calc_years_experience(exp_text: str) -> Optional[int]:
    years = [int(y) for y in re.findall(r"(19\\d{2}|20\\d{2})", exp_text)]
    return CURRENT_YEAR - min(years) if years else None

# ======================================================
# REBUILD ALL FREQUENCY SHEETS
# ======================================================
def rebuild_frequencies(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    result = {}

    # SkillFrequency
    skill_counts: Dict[str, Set[str]] = {}
    for _, r in df.iterrows():
        for c in CATEGORIES:
            for sk in split_items(r[c]):
                skill_counts.setdefault(sk, set()).add(r["Name"])

    result["SkillFrequency"] = pd.DataFrame(
        [{"Skill": k, "Candidate Count": len(v)} for k, v in sorted(skill_counts.items())]
    )

    # CertificationFrequency
    cert_counts: Dict[str, Set[str]] = {}
    for _, r in df.iterrows():
        for c in split_items(r["Certifications"]):
            cert_counts.setdefault(c, set()).add(r["Name"])

    result["CertificationFrequency"] = pd.DataFrame(
        [{"Certification": k, "Candidate Count": len(v)} for k, v in sorted(cert_counts.items())]
    )

    # Degree Frequencies
    for level in ["Associates", "Bachelors", "Masters", "Phds"]:
        col = f"Degree/{level}"
        deg_counts: Dict[str, Set[str]] = {}
        for _, r in df.iterrows():
            for d in split_items(r[col]):
                deg_counts.setdefault(d, set()).add(r["Name"])
        result[f"DegreeFrequency_{level}"] = pd.DataFrame(
            [{"Degree": k, "Candidate Count": len(v)} for k, v in sorted(deg_counts.items())]
        )

    return result

# ======================================================
# STREAMLIT APP
# ======================================================
st.set_page_config(layout="wide")
st.title("✅ Skills Matrix Resume Processor")

excel_file = st.file_uploader("Upload Skills Matrix Excel", type="xlsx")
resume_files = st.file_uploader(
    "Upload Resumes (.docx or .zip)",
    type=["docx", "zip"],
    accept_multiple_files=True,
)

if st.button("🚀 Process"):
    if not excel_file or not resume_files:
        st.error("Excel file and resumes are required.")
        st.stop()

    df = pd.read_excel(excel_file, sheet_name="By Category")
    df = df.reindex(columns=BY_CATEGORY_COLUMNS)

    excel_map = build_excel_skill_map(df)
    persistent_map = load_skill_map()
    skill_map = {**excel_map, **persistent_map}

    progress = st.progress(0)
    added = updated = 0

    for i, uploaded in enumerate(resume_files):
        progress.progress((i + 1) / len(resume_files))

        docs = []
        if uploaded.name.lower().endswith(".zip"):
            z = zipfile.ZipFile(uploaded)
            docs = [(n, z.read(n)) for n in z.namelist() if n.lower().endswith(".docx")]
        else:
            docs = [(uploaded.name, uploaded.read())]

        for fname, data in docs:
            doc = Document(io.BytesIO(data))
            text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
            ext = extract_resume(text)

            name = ext["name"]
            duplicate = df["Name"].str.lower().eq(name.lower()).any()

            replace = False
            if duplicate:
                replace = (
                    st.radio(
                        f"Duplicate found: {name}",
                        ["New Person", "Replace Existing"],
                        key=f"dup_{name}_{fname}",
                    )
                    == "Replace Existing"
                )

            row = {c: "" for c in BY_CATEGORY_COLUMNS}
            row["Name"] = name
            row["Years of Experience"] = calc_years_experience(ext["experience_section_text"])
            row["Certifications"] = "; ".join(split_items(ext["certifications_section_text"]))

            # EDUCATION
            for deg in split_education_lines(ext["education_section_text"]):
                lvl = classify_degree_level(deg)
                if lvl:
                    col = f"Degree/{lvl}"
                    row[col] += (", " if row[col] else "") + deg

            # SKILLS
            skills = dedupe(
                filter(None, [normalize_skill(s) for s in split_items(ext["skills_section_text"])])
            )

            unknown = []
            for sk in skills:
                if sk in skill_map:
                    cat = skill_map[sk]
                    row[cat] += (", " if row[cat] else "") + sk
                else:
                    unknown.append(sk)

            for sk in unknown:
                cat = st.selectbox(f"Categorize skill '{sk}'", CATEGORIES, key=f"{name}_{sk}")
                persistent_map[sk] = cat
                row[cat] += (", " if row[cat] else "") + sk

            if unknown:
                save_skill_map(persistent_map)
                skill_map.update(persistent_map)

            if duplicate and replace:
                df = df[df["Name"].str.lower() != name.lower()]
                updated += 1
            else:
                added += 1

            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

    frequencies = rebuild_frequencies(df)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="By Category", index=False)
        for sheet, fdf in frequencies.items():
            fdf.to_excel(writer, sheet_name=sheet, index=False)

    st.success(f"✅ Done — Added: {added}, Updated: {updated}")
    st.download_button(
        "⬇️ Download Updated Excel",
        output.getvalue(),
        "Skills_Matrix_Updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )