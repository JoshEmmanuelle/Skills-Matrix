# -------------------------------------------------------------
# FINAL PRODUCTION SKILL-MATRIX GENERATOR (OPTION 2: DYNAMIC)
# Streamlit App (Option B UI)
# -------------------------------------------------------------
# - Supports .docx and .zip resume uploads
# - Extracts SKILLS, CERTIFICATIONS, EDUCATION from correct sections
# - Dynamically classifies skills using keyword-based matching
# - Allows new skills to appear naturally from future resumes
# - Cleans, dedupes, sorts alphabetically, and normalizes values
# - Produces 7-sheet Excel like final-matrix-test-KS.xlsx
# -------------------------------------------------------------

import os
import zipfile
import tempfile
import re
import pandas as pd
from collections import defaultdict
from docx import Document
import streamlit as st

# -------------------------------------------------------------
# SECTION HEADERS
# -------------------------------------------------------------
SKILL_SECTIONS = ["skills", "technical skills", "tools", "technical summary", "technology stack"]
CERT_SECTIONS = ["certifications", "certification", "professional certifications"]
EDU_SECTIONS = ["education", "academic background", "degrees"]

# -------------------------------------------------------------
# KEYWORD-BASED CATEGORY CLASSIFICATION
# (Dynamic classification avoids hard-coded confidential data)
# -------------------------------------------------------------
CATEGORY_KEYWORDS = {
    "Language": [
        "python", "java", "c++", "c#", "c ", "bash", "perl", "groovy",
        "javascript", "typescript", "html", "css", "scala", "go", "rust"
    ],
    "Cloud": ["aws", "azure", "gcp", "cloud"],
    "Databases": ["mysql", "postgres", "sql", "mongodb", "oracle", "elasticsearch", "accumulo"],
    "OS": ["linux", "windows", "unix", "centos", "rhel", "ubuntu", "kali", "redhat"],
    "Framework/Libraries": ["react", "angular", "flask", "django", "spring", "bootstrap", "pytorch", "tensorflow", "keras"],
    "DevOps": ["jenkins", "ansible", "terraform", "cicd", "ci/cd", "argo"],
    "Container/Orchestration": ["docker", "kubernetes", "compose", "swarm"],
    "Machine Learning/AI": ["machine learning", "deep learning", "tensorflow", "keras", "pytorch", "ml", "ai"],
    "Networking": ["tcp", "dns", "dhcp", "ssl", "routing", "switching", "network"],
    "Version Control": ["git", "svn", "subversion"],
    "Tools": [
        "jira", "jupyter", "idapro", "ghidra", "autopsy", "splunk", "nifi", "vmware",
        "eclipse", "soapui", "postman", "visual studio", "vs code", "vscode"
    ],
    "Other": []  # fallback
}

# -------------------------------------------------------------
# NORMALIZATION MAP (Dynamic-based, no confidential data)
# -------------------------------------------------------------
NORMALIZE_MAP = {
    "VS Code": "VSCode",
    "UNIX": "Unix",
    "LINUX": "Linux",
    "CentOs": "CentOS",
    "PIG": "Pig",
    "IDA Pro": "IDAPro",
    "GIT": "Git",
    "Docker Compose": "Docker",
    "Docker Swarm": "Docker"
}

def normalize(v: str) -> str:
    v = v.strip()
    return NORMALIZE_MAP.get(v, v)

# -------------------------------------------------------------
# CLEANING UTILITIES
# -------------------------------------------------------------
def clean_list(text):
    if not text or str(text).strip() == "":
        return ""
    items = re.split(r",|;", text)
    cleaned = []
    for i in items:
        v = i.strip()
        if not v:
            continue
        v = normalize(v)
        cleaned.append(v)
    cleaned = list(dict.fromkeys(cleaned))        # dedupe
    cleaned = sorted(cleaned, key=str.lower)      # alphabetical
    return ", ".join(cleaned)

# -------------------------------------------------------------
# EXTRACT A SECTION FROM DOCX BY HEADERS
# -------------------------------------------------------------
def extract_section(doc, headers):
    lines = [p.text.strip() for p in doc.paragraphs]
    capture = False
    block = []

    for line in lines:
        low = line.lower()
        if any(low.startswith(h) for h in headers):
            capture = True
            continue

        if capture:
            if not line.strip() or re.match(r"^[A-Za-z ]+:$", line):
                break
            block.append(line)
    return block

# -------------------------------------------------------------
# FIND DOCX FILES (DIRECT, ZIP, FOLDER)
# -------------------------------------------------------------
def gather_docx(uploaded_files):
    temp_dir = tempfile.mkdtemp()
    found = []

    for file in uploaded_files:
        path = os.path.join(temp_dir, file.name)
        with open(path, "wb") as f:
            f.write(file.getbuffer())

        if path.lower().endswith(".docx"):
            found.append(path)
        elif path.lower().endswith(".zip"):
            with zipfile.ZipFile(path, "r") as z:
                z.extractall(temp_dir)
            for root, _, files in os.walk(temp_dir):
                for fn in files:
                    if fn.lower().endswith(".docx"):
                        found.append(os.path.join(root, fn))
    return found

# -------------------------------------------------------------
# DYNAMIC SKILL CLASSIFICATION
# -------------------------------------------------------------
def classify_skills(raw_skills):
    categorized = {c: [] for c in CATEGORY_KEYWORDS}

    for line in raw_skills:
        items = clean_list(line).split(", ")
        for skill in items:
            if not skill:
                continue

            skill_lower = skill.lower()
            matched = False

            # Keyword-based matching
            for category, keywords in CATEGORY_KEYWORDS.items():
                if any(kw in skill_lower for kw in keywords):
                    categorized[category].append(skill)
                    matched = True
                    break

            if not matched:
                categorized["Other"].append(skill)

    # Clean final results
    for c in categorized:
        categorized[c] = clean_list(", ".join(categorized[c]))

    return categorized

# -------------------------------------------------------------
# PARSE A SINGLE RESUME
# -------------------------------------------------------------
def parse_resume(path):
    doc = Document(path)
    name = os.path.splitext(os.path.basename(path))[0]

    skill_lines = extract_section(doc, SKILL_SECTIONS)
    cert_lines = extract_section(doc, CERT_SECTIONS)
    edu_lines = extract_section(doc, EDU_SECTIONS)

    skills = classify_skills(skill_lines)
    certs = clean_list("; ".join(cert_lines))

    assoc = bach = mast = phd = ""

    for line in edu_lines:
        L = line.lower()
        if "associate" in L:
            assoc = line
        elif "bachelor" in L or "b.s." in L:
            bach = line
        elif "master" in L or "m.s." in L or "mba" in L:
            mast = line
        elif "phd" in L or "doctor" in L:
            phd = line

    assoc = clean_list(assoc)
    bach = clean_list(bach)
    mast = clean_list(mast)
    phd = clean_list(phd)

    return {
        "Name": name,
        **skills,
        "Certifications": certs,
        "Degree/Associates": assoc,
        "Degree/Bachelors": bach,
        "Degree/Masters": mast,
        "Degree/Phd's": phd
    }

# -------------------------------------------------------------
# BUILD FREQUENCY SHEETS
# -------------------------------------------------------------
def build_frequencies(df):
    skill_freq = defaultdict(int)
    cert_freq = defaultdict(int)
    deg_freq = {
        "Degree/Associates": defaultdict(int),
        "Degree/Bachelors": defaultdict(int),
        "Degree/Masters": defaultdict(int),
        "Degree/Phd's": defaultdict(int)
    }

    for _, row in df.iterrows():
        # Skills
        for cat in CATEGORY_KEYWORDS:
            vals = clean_list(row[cat]).split(", ")
            for v in set([x for x in vals if x]):
                skill_freq[v] += 1

        # Certifications
        for c in set([x for x in clean_list(row["Certifications"]).split(", ") if x]):
            cert_freq[c] += 1

        # Degrees
        for col in deg_freq:
            for d in set([x for x in clean_list(row[col]).split(", ") if x]):
                deg_freq[col][d] += 1

    sf = pd.DataFrame(sorted(skill_freq.items(), key=lambda x: (-x[1], x[0].lower())),
                      columns=["Skill", "Candidate Count"])
    cf = pd.DataFrame(sorted(cert_freq.items(), key=lambda x: (-x[1], x[0].lower())),
                      columns=["Certification", "Candidate Count"])

    deg_tables = {}
    for col in deg_freq:
        deg_tables[col] = pd.DataFrame(
            sorted(deg_freq[col].items(), key=lambda x: (-x[1], x[0].lower())),
            columns=["Degree", "Candidate Count"]
        )

    return sf, cf, deg_tables

# -------------------------------------------------------------
# STREAMLIT UI — Option B
# -------------------------------------------------------------
def main():
    st.title("📊 Skill Matrix Generator (Dynamic Classification)")
    st.write("Upload resumes (.docx or .zip). The system will extract, classify, and generate a multi-sheet Excel file.")

    uploads = st.file_uploader("Upload files", type=["docx", "zip"], accept_multiple_files=True)

    if st.button("Process") and uploads:
        st.info("Scanning files...")
        files = gather_docx(uploads)

        if not files:
            st.error("No valid .docx files found.")
            return

        st.success(f"{len(files)} resumes detected.")

        rows = []
        for f in files:
            try:
                rows.append(parse_resume(f))
            except Exception as e:
                st.error(f"Error parsing {f}: {str(e)}")

        df = pd.DataFrame(rows)
        st.success("Extraction complete.")

        sf, cf, deg_tables = build_frequencies(df)

        outpath = os.path.join(tempfile.mkdtemp(), "final-matrix-with-all-frequencies.xlsx")
        with pd.ExcelWriter(outpath, engine="openpyxl") as w:
            df.to_excel(w, "By Category", index=False)
            sf.to_excel(w, "SkillFrequency", index=False)
            cf.to_excel(w, "CertificationFrequency", index=False)
            deg_tables["Degree/Associates"].to_excel(w, "DegreeFrequency_Associates", index=False)
            deg_tables["Degree/Bachelors"].to_excel(w, "DegreeFrequency_Bachelors", index=False)
            deg_tables["Degree/Masters"].to_excel(w, "DegreeFrequency_Masters", index=False)
            deg_tables["Degree/Phd's"].to_excel(w, "DegreeFrequency_Phds", index=False)

        st.success("Your Excel file is ready.")
        st.download_button(
            "Download final-matrix-with-all-frequencies.xlsx",
            data=open(outpath, "rb").read(),
            file_name="final-matrix-with-all-frequencies.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()