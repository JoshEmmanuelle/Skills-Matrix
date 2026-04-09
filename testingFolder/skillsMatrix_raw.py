# -------------------------------------------------------------
# FINAL PRODUCTION SKILL-MATRIX GENERATOR (STREAMLIT AGENT)
# Detailed UI (Option B)
# -------------------------------------------------------------
# - Drag & drop upload (DOCX, ZIP)
# - Extract SKILLS, CERTIFICATIONS, EDUCATION
# - Dynamic skill classification via keyword rules
# - Normalization + alphabetical sorting
# - Follows all rules from final-matrix-with-all-frequencies.xlsx
# - Produces 7-sheet Excel output
# -------------------------------------------------------------

import os
import re
import zipfile
import tempfile
import pandas as pd
from collections import defaultdict
from docx import Document
import streamlit as st

# -------------------------------------------------------------
# Allowed section identifiers
# -------------------------------------------------------------
SKILL_SECTIONS = [
    "skills", "technical skills", "technical summary",
    "toolset", "tools", "technology stack"
]

CERT_SECTIONS = [
    "certifications", "certification", "professional certifications"
]

EDU_SECTIONS = [
    "education", "academic background", "educational background",
    "degrees"
]

# -------------------------------------------------------------
# Keyword-based dynamic skill categorization
# -------------------------------------------------------------
CATEGORY_KEYWORDS = {
    "Language": [
        "python", "java", "c++", "c#", "c ",
        "bash", "perl", "groovy",
        "javascript", "typescript", "html", "css",
        "scala", "go", "rust"
    ],
    "Cloud": ["aws", "azure", "gcp", "cloud"],
    "Databases": ["mysql", "postgres", "sql server", "sql", "mongodb", "oracle", "elasticsearch", "accumulo"],
    "OS": ["linux", "windows", "unix", "centos", "rhel", "ubuntu", "kali", "redhat"],
    "Framework/Libraries": [
        "react", "angular", "flask", "django",
        "spring", "bootstrap",
        "pytorch", "tensorflow", "keras"
    ],
    "DevOps": [
        "jenkins", "ansible", "terraform",
        "cicd", "ci/cd", "argo"
    ],
    "Container/Orchestration": [
        "docker", "kubernetes", "compose", "swarm"
    ],
    "Machine Learning/AI": [
        "machine learning", "deep learning",
        "tensorflow", "keras", "pytorch",
        "ml", "ai"
    ],
    "Networking": [
        "tcp", "dns", "dhcp", "ssl", "routing", "switching", "network"
    ],
    "Version Control": ["git", "svn", "subversion"],
    "Tools": [
        "jira", "jupyter", "idapro", "ghidra",
        "autopsy", "splunk", "nifi",
        "vmware", "eclipse", "soapui",
        "postman", "visual studio", "vs code", "vscode"
    ],
    "Other": []
}

# -------------------------------------------------------------
# Normalization rules
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
# Clean + split + dedupe + normalize + alphabetize
# -------------------------------------------------------------
def clean_list(text):
    if not text or str(text).strip() == "":
        return ""
    parts = re.split(r",|;", text)
    items = []
    for p in parts:
        v = p.strip()
        if not v:
            continue
        v = normalize(v)
        items.append(v)
    items = list(dict.fromkeys(items))  # dedupe
    items = sorted(items, key=str.lower)
    return ", ".join(items)

# -------------------------------------------------------------
# Extract a section based on headers
# -------------------------------------------------------------
def extract_section(doc, header_list):
    lines = [p.text.strip() for p in doc.paragraphs]
    capture = False
    block = []

    for line in lines:
        low = line.lower()
        if any(low.startswith(h) for h in header_list):
            capture = True
            continue

        if capture:
            if not line.strip() or re.match(r"^[A-Za-z ]+:$", line):
                break
            block.append(line)
    return block

# -------------------------------------------------------------
# Find .docx files in uploads (supports ZIPs)
# -------------------------------------------------------------
def gather_docx(uploaded_files):
    temp_dir = tempfile.mkdtemp()
    collected = []

    for uf in uploaded_files:
        path = os.path.join(temp_dir, uf.name)
        with open(path, "wb") as f:
            f.write(uf.getbuffer())

        if path.lower().endswith(".docx"):
            collected.append(path)

        elif path.lower().endswith(".zip"):
            with zipfile.ZipFile(path, "r") as z:
                z.extractall(temp_dir)
            for root, _, files in os.walk(temp_dir):
                for name in files:
                    if name.lower().endswith(".docx"):
                        collected.append(os.path.join(root, name))

    return collected

# -------------------------------------------------------------
# Dynamic skill classification using keyword matching
# -------------------------------------------------------------
def classify_skills(raw_lines):
    categorized = {c: [] for c in CATEGORY_KEYWORDS}

    for line in raw_lines:
        parts = clean_list(line).split(", ")
        for skill in parts:
            if not skill:
                continue

            skill_lower = skill.lower()
            matched = False

            for category, keywords in CATEGORY_KEYWORDS.items():
                if any(kw in skill_lower for kw in keywords):
                    categorized[category].append(skill)
                    matched = True
                    break

            if not matched:
                categorized["Other"].append(skill)

    for c in categorized:
        categorized[c] = clean_list(", ".join(categorized[c]))

    return categorized

# -------------------------------------------------------------
# Parse a single resume
# -------------------------------------------------------------
def parse_resume(path):
    doc = Document(path)
    name = os.path.splitext(os.path.basename(path))[0]

    skill_lines = extract_section(doc, SKILL_SECTIONS)
    cert_lines  = extract_section(doc, CERT_SECTIONS)
    edu_lines   = extract_section(doc, EDU_SECTIONS)

    skills = classify_skills(skill_lines)
    certifications = clean_list("; ".join(cert_lines))

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

    return {
        "Name": name,
        **skills,
        "Certifications": certifications,
        "Degree/Associates": clean_list(assoc),
        "Degree/Bachelors": clean_list(bach),
        "Degree/Masters": clean_list(mast),
        "Degree/Phd's": clean_list(phd)
    }

# -------------------------------------------------------------
# Build frequency sheets
# -------------------------------------------------------------
def build_frequencies(df):
    skill_freq = defaultdict(int)
    cert_freq  = defaultdict(int)
    deg_freq = {
        "Degree/Associates": defaultdict(int),
        "Degree/Bachelors": defaultdict(int),
        "Degree/Masters": defaultdict(int),
        "Degree/Phd's": defaultdict(int)
    }

    for _, row in df.iterrows():
        for cat in CATEGORY_KEYWORDS:
            vals = clean_list(row[cat]).split(", ")
            for v in set([x for x in vals if x]):
                skill_freq[v] += 1

        certs = clean_list(row["Certifications"]).split(", ")
        for c in set([x for x in certs if x]):
            cert_freq[c] += 1

        for col in deg_freq:
            degs = clean_list(row[col]).split(", ")
            for d in set([x for x in degs if x]):
                deg_freq[col][d] += 1

    sf = pd.DataFrame(
        sorted(skill_freq.items(), key=lambda x: (-x[1], x[0].lower())),
        columns=["Skill", "Candidate Count"]
    )

    cf = pd.DataFrame(
        sorted(cert_freq.items(), key=lambda x: (-x[1], x[0].lower())),
        columns=["Certification", "Candidate Count"]
    )

    deg_tables = {}
    for col in deg_freq:
        deg_tables[col] = pd.DataFrame(
            sorted(deg_freq[col].items(), key=lambda x: (-x[1], x[0].lower())),
            columns=["Degree", "Candidate Count"]
        )

    return sf, cf, deg_tables

# -------------------------------------------------------------
# STREAMLIT UI (Detailed)
# -------------------------------------------------------------
def main():
    st.title("📊 Skill Matrix Agent")
    st.write("Automatically extract skills, certifications, and education from resumes.")

    st.subheader("📘 How to Use")
    st.markdown("""
    **1. Drag & drop your resumes** (.docx or .zip) into the upload box below.  
    **2. Click Process.**  
    **3. Download** your generated Excel file with:  
        - By Category sheet  
        - SkillFrequency  
        - CertificationFrequency  
        - DegreeFrequency sheets  
    """)

    uploads = st.file_uploader(
        "Drag and drop resumes here or click to browse",
        type=["docx", "zip"],
        accept_multiple_files=True
    )

    if uploads:
        st.info(f"{len(uploads)} file(s) selected.")
        for up in uploads:
            st.write("•", up.name)

    if st.button("Process Resumes"):
        if not uploads:
            st.error("Please upload at least one file.")
            return

        st.info("Extracting .docx files...")
        files = gather_docx(uploads)

        if not files:
            st.error("No .docx files found in uploads.")
            return

        st.success(f"{len(files)} resumes found.")

        rows = []
        for f in files:
            try:
                rows.append(parse_resume(f))
            except Exception as e:
                st.error(f"Error reading {f}: {str(e)}")

        df = pd.DataFrame(rows)

        st.info("Building frequency tables...")
        sf, cf, deg = build_frequencies(df)

        outpath = os.path.join(tempfile.mkdtemp(), "final-matrix-with-all-frequencies.xlsx")
        with pd.ExcelWriter(outpath, engine="openpyxl") as w:
            df.to_excel(w, "By Category", index=False)
            sf.to_excel(w, "SkillFrequency", index=False)
            cf.to_excel(w, "CertificationFrequency", index=False)
            deg["Degree/Associates"].to_excel(w, "DegreeFrequency_Associates", index=False)
            deg["Degree/Bachelors"].to_excel(w, "DegreeFrequency_Bachelors", index=False)
            deg["Degree/Masters"].to_excel(w, "DegreeFrequency_Masters", index=False)
            deg["Degree/Phd's"].to_excel(w, "DegreeFrequency_Phds", index=False)

        st.success("Your Skill Matrix is ready!")
        st.download_button(
            "Download final Excel file",
            data=open(outpath, "rb").read(),
            file_name="final-matrix-with-all-frequencies.xlsx"
        )


if __name__ == "__main__":
    main()