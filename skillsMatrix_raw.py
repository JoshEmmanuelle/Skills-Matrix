# -------------------------------------------------------------
# AUTOMATED RESUME SKILL–CERT–EDU MATRIX GENERATOR
# FINAL PRODUCTION VERSION FOR STREAMLIT DEPLOYMENT
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
# SECTION TITLES (YOUR EXACT RULES)
# -------------------------------------------------------------
SKILL_SECTIONS = [
    "skills", "technical skills", "tools", "technology stack"
]

CERT_SECTIONS = [
    "certification", "certifications"
]

EDU_SECTIONS = [
    "education", "academic background"
]

# -------------------------------------------------------------
# SKILL CATEGORY MAPPING (YOUR S1–S12 SYSTEM)
# -------------------------------------------------------------
CATEGORY_MAP = {
    "Language": [
        "python", "java", "c,", "c ", "c++", "bash", "perl", "groovy", 
        "javascript", "typescript", "vb", "visual basic", "scala", "rust"
    ],
    "Cloud": ["aws", "azure", "gcp"],
    "Databases": [
        "mysql", "postgres", "sql server", "mongodb", "oracle", 
        "elasticsearch", "accumulo", "postgresql", "maria"
    ],
    "OS": ["linux", "windows", "unix", "centos", "rhel", "ubuntu", "kali"],
    "Framework/Libraries": [
        "angular", "react", "flask", "django", "spring", 
        "bootstrap", "pytorch", "tensorflow"
    ],
    "DevOps": ["jenkins", "ansible", "terraform", "argo", "ci/cd", "cicd"],
    "Container/Orchestration": ["docker", "kubernetes", "compose", "swarm"],
    "Machine Learning/AI": [
        "tensorflow", "pytorch", "keras", "machine learning", 
        "ml", "deep learning"
    ],
    "Networking": ["tcp/ip", "dns", "dhcp", "ssl", "routing", "switching"],
    "Version Control": ["git", "svn", "subversion"],
    "Tools": [
        "jira", "jupyter", "ida pro", "idapro", "ghidra", "autopsy", 
        "splunk", "nifi", "vmware", "eclipse", "soapui", 
        "visual studio", "postman", "spectrum analyzer"
    ],
    "Other": []  # fallback category
}

# -------------------------------------------------------------
# NORMALIZATION RULES GIVEN BY YOU
# -------------------------------------------------------------
NORMALIZE = {
    "UNIX": "Unix",
    "LINUX": "Linux",
    "CentOs": "CentOS",
    "JIRA": "Jira",
    "VS Code": "VSCode",
    "REST Services": "REST",
    "Python3": "Python",
    "PIG": "Pig",
    "IDA Pro": "IDAPro",
    "GIT": "Git",
    "Docker Compose": "Docker",
    "Docker Swarm": "Docker",
    "Jupyter": "Jupyter Notebook",
    "Jupyter Notebooks": "Jupyter Notebook"
}

# -------------------------------------------------------------
# NORMALIZATION FUNCTION
# -------------------------------------------------------------
def normalize_skill(x):
    x = x.strip()
    return NORMALIZE.get(x, x)

# -------------------------------------------------------------
# EXTRACT SPECIFIC SECTIONS FROM DOCX
# -------------------------------------------------------------
def extract_section_text(doc, target_sections):
    lines = [p.text.strip() for p in doc.paragraphs]
    collected = []
    capture = False

    for line in lines:
        if any(line.lower().startswith(s) for s in target_sections):
            capture = True
            continue

        if capture:
            # Stop if next section heading or empty gap
            if line.strip() == "" or re.match(r"^[A-Za-z ]+:$", line):
                break
            collected.append(line)

    return collected

# -------------------------------------------------------------
# FIND ALL DOCX FILES FROM UPLOADED FILES / ZIP / FOLDER
# -------------------------------------------------------------
def gather_docx_files(uploaded_files):
    temp_dir = tempfile.mkdtemp()
    collected = []

    for uf in uploaded_files:
        filepath = os.path.join(temp_dir, uf.name)
        with open(filepath, "wb") as f:
            f.write(uf.getbuffer())

        if uf.name.lower().endswith(".docx"):
            collected.append(filepath)

        elif uf.name.lower().endswith(".zip"):
            with zipfile.ZipFile(filepath, "r") as z:
                z.extractall(temp_dir)

            for root, _, files in os.walk(temp_dir):
                for f in files:
                    if f.lower().endswith(".docx"):
                        collected.append(os.path.join(root, f))

    return collected

# -------------------------------------------------------------
# PARSE A SINGLE RESUME
# -------------------------------------------------------------
def parse_resume(filepath):
    doc = Document(filepath)
    name = os.path.splitext(os.path.basename(filepath))[0]

    raw_skills = extract_section_text(doc, SKILL_SECTIONS)
    raw_certs = extract_section_text(doc, CERT_SECTIONS)
    raw_edu = extract_section_text(doc, EDU_SECTIONS)

    # ---------- SKILL EXTRACTION ----------
    categorized = {c: [] for c in CATEGORY_MAP}
    for line in raw_skills:
        items = [
            normalize_skill(x) for x in re.split(r",|;", line) 
            if x.strip()
        ]
        for item in items:
            matched = False
            for category, keywords in CATEGORY_MAP.items():
                if any(k.lower() in item.lower() for k in keywords):
                    categorized[category].append(item)
                    matched = True
                    break
            if not matched:
                categorized["Other"].append(item)

    # Convert lists → clean strings
    for key in categorized:
        categorized[key] = ", ".join(dict.fromkeys(categorized[key]))

    # ---------- CERTIFICATIONS ----------
    certs = "; ".join([c.strip() for c in raw_certs]).strip()

    # ---------- EDUCATION ----------
    deg_assoc = ""
    deg_bach = ""
    deg_mast = ""
    deg_phd = ""

    for line in raw_edu:
        L = line.lower()
        if "associate" in L:
            deg_assoc = line
        elif "bachelor" in L or "b.s." in L:
            deg_bach = line
        elif "master" in L or "m.s." in L or "mba" in L:
            deg_mast = line
        elif "phd" in L or "doctor" in L:
            deg_phd = line

    return {
        "Name": name,
        **categorized,
        "Certifications": certs,
        "Degree/Associates": deg_assoc,
        "Degree/Bachelors": deg_bach,
        "Degree/Masters": deg_mast,
        "Degree/Phd's": deg_phd,
    }

# -------------------------------------------------------------
# BUILD FREQUENCY TABLES
# -------------------------------------------------------------
def build_frequency_tables(df):
    # Skills
    skill_count = defaultdict(int)
    for _, row in df.iterrows():
        seen = set()
        for col in CATEGORY_MAP.keys():
            vals = str(row[col]).split(",")
            for v in [x.strip() for x in vals if x.strip()]:
                seen.add(v)
        for item in seen:
            skill_count[item] += 1

    skill_df = pd.DataFrame(
        sorted(skill_count.items(), key=lambda x: (-x[1], x[0].lower())),
        columns=["Skill", "Candidate Count"]
    )

    # Certifications
    cert_count = defaultdict(int)
    for _, row in df.iterrows():
        if row["Certifications"]:
            for c in [x.strip() for x in row["Certifications"].split(";") if x.strip()]:
                cert_count[c] += 1

    cert_df = pd.DataFrame(
        sorted(cert_count.items(), key=lambda x: (-x[1], x[0].lower())),
        columns=["Certification", "Candidate Count"]
    )

    # Degrees
    degree_tables = {}
    for col in ["Degree/Associates", "Degree/Bachelors", "Degree/Masters", "Degree/Phd's"]:
        counts = defaultdict(int)
        for _, row in df.iterrows():
            if row[col]:
                for c in [x.strip() for x in row[col].split(";") if x.strip()]:
                    counts[c] += 1
        degree_tables[col] = pd.DataFrame(
            sorted(counts.items(), key=lambda x: (-x[1], x[0].lower())),
            columns=["Degree", "Candidate Count"]
        )

    return skill_df, cert_df, degree_tables

# -------------------------------------------------------------
# STREAMLIT APPLICATION UI
# -------------------------------------------------------------
def main():
    st.title("📊 Automated Skill Matrix Generator")
    st.write("Upload Word (.docx) files or ZIP folders containing resumes.")

    uploads = st.file_uploader(
        "Drag & Drop resumes or ZIP files here",
        type=["docx", "zip"],
        accept_multiple_files=True
    )

    if st.button("Process Resumes") and uploads:
        docx_files = gather_docx_files(uploads)

        if not docx_files:
            st.error("No .docx files found.")
            return

        st.info("Extracting resume data...")
        records = [parse_resume(f) for f in docx_files]
        df = pd.DataFrame(records)

        st.success("Extraction complete. Building frequency tables...")

        skill_df, cert_df, degree_tables = build_frequency_tables(df)

        # Save Excel
        output_path = os.path.join(
            tempfile.mkdtemp(), 
            "final-matrix-with-all-frequencies.xlsx"
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="By Category", index=False)
            skill_df.to_excel(writer, sheet_name="SkillFrequency", index=False)
            cert_df.to_excel(writer, sheet_name="CertificationFrequency", index=False)
            degree_tables["Degree/Associates"].to_excel(writer, sheet_name="DegreeFrequency_Associates", index=False)
            degree_tables["Degree/Bachelors"].to_excel(writer, sheet_name="DegreeFrequency_Bachelors", index=False)
            degree_tables["Degree/Masters"].to_excel(writer, sheet_name="DegreeFrequency_Masters", index=False)
            degree_tables["Degree/Phd's"].to_excel(writer, sheet_name="DegreeFrequency_Phds", index=False)

        st.success("Your final Skill Matrix Excel file is ready.")
        st.download_button(
            "Download Excel File",
            data=open(output_path, "rb").read(),
            file_name="final-matrix-with-all-frequencies.xlsx"
        )

if __name__ == "__main__":
    main()