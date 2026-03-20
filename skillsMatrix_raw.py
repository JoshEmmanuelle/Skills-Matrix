'''
The purpose of this notebook is to create an automated skills matrix for the team. 
This will help us identify the skills, certifications, and education of each team member.
'''

# Import the needed libraries

import os
import zipfile
import tempfile
import re
import pandas as pd
from collections import defaultdict
import streamlit as st
from docx import Document

# Define the sections you want to extract from the resumes

SKILL_SECTIONS = [
    "skills", "SKILLS", "Skills", "Technical Skills", "TECHNICAL SKILLS", "technical skills", "tools", "TOOLS", "Tools", "SKILLS/TOLLS/TECHNOLOGIES", "Skills/Tools/Technologies", "skills/tools/technologies"
]

CERT_SECTIONS = [
    "certifications", "CERTIFICATIONS", "Certifications", "CERTIFICATION", "Certification", "certification"
]

EDU_SECTIONS = [
    "education", "EDUCATION", "Education", "EDUCATIONAL BACKGROUND", "academic background", "ACADEMIC BACKGROUND", "Academic Background"
]



# Skill categories and keywords for mapping into S1–S12
CATEGORY_MAP = {
    "Language": ["python", "java", "c", "c++", "bash", "perl", "groovy", "javascript", "typescript", "vb", "visual basic", "scala"],
    "Cloud": ["aws", "azure", "gcp"],
    "Databases": ["mysql", "postgres", "sql server", "mongodb", "oracle", "elastic", "accumulo"],
    "OS": ["linux", "windows", "unix", "centos", "rhel", "ubuntu", "kali"],
    "Framework/Libraries": ["angular", "react", "flask", "django", "spring", "bootstrap"],
    "DevOps": ["jenkins", "ansible", "terraform", "argo", "cicd", "ci/cd"],
    "Container/Orchestration": ["docker", "kubernetes", "compose", "swarm"],
    "Machine Learning/AI": ["tensorflow", "keras", "pytorch", "machine learning"],
    "Networking": ["tcp/ip", "dns", "dhcp", "ssl", "routing", "switching"],
    "Version Control": ["git", "svn", "subversion"],
    "Tools": [
        "jira", "jupyter", "ida pro", "ghidra", "autopsy", "splunk", "nifi", "vmware",
        "eclipse", "soapui", "visual studio", "postman", "spectrum analyzer"
    ],
    "Other": []  # Fallback category. Probably depending on the results in this section, we may want to create additional categories or move some keywords around.
}

# Normalize skill keywords and text

def normalize_skill(x):
    x = x.strip()
    return NORMALIZE.get(x, x)

# Extract the text from .docx section

def extract_section_text(doc, target_sections):
    lines = [p.text.strip() for p in doc.paragraphs]
    collected = []
    capture = False

    for line in lines:
        if any(line.lower().startswith(s) for s in target_sections):
            capture = True
            continue
        if capture:
            if line.strip() == "" or re.match(r"^[A-Z\s]+:$", line):
                break
            collected.append(line)
    return collected

# Extract text from DOCX files, including DOCX files in folders or zip folders

def gather_docx_files(uploaded_files):
    temp_dir = tempfile.mkdtemp()
    all_docx = []

    for uf in uploaded_files:
        file_path = os.path.join(temp_dir, uf.name)
        with open(file_path, "wb") as f:
            f.write(uf.getbuffer())


        if uf.name.lower().endswith(".docx"):
            all_docx.append(file_path)

        elif uf.name.lower().endswith(".zip"):
            with zipfile.ZipFile(file_path, "r") as z:
                z.extractall(temp_dir)
                for root, _, files in os.walk(temp_dir):
                    for f in files:
                        if f.lower().endswith(".docx"):
                            all_docx.append(os.path.join(root, f))
        
        else:
            if os.path.isdir(file_path):
                for root, _, files in os.walk(file_path):
                    for f in files:
                        if f.lower().endswith(".docx"):
                            all_docx.append(os.path.join(root, f))
                
    return all_docx       


# Parse a single resume
def parse_resume(file_path):
    doc = Document(file_path)
    name = os.path.splitext(os.path.basename(file_path))[0]

    # Extract sections
    raw_skills = extract_section_text(doc, SKILL_SECTIONS)
    raw_certs = extract_section_text(doc, CERT_SECTIONS)
    raw_edu = extract_section_text(doc, EDU_SECTIONS)

    # Parse skills and categorize
    
    categorized = {c : [] for c in CATEGORY_MAP}
    for line in raw_skills:
        items = [normalize_skill(x) for x in re.split(r",|;", line) if x.strip()]
        for item in items:
            assigned = False
            for category, keywords in CATEGORY_MAP.items():
                if any(k.lower() in item.lower() for k in keywords):
                    categorized[category].append(item)
                    assigned = True
                    break
            if not assigned:
                categorized["Other"].append(item)

    #Parse Certifications
    certs = ";".join([x.strip() for x in raw_certs]).strip()

    # Parse Education
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
        "Degree/Associate": deg_assoc,
        "Degree/Bachelor": deg_bach,
        "Degree/Master": deg_mast,
        "Degree/PhD": deg_phd
    }


#Frequency computation

def build_frequency_tables(df):
    skill_freq = defaultdict(int)
    for _, row in df.iterrows():
        seen = set()
        for col in CATEGORY_MAP.keys():
            for item in [x.strip() for x in str(row).split(";") if x.strip()]:
                seen.add
        for s in seen:
            skill_freq[s] += 1

    skill_df = pd.DataFrame(
        sorted(skill_freq.items() , key = lambda x: (-x[1], x[0].lower())),
                            columns=["Skill", "Candidate Count"]
                            )

# Certification frequency
    cert_freq = defaultdict(int)
    for _, row in df.iterrows():
        raw = row["Certifications"]
        if raw:
            for item in [x.strip() for x in raw.split(";") if x.strip()]:
                cert_freq[item] += 1

    cert_df = pd.DataFrame(
        sorted(cert_freq.items() , key = lambda x: (-x[1], x[0].lower())),
                            columns=["Certification", "Candidate Count"]
                            )

 # Degree frequency
    degree_tables = {}
    for col in ["Degree/Associate", "Degree/Bachelor", "Degree/Master", "Degree/PhD"]:
        freq = defaultdict(int)
        for _, row in df.iterrows():
            raw = row[col]
            if raw:
                for item in [x.strip() for x in raw.split(";") if x.strip()]:
                    freq[item] += 1
        
        degree_tables[col] = pd.DataFrame(
            sorted(freq.items() , key = lambda x: (-x[1], x[0].lower())),
                                columns=["Degree", "Candidate Count"]
                                )
    return skill_df, cert_df, degree_tables


def main():
    st.title("Automated Resume Skills Matrix")
    st.write("Upload here the resumes needed for triage")

    uploaded = st.file_uploader(
        "Drag and drop the files or file here:",
        type = ["docx", "zip"],
        accept_multiple_files = True
    )

    if st.button('Process') and uploaded:
        docx_files = gather_docx_files(uploaded)

        if not docx_files:
            st.error("No DOCX files found in the uploaded content.")
            return
        
        records = [parse_resume(f) for f in docx_files]
        df = pd.DataFrame(records)

        #Build Frequencies
        skill_df, cert_df, degree_tables = build_frequency_tables(df)

        # Export Excel
         
        output_path = os.path.join(tempfile.mkdtemp(), "final-matrix-with-all-frequencies.xlsx")
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="By Category", index=False)
            skill_df.to_excel(writer, sheet_name="SkillFrequency", index=False)
            cert_df.to_excel(writer, sheet_name="CertificationFrequency", index=False)

            degree_tables["Degree/Associates"].to_excel(writer, sheet_name="DegreeFrequency_Associates", index=False)
            degree_tables["Degree/Bachelors"].to_excel(writer, sheet_name="DegreeFrequency_Bachelors", index=False)
            degree_tables["Degree/Masters"].to_excel(writer, sheet_name="DegreeFrequency_Masters", index=False)
            degree_tables["Degree/Phd's"].to_excel(writer, sheet_name="DegreeFrequency_Phds", index=False)


        st.success("Processing complete!")
        st.download_button(
            label = "Download Skills Matrix Excel File",
            data = open(output_path, "rb").read(),
            file_name = "skill-matrix-frequencies.xlsx",
        )

if __name__ == "__main__":
    main()