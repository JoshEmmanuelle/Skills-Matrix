# Skills Matrix Resume Processor

A deterministic, rule-driven Streamlit application for processing résumés into a structured Skills Matrix Excel file.

This tool extracts **skills, certifications, education, and years of experience** from Word (.docx) résumés and produces a standardized Excel workbook suitable for analytics and reporting.

---

## 🔒 Core Principles

- **No inference** — only explicitly stated résumé data is extracted
- **Deterministic processing** — same input always produces the same output
- **Rulebook-driven** — all logic follows a fixed, cumulative rule set
- **Human-in-the-loop categorization** — new skills require explicit user mapping
- **Full rebuilds** — all frequency sheets are rebuilt every run

---

## 📄 Supported Résumé Sections

The processor extracts data only from the following sections (case-insensitive):

### ✅ Skills
- skills
- technical skills
- skills/tools
- skills/tools/technologies
- skills & tools

### ✅ Education
- education
- educations

### ✅ Certifications
- certification
- certifications

### ✅ Experience
- Used **only** to calculate Years of Experience (YOE)

All other sections (projects, training, summaries, memberships, etc.) are ignored.

---

## 📊 Output Excel Structure

The generated Excel file contains the following sheets:

- **By Category**
- **SkillFrequency**
- **CertificationFrequency**
- **DegreeFrequency_Associates**
- **DegreeFrequency_Bachelors**
- **DegreeFrequency_Masters**
- **DegreeFrequency_Phds**

### Required Columns (By Category)
