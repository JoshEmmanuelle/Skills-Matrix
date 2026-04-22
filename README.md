# Skills Matrix Resume Processor

A **deterministic, rule-driven Streamlit application** for processing résumés into a structured **Skills Matrix Excel file**, with an **optional analytical chatbot** for querying the output.

This tool extracts **skills, certifications, education, and years of experience** from Word (`.docx`) résumés and produces a standardized Excel workbook suitable for analytics, reporting, and downstream workforce planning.

---

## 🔒 Core Principles

- **No inference** — only explicitly stated résumé data is extracted or analyzed
- **Deterministic processing** — same input always produces the same output
- **Rulebook-driven** — all logic follows a fixed, cumulative rule set
- **Human-in-the-loop categorization** — new skills require explicit user mapping
- **Full rebuilds** — all frequency sheets are rebuilt every run
- **Separation of concerns** — résumé processing and AI analysis are fully decoupled

All rules are **permanent and cumulative**. New rules extend the system; none are removed unless explicitly stated.

---

## 📄 Supported Résumé Sections

The processor extracts data **only** from the following sections (case-insensitive):

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
- Used **only** to calculate **Years of Experience (YOE)**

🚫 All other sections (summaries, projects, descriptions, narrative text, training, memberships, etc.) are **ignored entirely**.

---

## 📊 Output Excel Structure

The generated Excel file always contains the following sheets:

- **By Category**
- **SkillFrequency**
- **CertificationFrequency**
- **DegreeFrequency_Associates**
- **DegreeFrequency_Bachelors**
- **DegreeFrequency_Masters**
- **DegreeFrequency_Phds**

---

### ✅ Required Columns (By Category)

The `By Category` sheet must contain **exactly** the following columns:

Name
Years of Experience
Language
Cloud
Databases
OS
Framework/Libraries
DevOps
Container/Orchestration
Machine Learning/AI
Networking
Version Control
Tools
Other
Certifications
Degree/Associates
Degree/Bachelors
Degree/Masters
Degree/Phds

❗ No additional columns are permitted.

---

## 🔁 Duplicate Name Handling (Mandatory)

Before any résumé parsing occurs:

1. Each résumé’s **Name** is compared against `By Category → Name`
2. Matching rules:
   - Case-insensitive
   - Exact string match only
   - No fuzzy matching, aliases, or abbreviations
3. If no match is found:
   - Résumé is processed as a **new candidate**
4. If a match is found:
   - Processing **halts immediately**
   - The user must choose **one and only one** action:
     - New person with the same name
     - Replacement résumé for the existing person
5. After resolution:
   - All frequency sheets are **rebuilt from scratch**

This duplicate-name logic runs **every time**, without exception.

---

## 📈 Years of Experience (YOE)

- Calculated as: Years of Experience = 2026(or current year) − year of oldest job experience

## 🧠 Skill Processing Rules (High Level)

Skills are processed in the following **mandatory order**:

1. Literal extraction  
2. Duplicate removal (within résumé)  
3. Compound splitting  
4. Skill normalization / merging  
5. Skill categorization  

At no point is inference, guessing, or enrichment allowed.

---

## 🧩 Skill Categorization

- Skills may belong to **exactly one** predefined category
- Categorization sources:
- Previously approved mappings (automatic)
- Explicit user instructions
- If a skill has never been categorized:
- Processing pauses
- User must assign it to one category
- Once mapped:
- The mapping is **permanent**
- The same skill will never prompt again
- Approved mappings persist across:
- All future résumés
- All sessions
- All batches

---

## 📈 Frequency Sheet Rules

- **SkillFrequency**, **CertificationFrequency**, and **DegreeFrequency** sheets:
- Are rebuilt **from scratch every run**
- Never updated incrementally
- Never carry over historical counts
- Counts represent **unique candidates**, not occurrences

---

## Optional AI Chatbot (Analysis Layer)

> **Important:** The chatbot is **optional** and **read-only**.  
> It never modifies résumé data or Excel outputs.

The chatbot allows users to **ask questions about the generated Excel file**, such as:

- “Who has Kubernetes experience?”
- “Which candidates have AWS certifications?”
- “Show candidates with Machine Learning/AI skills”
- “What skills appear most frequently?”
- “Which rows match ‘Terraform’?”

### Key Characteristics

- ✅ Read-only (no Excel modification)
- ✅ Rule-safe (no inference, no recommendations)
- ✅ Deterministic, question-scoped context slicing
- ✅ Rate-limit aware (GitHub Models free tier)
- ✅ Fully detachable via configuration

---

## Free Access & Rate Limits

The chatbot uses **GitHub Models free-tier access**, which is subject to **rate limits**.

When limits are reached:
- The app shows a friendly message
- The app never crashes
- Requests are throttled automatically

**For deeper or high-volume analysis**, use your company’s **Microsoft Copilot Chat** instead.

This notice is shown directly in the UI.

---

##  Enabling / Disabling the Chatbot

### Disable instantly (recommended for production reviews)

In **Streamlit Cloud Secrets**:

CHATBOT_ENABLED = false/true

# More information about Streamlit Cloud Secrets

GITHUB_TOKEN="<your_token>"

GITHUB_MODELS_MODEL = "openai/gpt-4.1"

GITHUB_MODELS_ENDPOINT="https://models.github.ai/inference"

CHATBOT_MIN_SECONDS = 30.0



