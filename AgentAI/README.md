# Skills Matrix Resume Processor (Streamlit)

Rulebook-compliant Streamlit app for processing resumes into a Skills Matrix Excel.

## Key Behaviors (Rulebook)
- Extract ONLY from: **Skills**, **Education/Educations**, **Certification/Certifications**
- No inference. No guessing.
- Duplicate name check happens before parsing.
- Skill categorization uses:
  - previously approved mappings learned from existing Excel rows
  - persistent mappings in `data/skill_category_map.json`
  - user selection for new skills (required; no default)
- Frequency sheets rebuilt from scratch every run (with normalization/removal applied)

## Setup
```bash
# skm = Skills Matrix
python -m venv .skm
# Windows
.\.skm\Scripts\activate
# macOS/Linux
source .skm/bin/activate

pip install -r requirements.txt
streamlit run app.py