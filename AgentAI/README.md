# Skills Matrix Resume Processor (Streamlit)

Rulebook-compliant Streamlit app for processing resumes into a Skills Matrix Excel.

## Key Behaviors (Rulebook)
- Extract ONLY from: 
      - **Skills**, **Education/Educations**, **Certification/Certifications**, **ProfessionalExperience**

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
# Create and activate a virtual enviroment
python -m venv .skm

# Windows
.\.skm\Scripts\activate

# macOS/Linux
source .skm/bin/activate

#Install the necessary libraries
pip install -r requirements.txt

# Go to the correct path, and run app.py so you can develop in it.
cd Skills-Matrix/AgentAI/app/app.py

# To use the free github models you need to export your token in your terminal
# This is only for local development, and must have no spacing
export GITHUB_TOKEN=<"your_token">
export CHATBOT_ENABLED=true # or CHATBOT_ENABLED = false ; if you dont want to use it.

# Then run the streamlit command
streamlit run app.py

# This app is running 24/7 in the StreamLit Cloud Services, the URL is below
https://skills-matrix-masterpeace.streamlit.app/

# The Streamlit Cloud services require the following secrets
GITHUB_TOKEN="<your_token>"

GITHUB_MODELS_MODEL = "openai/gpt-4o-mini"

GITHUB_MODELS_ENDPOINT="https://models.github.ai/inference"

CHATBOT_ENABLED = true # or CHATBOT_ENABLED = false ; if you dont want to use it.

CHATBOT_MIN_SECONDS = 1.0

# We have an Agent created with Microsoft Copilot, that works very similar to this Skills Matrix Resume Processor (SMRP)
# Find below the link to use it with your compaby's Microsoft 365 Copilot Chat.

https://m365.cloud.microsoft/chat/?titleId=T_28b87335-6fef-cbd1-ce7c-40f08c97d834&source=embedded-builder
