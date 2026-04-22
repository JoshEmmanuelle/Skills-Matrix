import io
import os
from pathlib import Path
from collections import Counter

import pandas as pd
import streamlit as st

from processor import (
    load_excel_by_category,
    create_empty_by_category_df,
    peek_name_from_docx,
    parse_resume_sections,
    process_skills,
    load_category_map,
    save_category_map,
    build_category_map_from_by_category,
    resolve_conflicts_with_user,
    categorize_skills_with_user,
    apply_degrees,
    extract_oldest_experience_year,
    upsert_candidate_row,
    rebuild_skill_frequency,
    rebuild_cert_frequency,
    rebuild_degree_frequency,
    write_excel_output,
)

# Optional LLM support (GitHub Models via OpenAI-compatible client)
# Requires: openai in requirements.txt and GITHUB_TOKEN in Streamlit Secrets (or env var locally).
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

# ============================================================
# Helpers: safe secrets + token retrieval
# ============================================================

def safe_secret(key: str, default=None):
    """Return a Streamlit secret if available; otherwise default. Does not crash if no secrets.toml."""
    try:
        # st.secrets behaves like a mapping; reading it can raise if no secrets file exists locally
        return st.secrets.get(key, default)
    except Exception:
        return default


def get_github_token():
    """Safest approach: env var locally, Streamlit Cloud secrets in deployment."""
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        return token
    return safe_secret("GITHUB_TOKEN", None)


def llm_status():
    """Return (enabled: bool, reason: str)."""
    if OpenAI is None:
        return False, "openai package not installed"
    token = get_github_token()
    if not token:
        return False, "GITHUB_TOKEN not set"
    return True, "enabled"

# ============================================================
# App setup
# ============================================================

st.set_page_config(page_title="Skills Matrix Resume Processor", layout="wide")
st.title("Skills Matrix Resume Processor")

st.markdown(
    """
### Processing Mode
- **Option 1**: Update an existing Skills Matrix Excel (merge)
- **Option 2**: Generate a NEW Skills Matrix Excel using resumes only
"""
)

mode = st.radio(
    "Choose processing mode:",
    options=[
        "Option 1 – Update existing Skills Matrix Excel (merge)",
        "Option 2 – Generate NEW Skills Matrix Excel using resumes only",
    ],
)

# ============================================================
# Inputs
# ============================================================

st.sidebar.header("Inputs")

excel_file = None
if mode.startswith("Option 1"):
    excel_file = st.sidebar.file_uploader(
        "Upload MOST UPDATED Skills Matrix Excel (.xlsx) [Required for Option 1]",
        type=["xlsx"],
    )

resume_files = st.sidebar.file_uploader(
    "Upload resume(s) (.docx) [Required]",
    type=["docx"],
    accept_multiple_files=True,
)

# LLM badge (Enabled/Disabled)
_enabled, _reason = llm_status()
if _enabled:
    st.sidebar.success("LLM: Enabled")
else:
    st.sidebar.warning(f"LLM: Disabled ({_reason})")

st.sidebar.header("Mappings")

# IMPORTANT: anchor data path to this file's folder (works locally + Streamlit Cloud)
CATEGORY_MAP_PATH = Path(__file__).resolve().parent / "data" / "skill_category_map.json"

progress = st.progress(0, text="Ready")

def set_progress(pct: int, msg: str):
    pct = max(0, min(100, int(pct)))
    progress.progress(pct, text=f"{pct}% — {msg}")

# ============================================================
# Validate inputs
# ============================================================

if mode.startswith("Option 1") and not excel_file:
    st.info("Option 1 requires the most updated Skills Matrix Excel.")
    st.stop()

if not resume_files:
    st.info("Upload one or more resumes to proceed.")
    st.stop()

# ============================================================
# Step A — Initialize base dataset
# ============================================================

set_progress(5, "Initializing dataset")

if mode.startswith("Option 1"):
    by_cat = load_excel_by_category(excel_file)
    st.caption("Base dataset loaded from existing Skills Matrix Excel.")
else:
    by_cat = create_empty_by_category_df()
    st.caption("Base dataset initialized as empty (resumes-only mode).")

# ============================================================
# Step B — Merge-safe category map load
# ============================================================

set_progress(15, "Loading & merging skill mappings (merge-safe)")

category_map = load_category_map(CATEGORY_MAP_PATH)

if mode.startswith("Option 1") and not by_cat.empty:
    excel_map, conflicts_from_excel = build_category_map_from_by_category(by_cat)

    # JSON is authoritative for existing keys; Excel fills missing only
    for skill, cat in excel_map.items():
        if skill not in category_map:
            category_map[skill] = cat

    excel_conflicts = {
        s: {category_map[s], excel_map[s]}
        for s in excel_map
        if s in category_map and category_map[s] != excel_map[s]
    }

    combined_conflicts = {}
    combined_conflicts.update(conflicts_from_excel or {})
    combined_conflicts.update(excel_conflicts or {})

    if combined_conflicts:
        st.warning("Conflicting skill categories detected. Resolve to continue.")
        category_map = resolve_conflicts_with_user(combined_conflicts, category_map)

# Persist merged mapping
save_category_map(CATEGORY_MAP_PATH, category_map)
st.session_state.category_map = category_map

st.success(f"Loaded {len(st.session_state.category_map)} skill mappings (merge-safe).")

# ============================================================
# Step C — Duplicate Name Check (MANDATORY)
# ============================================================

set_progress(25, "Duplicate name check")

names_peeked = []
for rf in resume_files:
    rf.seek(0)
    nm = peek_name_from_docx(rf)
    names_peeked.append((rf.name, (nm or "").strip()))

existing_names = set(by_cat["Name"].astype(str).str.lower()) if not by_cat.empty else set()
batch_counts = Counter([nm.lower() for _, nm in names_peeked if nm])

rows = []
actions = {}

for fn, nm in names_peeked:
    nm_norm = nm.strip()
    in_existing = nm_norm.lower() in existing_names if nm_norm else False
    in_batch_dup = batch_counts.get(nm_norm.lower(), 0) > 1 if nm_norm else False
    is_dup = in_existing or in_batch_dup

    rows.append({
        "Resume File": fn,
        "Parsed Name": nm_norm,
        "Duplicate?": "YES" if is_dup else "NO",
        "Reason": ("Exists in Excel" if in_existing else "") +
                  ("; " if in_existing and in_batch_dup else "") +
                  ("Duplicate in batch" if in_batch_dup else ""),
    })

st.subheader("Duplicate Name Check (Mandatory)")
st.dataframe(pd.DataFrame(rows), use_container_width=True)

any_dup = any(r["Duplicate?"] == "YES" for r in rows)

if any_dup:
    st.warning("Duplicates detected. Choose an action for each duplicate before continuing.")
    for r in rows:
        fn = r["Resume File"]
        nm = r["Parsed Name"]
        if r["Duplicate?"] == "NO":
            actions[fn] = "new"
            continue
        choice = st.radio(
            f"Action for duplicate candidate '{nm}' (file: {fn})",
            options=["replacement (existing person)", "new person (same name)"],
            horizontal=True,
            key=f"dup_action_{fn}",
        )
        actions[fn] = "replace" if choice.startswith("replacement") else "new"
else:
    for fn, _ in names_peeked:
        actions[fn] = "new"

set_progress(35, "Ready to process resumes")

# ============================================================
# Step D — Process resumes
# ============================================================

st.subheader("Process Resumes")

if st.button("Run Processor", type="primary"):
    updated_by_cat = by_cat.copy()
    total = len(resume_files)

    for i, rf in enumerate(sorted(resume_files, key=lambda x: x.name.lower()), start=1):
        pct = 35 + int(50 * (i / max(1, total)))
        set_progress(pct, f"Processing resume {i} of {total}: {rf.name}")

        action = actions.get(rf.name, "new")

        rf.seek(0)
        parsed = parse_resume_sections(rf)

        # Skills
        skills = process_skills(parsed["skills_raw"])
        categorized, st.session_state.category_map = categorize_skills_with_user(
            skills,
            st.session_state.category_map,
            resume_label=rf.name,
        )

        # Degrees
        degrees_by_col = apply_degrees(
            parsed["education_lines"],
            parsed["education_text"],
        )

        # Years of Experience
        oldest_job_year = extract_oldest_experience_year(parsed.get("experience_lines", []))

        updated_by_cat = upsert_candidate_row(
            updated_by_cat,
            name=parsed["name"],
            skills_by_category=categorized,
            certs_raw=parsed["certs_raw"],
            degrees_by_col=degrees_by_col,
            oldest_job_year=oldest_job_year,
            action=action,
        )

    # Persist mapping
    save_category_map(CATEGORY_MAP_PATH, st.session_state.category_map)

    set_progress(90, "Rebuilding frequency sheets")

    skill_freq = rebuild_skill_frequency(updated_by_cat)
    cert_freq = rebuild_cert_frequency(updated_by_cat)
    deg_assoc = rebuild_degree_frequency(updated_by_cat, "Degree/Associates")
    deg_bach = rebuild_degree_frequency(updated_by_cat, "Degree/Bachelors")
    deg_mast = rebuild_degree_frequency(updated_by_cat, "Degree/Masters")
    deg_phd = rebuild_degree_frequency(updated_by_cat, "Degree/Phds")

    set_progress(98, "Writing output workbook")

    out_bytes = write_excel_output(
        updated_by_cat,
        skill_freq,
        cert_freq,
        deg_assoc,
        deg_bach,
        deg_mast,
        deg_phd,
    )

    # Store workbook in memory for Q&A (no download required)
    st.session_state["excel_bytes"] = out_bytes
    st.session_state.pop("excel_context", None)  # refresh cached context

    set_progress(100, "Complete")

    filename = (
        "Skills_Matrix_UPDATED.xlsx"
        if mode.startswith("Option 1")
        else "Skills_Matrix_FROM_RESUMES_ONLY.xlsx"
    )

    st.success("Processing complete.")
    st.download_button(
        "Download Excel",
        data=out_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ============================================================
# Q&A about output workbook (GitHub Models)
# ============================================================

def _get_github_models_client():
    """OpenAI-compatible client configured for GitHub Models."""
    if OpenAI is None:
        return None, "The 'openai' package is not installed. Add 'openai' to requirements.txt."

    token = get_github_token()
    if not token:
        return None, "Missing GITHUB_TOKEN. Add it to Streamlit Cloud Secrets or set env var locally."

    endpoint = safe_secret("GITHUB_MODELS_ENDPOINT", "https://models.github.ai/inference")
    return OpenAI(base_url=endpoint, api_key=token), None


def _build_excel_context(excel_bytes: bytes) -> str:
    """Build a compact, deterministic summary of the generated workbook."""
    xf = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    sheets = xf.sheet_names

    ctx = []
    ctx.append(f"Sheets: {', '.join(sheets)}")

    # By Category
    if "By Category" in sheets:
        by_cat_df = xf.parse("By Category")
        ctx.append(f"Total candidates (By Category): {len(by_cat_df)}")
        ctx.append(f"Columns (By Category): {', '.join(list(by_cat_df.columns))}")

    # SkillFrequency
    if "SkillFrequency" in sheets:
        sf = xf.parse("SkillFrequency")
        if not sf.empty and "Candidate Count" in sf.columns:
            top = sf.sort_values("Candidate Count", ascending=False).head(20)
            ctx.append("Top Skills (top 20):")
            for _, r in top.iterrows():
                ctx.append(f"- {r['Skill']}: {int(r['Candidate Count'])}")

    # CertificationFrequency
    if "CertificationFrequency" in sheets:
        cf = xf.parse("CertificationFrequency")
        if not cf.empty and "Candidate Count" in cf.columns:
            top = cf.sort_values("Candidate Count", ascending=False).head(20)
            ctx.append("Top Certifications (top 20):")
            for _, r in top.iterrows():
                ctx.append(f"- {r['Certification']}: {int(r['Candidate Count'])}")

    # Degree frequencies
    deg_sheets = [
        "DegreeFrequency_Associates",
        "DegreeFrequency_Bachelors",
        "DegreeFrequency_Masters",
        "DegreeFrequency_Phds",
    ]
    for sh in deg_sheets:
        if sh in sheets:
            df = xf.parse(sh)
            if not df.empty and "Candidate Count" in df.columns:
                top = df.sort_values("Candidate Count", ascending=False).head(10)
                ctx.append(f"Top Degrees ({sh.replace('DegreeFrequency_', '')}, top 10):")
                for _, r in top.iterrows():
                    ctx.append(f"- {r['Degree']}: {int(r['Candidate Count'])}")

    return "\n".join(ctx)


def _ask_llm(question: str, context: str) -> str:
    client, err = _get_github_models_client()
    if err:
        return f"LLM not available: {err}"

    model = safe_secret("GITHUB_MODELS_MODEL", "openai/gpt-4.1")

    system = (
        "You answer questions using ONLY the provided Excel Summary Context. "
        "Do not guess. If the answer is not contained in the context, say you do not have enough information."
    )

    prompt = f"Excel Summary Context:\n{context}\n\nUser Question:\n{question}"

    resp = client.chat.completions.create(
        model=model,
        temperature=0.2,
        top_p=1.0,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ],
    )

    return resp.choices[0].message.content


st.divider()
st.subheader("Ask questions about the generated Excel output")

if "excel_bytes" not in st.session_state:
    st.info("Run the processor above to generate the Excel output. Then you can ask questions here.")
else:
    # Build / cache context once per output
    if "excel_context" not in st.session_state:
        st.session_state["excel_context"] = _build_excel_context(st.session_state["excel_bytes"])

    # Optional: show context preview
    with st.expander("Show Excel summary context (what the assistant can use)"):
        st.text(st.session_state["excel_context"])

    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []

    user_q = st.chat_input("Ask a question about the output (e.g., top certifications, how many candidates, top skills)")
    if user_q:
        answer = _ask_llm(user_q, st.session_state["excel_context"])
        st.session_state["chat_history"].append(("user", user_q))
        st.session_state["chat_history"].append(("assistant", answer))

    for role, msg in st.session_state["chat_history"]:
        with st.chat_message(role):
            st.write(msg)