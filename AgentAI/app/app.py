import io
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

# ------------------------------------------------------------
# Optional chatbot integration (separable)
# ------------------------------------------------------------
try:
    import github_models as gh_llm
except Exception:
    gh_llm = None


# ============================================================
# App setup
# ============================================================

st.set_page_config(page_title="Skills Matrix Resume Processor", layout="wide")
st.title("Skills Matrix Resume Processor - SMRP")

st.markdown(
    """
### Processing Mode
- **Option 1**: Update an existing Skills Matrix Excel (merge)
- **Option 2**: Generate a NEW Skills Matrix Excel

- You can drag and drop the accepted files and folders into the 'Inputs' section
- The Skills Matrix Resume Processor will generate a new excel file depending of the options you select, and it will be automaically downloaded in your downloads folder
"""
)

mode = st.radio(
    "Choose processing mode:",
    options=[
        "Option 1 – Update an existing Skills Matrix Excel with NEW resumes (.docx)",
        "Option 2 – Generate a NEW Skills Matrix Excel using only resumes (.docx)",
    
    ],
)

# ============================================================
# Inputs
# ============================================================

st.sidebar.header("Inputs")

excel_file = None
if mode.startswith("Option 1"):
    excel_file = st.sidebar.file_uploader(
        "Upload the MOST UPDATED Skills Matrix Excel (.xlsx) [Required for Option 1]",
        type=["xlsx"],
    )

resume_files = st.sidebar.file_uploader(
    "Upload resume(s) (.docx) [Required]",
    type=["docx"],
    accept_multiple_files=True,
)

# LLM badge (Enabled / Disabled)
if gh_llm is not None:
    gh_llm.render_llm_badge(where="sidebar")
else:
    st.sidebar.warning("LLM: Disabled (module not installed)")

# st.sidebar.header("Mappings")

# IMPORTANT: anchor data path to this file's folder
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

    # JSON is authoritative; Excel fills missing only
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
    st.session_state.run_requested = True

if st.session_state.get("run_requested"):
    updated_by_cat = by_cat.copy()
    total = len(resume_files)

    for i, rf in enumerate(sorted(resume_files, key=lambda x: x.name.lower()), start=1):
        pct = 35 + int(50 * (i / max(1, total)))
        set_progress(pct, f"Processing resume {i} of {total}: {rf.name}")

        action = actions.get(rf.name, "new")

        rf.seek(0)
        parsed = parse_resume_sections(rf)

        skills = process_skills(parsed["skills_raw"])
        categorized, st.session_state.category_map = categorize_skills_with_user(
            skills,
            st.session_state.category_map,
            resume_label=rf.name,
        )

        degrees_by_col = apply_degrees(
            parsed["education_lines"],
            parsed["education_text"],
        )

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

    st.session_state["excel_bytes"] = out_bytes
    set_progress(100, "Complete")
    st.success("Processing complete.")
    st.session_state.run_requested = False

# ============================================================
# ✅ Persistent Download Button (OPTION A)
# ============================================================

if "excel_bytes" in st.session_state:
    filename = (
        "Skills_Matrix_UPDATED.xlsx"
        if mode.startswith("Option 1")
        else "Skills_Matrix_FROM_RESUMES_ONLY.xlsx"
    )

    st.download_button(
        label="Download Excel",
        data=st.session_state["excel_bytes"],
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ============================================================
# Chatbot (optional)
# ============================================================

if gh_llm is not None and "excel_bytes" in st.session_state:
    gh_llm.render_chat(st.session_state["excel_bytes"], state_key_prefix="chat")