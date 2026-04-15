import streamlit as st
from pathlib import Path
import pandas as pd
from collections import Counter

from app.processor import (
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
    upsert_candidate_row,
    rebuild_skill_frequency,
    rebuild_cert_frequency,
    rebuild_degree_frequency,
    write_excel_output,
)

# ============================================================
# App setup
# ============================================================

st.set_page_config(page_title="Skills Matrix Resume Processor", layout="wide")
st.title("Skills Matrix Resume Processor")

st.markdown(
    """
### Processing Mode
- **Option 1**: Update an existing Skills Matrix Excel (merge)
- **Option 2**: Generate a NEW Skills Matrix Excel from resumes only  
  (same rules, same normalization, same frequency rebuilds)
"""
)

# ============================================================
# Mode selection
# ============================================================

mode = st.radio(
    "Choose processing mode:",
    options=[
        "Option 1 – Update existing Skills Matrix Excel (merge)",
        "Option 2 – Generate NEW Skills Matrix Excel from resumes only",
    ],
)

# ============================================================
# Sidebar inputs
# ============================================================

st.sidebar.header("Inputs")

excel_file = None
if mode.startswith("Option 1"):
    excel_file = st.sidebar.file_uploader(
        "Upload MOST UPDATED Skills Matrix Excel (.xlsx)",
        type=["xlsx"],
    )

resume_files = st.sidebar.file_uploader(
    "Upload resume(s) (.docx)",
    type=["docx"],
    accept_multiple_files=True,
)

category_map_path = Path("data/skill_category_map.json")

# Progress bar
progress = st.progress(0, text="Ready")

def set_progress(pct: int, msg: str):
    pct = max(0, min(100, int(pct)))
    progress.progress(pct, text=f"{pct}% — {msg}")

# ============================================================
# Load persistent category map (canonicalized in processor)
# ============================================================

if "category_map" not in st.session_state:
    st.session_state.category_map = load_category_map(category_map_path)

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
# Step B — Learn mappings from base Excel (if present)
# ============================================================

set_progress(15, "Loading existing skill mappings")

if not by_cat.empty:
    excel_map, conflicts = build_category_map_from_by_category(by_cat)

    for skill, cat in excel_map.items():
        if skill not in st.session_state.category_map:
            st.session_state.category_map[skill] = cat

    if conflicts:
        st.warning("Conflicting skill categories detected in base Excel.")
        st.session_state.category_map = resolve_conflicts_with_user(
            conflicts, st.session_state.category_map
        )

save_category_map(category_map_path, st.session_state.category_map)
st.success(f"Loaded {len(st.session_state.category_map)} skill mappings.")

# ============================================================
# Step C — Duplicate name check (Rule 4)
# ============================================================

set_progress(25, "Duplicate name check")

names_peeked = []
for rf in resume_files:
    rf.seek(0)
    nm = peek_name_from_docx(rf)
    names_peeked.append((rf.name, (nm or "").strip()))

existing_names = set(by_cat["Name"].astype(str).str.lower())
batch_counts = Counter([nm.lower() for _, nm in names_peeked if nm])

dup_table = []
actions = {}

for fn, nm in names_peeked:
    nm_norm = nm.strip()
    in_existing = nm_norm.lower() in existing_names if nm_norm else False
    in_batch_dup = batch_counts.get(nm_norm.lower(), 0) > 1 if nm_norm else False
    is_dup = in_existing or in_batch_dup

    dup_table.append({
        "Resume File": fn,
        "Parsed Name": nm_norm,
        "Duplicate?": "YES" if is_dup else "NO",
    })

st.subheader("Duplicate Name Check (Mandatory)")
st.dataframe(pd.DataFrame(dup_table), use_container_width=True)

if any(row["Duplicate?"] == "YES" for row in dup_table):
    st.warning("Duplicates detected. Choose an action for each duplicate.")
    for row in dup_table:
        fn = row["Resume File"]
        nm = row["Parsed Name"]
        if row["Duplicate?"] == "NO":
            actions[fn] = "new"
            continue

        choice = st.radio(
            f"Action for '{nm}' ({fn})",
            options=["replacement (existing person)", "new person (same name)"],
            key=f"dup_{fn}",
        )
        actions[fn] = "replace" if choice.startswith("replacement") else "new"
else:
    for fn, _ in names_peeked:
        actions[fn] = "new"

# ============================================================
# Step D — Process resumes
# ============================================================

set_progress(35, "Ready to process resumes")

if st.button("Run Processor", type="primary"):
    updated_by_cat = by_cat.copy()
    total = len(resume_files)

    for i, rf in enumerate(sorted(resume_files, key=lambda x: x.name.lower()), start=1):
        set_progress(35 + int(50 * i / total), f"Processing {rf.name}")

        rf.seek(0)
        parsed = parse_resume_sections(rf)

        # ✅ Rule 9–11 enforced here
        skills = process_skills(parsed["skills_raw"])

        categorized, st.session_state.category_map = categorize_skills_with_user(
            skills,
            st.session_state.category_map,
            resume_label=rf.name,
        )

        degrees_by_col, earliest_degree_year = apply_degrees(
            parsed["education_lines"],
            parsed["education_text"],
        )

        updated_by_cat = upsert_candidate_row(
            updated_by_cat,
            name=parsed["name"],
            skills_by_category=categorized,
            certs_raw=parsed["certs_raw"],
            degrees_by_col=degrees_by_col,
            earliest_degree_year=earliest_degree_year,
            action=actions.get(rf.name, "new"),
        )

    save_category_map(category_map_path, st.session_state.category_map)

    set_progress(90, "Rebuilding frequency sheets")

    skill_freq = rebuild_skill_frequency(updated_by_cat)
    cert_freq = rebuild_cert_frequency(updated_by_cat)
    deg_assoc = rebuild_degree_frequency(updated_by_cat, "Degree/Associates")
    deg_bach = rebuild_degree_frequency(updated_by_cat, "Degree/Bachelors")
    deg_mast = rebuild_degree_frequency(updated_by_cat, "Degree/Masters")
    deg_phd = rebuild_degree_frequency(updated_by_cat, "Degree/Phds")

    set_progress(98, "Writing output file")

    out_bytes = write_excel_output(
        updated_by_cat,
        skill_freq,
        cert_freq,
        deg_assoc,
        deg_bach,
        deg_mast,
        deg_phd,
    )

    set_progress(100, "Complete")

    filename = (
        "Skills_Matrix_UPDATED.xlsx" # add datetime
        if mode.startswith("Option 1")
        else "Skills_Matrix_from_resumes_ONLY.xlsx" # add datetime
    )

    st.success("Processing complete.")
    st.download_button(
        "Download Excel",
        data=out_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )