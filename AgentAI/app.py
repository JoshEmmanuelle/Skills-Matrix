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

st.set_page_config(page_title="Skills Matrix Resume Processor", layout="wide")
st.title("Skills Matrix Resume Processor (Streamlit)")

st.markdown(
    """
### Modes
**Option 1 — Update existing Skills Matrix Excel (merge):**
Upload the most updated Excel + resumes; output includes existing rows + new rows.

**Option 2 — Generate NEW Skills Matrix Excel from uploaded resumes ONLY:**
Upload only resumes; output contains only those resumes (same rules applied).
Optional: upload a reference Excel to help auto-map skills (no rows copied).
"""
)

# -----------------------------
# UI: Mode selection
# -----------------------------
mode = st.radio(
    "Choose processing mode",
    options=[
        "Option 1: Update existing Skills Matrix Excel (merge)",
        "Option 2: Generate NEW Skills Matrix Excel from uploaded resumes ONLY",
    ],
    horizontal=False,
)

# -----------------------------
# Inputs
# -----------------------------
st.sidebar.header("Inputs")

excel_file = None
reference_excel = None

if mode.startswith("Option 1"):
    excel_file = st.sidebar.file_uploader(
        "Upload MOST UPDATED Skills Matrix Excel (.xlsx) [Required]",
        type=["xlsx"],
    )
else:
    reference_excel = st.sidebar.file_uploader(
        "Optional: Upload reference Skills Matrix Excel (.xlsx) to expand mappings (no rows copied)",
        type=["xlsx"],
    )

resume_files = st.sidebar.file_uploader(
    "Upload resume(s) (.docx)",
    type=["docx"],
    accept_multiple_files=True
)

st.sidebar.header("Mappings")
category_map_path = Path("data/skill_category_map.json")

# Load persistent mappings (skill_category_map.json on disk)
if "category_map" not in st.session_state:
    st.session_state.category_map = load_category_map(category_map_path)

# Progress bar
progress = st.progress(0, text="Ready")

def set_progress(pct: int, msg: str):
    pct = max(0, min(100, int(pct)))
    progress.progress(pct, text=f"{pct}% — {msg}")

# -----------------------------
# Validate inputs
# -----------------------------
if mode.startswith("Option 1") and not excel_file:
    st.info("Option 1 requires the most updated Excel file.")
    st.stop()

if not resume_files:
    st.info("Upload one or more resumes to proceed.")
    st.stop()

# -----------------------------
# Step A: Establish base By Category dataframe
# -----------------------------
set_progress(5, "Initializing base dataset")

if mode.startswith("Option 1"):
    by_cat = load_excel_by_category(excel_file)
    st.caption("Base dataset: loaded from the uploaded Skills Matrix Excel.")
else:
    by_cat = create_empty_by_category_df()
    st.caption("Base dataset: starting from an empty Skills Matrix (resumes-only mode).")

set_progress(10, "Loading mappings")

# -----------------------------
# Step B: (Optional) expand mappings from reference Excel in Option 2
# -----------------------------
if mode.startswith("Option 2") and reference_excel is not None:
    set_progress(15, "Reading reference Excel to expand mappings")

    ref_by_cat = load_excel_by_category(reference_excel)
    ref_map, ref_conflicts = build_category_map_from_by_category(ref_by_cat)

    # Merge reference mappings into session mapping (do not overwrite existing)
    for skill, cat in ref_map.items():
        if skill not in st.session_state.category_map:
            st.session_state.category_map[skill] = cat

    # If conflicts exist in reference file, rulebook requires user resolution
    if ref_conflicts:
        st.warning(
            "Reference Excel has conflicting skill categories. Resolve to continue (no inference)."
        )
        st.session_state.category_map = resolve_conflicts_with_user(
            ref_conflicts, st.session_state.category_map
        )
        save_category_map(category_map_path, st.session_state.category_map)

    set_progress(20, "Reference mappings merged")

# -----------------------------
# Step C: Learn mappings from base by_cat (Option 1 always has a base Excel)
# -----------------------------
set_progress(25, "Learning mappings from base dataset")

base_map, base_conflicts = build_category_map_from_by_category(by_cat)

# Merge base mappings into JSON map without overwriting existing
for skill, cat in base_map.items():
    if skill not in st.session_state.category_map:
        st.session_state.category_map[skill] = cat

# If base Excel has conflicts, must resolve
if base_conflicts:
    st.warning(
        "The base dataset has conflicting skill categories. Resolve to continue (no inference)."
    )
    st.session_state.category_map = resolve_conflicts_with_user(
        base_conflicts, st.session_state.category_map
    )

save_category_map(category_map_path, st.session_state.category_map)
st.success(f"Loaded {len(st.session_state.category_map)} skill mappings (JSON + Excel learned).")

# -----------------------------
# Step D: Duplicate Name Check (Mandatory First Step)
# -----------------------------
set_progress(30, "Running duplicate name check")

# Peek candidate names for duplicate check (name-only)
names_peeked = []
for rf in resume_files:
    rf.seek(0)
    nm = peek_name_from_docx(rf)
    names_peeked.append((rf.name, nm))

# Determine duplicates vs existing base dataset and duplicates inside the batch
existing_names = set(by_cat["Name"].astype(str).str.lower()) if len(by_cat) else set()
batch_counts = Counter([nm.lower() for _, nm in names_peeked if nm])

dup_table = []
actions = {}

for fn, nm in names_peeked:
    nm_norm = (nm or "").strip()
    in_existing = nm_norm.lower() in existing_names if nm_norm else False
    in_batch_dup = batch_counts.get(nm_norm.lower(), 0) > 1 if nm_norm else False
    is_dup = in_existing or in_batch_dup

    dup_table.append({
        "Resume File": fn,
        "Parsed Name": nm_norm,
        "Duplicate?": "YES" if is_dup else "NO",
        "Reason": ("Exists in base Excel" if in_existing else "") + ("; " if in_existing and in_batch_dup else "") + ("Duplicate in upload batch" if in_batch_dup else ""),
    })

st.subheader("Duplicate Name Check (Mandatory)")
st.dataframe(pd.DataFrame(dup_table), use_container_width=True)

# Ask for user decision for every duplicate (Rule 4/5)
any_dup = any(row["Duplicate?"] == "YES" for row in dup_table)

if any_dup:
    st.warning("Duplicates detected. Choose one action for each duplicate before continuing.")
    for row in dup_table:
        fn = row["Resume File"]
        nm = row["Parsed Name"]
        if row["Duplicate?"] == "NO":
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

set_progress(35, "Duplicate decisions captured")

# -----------------------------
# Step E: Process resumes
# -----------------------------
st.subheader("Process Resumes")
if st.button("Run Processor", type="primary"):
    updated_by_cat = by_cat.copy()

    total = len(resume_files)
    for i, rf in enumerate(sorted(resume_files, key=lambda x: x.name.lower()), start=1):
        pct = 35 + int(50 * (i / max(1, total)))
        set_progress(pct, f"Processing resume {i} of {total}: {rf.name}")

        action = actions.get(rf.name, "new")

        rf.seek(0)
        parsed = parse_resume_sections(rf, name_override=None)

        # Skills
        skills = process_skills(parsed["skills_raw"])

        # Categorize skills (prompts only if missing mapping, permanent via json)
        categorized, st.session_state.category_map = categorize_skills_with_user(
            skills,
            st.session_state.category_map,
            resume_label=rf.name
        )

        # Degrees + experience
        degrees_by_col, earliest_degree_year = apply_degrees(
            parsed["education_lines"], parsed["education_text"]
        )

        # Upsert row (supports replace and new for duplicate name)
        updated_by_cat = upsert_candidate_row(
            updated_by_cat,
            name=parsed["name"],
            skills_by_category=categorized,
            certs_raw=parsed["certs_raw"],
            degrees_by_col=degrees_by_col,
            earliest_degree_year=earliest_degree_year,
            action=action
        )

    # Save mappings after successful run
    save_category_map(category_map_path, st.session_state.category_map)

    set_progress(90, "Rebuilding frequency sheets")

    skill_freq = rebuild_skill_frequency(updated_by_cat)
    cert_freq = rebuild_cert_frequency(updated_by_cat)
    deg_assoc = rebuild_degree_frequency(updated_by_cat, "Degree/Associates")
    deg_bach = rebuild_degree_frequency(updated_by_cat, "Degree/Bachelors")
    deg_mast = rebuild_degree_frequency(updated_by_cat, "Degree/Masters")
    deg_phd = rebuild_degree_frequency(updated_by_cat, "Degree/Phds")

    set_progress(98, "Writing output workbook")

    out_bytes = write_excel_output(
        updated_by_cat, skill_freq, cert_freq, deg_assoc, deg_bach, deg_mast, deg_phd
    )

    set_progress(100, "Complete")

    st.success("Processing complete.")
    out_name = "Skills_Matrix_UPDATED.xlsx" if mode.startswith("Option 1") else "Skills_Matrix_FROM_RESUMES_ONLY.xlsx"
    st.download_button(
        "Download output Excel",
        data=out_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if mode.startswith("Option 2"):
        st.info("Note: Output contains ONLY the uploaded resumes (no base rows).")