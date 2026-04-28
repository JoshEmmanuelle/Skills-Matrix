"""github_models.py

GitHub Models chatbot integration (OpenAI-compatible via OpenAI Python SDK).

Purpose
- Provide a detachable chatbot layer for the Skills Matrix Resume Processor.
- Keep all LLM logic out of app.py so it can be disabled easily.

How to disable quickly
- Streamlit Cloud Secrets: CHATBOT_ENABLED = false
- Or environment variable: CHATBOT_ENABLED=false

Free access notice
- GitHub Models free tiers have rate limits. When limits are hit, the chatbot will
  show a friendly message instead of crashing.
- For deeper / high-volume analysis, use your company's Microsoft Copilot Chat.

Required (Streamlit Cloud Secrets)
- GITHUB_TOKEN

Optional (Streamlit Cloud Secrets)
- GITHUB_TOKEN = "<your_token>"
- GITHUB_MODELS_MODEL (default: openai/gpt-4o-mini)
- GITHUB_MODELS_ENDPOINT (default: https://models.github.ai/inference)
- CHATBOT_ENABLED (true/false)

Local development
- Preferred: export GITHUB_TOKEN in your shell
- Streamlit Cloud secrets are NOT accessible locally.

Notes on context limits
- Do NOT send the entire Excel workbook content to the LLM.
- This module builds a deterministic, question-scoped context slice to stay within
  hosted model limits and reduce rate-limit pressure.
"""

import hashlib
import io
import os
import re
import time
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

try:
    from openai import OpenAI
    from openai import APIStatusError, RateLimitError
except Exception:
    OpenAI = None
    APIStatusError = Exception
    RateLimitError = Exception


# ----------------------------
# Secrets & configuration
# ----------------------------

def safe_secret(key: str, default=None):
    """Safely read Streamlit secrets.

    Accessing st.secrets can raise StreamlitSecretNotFoundError locally if no
    secrets.toml exists. This wrapper prevents crashes.
    """
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def _env_bool(name: str) -> Optional[bool]:
    v = os.environ.get(name)
    if v is None:
        return None
    v = v.strip().lower()
    if v in {"1", "true", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "no", "n", "off"}:
        return False
    return None


def chatbot_enabled() -> bool:
    """Whether chatbot is enabled."""
    env_override = _env_bool("CHATBOT_ENABLED")
    if env_override is not None:
        return env_override
    return bool(safe_secret("CHATBOT_ENABLED", True))


def get_github_token() -> Optional[str]:
    """Token lookup order:

    1) Local: environment variable GITHUB_TOKEN
    2) Streamlit secrets: Cloud Secrets UI (or local secrets.toml if you created one)

    NOTE: Streamlit Cloud secrets are not accessible from local VS Code runs.
    """
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        return token
    return safe_secret("GITHUB_TOKEN", None)


def github_models_endpoint() -> str:
    return safe_secret("GITHUB_MODELS_ENDPOINT", "https://models.github.ai/inference")


def model_name() -> str:
    return safe_secret("GITHUB_MODELS_MODEL", "openai/gpt-4o-mini")


def llm_status() -> Tuple[bool, str]:
    """Return (enabled: bool, reason: str)."""
    if not chatbot_enabled():
        return False, "disabled by configuration"
    if OpenAI is None:
        return False, "openai package not installed"
    if not get_github_token():
        return False, "GITHUB_TOKEN not set"
    return True, "enabled"


def _get_client() -> Tuple[Optional[OpenAI], Optional[str]]:
    if OpenAI is None:
        return None, "The 'openai' package is not installed. Add 'openai' to requirements.txt."

    if not chatbot_enabled():
        return None, "Chatbot disabled by configuration (CHATBOT_ENABLED=false)."

    token = get_github_token()
    if not token:
        return None, "Missing GITHUB_TOKEN. Add it to Streamlit Cloud Secrets or set env var locally."

    return OpenAI(base_url=github_models_endpoint(), api_key=token), None


# ----------------------------
# Rate-limit friendly guardrails
# ----------------------------

def _throttle_seconds() -> float:
    """Minimum seconds between LLM calls (deterministic)."""

    env = os.environ.get("CHATBOT_MIN_SECONDS")
    if env:
        try:
            return max(0.0, float(env))
        except Exception:
            pass
    return float(safe_secret("CHATBOT_MIN_SECONDS", 1.0))


def _free_tier_notice() -> str:
    return (
        "Note: This chatbot uses MasterPeace GitHub Models access and may be rate-limited. "
        "If you need deeper or high-volume analysis, use your company's Microsoft 365 Copilot Chat."
    )


# ----------------------------
# Context builder (question-scoped)
# ----------------------------

def build_excel_context(excel_bytes: bytes, question: str = "") -> str:
    """Build a deterministic, question-scoped context from the workbook.

    This is designed to support 'ask anything' without sending the entire workbook.

    Always includes:
    - Workbook index (sheets + row counts + columns)

    Then adds deterministic slices based on question:
    - If question contains a sheet name -> preview that sheet
    - If question contains candidate name -> that candidate row from By Category
    - Otherwise -> previews of key sheets + token-matched By Category rows

    Hard-capped to avoid oversize payloads.
    """

    MAX_TEXT_CHARS = 60_000
    MAX_PREVIEW_ROWS = 40
    MAX_MATCH_ROWS = 60
    MAX_EXTRA_SHEETS = 6

    q = (question or "").strip()
    q_low = q.lower()

    xf = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    sheets = xf.sheet_names

    ctx = []
    ctx.append("WORKBOOK INDEX")
    ctx.append(f"Sheets: {', '.join(sheets)}")

    def tsv(df: pd.DataFrame) -> str:
        return df.to_csv(sep="\t", index=False)

    def safe_parse(sheet_name: str) -> pd.DataFrame:
        try:
            return xf.parse(sheet_name).fillna("")
        except Exception:
            return pd.DataFrame()

    def head(df: pd.DataFrame, n: int) -> pd.DataFrame:
        return df.head(n)

    def truncate_text(text: str) -> str:
        if len(text) <= MAX_TEXT_CHARS:
            return text
        return text[:MAX_TEXT_CHARS] + "\n\nNOTE: Context truncated by MAX_TEXT_CHARS limit."

    def normalize_token(s: str) -> str:
        return re.sub(r"\s+", " ", (s or "").strip().lower())

    def find_rows_containing(df: pd.DataFrame, needle: str, max_rows: int) -> pd.DataFrame:
        if df.empty or not needle:
            return df.head(0)
        n = needle.lower()
        mask = df.apply(lambda row: any(n in str(v).lower() for v in row.values), axis=1)
        return df[mask].head(max_rows)

    # metadata
    meta = []
    for sh in sheets:
        df = safe_parse(sh)
        cols = list(df.columns) if not df.empty else []
        meta.append((sh, len(df), cols))

    for sh, nrows, cols in meta:
        ctx.append(f"- {sh}: {nrows} rows; columns: {', '.join(cols) if cols else '(unreadable/empty)'}")

    # If no question, include small previews of key sheets
    if not q:
        for sh in ["By Category", "SkillFrequency", "CertificationFrequency"]:
            if sh in sheets:
                df = safe_parse(sh)
                ctx.append(f"\nPREVIEW: {sh} (first {MAX_PREVIEW_ROWS} rows, TSV)")
                ctx.append(tsv(head(df, MAX_PREVIEW_ROWS)))
        return truncate_text("\n".join(ctx))

    # 1) sheet-name targeting
    matched_sheets = []
    q_norm = normalize_token(q)
    for sh in sheets:
        if normalize_token(sh) in q_norm:
            matched_sheets.append(sh)

    if matched_sheets:
        for sh in matched_sheets:
            df = safe_parse(sh)
            ctx.append(f"\nFOCUSED SHEET: {sh} (first {MAX_PREVIEW_ROWS} rows, TSV)")
            ctx.append(tsv(head(df, MAX_PREVIEW_ROWS)))
        return truncate_text("\n".join(ctx))

    # 2) candidate-name targeting (By Category)
    by_cat = safe_parse("By Category") if "By Category" in sheets else pd.DataFrame()
    if not by_cat.empty and "Name" in by_cat.columns:
        names = [str(n).strip() for n in by_cat["Name"].tolist() if str(n).strip()]
        name_hits = [nm for nm in names if nm.lower() in q_low]
        if name_hits:
            subset = by_cat[by_cat["Name"].isin(name_hits)]
            ctx.append("\nFOCUSED CANDIDATE ROW(S) FROM BY CATEGORY (TSV)")
            ctx.append(tsv(subset))
            return truncate_text("\n".join(ctx))

    # 3) intent-based previews
    wants_skills = any(k in q_low for k in ["skill", "skills", "technology", "technologies"])
    wants_certs = "cert" in q_low
    wants_degrees = any(k in q_low for k in ["degree", "degrees", "education", "bachelor", "master", "phd", "associate"])

    if wants_skills and "SkillFrequency" in sheets:
        df = safe_parse("SkillFrequency")
        ctx.append(f"\nSkillFrequency (first {MAX_PREVIEW_ROWS} rows, TSV)")
        ctx.append(tsv(head(df, MAX_PREVIEW_ROWS)))

    if wants_certs and "CertificationFrequency" in sheets:
        df = safe_parse("CertificationFrequency")
        ctx.append(f"\nCertificationFrequency (first {MAX_PREVIEW_ROWS} rows, TSV)")
        ctx.append(tsv(head(df, MAX_PREVIEW_ROWS)))

    if wants_degrees:
        for sh in [
            "DegreeFrequency_Associates",
            "DegreeFrequency_Bachelors",
            "DegreeFrequency_Masters",
            "DegreeFrequency_Phds",
        ]:
            if sh in sheets:
                df = safe_parse(sh)
                ctx.append(f"\n{sh} (first {MAX_PREVIEW_ROWS} rows, TSV)")
                ctx.append(tsv(head(df, MAX_PREVIEW_ROWS)))

    # 4) token match against By Category for broad queries
    if not by_cat.empty:
        tokens = [t for t in re.split(r"[^a-zA-Z0-9\+\.\-/_]+", q_low) if len(t) >= 4]
        tokens = sorted(set(tokens), key=lambda x: (-len(x), x))

        matched = pd.DataFrame()
        for tok in tokens[:5]:
            matched = find_rows_containing(by_cat, tok, MAX_MATCH_ROWS)
            if not matched.empty:
                ctx.append(f"\nBY CATEGORY ROWS MATCHING TOKEN '{tok}' (up to {MAX_MATCH_ROWS} rows, TSV)")
                ctx.append(tsv(matched))
                break

        if matched.empty:
            ctx.append(f"\nBY CATEGORY PREVIEW (first {MAX_PREVIEW_ROWS} rows, TSV)")
            ctx.append(tsv(head(by_cat, MAX_PREVIEW_ROWS)))

    # 5) extra sheet previews (broad fallback)
    extras = [sh for sh in sheets if sh not in {"By Category", "SkillFrequency", "CertificationFrequency"}]
    for sh in extras[:MAX_EXTRA_SHEETS]:
        df = safe_parse(sh)
        if df.empty:
            continue
        ctx.append(f"\nPREVIEW: {sh} (first {MAX_PREVIEW_ROWS} rows, TSV)")
        ctx.append(tsv(head(df, MAX_PREVIEW_ROWS)))

    return truncate_text("\n".join(ctx))


# ----------------------------
# LLM call (with throttle + user-friendly errors)
# ----------------------------

def ask_llm(question: str, context: str) -> str:
    """Ask GitHub Models about the workbook context.

    This function is rate-limit aware and will not crash the app.
    """
    client, err = _get_client()
    if err:
        return f"LLM not available: {err}"

    # Throttle to protect MasterPeace GitHub Models access limits
    now = time.time()
    last = st.session_state.get("_llm_last_call", 0.0)
    min_s = _throttle_seconds()
    if now - last < min_s:
        return (
            f"Please wait {max(0, int(min_s - (now - last)))}s and try again. "
            + _free_tier_notice()
        )
    st.session_state["_llm_last_call"] = now

    system = (
        "You answer questions using ONLY the provided/output Excel Context. "
        "Do not guess. If the answer is not contained in the context, say you do not have enough information."
        "You can make recommendations using the excel file as context"
    )
    prompt = f"Excel Context:\n{context}\n\nUser Question:\n{question}"

    try:
        resp = client.chat.completions.create(
            model=model_name(),
            temperature=0.2,
            top_p=1.0,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": prompt},
            ],
        )
        return resp.choices[0].message.content

    except RateLimitError:
        return (
            "Rate limit reached for MasterPeace GitHub Models Access. Please try again later.\n\n"
            + _free_tier_notice()
        )

    except APIStatusError as e:
        # Many providers use generic status errors for rate-limit, payload, or policy.
        # We show a safe, non-leaky message.
        return (
            "The AI service returned an error (possibly rate limits or request size). "
            "Please try a shorter question or wait and try again.\n\n"
            + _free_tier_notice()
        )

    except Exception:
        return (
            "The AI service encountered an unexpected error. Please try again.\n\n"
            + _free_tier_notice()
        )


# ----------------------------
# UI helpers
# ----------------------------

def render_llm_badge(where: str = "sidebar") -> None:
    """Render an LLM Enabled/Disabled badge."""
    enabled, reason = llm_status()
    badge_text = "LLM: Enabled" if enabled else f"LLM: Disabled ({reason})"

    if where == "sidebar":
        if enabled:
            st.sidebar.success(badge_text)
            st.sidebar.caption("MasterPeace LLM access: GitHub Models may be rate-limited. For deeper analysis use company's Microsoft Copilot Chat.")
        else:
            st.sidebar.warning(badge_text)
    else:
        if enabled:
            st.success(badge_text)
            st.caption("MasterPeace LLM access: GitHub Models may be rate-limited. For deeper analysis use company's Microsoft Copilot Chat.")
        else:
            st.warning(badge_text)


def render_chat(excel_bytes: bytes, state_key_prefix: str = "chat") -> None:
    """Render chatbot UI for a workbook bytes.

    Includes:
    - Optional Excel upload to ask questions about any workbook
    - Per-workbook and per-question caching to reduce repeated calls
    - Throttling to respect free-tier limits
    """

    enabled, reason = llm_status()

    st.divider()
    st.subheader("Ask questions about the Excel output")
    st.caption(_free_tier_notice())

    if not enabled:
        st.info(f"Chatbot is disabled: {reason}")
        return

    # Optional Excel upload
    upload_key = f"{state_key_prefix}_uploaded_excel"
    uploaded = st.file_uploader(
        "Optional: Upload an Excel (.xlsx) to ask questions about it",
        type=["xlsx"],
        key=upload_key,
        help="If provided, the chatbot will use this uploaded workbook instead of the most recent generated output.",
    )

    if uploaded is not None:
        active_bytes = uploaded.getvalue()
        st.caption(f"Using uploaded file: {uploaded.name}")
    else:
        active_bytes = excel_bytes
        st.caption("Using generated output from this session")

    # Stable signature per workbook
    sig = hashlib.sha256(active_bytes).hexdigest()[:12]
    hist_key = f"{state_key_prefix}_history_{sig}"
    cache_key = f"{state_key_prefix}_answer_cache_{sig}"

    if hist_key not in st.session_state:
        st.session_state[hist_key] = []
    if cache_key not in st.session_state:
        st.session_state[cache_key] = {}

    # Ask question first
    user_q = st.chat_input("Ask a question about the Excel just produced (e.g., who has Kubernetes, top certifications, etc.)")

    if user_q:
        q_norm = re.sub(r"\s+", " ", user_q.strip())
        q_hash = hashlib.sha256(q_norm.encode("utf-8")).hexdigest()[:12]

        # Cache per question to avoid re-calling on reruns
        if q_hash in st.session_state[cache_key]:
            answer, context_used = st.session_state[cache_key][q_hash]
        else:
            context_used = build_excel_context(active_bytes, question=user_q)
            answer = ask_llm(user_q, context_used)
            st.session_state[cache_key][q_hash] = (answer, context_used)

        st.session_state[hist_key].append(("user", user_q))
        st.session_state[hist_key].append(("assistant", answer))

        # Provide a transparency expander for the last question
        with st.expander("Show context used for the last question"):
            st.text(context_used)

    # Render chat history
    for role, msg in st.session_state[hist_key]:
        with st.chat_message(role):
            st.write(msg)