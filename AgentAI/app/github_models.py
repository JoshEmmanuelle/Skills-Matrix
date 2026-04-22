import os
import streamlit as st
from openai import OpenAI

GITHUB_MODELS_ENDPOINT = "https://models.github.ai/inference"

def get_client():
    token = st.secrets.get("GITHUB_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if not token:
        raise RuntimeError("Missing GITHUB_TOKEN (set Streamlit Secrets or env var).")
    return OpenAI(base_url=GITHUB_MODELS_ENDPOINT, api_key=token)

def ask_github_model(prompt: str, system: str = "", model: str = "openai/gpt-4.1",
                     temperature: float = 0.2, top_p: float = 1.0) -> str:
    client = get_client()
    resp = client.chat.completions.create(
        model=model,
        temperature=temperature,
        top_p=top_p,
        messages=[
            {"role": "system", "content": system or ""},
            {"role": "user", "content": prompt},
        ],
    )
    return resp.choices[0].message.content