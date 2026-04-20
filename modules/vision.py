# ============================================================
# vision.py — Gemini Vision analysis for Keyword.com screenshots
# ============================================================

from google import genai
from google.genai import types
import json
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import GEMINI_API_KEY, GEMINI_MODEL


def analyse_keyword_screenshot(image_bytes: bytes) -> dict:
    api_key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY", "")
    # Streamlit Cloud secrets fallback
    if not api_key:
        try:
            import streamlit as st
            api_key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            pass
    if not api_key:
        raise ValueError("Gemini API key not set. Add GEMINI_API_KEY to Streamlit Secrets.")

    client = genai.Client(api_key=api_key)

    prompt = """Analyse this Keyword.com screenshot showing SEO keyword rankings.
Return ONLY valid JSON with no markdown or explanation:

{
  "total_keywords": <total keywords shown>,
  "top1_count":  <keywords ranked exactly #1>,
  "top3_count":  <keywords ranked in positions 1-3>,
  "top10_count": <keywords ranked in positions 1-10 (page 1)>,
  "top20_count": <keywords ranked in positions 1-20 (top 2 pages)>,
  "top30_count": <keywords ranked in positions 1-30 (top 3 pages)>,
  "improved_count":   <keywords that improved in ranking this period>,
  "maintained_count": <keywords that maintained their ranking>,
  "dropped_count":    <keywords that dropped in ranking>,
  "keywords": [
    {
      "keyword": "<keyword text>",
      "current_rank": <current position>,
      "change": <positive=improved, negative=dropped, 0=same>,
      "volume": <monthly search volume or 0 if not visible>
    }
  ],
  "insight_summary": "<1-2 professional sentences for an SEO monthly report>"
}

Use 0 or empty list for any field not visible in the screenshot."""

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type="image/png"),
            prompt
        ]
    )

    raw = response.text.strip()
    if raw.startswith("```"):
        lines = raw.split("\n")
        raw = "\n".join(lines[1:-1])

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        data = {
            "total_keywords": 0, "top1_count": 0, "top3_count": 0,
            "top10_count": 0, "top20_count": 0, "top30_count": 0,
            "improved_count": 0, "maintained_count": 0, "dropped_count": 0,
            "keywords": [],
            "insight_summary": "Keyword ranking data extracted from screenshot."
        }

    return data
