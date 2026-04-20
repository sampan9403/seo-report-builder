# ============================================================
# SEO Report Builder — Main Streamlit App
# ============================================================
# Run with:  python -m streamlit run app.py

import streamlit as st
import os, sys, re
import requests as http_requests

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from modules.vision import analyse_keyword_screenshot
from modules.keywords_builder import build_keyword_slides
from modules.task_builder import build_task_slides

st.set_page_config(page_title="SEO Report Builder", page_icon="📊", layout="centered")
st.title("📊 SEO Report Builder")
st.caption("Internal tool — inserts SEO report slides into a Google Slides presentation")
st.divider()


# ── Helpers ───────────────────────────────────────────────────
def extract_pres_id(url: str) -> str:
    if "presentation/d/" in url:
        return url.split("presentation/d/")[1].split("/")[0]
    return url.strip()

def extract_drive_id(url: str) -> str:
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if m: return m.group(1)
    m = re.search(r"id=([a-zA-Z0-9_-]+)", url)
    if m: return m.group(1)
    if re.match(r"^[a-zA-Z0-9_-]+$", url.strip()):
        return url.strip()
    return ""

def drive_image_url(file_id: str) -> str:
    return f"https://drive.google.com/uc?export=view&id={file_id}"

def fetch_image_bytes(file_id: str) -> bytes:
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    r = http_requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content


# ── Step 1: Target Presentation ──────────────────────────────
st.subheader("Step 1 — Target Google Slides")
st.caption("The presentation where new slides will be inserted.")

pres_input = st.text_input(
    "Google Slides URL or ID",
    placeholder="https://docs.google.com/presentation/d/XXXX/edit"
)
pres_id = extract_pres_id(pres_input) if pres_input else ""

if pres_id:
    st.success(f"Presentation ID: `{pres_id}`")
    st.markdown(f"[Open in Google Slides](https://docs.google.com/presentation/d/{pres_id}/edit)")

st.divider()

# ── Section 1: Tasks Completed ────────────────────────────────
st.subheader("Section 1 — Tasks Completed")
st.caption("Slides inserted after the **'Tasks Completed'** header slide.")

tasks_text = st.text_area(
    "Enter completed task names (one per line)",
    placeholder="燈片 Category Page Optimization Suggestion\n便攜式背幕 Product Page Optimization Suggestion\n【貼紙印刷攻略】Blog Writing",
    height=180,
)

tasks_ready = bool(pres_id and tasks_text.strip())

if st.button("Insert Tasks Slides", type="primary", disabled=not tasks_ready, key="btn_tasks"):
    with st.spinner("Categorising tasks with Gemini AI..."):
        try:
            n = build_task_slides(pres_id, tasks_text)
            st.success(f"✓ {n} task slide(s) inserted after 'Tasks Completed' header.")
            st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)

st.divider()

# ── Section 2: Target Keywords Performance ────────────────────
st.subheader("Section 2 — Target Keywords Performance")
st.caption("Slides inserted after the **'Website SEO Performance'** header slide.")

tracking_link = st.text_input(
    "Live Tracking Link URL",
    placeholder="https://app.keyword.com/..."
)

st.info(
    "**How to get Drive image links:**\n"
    "1. Upload screenshot to Google Drive\n"
    "2. Right-click → Share → **'Anyone with the link'** (Viewer)\n"
    "3. Copy link and paste below"
)

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**Overview screenshot** (P.10)")
    link10 = st.text_input("Google Drive link", placeholder="https://drive.google.com/file/d/...", key="link10")
    id10 = extract_drive_id(link10) if link10 else ""
    if id10:
        st.success(f"File ID: `{id10}`")
        try:
            st.image(fetch_image_bytes(id10), caption="Overview preview", width=300)
        except Exception:
            st.warning("Preview unavailable.")
    elif link10:
        st.error("Could not extract file ID.")

with col_b:
    st.markdown("**Ranking screenshot** (P.11, optional)")
    link11 = st.text_input("Google Drive link", placeholder="Leave blank to use overview image", key="link11")
    id11 = extract_drive_id(link11) if link11 else id10
    if link11 and id11:
        st.success(f"File ID: `{id11}`")
        try:
            st.image(fetch_image_bytes(id11), caption="Ranking preview", width=300)
        except Exception:
            st.warning("Preview unavailable.")
    elif link11:
        st.error("Could not extract file ID.")

kw_ready = bool(pres_id and id10)

if not kw_ready:
    missing = []
    if not pres_id: missing.append("presentation ID")
    if not id10:    missing.append("overview screenshot link")
    if missing:
        st.warning(f"Please complete: {', '.join(missing)}")

if st.button("Insert Keywords Slides", type="primary", disabled=not kw_ready, key="btn_kw"):
    try:
        # 1. Download image for Gemini analysis
        with st.spinner("Downloading overview screenshot..."):
            img_bytes = fetch_image_bytes(id10)

        # 2. Analyse with Gemini Vision
        with st.spinner("Analysing keywords with Gemini AI..."):
            keyword_data = analyse_keyword_screenshot(img_bytes)

        total = keyword_data.get("total_keywords", 0)
        st.success(f"Extracted: {total} keywords detected")
        with st.expander("View extracted data"):
            st.json(keyword_data)

        # 3. Build image URLs for Slides API
        url10 = drive_image_url(id10)
        url11 = drive_image_url(id11)

        # 4. Insert slides
        with st.spinner("Inserting keyword slides into presentation..."):
            build_keyword_slides(
                presentation_id=pres_id,
                data=keyword_data,
                image_url_slide10=url10,
                image_url_slide11=url11,
                tracking_link=tracking_link,
            )

        st.success("✓ Keywords Performance slides (P.10 & P.11) inserted.")
        st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)

st.divider()
st.caption("SEO Report Builder · Internal Use Only")
