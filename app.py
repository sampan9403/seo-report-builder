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
from modules.task_detail_builder import analyze_task_detail, build_task_detail_slides

st.set_page_config(page_title="SEO Report Builder", page_icon="📊", layout="centered")
st.title("📊 SEO Report Builder")
st.caption("Internal tool — inserts SEO report slides into a Google Slides presentation")

DRIVE_UPLOAD_FOLDER = "https://drive.google.com/drive/u/0/folders/1_8N9YLkf9qCk1TJ7KHROJWsxH5gXfRG9"
st.info(
    f"📁 **Upload images here first:** [Google Drive Folder]({DRIVE_UPLOAD_FOLDER})  \n"
    "After uploading, right-click each file → Share → **'Anyone with the link'** (Viewer), then copy the link."
)
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

# ── Section 1a: Tasks Completed — Overview ───────────────────
st.subheader("Section 1a — Tasks Completed (Overview)")
st.caption("Generates a categorised overview slide after the **'Tasks Completed'** header.")

tasks_text = st.text_area(
    "Enter completed task names (one per line)",
    placeholder="燈片 Category Page Optimization Suggestion\n便攜式背幕 Product Page Optimization Suggestion\n【貼紙印刷攻略】Blog Writing",
    height=180,
)

tasks_ready = bool(pres_id and tasks_text.strip())

if st.button("Insert Tasks Overview Slides", type="primary", disabled=not tasks_ready, key="btn_tasks"):
    with st.spinner("Categorising tasks with Gemini AI..."):
        try:
            n = build_task_slides(pres_id, tasks_text)
            st.success(f"✓ {n} overview slide(s) inserted after 'Tasks Completed' header.")
            st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)

st.divider()

# ── Section 1b: Tasks Completed — Detail Slides ──────────────
st.subheader("Section 1b — Task Details")
st.caption(
    "One detailed slide per task — inserted after the overview. "
    "Each task needs at least one of: description, images, or document link."
)

# Initialise session state for dynamic task cards
if "task_detail_cards" not in st.session_state:
    st.session_state.task_detail_cards = [{"id": 0}]
if "task_card_counter" not in st.session_state:
    st.session_state.task_card_counter = 1

# ── "Add Task" button ─────────────────────────────────────────
def add_task_card():
    st.session_state.task_card_counter += 1
    st.session_state.task_detail_cards.append({"id": st.session_state.task_card_counter})

# Render each task card
cards_to_remove = []
for idx, card in enumerate(st.session_state.task_detail_cards):
    cid = card["id"]
    with st.expander(f"Task {idx + 1}", expanded=True):
        col_name, col_del = st.columns([5, 1])
        with col_name:
            st.text_input(
                "Task name",
                key=f"td_name_{cid}",
                placeholder="e.g. Target Blog 4 Content Writing",
            )
        with col_del:
            st.write("")  # spacer
            if st.button("✕ Remove", key=f"td_remove_{cid}"):
                cards_to_remove.append(idx)

        st.text_area(
            "Description (optional) — what was done and why",
            key=f"td_desc_{cid}",
            placeholder="Wrote and optimised blog content targeting keywords X and Y, including title tags and meta descriptions…",
            height=100,
        )

        st.text_area(
            "Image Drive links (one per line, optional)",
            key=f"td_imgs_{cid}",
            placeholder="https://drive.google.com/file/d/...\nhttps://drive.google.com/file/d/...",
            height=80,
        )

        # Preview images
        raw_img_links = st.session_state.get(f"td_imgs_{cid}", "")
        if raw_img_links:
            img_ids_preview = [
                extract_drive_id(l.strip())
                for l in raw_img_links.strip().splitlines()
                if l.strip()
            ]
            valid_ids = [i for i in img_ids_preview if i]
            if valid_ids:
                prev_cols = st.columns(min(len(valid_ids), 3))
                for pi, fid in enumerate(valid_ids[:3]):
                    with prev_cols[pi]:
                        try:
                            st.image(fetch_image_bytes(fid), use_container_width=True)
                        except Exception:
                            st.caption("Preview unavailable")

        st.text_input(
            "Document link (optional) — URL to reference document or page",
            key=f"td_doc_{cid}",
            placeholder="https://docs.google.com/...",
        )

# Remove cards flagged for deletion
for idx in sorted(cards_to_remove, reverse=True):
    st.session_state.task_detail_cards.pop(idx)
if cards_to_remove:
    st.rerun()

# Add Task + Insert buttons
col_add, col_spacer, col_insert = st.columns([2, 1, 3])
with col_add:
    st.button("＋ Add Another Task", on_click=add_task_card)

# Validate: at least one card has some content
def cards_have_content() -> bool:
    for card in st.session_state.task_detail_cards:
        cid = card["id"]
        if (
            st.session_state.get(f"td_name_{cid}", "").strip()
            or st.session_state.get(f"td_desc_{cid}", "").strip()
            or st.session_state.get(f"td_imgs_{cid}", "").strip()
            or st.session_state.get(f"td_doc_{cid}", "").strip()
        ):
            return True
    return False

detail_ready = bool(pres_id and cards_have_content())

with col_insert:
    insert_detail = st.button(
        "Insert Task Detail Slides",
        type="primary",
        disabled=not detail_ready,
        key="btn_task_detail",
    )

if insert_detail:
    tasks_payload = []
    progress = st.progress(0, text="Preparing tasks…")
    total_cards = len(st.session_state.task_detail_cards)

    for i, card in enumerate(st.session_state.task_detail_cards):
        cid   = card["id"]
        name  = st.session_state.get(f"td_name_{cid}", "").strip()
        desc  = st.session_state.get(f"td_desc_{cid}", "").strip()
        imgs  = st.session_state.get(f"td_imgs_{cid}", "").strip()
        doc   = st.session_state.get(f"td_doc_{cid}", "").strip()

        # Skip completely empty cards
        if not any([name, desc, imgs, doc]):
            continue

        progress.progress((i + 0.2) / total_cards, text=f"Analysing task {i+1}: {name or '(unnamed)'}…")

        # Fetch image bytes for Gemini analysis
        img_ids   = [extract_drive_id(l.strip()) for l in imgs.splitlines() if l.strip()]
        img_ids   = [fid for fid in img_ids if fid]
        img_bytes_list = []
        img_urls  = []
        for fid in img_ids[:3]:
            try:
                b = fetch_image_bytes(fid)
                img_bytes_list.append(b)
                img_urls.append(drive_image_url(fid))
            except Exception as e:
                st.warning(f"Could not load image {fid}: {e}")

        # Gemini analysis
        with st.spinner(f"AI analysing task {i+1}…"):
            try:
                result = analyze_task_detail(
                    task_name        = name or "SEO Task",
                    description      = desc,
                    image_bytes_list = img_bytes_list,
                    doc_url          = doc,
                )
            except Exception as e:
                result = {
                    "slide_title": name or "SEO Task",
                    "insight":     desc or "Task completed.",
                    "link_anchor": "View Document" if doc else "",
                }
                st.warning(f"AI analysis failed for task {i+1}, using fallback: {e}")

        tasks_payload.append({
            "name":        name,
            "slide_title": result.get("slide_title", name),
            "insight":     result.get("insight", desc),
            "image_urls":  img_urls,
            "doc_url":     doc,
            "link_anchor": result.get("link_anchor", ""),
        })

        progress.progress((i + 1) / total_cards, text=f"Task {i+1} analysed ✓")

    if tasks_payload:
        with st.spinner("Inserting detail slides into presentation…"):
            try:
                n = build_task_detail_slides(pres_id, tasks_payload)
                st.success(f"✓ {n} task detail slide(s) inserted.")
                st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
            except Exception as e:
                st.error(f"Error inserting slides: {e}")
                st.exception(e)
    else:
        st.warning("No valid tasks to insert.")

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
