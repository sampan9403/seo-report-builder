# ============================================================
# SEO Report Builder — Main Streamlit App
# ============================================================
# Run with:  python -m streamlit run app.py

import streamlit as st
import os, sys, re
import requests as http_requests

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from modules.keywords_builder import (
    build_keyword_overview_slide,
    build_keyword_table_slides,
    parse_keyword_csv,
)
from modules.task_builder import build_task_slides
from modules.task_detail_builder import analyze_task_detail, build_task_detail_slides
from modules.ahrefs_builder import (
    analyze_ahrefs_slide,
    build_ahrefs_slides,
    parse_competitor_csv,
    analyze_organic_competitors,
    build_organic_competitors_slide,
)
from modules.gsc_builder import (
    analyze_gsc_images,
    build_gsc_image_slides,
    parse_gsc_csv,
    analyze_gsc_csv,
    build_gsc_csv_slide,
)

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

def read_as_csv(uploaded_file, key_prefix: str) -> bytes | None:
    """
    Accept CSV or Excel uploads; return UTF-8 CSV bytes.
    For Excel files with multiple sheets, show a sheet selector so the
    uploader can specify which tab they are importing.
    """
    import pandas as pd

    fname = (uploaded_file.name or "").lower()
    is_excel = fname.endswith(".xlsx") or fname.endswith(".xls")

    if not is_excel:
        data = uploaded_file.read()
        uploaded_file.seek(0)
        return data

    # ── Excel path ────────────────────────────────────────────
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        sheets = xl.sheet_names
    except Exception as e:
        st.error(f"Could not open Excel file: {e}")
        return None

    if len(sheets) == 1:
        selected = sheets[0]
        st.caption(f"Excel file detected — importing from sheet: **{selected}**")
    else:
        st.info(
            f"This Excel file contains **{len(sheets)} sheets**. "
            "Please select the tab you want to import:"
        )
        selected = st.selectbox(
            "Select sheet / tab to import",
            options=sheets,
            key=f"{key_prefix}_sheet_sel",
        )

    try:
        df = xl.parse(selected)
        return df.to_csv(index=False).encode("utf-8-sig")
    except Exception as e:
        st.error(f"Could not read sheet '{selected}': {e}")
        return None


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

# ════════════════════════════════════════════════════════════
# Section 1a — Tasks Completed (Overview)
# ════════════════════════════════════════════════════════════
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

# ════════════════════════════════════════════════════════════
# Section 1b — Task Details
# ════════════════════════════════════════════════════════════
st.subheader("Section 1b — Task Details")
st.caption(
    "One detailed slide per task — inserted after the overview. "
    "Each task needs at least one of: description, images, or document link."
)

if "task_detail_cards" not in st.session_state:
    st.session_state.task_detail_cards = [{"id": 0}]
if "task_card_counter" not in st.session_state:
    st.session_state.task_card_counter = 1

def add_task_card():
    st.session_state.task_card_counter += 1
    st.session_state.task_detail_cards.append({"id": st.session_state.task_card_counter})

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
            st.write("")
            if st.button("✕ Remove", key=f"td_remove_{cid}"):
                cards_to_remove.append(idx)

        st.text_area(
            "Description (optional) — what was done and why",
            key=f"td_desc_{cid}",
            placeholder="Wrote and optimised blog content targeting keywords X and Y…",
            height=100,
        )
        st.text_area(
            "Image Drive links (one per line, optional)",
            key=f"td_imgs_{cid}",
            placeholder="https://drive.google.com/file/d/...",
            height=80,
        )

        raw_img_links = st.session_state.get(f"td_imgs_{cid}", "")
        if raw_img_links:
            img_ids_preview = [
                extract_drive_id(l.strip())
                for l in raw_img_links.strip().splitlines() if l.strip()
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

for idx in sorted(cards_to_remove, reverse=True):
    st.session_state.task_detail_cards.pop(idx)
if cards_to_remove:
    st.rerun()

col_add, col_spacer, col_insert = st.columns([2, 1, 3])
with col_add:
    st.button("＋ Add Another Task", on_click=add_task_card)

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

        if not any([name, desc, imgs, doc]):
            continue

        progress.progress((i + 0.2) / total_cards, text=f"Analysing task {i+1}: {name or '(unnamed)'}…")

        img_ids = [extract_drive_id(l.strip()) for l in imgs.splitlines() if l.strip()]
        img_ids = [fid for fid in img_ids if fid]
        img_bytes_list = []
        img_urls = []
        for fid in img_ids[:3]:
            try:
                b = fetch_image_bytes(fid)
                img_bytes_list.append(b)
                img_urls.append(drive_image_url(fid))
            except Exception as e:
                st.warning(f"Could not load image {fid}: {e}")

        with st.spinner(f"AI analysing task {i+1}…"):
            try:
                result = analyze_task_detail(
                    task_name=name or "SEO Task",
                    description=desc,
                    image_bytes_list=img_bytes_list,
                    doc_url=doc,
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

# ════════════════════════════════════════════════════════════
# Section 2a — Keywords Overview
# ════════════════════════════════════════════════════════════
st.subheader("Section 2a — Keywords Overview")
st.caption(
    "One overview slide with ranking stats, live tracking link, and screenshot. "
    "Inserted at the end of the **'Website SEO Performance'** section."
)

tracking_link = st.text_input(
    "Live Tracking Link URL (optional)",
    placeholder="https://app.keyword.com/...",
    key="kw_tracking_link",
)

st.markdown("**Overview screenshot** (required)")
st.caption(
    "Upload a keyword.com overview screenshot to Google Drive, "
    "share as 'Anyone with the link', then paste the link below."
)
link10 = st.text_input(
    "Google Drive link",
    placeholder="https://drive.google.com/file/d/...",
    key="link10",
)
id10 = extract_drive_id(link10) if link10 else ""
if id10:
    st.success(f"File ID: `{id10}`")
    try:
        st.image(fetch_image_bytes(id10), caption="Overview screenshot preview", width=400)
    except Exception:
        st.warning("Preview unavailable.")
elif link10:
    st.error("Could not extract file ID.")

st.markdown("**Keyword CSV (optional — for ranking stats)**")
st.caption("If uploaded, the slide will display top-3 / page-1 / improved counts.")
kw_csv_for_overview = st.file_uploader(
    "Upload keyword.com CSV or Excel (optional)",
    type=["csv", "xlsx", "xls"],
    key="kw_csv_overview",
)
kw_data_for_overview = None
if kw_csv_for_overview:
    try:
        _bytes = read_as_csv(kw_csv_for_overview, "kw_ov")
        if _bytes:
            kw_data_for_overview = parse_keyword_csv(_bytes)
            st.success(
                f"Loaded: {kw_data_for_overview['total']} keywords — "
                f"Top 3: {kw_data_for_overview['top3_count']} | "
                f"Page 1: {kw_data_for_overview['top10_count']} | "
                f"Improved: {kw_data_for_overview['improved_count']}"
            )
    except Exception as e:
        st.error(f"Could not parse file: {e}")

custom_text_a = st.text_area(
    "Additional notes for this slide (optional)",
    key="kw_custom_a",
    placeholder="e.g. 本月關鍵字整體排名持續上升，特別是 foam board 和 backdrop 表現突出…",
    height=80,
)

ov_ready = bool(pres_id and id10)
if not ov_ready:
    missing = []
    if not pres_id: missing.append("presentation ID")
    if not id10:    missing.append("overview screenshot")
    if missing:
        st.warning(f"Please complete: {', '.join(missing)}")

if st.button("Insert Keywords Overview Slide", type="primary", disabled=not ov_ready, key="btn_kw_overview"):
    try:
        with st.spinner("Inserting overview slide…"):
            url_overview = drive_image_url(id10)
            build_keyword_overview_slide(
                presentation_id    = pres_id,
                kw_data            = kw_data_for_overview,
                tracking_link      = tracking_link,
                image_url_overview = url_overview,
                custom_text_a      = custom_text_a,
            )
        st.success("✓ Keywords Overview slide inserted.")
        st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)

st.divider()

# ════════════════════════════════════════════════════════════
# Section 2b — Keywords Ranking Table
# ════════════════════════════════════════════════════════════
st.subheader("Section 2b — Keywords Ranking Table")
st.caption(
    "Generates a ranking table from keyword.com CSV export. "
    "Columns: Keyword | Start | Rank | Life | Volume. "
    "Paginated automatically if there are many keywords."
)

st.markdown("**keyword.com CSV export** (required)")
st.caption("In keyword.com: select all keywords → Export → CSV.")
kw_csv_file = st.file_uploader(
    "Upload keyword.com CSV or Excel",
    type=["csv", "xlsx", "xls"],
    key="kw_csv",
)

kw_csv_preview = None
if kw_csv_file:
    try:
        _kw_bytes = read_as_csv(kw_csv_file, "kw_tbl")
        if _kw_bytes:
            kw_csv_preview = parse_keyword_csv(_kw_bytes)
            n_slides_needed = max(1, -(-len(kw_csv_preview["table_rows"]) // 10))
            st.success(
                f"Loaded: **{kw_csv_preview['total']} keywords** — "
                f"Top 3: {kw_csv_preview['top3_count']} | "
                f"Page 1: {kw_csv_preview['top10_count']} | "
                f"Improved: {kw_csv_preview['improved_count']} | "
                f"Table rows: {len(kw_csv_preview['table_rows'])} "
                f"({'1 slide' if n_slides_needed == 1 else f'{n_slides_needed} slides'})"
            )
            with st.expander("Preview ranking table data"):
                import pandas as pd
                df = pd.DataFrame(kw_csv_preview["table_rows"])[
                    ["keyword", "start", "rank", "life", "volume", "category"]
                ].rename(columns={
                    "keyword":  "Keyword",
                    "start":    "Start",
                    "rank":     "Rank",
                    "life":     "Life",
                    "volume":   "Volume",
                    "category": "Group",
                })
                st.dataframe(df, use_container_width=True)
    except Exception as e:
        st.error(f"Could not parse file: {e}")

custom_text_b = st.text_area(
    "Additional notes for ranking table slide (optional)",
    key="kw_custom_b",
    placeholder="e.g. 本月排名改善最顯著的是 printshop (+56 positions)，見證了 on-page 優化的成效…",
    height=80,
)

kw_tbl_ready = bool(pres_id and kw_csv_file and kw_csv_preview)
if not kw_tbl_ready:
    missing = []
    if not pres_id:     missing.append("presentation ID")
    if not kw_csv_file: missing.append("keyword.com CSV")
    if missing:
        st.warning(f"Please complete: {', '.join(missing)}")

if st.button("Insert Keywords Ranking Table Slides", type="primary", disabled=not kw_tbl_ready, key="btn_kw_table"):
    try:
        with st.spinner("Generating insights and inserting slides…"):
            result = build_keyword_table_slides(
                presentation_id = pres_id,
                csv_bytes       = b"",
                kw_data         = kw_csv_preview,
                custom_text_b   = custom_text_b,
            )

        n_tbl = result["table_slides"]
        st.success(
            f"✓ {n_tbl} ranking table slide{'s' if n_tbl > 1 else ''} inserted "
            f"({result['table_rows_shown']} keywords shown)."
        )
        st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)

st.divider()

# ════════════════════════════════════════════════════════════
# Section 3a — SEO Performance by Ahrefs
# ════════════════════════════════════════════════════════════
st.subheader("Section 3a — SEO Performance by Ahrefs")
st.caption(
    "One slide per analysis group — inserted at end of **'Website SEO Performance'** section. "
    "Gemini analyses screenshots and description to generate slide title and insights. "
    "Each group needs at least a description or image."
)

if "ahrefs_cards" not in st.session_state:
    st.session_state.ahrefs_cards = [{"id": 0}]
if "ahrefs_counter" not in st.session_state:
    st.session_state.ahrefs_counter = 1

def add_ahrefs_card():
    st.session_state.ahrefs_counter += 1
    st.session_state.ahrefs_cards.append({"id": st.session_state.ahrefs_counter})

ah_to_remove = []
for idx, card in enumerate(st.session_state.ahrefs_cards):
    cid = card["id"]
    with st.expander(f"Slide {idx + 1}", expanded=True):
        col_lbl, col_del = st.columns([5, 1])
        with col_lbl:
            st.caption(f"Ahrefs slide {idx + 1} — description and/or images required")
        with col_del:
            st.write("")
            if st.button("✕ Remove", key=f"ah_remove_{cid}"):
                ah_to_remove.append(idx)

        st.text_area(
            "Description (optional) — what you want to highlight on this slide",
            key=f"ah_text_{cid}",
            placeholder="e.g. Backlink profile has grown by 12% this month, with 45 new referring domains acquired…",
            height=100,
        )
        st.text_area(
            "Ahrefs screenshot Drive links (one per line)",
            key=f"ah_imgs_{cid}",
            placeholder="https://drive.google.com/file/d/...",
            height=80,
        )

        raw_ah_imgs = st.session_state.get(f"ah_imgs_{cid}", "")
        if raw_ah_imgs:
            ah_img_ids = [
                extract_drive_id(l.strip())
                for l in raw_ah_imgs.strip().splitlines() if l.strip()
            ]
            valid_ah_ids = [i for i in ah_img_ids if i]
            if valid_ah_ids:
                prev_cols = st.columns(min(len(valid_ah_ids), 3))
                for pi, fid in enumerate(valid_ah_ids[:3]):
                    with prev_cols[pi]:
                        try:
                            st.image(fetch_image_bytes(fid), use_container_width=True)
                        except Exception:
                            st.caption("Preview unavailable")

for idx in sorted(ah_to_remove, reverse=True):
    st.session_state.ahrefs_cards.pop(idx)
if ah_to_remove:
    st.rerun()

col_add_ah, _, col_ins_ah = st.columns([2, 1, 3])
with col_add_ah:
    st.button("＋ Add Another Slide", on_click=add_ahrefs_card, key="btn_add_ahrefs")

def ahrefs_cards_have_content() -> bool:
    for card in st.session_state.ahrefs_cards:
        cid = card["id"]
        if (
            st.session_state.get(f"ah_text_{cid}", "").strip()
            or st.session_state.get(f"ah_imgs_{cid}", "").strip()
        ):
            return True
    return False

ahrefs_ready = bool(pres_id and ahrefs_cards_have_content())

with col_ins_ah:
    insert_ahrefs = st.button(
        "Insert Ahrefs Analysis Slides",
        type="primary",
        disabled=not ahrefs_ready,
        key="btn_ahrefs",
    )

if insert_ahrefs:
    ah_payload = []
    ah_progress = st.progress(0, text="Preparing Ahrefs slides…")
    total_ah = len(st.session_state.ahrefs_cards)

    for i, card in enumerate(st.session_state.ahrefs_cards):
        cid  = card["id"]
        desc = st.session_state.get(f"ah_text_{cid}", "").strip()
        imgs = st.session_state.get(f"ah_imgs_{cid}", "").strip()

        if not desc and not imgs:
            continue

        ah_progress.progress((i + 0.2) / total_ah, text=f"Analysing slide {i+1}…")

        img_ids = [extract_drive_id(l.strip()) for l in imgs.splitlines() if l.strip()]
        img_ids = [fid for fid in img_ids if fid]
        img_bytes_list = []
        img_urls = []
        for fid in img_ids[:3]:
            try:
                b = fetch_image_bytes(fid)
                img_bytes_list.append(b)
                img_urls.append(drive_image_url(fid))
            except Exception as e:
                st.warning(f"Could not load image {fid}: {e}")

        with st.spinner(f"AI analysing Ahrefs slide {i+1}…"):
            try:
                result = analyze_ahrefs_slide(
                    description=desc,
                    image_bytes_list=img_bytes_list,
                )
            except Exception as e:
                result = {
                    "slide_title": "Ahrefs SEO Analysis",
                    "insight":     desc or "Ahrefs data reviewed.",
                }
                st.warning(f"AI analysis failed for slide {i+1}, using fallback: {e}")

        ah_payload.append({
            "slide_title": result.get("slide_title", "Ahrefs SEO Analysis"),
            "insight":     result.get("insight", desc),
            "image_urls":  img_urls,
            "doc_url":     "",
            "link_anchor": "",
        })
        ah_progress.progress((i + 1) / total_ah, text=f"Slide {i+1} analysed ✓")

    if ah_payload:
        with st.spinner("Inserting Ahrefs slides into presentation…"):
            try:
                n = build_ahrefs_slides(pres_id, ah_payload)
                st.success(f"✓ {n} Ahrefs slide(s) inserted.")
                st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
            except Exception as e:
                st.error(f"Error inserting slides: {e}")
                st.exception(e)
    else:
        st.warning("No valid slides to insert.")

st.divider()

# ════════════════════════════════════════════════════════════
# Section 3b — Organic Competitors (Ahrefs)
# ════════════════════════════════════════════════════════════
st.subheader("Section 3b — Organic Competitors (Ahrefs)")
st.caption(
    "Upload an Ahrefs Organic Competitors CSV to generate a comparison table. "
    "Gemini analyses the data and highlights your client's position vs competitors."
)

client_domain = st.text_input(
    "Client domain (optional — used to highlight client row in table)",
    placeholder="e.g. printcity.com.hk",
    key="oc_domain",
)

st.markdown("**Ahrefs Organic Competitors CSV** (required)")
st.caption("In Ahrefs: Site Explorer → Organic Competitors → Export CSV.")
oc_csv_file = st.file_uploader(
    "Upload Ahrefs Competitors CSV or Excel",
    type=["csv", "xlsx", "xls"],
    key="oc_csv",
)

oc_comp_data = None
if oc_csv_file:
    try:
        _oc_bytes = read_as_csv(oc_csv_file, "oc")
        if _oc_bytes:
            oc_comp_data = parse_competitor_csv(_oc_bytes)
            st.success(
                f"Loaded: **{oc_comp_data['total_found']} competitors** found "
                f"(showing top {len(oc_comp_data['rows'])})."
            )
            with st.expander("Preview competitor data"):
                import pandas as pd
                df_comp = pd.DataFrame(oc_comp_data["rows"]).rename(columns={
                    "domain":  "Domain",
                    "dr":      "DR",
                    "org_kw":  "Org. Keywords",
                    "traffic": "Org. Traffic",
                    "tv":      "Traffic Value",
                })
                st.dataframe(df_comp, use_container_width=True)
    except Exception as e:
        st.error(f"Could not parse file: {e}")

st.markdown("**Ahrefs screenshot(s) (optional)**")
st.caption("Share via Google Drive and paste links below — included in AI analysis but not always added to slide.")
oc_img_links = st.text_area(
    "Screenshot Drive links (one per line)",
    key="oc_imgs",
    placeholder="https://drive.google.com/file/d/...",
    height=70,
)

oc_desc = st.text_area(
    "Description (optional) — key points to highlight",
    key="oc_desc",
    placeholder="e.g. Our client is 3rd in organic traffic among the top 5 competitors but has the highest growth rate…",
    height=80,
)

oc_ready = bool(pres_id and oc_csv_file and oc_comp_data)
if not oc_ready:
    missing = []
    if not pres_id:      missing.append("presentation ID")
    if not oc_csv_file:  missing.append("competitors CSV")
    if missing:
        st.warning(f"Please complete: {', '.join(missing)}")

if st.button("Insert Organic Competitors Slide", type="primary", disabled=not oc_ready, key="btn_oc"):
    oc_img_bytes_list = []
    raw_oc_imgs = oc_img_links.strip()
    if raw_oc_imgs:
        oc_img_ids = [extract_drive_id(l.strip()) for l in raw_oc_imgs.splitlines() if l.strip()]
        for fid in [i for i in oc_img_ids if i][:3]:
            try:
                oc_img_bytes_list.append(fetch_image_bytes(fid))
            except Exception as e:
                st.warning(f"Could not load image {fid}: {e}")

    with st.spinner("AI analysing competitor data…"):
        try:
            analysis = analyze_organic_competitors(
                description      = oc_desc.strip(),
                image_bytes_list = oc_img_bytes_list,
                competitor_rows  = oc_comp_data["rows"],
                client_domain    = client_domain.strip(),
            )
        except Exception as e:
            analysis = {
                "slide_title": "Organic Competitors Overview",
                "insight":     oc_desc.strip() or "Competitor analysis completed via Ahrefs.",
            }
            st.warning(f"AI analysis failed, using fallback: {e}")

    with st.spinner("Inserting competitors slide…"):
        try:
            build_organic_competitors_slide(
                presentation_id = pres_id,
                comp_data = {
                    "slide_title":  analysis.get("slide_title", "Organic Competitors Overview"),
                    "insight":      analysis.get("insight", ""),
                    "rows":         oc_comp_data["rows"],
                    "client_domain": client_domain.strip(),
                },
            )
            st.success("✓ Organic Competitors slide inserted.")
            st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
        except Exception as e:
            st.error(f"Error inserting slide: {e}")
            st.exception(e)

st.divider()

# ════════════════════════════════════════════════════════════
# Section 4a — Google Search Console Analysis (Screenshots)
# ════════════════════════════════════════════════════════════
st.subheader("Section 4a — Google Search Console Analysis")
st.caption(
    "Upload GSC screenshots for AI analysis — one slide per group. "
    "Gemini identifies the GSC report type and generates slide title + insights. "
    "Inserted at end of **'Website SEO Performance'** section."
)

if "gsc_cards" not in st.session_state:
    st.session_state.gsc_cards = [{"id": 0}]
if "gsc_counter" not in st.session_state:
    st.session_state.gsc_counter = 1

def add_gsc_card():
    st.session_state.gsc_counter += 1
    st.session_state.gsc_cards.append({"id": st.session_state.gsc_counter})

gsc_to_remove = []
for idx, card in enumerate(st.session_state.gsc_cards):
    cid = card["id"]
    with st.expander(f"Slide {idx + 1}", expanded=True):
        col_lbl, col_del = st.columns([5, 1])
        with col_lbl:
            st.caption(f"GSC slide {idx + 1} — description and/or screenshots required")
        with col_del:
            st.write("")
            if st.button("✕ Remove", key=f"gsc_remove_{cid}"):
                gsc_to_remove.append(idx)

        st.text_area(
            "Description (optional) — key points you want to highlight on this slide",
            key=f"gsc_text_{cid}",
            placeholder="e.g. Overall clicks increased by 18% MoM; CTR improved on mobile after the title tag optimisation in week 2…",
            height=100,
        )
        st.text_area(
            "GSC screenshot Drive links (one per line)",
            key=f"gsc_imgs_{cid}",
            placeholder="https://drive.google.com/file/d/...",
            height=80,
        )

        raw_gsc_imgs = st.session_state.get(f"gsc_imgs_{cid}", "")
        if raw_gsc_imgs:
            gsc_img_ids = [
                extract_drive_id(l.strip())
                for l in raw_gsc_imgs.strip().splitlines() if l.strip()
            ]
            valid_gsc_ids = [i for i in gsc_img_ids if i]
            if valid_gsc_ids:
                prev_cols = st.columns(min(len(valid_gsc_ids), 3))
                for pi, fid in enumerate(valid_gsc_ids[:3]):
                    with prev_cols[pi]:
                        try:
                            st.image(fetch_image_bytes(fid), use_container_width=True)
                        except Exception:
                            st.caption("Preview unavailable")

for idx in sorted(gsc_to_remove, reverse=True):
    st.session_state.gsc_cards.pop(idx)
if gsc_to_remove:
    st.rerun()

col_add_gsc, _, col_ins_gsc = st.columns([2, 1, 3])
with col_add_gsc:
    st.button("＋ Add Another Slide", on_click=add_gsc_card, key="btn_add_gsc")

def gsc_cards_have_content() -> bool:
    for card in st.session_state.gsc_cards:
        cid = card["id"]
        if (
            st.session_state.get(f"gsc_text_{cid}", "").strip()
            or st.session_state.get(f"gsc_imgs_{cid}", "").strip()
        ):
            return True
    return False

gsc_img_ready = bool(pres_id and gsc_cards_have_content())

with col_ins_gsc:
    insert_gsc = st.button(
        "Insert GSC Analysis Slides",
        type="primary",
        disabled=not gsc_img_ready,
        key="btn_gsc_imgs",
    )

if insert_gsc:
    gsc_payload = []
    gsc_progress = st.progress(0, text="Preparing GSC slides…")
    total_gsc = len(st.session_state.gsc_cards)

    for i, card in enumerate(st.session_state.gsc_cards):
        cid  = card["id"]
        desc = st.session_state.get(f"gsc_text_{cid}", "").strip()
        imgs = st.session_state.get(f"gsc_imgs_{cid}", "").strip()

        if not desc and not imgs:
            continue

        gsc_progress.progress((i + 0.2) / total_gsc, text=f"Analysing slide {i+1}…")

        img_ids = [extract_drive_id(l.strip()) for l in imgs.splitlines() if l.strip()]
        img_ids = [fid for fid in img_ids if fid]
        img_bytes_list, img_urls = [], []
        for fid in img_ids[:3]:
            try:
                b = fetch_image_bytes(fid)
                img_bytes_list.append(b)
                img_urls.append(drive_image_url(fid))
            except Exception as e:
                st.warning(f"Could not load image {fid}: {e}")

        with st.spinner(f"AI analysing GSC slide {i+1}…"):
            try:
                result = analyze_gsc_images(
                    description=desc,
                    image_bytes_list=img_bytes_list,
                )
            except Exception as e:
                result = {
                    "slide_title": "GSC Search Performance",
                    "insight":     desc or "GSC data reviewed.",
                }
                st.warning(f"AI analysis failed for slide {i+1}, using fallback: {e}")

        gsc_payload.append({
            "slide_title": result.get("slide_title", "GSC Search Performance"),
            "insight":     result.get("insight", desc),
            "image_urls":  img_urls,
        })
        gsc_progress.progress((i + 1) / total_gsc, text=f"Slide {i+1} analysed ✓")

    if gsc_payload:
        with st.spinner("Inserting GSC slides into presentation…"):
            try:
                n = build_gsc_image_slides(pres_id, gsc_payload)
                st.success(f"✓ {n} GSC analysis slide(s) inserted.")
                st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
            except Exception as e:
                st.error(f"Error inserting slides: {e}")
                st.exception(e)
    else:
        st.warning("No valid slides to insert.")

st.divider()

# ════════════════════════════════════════════════════════════
# Section 4b — GSC Performance Data (CSV)
# ════════════════════════════════════════════════════════════
st.subheader("Section 4b — GSC Performance Data (CSV)")
st.caption(
    "Upload a GSC performance CSV (Queries, Pages, Countries, etc.). "
    "Generates a data table + AI insight. "
    "Add notes to focus analysis on specific keywords or pages."
)

st.markdown("**GSC performance CSV** (required)")
st.caption("In GSC: Performance → select date range → Export → Download CSV.")
gsc_csv_file = st.file_uploader(
    "Upload GSC CSV or Excel",
    type=["csv", "xlsx", "xls"],
    key="gsc_csv",
)

gsc_csv_data = None
if gsc_csv_file:
    try:
        _gsc_bytes = read_as_csv(gsc_csv_file, "gsc")
        if _gsc_bytes:
            gsc_csv_data = parse_gsc_csv(_gsc_bytes)
            st.success(
                f"Loaded: **{gsc_csv_data['report_type']} report** — "
                f"{gsc_csv_data['total_rows']} entries | "
                f"Total clicks: {gsc_csv_data['total_clicks']:,} | "
                f"Total impressions: {gsc_csv_data['total_impressions']:,}"
            )
            with st.expander("Preview top entries"):
                import pandas as pd
                df_gsc = pd.DataFrame(gsc_csv_data["table_rows"]).rename(columns={
                    "dimension":   gsc_csv_data["dim_col"] or "Dimension",
                    "clicks":      "Clicks",
                    "impressions": "Impressions",
                    "ctr":         "CTR",
                    "position":    "Avg. Position",
                })
                st.dataframe(df_gsc, use_container_width=True)
    except Exception as e:
        st.error(f"Could not parse file: {e}")

gsc_csv_notes = st.text_area(
    "Analysis notes (optional) — describe focus areas or specific keywords/pages to analyse",
    key="gsc_csv_notes",
    placeholder=(
        "e.g. Please focus on keywords related to 'foam board' and 'backdrop' — "
        "I want to understand their click and position trends this month. "
        "Also highlight any queries ranking in positions 11-20 that are close to page 1."
    ),
    height=100,
)

st.markdown("**Additional screenshots for AI context (optional)**")
st.caption("Screenshots are used for AI analysis only and will not appear in the slide.")
gsc_csv_imgs = st.text_area(
    "GSC screenshot Drive links (one per line, optional)",
    key="gsc_csv_imgs",
    placeholder="https://drive.google.com/file/d/...",
    height=70,
)

gsc_csv_ready = bool(pres_id and gsc_csv_file and gsc_csv_data)
if not gsc_csv_ready:
    missing = []
    if not pres_id:      missing.append("presentation ID")
    if not gsc_csv_file: missing.append("GSC CSV")
    if missing:
        st.warning(f"Please complete: {', '.join(missing)}")

if st.button("Insert GSC Data Slide", type="primary", disabled=not gsc_csv_ready, key="btn_gsc_csv"):
    # Load optional context screenshots (for Gemini only)
    ctx_img_bytes = []
    for link in gsc_csv_imgs.strip().splitlines():
        fid = extract_drive_id(link.strip())
        if fid:
            try:
                ctx_img_bytes.append(fetch_image_bytes(fid))
            except Exception as e:
                st.warning(f"Could not load context image {fid}: {e}")

    with st.spinner("AI analysing GSC data…"):
        try:
            analysis = analyze_gsc_csv(
                csv_data         = gsc_csv_data,
                user_notes       = gsc_csv_notes.strip(),
                image_bytes_list = ctx_img_bytes,
            )
        except Exception as e:
            analysis = {
                "slide_title": f"GSC {gsc_csv_data['report_type']} Performance",
                "insight":     gsc_csv_notes.strip() or "GSC performance data reviewed.",
            }
            st.warning(f"AI analysis failed, using fallback: {e}")

    with st.spinner("Inserting GSC data slide…"):
        try:
            build_gsc_csv_slide(
                presentation_id = pres_id,
                gsc_slide_data = {
                    "slide_title": analysis.get("slide_title", f"GSC {gsc_csv_data['report_type']} Performance"),
                    "insight":     analysis.get("insight", ""),
                    "dim_col":     gsc_csv_data["dim_col"],
                    "report_type": gsc_csv_data["report_type"],
                    "table_rows":  gsc_csv_data["table_rows"],
                },
            )
            st.success(
                f"✓ GSC {gsc_csv_data['report_type']} data slide inserted "
                f"(top {len(gsc_csv_data['table_rows'])} of {gsc_csv_data['total_rows']} entries shown)."
            )
            st.markdown(f"[Open presentation](https://docs.google.com/presentation/d/{pres_id}/edit)")
        except Exception as e:
            st.error(f"Error inserting slide: {e}")
            st.exception(e)

st.divider()
st.caption("SEO Report Builder · Internal Use Only")
