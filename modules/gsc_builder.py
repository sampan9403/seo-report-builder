# ============================================================
# gsc_builder.py — Google Search Console analysis slides
# ============================================================
# Section 4a: GSC screenshot slides  (P.5-7 layout, "GSC Analysis" label)
# Section 4b: GSC performance CSV    (table + Gemini insight)
# ============================================================

import io, csv, json, os, sys, uuid
from urllib.parse import urlparse
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google import genai
from google.genai import types as genai_types

from config import GEMINI_API_KEY, GEMINI_MODEL
from modules.slide_utils import (
    get_service, elem_props, frame_requests, title_requests,
    text_segments_requests, find_seo_performance_end_index,
    FRAME_P10, TITLE_P10, BODY_FONT,
    ORANGE, BLACK, WHITE, BLUE, GREEN,
)
from modules.task_detail_builder import build_task_detail_slides

# ── Colours ──────────────────────────────────────────────────
LIGHT_GREEN = {"red": 0.878, "green": 0.961, "blue": 0.875}   # avg position 1-3
LIGHT_BLUE  = {"red": 0.878, "green": 0.937, "blue": 0.961}   # avg position 4-10
LIGHT_GREY  = {"red": 0.95,  "green": 0.95,  "blue": 0.95}    # alternating rows

# ── Table layout (same x,y base as keywords for consistency) ─
# 5 columns: Dimension | Clicks | Impressions | CTR | Avg.Pos
# Total width ≈ 421pt to fit left panel of FRAME_P10
GSC_COL_WIDTHS = [
    int(175 * 12700),   # Dimension  = 2222500
    int(52  * 12700),   # Clicks     =  660400
    int(62  * 12700),   # Impressions=  787400
    int(55  * 12700),   # CTR        =  698500
    int(77  * 12700),   # Avg.Pos    =  977900
]                       # Total      = 5346700 EMU ≈ 421pt
_GSC_COL_SUM = sum(GSC_COL_WIDTHS)

GSC_TABLE = {
    "x": 630425,  "y": 1490188,
    "w": _GSC_COL_SUM,
    "h": int(220 * 12700),   # overridden dynamically
}
_GSC_AVAIL_H = (FRAME_P10["y"] + FRAME_P10["h"]) - GSC_TABLE["y"]  # 2960053 EMU ≈ 233pt

GSC_INSIGHT = {
    "x": 5914650,  "y": 1709850,   # right panel, same as keyword table insight
    "w": int(213 * 12700),
    "h": int(205 * 12700),
}

MAX_GSC_TABLE_ROWS = 10   # per slide (same as keywords to avoid overflow)
MAX_GSC_CONTEXT_ROWS = 50  # rows sent to Gemini for analysis

# ── Report type detection ─────────────────────────────────────
REPORT_TYPE_MAP = {
    "query":             "Queries",
    "top pages":         "Pages",
    "page":              "Pages",
    "landing page":      "Pages",
    "country":           "Countries",
    "device":            "Devices",
    "date":              "Dates",
    "search type":       "Search Type",
    "search appearance": "Search Appearance",
}


# ── Internal helpers ──────────────────────────────────────────

def _get_gemini_key() -> str:
    key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY", "")
    if not key:
        try:
            import streamlit as st
            key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            pass
    return key


def _gemini_client():
    return genai.Client(api_key=_get_gemini_key())


def _find_seo_end(slides: list) -> int:
    idx = find_seo_performance_end_index(slides)
    if idx is None:
        raise ValueError(
            "Could not find 'Website SEO Performance' header slide. "
            "Ensure the presentation has that section header."
        )
    return idx


def _fmt_dim(dim_val: str, report_type: str) -> str:
    """Truncate dimension value to prevent row height expansion."""
    if report_type == "Pages":
        try:
            path = urlparse(dim_val).path or dim_val
        except Exception:
            path = dim_val
        return (path[:31] + "…") if len(path) > 32 else path
    elif report_type == "Queries":
        return (dim_val[:27] + "…") if len(dim_val) > 28 else dim_val
    else:
        return (dim_val[:21] + "…") if len(dim_val) > 22 else dim_val


# ── CSV Parsing ──────────────────────────────────────────────

def parse_gsc_csv(csv_bytes: bytes) -> dict:
    """
    Parse a Google Search Console CSV export (Queries / Pages / Countries / etc.).
    Auto-detects report type from the first column header.
    """
    text    = csv_bytes.decode("utf-8-sig")
    reader  = csv.DictReader(io.StringIO(text))
    headers = reader.fieldnames or []

    first_col   = (headers[0] if headers else "").lower().strip()
    report_type = "Queries"
    for key, val in REPORT_TYPE_MAP.items():
        if key in first_col:
            report_type = val
            break

    dim_col = headers[0] if headers else ""

    def find_col(aliases: list) -> str:
        for alias in aliases:
            for h in headers:
                if alias.lower() in h.lower():
                    return h
        return ""

    col_clicks = find_col(["clicks"])
    col_impr   = find_col(["impressions", "impr."])
    col_ctr    = find_col(["ctr"])
    col_pos    = find_col(["position", "avg. position"])

    def safe_int(s: str) -> int:
        try:
            return int(str(s).replace(",", "").replace(" ", "") or 0)
        except (ValueError, TypeError):
            return 0

    rows = []
    for row in reader:
        dim_val = (row.get(dim_col) or "").strip()
        if not dim_val:
            continue

        def cell(col):
            return (row.get(col) or "").strip() if col else "—"

        rows.append({
            "dimension":   dim_val,
            "clicks":      cell(col_clicks),
            "impressions": cell(col_impr),
            "ctr":         cell(col_ctr),
            "position":    cell(col_pos),
        })

    rows.sort(key=lambda x: safe_int(x["clicks"]), reverse=True)

    total_clicks = sum(safe_int(r["clicks"]) for r in rows)
    total_impr   = sum(safe_int(r["impressions"]) for r in rows)

    return {
        "report_type":       report_type,
        "dim_col":           dim_col,
        "rows":              rows,
        "table_rows":        rows[:MAX_GSC_TABLE_ROWS],
        "context_rows":      rows[:MAX_GSC_CONTEXT_ROWS],
        "total_rows":        len(rows),
        "total_clicks":      total_clicks,
        "total_impressions": total_impr,
        "top_items":         rows[:5],
    }


# ════════════════════════════════════════════════════════════
# Section 4a — GSC Screenshot Analysis
# ════════════════════════════════════════════════════════════

def analyze_gsc_images(description: str, image_bytes_list: list) -> dict:
    """
    Gemini analysis for one GSC screenshot slide.
    Returns {slide_title, insight}.
    """
    context = []
    if description:
        context.append(f"User notes: {description}")
    if image_bytes_list:
        context.append(f"{len(image_bytes_list)} GSC screenshot(s) attached.")

    prompt = f"""You are an SEO consultant reviewing Google Search Console data for a client monthly report.

{chr(10).join(context)}

The screenshot(s) are from Google Search Console and may show:
- Performance overview (clicks, impressions, CTR, average position trends)
- Coverage / Indexing status and errors
- Core Web Vitals (LCP, INP, CLS) results
- Page experience signals
- Mobile usability issues
- Rich results / Structured data enhancements
- URL inspection results
- Or any other GSC report

Based on what is shown:
1. Identify which GSC report/feature is displayed
2. Extract the key metrics, trends, and anomalies
3. Write professional content suitable for an SEO monthly client report

Return ONLY valid JSON — no markdown:
{{
  "slide_title": "Short title, max 8 words (e.g. 'Search Performance Overview', 'Core Web Vitals Status', 'Index Coverage Report')",
  "insight": "3-4 professional sentences covering key findings, notable metrics, and their SEO implications. Be specific with numbers where visible."
}}"""

    parts = [genai_types.Part.from_text(text=prompt)]
    for img in image_bytes_list:
        parts.append(genai_types.Part.from_bytes(data=img, mime_type="image/png"))

    try:
        resp = _gemini_client().models.generate_content(
            model=GEMINI_MODEL,
            contents=[genai_types.Content(role="user", parts=parts)],
        )
        raw = resp.text.strip()
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:-1])
        return json.loads(raw)
    except Exception:
        return {
            "slide_title": "GSC Search Performance",
            "insight": description or "Google Search Console data reviewed as part of the monthly SEO engagement.",
        }


def build_gsc_image_slides(presentation_id: str, slide_groups: list) -> int:
    """
    Insert one P.5-7 style slide per group at end of 'Website SEO Performance' section.

    Each slide_group:
    {
        "slide_title":  str,
        "insight":      str,
        "image_urls":   [str],   # public Drive image URLs
        "doc_url":      str,     # optional
        "link_anchor":  str,     # optional
    }
    """
    service = get_service()
    pres    = service.presentations().get(presentationId=presentation_id).execute()
    slides  = pres["slides"]
    ins_at  = _find_seo_end(slides)

    tasks = [
        {
            "name":        g.get("slide_title", "GSC Analysis"),
            "slide_title": g.get("slide_title", "GSC Analysis"),
            "insight":     g.get("insight", ""),
            "image_urls":  g.get("image_urls", []),
            "doc_url":     g.get("doc_url", ""),
            "link_anchor": g.get("link_anchor", ""),
        }
        for g in slide_groups
    ]

    return build_task_detail_slides(
        presentation_id=presentation_id,
        tasks=tasks,
        section_label="GSC Analysis",
        insert_at=ins_at,
    )


# ════════════════════════════════════════════════════════════
# Section 4b — GSC Performance Data (CSV)
# ════════════════════════════════════════════════════════════

def analyze_gsc_csv(
    csv_data: dict,
    user_notes: str = "",
    image_bytes_list: list = None,
) -> dict:
    """
    Gemini insight from GSC CSV data + optional user notes/images.
    If user notes mention specific keywords/pages, Gemini will focus on those.
    Returns {slide_title, insight}.
    """
    image_bytes_list = image_bytes_list or []
    report_type = csv_data["report_type"]
    fallback = {
        "slide_title": f"GSC {report_type} Performance",
        "insight": user_notes or f"GSC {report_type} data reviewed as part of the monthly SEO engagement.",
    }

    if not _get_gemini_key():
        return fallback

    context_rows = csv_data.get("context_rows", csv_data.get("rows", []))

    prompt = f"""You are an SEO consultant writing a slide insight for a client monthly SEO report.

Google Search Console {report_type} report:
- Report type: {report_type}
- Total entries: {csv_data['total_rows']}
- Total clicks: {csv_data['total_clicks']:,}
- Total impressions: {csv_data['total_impressions']:,}

{"User focus / notes: " + user_notes if user_notes else "General performance analysis requested."}

Data (sorted by clicks, top {len(context_rows)} entries):
{json.dumps(context_rows, indent=2)}

Write a professional insight (3-4 sentences) that:
- Summarises overall search performance (clicks, impressions)
- Highlights the top performers by clicks
{"- Specifically analyses any keywords or pages the user mentioned if they appear in the data" if user_notes else ""}
- Provides at least one actionable SEO observation

Return ONLY valid JSON:
{{
  "slide_title": "GSC {report_type} Performance",
  "insight": "..."
}}"""

    parts = [genai_types.Part.from_text(text=prompt)]
    for img in image_bytes_list:
        parts.append(genai_types.Part.from_bytes(data=img, mime_type="image/png"))

    try:
        resp = _gemini_client().models.generate_content(
            model=GEMINI_MODEL,
            contents=[genai_types.Content(role="user", parts=parts)],
        )
        raw = resp.text.strip()
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:-1])
        return json.loads(raw)
    except Exception:
        return fallback


def _gsc_table_requests(
    prefix: str,
    slide_id: str,
    dim_label: str,
    report_type: str,
    table_rows: list,
) -> list:
    """Build createTable + all cell styling for one GSC data slide."""
    table_id = f"{prefix}_gsc_tbl"
    n_data   = len(table_rows)
    n_rows   = n_data + 1

    row_h   = min(int(20 * 12700), _GSC_AVAIL_H // max(n_rows, 1))
    table_h = row_h * n_rows
    table_pos = {**GSC_TABLE, "h": table_h}
    reqs = []

    reqs.append({"createTable": {
        "objectId": table_id,
        "elementProperties": elem_props(slide_id, table_pos),
        "rows":    n_rows,
        "columns": 5,
    }})

    for ci, w in enumerate(GSC_COL_WIDTHS):
        reqs.append({"updateTableColumnProperties": {
            "objectId": table_id,
            "columnIndices": [ci],
            "tableColumnProperties": {"columnWidth": {"magnitude": w, "unit": "EMU"}},
            "fields": "columnWidth",
        }})

    for ri in range(n_rows):
        reqs.append({"updateTableRowProperties": {
            "objectId": table_id,
            "rowIndices": [ri],
            "tableRowProperties": {"minRowHeight": {"magnitude": row_h, "unit": "EMU"}},
            "fields": "minRowHeight",
        }})

    # Header labels
    dim_header = (dim_label[:9] + "…") if len(dim_label) > 10 else dim_label
    col_headers = [dim_header, "Clicks", "Impr.", "CTR", "Avg.Pos"]

    for ci, label in enumerate(col_headers):
        reqs.append({"insertText": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": ci},
            "text": label, "insertionIndex": 0,
        }})
        reqs.append({"updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 0, "columnIndex": ci},
            "textRange": {"type": "ALL"},
            "style": {
                "fontFamily": BODY_FONT,
                "fontSize": {"magnitude": 8, "unit": "PT"},
                "foregroundColor": {"opaqueColor": {"rgbColor": WHITE}},
                "bold": True,
            },
            "fields": "fontFamily,fontSize,foregroundColor,bold",
        }})
        reqs.append({"updateTableCellProperties": {
            "objectId": table_id,
            "tableRange": {"location": {"rowIndex": 0, "columnIndex": ci}, "rowSpan": 1, "columnSpan": 1},
            "tableCellProperties": {
                "tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": ORANGE}, "alpha": 1}},
                "contentAlignment": "MIDDLE",
            },
            "fields": "tableCellBackgroundFill,contentAlignment",
        }})

    # Data rows — colour by average position
    for ri, row in enumerate(table_rows):
        actual_ri = ri + 1

        try:
            pos_val = float(row.get("position", "999").replace(",", "") or 999)
        except (ValueError, TypeError):
            pos_val = 999

        row_bg = (LIGHT_GREEN if pos_val <= 3
                  else LIGHT_BLUE if pos_val <= 10
                  else LIGHT_GREY if ri % 2 == 0
                  else WHITE)

        dim_disp = _fmt_dim(row.get("dimension", ""), report_type)

        cells = [
            (dim_disp,                    BLACK, False),
            (row.get("clicks", "—"),      BLACK, True),
            (row.get("impressions", "—"), BLACK, False),
            (row.get("ctr", "—"),         BLACK, False),
            (row.get("position", "—"),    BLACK, False),
        ]

        for ci, (text, color, bold) in enumerate(cells):
            reqs.append({"updateTableCellProperties": {
                "objectId": table_id,
                "tableRange": {"location": {"rowIndex": actual_ri, "columnIndex": ci}, "rowSpan": 1, "columnSpan": 1},
                "tableCellProperties": {
                    "tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": row_bg}, "alpha": 1}},
                    "contentAlignment": "MIDDLE",
                },
                "fields": "tableCellBackgroundFill,contentAlignment",
            }})
            if text:
                reqs.append({"insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": actual_ri, "columnIndex": ci},
                    "text": text, "insertionIndex": 0,
                }})
                reqs.append({"updateTextStyle": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": actual_ri, "columnIndex": ci},
                    "textRange": {"type": "ALL"},
                    "style": {
                        "fontFamily": BODY_FONT,
                        "fontSize": {"magnitude": 7, "unit": "PT"},
                        "foregroundColor": {"opaqueColor": {"rgbColor": color}},
                        "bold": bold,
                    },
                    "fields": "fontFamily,fontSize,foregroundColor,bold",
                }})

    return reqs


def build_gsc_csv_slide(presentation_id: str, gsc_slide_data: dict) -> int:
    """
    Insert one GSC data slide at end of 'Website SEO Performance' section.

    gsc_slide_data:
    {
        "slide_title":  str,
        "insight":      str,
        "dim_col":      str,    # raw column name from CSV (e.g. "Query", "Top pages")
        "report_type":  str,    # detected type  (e.g. "Queries", "Pages")
        "table_rows":   list,
    }
    """
    service  = get_service()
    pres     = service.presentations().get(presentationId=presentation_id).execute()
    slides   = pres["slides"]
    ins_idx  = _find_seo_end(slides)

    run_id   = uuid.uuid4().hex[:8]
    slide_id = f"seo_gsc_{run_id}"
    pfx      = f"seo_gsc_{run_id}"
    reqs     = []

    reqs.append({"createSlide": {
        "objectId": slide_id,
        "insertionIndex": ins_idx,
        "slideLayoutReference": {"predefinedLayout": "BLANK"},
    }})
    reqs.extend(frame_requests(pfx, slide_id, FRAME_P10))
    reqs.extend(title_requests(pfx, slide_id, gsc_slide_data.get("slide_title", "GSC Performance"), TITLE_P10))

    # Table (if rows available)
    table_rows = gsc_slide_data.get("table_rows", [])
    if table_rows:
        reqs.extend(_gsc_table_requests(
            pfx, slide_id,
            gsc_slide_data.get("dim_col", "Query"),
            gsc_slide_data.get("report_type", "Queries"),
            table_rows,
        ))

    # Insight on right panel
    insight = gsc_slide_data.get("insight", "")
    if insight:
        reqs.append({"createShape": {
            "objectId": f"{pfx}_insight",
            "shapeType": "TEXT_BOX",
            "elementProperties": elem_props(slide_id, GSC_INSIGHT),
        }})
        reqs.extend(text_segments_requests(f"{pfx}_insight", [
            (insight, {"fontFamily": BODY_FONT, "fontSize": 10, "bold": False, "color": BLACK}),
        ]))

    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": reqs},
    ).execute()

    return 1
