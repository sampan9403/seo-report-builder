# ============================================================
# ahrefs_builder.py — SEO Performance by Ahrefs slides
# ============================================================
# Section 3a: General Ahrefs analysis slides (reuses P.5-7 layout)
# Section 3b: Organic Competitors comparison table
# ============================================================

import io, csv, json, os, sys, uuid
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google import genai
from google.genai import types as genai_types

from config import GEMINI_API_KEY, GEMINI_MODEL
from modules.slide_utils import (
    get_service, elem_props, frame_requests, title_requests,
    text_segments_requests, find_seo_performance_end_index,
    FRAME_P10, TITLE_P10, BODY_FONT,
    ORANGE, BLACK, WHITE, BLUE, GREEN, RED,
)
from modules.task_detail_builder import build_task_detail_slides


# ── Colour helpers ────────────────────────────────────────────
LIGHT_ORANGE = {"red": 1.0,   "green": 0.949, "blue": 0.933}
LIGHT_GREY   = {"red": 0.95,  "green": 0.95,  "blue": 0.95}
HIGHLIGHT    = {"red": 0.878, "green": 0.937, "blue": 0.961}

# ── Competitor table layout ───────────────────────────────────
# Full-width table within FRAME_P10, enough room for insight below.
COMP_COL_WIDTHS = [
    int(200 * 12700),   # Domain
    int(60  * 12700),   # DR
    int(130 * 12700),   # Org. Keywords
    int(130 * 12700),   # Org. Traffic
    int(110 * 12700),   # Traffic Value
]
_COMP_COL_SUM = sum(COMP_COL_WIDTHS)  # 630pt total

COMP_TABLE = {
    "x": 541275,  "y": 1490188,
    "w": _COMP_COL_SUM,
    "h": int(180 * 12700),   # dynamic, overridden in builder
}
_COMP_AVAIL_H = (FRAME_P10["y"] + FRAME_P10["h"]) - COMP_TABLE["y"]  # ~2960053 EMU

COMP_INSIGHT = {
    "x": 541275,   "y": int(315 * 12700),
    "w": _COMP_COL_SUM, "h": int(32 * 12700),
}
COMP_HEADERS = ["Domain", "DR", "Org. Keywords", "Org. Traffic", "Traffic Value"]

MAX_COMP_ROWS = 10


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


# ════════════════════════════════════════════════════════════
# Section 3a — General Ahrefs Analysis
# ════════════════════════════════════════════════════════════

def analyze_ahrefs_slide(
    description: str,
    image_bytes_list: list,
) -> dict:
    """
    Use Gemini to generate slide title + insight for one Ahrefs slide.
    Returns: {slide_title, insight}
    """
    client = _gemini_client()

    context_parts = []
    if description:
        context_parts.append(f"User notes: {description}")
    if image_bytes_list:
        context_parts.append(f"{len(image_bytes_list)} Ahrefs screenshot(s) attached.")

    prompt = f"""You are an SEO consultant writing content for a client monthly report slide.

{chr(10).join(context_parts)}

The screenshot(s) are from the Ahrefs SEO tool. They may show backlink profiles, organic search performance, site audit results, keyword rankings, competitor data, or other Ahrefs features.

1. Identify which Ahrefs feature/report is shown.
2. Extract the most important metrics and trends.
3. Write professional slide content appropriate for an SEO monthly report.

Return ONLY valid JSON — no markdown, no explanation:
{{
  "slide_title": "Short descriptive title, max 8 words (e.g. 'Backlink Profile Overview', 'Organic Traffic Growth')",
  "insight": "3-4 professional sentences covering key findings, notable metrics, and their SEO implications. Be specific with numbers where visible."
}}"""

    parts = [genai_types.Part.from_text(text=prompt)]
    for img_bytes in image_bytes_list:
        parts.append(genai_types.Part.from_bytes(data=img_bytes, mime_type="image/png"))

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[genai_types.Content(role="user", parts=parts)],
    )
    raw = response.text.strip()
    if raw.startswith("```"):
        raw = "\n".join(raw.split("\n")[1:-1])
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {
            "slide_title": "Ahrefs SEO Analysis",
            "insight": description or "Ahrefs data reviewed as part of the monthly SEO engagement.",
        }


def build_ahrefs_slides(
    presentation_id: str,
    slide_groups: list,
) -> int:
    """
    Insert one P.5-7 style slide per group at end of 'Website SEO Performance' section.

    Each slide_group:
    {
        "slide_title":  str,       # from analyze_ahrefs_slide
        "insight":      str,
        "image_urls":   [str],     # public Drive image URLs
        "doc_url":      str,       # optional
        "link_anchor":  str,       # optional
    }
    """
    service = get_service()
    pres    = service.presentations().get(presentationId=presentation_id).execute()
    slides  = pres["slides"]
    ins_at  = _find_seo_end(slides)

    tasks = [
        {
            "name":        g.get("slide_title", "SEO Analysis"),
            "slide_title": g.get("slide_title", "SEO Analysis"),
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
        section_label="SEO Performance",
        insert_at=ins_at,
    )


# ════════════════════════════════════════════════════════════
# Section 3b — Organic Competitors
# ════════════════════════════════════════════════════════════

def parse_competitor_csv(csv_bytes: bytes) -> dict:
    """
    Parse an Ahrefs Organic Competitors CSV export.
    Flexibly maps common column name variations.
    Returns: {headers, rows, raw_headers}
    """
    text   = csv_bytes.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    raw_headers = reader.fieldnames or []

    def find_col(priorities: list) -> str:
        for p in priorities:
            for h in raw_headers:
                if p.lower() in h.lower():
                    return h
        return ""

    col_domain  = find_col(["url", "target", "domain", "website"])
    col_dr      = find_col(["domain rating", " dr", "dr "])
    col_org_kw  = find_col(["organic keyword", "org. keyword", "keywords"])
    col_traffic = find_col(["organic traffic", "org. traffic", "traffic"])
    col_tv      = find_col(["traffic value", "org. traffic value"])

    rows = []
    for row in reader:
        def val(col):
            if not col:
                return ""
            v = (row.get(col) or "").strip()
            return v

        domain = val(col_domain)
        if not domain:
            continue

        def fmt_num(col):
            v = val(col).replace(",", "").replace(" ", "")
            try:
                n = int(float(v))
                return f"{n:,}" if n else "—"
            except (ValueError, TypeError):
                return v or "—"

        rows.append({
            "domain":   domain,
            "dr":       val(col_dr) or "—",
            "org_kw":   fmt_num(col_org_kw),
            "traffic":  fmt_num(col_traffic),
            "tv":       fmt_num(col_tv),
        })

    return {
        "rows":        rows[:MAX_COMP_ROWS],
        "total_found": len(rows),
        "raw_headers": raw_headers,
    }


def analyze_organic_competitors(
    description: str,
    image_bytes_list: list,
    competitor_rows: list,
    client_domain: str,
) -> dict:
    """
    Gemini analysis of competitor data; returns {slide_title, insight}.
    Falls back gracefully if API unavailable.
    """
    key = _get_gemini_key()
    if not key:
        return {
            "slide_title": "Organic Competitors Overview",
            "insight":     description or "Competitor analysis completed via Ahrefs.",
        }

    client = _gemini_client()

    rows_summary = json.dumps(competitor_rows[:8], indent=2)
    context = f"Client domain: {client_domain}\n" if client_domain else ""
    if description:
        context += f"User notes: {description}\n"

    prompt = f"""You are an SEO consultant writing a competitor analysis for a monthly report.

{context}
Ahrefs Organic Competitors data (top entries):
{rows_summary}

{"Ahrefs screenshot(s) attached." if image_bytes_list else ""}

Write a professional competitor analysis (3-4 sentences) that:
- Identifies the client's current competitive position
- Highlights key gaps or opportunities vs top competitors
- References specific domains and metrics where relevant
- Maintains a factual, consultant tone

Return ONLY valid JSON:
{{
  "slide_title": "Organic Competitors Overview",
  "insight": "..."
}}"""

    parts = [genai_types.Part.from_text(text=prompt)]
    for img_bytes in image_bytes_list:
        parts.append(genai_types.Part.from_bytes(data=img_bytes, mime_type="image/png"))

    try:
        response = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[genai_types.Content(role="user", parts=parts)],
        )
        raw = response.text.strip()
        if raw.startswith("```"):
            raw = "\n".join(raw.split("\n")[1:-1])
        return json.loads(raw)
    except Exception:
        return {
            "slide_title": "Organic Competitors Overview",
            "insight":     description or "Competitor analysis completed via Ahrefs.",
        }


def _competitor_table_requests(
    prefix: str,
    slide_id: str,
    rows: list,
    client_domain: str,
) -> list:
    """Build createTable + all cell requests for the competitor table."""
    table_id = f"{prefix}_comp_table"
    n_data   = len(rows)
    n_rows   = n_data + 1

    row_h  = min(int(19 * 12700), _COMP_AVAIL_H // max(n_rows, 1))
    table_h = row_h * n_rows
    table_pos = {**COMP_TABLE, "h": table_h}
    reqs = []

    reqs.append({"createTable": {
        "objectId": table_id,
        "elementProperties": elem_props(slide_id, table_pos),
        "rows":    n_rows,
        "columns": 5,
    }})

    for ci, w in enumerate(COMP_COL_WIDTHS):
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

    # Header
    for ci, label in enumerate(COMP_HEADERS):
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

    # Data rows
    client_domain_clean = (client_domain or "").lower().strip().rstrip("/")
    for ri, row in enumerate(rows):
        actual_ri  = ri + 1
        domain_val = row.get("domain", "")
        is_client  = client_domain_clean and client_domain_clean in domain_val.lower()
        row_bg     = HIGHLIGHT if is_client else (LIGHT_GREY if ri % 2 == 0 else WHITE)

        domain_disp = domain_val
        if len(domain_disp) > 28:
            domain_disp = domain_disp[:27] + "…"

        cells = [
            (domain_disp,          BLACK if not is_client else ORANGE, is_client),
            (row.get("dr", "—"),   BLACK, False),
            (row.get("org_kw", "—"), BLACK, False),
            (row.get("traffic", "—"), BLACK, False),
            (row.get("tv", "—"),   BLACK, False),
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


def build_organic_competitors_slide(
    presentation_id: str,
    comp_data: dict,
) -> int:
    """
    Insert one organic competitors slide at end of 'Website SEO Performance' section.

    comp_data:
    {
        "slide_title":  str,
        "insight":      str,
        "rows":         list,   # from parse_competitor_csv
        "client_domain": str,
    }
    """
    service = get_service()
    pres    = service.presentations().get(presentationId=presentation_id).execute()
    slides  = pres["slides"]
    ins_idx = _find_seo_end(slides)

    run_id   = uuid.uuid4().hex[:8]
    slide_id = f"seo_comp_{run_id}"
    pfx      = f"seo_comp_{run_id}"
    reqs     = []

    title_text   = comp_data.get("slide_title", "Organic Competitors Overview")
    insight_text = comp_data.get("insight", "")
    rows         = comp_data.get("rows", [])
    client_domain = comp_data.get("client_domain", "")

    reqs.append({"createSlide": {
        "objectId": slide_id,
        "insertionIndex": ins_idx,
        "slideLayoutReference": {"predefinedLayout": "BLANK"},
    }})
    reqs.extend(frame_requests(pfx, slide_id, FRAME_P10))
    reqs.extend(title_requests(pfx, slide_id, title_text, TITLE_P10))

    # Competitor table
    reqs.extend(_competitor_table_requests(pfx, slide_id, rows, client_domain))

    # Insight text below table
    if insight_text:
        reqs.append({"createShape": {
            "objectId": f"{pfx}_insight",
            "shapeType": "TEXT_BOX",
            "elementProperties": elem_props(slide_id, COMP_INSIGHT),
        }})
        reqs.extend(text_segments_requests(f"{pfx}_insight", [
            (insight_text, {"fontFamily": BODY_FONT, "fontSize": 10, "bold": False, "color": BLACK}),
        ]))

    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": reqs},
    ).execute()

    return 1
