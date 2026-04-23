# ============================================================
# keywords_builder.py — Target Keywords Performance slides
# ============================================================

import io, csv, sys, os, uuid, json
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from modules.slide_utils import (
    get_service, elem_props, frame_requests, title_requests,
    text_segments_requests, find_header_index, find_seo_performance_end_index,
    FRAME_P10, TITLE_P10, BODY_FONT,
    ORANGE, BLACK, BLUE, GREEN, RED, DARK_BLUE, WHITE,
)
from config import GEMINI_API_KEY, GEMINI_MODEL

# ── Colours ──────────────────────────────────────────────────
LIGHT_GREEN  = {"red": 0.878, "green": 0.961, "blue": 0.875}
LIGHT_BLUE   = {"red": 0.878, "green": 0.937, "blue": 0.961}
PALE_ORANGE  = {"red": 1.0,   "green": 0.949, "blue": 0.933}
TEXT_GREEN   = {"red": 0.0,   "green": 0.55,  "blue": 0.2}
TEXT_RED     = {"red": 0.85,  "green": 0.1,   "blue": 0.1}

# ── Pagination ────────────────────────────────────────────────
MAX_TABLE_ROWS     = 20   # total rows across all slides
MAX_ROWS_PER_SLIDE = 10   # rows per table slide

# ── Overview slide elements ───────────────────────────────────
S_OVERVIEW_STATS = {
    "x": 4400896, "y": 1788675,
    "w": int(329.08 * 12700), "h": int(187.87 * 12700),
}
S_OVERVIEW_LINK = {
    "x": 681300, "y": 1330875,
    "w": int(582.33 * 12700), "h": int(31.51 * 12700),
}
S_OVERVIEW_IMAGE = {
    "x": 632475, "y": 2123950,
    "w": int(290.73 * 12700), "h": int(134.37 * 12700),
}

# ── Table slide elements ──────────────────────────────────────
# Table fits left ~66% of the frame; legend+insight occupy the right third.
# S_TABLE y=117.3pt; FRAME_P10 bottom=350.4pt → 233pt available.
# Row height capped at 20pt so 11 rows (10+header) = 220pt — safely within frame.

# Column widths must be set before using in S_TABLE w
TABLE_COL_WIDTHS = [
    int(183 * 12700),   # Keyword  = 2324100
    int(48  * 12700),   # Start    =  609600
    int(48  * 12700),   # Rank     =  609600
    int(48  * 12700),   # Life     =  609600
    int(94  * 12700),   # Volume   = 1193800
]                       # Total    = 5346700 EMU ≈ 421pt
TABLE_COL_HEADERS = ["Keyword", "Start", "Rank", "Life", "Volume"]

_COL_W_SUM = sum(TABLE_COL_WIDTHS)   # 5346700 EMU

S_TABLE = {
    "x": 630425,  "y": 1490188,
    "w": _COL_W_SUM,
    "h": int(220 * 12700),  # container height (row heights calculated dynamically)
}
# Available vertical space from table top to frame bottom (for row height calc)
_TABLE_AVAILABLE_H = (FRAME_P10["y"] + FRAME_P10["h"]) - S_TABLE["y"]  # 2960053 EMU

S_LEGEND = {
    "x": 5976802, "y": 1709850,
    "w": int(139.46 * 12700), "h": int(90 * 12700),
}
S_TABLE_INSIGHT = {
    "x": 5914650, "y": 2477675,
    "w": int(213.94 * 12700), "h": int(150 * 12700),
}

MAX_KW_CHARS = 25   # truncate long keywords to prevent row height expansion


# ── Helpers ──────────────────────────────────────────────────

def _get_gemini_key() -> str:
    key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY", "")
    if not key:
        try:
            import streamlit as st
            key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            pass
    return key


def _find_seo_insert_index(slides: list) -> int:
    idx = find_seo_performance_end_index(slides)
    if idx is None:
        raise ValueError(
            "Could not find 'Website SEO Performance' header slide. "
            "Ensure the presentation has that section header."
        )
    return idx


# ── CSV Parsing ──────────────────────────────────────────────

def parse_keyword_csv(csv_bytes: bytes) -> dict:
    text   = csv_bytes.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))

    keywords = []
    for row in reader:
        def safe_int(key, default=0):
            try:
                return int(float(row.get(key, default) or default))
            except (ValueError, TypeError):
                return default

        kw = (row.get("Keyword") or "").strip()
        if not kw:
            continue
        keywords.append({
            "keyword": kw,
            "rank":    safe_int("Google", -1),
            "life":    safe_int("Life",    0),
            "start":   safe_int("Start",   0),
            "volume":  safe_int("Search Volume", 0),
            "week":    safe_int("Week",    0),
        })

    ranked   = [k for k in keywords if k["rank"] > 0]
    top3     = sorted([k for k in ranked if k["rank"] <= 3],        key=lambda x: (x["rank"], -x["volume"]))
    top4_10  = sorted([k for k in ranked if 4 <= k["rank"] <= 10],  key=lambda x: (x["rank"], -x["volume"]))
    top11_20 = [k for k in ranked if 11 <= k["rank"] <= 20]
    top21_30 = [k for k in ranked if 21 <= k["rank"] <= 30]
    improved = sorted([k for k in keywords if k["life"] > 0],  key=lambda x: -x["life"])
    maintained = [k for k in ranked  if k["life"] == 0]
    dropped  = [k for k in keywords if k["life"] < 0]

    table_rows, seen = [], set()

    for k in sorted(top3, key=lambda x: -x["volume"])[:8]:
        table_rows.append({**k, "category": "top3"})
        seen.add(k["keyword"])

    for k in sorted(top4_10, key=lambda x: -x["volume"])[:8]:
        if k["keyword"] not in seen:
            table_rows.append({**k, "category": "top10"})
            seen.add(k["keyword"])

    for k in improved[:4]:
        if k["keyword"] not in seen and len(table_rows) < MAX_TABLE_ROWS:
            table_rows.append({**k, "category": "improved"})
            seen.add(k["keyword"])

    table_rows = table_rows[:MAX_TABLE_ROWS]

    return {
        "total":            len(keywords),
        "top3_count":       len(top3),
        "top10_count":      len(top3) + len(top4_10),
        "top20_count":      len(top3) + len(top4_10) + len(top11_20),
        "top30_count":      len(top3) + len(top4_10) + len(top11_20) + len(top21_30),
        "improved_count":   len(improved),
        "maintained_count": len(maintained),
        "dropped_count":    len(dropped),
        "table_rows":       table_rows,
        "top3_samples":     sorted(top3,    key=lambda x: -x["volume"])[:5],
        "top10_samples":    sorted(top4_10, key=lambda x: -x["volume"])[:5],
        "improved_samples": improved[:5],
    }


# ── Insight generation ────────────────────────────────────────

def _computed_insight(kw_data: dict) -> str:
    top3     = kw_data["top3_count"]
    top10    = kw_data["top10_count"]
    total    = kw_data["total"]
    improved = kw_data["improved_count"]

    top3_names = [k["keyword"] for k in kw_data["top3_samples"][:3]]
    imp_names  = [k["keyword"] for k in kw_data["improved_samples"][:3]]

    parts = []
    if top3 > 0:
        kws = ", ".join(f'"{k}"' for k in top3_names)
        parts.append(f"{top3} keyword{'s' if top3>1 else ''} ranked in top 3 ({kws}).")
    if top10 > top3:
        parts.append(f"{top10} out of {total} keywords are ranked on page 1.")
    if improved > 0:
        kws = ", ".join(f'"{k}"' for k in imp_names)
        parts.append(
            f"{improved} keyword{'s' if improved>1 else ''} improved in ranking "
            f"since tracking started, including {kws}."
        )
    return " ".join(parts) if parts else "Keyword rankings tracked as part of the SEO engagement."


def generate_keyword_insight(kw_data: dict, custom_text: str = "") -> str:
    if not custom_text.strip():
        return _computed_insight(kw_data)

    key = _get_gemini_key()
    if not key:
        return custom_text.strip()

    try:
        from google import genai
        client = genai.Client(api_key=key)

        prompt = f"""You are an SEO consultant writing a slide insight for a monthly report.

SEO manager's notes:
{custom_text}

Keyword performance data:
- Total tracked keywords: {kw_data['total']}
- Ranked in top 3: {kw_data['top3_count']}
- Ranked on page 1: {kw_data['top10_count']}
- Improved since tracking started: {kw_data['improved_count']}
- Notable top-3 keywords: {', '.join(k['keyword'] for k in kw_data['top3_samples'][:3])}
- Most improved: {', '.join(k['keyword'] for k in kw_data['improved_samples'][:3])}

Write 2-3 professional sentences combining the manager's notes with the keyword data.
Write in English. Return only the insight text, no JSON, no markdown, no bullet points."""

        response = client.models.generate_content(model=GEMINI_MODEL, contents=[prompt])
        return response.text.strip()
    except Exception:
        return custom_text.strip()


# ── Table request builder ─────────────────────────────────────

def _table_requests(prefix: str, slide_id: str, table_rows: list) -> list:
    """
    Build createTable + all cell styling for one slide.
    Truncates keywords to MAX_KW_CHARS to prevent row height expansion.
    Row heights are calculated to guarantee the table fits within FRAME_P10.
    """
    table_id = f"{prefix}_table"
    n_data   = len(table_rows)
    n_rows   = n_data + 1   # +1 header

    # Row height: fit all rows within available frame space, cap at 20pt
    row_h = min(int(20 * 12700), _TABLE_AVAILABLE_H // n_rows)
    table_h = row_h * n_rows

    table_pos = {**S_TABLE, "h": table_h}
    reqs = []

    reqs.append({"createTable": {
        "objectId": table_id,
        "elementProperties": elem_props(slide_id, table_pos),
        "rows":    n_rows,
        "columns": 5,
    }})

    for ci, w in enumerate(TABLE_COL_WIDTHS):
        reqs.append({"updateTableColumnProperties": {
            "objectId": table_id,
            "columnIndices": [ci],
            "tableColumnProperties": {
                "columnWidth": {"magnitude": w, "unit": "EMU"}
            },
            "fields": "columnWidth",
        }})

    for ri in range(n_rows):
        reqs.append({"updateTableRowProperties": {
            "objectId": table_id,
            "rowIndices": [ri],
            "tableRowProperties": {
                "minRowHeight": {"magnitude": row_h, "unit": "EMU"}
            },
            "fields": "minRowHeight",
        }})

    # Header row
    for ci, label in enumerate(TABLE_COL_HEADERS):
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
    for ri, row in enumerate(table_rows):
        actual_ri = ri + 1
        cat   = row.get("category", "other")
        rank  = row.get("rank",   -1)
        start = row.get("start",   0)
        life  = row.get("life",    0)
        vol   = row.get("volume",  0)

        row_bg = (LIGHT_GREEN if cat == "top3"
                  else LIGHT_BLUE if cat == "top10"
                  else PALE_ORANGE if cat == "improved"
                  else WHITE)

        life_str = f"+{life}" if life > 0 else ("—" if life == 0 else str(life))
        life_col = TEXT_GREEN if life > 0 else (TEXT_RED if life < 0 else BLACK)

        kw_text = row.get("keyword", "")
        if len(kw_text) > MAX_KW_CHARS:
            kw_text = kw_text[:MAX_KW_CHARS - 1] + "…"

        cells = [
            (kw_text,                                   BLACK,    False),
            (str(start) if start > 0 else "—",          BLACK,    False),
            (str(rank)  if rank  > 0 else "—",          BLACK,    True),
            (life_str,                                   life_col, True),
            (f"{vol:,}" if vol > 0 else "—",            BLACK,    False),
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


# ── Overview slide builder ────────────────────────────────────

def build_keyword_overview_slide(
    presentation_id:    str,
    csv_bytes:          bytes = b"",
    kw_data:            dict  = None,
    tracking_link:      str   = "",
    image_url_overview: str   = "",
    custom_text_a:      str   = "",
) -> dict:
    """
    Insert 1 Keywords Overview slide at end of 'Website SEO Performance' section.
    Pass either csv_bytes or pre-parsed kw_data (kw_data takes precedence).
    """
    if kw_data is None:
        kw_data = parse_keyword_csv(csv_bytes) if csv_bytes else None

    insight_overview = generate_keyword_insight(kw_data, custom_text_a) if kw_data else (custom_text_a or "")

    service = get_service()
    pres    = service.presentations().get(presentationId=presentation_id).execute()
    slides  = pres["slides"]
    ins_idx = _find_seo_insert_index(slides)

    run_id  = uuid.uuid4().hex[:8]
    ov_id   = f"seo_kw_{run_id}_ov"
    ov_pfx  = f"seo_kw_{run_id}_ov"
    base_font = {"fontFamily": BODY_FONT, "fontSize": 13}
    reqs = []

    reqs.append({"createSlide": {
        "objectId": ov_id,
        "insertionIndex": ins_idx,
        "slideLayoutReference": {"predefinedLayout": "BLANK"},
    }})
    reqs.extend(frame_requests(ov_pfx, ov_id, FRAME_P10))
    reqs.extend(title_requests(ov_pfx, ov_id, "Target Keywords performance", TITLE_P10))

    # Stats block (only when we have CSV data)
    if kw_data:
        total    = kw_data["total"]
        top3     = kw_data["top3_count"]
        top10    = kw_data["top10_count"]
        top20    = kw_data["top20_count"]
        top30    = kw_data["top30_count"]

        reqs.append({"createShape": {
            "objectId": f"{ov_pfx}_stats",
            "shapeType": "TEXT_BOX",
            "elementProperties": elem_props(ov_id, S_OVERVIEW_STATS),
        }})
        reqs.extend(text_segments_requests(f"{ov_pfx}_stats", [
            ("Ranking of ",                                               {**base_font, "bold": False, "color": BLACK}),
            (str(total),                                                  {**base_font, "bold": True,  "color": BLACK}),
            (" target keywords has been up since the project start\n\n", {**base_font, "bold": False, "color": BLACK}),
            (f"{top3}/{total}",                                           {**base_font, "bold": True,  "color": GREEN}),
            (" are ranked in top 3\n",                                   {**base_font, "bold": False, "color": BLACK}),
            (f"{top10}/{total}",                                          {**base_font, "bold": True,  "color": GREEN}),
            (" are ranked in top 1 page\n",                              {**base_font, "bold": False, "color": BLACK}),
            (f"{top20}/{total}",                                          {**base_font, "bold": True,  "color": GREEN}),
            (" are ranked in top 2 pages\n",                             {**base_font, "bold": False, "color": BLACK}),
            (f"{top30}/{total}",                                          {**base_font, "bold": True,  "color": GREEN}),
            (" are ranked in top 3 pages",                               {**base_font, "bold": False, "color": BLACK}),
        ]))

    # Tracking link
    reqs.append({"createShape": {
        "objectId": f"{ov_pfx}_link",
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(ov_id, S_OVERVIEW_LINK),
    }})
    if tracking_link:
        reqs.append({"insertText": {"objectId": f"{ov_pfx}_link", "text": "Live Tracking Link", "insertionIndex": 0}})
        reqs.append({"updateTextStyle": {
            "objectId": f"{ov_pfx}_link",
            "textRange": {"type": "ALL"},
            "style": {
                "fontFamily": BODY_FONT,
                "fontSize": {"magnitude": 14, "unit": "PT"},
                "foregroundColor": {"opaqueColor": {"rgbColor": BLUE}},
                "link": {"url": tracking_link},
                "underline": True,
            },
            "fields": "fontFamily,fontSize,foregroundColor,link,underline",
        }})

    # Overview screenshot
    if image_url_overview:
        reqs.append({"createImage": {
            "objectId": f"{ov_pfx}_image",
            "url": image_url_overview,
            "elementProperties": elem_props(ov_id, S_OVERVIEW_IMAGE),
        }})

    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": reqs},
    ).execute()

    return {"slides_inserted": 1}


# ── Ranking table slides builder ──────────────────────────────

def build_keyword_table_slides(
    presentation_id: str,
    csv_bytes:       bytes,
    kw_data:         dict = None,
    custom_text_b:   str  = "",
) -> dict:
    """
    Insert 1-N Ranking Table slides at end of 'Website SEO Performance' section.
    """
    if kw_data is None:
        kw_data = parse_keyword_csv(csv_bytes)

    insight_table = generate_keyword_insight(kw_data, custom_text_b)
    table_rows    = kw_data["table_rows"]
    pages = [table_rows[i: i + MAX_ROWS_PER_SLIDE]
             for i in range(0, max(len(table_rows), 1), MAX_ROWS_PER_SLIDE)]
    n_pages = len(pages)

    service = get_service()
    pres    = service.presentations().get(presentationId=presentation_id).execute()
    slides  = pres["slides"]
    ins_idx = _find_seo_insert_index(slides)

    run_id  = uuid.uuid4().hex[:8]
    reqs    = []

    for page_num, page_rows in enumerate(pages):
        slide_id = f"seo_kw_{run_id}_tbl{page_num}"
        pfx      = f"seo_kw_{run_id}_t{page_num}"
        pos      = ins_idx + page_num

        title_text = (
            f"Target Keywords performance ({page_num + 1})"
            if n_pages > 1
            else "Target Keywords performance"
        )

        reqs.append({"createSlide": {
            "objectId": slide_id,
            "insertionIndex": pos,
            "slideLayoutReference": {"predefinedLayout": "BLANK"},
        }})
        reqs.extend(frame_requests(pfx, slide_id, FRAME_P10))
        reqs.extend(title_requests(pfx, slide_id, title_text, TITLE_P10))
        reqs.extend(_table_requests(pfx, slide_id, page_rows))

        if page_num == 0:
            # Legend
            reqs.append({"createShape": {
                "objectId": f"{pfx}_legend",
                "shapeType": "TEXT_BOX",
                "elementProperties": elem_props(slide_id, S_LEGEND),
            }})
            reqs.extend(text_segments_requests(f"{pfx}_legend", [
                ("■ Ranked #1–3\n\n",  {"fontFamily": BODY_FONT, "fontSize": 10, "bold": True, "color": TEXT_GREEN}),
                ("■ Ranked #4–10\n\n", {"fontFamily": BODY_FONT, "fontSize": 10, "bold": True, "color": BLUE}),
                ("■ Life improved",    {"fontFamily": BODY_FONT, "fontSize": 10, "bold": True,
                                        "color": {"red": 0.8, "green": 0.4, "blue": 0.0}}),
            ]))

            # Insight
            reqs.append({"createShape": {
                "objectId": f"{pfx}_insight",
                "shapeType": "TEXT_BOX",
                "elementProperties": elem_props(slide_id, S_TABLE_INSIGHT),
            }})
            reqs.extend(text_segments_requests(f"{pfx}_insight", [
                (insight_table, {"fontFamily": BODY_FONT, "fontSize": 10, "bold": False, "color": BLACK}),
            ]))

    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": reqs},
    ).execute()

    return {
        "total_keywords":   kw_data["total"],
        "table_rows_shown": len(table_rows),
        "table_slides":     n_pages,
    }


# ── Combined wrapper (backward compat) ────────────────────────

def build_keyword_slides(
    presentation_id:    str,
    csv_bytes:          bytes,
    tracking_link:      str = "",
    image_url_overview: str = "",
    custom_text_a:      str = "",
    custom_text_b:      str = "",
) -> dict:
    kw_data = parse_keyword_csv(csv_bytes)
    build_keyword_overview_slide(
        presentation_id, csv_bytes=b"", kw_data=kw_data,
        tracking_link=tracking_link,
        image_url_overview=image_url_overview,
        custom_text_a=custom_text_a,
    )
    result = build_keyword_table_slides(
        presentation_id, csv_bytes=b"", kw_data=kw_data,
        custom_text_b=custom_text_b,
    )
    return result
