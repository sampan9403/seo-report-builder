# ============================================================
# task_detail_builder.py — Task Detail slides (P.5-7 style)
# One slide per task with title, images, insight, optional link
# ============================================================

import json, uuid, os, sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from google import genai
from google.genai import types as genai_types
from config import GEMINI_API_KEY, GEMINI_MODEL
from modules.slide_utils import (
    get_service, elem_props, text_segments_requests,
    find_header_index, all_text_from_slide,
    ORANGE, BLACK, WHITE, BLUE,
    BODY_FONT, TITLE_FONT,
)

# ── Layout constants (measured from template P.5-7) ──────────
# All values in EMU (1pt = 12700 EMU). Slide = 720pt × 405pt.

# Decorative frame (ROUND_RECTANGLE)
DETAIL_FRAME = {
    "x": 702200,  "y": 854825,
    "w": 7783199, "h": 3854400,   # 612.9pt × 303.5pt
}

# Section label (Bebas Neue 35pt orange, top-left area)
DETAIL_SECTION_LABEL = {
    "x": 283250,  "y": 176300,
    "w": 6993900, "h": 572700,    # 550.7pt × 45.1pt
}

# Task title bar (Montserrat 15pt bold, DARK1, white bg — overlays top of frame)
DETAIL_TITLE_BAR = {
    "x": 680400,  "y": 854825,
    "w": 7783199, "h": 395699,    # 612.9pt × 31.2pt
}

# Body area (images live here)
BODY_X   = 702200
BODY_Y   = 1250524   # immediately below title bar
BODY_W   = 7783199
BODY_H   = 2642426   # extends to y=3892950
BODY_PAD = 100000    # left/right padding inside body

# Insight/description bar (Montserrat 12pt, white bg — overlays bottom of frame)
DETAIL_DESC = {
    "x": 702200,  "y": 3892950,
    "w": 7587900, "h": 794100,    # 597.5pt × 62.5pt
}

# Description box when NO images provided (fills full body area)
DETAIL_DESC_NOIMG = {
    "x": BODY_X + BODY_PAD, "y": BODY_Y + 50000,
    "w": BODY_W - 2 * BODY_PAD,  "h": BODY_H - 100000,
}


def _get_api_key() -> str:
    key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY", "")
    if not key:
        try:
            import streamlit as st
            key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            pass
    return key


def _gemini_client():
    return genai.Client(api_key=_get_api_key())


def analyze_task_detail(
    task_name: str,
    description: str,
    image_bytes_list: list,
    doc_url: str,
) -> dict:
    """
    Use Gemini to generate slide content from all available inputs.
    Returns: {slide_title, insight, link_anchor}
    """
    client = _gemini_client()

    context_parts = []
    if description:
        context_parts.append(f"Description provided by user:\n{description}")
    if image_bytes_list:
        context_parts.append(f"{len(image_bytes_list)} screenshot(s) attached — analyse their content.")
    if doc_url:
        context_parts.append(f"Document/reference URL: {doc_url}")

    prompt = f"""You are an SEO consultant writing content for a monthly SEO report slide.

Task name: {task_name}
{chr(10).join(context_parts)}

Generate professional slide content in JSON:
{{
  "slide_title": "Short task title (max 8 words, professional)",
  "insight": "2-3 sentence explanation of what was done and why it matters for SEO. Be specific and professional.",
  "link_anchor": "Concise anchor text for the document link (4-6 words, only if doc URL was provided, else empty string)"
}}

Rules:
- slide_title should be a clean professional version of the task name
- insight should summarise the work done and its SEO impact
- Write in English
- Return ONLY valid JSON, no markdown"""

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
            "slide_title": task_name[:60],
            "insight": description or "Task completed as part of the monthly SEO engagement.",
            "link_anchor": "View Document" if doc_url else "",
        }


def _image_positions(n: int) -> list:
    """Return list of {x,y,w,h} dicts for n images laid out in body area."""
    if n <= 0:
        return []
    gap = 100000
    x0  = BODY_X + BODY_PAD
    y0  = BODY_Y + 50000
    h   = BODY_H - 100000
    avail_w = BODY_W - 2 * BODY_PAD - gap * max(n - 1, 0)
    each_w  = avail_w // n
    return [
        {"x": x0 + i * (each_w + gap), "y": y0, "w": each_w, "h": h}
        for i in range(n)
    ]


def _frame_requests(prefix: str, slide_id: str) -> list:
    """Create ROUND_RECTANGLE frame for detail slide."""
    fid  = f"{prefix}_frame"
    reqs = []
    reqs.append({"createShape": {
        "objectId": fid,
        "shapeType": "ROUND_RECTANGLE",
        "elementProperties": elem_props(slide_id, DETAIL_FRAME),
    }})
    reqs.append({"updateShapeProperties": {
        "objectId": fid,
        "shapeProperties": {
            "shapeBackgroundFill": {
                "solidFill": {"color": {"themeColor": "LIGHT2"}, "alpha": 1}
            },
            "outline": {
                "outlineFill": {
                    "solidFill": {"color": {"themeColor": "ACCENT2"}, "alpha": 1}
                },
                "weight": {"magnitude": 19050, "unit": "EMU"},
                "dashStyle": "SOLID",
            },
        },
        "fields": "shapeBackgroundFill,outline",
    }})
    return reqs


def _section_label_requests(prefix: str, slide_id: str, section_text: str) -> list:
    """Create section label (orange Bebas Neue 35pt) at top-left."""
    lid  = f"{prefix}_label"
    reqs = [{"createShape": {
        "objectId": lid,
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(slide_id, DETAIL_SECTION_LABEL),
    }}]
    reqs.append({"insertText": {"objectId": lid, "text": section_text, "insertionIndex": 0}})
    reqs.append({"updateTextStyle": {
        "objectId": lid,
        "textRange": {"type": "ALL"},
        "style": {
            "fontFamily": TITLE_FONT,
            "fontSize": {"magnitude": 35, "unit": "PT"},
            "foregroundColor": {"opaqueColor": {"rgbColor": ORANGE}},
            "bold": False,
        },
        "fields": "fontFamily,fontSize,foregroundColor,bold",
    }})
    return reqs


def _title_bar_requests(prefix: str, slide_id: str, title_text: str) -> list:
    """Create white title bar with Montserrat 15pt bold dark text."""
    tid  = f"{prefix}_titlebar"
    reqs = [{"createShape": {
        "objectId": tid,
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(slide_id, DETAIL_TITLE_BAR),
    }}]
    reqs.append({"updateShapeProperties": {
        "objectId": tid,
        "shapeProperties": {
            "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": WHITE}, "alpha": 1}}
        },
        "fields": "shapeBackgroundFill",
    }})
    reqs.append({"insertText": {"objectId": tid, "text": title_text, "insertionIndex": 0}})
    reqs.append({"updateTextStyle": {
        "objectId": tid,
        "textRange": {"type": "ALL"},
        "style": {
            "fontFamily": BODY_FONT,
            "fontSize": {"magnitude": 15, "unit": "PT"},
            "foregroundColor": {"opaqueColor": {"rgbColor": BLACK}},
            "bold": True,
        },
        "fields": "fontFamily,fontSize,foregroundColor,bold",
    }})
    return reqs


def _description_requests(
    prefix: str,
    slide_id: str,
    insight: str,
    doc_url: str = "",
    link_anchor: str = "",
    has_images: bool = True,
) -> list:
    """Create insight text box (with optional clickable link appended)."""
    did  = f"{prefix}_desc"
    pos  = DETAIL_DESC if has_images else DETAIL_DESC_NOIMG
    reqs = [{"createShape": {
        "objectId": did,
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(slide_id, pos),
    }}]
    reqs.append({"updateShapeProperties": {
        "objectId": did,
        "shapeProperties": {
            "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": WHITE}, "alpha": 1}}
        },
        "fields": "shapeBackgroundFill",
    }})

    font_sz = 12 if has_images else 13
    base = {"fontFamily": BODY_FONT, "fontSize": font_sz, "bold": False, "color": BLACK}

    if doc_url and link_anchor:
        link_text = f"  |  {link_anchor}"
        segments  = [(insight, base), (link_text, {**base, "color": BLUE})]
    else:
        segments = [(insight, base)]

    reqs.extend(text_segments_requests(did, segments))

    # Apply hyperlink on link text
    if doc_url and link_anchor:
        link_start = len(insight)
        link_end   = link_start + len(f"  |  {link_anchor}")
        reqs.append({"updateTextStyle": {
            "objectId": did,
            "textRange": {"type": "FIXED_RANGE", "startIndex": link_start, "endIndex": link_end},
            "style": {"link": {"url": doc_url}},
            "fields": "link",
        }})

    return reqs


def _find_detail_insert_index(slides: list) -> int:
    """
    Insert detail slides just before the next major section header
    (Website SEO / Keywords / Coming Tasks) that comes after the
    Tasks Completed header. Falls back to appending at end.
    """
    TASKS_EXCLUDE = ["website seo", "coming tasks"]
    task_header = find_header_index(slides, ["tasks completed"], exclude_keywords=TASKS_EXCLUDE)
    if task_header is None:
        task_header = find_header_index(slides, ["tasks"], exclude_keywords=TASKS_EXCLUDE)
    start = (task_header + 1) if task_header is not None else 0

    for i in range(start, len(slides)):
        text = all_text_from_slide(slides[i]).lower()
        if any(kw in text for kw in ["website seo", "keywords performance", "coming tasks", "coming soon"]):
            return i

    return len(slides)


def build_task_detail_slides(
    presentation_id: str,
    tasks: list,
) -> int:
    """
    Build one detail slide per task and insert them into the presentation.

    Each task dict:
    {
        "name":         str,          # task name
        "insight":      str,          # Gemini-generated insight
        "slide_title":  str,          # Gemini-generated title
        "image_urls":   [str],        # public Drive image URLs (for Slides API)
        "doc_url":      str,          # optional document link
        "link_anchor":  str,          # optional anchor text
    }

    Returns number of slides created.
    """
    service = get_service()
    pres    = service.presentations().get(presentationId=presentation_id).execute()
    slides  = pres["slides"]

    insert_at = _find_detail_insert_index(slides)
    run_id    = uuid.uuid4().hex[:8]
    all_reqs  = []

    for i, task in enumerate(tasks):
        slide_id = f"seo_detail_{run_id}_{i}"
        prefix   = f"seo_det_{run_id}_{i}"
        pos      = insert_at + i

        # 1. Create blank slide
        all_reqs.append({"createSlide": {
            "objectId": slide_id,
            "insertionIndex": pos,
            "slideLayoutReference": {"predefinedLayout": "BLANK"},
        }})

        # 2. Decorative frame
        all_reqs.extend(_frame_requests(prefix, slide_id))

        # 3. Section label (orange Bebas Neue)
        all_reqs.extend(_section_label_requests(prefix, slide_id, "Completed Tasks"))

        # 4. Task title bar (white, dark bold)
        title_text = task.get("slide_title") or task.get("name", "Task")
        all_reqs.extend(_title_bar_requests(prefix, slide_id, title_text))

        # 5. Images
        image_urls  = task.get("image_urls", [])
        n_imgs      = min(len(image_urls), 3)  # max 3 per slide
        img_positions = _image_positions(n_imgs)

        for j, (img_url, img_pos) in enumerate(zip(image_urls[:3], img_positions)):
            img_id = f"{prefix}_img{j}"
            all_reqs.append({"createImage": {
                "objectId": img_id,
                "url": img_url,
                "elementProperties": elem_props(slide_id, img_pos),
            }})

        # 6. Insight + optional link
        all_reqs.extend(_description_requests(
            prefix        = prefix,
            slide_id      = slide_id,
            insight       = task.get("insight", ""),
            doc_url       = task.get("doc_url", ""),
            link_anchor   = task.get("link_anchor", ""),
            has_images    = n_imgs > 0,
        ))

    service.presentations().batchUpdate(
        presentationId=presentation_id,
        body={"requests": all_reqs},
    ).execute()

    return len(tasks)
