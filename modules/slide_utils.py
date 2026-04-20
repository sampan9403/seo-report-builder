# ============================================================
# slide_utils.py — Common primitives for all slide builders
# Layout constants extracted from Bannershop template reference
# ============================================================

from google.oauth2 import service_account
from googleapiclient.discovery import build
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import CREDENTIALS_FILE, SCOPES


def _get_credentials():
    """
    Load Google credentials from Streamlit Secrets (cloud) or
    local JSON file (development). Falls back gracefully.
    """
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            return service_account.Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=SCOPES
            )
    except Exception:
        pass
    return service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES
    )

# ── Brand colours ────────────────────────────────────────────
ORANGE = {"red": 0.9372549, "green": 0.36862746, "blue": 0.2627451}
BLACK  = {"red": 0.0, "green": 0.0, "blue": 0.0}
WHITE  = {"red": 1.0, "green": 1.0, "blue": 1.0}
BLUE   = {"red": 0.06666667, "green": 0.33333334, "blue": 0.8}
GREEN  = {"red": 0.42, "green": 0.66, "blue": 0.31}
RED    = {"red": 1.0, "green": 0.0, "blue": 0.0}
DARK_BLUE = {"red": 0.0, "green": 0.0, "blue": 1.0}
TITLE_FONT = "Bebas Neue"
BODY_FONT  = "Montserrat"

# ── Layout: Decorative frame (ROUND_RECTANGLE) ───────────────
# Used on P.4 style (Tasks) slides
FRAME_P4 = {
    "x": 680400,            "y": 880224,
    "w": int(612.9 * 12700), "h": int(293.3 * 12700),
}
# Used on P.10/P.11 style (Keywords) slides
FRAME_P10 = {
    "x": 512250,             "y": 1163100,
    "w": int(639.33 * 12700), "h": int(258.83 * 12700),
}

# ── Layout: Decorative dots (3 ELLIPSEs, upper right) ────────
DOT_SIZE = 108300  # EMU  ≈ 8.53pt
DOTS = [
    {"x": 8059250, "y": 431700},
    {"x": 8277475, "y": 431700},
    {"x": 8495700, "y": 431700},
]

# ── Layout: Title text box ───────────────────────────────────
TITLE_P4 = {
    "x": 612975, "y": 218275,
    "w": int(550.7 * 12700), "h": int(45.1 * 12700),
}
TITLE_P10 = {
    "x": 512250, "y": 365300,
    "w": int(492.24 * 12700), "h": int(45.09 * 12700),
}


def get_service():
    return build("slides", "v1", credentials=_get_credentials())


def elem_props(slide_id: str, pos: dict) -> dict:
    """Build elementProperties for createShape / createImage."""
    return {
        "pageObjectId": slide_id,
        "size": {
            "width":  {"magnitude": pos["w"], "unit": "EMU"},
            "height": {"magnitude": pos["h"], "unit": "EMU"},
        },
        "transform": {
            "scaleX": 1, "scaleY": 1,
            "translateX": pos["x"], "translateY": pos["y"],
            "unit": "EMU",
        },
    }


def all_text_from_slide(slide: dict) -> str:
    """Collect all visible text from a slide (recursively into groups)."""
    parts = []
    def collect(elems):
        for e in elems:
            for te in e.get("shape", {}).get("text", {}).get("textElements", []):
                t = te.get("textRun", {}).get("content", "").strip()
                if t:
                    parts.append(t)
            collect(e.get("elementGroup", {}).get("children", []))
    collect(slide.get("pageElements", []))
    return " ".join(parts)


def find_header_index(slides: list, keywords: list,
                      exclude_keywords: list | None = None) -> int | None:
    """
    Return the 0-based index of the first slide whose text:
      - contains ALL of `keywords` (case-insensitive)
      - does NOT contain any of `exclude_keywords` (case-insensitive)
    The exclusion prevents matching the TOC slide which lists all section names.
    """
    for i, slide in enumerate(slides):
        text = all_text_from_slide(slide).lower()
        if not all(kw.lower() in text for kw in keywords):
            continue
        if exclude_keywords and any(ex.lower() in text for ex in exclude_keywords):
            continue
        return i
    return None


def delete_tool_slides_requests(slides: list, id_set: set) -> list:
    """
    Return deleteObject requests for slides whose objectId is in id_set.
    Returns empty list if none found.
    """
    reqs = []
    for slide in slides:
        if slide["objectId"] in id_set:
            reqs.append({"deleteObject": {"objectId": slide["objectId"]}})
    return reqs


def frame_requests(obj_prefix: str, slide_id: str, frame_pos: dict) -> list:
    """
    Create ROUND_RECTANGLE decorative frame + 3 dot ellipses.
    Returns list of batchUpdate request dicts.
    """
    frame_id = f"{obj_prefix}_frame"
    reqs = []

    # ── Rounded rectangle ────────────────────────────────────
    reqs.append({"createShape": {
        "objectId": frame_id,
        "shapeType": "ROUND_RECTANGLE",
        "elementProperties": elem_props(slide_id, frame_pos),
    }})
    reqs.append({"updateShapeProperties": {
        "objectId": frame_id,
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

    # ── 3 decorative dots ─────────────────────────────────────
    for di, dot in enumerate(DOTS):
        dot_id = f"{obj_prefix}_dot{di}"
        dot_pos = {**dot, "w": DOT_SIZE, "h": DOT_SIZE}
        reqs.append({"createShape": {
            "objectId": dot_id,
            "shapeType": "ELLIPSE",
            "elementProperties": elem_props(slide_id, dot_pos),
        }})
        reqs.append({"updateShapeProperties": {
            "objectId": dot_id,
            "shapeProperties": {
                "shapeBackgroundFill": {
                    "solidFill": {"color": {"themeColor": "ACCENT1"}, "alpha": 1}
                },
                "outline": {
                    "outlineFill": {"solidFill": {"color": {"rgbColor": {}}}},
                    "propertyState": "NOT_RENDERED",
                },
            },
            "fields": "shapeBackgroundFill,outline",
        }})

    return reqs


def title_requests(obj_prefix: str, slide_id: str,
                   title_text: str, title_pos: dict) -> list:
    """Create a styled title text box (Bebas Neue 35pt orange)."""
    tid = f"{obj_prefix}_title"
    reqs = [{"createShape": {
        "objectId": tid,
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(slide_id, title_pos),
    }}]
    reqs.append({"insertText": {"objectId": tid, "text": title_text, "insertionIndex": 0}})
    reqs.append({"updateTextStyle": {
        "objectId": tid,
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


def text_segments_requests(obj_id: str, segments: list) -> list:
    """
    Build insertText + updateTextStyle requests for mixed-style text
    in a NEW (empty) shape.
    segments: list of (text, style_dict)
    style_dict keys: bold, color, fontFamily, fontSize
    """
    full_text = "".join(t for t, _ in segments)
    reqs = [{"insertText": {"objectId": obj_id, "text": full_text, "insertionIndex": 0}}]

    cursor = 0
    for text, style in segments:
        start, end = cursor, cursor + len(text)
        cursor = end
        if not style or start == end:
            continue
        sd, fields = {}, []
        if "bold" in style:
            sd["bold"] = style["bold"]; fields.append("bold")
        if "color" in style and style["color"] is not None:
            sd["foregroundColor"] = {"opaqueColor": {"rgbColor": style["color"]}}
            fields.append("foregroundColor")
        if "fontFamily" in style:
            sd["fontFamily"] = style["fontFamily"]; fields.append("fontFamily")
        if "fontSize" in style:
            sd["fontSize"] = {"magnitude": style["fontSize"], "unit": "PT"}
            fields.append("fontSize")
        if sd:
            reqs.append({"updateTextStyle": {
                "objectId": obj_id,
                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end},
                "style": sd,
                "fields": ",".join(fields),
            }})
    return reqs
