# ============================================================
# slides_builder.py — Creates Slides 10 & 11 elements from scratch
# Layout hardcoded from Bannershop template reference inspection
# ============================================================

from google.oauth2 import service_account
from googleapiclient.discovery import build
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import CREDENTIALS_FILE, SCOPES

# ── Brand colors ──────────────────────────────────────────────
GREEN = {"red": 0.42, "green": 0.66, "blue": 0.31}
BLACK = {"red": 0.0,  "green": 0.0,  "blue": 0.0}
RED   = {"red": 1.0,  "green": 0.0,  "blue": 0.0}
BLUE  = {"red": 0.0,  "green": 0.0,  "blue": 1.0}
BODY_FONT = "Montserrat"
BODY_SIZE = 13

# ── Layout from Bannershop template (positions in EMU) ────────
# Slide 10
S10_STATS = {
    "x": 4400896, "y": 1788675,
    "w": int(329.08 * 12700), "h": int(187.87 * 12700),
}
S10_IMAGE = {
    "x": 632475, "y": 2123950,
    "w": int(290.73 * 12700), "h": int(134.37 * 12700),
}
# Slide 11
S11_INSIGHT = {
    "x": 5914650, "y": 2477675,
    "w": int(213.94 * 12700), "h": int(157.13 * 12700),
}
S11_IMAGE = {
    "x": 630425, "y": 1594172,
    "w": int(420.97 * 12700), "h": int(213.94 * 12700),
}

# Deterministic objectIds — allows idempotent re-runs (delete + recreate)
ID_S10_STATS   = "seo_tool_s10_stats"
ID_S10_IMAGE   = "seo_tool_s10_image"
ID_S11_INSIGHT = "seo_tool_s11_insight"
ID_S11_IMAGE   = "seo_tool_s11_image"
ALL_TOOL_IDS   = {ID_S10_STATS, ID_S10_IMAGE, ID_S11_INSIGHT, ID_S11_IMAGE}


def get_service():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES
    )
    return build("slides", "v1", credentials=creds)


def elem_props(slide_id: str, pos: dict) -> dict:
    """Build elementProperties dict for createShape / createImage."""
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


def text_requests(obj_id: str, segments: list) -> list:
    """
    Build batchUpdate requests for mixed-style text in a NEW empty shape.
    segments: list of (text, style_dict)
    style_dict keys: bold, color (rgb dict), fontFamily, fontSize
    """
    full_text = "".join(t for t, _ in segments)
    reqs = [{"insertText": {"objectId": obj_id, "text": full_text, "insertionIndex": 0}}]

    cursor = 0
    for text, style in segments:
        start, end = cursor, cursor + len(text)
        cursor = end
        if not style or start == end:
            continue

        style_dict, fields = {}, []
        if "bold" in style:
            style_dict["bold"] = style["bold"]
            fields.append("bold")
        if "color" in style and style["color"] is not None:
            style_dict["foregroundColor"] = {"opaqueColor": {"rgbColor": style["color"]}}
            fields.append("foregroundColor")
        if "fontFamily" in style:
            style_dict["fontFamily"] = style["fontFamily"]
            fields.append("fontFamily")
        if "fontSize" in style:
            style_dict["fontSize"] = {"magnitude": style["fontSize"], "unit": "PT"}
            fields.append("fontSize")

        if style_dict:
            reqs.append({"updateTextStyle": {
                "objectId": obj_id,
                "textRange": {"type": "FIXED_RANGE", "startIndex": start, "endIndex": end},
                "style": style_dict,
                "fields": ",".join(fields),
            }})

    return reqs


def build_keyword_slides(presentation_id: str, data: dict,
                          image_url_slide10: str, image_url_slide11: str):
    service = get_service()
    pres = service.presentations().get(presentationId=presentation_id).execute()
    slides = pres["slides"]
    slide10_id = slides[9]["objectId"]
    slide11_id = slides[10]["objectId"]

    requests = []

    # ── Step 1: Delete any previously created tool elements ──────
    existing_ids = set()
    for slide in [slides[9], slides[10]]:
        for elem in slide.get("pageElements", []):
            existing_ids.add(elem.get("objectId", ""))
            for child in elem.get("elementGroup", {}).get("children", []):
                existing_ids.add(child.get("objectId", ""))

    for tool_id in ALL_TOOL_IDS:
        if tool_id in existing_ids:
            requests.append({"deleteObject": {"objectId": tool_id}})

    # ── Step 2: Slide 10 — stats text box ────────────────────────
    total  = data.get("total_keywords", 0)
    top3   = data.get("top3_count", 0)
    top10  = data.get("top10_count", 0)
    top20  = data.get("top20_count", top10)
    top30  = data.get("top30_count", top10)
    base   = {"fontFamily": BODY_FONT, "fontSize": BODY_SIZE}

    requests.append({"createShape": {
        "objectId": ID_S10_STATS,
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(slide10_id, S10_STATS),
    }})
    requests.extend(text_requests(ID_S10_STATS, [
        ("Ranking of ",                                             {**base, "bold": False, "color": BLACK}),
        (str(total),                                               {**base, "bold": True,  "color": BLACK}),
        (" target keywords has been up since the project start\n\n",{**base, "bold": False, "color": BLACK}),
        (f"{top3}/{total}",                                        {**base, "bold": True,  "color": GREEN}),
        (" are ranked in top 3\n",                                 {**base, "bold": False, "color": BLACK}),
        (f"{top10}/{total}",                                       {**base, "bold": True,  "color": GREEN}),
        (" are ranked in top 1 page\n",                            {**base, "bold": False, "color": BLACK}),
        (f"{top20}/{total}",                                       {**base, "bold": True,  "color": GREEN}),
        (" are ranked in top 2 pages\n",                           {**base, "bold": False, "color": BLACK}),
        (f"{top30}/{total}",                                       {**base, "bold": True,  "color": GREEN}),
        (" are ranked in top 3 pages",                             {**base, "bold": False, "color": BLACK}),
    ]))

    # ── Step 3: Slide 10 — screenshot image ──────────────────────
    requests.append({"createImage": {
        "objectId": ID_S10_IMAGE,
        "url": image_url_slide10,
        "elementProperties": elem_props(slide10_id, S10_IMAGE),
    }})

    # ── Step 4: Slide 11 — insight text box ──────────────────────
    top1       = data.get("top1_count", 0)
    improved   = data.get("improved_count", 0)
    maintained = data.get("maintained_count", 0)

    requests.append({"createShape": {
        "objectId": ID_S11_INSIGHT,
        "shapeType": "TEXT_BOX",
        "elementProperties": elem_props(slide11_id, S11_INSIGHT),
    }})
    requests.extend(text_requests(ID_S11_INSIGHT, [
        ("Since the project started, ",                 {**base, "bold": False, "color": BLACK}),
        (str(top1),                                     {**base, "bold": True,  "color": BLACK}),
        (" keywords reached ",                          {**base, "bold": False, "color": BLACK}),
        ("position 1",                                  {**base, "bold": True,  "color": BLACK}),
        (", with ",                                     {**base, "bold": False, "color": BLACK}),
        (f"{maintained} maintaining their top ranking", {**base, "bold": True,  "color": BLUE}),
        (". ",                                          {**base, "bold": False, "color": BLACK}),
        (f"{improved} keywords improved their rankings",{**base, "bold": True,  "color": RED}),
        (".",                                           {**base, "bold": False, "color": BLACK}),
    ]))

    # ── Step 5: Slide 11 — screenshot image ──────────────────────
    requests.append({"createImage": {
        "objectId": ID_S11_IMAGE,
        "url": image_url_slide11,
        "elementProperties": elem_props(slide11_id, S11_IMAGE),
    }})

    # ── Execute all at once ───────────────────────────────────────
    if requests:
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests}
        ).execute()
        print(f"[OK] Created {len(requests)} requests on {presentation_id}")
