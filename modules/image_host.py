# ============================================================
# image_host.py
# Uploads images to imgbb (free) to get a public URL
# needed by the Google Slides API to insert images.
# ============================================================

import requests
import base64
import sys, os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import IMGBB_API_KEY


def upload_image(image_bytes: bytes) -> str:
    """
    Upload image bytes to imgbb and return a public URL.
    The URL is used by Google Slides API to insert the image.
    """
    api_key = IMGBB_API_KEY or os.environ.get("IMGBB_API_KEY", "")
    if not api_key:
        raise ValueError(
            "imgbb API key not set.\n"
            "1. Sign up free at https://imgbb.com\n"
            "2. Get key at https://api.imgbb.com\n"
            "3. Add to config.py: IMGBB_API_KEY = 'your-key'"
        )

    b64 = base64.b64encode(image_bytes).decode("utf-8")
    response = requests.post(
        "https://api.imgbb.com/1/upload",
        data={"key": api_key, "image": b64},
        timeout=30
    )
    response.raise_for_status()
    result = response.json()

    if not result.get("success"):
        raise RuntimeError(f"imgbb upload failed: {result}")

    return result["data"]["url"]
