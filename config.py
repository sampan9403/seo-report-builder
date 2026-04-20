# ============================================================
# SEO Report Builder — Configuration
# ============================================================

# Google API credentials
CREDENTIALS_FILE = "Google-Slides-Keys.json"

# Bannershop master template (source to copy from)
BANNERSHOP_TEMPLATE_ID = "1fAxAZUqg3LDF_CvRHnvjD7Zr0TbzVx4KRBwF4QVBT4g"

# Target Google Drive folder (new reports will be saved here)
TARGET_DRIVE_FOLDER_ID = "1HAAgVUuSaQiaKvZgjtacdQ4OyvcArl6x"

# Google API scopes needed
SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
]

# Bannershop brand style
STYLE = {
    "accent_color": {"red": 0.9372549, "green": 0.36862746, "blue": 0.2627451},  # #EF5E43
    "dark_color":   {"red": 0.0,       "green": 0.0,        "blue": 0.0},
    "muted_color":  {"red": 0.42,      "green": 0.44,       "blue": 0.48},
    "blue_link":    {"red": 0.07,      "green": 0.33,       "blue": 0.80},
    "title_font":   "Bebas Neue",
    "body_font":    "Montserrat",
    "title_size":   35,
    "body_size":    13,
}

# Slide indices in Bannershop template (0-based)
SLIDE_10_INDEX = 9   # Target Keywords Performance (summary)
SLIDE_11_INDEX = 10  # Target Keywords Performance (detail table)

# ── imgbb (free image hosting for inserting images into Slides) ──
# Service Accounts have no Drive storage quota, so images are hosted here instead.
# Free API key at: https://api.imgbb.com  (sign up at imgbb.com, no credit card)
IMGBB_API_KEY = ""  # Paste your key here

# ── Gemini API (Google AI Studio) ───────────────────────────
# Used for: analysing Keyword.com screenshots (vision task)
# Get your key at: https://aistudio.google.com/app/apikey
GEMINI_API_KEY = ""  # Set via Streamlit Secrets (GEMINI_API_KEY) or local config

# Model selection — change this line to upgrade/downgrade
# Recommended for screenshot analysis (cheap, fast, free tier available):
GEMINI_MODEL = "gemini-2.5-flash"
#
# Other options (uncomment to switch):
# GEMINI_MODEL = "gemini-2.5-flash-lite"  # Cheapest, free tier
# GEMINI_MODEL = "gemini-2.5-flash"       # Better quality, low cost (current)
# GEMINI_MODEL = "gemini-2.5-pro"         # Highest quality, highest cost
