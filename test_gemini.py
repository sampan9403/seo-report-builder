"""Quick test — verifies Gemini API key and model are working."""
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from config import GEMINI_API_KEY, GEMINI_MODEL
from google import genai

api_key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY", "")
if not api_key:
    print("ERROR: GEMINI_API_KEY is empty in config.py")
    sys.exit(1)

client = genai.Client(api_key=api_key)
response = client.models.generate_content(
    model=GEMINI_MODEL,
    contents="Reply with exactly: OK"
)
print(f"Model   : {GEMINI_MODEL}")
print(f"Response: {response.text.strip()}")
print("Gemini API connection successful.")
