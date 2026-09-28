from pathlib import Path
import os
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# Project Paths
# ==========================================

BASE_DIR = Path(__file__).parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

CLIENT_TIMESHEET = INPUT_DIR / "Infogain_Test (49).xlsx"
PARENT_TIMESHEET = INPUT_DIR / "June Reco-7.xlsx"

OUTPUT_FILE = OUTPUT_DIR / "updated_client_timesheet.xlsx"

PENDING_REVIEW_FILE = OUTPUT_DIR / "pending_review.json"

MAPPING_FILE = BASE_DIR / "employee_mapping.json"

# ==========================================
# Gemini Configuration
# ==========================================

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

# ==========================================
# Weekend Labels
# ==========================================

SATURDAY_TEXT = "ASPIRE"
SUNDAY_TEXT = "FC"

# ==========================================
# Excel Highlight Colors
# ==========================================

MATCH_COLOR = "90EE90"      # Light Green
MISMATCH_COLOR = "FFC7CE"   # Light Red
MISSING_CLIENT_COLOR = "FFFF00"  # Yellow
CLIENT_ONLY_COLOR = "ADD8E6"  # Light Blue for client-only employees
PENDING_REVIEW_COLOR = "FFB347"  # Orange for weeks needing manual review