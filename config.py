from datetime import date

# ─── File Paths ───────────────────────────────────────────────────────────────
EXCEL_FILE_PATH = "HARTS_Events.xlsx"
LOG_FILE = "event_scout.log"
TODAY = date.today()
REMINDER_DAYS_BEFORE = 15

# ─── Email ────────────────────────────────────────────────────────────────────
RECIPIENT_EMAIL = "rethik.viswanathan@globalharts.com"

# ─── M365 Credentials (TODO: fill in when enabling email reminders) ───────────
# TODO: CLIENT_ID = ""
# TODO: CLIENT_SECRET = ""
# TODO: TENANT_ID = ""

# ─── Search Queries (50 total) ────────────────────────────────────────────────
SEARCH_QUERIES = [
    # GCC & Shared Services
    "GCC Global Capability Center conference 2025 India",
    "Global Capability Center summit leadership 2025",
    "NASSCOM GCC summit 2025",
    "Shared Services conference summit 2025",
    "SSON Shared Services excellence conference 2025",
    "Global Business Services GBS summit 2025",
    "GCC operating model strategy forum 2025",
    "Global In-house Center conference India 2025",
    "GCC scaling leadership forum Asia 2025",
    "Shared Services transformation conference 2025",
    "GBS delivery model conference 2025",
    "GCC India Bangalore Hyderabad summit 2025",
    # OD & Implementation
    "Organization Design conference 2025",
    "Operating model transformation summit 2025",
    "Enterprise transformation conference 2025",
    "Organizational effectiveness leadership forum 2025",
    "Change management conference senior leadership 2025",
    "OD organizational development conference 2025",
    "Enterprise restructuring transformation summit 2025",
    "Business transformation conference CXO 2025",
    # Centre of Excellence
    "Centre of Excellence setup strategy conference 2025",
    "Enterprise capability building COE summit 2025",
    "COE strategy leadership forum 2025",
    "Center of Excellence governance conference 2025",
    # Post M&A Integration
    "Post merger integration conference 2025",
    "M&A integration summit leadership 2025",
    "Corporate restructuring conference 2025",
    "M&A value creation leadership forum 2025",
    "Post acquisition integration strategy conference 2025",
    "Enterprise integration management summit 2025",
    # Recruitment as a Service
    "RPO recruitment process outsourcing summit 2025",
    "Talent acquisition leadership forum 2025",
    "Recruitment as a Service conference 2025",
    "HR talent strategy conference 2025",
    "Workforce planning leadership summit 2025",
    # Executive Coaching & Board Advisory
    "Executive coaching conference ICF 2025",
    "Board advisory leadership summit 2025",
    "C-suite CXO strategy leadership forum 2025",
    "Senior executive leadership development summit 2025",
    "Board governance directors conference 2025",
    "Executive leadership coaching forum 2025",
    # Industry Analyst / Organizer Events
    "Everest Group analyst summit outsourcing 2025",
    "ISG outsourcing summit 2025",
    "Deloitte consulting leadership summit 2025",
    "EY strategy transformation conference 2025",
    "CII industry leadership summit India 2025",
    "SHRM HR leadership conference 2025",
    "People Matters HR summit 2025",
    "ETHRWorld HR transformation summit 2025",
    "CHRO chief HR officer leadership forum 2025",
]
