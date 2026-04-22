"""
Ontario Intertie Congestion Scanner — Configuration
"""
from datetime import timezone, timedelta

# ── Time zones ──────────────────────────────────────────────────────
# IESO operates on Eastern Standard Time year-round (no DST observed in
# market hour conventions). HE 1 = 00:00–01:00 EST regardless of season.
EST = timezone(timedelta(hours=-5))

# ── IESO Report URLs ────────────────────────────────────────────────
IESO_PREDISP_LIMITS_URL = (
    "https://reports-public.ieso.ca/public/PredispIntertieSchedLimits/"
    "PUB_PredispIntertieSchedLimits.xml"
)
IESO_ADEQUACY3_DIR_URL = "https://reports-public.ieso.ca/public/Adequacy3/"
IESO_ADEQUACY3_BASE_URL = "https://reports-public.ieso.ca/public/Adequacy3"

# ── NYISO ATC/TTC ───────────────────────────────────────────────────
NYISO_ATC_TTC_URL = "https://mis.nyiso.com/public/htm/atc_ttc/{date}atc_ttc.htm"
# Directions are named from NYISO's perspective:
#   IMO-NYISO  = Ontario → NY (ON EXPORT)
#   NYISO-IMO  = NY → Ontario (ON IMPORT)
NYISO_IMO_NY_INTERFACE = "IMO-NYISO"
NYISO_NY_IMO_INTERFACE = "NYISO-IMO"
NYISO_SECURITY_BUFFER_MW = 300  # TTC overstatement — subtract for real capability

# ── IESO Intertie zone codes ────────────────────────────────────────
# PredispIntertieSchedLimits zone naming convention (per IESO XSL stylesheet):
#   suffix N = "to <zone>"   → flow OUT of Ontario → EXPORT from Ontario
#   suffix X = "from <zone>" → flow INTO Ontario → IMPORT into Ontario
# Values are signed (N negative, X positive); take abs() to get capacity.
IESO_ZONES = {
    "NY": {
        "import": "NYSIX", "export": "NYSIN",
        "name": "NYSI (New York — Roseton)",
        "adeq": "New York",
    },
    "MI": {
        "import": "MISIX", "export": "MISIN",
        "name": "MISI (Michigan — Ludington)",
        "adeq": "Michigan",
    },
    "MN": {
        "import": "MNSIX", "export": "MNSIN",
        "name": "MNSI (Minnesota — Intl Falls)",
        "adeq": "Minnesota",
    },
    "MB": {
        "import": "MBSIX", "export": "MBSIN",
        "name": "MBSI (Manitoba)",
        "adeq": "Manitoba",
    },
    "PQ": {
        "import": None, "export": None,  # Quebec is summed from sub-zones below
        "name": "PQSI (Quebec — summed)",
        "adeq": "Quebec",
    },
}

# Display order (primary strategy corridor first)
DISPLAY_ORDER = ["NY", "MI", "MN", "MB", "PQ"]

# Joint MISI+NYSI constraint (usually non-binding at ±9999)
IESO_JOINT_LIMIT = {
    "import": "MISI+NYSIN", "export": "MISI+NYSIX",
    "name": "MISI+NYSI joint",
}

# Quebec sub-zones that roll up into the Quebec total
QUEBEC_SUBZONES = ["AT", "BE", "DA", "DZ", "HA", "HZ", "PC", "QC", "SK", "XY"]

# ── Thresholds for status classification ────────────────────────────
CRITICAL_MW = 100      # <100 MW remaining → 🔴
CRITICAL_PCT = 90      # >90% utilized   → 🔴
WARNING_MW = 300       # <300 MW remaining → ⚠️
WARNING_PCT = 70       # >70% utilized   → ⚠️

# ── Output ──────────────────────────────────────────────────────────
FORECAST_HOURS = 8     # how many forward HEs to display (includes current)
PAST_HOURS = 2         # how many past HEs to display
