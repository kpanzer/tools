"""
Fetchers for IESO (PredispIntertieSchedLimits + Adequacy3) and NYISO ATC/TTC.
"""
from __future__ import annotations

import re
import urllib.request

import config

USER_AGENT = "Intertie-Congestion-Scanner/1.0"


def _fetch(url: str, timeout: int = 15) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


# ── IESO PredispIntertieSchedLimits ─────────────────────────────────
IESO_PREDISP_LIMITS_DIR = (
    "https://reports-public.ieso.ca/public/PredispIntertieSchedLimits/"
)
IESO_PREDISP_LIMITS_BASE = (
    "https://reports-public.ieso.ca/public/PredispIntertieSchedLimits"
)


def fetch_predisp_intertie_limits() -> bytes:
    """Fetch the always-current Pre-Dispatch Intertie Scheduling Limits XML.
    Returns the most recent "current" file regardless of delivery date."""
    return _fetch(config.IESO_PREDISP_LIMITS_URL)


def fetch_predisp_limits_for_date(date_str: str) -> tuple[bytes, str] | tuple[None, None]:
    """
    Fetch the latest (highest-version) PredispIntertieSchedLimits file for
    YYYYMMDD. Falls back to the un-versioned dated file. Returns
    (xml_bytes, filename_used). If the date isn't published yet (common for
    "tomorrow" earlier in the day), returns (None, None).
    """
    try:
        dir_html = _fetch(IESO_PREDISP_LIMITS_DIR).decode("utf-8", errors="replace")
    except Exception:
        return (None, None)

    version_pattern = rf"PUB_PredispIntertieSchedLimits_{date_str}_v(\d+)\.xml"
    versions = [int(v) for v in re.findall(version_pattern, dir_html)]
    if versions:
        fname = f"PUB_PredispIntertieSchedLimits_{date_str}_v{max(versions)}.xml"
    else:
        fname = f"PUB_PredispIntertieSchedLimits_{date_str}.xml"

    url = f"{IESO_PREDISP_LIMITS_BASE}/{fname}"
    try:
        return _fetch(url), fname
    except Exception:
        return (None, None)


# ── IESO Adequacy3 ──────────────────────────────────────────────────
def fetch_adequacy3_for_date(date_str: str) -> tuple[bytes, str]:
    """
    Fetch the latest (highest-version) Adequacy3 file for YYYYMMDD.
    Falls back to the un-versioned file if no versioned files exist yet.
    Returns (xml_bytes, filename_used).
    """
    dir_html = _fetch(config.IESO_ADEQUACY3_DIR_URL).decode("utf-8", errors="replace")
    version_pattern = rf"PUB_Adequacy3_{date_str}_v(\d+)\.xml"
    versions = [int(v) for v in re.findall(version_pattern, dir_html)]

    if versions:
        latest = max(versions)
        fname = f"PUB_Adequacy3_{date_str}_v{latest}.xml"
    else:
        # Fall back to the un-versioned daily file (also exists on the server)
        fname = f"PUB_Adequacy3_{date_str}.xml"

    url = f"{config.IESO_ADEQUACY3_BASE_URL}/{fname}"
    return _fetch(url), fname


# ── NYISO ATC/TTC ───────────────────────────────────────────────────
def fetch_nyiso_atc_ttc(date_str: str) -> bytes:
    """Fetch NYISO ATC/TTC HTML for YYYYMMDD (no auth required)."""
    url = config.NYISO_ATC_TTC_URL.format(date=date_str)
    return _fetch(url)
