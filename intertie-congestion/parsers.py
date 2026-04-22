"""
Parsers for IESO XML reports and NYISO ATC/TTC HTML.
"""
from __future__ import annotations

import re
import xml.etree.ElementTree as ET

IESO_NS = "{http://www.ieso.ca/schema}"


# ── Small XML helpers ───────────────────────────────────────────────
def _find(el, name):
    return el.find(IESO_NS + name) if el is not None else None


def _findall(el, name):
    return el.findall(IESO_NS + name) if el is not None else []


def _text(el, name, default=None):
    child = _find(el, name)
    return child.text if child is not None else default


def _parse_hourly_mw(container, child_tag: str, value_tag: str = "EnergyMW") -> dict:
    """Generic helper: container → [child_tag: {DeliveryHour, value_tag}] → {hour: int}."""
    out = {}
    if container is None:
        return out
    for entry in _findall(container, child_tag):
        h_el = _find(entry, "DeliveryHour")
        v_el = _find(entry, value_tag)
        if h_el is None or v_el is None or h_el.text is None or v_el.text is None:
            continue
        try:
            out[int(h_el.text)] = int(v_el.text)
        except ValueError:
            continue
    return out


# ── PredispIntertieSchedLimits ──────────────────────────────────────
def parse_predisp_limits(xml_bytes: bytes) -> dict:
    """
    Returns:
      {
        "created_at":   "2026-04-20T18:08:33",
        "delivery_date": "2026-04-20",
        "zones": {
          "MBSIN": {1: -50, 2: -50, ..., 24: -50},
          "NYSIX": {1: 1400, ...},
          ...
        }
      }
    """
    root = ET.fromstring(xml_bytes)
    header = _find(root, "DocHeader")
    body = _find(root, "DocBody")
    out = {
        "created_at": _text(header, "CreatedAt"),
        "delivery_date": _text(body, "DeliveryDate"),
        "zones": {},
    }
    for zone_el in _findall(body, "IntertieZonalEnergies"):
        zname = _text(zone_el, "IntertieZoneName")
        hourly = _find(zone_el, "HourlyEnergies")
        if zname is None or hourly is None:
            continue
        out["zones"][zname] = _parse_hourly_mw(hourly, "HourlyEnergy")
    return out


# ── Adequacy3 ───────────────────────────────────────────────────────
def parse_adequacy3(xml_bytes: bytes) -> dict:
    """
    Returns:
      {
        "created_at": str, "delivery_date": str,
        "ontario_demand":   {hour: MW},
        "excess":           {hour: MW},   # Excess Capacity (adequacy metric)
        "excess_offered":   {hour: MW},
        "total_imports":    {"offered": {h: MW}, "scheduled": {h: MW}},
        "total_exports":    {"bid":     {h: MW}, "scheduled": {h: MW}},
        "zonal_imports":    {
            "New York": {"offered": {h: MW}, "scheduled": {h: MW}}, ...
        },
        "zonal_exports":    {
            "New York": {"bid": {h: MW}, "scheduled": {h: MW}}, ...
        },
      }
    """
    root = ET.fromstring(xml_bytes)
    header = _find(root, "DocHeader")
    body = _find(root, "DocBody")
    supply = _find(body, "ForecastSupply")
    demand = _find(body, "ForecastDemand")

    out = {
        "created_at": _text(header, "CreatedAt"),
        "delivery_date": _text(body, "DeliveryDate"),
        "ontario_demand": {},
        "excess": {},
        "excess_offered": {},
        "zonal_imports": {},
        "zonal_exports": {},
        "total_imports": {"offered": {}, "scheduled": {}},
        "total_exports": {"bid": {}, "scheduled": {}},
    }

    # Ontario forecast demand
    ont = _find(demand, "OntarioDemand")
    if ont is not None:
        fcst = _find(ont, "ForecastOntDemand")
        out["ontario_demand"] = _parse_hourly_mw(fcst, "Demand")

    # Excess / adequacy metrics
    out["excess"] = _parse_hourly_mw(_find(demand, "ExcessCapacities"), "Capacity")
    out["excess_offered"] = _parse_hourly_mw(
        _find(demand, "ExcessOfferedCapacities"), "Capacity"
    )

    # Zonal imports (supply side)
    zimps = _find(supply, "ZonalImports")
    if zimps is not None:
        for zimp in _findall(zimps, "ZonalImport"):
            name = _text(zimp, "ZoneName")
            if not name:
                continue
            out["zonal_imports"][name] = {
                "offered": _parse_hourly_mw(_find(zimp, "Offers"), "Offer"),
                "scheduled": _parse_hourly_mw(_find(zimp, "Schedules"), "Schedule"),
            }
        total = _find(zimps, "TotalImports")
        if total is not None:
            out["total_imports"]["offered"] = _parse_hourly_mw(
                _find(total, "Offers"), "Offer"
            )
            out["total_imports"]["scheduled"] = _parse_hourly_mw(
                _find(total, "Schedules"), "Schedule"
            )

    # Zonal exports (demand side)
    zexps = _find(demand, "ZonalExports")
    if zexps is not None:
        for zexp in _findall(zexps, "ZonalExport"):
            name = _text(zexp, "ZoneName")
            if not name:
                continue
            out["zonal_exports"][name] = {
                "bid": _parse_hourly_mw(_find(zexp, "Bids"), "Bid"),
                "scheduled": _parse_hourly_mw(_find(zexp, "Schedules"), "Schedule"),
            }
        total = _find(zexps, "TotalExports")
        if total is not None:
            out["total_exports"]["bid"] = _parse_hourly_mw(
                _find(total, "Bids"), "Bid"
            )
            out["total_exports"]["scheduled"] = _parse_hourly_mw(
                _find(total, "Schedules"), "Schedule"
            )

    return out


# ── NYISO ATC/TTC HTML ──────────────────────────────────────────────
def parse_nyiso_atc(html_bytes: bytes, interface: str) -> list[dict]:
    """
    Parse one interface section (e.g. 'IMO-NYISO') from the NYISO ATC/TTC HTML.

    Returns a list of hourly dicts, one per row, sorted by hour:
      {"hour": 0-23, "dam_ttc": int, "dam_atc": int,
       "ham_ttc_00": int|None, "ham_atc_00": int|None, ...,
       "ham_ttc_45": int|None, "ham_atc_45": int|None}

    A row may have only DAM values (2 columns) near end-of-day; HAM fields
    will be None in that case.
    """
    html = html_bytes.decode("utf-8", errors="replace")

    # Locate the interface section
    anchor_re = re.compile(
        rf'<a name="{re.escape(interface)}"></a>\s*Interface:',
        re.IGNORECASE,
    )
    m = anchor_re.search(html)
    if not m:
        return []
    section = html[m.end():]
    next_anchor = re.search(r'<a name="[^"]+"></a>\s*Interface:', section)
    if next_anchor:
        section = section[:next_anchor.start()]

    # Split into per-row chunks (more reliable than spanning regexes across
    # multiple rows, especially when short rows follow full rows).
    time_re = re.compile(r'<td>\s*(\d{2}):00\s*[A-Z]{3}\s*</td>(.*?)</tr>',
                         re.IGNORECASE | re.DOTALL)
    cell_re = re.compile(r'<td>\s*(-?\d+)\s*</td>', re.IGNORECASE)
    ham_keys = ["ham_ttc_00", "ham_atc_00", "ham_ttc_15", "ham_atc_15",
                "ham_ttc_30", "ham_atc_30", "ham_ttc_45", "ham_atc_45"]

    rows: dict[int, dict] = {}
    for m in time_re.finditer(section):
        h = int(m.group(1))
        cells = [int(c) for c in cell_re.findall(m.group(2))]
        row = {
            "hour": h,
            "dam_ttc": cells[0] if len(cells) > 0 else None,
            "dam_atc": cells[1] if len(cells) > 1 else None,
        }
        # Fill HAM fields (8 columns after DAM) with None if missing
        for i, key in enumerate(ham_keys):
            row[key] = cells[i + 2] if len(cells) > i + 2 else None
        rows[h] = row
    return [rows[h] for h in sorted(rows)]
