#!/usr/bin/env python3
"""
Ontario Intertie Congestion Scanner
===================================
Pulls IESO PredispIntertieSchedLimits + today's Adequacy3 + NYISO ATC/TTC and
prints a multi-hour congestion picture for every Ontario intertie
(NY, MI, MN, MB, PQ). The NY tie gets the dual-sided view with the
300 MW security buffer applied.

Usage:
    python3 scan.py                 # print report to terminal
    python3 scan.py --notify        # also send macOS notification if alerts
    python3 scan.py --hours 8       # show 8 forward HEs instead of 6

Press Ctrl+C to stop.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta

import config
import fetchers
import parsers


# ── Small utilities ─────────────────────────────────────────────────
def get_now_et() -> datetime:
    """Return current time in Eastern Standard Time (fixed UTC-5, no DST)."""
    return datetime.now(config.EST)


def mw_abs(v):
    return abs(v) if isinstance(v, int) else None


def classify(remaining, limit):
    """
    Return (icon, label) for a given remaining-MW and total capacity.

    Uses % utilization as the primary signal. The absolute-MW thresholds
    only apply on large ties (capacity > 500 MW), because on small ties
    like Manitoba (~125 MW) an absolute "<100 MW remaining" check is
    almost always true regardless of utilization.
    """
    if remaining is None or limit is None or limit <= 0:
        return "⚪", "n/a"
    util = 100 - (remaining / limit * 100)
    big_tie = limit > 500

    if util > config.CRITICAL_PCT or (big_tie and remaining < config.CRITICAL_MW):
        return "🔴", "critical"
    if util > config.WARNING_PCT or (big_tie and remaining < config.WARNING_MW):
        return "⚠️ ", "warning"
    return "✅", "open"


# ── Row formatting ──────────────────────────────────────────────────
def _row_int(label, values, width=24):
    cells = [f"{v:>5}" if isinstance(v, int) else "    —" for v in values]
    return f"{label:<{width}}" + "  ".join(cells)


def _row_str(label, values, width=24):
    cells = [f"{str(v):>5}" for v in values]
    return f"{label:<{width}}" + "  ".join(cells)


# ── Zone data builder ───────────────────────────────────────────────
def build_zone_data(limits, adeq):
    """
    Combine IESO limits + Adequacy3 into per-zone dicts.
    Returns {zkey: {name, import_limit, export_limit,
                    import_scheduled, export_scheduled,
                    import_offered, export_bid}}
    """
    zones = {}
    for zkey in config.DISPLAY_ORDER:
        spec = config.IESO_ZONES[zkey]
        z = {
            "name": spec["name"],
            "import_limit": {}, "export_limit": {},
            "import_scheduled": {}, "export_scheduled": {},
            "import_offered": {}, "export_bid": {},
        }

        # IESO limits (direct zone lookup)
        if spec["import"]:
            raw = limits["zones"].get(spec["import"], {})
            z["import_limit"] = {h: abs(v) for h, v in raw.items()}
        if spec["export"]:
            raw = limits["zones"].get(spec["export"], {})
            z["export_limit"] = {h: abs(v) for h, v in raw.items()}

        # Adequacy3 scheduled/offered/bid.
        # Adequacy3 stores exports as NEGATIVE values (sign convention on
        # the demand side). Take abs() so "scheduled" is always magnitude.
        if adeq:
            adeq_name = spec["adeq"]
            zi = adeq["zonal_imports"].get(adeq_name, {})
            ze = adeq["zonal_exports"].get(adeq_name, {})
            z["import_scheduled"] = {h: abs(v) for h, v in zi.get("scheduled", {}).items()}
            z["import_offered"] = {h: abs(v) for h, v in zi.get("offered", {}).items()}
            z["export_scheduled"] = {h: abs(v) for h, v in ze.get("scheduled", {}).items()}
            z["export_bid"] = {h: abs(v) for h, v in ze.get("bid", {}).items()}

        zones[zkey] = z

    # Quebec: sum all PQ sub-zones
    # IESO convention: N = export ("to"), X = import ("from")
    pq_imp, pq_exp = {}, {}
    for sub in config.QUEBEC_SUBZONES:
        for h, v in limits["zones"].get(f"PQ{sub}X", {}).items():
            pq_imp[h] = pq_imp.get(h, 0) + abs(v)
        for h, v in limits["zones"].get(f"PQ{sub}N", {}).items():
            pq_exp[h] = pq_exp.get(h, 0) + abs(v)
    zones["PQ"]["import_limit"] = pq_imp
    zones["PQ"]["export_limit"] = pq_exp
    return zones


# ── NYISO row helper ────────────────────────────────────────────────
def _nyiso_for_he(nyiso_rows, he):
    """HE h covers clock hour (h-1) → (h). Get the NYISO row for clock hour h-1."""
    if not nyiso_rows:
        return None
    clock_h = he - 1
    return next((r for r in nyiso_rows if r["hour"] == clock_h), None)


def _binding_limit(iso_limit, nyiso_row):
    """Return min(IESO limit, NYISO TTC−300). None-safe."""
    ny_real = None
    if nyiso_row is not None and nyiso_row.get("dam_ttc") is not None:
        ny_real = max(0, nyiso_row["dam_ttc"] - config.NYISO_SECURITY_BUFFER_MW)
    if iso_limit is not None and ny_real is not None:
        return min(iso_limit, ny_real)
    return iso_limit if iso_limit is not None else ny_real


# ── Direction table renderer ────────────────────────────────────────
def render_direction(z, hours, direction, nyiso_rows=None):
    """Print the MW/hour table for one direction (export or import) of one zone."""
    limit_map = z[f"{direction}_limit"]
    sched_map = z[f"{direction}_scheduled"]
    dir_word = "Export" if direction == "export" else "Import"

    # Header line: HE columns
    header = "  ".join(f"HE{h:>2}" for h in hours)
    print(f"                        {header}")

    # IESO limit row
    print(_row_int(f"  IESO {dir_word} Limit:", [limit_map.get(h) for h in hours]))

    # NYISO (only for NY)
    binding_row = []
    if nyiso_rows:
        ttc_row, real_row, atc_row = [], [], []
        for he in hours:
            nr = _nyiso_for_he(nyiso_rows, he)
            if nr:
                ttc = nr["dam_ttc"]
                atc = nr["dam_atc"]
                real = max(0, ttc - config.NYISO_SECURITY_BUFFER_MW)
                ttc_row.append(ttc); atc_row.append(atc); real_row.append(real)
            else:
                ttc_row.append(None); atc_row.append(None); real_row.append(None)
        print(_row_int("  NYISO TTC (listed):", ttc_row))
        print(_row_int("  NYISO TTC (−300):", real_row))
        print(_row_int("  NYISO ATC (DAM):", atc_row))

        for i, he in enumerate(hours):
            binding_row.append(_binding_limit(limit_map.get(he),
                                              _nyiso_for_he(nyiso_rows, he)))
        print(_row_int("  Binding Limit:", binding_row))
    else:
        for he in hours:
            binding_row.append(limit_map.get(he))

    # Scheduled / remaining / utilization / status
    print(_row_int("  Already Scheduled:", [sched_map.get(h) for h in hours]))

    remaining_row = []
    for i, he in enumerate(hours):
        lim = binding_row[i]
        sched = sched_map.get(he)
        if lim is not None and sched is not None:
            remaining_row.append(max(0, lim - sched))
        else:
            remaining_row.append(None)
    print(_row_int("  Remaining Room:", remaining_row))

    util_row = []
    for i, he in enumerate(hours):
        lim = binding_row[i]
        sched = sched_map.get(he)
        if lim and lim > 0 and sched is not None:
            util_row.append(f"{int(round(sched / lim * 100))}%")
        else:
            util_row.append("—")
    print(_row_str("  Utilization:", util_row))

    status_row = []
    for i, he in enumerate(hours):
        icon, _ = classify(remaining_row[i], binding_row[i])
        status_row.append(icon)
    print(_row_str("  Status:", status_row))


# ── Full report ─────────────────────────────────────────────────────
def print_report(now, limits, adeq, zones, nyiso_exp, nyiso_imp, hours):
    w = 62
    print()
    print("═" * w)
    print("  ONTARIO INTERTIE CONGESTION SUMMARY")
    print(f"  Generated: {now.strftime('%Y-%m-%d %H:%M')} EST")
    print(f"  IESO limits report: {limits['created_at']}  (delivery {limits['delivery_date']})")
    if adeq:
        print(f"  Adequacy3 report:   {adeq['created_at']}  (delivery {adeq['delivery_date']})")
    print("═" * w)

    # ── System status ──
    if adeq:
        print()
        print("ONTARIO SYSTEM STATUS")
        dem_first = adeq["ontario_demand"].get(hours[0])
        dem_last = adeq["ontario_demand"].get(hours[-1])
        exc_vals = [adeq["excess"].get(h) for h in hours if isinstance(adeq["excess"].get(h), int)]
        print(f"  Demand Forecast (HE{hours[0]}–HE{hours[-1]}): "
              f"{dem_first if dem_first is not None else '—'} → "
              f"{dem_last if dem_last is not None else '—'} MW")
        if exc_vals:
            print(f"  Surplus/Shortfall (Excess Capacity): "
                  f"{min(exc_vals):+} → {max(exc_vals):+} MW "
                  f"(range over window)")
            if min(exc_vals) <= 0:
                status = "🔴 SHORTFALL"
            elif min(exc_vals) < 500:
                status = "⚠️  TIGHT"
            else:
                status = "✅ Adequate"
            print(f"  Status: {status}  (min excess {min(exc_vals):+} MW)")

    # ── Per-zone ──
    for zkey in config.DISPLAY_ORDER:
        z = zones[zkey]
        print()
        print("─" * w)
        tag = "  ← PRIMARY CORRIDOR" if zkey == "NY" else ""
        print(f"{z['name']}{tag}")
        print("─" * w)

        print(f"Direction: ON → {zkey} (Export)")
        render_direction(z, hours, "export",
                         nyiso_rows=nyiso_exp if zkey == "NY" else None)

        print()
        print(f"Direction: {zkey} → ON (Import)")
        render_direction(z, hours, "import",
                         nyiso_rows=nyiso_imp if zkey == "NY" else None)


# ── Alerts ──────────────────────────────────────────────────────────
def build_alerts(zones, nyiso_exp, nyiso_imp, hours):
    """Build short alert bullets for the ALERTS section + notifications."""
    alerts = []
    good_zones = []

    for zkey in config.DISPLAY_ORDER:
        z = zones[zkey]
        for direction in ("export", "import"):
            limit_map = z[f"{direction}_limit"]
            sched_map = z[f"{direction}_scheduled"]
            nyiso_rows = None
            if zkey == "NY":
                nyiso_rows = nyiso_exp if direction == "export" else nyiso_imp

            crit_h, warn_h = [], []
            any_data = False
            for he in hours:
                iso_lim = limit_map.get(he)
                lim = _binding_limit(iso_lim, _nyiso_for_he(nyiso_rows, he)) if nyiso_rows else iso_lim
                sched = sched_map.get(he)
                if lim is None or sched is None:
                    continue
                any_data = True
                remaining = max(0, lim - sched)
                icon, label = classify(remaining, lim)
                if label == "critical":
                    crit_h.append(he)
                elif label == "warning":
                    warn_h.append(he)

            dir_label = f"ON → {zkey}" if direction == "export" else f"{zkey} → ON"
            if crit_h:
                alerts.append(
                    f"🔴 {dir_label}: CRITICAL for HE{crit_h[0]}"
                    + (f"–HE{crit_h[-1]}" if len(crit_h) > 1 else "")
                    + " (near full)"
                )
            elif warn_h:
                alerts.append(
                    f"⚠️  {dir_label}: tight HE{warn_h[0]}"
                    + (f"–HE{warn_h[-1]}" if len(warn_h) > 1 else "")
                )
            elif any_data:
                good_zones.append(dir_label)

    return alerts, good_zones


# ── Notification ────────────────────────────────────────────────────
def notify(title, message):
    try:
        subprocess.run(
            [
                "terminal-notifier",
                "-title", title,
                "-message", message,
                "-sound", "Glass",
                "-group", "intertie-congestion",
            ],
            capture_output=True, timeout=5,
        )
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass


# ── JSON export (for HTML dashboard) ────────────────────────────────
def export_payload(now, limits, adeq, zones, nyiso_exp, nyiso_imp, hours, alerts):
    """Build the payload used by the HTML dashboard."""
    payload = {
        "generated_at": now.isoformat(),
        "hours": hours,
        "limits_created_at": limits["created_at"],
        "limits_delivery_date": limits["delivery_date"],
        "adeq_created_at": adeq["created_at"] if adeq else None,
        "nyiso_url": config.NYISO_ATC_TTC_URL.format(date=now.strftime("%Y%m%d")),
        "ontario_demand": {h: adeq["ontario_demand"].get(h) for h in hours} if adeq else {},
        "excess": {h: adeq["excess"].get(h) for h in hours} if adeq else {},
        "zones": {},
        "alerts": alerts,
    }

    for zkey in config.DISPLAY_ORDER:
        z = zones[zkey]
        z_out = {"name": z["name"], "export": {}, "import": {}}
        for direction in ("export", "import"):
            limit_map = z[f"{direction}_limit"]
            sched_map = z[f"{direction}_scheduled"]
            rows = []
            for he in hours:
                iso_lim = limit_map.get(he)
                nyiso_rows = None
                if zkey == "NY":
                    nyiso_rows = nyiso_exp if direction == "export" else nyiso_imp
                nr = _nyiso_for_he(nyiso_rows, he) if nyiso_rows else None
                ny_ttc = nr["dam_ttc"] if nr else None
                ny_real = max(0, ny_ttc - config.NYISO_SECURITY_BUFFER_MW) if ny_ttc is not None else None
                ny_atc = nr["dam_atc"] if nr else None
                binding = _binding_limit(iso_lim, nr) if nyiso_rows else iso_lim
                sched = sched_map.get(he)
                remaining = max(0, binding - sched) if (binding is not None and sched is not None) else None
                util = int(round(sched / binding * 100)) if (binding and binding > 0 and sched is not None) else None
                icon, status = classify(remaining, binding)
                rows.append({
                    "he": he,
                    "iso_limit": iso_lim,
                    "ny_ttc": ny_ttc,
                    "ny_real": ny_real,
                    "ny_atc": ny_atc,
                    "binding": binding,
                    "scheduled": sched,
                    "remaining": remaining,
                    "utilization": util,
                    "status": status,  # "open"|"warning"|"critical"|"n/a"
                })
            z_out[direction] = rows
        payload["zones"][zkey] = z_out
    return payload


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="300">
<title>IESO Congestion Picture __WINDOW__</title>
<style>
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,monospace;
       background:#0d1117;color:#c9d1d9;padding:20px 24px;min-height:100vh}
  h1{font-size:20px;font-weight:700;letter-spacing:.3px;margin-bottom:6px;color:#f0f6fc}
  h1 span{color:#3fb950}
  .meta{font-size:12px;color:#c9d1d9;margin-bottom:6px}
  .sources{font-size:12px;color:#c9d1d9;margin-bottom:20px}
  .sources a{color:#58a6ff;text-decoration:none;margin-right:14px;font-weight:500}
  .sources a:hover{text-decoration:underline}

  /* Notification preview — macOS style */
  .notif-label{font-size:11px;color:#8b949e;margin-bottom:6px;
               text-transform:uppercase;letter-spacing:.5px}
  .notif{background:rgba(28,33,40,.85);backdrop-filter:blur(20px);
         border:1px solid #30363d;border-radius:12px;padding:10px 14px;
         display:flex;gap:12px;align-items:flex-start;
         max-width:420px;margin-bottom:20px;
         box-shadow:0 8px 24px rgba(0,0,0,.4)}
  .notif-icon{width:36px;height:36px;border-radius:8px;
              background:linear-gradient(135deg,#3fb950,#238636);
              display:flex;align-items:center;justify-content:center;
              flex-shrink:0;font-size:18px;color:#fff}
  .notif-body{flex:1;min-width:0}
  .notif-app{font-size:10px;color:#8b949e;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}
  .notif-title{font-size:13px;font-weight:600;color:#f0f6fc;margin-bottom:3px}
  .notif-msg{font-size:12px;color:#c9d1d9;white-space:pre-wrap;line-height:1.4}
  .notif-time{font-size:10px;color:#8b949e;flex-shrink:0}

  .zone{background:#161b22;border:1px solid #30363d;border-radius:8px;
        margin-bottom:14px;padding:14px 18px}
  .zone-head{display:flex;align-items:baseline;gap:12px;margin-bottom:10px}
  .zone-name{font-size:15px;font-weight:600;color:#f0f6fc}
  .system-box{background:#161b22;border:1px solid #30363d;border-radius:8px;
              margin-bottom:14px;padding:14px 18px}
  .system-box h3{font-size:14px;color:#f0f6fc;margin-bottom:10px;font-weight:600}
  tr.demand-row td{color:#f0f6fc;font-weight:500}
  tr.excess-row td{color:#c9d1d9;font-weight:500}
  tr.excess-tight td:first-child,tr.excess-tight td{color:#fbbf24}
  tr.excess-short td:first-child,tr.excess-short td{color:#f85149}
  .dir-block{margin-top:10px}
  .dir-block + .dir-block{
    margin-top:28px;
    padding-top:20px;
    border-top:2px solid #30363d;
  }
  .dir-label{font-size:14px;color:#f0f6fc;margin-bottom:10px;
             text-transform:uppercase;letter-spacing:.6px;font-weight:700;
             display:inline-block;padding:4px 10px;border-radius:4px;
             background:rgba(139,148,158,.12);border-left:3px solid #6e7681}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{padding:6px 10px;text-align:right;border:1px solid #30363d;white-space:nowrap}
  th{background:#1c2128;color:#e5e7eb;font-weight:600;text-align:center}
  th:first-child,td:first-child{text-align:left;color:#f0f6fc;font-weight:600}
  td.status{text-align:center;font-size:14px}
  tr.status-critical td.status{background:rgba(248,81,73,.25)}
  tr.status-warning  td.status{background:rgba(251,191,36,.22)}
  tr.status-open     td.status{background:rgba(63,185,80,.18)}
  td.status-critical{background:rgba(248,81,73,.15);color:#f85149}
  td.status-warning {background:rgba(251,191,36,.12);color:#fbbf24}
  td.status-open    {background:rgba(63,185,80,.10);color:#3fb950}
  .na{color:#c9d1d9}
  .corridor-picker{display:flex;gap:10px;align-items:center;margin-bottom:14px;
                   font-size:14px}
  .corridor-picker label{color:#e5e7eb;font-weight:500}
  .corridor-picker select{background:#1c2128;color:#c9d1d9;border:1px solid #30363d;
                          border-radius:6px;padding:6px 10px;font-size:13px;cursor:pointer;
                          font-family:inherit}
  .corridor-picker select:focus{outline:2px solid #3fb950;outline-offset:1px}
  .legend-swatch{display:inline-block;width:12px;height:12px;border-radius:2px;
                 vertical-align:middle;margin-right:4px}
  .legend-bar{display:flex;gap:16px;font-size:12px;color:#c9d1d9;margin-top:4px;
              margin-bottom:10px;flex-wrap:wrap}

  /* HE-row tables + capacity gauge */
  .flow-table{width:100%;border-collapse:separate;border-spacing:0;font-size:13px}
  .flow-table th,.flow-table td{padding:8px 10px;border-bottom:1px solid #21262d;white-space:nowrap}
  .flow-table th{background:#1c2128;color:#e5e7eb;font-weight:600;text-align:right;
                 font-size:12px;text-transform:uppercase;letter-spacing:.4px}
  .flow-table th:first-child,.flow-table td:first-child{text-align:left;font-weight:600;color:#f0f6fc}
  .flow-table td{text-align:right;color:#e5e7eb}
  .flow-table td.sched-exp{color:#f59e0b;font-weight:700}
  .flow-table td.sched-imp{color:#60a5fa;font-weight:700}
  .flow-table td.util-critical{color:#f85149;background:rgba(248,81,73,.18);font-weight:700}
  .flow-table td.util-warning {color:#fbbf24;background:rgba(251,191,36,.16);font-weight:700}
  .flow-table td.util-open    {color:#3fb950;background:rgba(63,185,80,.10);font-weight:700}
  /* Focus columns: Remain + Util — always visually lifted */
  .flow-table td.focus-col{background:rgba(88,166,255,.06);font-weight:600;color:#f0f6fc}
  .flow-table td.util-critical.focus-col{background:rgba(248,81,73,.22)}
  .flow-table td.util-warning.focus-col {background:rgba(251,191,36,.20)}
  .flow-table td.util-open.focus-col    {background:rgba(63,185,80,.14)}
  .flow-table th.focus-col{background:#262d36;color:#f0f6fc;border-left:1px solid #3b434c;border-right:1px solid #3b434c}
  .flow-table td.focus-col{border-left:1px solid rgba(88,166,255,.20);border-right:1px solid rgba(88,166,255,.20)}
  .flow-table td.status-cell{text-align:center;width:48px;font-size:11px;font-weight:700;letter-spacing:.3px}
  .status-dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:4px;vertical-align:middle}
  .status-dot.critical{background:#f85149;box-shadow:0 0 6px rgba(248,81,73,.6)}
  .status-dot.warning{background:#fbbf24;box-shadow:0 0 6px rgba(251,191,36,.5)}
  .status-dot.open{background:#3fb950;box-shadow:0 0 6px rgba(63,185,80,.4)}
  .status-dot.na{background:#6e7681}
  .flow-table td.he-cell{color:#f0f6fc;font-weight:600}
  /* Past HEs: slightly darker bg band */
  .flow-table tr.past td{background:rgba(110,118,129,.10)}
  .flow-table tr.past td.he-cell{color:#c9d1d9}
  .flow-table tr.past td.he-cell::before{content:"← ";color:#8b949e;font-weight:400}
  /* Current HE: just mark the HE label with "· NOW" — no row highlight */
  .flow-table tr.current td.he-cell{color:#58a6ff;font-weight:800}
  .flow-table tr.current td.he-cell::after{content:" · NOW";font-size:10px;font-weight:700;color:#58a6ff;letter-spacing:.5px}
  /* Section heading rows (Past / Now-and-forward) */
  .flow-table tr.sec-head td{background:#0d1117;color:#8b949e;
    font-size:10px;text-transform:uppercase;letter-spacing:.5px;
    padding:8px 10px 4px;border-bottom:1px solid #30363d;text-align:left;font-weight:700}

  /* Capacity gauge cell */
  .gauge-cell{width:40%;padding:8px 12px !important}
  .gauge{position:relative;height:22px;background:transparent;
         border-radius:3px;overflow:visible}
  .gauge-track{position:absolute;top:0;height:100%;display:flex}
  .gauge-seg{height:100%}
  .gauge-seg.sched-exp{background:#fbbf24}
  .gauge-seg.sched-imp{background:#60a5fa}
  .gauge-seg.usable{background:rgba(255,255,255,.10)}
  .gauge-seg.blocked{background:repeating-linear-gradient(
         45deg,rgba(167,139,250,.32) 0 5px,rgba(124,92,235,.12) 5px 10px)}
  .gauge-tick{position:absolute;top:-2px;bottom:-2px;width:2px;background:#6e7681}
  .gauge-tick.bind{background:#f0f6fc}
  .gauge-label{position:absolute;font-size:10px;color:#c9d1d9;
               top:22px;transform:translateX(-50%);white-space:nowrap;line-height:1.2}
  .gauge-label.bind{color:#f0f6fc;font-weight:700}
  .gauge-zero{position:absolute;top:-3px;bottom:-3px;width:2px;background:#6e7681;left:50%}
</style>
</head>
<body>
  <h1>IESO Congestion Picture <span id="windowTitle" style="color:#c9d1d9"></span></h1>
  <div class="meta" id="meta"></div>
  <div class="sources" id="sources"></div>

  <div id="systemTable"></div>

  <div class="corridor-picker">
    <label for="corridorSelect">View corridor:</label>
    <select id="corridorSelect">
      <option value="NY">NYSI (New York — Roseton)</option>
      <option value="MI">MISI (Michigan — Ludington)</option>
      <option value="MN">MNSI (Minnesota — Intl Falls)</option>
      <option value="MB">MBSI (Manitoba)</option>
      <option value="PQ">PQSI (Quebec — summed)</option>
    </select>
  </div>

  <div id="zones"></div>

<script>
const DATA = __PAYLOAD__;
const ICONS = {critical:"🔴", warning:"⚠️", open:"✅", "n/a":"⚪"};
const FIRST_DATE = DATA.hours_spec[0][0];

function fmt(v){ return v==null ? '<span class="na">—</span>' : v; }
function fmtPct(v){ return v==null ? '<span class="na">—</span>' : v+'%'; }

function renderMeta(){
  const d = DATA.generated_at.slice(0,16).replace('T',' ');
  document.getElementById('windowTitle').textContent = DATA.title_window;
  document.getElementById('meta').innerHTML = `Generated at ${d} (EST) — page auto-refreshes every 5 min`;

  // Predisp limits: link per unique date (may be 1 or 2 files)
  const pdLinks = Object.entries(DATA.predisp_urls || {}).map(([d,u]) =>
    `<a href="${u}" target="_blank">ONT PD limits (${d.slice(4,6)}/${d.slice(6,8)}) ↗</a>`
  ).join(' ');
  const adeqLinks = Object.entries(DATA.adeq_urls).map(([d,u]) =>
    `<a href="${u}" target="_blank">IESO Adequacy3 (${d.slice(4,6)}/${d.slice(6,8)}) ↗</a>`
  ).join(' ');
  const nyLinks = Object.entries(DATA.nyiso_urls).map(([d,u]) =>
    `<a href="${u}" target="_blank">NYISO ATC/TTC (${d.slice(4,6)}/${d.slice(6,8)}) ↗</a>`
  ).join(' ');
  document.getElementById('sources').innerHTML = `${pdLinks} ${adeqLinks} ${nyLinks}`;
}

function renderNotifPreview(){
  const hasAlerts = DATA.has_alerts;
  const iconBg = hasAlerts ? 'linear-gradient(135deg,#f85149,#9e2626)' : 'linear-gradient(135deg,#3fb950,#238636)';
  const icon = hasAlerts ? '⚠️' : '✅';
  const body = (DATA.notif_body || '').replace(/</g,'&lt;');
  const time = new Date(DATA.generated_at).toLocaleTimeString('en-US',
    {hour:'numeric', minute:'2-digit', hour12:false});
  document.getElementById('notifBox').innerHTML = `
    <div class="notif">
      <div class="notif-icon" style="background:${iconBg}">${icon}</div>
      <div class="notif-body">
        <div class="notif-app">Terminal · now</div>
        <div class="notif-title">${DATA.notif_title}</div>
        <div class="notif-msg">${body}</div>
      </div>
      <div class="notif-time">${time}</div>
    </div>`;
}

function heHeaders(){
  return DATA.hours_spec.map(([d,he]) => `<th>HE ${he}</th>`).join('');
}

function renderSystemTable(){
  const hdrs = heHeaders();
  const demCells = DATA.ontario_demand.map(v => `<td>${fmt(v)}</td>`).join('');
  const excCells = DATA.excess.map(v => {
    if (v == null) return '<td><span class="na">—</span></td>';
    return `<td>${v>=0?'+':''}${v}</td>`;
  }).join('');
  const excVals = DATA.excess.filter(v => v!=null);
  const minExc = excVals.length ? Math.min(...excVals) : null;
  let excClass = '';
  if (minExc != null){
    if (minExc <= 0) excClass = 'excess-short';
    else if (minExc < 500) excClass = 'excess-tight';
  }
  document.getElementById('systemTable').innerHTML = `
    <div class="system-box">
      <h3>Ontario System</h3>
      <table>
        <thead><tr><th>Limits / Availability</th>${hdrs}</tr></thead>
        <tbody>
          <tr class="demand-row"><td>Demand Forecast (MW)</td>${demCells}</tr>
          <tr class="excess-row ${excClass}"><td>Excess Capacity (MW)</td>${excCells}</tr>
        </tbody>
      </table>
    </div>`;
}

// ── HE-row tables with inline capacity gauges ──────────────────────
function gaugeCell(row, direction, scaleMax, isNY){
  // Each segment width is a % of scaleMax.
  // Order (left-to-right): scheduled → usable → blocked.
  if (row.binding == null){
    return '<div class="gauge"><span style="color:#8b949e;font-size:11px">—</span></div>';
  }
  const sched = row.scheduled != null ? Math.abs(row.scheduled) : 0;
  const usable = row.remaining != null ? row.remaining : 0;
  const other = (isNY && row.ny_real != null && row.iso_limit != null)
                ? Math.max(row.ny_real, row.iso_limit) : row.binding;
  const blocked = Math.max(0, other - row.binding);
  const pct = v => (v / scaleMax * 100).toFixed(2) + '%';
  const schedCls = direction === 'export' ? 'sched-exp' : 'sched-imp';

  // Segments
  let segs = `<div class="gauge-seg ${schedCls}" style="width:${pct(sched)}"
                title="Scheduled ${sched} MW"></div>`;
  segs += `<div class="gauge-seg usable" style="width:${pct(usable)}"
              title="Usable room ${usable} MW"></div>`;
  if (blocked > 0 && isNY){
    segs += `<div class="gauge-seg blocked" style="width:${pct(blocked)}"
                title="Held back by the tighter limit: ${blocked} MW"></div>`;
  }

  // Ticks: binding (bright white), IESO & NYISO (muted)
  const ticks = [];
  ticks.push(`<div class="gauge-tick bind" style="left:${pct(row.binding)}" title="Binding ${row.binding} MW"></div>`);
  if (isNY && row.iso_limit != null && row.iso_limit !== row.binding){
    ticks.push(`<div class="gauge-tick" style="left:${pct(row.iso_limit)}" title="IESO limit ${row.iso_limit} MW"></div>`);
  }
  if (isNY && row.ny_real != null && row.ny_real !== row.binding){
    ticks.push(`<div class="gauge-tick" style="left:${pct(row.ny_real)}" title="NYISO TTC−300 = ${row.ny_real} MW"></div>`);
  }

  return `<div class="gauge"><div class="gauge-track" style="width:100%">${segs}</div>${ticks.join('')}</div>`;
}

function renderDirectionTable(k, direction, rows, isNY){
  // Scale max = max binding + blocked across all rows in this direction (so all gauges share scale).
  let scaleMax = 0;
  for (const r of rows){
    if (r.binding != null) scaleMax = Math.max(scaleMax, r.binding);
    if (isNY && r.iso_limit != null) scaleMax = Math.max(scaleMax, r.iso_limit);
    if (isNY && r.ny_real != null) scaleMax = Math.max(scaleMax, r.ny_real);
  }
  if (scaleMax <= 0) scaleMax = 1;

  const dirWord = direction === 'export' ? 'Export' : 'Import';
  const arrow = direction === 'export' ? 'ON → ' + k : k + ' → ON';

  const cols = [];
  cols.push(`<th>HE</th>`);
  cols.push(`<th>ONT PD<br><span style="font-size:10px;font-weight:400">scheduling limit</span></th>`);
  if (isNY){
    cols.push(`<th>NY TTC<br><span style="color:#60a5fa;font-size:10px">(listed)</span></th>`);
    cols.push(`<th>NY TTC<br><span style="color:#60a5fa;font-size:10px">(−300)</span></th>`);
    cols.push(`<th>NY ATC</th>`);
    cols.push(`<th>Binding</th>`);
  }
  cols.push(`<th>Scheduled</th>`);
  cols.push(`<th class="focus-col">Remain</th>`);
  cols.push(`<th class="focus-col">Util</th>`);
  cols.push(`<th class="gauge-cell">0 MW → ${scaleMax} MW (capacity envelope)</th>`);
  cols.push(`<th>Status</th>`);

  let html = `<div class="dir-block">
    <div class="dir-label">${arrow} (${dirWord})</div>
    <table class="flow-table">
      <thead><tr>${cols.join('')}</tr></thead>
      <tbody>`;

  const STATUS_LABEL = {critical:'CRIT', warning:'WARN', open:'OK', 'n/a':'—'};
  const STATUS_COLOR = {critical:'#f85149', warning:'#fbbf24', open:'#3fb950', 'n/a':'#c9d1d9'};

  // Determine which hour-spec index is "current" based on now wall-clock (EST)
  // Current HE = hour+1 in EST. Server-calculated once at page load time.
  const nowEst = new Date(Date.now() + (new Date().getTimezoneOffset() - 300) * 60000);
  const currentHe = (nowEst.getHours() + 1);
  const currentDate = [nowEst.getFullYear(),
    String(nowEst.getMonth()+1).padStart(2,'0'),
    String(nowEst.getDate()).padStart(2,'0')].join('');
  let currentIdx = DATA.hours_spec.findIndex(([d,he]) => d === currentDate && he === currentHe);
  if (currentIdx < 0) currentIdx = 0;  // fallback: first row is "current"

  const colSpan = isNY ? 11 : 8;

  for (let i = 0; i < rows.length; i++){
    const r = rows[i];
    const [d, he] = DATA.hours_spec[i];
    const utilCls = r.status === 'critical' ? 'util-critical'
                  : r.status === 'warning'  ? 'util-warning'
                  : r.status === 'open'     ? 'util-open' : '';
    const schedCls = direction === 'export' ? 'sched-exp' : 'sched-imp';
    const schedVal = r.scheduled == null ? '—'
                   : (r.scheduled > 0 ? '+'+r.scheduled : r.scheduled);

    // Section headers
    if (i === 0 && currentIdx > 0){
      html += `<tr class="sec-head"><td colspan="${colSpan}">Past hours (last ${currentIdx}h · already cleared)</td></tr>`;
    }
    if (i === currentIdx && currentIdx >= 0){
      html += `<tr class="sec-head"><td colspan="${colSpan}">Current hour &amp; forecast (pre-dispatch)</td></tr>`;
    }

    const rowCls = i < currentIdx ? 'past' : (i === currentIdx ? 'current' : '');
    const cells = [];
    cells.push(`<td class="he-cell">HE ${he}</td>`);
    cells.push(`<td>${fmt(r.iso_limit)}</td>`);
    if (isNY){
      cells.push(`<td>${fmt(r.ny_ttc)}</td>`);
      cells.push(`<td>${fmt(r.ny_real)}</td>`);
      cells.push(`<td>${fmt(r.ny_atc)}</td>`);
      cells.push(`<td>${fmt(r.binding)}</td>`);
    }
    cells.push(`<td class="${schedCls}">${schedVal}</td>`);
    cells.push(`<td class="focus-col">${fmt(r.remaining)}</td>`);
    cells.push(`<td class="focus-col ${utilCls}">${r.utilization != null ? r.utilization+'%' : '—'}</td>`);
    cells.push(`<td class="gauge-cell">${gaugeCell(r, direction, scaleMax, isNY)}</td>`);
    const sColor = STATUS_COLOR[r.status] || '#8b949e';
    const sLabel = STATUS_LABEL[r.status] || '—';
    cells.push(`<td class="status-cell" style="color:${sColor}">
      <span class="status-dot ${r.status}"></span>${sLabel}</td>`);

    html += `<tr class="${rowCls}">${cells.join('')}</tr>`;
  }

  html += `</tbody></table></div>`;
  return html;
}

function renderSelectedZone(k){
  const container = document.getElementById('zones');
  const z = DATA.zones[k];
  if (!z){ container.innerHTML = ''; return; }
  const isNY = (k === 'NY');
  const legend = `
    <div class="legend-bar">
      <span><span class="legend-swatch" style="background:#fbbf24"></span>Export scheduled</span>
      <span><span class="legend-swatch" style="background:#60a5fa"></span>Import scheduled</span>
      <span><span class="legend-swatch" style="background:rgba(255,255,255,.10)"></span>Usable room (up to binding)</span>
      ${isNY ? '<span><span class="legend-swatch" style="background:repeating-linear-gradient(45deg,rgba(167,139,250,.32) 0 4px,rgba(124,92,235,.12) 4px 8px)"></span>Blocked by tighter limit</span>' : ''}
      <span><span class="legend-swatch" style="background:#f0f6fc"></span>Binding tick (bright)</span>
      ${isNY ? '<span><span class="legend-swatch" style="background:#6e7681"></span>Other limit tick</span>' : ''}
    </div>`;
  container.innerHTML = `<div class="zone">
    <div class="zone-head"><div class="zone-name">${z.name}</div></div>
    ${legend}
    ${renderDirectionTable(k, 'export', z.export, isNY)}
    ${renderDirectionTable(k, 'import', z.import, isNY)}
  </div>`;
}

function render(){
  const k = document.getElementById('corridorSelect').value;
  renderSelectedZone(k);
}

renderMeta();
renderSystemTable();
document.getElementById('corridorSelect').addEventListener('change', render);
render();
</script>
</body>
</html>
"""


def write_html_dashboard(payload, out_path):
    html = (HTML_TEMPLATE
            .replace("__WINDOW__", payload.get("title_window", ""))
            .replace("__PAYLOAD__", json.dumps(payload, default=str)))
    with open(out_path, "w") as f:
        f.write(html)
    return out_path


# ── Hour-window planning (past + forward, cross-midnight) ──────────
def get_hour_window(now, past, forward):
    """
    Return list of (date_str, he) tuples for a window covering `past` HEs
    before the current HE plus `forward` HEs starting from the current HE.

    Current HE = the hour ending at the *next* wall-clock hour boundary.
    E.g. clock 20:15 EST → current HE is 21 (covers 20:00-21:00).

    Rolls across midnight in both directions.
    """
    current_he = now.hour + 1  # HE numbering convention
    today = now
    # Build a flat list of (date, he) where we walk from -past to +forward-1
    out = []
    for offset in range(-past, forward):
        he = current_he + offset
        day_shift = 0
        # Handle wrap-around
        while he < 1:
            he += 24
            day_shift -= 1
        while he > 24:
            he -= 24
            day_shift += 1
        date_str = (today + timedelta(days=day_shift)).strftime("%Y%m%d")
        out.append((date_str, he))
    return out


def window_title(hours_spec):
    """Format a title like 'HE 22 to HE 4' spanning the hours_spec window."""
    if not hours_spec:
        return ""
    start = f"HE {hours_spec[0][1]}"
    end = f"HE {hours_spec[-1][1]}"
    return f"{start} to {end}"


# ── Build forward-row rows (replaces build_zone_data + payload assembly) ──
def build_forward_rows(hours_spec, limits_by_date, adeq_by_date, nyiso_by_date):
    """
    For each zone and direction, produce a list of row dicts aligned with hours_spec.
    limits_by_date = {date_str: parsed_predisp_limits_dict}
    adeq_by_date  = {date_str: parsed_adequacy3_dict}
    nyiso_by_date = {date_str: (nyiso_exp_rows, nyiso_imp_rows)}
    """
    def pq_limit(limits, direction, he):
        """Sum all PQ sub-zone limits for the given direction & hour.
        IESO convention: N = export ("to"), X = import ("from")."""
        if limits is None:
            return None
        suffix = "X" if direction == "import" else "N"
        total = 0
        any_data = False
        for sub in config.QUEBEC_SUBZONES:
            v = limits["zones"].get(f"PQ{sub}{suffix}", {}).get(he)
            if v is not None:
                total += abs(v)
                any_data = True
        return total if any_data else None

    def iso_limit_for(zkey, direction, date_str, he):
        """Look up the IESO Predisp scheduling limit for this zone/direction/date/HE."""
        limits = limits_by_date.get(date_str)
        if limits is None:
            return None
        # Only serve if the file's declared delivery date matches
        if limits.get("delivery_date", "").replace("-", "") != date_str:
            return None
        spec = config.IESO_ZONES[zkey]
        if zkey == "PQ":
            return pq_limit(limits, direction, he)
        code = spec[direction]
        if not code:
            return None
        v = limits["zones"].get(code, {}).get(he)
        return abs(v) if v is not None else None

    zones = {}
    for zkey in config.DISPLAY_ORDER:
        spec = config.IESO_ZONES[zkey]
        zones[zkey] = {"name": spec["name"], "export": [], "import": []}
        for direction in ("export", "import"):
            for (date_str, he) in hours_spec:
                iso_lim = iso_limit_for(zkey, direction, date_str, he)

                # Scheduled from that date's Adequacy3.
                # Sign convention (Ontario-centric):
                #   exports are negative (off-take from ON)
                #   imports are positive (flow into ON)
                # IESO's Adequacy3 already stores exports as negative; we keep it.
                adeq = adeq_by_date.get(date_str)
                sched = None       # signed value for display
                sched_mag = None   # magnitude for calculations
                if adeq:
                    adeq_name = spec["adeq"]
                    key_group = "zonal_imports" if direction == "import" else "zonal_exports"
                    bucket = adeq[key_group].get(adeq_name, {})
                    raw = bucket.get("scheduled", {}).get(he)
                    if raw is not None:
                        sched = raw             # keep IESO's sign
                        sched_mag = abs(raw)

                # NYISO only for NY
                ny_ttc = ny_real = ny_atc = None
                if zkey == "NY":
                    ny_rows_exp, ny_rows_imp = nyiso_by_date.get(date_str, ([], []))
                    ny_rows = ny_rows_exp if direction == "export" else ny_rows_imp
                    nr = _nyiso_for_he(ny_rows, he)
                    if nr:
                        ny_ttc = nr["dam_ttc"]
                        ny_atc = nr["dam_atc"]
                        if ny_ttc is not None:
                            ny_real = max(0, ny_ttc - config.NYISO_SECURITY_BUFFER_MW)

                binding = _binding_limit(iso_lim, {"dam_ttc": ny_ttc}) if ny_ttc is not None else iso_lim
                remaining = max(0, binding - sched_mag) if (binding is not None and sched_mag is not None) else None
                util = int(round(sched_mag / binding * 100)) if (binding and binding > 0 and sched_mag is not None) else None
                icon, status = classify(remaining, binding)
                zones[zkey][direction].append({
                    "date": date_str, "he": he,
                    "iso_limit": iso_lim,
                    "ny_ttc": ny_ttc, "ny_real": ny_real, "ny_atc": ny_atc,
                    "binding": binding, "scheduled": sched,
                    "remaining": remaining, "utilization": util, "status": status,
                })
    return zones


def build_alerts_from_rows(zones):
    """Build short alert bullets, iterating over the pre-computed rows."""
    alerts = []
    for zkey in config.DISPLAY_ORDER:
        z = zones[zkey]
        for direction in ("export", "import"):
            crit_h, warn_h = [], []
            for r in z[direction]:
                if r["status"] == "critical":
                    crit_h.append(r["he"])
                elif r["status"] == "warning":
                    warn_h.append(r["he"])
            dir_label = f"ON → {zkey}" if direction == "export" else f"{zkey} → ON"
            if crit_h:
                alerts.append(
                    f"🔴 {dir_label}: CRITICAL HE {crit_h[0]}"
                    + (f"–HE {crit_h[-1]}" if len(crit_h) > 1 else "")
                )
            elif warn_h:
                alerts.append(
                    f"⚠️ {dir_label}: tight HE {warn_h[0]}"
                    + (f"–HE {warn_h[-1]}" if len(warn_h) > 1 else "")
                )
    return alerts


# ── Main ────────────────────────────────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description="Ontario Intertie Congestion Scanner")
    ap.add_argument("--notify", action="store_true",
                    help="Send a macOS notification with the alerts summary")
    ap.add_argument("--hours", type=int, default=config.FORECAST_HOURS,
                    help=f"Forward HEs to display (default {config.FORECAST_HOURS})")
    ap.add_argument("--past", type=int, default=config.PAST_HOURS,
                    help=f"Past HEs to display (default {config.PAST_HOURS})")
    ap.add_argument("--html", metavar="PATH", default=None,
                    help="Also write a self-contained HTML dashboard to PATH")
    args = ap.parse_args()

    now = get_now_et()
    hours_spec = get_hour_window(now, past=args.past, forward=args.hours)
    unique_dates = sorted({d for d, _ in hours_spec})

    # ── Fetch ──
    print("── Fetching ──")
    limits_by_date = {}
    limits_filenames = {}
    for d in unique_dates:
        try:
            print(f"  IESO PredispIntertieSchedLimits ({d})… ", end="", flush=True)
            b, fname = fetchers.fetch_predisp_limits_for_date(d)
            if b is None:
                print("not yet published")
                continue
            limits_by_date[d] = parsers.parse_predisp_limits(b)
            limits_filenames[d] = fname
            print(f"ok ({fname})")
        except Exception as e:
            print(f"WARN: {e}")
    if not limits_by_date:
        # Fall back to the always-current file
        try:
            print("  IESO PredispIntertieSchedLimits (fallback current)… ", end="", flush=True)
            b = fetchers.fetch_predisp_intertie_limits()
            parsed = parsers.parse_predisp_limits(b)
            key = parsed["delivery_date"].replace("-", "")
            limits_by_date[key] = parsed
            limits_filenames[key] = "PUB_PredispIntertieSchedLimits.xml"
            print(f"ok (delivery {parsed['delivery_date']})")
        except Exception as e:
            print(f"FAIL: {e}", file=sys.stderr)

    adeq_by_date = {}
    adeq_filenames = {}
    for d in unique_dates:
        try:
            print(f"  IESO Adequacy3 ({d})…      ", end="", flush=True)
            b, fname = fetchers.fetch_adequacy3_for_date(d)
            adeq_by_date[d] = parsers.parse_adequacy3(b)
            adeq_filenames[d] = fname
            print(f"ok ({fname})")
        except Exception as e:
            print(f"WARN: {e}")

    nyiso_by_date = {}
    for d in unique_dates:
        try:
            print(f"  NYISO ATC/TTC ({d})…       ", end="", flush=True)
            b = fetchers.fetch_nyiso_atc_ttc(d)
            ny_exp = parsers.parse_nyiso_atc(b, config.NYISO_IMO_NY_INTERFACE)
            ny_imp = parsers.parse_nyiso_atc(b, config.NYISO_NY_IMO_INTERFACE)
            nyiso_by_date[d] = (ny_exp, ny_imp)
            print("ok")
        except Exception as e:
            print(f"WARN: {e}")
            nyiso_by_date[d] = ([], [])

    # ── Build analysis ──
    zones_rows = build_forward_rows(hours_spec, limits_by_date, adeq_by_date, nyiso_by_date)
    alerts = build_alerts_from_rows(zones_rows)

    # ── Terminal report (compact) ──
    title_window = window_title(hours_spec)
    # Pick a representative "limits report" status line from whichever date has data
    lim_desc_parts = []
    for d in sorted(limits_by_date.keys()):
        lim_desc_parts.append(f"{limits_by_date[d]['delivery_date']}@{limits_by_date[d]['created_at'][-8:]}")
    lim_desc = ", ".join(lim_desc_parts) if lim_desc_parts else "none"
    print()
    print("═" * 62)
    print(f"  IESO Congestion Picture  {title_window}")
    print(f"  Generated: {now.strftime('%Y-%m-%d %H:%M')} EST")
    print(f"  Predisp limits loaded for: {lim_desc}")
    print("═" * 62)
    if alerts:
        print("ALERTS")
        for a in alerts:
            print(f"  {a}")
    else:
        print(f"  ✅ All interties comfortable for {title_window}")
    print("═" * 62)

    # ── Notification ──
    notif_title = f"IESO Congestion Picture {title_window}"
    notif_body = now.strftime("Updated %H:%M EST")
    if args.notify:
        notify(notif_title, notif_body)

    # ── HTML dashboard ──
    if args.html:
        payload = {
            "generated_at": now.isoformat(),
            "title_window": title_window,
            "notif_title": notif_title,
            "notif_body": notif_body,
            "hours_spec": hours_spec,  # list of [date_str, he]
            "limits_by_date": {
                d: {
                    "created_at": limits_by_date[d]["created_at"],
                    "delivery_date": limits_by_date[d]["delivery_date"],
                } for d in limits_by_date
            },
            "predisp_urls": {
                d: f"{config.IESO_ADEQUACY3_BASE_URL.rsplit('/',1)[0]}/PredispIntertieSchedLimits/{limits_filenames[d]}"
                for d in limits_filenames
            },
            "nyiso_urls": {d: config.NYISO_ATC_TTC_URL.format(date=d) for d in unique_dates},
            "adeq_urls": {
                d: f"{config.IESO_ADEQUACY3_BASE_URL}/{adeq_filenames[d]}"
                for d in adeq_filenames
            },
            "ontario_demand": [
                (adeq_by_date[d]["ontario_demand"].get(he) if d in adeq_by_date else None)
                for (d, he) in hours_spec
            ],
            "excess": [
                (adeq_by_date[d]["excess"].get(he) if d in adeq_by_date else None)
                for (d, he) in hours_spec
            ],
            "zones": zones_rows,
            "alerts": alerts,
            "has_alerts": bool(alerts),
        }
        out = write_html_dashboard(payload, args.html)
        print(f"\nHTML dashboard: {out}")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrupted.")
        sys.exit(130)
