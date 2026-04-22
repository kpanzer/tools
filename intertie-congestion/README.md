# Ontario Intertie Congestion Scanner

A forward-looking view of scheduling congestion on every Ontario intertie
(NY, MI, MN, MB, PQ). Built for Command Power trading operations.

## What it does

Every 5 minutes the tool pulls three public reports:

1. **IESO Pre-Dispatch Intertie Scheduling Limits** — hour-by-hour scheduling
   limits for each Ontario intertie zone (ON PD limits).
2. **IESO Adequacy3** — forecast Ontario demand, excess capacity, and zonal
   imports/exports scheduled through pre-dispatch.
3. **NYISO ATC/TTC** — Total / Available Transfer Capability for the NYISO-IMO
   interface, with the 300 MW security buffer applied to get the real binding
   capability.

It cross-references the three, computes a **binding limit** per hour and
direction, and writes a self-contained HTML dashboard showing:

- **2 past hours** (already cleared) + **8 forward hours** (current HE + 7 ahead)
- Per-hour scheduled flow, remaining room, utilization %, and status
  (OK / WARN / CRIT)
- A visual capacity gauge per row showing scheduled vs. usable room
  vs. capacity blocked by the tighter side (e.g., NY TTC−300 vs. IESO)
- Ontario system demand and excess-capacity strip
- Corridor dropdown to switch between NY / MI / MN / MB / PQ

Hours are in **EST** (fixed UTC-5, no DST) to match IESO market convention.

## Project layout

| File | Purpose |
|---|---|
| `config.py` | URLs, zone-code mappings (IESO's `N` = export-to-zone, `X` = import-from-zone), thresholds |
| `fetchers.py` | URL fetchers: latest versioned IESO XML files + NYISO HTML |
| `parsers.py` | `xml.etree` for IESO XMLs, per-row regex for NYISO HTML |
| `scan.py` | Orchestrator: fetch → parse → analyze → write dashboard + (optional) notify |

## Running manually

```bash
cd intertie-congestion
python3 scan.py --html dashboard.html
```

Flags:
- `--html PATH` write the self-contained HTML dashboard
- `--notify` send a macOS notification (requires `terminal-notifier`)
- `--past N` number of past HEs (default 2)
- `--hours N` number of forward HEs including current (default 8)

## Running automatically (macOS LaunchAgent)

A LaunchAgent at `~/Library/LaunchAgents/com.commandpower.intertie-congestion.plist`
runs the scan every 5 minutes and overwrites `dashboard.html`. The HTML has a
5-minute meta-refresh so an open browser tab stays current.

## Data source conventions

- IESO **PredispIntertieSchedLimits** zone codes use the IESO convention
  (per their XSL stylesheet):
  - `<ZONE>N` (e.g. `NYSIN`) → "to &lt;zone&gt;" = **EXPORT from Ontario**
  - `<ZONE>X` (e.g. `NYSIX`) → "from &lt;zone&gt;" = **IMPORT to Ontario**
  - Values are signed from Ontario's system perspective: exports negative,
    imports positive. We take `abs()` to get capacity magnitudes.
- IESO **Adequacy3** `ZonalExport` schedules are stored as negative MW
  (Ontario load convention). Displayed signed on the dashboard to keep
  the convention obvious (yellow = export, blue = import).
- NYISO **TTC listed** is overstated by 300 MW for security; the real
  usable TTC is `listed − 300`. The binding limit at NY is
  `min(IESO PD limit, NYISO TTC − 300)`.

## Tomorrow's data

IESO publishes tomorrow's PredispIntertieSchedLimits around 20:00 EST.
When the scanner sees it, tomorrow's HE limits populate automatically on
the next tick. Until then, tomorrow-dated rows show `—` for the IESO
column and fall back to NYISO TTC as the binder (for the NY corridor).

## License / audience

Internal Command Power tooling. Public IESO / NYISO data only — no
market-sensitive information embedded.
