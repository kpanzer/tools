#!/usr/bin/env python3
"""
IESO Predispatch Hourly Intertie LMP Watcher
=============================================
Polls the IESO public reports directory every 30 seconds starting at :05 past
each hour, and sends a macOS notification when a new predispatch file drops.

Instead of guessing the version number, it scrapes the directory listing and
detects when the latest file for today changes.

Usage:
    python3 ieso_predispatch_watcher.py

Press Ctrl+C to stop.
"""

import re
import subprocess
import time
import urllib.request
import urllib.error
from datetime import datetime, timezone, timedelta

# ── Config ──────────────────────────────────────────────────────────
POLL_INTERVAL = 30          # seconds between checks
POLL_START_MINUTE = 5       # start polling at :05 past the hour
POLL_WARN_MINUTE = 20       # send warning notification if not found by :20
POLL_TIMEOUT_MINUTE = 25    # stop polling at :25 if still not found
BASE_URL = "https://reports-public.ieso.ca/public/PredispHourlyIntertieLMP"
EST = timezone(timedelta(hours=-5))

# ── macOS Notification ──────────────────────────────────────────────
def notify(title, message, sound="Glass"):
    subprocess.run([
        "terminal-notifier",
        "-title", title,
        "-message", message,
        "-sound", sound,
        "-group", "ieso-watcher",
    ], capture_output=True)


def get_est_now():
    return datetime.now(EST)


def get_latest_version(est_dt):
    """
    Scrape the IESO directory listing and return the highest version number
    found for today's date. Returns None if no files found.
    """
    date_str = est_dt.strftime("%Y%m%d")
    url = f"{BASE_URL}/"
    try:
        req = urllib.request.Request(url)
        req.add_header("User-Agent", "IESO-Watcher/1.0")
        resp = urllib.request.urlopen(req, timeout=10)
        html = resp.read().decode()
        pattern = rf"PUB_PredispHourlyIntertieLMP_{date_str}_v(\d+)\.xml"
        versions = [int(v) for v in re.findall(pattern, html)]
        if not versions:
            return None
        return max(versions)
    except (urllib.error.HTTPError, urllib.error.URLError, OSError):
        return None


def build_url(est_dt, version):
    date_str = est_dt.strftime("%Y%m%d")
    return f"{BASE_URL}/PUB_PredispHourlyIntertieLMP_{date_str}_v{version}.xml"


def main():
    print("=" * 60)
    print("  IESO Predispatch Intertie LMP Watcher")
    print("=" * 60)
    print(f"  Polling every {POLL_INTERVAL}s from :{POLL_START_MINUTE:02d} to :{POLL_TIMEOUT_MINUTE:02d} each hour")
    print(f"  Warning notification at :{POLL_WARN_MINUTE:02d} if no new file found")
    print(f"  Report: PredispHourlyIntertieLMP")
    print("=" * 60)
    print()

    last_notified_version = None   # highest version we already notified for
    warned_hour = None             # hour we already sent a delay warning for

    while True:
        now = get_est_now()
        minute = now.minute
        hour_key = now.strftime("%Y%m%d_%H")

        if POLL_START_MINUTE <= minute < POLL_TIMEOUT_MINUTE:
            print(f"[{now.strftime('%H:%M:%S')} EST]  Checking directory for latest version...", end="", flush=True)

            latest = get_latest_version(now)

            if latest is None:
                print(" no files found.", flush=True)
                time.sleep(POLL_INTERVAL)
                continue

            version_key = f"{now.strftime('%Y%m%d')}_v{latest}"

            if version_key == last_notified_version:
                print(f" v{latest} already reported. Waiting for next update...", end="\r", flush=True)
                time.sleep(POLL_INTERVAL)
                continue

            elapsed = minute - POLL_START_MINUTE
            url = build_url(now, latest)
            print(f" NEW: v{latest} (+{elapsed}min after :{POLL_START_MINUTE:02d})")
            print(f"  → {url}")
            print()

            notify(
                "⚠️ IESO Reports",
                f"{now.strftime('%b %d')} {now.strftime('%H:%M')} EST - AVAILABLE"
            )

            last_notified_version = version_key
            print(f"  ✅ Notification sent at {now.strftime('%H:%M:%S')} EST")
            print(f"  ⏳ Next check at :{POLL_START_MINUTE:02d} past the next hour")
            print()

        else:
            if minute < POLL_START_MINUTE:
                wait_min = POLL_START_MINUTE - minute
                status = f"[{now.strftime('%H:%M:%S')} EST]  Waiting {wait_min}min until :{POLL_START_MINUTE:02d}..."
            else:
                next_hour = (now.hour + 1) % 24
                status = f"[{now.strftime('%H:%M:%S')} EST]  Polling window closed. Next check at :{POLL_START_MINUTE:02d} past {next_hour:02d}:00 EST"

            print(status, end="\r", flush=True)

        # Warning at :20 if no new version found this hour
        if minute >= POLL_WARN_MINUTE and hour_key != warned_hour:
            latest = get_latest_version(now)
            version_key = f"{now.strftime('%Y%m%d')}_v{latest}" if latest else None
            if version_key != last_notified_version:
                print(f"\n  ⚠️  No new file found by :{POLL_WARN_MINUTE:02d}. May be delayed.\n")
                notify(
                    "⚠️ IESO Reports",
                    f"{now.strftime('%b %d')} {now.strftime('%H:%M')} EST - DELAYED",
                    sound="Basso"
                )
                warned_hour = hour_key

        time.sleep(10)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nWatcher stopped. Goodbye!")
