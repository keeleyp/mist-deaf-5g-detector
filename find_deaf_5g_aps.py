#!/usr/bin/env python3
# find_deaf_5g_aps.py
#
# Goes through every site in the org and looks for APs whose 5GHz radio
# has gone "deaf-in-transmit": the AP still hears its neighbours on 5GHz,
# but no neighbour hears it back.
#
# Signature:
#   - AP reports neighbours on 5GHz, some of them strong
#   - AP does not appear in ANY other AP's 5GHz neighbour list
#   - (usually) AP is still heard normally on 2.4GHz
#
# Output:
#   - summary printed to screen
#   - one row per suspect AP appended to a CSV (so repeated runs build a history)
#   - an Excel report of the failed APs for this run, named with the site(s) and time
#
# Usage:
#   Edit find_deaf_5g_aps.ini (same folder as this script), then:
#   python3 find_deaf_5g_aps.py
#   python3 find_deaf_5g_aps.py other_org.ini     (optional: use a different .ini)

import os
import sys
import csv
import time
import datetime
import configparser
import re
import json
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ---------------- read settings from the .ini file ----------------

if len(sys.argv) > 1:
    INI_FILE = sys.argv[1]
else:
    INI_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "find_deaf_5g_aps.ini")

if not os.path.exists(INI_FILE):
    print("Can't find settings file: " + INI_FILE)
    sys.exit(1)

config = configparser.ConfigParser()
config.read(INI_FILE)

API_HOST = config.get("mist", "api_host", fallback="https://api.eu.mist.com").rstrip("/")
API_TOKEN = config.get("mist", "api_token", fallback="").strip()
ORG_ID = config.get("mist", "org_id", fallback="").strip()

MIN_NEIGHBOURS_HEARD = config.getint("detection", "min_neighbours_heard", fallback=2)
STRONGEST_RSSI_THRESHOLD = config.getint("detection", "strongest_rssi_threshold", fallback=-75)
# Ignore APs up for less than this many days - RRM won't have built a neighbour view yet
MIN_UPTIME_DAYS = config.getfloat("detection", "min_uptime_days", fallback=1)
# Ignore APs that have clients on 5GHz - the radio must be working
IGNORE_IF_5G_CLIENTS = config.getboolean("detection", "ignore_if_5g_clients", fallback=True)
# Ignore APs whose 5GHz power is 0 - no SSIDs configured on that radio
IGNORE_IF_5G_POWER_ZERO = config.getboolean("detection", "ignore_if_5g_power_zero", fallback=True)
# Ignore APs whose 5GHz channel or power is unknown (no 5GHz radio stats reported)
IGNORE_IF_5G_UNKNOWN = config.getboolean("detection", "ignore_if_5g_unknown", fallback=True)
# Ignore APs not heard by any other AP on 2.4GHz either - probably just isolated
IGNORE_IF_ISOLATED = config.getboolean("detection", "ignore_if_isolated", fallback=True)

SITE_NAME_FILTER = config.get("run", "site_name_filter", fallback="").strip()
PAUSE_BETWEEN_CALLS = config.getfloat("run", "pause_between_calls", fallback=0.3)
# {org} in the name is replaced with the org name, so each org keeps its own history file
CSV_FILE = config.get("run", "csv_file", fallback="deaf_5g_aps_{org}.csv").strip()
REPORT_FOLDER = config.get("run", "report_folder", fallback=".").strip() or "."

# Offer to reset (disable, wait, re-enable) the radio on the failed APs at the end of the run
OFFER_RADIO_RESET = config.getboolean("reset", "offer_radio_reset", fallback=True)
# Which radios to reset: 5 (default), or e.g. 24,5 for both 2.4GHz and 5GHz
RESET_BANDS = ["band_" + b.strip() for b in config.get("reset", "reset_bands", fallback="5").split(",") if b.strip() in ("24", "5", "6")]
if len(RESET_BANDS) == 0:
    RESET_BANDS = ["band_5"]
RESET_WAIT_SECONDS = config.getfloat("reset", "reset_wait_seconds", fallback=5)

# -----------------------------------------------------------------

if API_TOKEN == "" or API_TOKEN.startswith("PASTE"):
    print("Put your Mist API token in " + INI_FILE + " ([mist] api_token)")
    sys.exit(1)
if ORG_ID == "" or ORG_ID.startswith("PASTE"):
    print("Put the org ID in " + INI_FILE + " ([mist] org_id)")
    sys.exit(1)

# ---------- HTTPS certificate checking ----------
# On corporate networks a proxy (Zscaler, Netskope, Palo Alto etc.) often re-signs
# HTTPS traffic with the company's own root certificate. That certificate is in the
# Windows/macOS certificate store, but not in the bundle Python uses by default,
# which gives "CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate".
#   use_system_certs = true -> trust the operating system's certificate store (needs truststore)
#   ca_bundle = <path>      -> or point at the company root certificate (.pem) instead
USE_SYSTEM_CERTS = config.getboolean("mist", "use_system_certs", fallback=True)
CA_BUNDLE = config.get("mist", "ca_bundle", fallback="").strip()

if USE_SYSTEM_CERTS and CA_BUNDLE == "":
    try:
        import truststore
        truststore.inject_into_ssl()
    except ImportError:
        print("Note: 'truststore' isn't installed, so the OS certificate store isn't used.")
        print("      If you get CERTIFICATE_VERIFY_FAILED, run: pip install -r requirements.txt")
        print("")

session = requests.Session()
session.headers.update({"Authorization": "Token " + API_TOKEN})
if CA_BUNDLE != "":
    if not os.path.exists(CA_BUNDLE):
        print("ca_bundle file not found: " + CA_BUNDLE)
        sys.exit(1)
    session.verify = CA_BUNDLE
    # requests lets these environment variables override session.verify, so set them too
    os.environ["REQUESTS_CA_BUNDLE"] = CA_BUNDLE
    os.environ["CURL_CA_BUNDLE"] = CA_BUNDLE

# Quick connection test, so a certificate/proxy problem gives a clear message
try:
    session.get(API_HOST + "/api/v1/self", timeout=30)
except requests.exceptions.SSLError as e:
    print("HTTPS certificate check failed connecting to " + API_HOST)
    print("")
    print("This usually means a company proxy is inspecting HTTPS traffic. To fix it, either:")
    print("  1. pip install -r requirements.txt   (installs truststore, so Windows/macOS")
    print("     certificates are trusted - leave use_system_certs = true in the .ini), or")
    print("  2. get your company's root certificate as a .pem file and set")
    print("     ca_bundle = C:\\path\\to\\company-root.pem   under [mist] in the .ini")
    print("")
    print("Details: " + str(e))
    sys.exit(1)
except requests.exceptions.ConnectionError as e:
    print("Couldn't connect to " + API_HOST + " - check the api_host setting and network/proxy access.")
    print("Details: " + str(e))
    sys.exit(1)

run_now = datetime.datetime.now(datetime.timezone.utc)
run_time = run_now.strftime("%Y-%m-%d %H:%M:%S UTC")
print("Run started " + run_time)
# ---------- look up the org name (used in the CSV and Excel filenames) ----------

org_name = ORG_ID
while True:
    resp = session.get(API_HOST + "/api/v1/orgs/" + ORG_ID)
    if resp.status_code == 429:
        print("Rate limited getting org, waiting 60s...")
        time.sleep(60)
        continue
    if resp.status_code == 200 and isinstance(resp.json(), dict):
        org_name = resp.json().get("name", ORG_ID) or ORG_ID
    else:
        print("Couldn't look up org name (HTTP " + str(resp.status_code) + ") - using org ID in filenames")
    break

# Safe version for filenames: letters, digits, _ and - only
org_file_part = re.sub(r"[^A-Za-z0-9_-]+", "-", org_name).strip("-") or "org"
CSV_FILE = CSV_FILE.replace("{org}", org_file_part)

print("Org " + org_name + " (" + ORG_ID + ") on " + API_HOST)
print("")

# ---------- 1. get all sites in the org (paged) ----------

sites = []
page = 1
while True:
    url = API_HOST + "/api/v1/orgs/" + ORG_ID + "/sites"
    resp = session.get(url, params={"limit": 1000, "page": page})
    if resp.status_code == 429:
        print("Rate limited getting sites, waiting 60s...")
        time.sleep(60)
        continue
    resp.raise_for_status()
    batch = resp.json()
    sites.extend(batch)
    if len(batch) < 1000:
        break
    page = page + 1
    time.sleep(PAUSE_BETWEEN_CALLS)

if SITE_NAME_FILTER != "":
    sites = [s for s in sites if SITE_NAME_FILTER.lower() in s.get("name", "").lower()]

print("Sites to check: " + str(len(sites)))
print("")

# ---------- prepare the CSV ----------

csv_exists = os.path.exists(CSV_FILE)
csv_handle = open(CSV_FILE, "a", newline="")
writer = csv.writer(csv_handle)
if not csv_exists:
    writer.writerow([
        "run_time", "site_name", "site_id", "ap_name", "ap_mac", "model", "firmware",
        "uptime_days", "status", "ch_5g", "power_5g", "clients_5g", "clients_24g",
        "neighbours_heard_5g", "strongest_heard_rssi_5g", "strongest_heard_mac",
        "heard_on_24g", "heard_by_count_24g",
    ])

report_rows = []
failed_aps = []
total_suspects = 0
total_ignored_uptime = 0
total_ignored_clients = 0
total_ignored_power = 0
total_ignored_unknown = 0
total_ignored_isolated = 0
sites_with_suspects = 0
sites_done = 0
sites_skipped = 0

# ---------- 2. loop through every site ----------

for site in sites:
    site_id = site["id"]
    site_name = site.get("name", site_id)
    sites_done = sites_done + 1

    # ----- 2a. 5GHz neighbours (paged) -----
    results_5 = []
    page = 1
    while True:
        url = API_HOST + "/api/v1/sites/" + site_id + "/rrm/neighbors/band/5"
        resp = session.get(url, params={"limit": 100, "page": page})
        if resp.status_code == 429:
            print("  rate limited, waiting 60s...")
            time.sleep(60)
            continue
        if resp.status_code != 200:
            print("  [" + site_name + "] 5GHz neighbour call failed: HTTP " + str(resp.status_code))
            break
        body = resp.json()
        results_5.extend(body.get("results", []))
        total = body.get("total", 0)
        time.sleep(PAUSE_BETWEEN_CALLS)
        if len(results_5) >= total or len(body.get("results", [])) == 0:
            break
        page = page + 1

    # Sites with 0 or 1 reporting APs can't be judged
    if len(results_5) < 2:
        sites_skipped = sites_skipped + 1
        continue

    # Every AP MAC that some OTHER AP hears on 5GHz
    heard_by_others_5 = set()
    for entry in results_5:
        for nb in entry.get("neighbors", []):
            if nb["mac"] != entry["mac"]:
                heard_by_others_5.add(nb["mac"])

    # Find suspects: hears strong neighbours, but nobody hears it
    suspects = []
    for entry in results_5:
        ap_mac = entry["mac"]
        neighbours = entry.get("neighbors", [])
        if len(neighbours) < MIN_NEIGHBOURS_HEARD:
            continue
        if ap_mac in heard_by_others_5:
            continue
        strongest = None
        for nb in neighbours:
            if strongest is None or nb["rssi"] > strongest["rssi"]:
                strongest = nb
        if strongest["rssi"] < STRONGEST_RSSI_THRESHOLD:
            continue
        suspects.append({
            "mac": ap_mac,
            "count": len(neighbours),
            "strongest_rssi": strongest["rssi"],
            "strongest_mac": strongest["mac"],
        })

    if len(suspects) == 0:
        print("[" + str(sites_done) + "/" + str(len(sites)) + "] " + site_name + ": OK (" + str(len(results_5)) + " APs)")
        continue

    # ----- 2b. 2.4GHz neighbours, to see if suspects are still heard there -----
    results_24 = []
    page = 1
    while True:
        url = API_HOST + "/api/v1/sites/" + site_id + "/rrm/neighbors/band/24"
        resp = session.get(url, params={"limit": 100, "page": page})
        if resp.status_code == 429:
            print("  rate limited, waiting 60s...")
            time.sleep(60)
            continue
        if resp.status_code != 200:
            print("  [" + site_name + "] 2.4GHz neighbour call failed: HTTP " + str(resp.status_code))
            break
        body = resp.json()
        results_24.extend(body.get("results", []))
        total = body.get("total", 0)
        time.sleep(PAUSE_BETWEEN_CALLS)
        if len(results_24) >= total or len(body.get("results", [])) == 0:
            break
        page = page + 1

    heard_count_24 = {}
    for entry in results_24:
        for nb in entry.get("neighbors", []):
            if nb["mac"] != entry["mac"]:
                heard_count_24[nb["mac"]] = heard_count_24.get(nb["mac"], 0) + 1

    # ----- 2c. AP stats for names, model, firmware, clients -----
    ap_stats = {}
    while True:
        url = API_HOST + "/api/v1/sites/" + site_id + "/stats/devices"
        resp = session.get(url, params={"type": "ap", "limit": 1000})
        if resp.status_code == 429:
            print("  rate limited, waiting 60s...")
            time.sleep(60)
            continue
        if resp.status_code == 200:
            for dev in resp.json():
                ap_stats[dev.get("mac", "")] = dev
        else:
            print("  [" + site_name + "] device stats call failed: HTTP " + str(resp.status_code))
        time.sleep(PAUSE_BETWEEN_CALLS)
        break

    # ----- 2d. drop APs that can't be judged or are clearly working -----
    #   - up for less than MIN_UPTIME_DAYS: RRM won't have run yet
    #   - has clients on 5GHz: the radio must be working
    #   - 5GHz power is 0: no SSIDs configured on that radio
    #   - 5GHz channel or power unknown: no 5GHz radio stats to judge it on
    #   - not heard on 2.4GHz either: probably an isolated AP, not a 5GHz fault
    #     (only applied if the 2.4GHz neighbour call actually returned data)
    # (if stats are missing for an AP we keep it, rather than hide a real case)
    kept = []
    ignored_uptime_here = 0
    ignored_clients_here = 0
    ignored_power_here = 0
    ignored_unknown_here = 0
    ignored_isolated_here = 0
    for s in suspects:
        dev = ap_stats.get(s["mac"], {})
        if "uptime" in dev and (dev.get("uptime") or 0) < MIN_UPTIME_DAYS * 86400:
            ignored_uptime_here = ignored_uptime_here + 1
            continue
        radio = dev.get("radio_stat", {}) or {}
        r5 = radio.get("band_5", {}) or {}
        if IGNORE_IF_5G_CLIENTS and (r5.get("num_clients") or 0) > 0:
            ignored_clients_here = ignored_clients_here + 1
            continue
        if IGNORE_IF_5G_POWER_ZERO and "power" in r5 and (r5.get("power") or 0) == 0:
            ignored_power_here = ignored_power_here + 1
            continue
        if IGNORE_IF_5G_UNKNOWN and (r5.get("channel") in (None, "", 0) or r5.get("power") in (None, "")):
            ignored_unknown_here = ignored_unknown_here + 1
            continue
        if IGNORE_IF_ISOLATED and len(results_24) > 0 and heard_count_24.get(s["mac"], 0) == 0:
            ignored_isolated_here = ignored_isolated_here + 1
            continue
        kept.append(s)
    suspects = kept
    total_ignored_uptime = total_ignored_uptime + ignored_uptime_here
    total_ignored_clients = total_ignored_clients + ignored_clients_here
    total_ignored_power = total_ignored_power + ignored_power_here
    total_ignored_unknown = total_ignored_unknown + ignored_unknown_here
    total_ignored_isolated = total_ignored_isolated + ignored_isolated_here

    ignored_note = ""
    if ignored_uptime_here > 0:
        ignored_note = ignored_note + ", " + str(ignored_uptime_here) + " ignored - uptime < " + str(MIN_UPTIME_DAYS) + " day"
    if ignored_clients_here > 0:
        ignored_note = ignored_note + ", " + str(ignored_clients_here) + " ignored - has 5GHz clients"
    if ignored_power_here > 0:
        ignored_note = ignored_note + ", " + str(ignored_power_here) + " ignored - 5GHz power 0 (no SSIDs)"
    if ignored_unknown_here > 0:
        ignored_note = ignored_note + ", " + str(ignored_unknown_here) + " ignored - 5GHz channel/power unknown"
    if ignored_isolated_here > 0:
        ignored_note = ignored_note + ", " + str(ignored_isolated_here) + " ignored - not heard on 2.4GHz either (isolated)"

    if len(suspects) == 0:
        print("[" + str(sites_done) + "/" + str(len(sites)) + "] " + site_name + ": OK (" + str(len(results_5))
              + " APs" + ignored_note + ")")
        continue

    sites_with_suspects = sites_with_suspects + 1

    # ----- 2e. report -----
    line = "[" + str(sites_done) + "/" + str(len(sites)) + "] " + site_name + ": " + str(len(suspects)) + " SUSPECT AP(s)"
    if ignored_note != "":
        line = line + " (" + ignored_note[2:] + ")"
    print(line)

    for s in suspects:
        total_suspects = total_suspects + 1
        dev = ap_stats.get(s["mac"], {})
        radio = dev.get("radio_stat", {}) or {}
        r5 = radio.get("band_5", {}) or {}
        r24 = radio.get("band_24", {}) or {}
        uptime = dev.get("uptime", 0) or 0
        uptime_days = round(uptime / 86400, 1)
        heard_24 = heard_count_24.get(s["mac"], 0)

        print("    " + str(dev.get("name", "?")) + "  " + s["mac"]
              + "  model " + str(dev.get("model", "?"))
              + "  fw " + str(dev.get("version", "?"))
              + "  up " + str(uptime_days) + "d")
        print("      5GHz ch " + str(r5.get("channel", "?"))
              + "  pwr " + str(r5.get("power", "?"))
              + "  clients " + str(r5.get("num_clients", "?"))
              + "  | hears " + str(s["count"]) + " nbrs, strongest " + str(s["strongest_rssi"])
              + " (" + s["strongest_mac"] + ")  | heard by 0")
        print("      2.4GHz clients " + str(r24.get("num_clients", "?"))
              + "  | heard by " + str(heard_24) + " AP(s) on 2.4GHz")

        writer.writerow([
            run_time, site_name, site_id, dev.get("name", ""), s["mac"], dev.get("model", ""),
            dev.get("version", ""), uptime_days, dev.get("status", ""),
            r5.get("channel", ""), r5.get("power", ""), r5.get("num_clients", ""),
            r24.get("num_clients", ""), s["count"], s["strongest_rssi"], s["strongest_mac"],
            "yes" if heard_24 > 0 else "no", heard_24,
        ])

        report_rows.append([
            site_name, dev.get("name", ""), s["mac"], dev.get("model", ""), dev.get("version", ""),
            uptime_days, dev.get("status", ""), r5.get("channel", ""), r5.get("power", ""),
            r5.get("num_clients", ""), r24.get("num_clients", ""), s["count"], s["strongest_rssi"],
            s["strongest_mac"], heard_24,
        ])

        failed_aps.append({
            "site_id": site_id, "site_name": site_name,
            "name": dev.get("name", "") or s["mac"], "mac": s["mac"],
        })

    csv_handle.flush()

csv_handle.close()

# ---------- 3. offer to reset the radios on the failed APs ----------
#
# Turning the radio off and on again clears the stuck state (the same thing an RRM
# channel change does). This shows the failed APs, asks for confirmation, then:
#   1. disables the radio(s) in reset_bands on every failed AP
#   2. waits reset_wait_seconds
#   3. puts each AP's original radio_config back (which re-enables the radio)
# With a read-only API token nothing can be changed, so it shows what WOULD
# have been done and says clearly that nothing was changed.

reset_result_by_mac = {}
reset_done = 0
reset_failed = 0
reset_mode = "not offered"

if OFFER_RADIO_RESET and len(failed_aps) > 0:
    print("")
    print("==== Radio reset ====")
    print("")
    print("APs with a stuck 5GHz radio (" + str(len(failed_aps)) + "):")
    n = 0
    for ap in sorted(failed_aps, key=lambda a: (a["site_name"], a["name"])):
        n = n + 1
        print("  " + str(n).rjust(3) + ". " + ap["site_name"] + "  /  " + ap["name"] + "  (" + ap["mac"] + ")")
    print("")

    # --- can this token write? Check /self first, then find out for sure on the first change ---
    can_write = None          # None = unknown, True = has write rights, False = read-only
    resp = session.get(API_HOST + "/api/v1/self")
    if resp.status_code == 200 and isinstance(resp.json(), dict):
        org_role = ""
        for priv in resp.json().get("privileges", []) or []:
            if priv.get("scope") == "org" and priv.get("org_id") == ORG_ID:
                org_role = priv.get("role", "")
        if org_role in ("admin", "write"):
            can_write = True
        elif org_role != "":
            # read / helpdesk / installer etc. at org level - can't change AP config
            # (site-level write rights would still be caught by the first change below)
            can_write = False
            for priv in resp.json().get("privileges", []) or []:
                if priv.get("scope") == "site" and priv.get("role") in ("admin", "write"):
                    can_write = None

    band_names = ", ".join([b.replace("band_24", "2.4GHz").replace("band_5", "5GHz").replace("band_6", "6GHz") for b in RESET_BANDS])

    if can_write is False:
        print("NOTE: this API token is READ-ONLY for this org, so the radios can't be changed.")
        print("      If you continue, the script will show what it WOULD do - nothing will be changed.")
        print("")

    answer = ""
    try:
        answer = input("Reset the " + band_names + " radio on these " + str(len(failed_aps))
                       + " APs (off, wait " + str(RESET_WAIT_SECONDS) + "s, back on)? [y/N]: ").strip().lower()
    except EOFError:
        print("")
        print("No keyboard input available (running unattended?) - skipping the radio reset.")

    if answer not in ("y", "yes"):
        print("Radio reset skipped - no changes made.")
        reset_mode = "skipped by user"
        for ap in failed_aps:
            reset_result_by_mac[ap["mac"]] = "Not reset (skipped)"
    else:
        dry_run = (can_write is False)
        reset_mode = "dry run (read-only token)" if dry_run else "live"

        # --- read each AP's current radio_config, so it can be put back exactly ---
        for ap in failed_aps:
            ap["device_id"] = "00000000-0000-0000-1000-" + ap["mac"]
            ap["original_rc"] = None
            ap["disabled_ok"] = False
            while True:
                resp = session.get(API_HOST + "/api/v1/sites/" + ap["site_id"] + "/devices/" + ap["device_id"])
                if resp.status_code == 429:
                    print("  rate limited, waiting 60s...")
                    time.sleep(60)
                    continue
                break
            if resp.status_code == 200:
                ap["original_rc"] = resp.json().get("radio_config") or {}
            else:
                print("  [FAILED] " + ap["site_name"] + " / " + ap["name"] + ": couldn't read device config (HTTP "
                      + str(resp.status_code) + ") - this AP will be left alone")
                reset_result_by_mac[ap["mac"]] = "Not reset (couldn't read config, HTTP " + str(resp.status_code) + ")"
                reset_failed = reset_failed + 1
            time.sleep(PAUSE_BETWEEN_CALLS)

        # --- step 1: disable ---
        print("")
        if dry_run:
            print("*" * 78)
            print("*  DRY RUN - READ-ONLY TOKEN - NOTHING BELOW WAS ACTUALLY CHANGED             *")
            print("*" * 78)
        print("Step 1/3  " + datetime.datetime.now().strftime("%H:%M:%S") + "  disabling " + band_names + " radio")
        for ap in failed_aps:
            if ap["original_rc"] is None:
                continue
            new_rc = json.loads(json.dumps(ap["original_rc"]))
            for band in RESET_BANDS:
                band_cfg = new_rc.get(band) or {}
                band_cfg["disabled"] = True
                new_rc[band] = band_cfg
            ap["disable_rc"] = new_rc
            url = API_HOST + "/api/v1/sites/" + ap["site_id"] + "/devices/" + ap["device_id"]

            if dry_run:
                print("  [WOULD DISABLE - NOT DONE] " + ap["site_name"] + " / " + ap["name"])
                print("        PUT " + url.replace(API_HOST, ""))
                print("        " + json.dumps({"radio_config": {b: new_rc[b] for b in RESET_BANDS}}))
                continue

            while True:
                resp = session.put(url, json={"radio_config": new_rc})
                if resp.status_code == 429:
                    print("  rate limited, waiting 60s...")
                    time.sleep(60)
                    continue
                break
            if resp.status_code == 200:
                ap["disabled_ok"] = True
                print("  [OK]     " + datetime.datetime.now().strftime("%H:%M:%S") + "  " + ap["site_name"] + " / " + ap["name"] + "  -> radio disabled")
            elif resp.status_code in (401, 403):
                # Token turned out to be read-only - switch to dry run for the rest
                print("  [DENIED] " + ap["site_name"] + " / " + ap["name"] + "  -> HTTP " + str(resp.status_code)
                      + " - this API token has no write access")
                dry_run = True
                reset_mode = "dry run (read-only token)"
                print("")
                print("*" * 78)
                print("*  READ-ONLY TOKEN - switching to DRY RUN. NOTHING BELOW WAS ACTUALLY CHANGED *")
                print("*" * 78)
                print("  [WOULD DISABLE - NOT DONE] " + ap["site_name"] + " / " + ap["name"])
                print("        PUT " + url.replace(API_HOST, ""))
                print("        " + json.dumps({"radio_config": {b: new_rc[b] for b in RESET_BANDS}}))
            else:
                print("  [FAILED] " + ap["site_name"] + " / " + ap["name"] + "  -> HTTP " + str(resp.status_code)
                      + " " + resp.text[:200] + " - NOT changed")
                reset_result_by_mac[ap["mac"]] = "Not reset (disable failed, HTTP " + str(resp.status_code) + ")"
                reset_failed = reset_failed + 1
            time.sleep(PAUSE_BETWEEN_CALLS)

        # --- step 2: wait ---
        any_disabled = len([ap for ap in failed_aps if ap["disabled_ok"]]) > 0
        if any_disabled or dry_run:
            prefix = "(dry run) " if dry_run and not any_disabled else ""
            print("Step 2/3  " + datetime.datetime.now().strftime("%H:%M:%S") + "  " + prefix + "waiting " + str(RESET_WAIT_SECONDS) + " seconds...")
            if any_disabled:
                try:
                    time.sleep(RESET_WAIT_SECONDS)
                except KeyboardInterrupt:
                    # Never leave radios switched off - carry straight on to re-enable them
                    print("  Interrupted - re-enabling the radios now")

        # --- step 3: re-enable by putting the original radio_config back ---
        print("Step 3/3  " + datetime.datetime.now().strftime("%H:%M:%S") + "  re-enabling " + band_names + " radio")
        left_disabled = []
        for ap in failed_aps:
            if ap["original_rc"] is None:
                continue
            url = API_HOST + "/api/v1/sites/" + ap["site_id"] + "/devices/" + ap["device_id"]

            if dry_run and not ap["disabled_ok"]:
                print("  [WOULD RE-ENABLE - NOT DONE] " + ap["site_name"] + " / " + ap["name"])
                print("        PUT " + url.replace(API_HOST, ""))
                print("        " + json.dumps({"radio_config": {b: ap["original_rc"].get(b, {}) for b in RESET_BANDS}}) + "   (original config put back)")
                reset_result_by_mac[ap["mac"]] = "DRY RUN - would have reset (NOT done, read-only token)"
                continue
            if not ap["disabled_ok"]:
                continue

            ok = False
            tries = 0
            while tries < 5:
                tries = tries + 1
                try:
                    resp = session.put(url, json={"radio_config": ap["original_rc"]})
                except requests.exceptions.RequestException as e:
                    print("  retrying " + ap["name"] + " after error: " + str(e))
                    time.sleep(5)
                    continue
                if resp.status_code == 429:
                    print("  rate limited, waiting 60s...")
                    time.sleep(60)
                    continue
                if resp.status_code == 200:
                    ok = True
                    break
                print("  retrying " + ap["name"] + " after HTTP " + str(resp.status_code))
                time.sleep(5)
            if ok:
                print("  [OK]     " + datetime.datetime.now().strftime("%H:%M:%S") + "  " + ap["site_name"] + " / " + ap["name"] + "  -> radio re-enabled")
                reset_result_by_mac[ap["mac"]] = "Reset OK " + datetime.datetime.now().strftime("%H:%M:%S")
                reset_done = reset_done + 1
            else:
                print("  [FAILED] " + ap["site_name"] + " / " + ap["name"] + "  -> could NOT re-enable the radio!")
                reset_result_by_mac[ap["mac"]] = "RADIO LEFT DISABLED - re-enable manually"
                reset_failed = reset_failed + 1
                left_disabled.append(ap)
            time.sleep(PAUSE_BETWEEN_CALLS)

        print("")
        if dry_run:
            print("*" * 78)
            if reset_done == 0:
                print("*  DRY RUN ONLY - the API token is read-only, so NO radios were changed.      *")
            else:
                print("*  PARTLY DRY RUN - " + str(reset_done).ljust(3) + " AP(s) were reset before write access was refused; *")
                print("*  the rest were NOT changed.                                                *")
            print("*  Use a token with write access to the org to actually reset them.          *")
            print("*" * 78)
        else:
            print("Radio reset finished: " + str(reset_done) + " reset OK, " + str(reset_failed) + " not reset")
        if len(left_disabled) > 0:
            print("")
            print("!!! WARNING: these APs still have the radio DISABLED - re-enable them in the Mist dashboard:")
            for ap in left_disabled:
                print("!!!   " + ap["site_name"] + " / " + ap["name"] + "  (" + ap["mac"] + ")")

for ap in failed_aps:
    if ap["mac"] not in reset_result_by_mac:
        reset_result_by_mac[ap["mac"]] = "Not reset"


# ---------- 4. Excel report of the failed APs ----------

# Filename: org name + site filter (or ALL-SITES) + run time, e.g.
#   deaf_5g_report_My-Org_SITE-01_2026-09-23_1015UTC.xlsx
if SITE_NAME_FILTER != "":
    name_part = SITE_NAME_FILTER
elif len(sites) == 1:
    name_part = sites[0].get("name", "site")
else:
    name_part = "ALL-SITES"
name_part = re.sub(r"[^A-Za-z0-9_-]+", "-", name_part).strip("-")
report_file = os.path.join(REPORT_FOLDER, "deaf_5g_report_" + org_file_part + "_" + name_part + "_"
                           + run_now.strftime("%Y-%m-%d_%H%M") + "UTC.xlsx")

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="1F3864")
red_fill = PatternFill("solid", fgColor="F8CBAD")
amber_fill = PatternFill("solid", fgColor="FFE699")
title_font = Font(bold=True, size=14)

wb = Workbook()

# --- sheet 1: failed APs ---
ws = wb.active
ws.title = "Failed APs"
headers = [
    "Site", "AP name", "AP MAC", "Model", "Firmware", "Uptime (days)", "Status",
    "5GHz channel", "5GHz power", "5GHz clients", "2.4GHz clients",
    "5GHz neighbours heard", "Strongest heard RSSI", "Strongest heard AP", "Heard by (2.4GHz)",
    "Radio reset",
]
ws.append(headers)
for cell in ws[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[1].height = 32

if len(report_rows) == 0:
    ws.append(["No failed APs found in this run"])
else:
    report_rows.sort(key=lambda r: (str(r[0]), str(r[1])))
    for r in report_rows:
        r = r + [reset_result_by_mac.get(r[2], "Not reset")]
        ws.append(r)
        row_num = ws.max_row
        # Colour by how strong the evidence is: hears a neighbour loudly but nobody hears it
        rssi = r[12]
        if isinstance(rssi, (int, float)) and rssi >= -65:
            fill = red_fill
        else:
            fill = amber_fill
        for col in range(1, len(headers) + 1):
            ws.cell(row=row_num, column=col).fill = fill

widths = [28, 22, 15, 8, 14, 10, 12, 10, 10, 10, 10, 12, 12, 16, 12, 34]
for i in range(len(widths)):
    ws.column_dimensions[chr(65 + i)].width = widths[i]
ws.freeze_panes = "A2"
ws.auto_filter.ref = "A1:" + chr(64 + len(headers)) + str(ws.max_row)

# --- sheet 2: summary ---
ws2 = wb.create_sheet("Summary")
ws2.append(["Deaf-in-transmit 5GHz AP report"])
ws2["A1"].font = title_font
ws2.append([])
summary_rows = [
    ["Run time", run_time],
    ["Org", org_name],
    ["Org ID", ORG_ID],
    ["Site filter", SITE_NAME_FILTER if SITE_NAME_FILTER != "" else "(all sites)"],
    ["Sites checked", sites_done],
    ["Sites skipped (<2 APs)", sites_skipped],
    ["Sites with failed APs", sites_with_suspects],
    ["Failed APs", total_suspects],
    [],
    ["Ignored - uptime < " + str(MIN_UPTIME_DAYS) + " day", total_ignored_uptime],
    ["Ignored - has 5GHz clients", total_ignored_clients],
    ["Ignored - 5GHz power 0", total_ignored_power],
    ["Ignored - 5GHz channel/power unknown", total_ignored_unknown],
    ["Ignored - isolated (not heard on 2.4GHz)", total_ignored_isolated],
    [],
    ["Radio reset", reset_mode],
    ["Radios reset OK", reset_done],
    ["Radios not reset", reset_failed],
    [],
    ["Detection rule", "Hears >= " + str(MIN_NEIGHBOURS_HEARD) + " 5GHz neighbours, strongest >= "
        + str(STRONGEST_RSSI_THRESHOLD) + " dBm, heard by 0 APs on 5GHz"],
    ["Colour key", "Red = strongest heard RSSI >= -65 dBm (strong evidence); Amber = weaker"],
]
for r in summary_rows:
    ws2.append(r)
for row in ws2.iter_rows(min_row=3, max_col=1):
    for cell in row:
        cell.font = Font(bold=True)
ws2.column_dimensions["A"].width = 40
ws2.column_dimensions["B"].width = 90

os.makedirs(REPORT_FOLDER, exist_ok=True)
wb.save(report_file)

# ---------- 5. summary ----------

print("")
print("==== Summary ====")
print("Sites checked:            " + str(sites_done))
print("Sites skipped (<2 APs):   " + str(sites_skipped))
print("Sites with suspects:      " + str(sites_with_suspects))
print("Suspect APs:              " + str(total_suspects))
print("Ignored (uptime < " + str(MIN_UPTIME_DAYS) + " day): " + str(total_ignored_uptime))
print("Ignored (has 5GHz clients): " + str(total_ignored_clients))
print("Ignored (5GHz power 0):     " + str(total_ignored_power))
print("Ignored (5GHz ch/pwr unknown): " + str(total_ignored_unknown))
print("Ignored (isolated - not heard on 2.4GHz): " + str(total_ignored_isolated))
print("Radio reset:              " + reset_mode + (" (" + str(reset_done) + " OK, " + str(reset_failed) + " not reset)" if reset_mode == "live" else ""))
print("Results appended to:      " + CSV_FILE)
print("Excel report:             " + report_file)
