# mist-deaf-5g-detector

Finds Mist APs whose 5GHz radio has stopped transmitting properly ("deaf-in-transmit"):
the AP still **hears** its neighbours on 5GHz, but **no neighbour hears it**.

Built while investigating intermittent AP32 5GHz failures: the radio kept beaconing according to
its own counters, but neighbours stopped seeing it and 5GHz clients dropped to zero.

## How it works

For every site in the org (or only the sites that match a name filter):

1. Pulls `/sites/{site_id}/rrm/neighbors/band/5`
2. Flags an AP when all three of these are true:
   - no other AP lists it as a 5GHz neighbour
   - it hears at least `min_neighbours_heard` neighbours
   - its strongest neighbour is at least `strongest_rssi_threshold` dBm
3. For each flagged AP, checks whether it is still heard on 2.4GHz (`band/24`), and pulls its name,
   model, firmware, uptime, 5GHz channel/power and client counts from `/sites/{site_id}/stats/devices`
4. Ignores flagged APs that have been up for less than `min_uptime_days` (default 1), because RRM won't have
   built a neighbour view for them yet
   - also ignores flagged APs with clients on 5GHz (`ignore_if_5g_clients`), because the radio must be working
   - also ignores flagged APs whose 5GHz power is 0 (`ignore_if_5g_power_zero`), because no SSIDs are configured on that radio
   - also ignores flagged APs whose 5GHz channel or power is unknown (`ignore_if_5g_unknown`), because there are no 5GHz radio stats to judge them on
   - also ignores flagged APs that no other AP hears on 2.4GHz either (`ignore_if_isolated`), because they are probably just isolated
5. Prints the results and appends them to a CSV named with the org (`deaf_5g_aps_<org>.csv`), so repeated runs build a history for each org
6. Writes an Excel report of the failed APs for that run, e.g. `deaf_5g_report_My-Org_SITE-01_2026-09-23_1015UTC.xlsx`
   (org name, then the site filter or ALL-SITES, then the run time). It has a Failed APs sheet (red = strong evidence,
   amber = weaker) and a Summary sheet

## Radio reset (optional)

Once the scan has finished, the script lists the failed APs and asks:

```
Reset the 5GHz radio on these 3 APs (off, wait 5.0s, back on)? [y/N]:
```

If you answer `y`, it does three things:
1. Disables the radio on every failed AP (`PUT /sites/{site_id}/devices/{device_id}`, `radio_config.<band>.disabled = true`)
2. Waits `reset_wait_seconds`
3. Puts back each AP's original `radio_config`, which re-enables the radio

Each step is timestamped in the output. The result for each AP goes in the Excel report's "Radio reset" column.

- **Read-only API token:** nothing can be changed. The script shows exactly what it *would* have sent and
  marks it clearly as a **DRY RUN, nothing changed**. It detects this from `/self`, or from the first
  HTTP 401/403 response.
- **Radio failed to come back on:** the script retries 5 times. If it still fails, it lists the AP in a
  warning, so you can re-enable it in the dashboard.
- **Ctrl-C during the wait:** the script still turns the radios back on.
- **Running unattended (no keyboard):** the reset is skipped.
- **Settings:** under `[reset]` in the `.ini`, set `offer_radio_reset`, `reset_bands` (`5`, or `24,5`)
  and `reset_wait_seconds`.

## Setup

```
pip3 install -r requirements.txt
cp find_deaf_5g_aps.ini.example find_deaf_5g_aps.ini
# edit find_deaf_5g_aps.ini: add api_token (and org_id / api_host if different)
python3 find_deaf_5g_aps.py
```

### `CERTIFICATE_VERIFY_FAILED` on a company network

If you see `certificate verify failed: unable to get local issuer certificate`, a company proxy is
probably inspecting HTTPS traffic. `pip install -r requirements.txt` installs `truststore`, which
lets the script trust the Windows/macOS certificate store (`use_system_certs = true`, the default).
You can also set `ca_bundle` in the `.ini` to the path of your company's root certificate (`.pem`).

Use a different settings file: `python3 find_deaf_5g_aps.py other_org.ini`

`find_deaf_5g_aps.ini`, `*.csv` and `*.xlsx` are git-ignored, so tokens and results never get committed.

## Reading the results

A real case usually looks like this:
- it hears strong 5GHz neighbours but is heard by 0
- it has 0 5GHz clients
- it is still heard by neighbours on 2.4GHz

APs at the edge of a building can be flagged once in a while, because one-way links happen naturally
there. If an AP gets flagged run after run, it is probably a real case.

## API usage

A healthy site costs 1 call. A site with a flagged AP costs 3. `pause_between_calls` keeps the
script well under the Mist limit of 5000 calls/hour. If it gets HTTP 429, it waits 60s and retries.
