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
5. Prints the results and appends them to a CSV, so repeated runs build a history
6. Writes an Excel report of the failed APs for that run, e.g. `deaf_5g_report_SITE-01_2026-09-23_1015UTC.xlsx`
   (the site filter, or ALL-SITES, plus the run time). It has a Failed APs sheet (red = strong evidence,
   amber = weaker) and a Summary sheet

## Setup

```
pip3 install requests openpyxl
cp find_deaf_5g_aps.ini.example find_deaf_5g_aps.ini
# edit find_deaf_5g_aps.ini: add api_token (and org_id / api_host if different)
python3 find_deaf_5g_aps.py
```

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
