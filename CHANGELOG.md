# Changelog

## v1.2.0

- Works behind company HTTPS proxies. It trusts the Windows/macOS certificate store via `truststore`
  (`use_system_certs`, on by default), or you can set `ca_bundle` to a company root certificate
- Checks the connection first, so a certificate or proxy problem gives a clear message and not a traceback
- Adds `truststore` to `requirements.txt`

## v1.1.1

- Adds `requirements.txt` (`requests`, `openpyxl`) so setup is `pip3 install -r requirements.txt`

## v1.1.0

- Looks up the org name and uses it in the filenames:
  - Excel report: `deaf_5g_report_<org>_<site filter or ALL-SITES>_<time>UTC.xlsx`
  - CSV history: `deaf_5g_aps_<org>.csv` (set by `csv_file`; `{org}` is replaced with the org name)
- The org name is also shown in the Summary sheet and at the start of each run

## v1.0.0

First public release.

- Scans every site in a Mist org (or only the sites that match `site_name_filter`) for APs whose
  5GHz radio has gone deaf-in-transmit: the AP hears its 5GHz neighbours, but none of them hears it
- Filters out false positives. It ignores an AP if:
  - its uptime is under `min_uptime_days`, because RRM won't have run yet
  - it has 5GHz clients connected, because the radio must be working
  - its 5GHz power is 0, because no SSIDs are configured on that radio
  - its 5GHz channel or power is unknown
  - it is not heard on 2.4GHz either, because it is probably just isolated
- Settings come from an `.ini` file (the real `.ini` is git-ignored; an `.example` is included)
- Appends results to a CSV history file each run
- Writes an Excel report for each run, named with the site filter and the run time
