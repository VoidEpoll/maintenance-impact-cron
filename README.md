# Get-MaintenanceCrons.ps1

Reads one or many crontab text files (one per server/user) and lists every
cron that will fire during a maintenance window, in run order, with its
name/comment tag shown above it.
Commented-out crons are listed too, clearly marked as **COMMENTED - will NOT run**.

Default window: **Sat 14-Nov-2026 00:30 -> 12:30** (end time included).

---

## 1. Prepare the cron files

On each Linux server, export each user's crontab to a file named **`<hostname>-<username>.txt`**:

```bash
crontab -l > $(hostname -s)-$(whoami).txt          # current user
sudo crontab -l -u webid > $(hostname -s)-webid.txt # another user
```

Example: `lbat017-webid.txt`, `lbat017-root.txt`, `lbat018-oracle.txt`.

Copy all files into one folder on your Windows machine (WinSCP, scp, etc.), e.g. `.\crons\`.

> **Host/User from file name:** the name is split on the **last** `-`, so
> `lb-app-02-oracle.txt` -> host `lb-app-02`, user `oracle`. A file without a
> `-` gets the whole name as host and an empty user.

> **Timezone:** the script compares cron times exactly as written. If the server
> runs in a different timezone (e.g. UTC) than the maintenance window (e.g. IST),
> convert the window to server time and pass it with `-Start` / `-End`.
> Check server time with `date` or `timedatectl`.

---

## 2. Run the script

Open PowerShell in this folder:

```powershell
# All files in a folder + audit + combined Excel CSV (recommended)
.\Get-MaintenanceCrons.ps1 -CronFile .\crons -Audit -Csv

# Specific files / wildcard
.\Get-MaintenanceCrons.ps1 -CronFile .\crons\lbat017-webid.txt, .\crons\lbat017-root.txt
.\Get-MaintenanceCrons.ps1 -CronFile .\crons\lbat017-*.txt

# Single file
.\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt

# Report + validation audit (recommended)
.\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt -Audit

# Audit on screen only - no .txt, no .csv, no clipboard
.\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt -Audit -ScreenOnly

# Custom window
.\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt -Start "2026-11-14 00:30" -End "2026-11-14 12:30"
```

If you get *"running scripts is disabled on this system"*:

```powershell
powershell -ExecutionPolicy Bypass -File .\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt -Audit
```

### Parameters

| Parameter      | Required | Default              | Description                                   |
|----------------|----------|----------------------|-----------------------------------------------|
| `-CronFile`    | Yes      |                      | One or more files, folders (all `*.txt` inside) or wildcards |
| `-Start`       | No       | `2026-11-14 00:30`   | Window start (inclusive)                      |
| `-End`         | No       | `2026-11-14 12:30`   | Window end (inclusive)                        |
| `-OutFile`     | No       | `maintenance_crons_<yyyyMMdd>.txt` in current folder | Where to save the report |
| `-Audit`       | No       | off                  | Add validation checks and an all-crons table to the output |
| `-NoClipboard` | No       | off                  | Don't copy the report to the clipboard        |
| `-ScreenOnly`  | No       | off                  | Print only: no .txt, no .csv, no clipboard    |
| `-Csv`         | No       | off                  | Also export all crons of all files to one `.csv` for Excel, with **Host** and **User** columns |

---

## 3. Output

Everything (report and, with `-Audit`, the audit and all-crons table) is **printed**,
**saved to one .txt file** (identical to the screen) and **copied to the clipboard**,
ready to paste into email, Teams or a change ticket.

```
MAINTENANCE WINDOW : Sat 14-Nov-2026 00:30  ->  Sat 14-Nov-2026 12:30
CRONS IN WINDOW    : 4 active, 1 commented (will not run)
----------------------------------------------------------------------
1) 00:30  [ACTIVE]  (line 9)
   ###### Health check every 15 min
   */15 * * * * /opt/scripts/health.sh
   Runs: 00:30, 00:45, 01:00 ... 12:30  (49 runs)

2) 02:30  [ACTIVE]  (line 6)
   ###### Daily DB backup
   30 2 * * * /opt/scripts/db_backup.sh
   Runs: 02:30  (1 run(s))

3) 03:00  [COMMENTED - will NOT run]  (line 12)
   ###### Old cleanup job - disabled
   #0 3 * * * /opt/scripts/cleanup.sh
   Runs: 03:00  (1 run(s))
```

The report starts with a summary per server/user:

```
Host      User   TotalCrons ActiveInWindow CommentedInWindow Audit
----      ----   ---------- -------------- ----------------- -----
lbat017   root            5              3                 2 OK
lbat017   webid          10              4                 1 CHECK
```

`Audit = CHECK` means that file has lines needing a manual look (run with `-Audit` to see them).
Then one section per file, headed `HOST: lbat017    USER: webid`, containing:

1. **Crons in window**: active and commented, sorted by first run time.
2. **Other commented crons**: commented crons outside the window, for reference.
3. **@reboot crons**: these run if the server is restarted during maintenance.

---

## 4. How the file is read

| Line in crontab                              | Treated as                                  |
|----------------------------------------------|---------------------------------------------|
| `30 2 * * * /opt/job.sh`                     | Active cron                                 |
| `#30 2 * * * ...` / `###### 30 2 * * * ...`  | Commented cron (will NOT run)               |
| `###### Daily backup` / `# any text`         | Name/tag for the cron directly below        |
| `##########` (only hashes)                   | Ignored (separator)                         |
| empty / whitespace-only line                 | Blank; breaks the tag-to-cron link        |
| `SHELL=/bin/bash`, `MAILTO=...`              | Environment line, ignored                   |
| `@daily`, `@hourly`, `@weekly`, `@monthly`, `@yearly` | Converted to normal schedules      |
| `@reboot ...`                                | Listed in its own section                   |

**Tags:** all comment lines *directly* above a cron (no blank line in between)
are shown as its name. If a cron shows no name, there's a blank line between it
and its tag in the file.

**Supported cron syntax:** `*`, `5`, `1,15,30`, `9-17`, `*/10`, `0-30/5`,
month names (`jan`-`dec`), day names (`sun`-`sat`), `7` = Sunday. If both
day-of-month and day-of-week are set, the job runs when **either** matches
(standard cron behaviour).

---

## 5. Validating the result (`-Audit`)

Use `-Audit` on large files (200+ crons) to confirm nothing was missed.

```
==================== AUDIT ====================
Total lines in file        : 412
  BLANK                    : 60
  COMMENT/TAG              : 130
  CRON-ACTIVE              : 198
  CRON-COMMENTED           : 22
  ENV                      : 2
Line reconciliation        : OK (every line classified)
Raw lines that look active : 198
Active crons parsed        : 198
Active cron cross-check    : OK

Lines NOT understood       : none

ALL 220 CRONS:
LineNo Status    InWindow FirstRun RunCount Name                   Cron
------ ------    -------- -------- -------- ----                   ----
     6 ACTIVE    YES      02:30           1 ###### Daily DB backup 30 2 * * * /opt/scripts/db_backup.sh
   ...
```

### Checklist

1. **Line reconciliation = OK**: every line of the file was accounted for.
2. **Active cron cross-check = OK**: an independent count matches the parsed count.
3. **"Lines NOT understood" is empty**: if not, check each listed line by hand
   (e.g. `0 25 * * *` has an invalid hour).
4. **Match the server**: on Linux run
   ```bash
   crontab -l | grep -cE '^\s*[0-9*@]'
   ```
   The number must equal **Active crons parsed** for that host/user.
5. **Spot-check the ALL CRONS table** (in the .txt, or in Excel with `-Csv`; filter by Host/User):
   - Filter `InWindow = NO` and look at crons near the window edges (hour `0`, `12`).
   - Filter `InWindow = YES` and verify 5-10 by hand (or paste into https://crontab.guru).

If all five pass, the report can be trusted.

---

## 6. Files

| File                                    | Purpose                              |
|-----------------------------------------|--------------------------------------|
| `Get-MaintenanceCrons.ps1`              | The script                           |
| `maintenance_crons_<date>.txt`          | Everything shown on screen (copy-paste ready) |
| `maintenance_crons_<date>.csv`          | All crons of all files for Excel, columns `Host, User, LineNo, Status, InWindow, FirstRun, RunCount, Runs, Name, Cron, File` (`-Csv` only) |
