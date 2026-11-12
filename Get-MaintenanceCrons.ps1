<#
.SYNOPSIS
    Lists the cron jobs that fire during a maintenance window, in run order,
    with the comment/name lines that sit above each cron.

.DESCRIPTION
    - Active crons that run in the window are listed in time order.
    - Commented-out crons (e.g. "#30 2 * * * /opt/job.sh") that would have run
      in the window are listed in the same order, marked as COMMENTED.
    - All other commented-out crons are listed at the end for reference.
    - Comment lines directly above a cron (no blank line in between) are
      treated as its name/tag.
    Output is printed, saved to a text file and copied to the clipboard.

.EXAMPLE
    .\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt

.EXAMPLE
    .\Get-MaintenanceCrons.ps1 -CronFile .\crontab.txt -Start "2026-11-14 00:30" -End "2026-11-14 12:30"
#>
param(
    [Parameter(Mandatory = $true)][string]$CronFile,
    [datetime]$Start = '2026-11-14 00:30',
    [datetime]$End   = '2026-11-14 12:30',
    [string]$OutFile,
    [switch]$NoClipboard,
    [switch]$Audit,
    [switch]$ScreenOnly,  # print only: no .txt, no .csv, no clipboard
    [switch]$Csv          # with -Audit, also export all crons to .audit.csv
)

$ErrorActionPreference = 'Stop'

$MonthNames = @{ jan = 1; feb = 2; mar = 3; apr = 4; may = 5; jun = 6; jul = 7; aug = 8; sep = 9; oct = 10; nov = 11; dec = 12 }
$DowNames   = @{ sun = 0; mon = 1; tue = 2; wed = 3; thu = 4; fri = 5; sat = 6 }
$Specials   = @{
    '@yearly'   = '0 0 1 1 *'
    '@annually' = '0 0 1 1 *'
    '@monthly'  = '0 0 1 * *'
    '@weekly'   = '0 0 * * 0'
    '@daily'    = '0 0 * * *'
    '@midnight' = '0 0 * * *'
    '@hourly'   = '0 * * * *'
}

function Resolve-CronValue {
    param([string]$Token, [hashtable]$Names)
    if ($Token -match '^\d+$') { return [int]$Token }
    if ($Names -and $Names.ContainsKey($Token)) { return $Names[$Token] }
    throw "Invalid cron value '$Token'"
}

function Expand-CronField {
    param([string]$Field, [int]$Min, [int]$Max, [hashtable]$Names)
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($part in $Field.ToLower().Split(',')) {
        if ($part -notmatch '^(\*|\?|[0-9a-z]+(?:-[0-9a-z]+)?)(?:/(\d+))?$') { throw "Invalid cron field '$Field'" }
        $range   = $Matches[1]
        $stepStr = $Matches[2]
        $step = 1
        if ($stepStr) { $step = [int]$stepStr }
        if ($step -lt 1) { throw "Invalid step in '$Field'" }

        if ($range -eq '*' -or $range -eq '?') {
            $lo = $Min; $hi = $Max
        } else {
            $ends = $range.Split('-')
            $lo = Resolve-CronValue $ends[0] $Names
            if ($ends.Count -gt 1) { $hi = Resolve-CronValue $ends[1] $Names }
            elseif ($stepStr)      { $hi = $Max }   # "5/10" means 5-max/10
            else                   { $hi = $lo }
        }
        if ($lo -lt $Min -or $hi -gt $Max -or $lo -gt $hi) { throw "Out of range in '$Field'" }
        for ($v = $lo; $v -le $hi; $v += $step) { [void]$set.Add($v) }
    }
    return , $set
}

# Returns a parsed cron object, or $null when the text is not a cron line.
function ConvertFrom-CronLine {
    param([string]$Text)
    $t = $Text.Trim()
    if ($t -match '^(@\w+)\s+(.+)$') {
        $key = $Matches[1].ToLower()
        $cmd = $Matches[2]
        if ($key -eq '@reboot') {
            return [pscustomobject]@{ Reboot = $true; Command = $cmd }
        }
        if (-not $Specials.ContainsKey($key)) { return $null }
        $fields = $Specials[$key] -split '\s+'
    } else {
        $tokens = $t -split '\s+', 6
        if ($tokens.Count -lt 6) { return $null }
        $fields = $tokens[0..4]
        $cmd = $tokens[5]
    }
    try {
        $dow = Expand-CronField $fields[4] 0 7 $DowNames
        if ($dow.Remove(7)) { [void]$dow.Add(0) }   # 7 = Sunday
        return [pscustomobject]@{
            Reboot  = $false
            Minutes = Expand-CronField $fields[0] 0 59 $null
            Hours   = Expand-CronField $fields[1] 0 23 $null
            Dom     = Expand-CronField $fields[2] 1 31 $null
            Months  = Expand-CronField $fields[3] 1 12 $MonthNames
            Dow     = $dow
            DomStar = $fields[2] -match '^[\*\?]'
            DowStar = $fields[4] -match '^[\*\?]'
            Command = $cmd
        }
    } catch {
        return $null
    }
}

function Get-CronRunTimes {
    param($Entry, [datetime]$From, [datetime]$To)
    $runs = New-Object 'System.Collections.Generic.List[datetime]'
    $t = $From.Date.AddHours($From.Hour).AddMinutes($From.Minute)
    if ($t -lt $From) { $t = $t.AddMinutes(1) }
    while ($t -le $To) {
        if ($Entry.Minutes.Contains($t.Minute) -and $Entry.Hours.Contains($t.Hour) -and $Entry.Months.Contains($t.Month)) {
            $domOk = $Entry.Dom.Contains($t.Day)
            $dowOk = $Entry.Dow.Contains([int]$t.DayOfWeek)
            # Standard cron: if both day fields are restricted, either one matching is enough
            if ($Entry.DomStar -or $Entry.DowStar) { $dayOk = $domOk -and $dowOk } else { $dayOk = $domOk -or $dowOk }
            if ($dayOk) { $runs.Add($t) }
        }
        $t = $t.AddMinutes(1)
    }
    return , $runs
}

function Format-RunTimes {
    param($Runs, [string]$Fmt)
    $times = @($Runs | ForEach-Object { $_.ToString($Fmt) })
    if ($times.Count -le 12) { return ($times -join ', ') + "  ($($times.Count) run(s))" }
    return "$($times[0..2] -join ', ') ... $($times[-1])  ($($times.Count) runs)"
}

# ---------------------------------------------------------------- parse file
if (-not (Test-Path -LiteralPath $CronFile)) { throw "Cron file not found: $CronFile" }
if ($End -lt $Start) { throw "End time is before start time." }

$lines   = @(Get-Content -LiteralPath $CronFile)
$entries = New-Object 'System.Collections.Generic.List[object]'
$pending = @()
$lineTypes = @{}   # line number -> classification, used by -Audit

for ($i = 0; $i -lt $lines.Count; $i++) {
    $trim = $lines[$i].Trim()
    if ($trim -eq '') { $pending = @(); $lineTypes[$i + 1] = 'BLANK'; continue }

    $isComment = $trim.StartsWith('#')
    $body = $trim
    if ($isComment) { $body = $trim.TrimStart('#').Trim() }

    $parsed = ConvertFrom-CronLine $body
    if (-not $parsed) {
        # Plain comment = name/tag for the next cron; env lines (SHELL=, MAILTO=) are skipped
        if ($isComment -and $body -ne '') { $pending += $trim }
        # Starts like a cron but could not be parsed - needs a manual look.
        # For comments, require two schedule-like fields so tags such as "###### 2 AM jobs" aren't flagged.
        $looksLikeCron = $body -match '^[\d\*@]' -and $body -notmatch '^\w+\s*='
        if ($isComment) { $looksLikeCron = $body -match '^(@\w+|[\d\*][\d\*/,\-]*\s+[\d\*][\d\*/,\-]*)\s' }
        if ($looksLikeCron) {
            if ($isComment) { $lineTypes[$i + 1] = 'UNPARSED-COMMENTED' } else { $lineTypes[$i + 1] = 'UNPARSED' }
        } elseif ($isComment)       { $lineTypes[$i + 1] = 'COMMENT/TAG' }
        elseif ($body -match '^\w+\s*=') { $lineTypes[$i + 1] = 'ENV' }
        else                        { $lineTypes[$i + 1] = 'UNPARSED' }
        continue
    }
    if ($isComment) { $lineTypes[$i + 1] = 'CRON-COMMENTED' } else { $lineTypes[$i + 1] = 'CRON-ACTIVE' }

    $parsed | Add-Member -NotePropertyMembers @{
        LineNo    = $i + 1
        Commented = $isComment
        Raw       = $trim
        Tags      = $pending
    }
    $pending = @()
    $entries.Add($parsed)
}

# ---------------------------------------------------------------- evaluate
$fmt = 'HH:mm'
if ($Start.Date -ne $End.Date) { $fmt = 'dd-MMM HH:mm' }

$inWindow = @()
$commentedOutside = @()
$reboot = @()

foreach ($e in $entries) {
    if ($e.Reboot) { $reboot += $e; continue }
    $runs = Get-CronRunTimes $e $Start $End
    if ($runs.Count -gt 0) {
        $e | Add-Member -NotePropertyMembers @{ Runs = $runs; FirstRun = $runs[0] }
        $inWindow += $e
    } elseif ($e.Commented) {
        $commentedOutside += $e
    }
}
$inWindow = @($inWindow | Sort-Object FirstRun, LineNo)

# ---------------------------------------------------------------- output
$activeCount    = @($inWindow | Where-Object { -not $_.Commented }).Count
$commentedCount = @($inWindow | Where-Object { $_.Commented }).Count
$sep = '-' * 70

$out = New-Object 'System.Collections.Generic.List[string]'
$out.Add("MAINTENANCE WINDOW : $($Start.ToString('ddd dd-MMM-yyyy HH:mm'))  ->  $($End.ToString('ddd dd-MMM-yyyy HH:mm'))")
$out.Add("CRON FILE          : $((Resolve-Path -LiteralPath $CronFile).Path)")
$out.Add("CRONS IN WINDOW    : $activeCount active, $commentedCount commented (will not run)")
$out.Add($sep)

if ($inWindow.Count -eq 0) {
    $out.Add('No crons scheduled during the maintenance window.')
}

$n = 0
foreach ($e in $inWindow) {
    $n++
    $status = 'ACTIVE'
    if ($e.Commented) { $status = 'COMMENTED - will NOT run' }
    $out.Add("$n) $($e.FirstRun.ToString($fmt))  [$status]  (line $($e.LineNo))")
    foreach ($tag in $e.Tags) { $out.Add("   $tag") }
    $out.Add("   $($e.Raw)")
    $out.Add("   Runs: $(Format-RunTimes $e.Runs $fmt)")
    $out.Add('')
}

if ($commentedOutside.Count -gt 0) {
    $out.Add($sep)
    $out.Add('OTHER COMMENTED CRONS (outside window - for reference only)')
    $out.Add($sep)
    foreach ($e in $commentedOutside) {
        $out.Add("- (line $($e.LineNo)) COMMENTED")
        foreach ($tag in $e.Tags) { $out.Add("   $tag") }
        $out.Add("   $($e.Raw)")
        $out.Add('')
    }
}

if ($reboot.Count -gt 0) {
    $out.Add($sep)
    $out.Add('@REBOOT CRONS (run if the server is restarted during maintenance)')
    $out.Add($sep)
    foreach ($e in $reboot) {
        $state = 'ACTIVE'
        if ($e.Commented) { $state = 'COMMENTED' }
        $out.Add("- (line $($e.LineNo)) $state")
        foreach ($tag in $e.Tags) { $out.Add("   $tag") }
        $out.Add("   $($e.Raw)")
        $out.Add('')
    }
}

# ---------------------------------------------------------------- audit
if ($Audit) {
    $total   = $lines.Count
    $byType  = $lineTypes.Values | Group-Object | Sort-Object Name
    $sum     = ($byType | Measure-Object Count -Sum).Sum
    # Independent rough count: any non-comment line starting like a cron schedule
    $rawActive    = @($lines | Where-Object { $_ -match '^\s*[\d\*@]' }).Count
    $activeParsed = @($entries | Where-Object { -not $_.Commented }).Count

    $out.Add('')
    $out.Add('==================== AUDIT ====================')
    $out.Add(("Total lines in file        : {0}" -f $total))
    foreach ($g in $byType) { $out.Add(("  {0,-24} : {1}" -f $g.Name, $g.Count)) }
    if ($sum -eq $total) {
        $out.Add('Line reconciliation        : OK (every line classified)')
    } else {
        $out.Add("Line reconciliation        : MISMATCH ($sum classified vs $total lines)")
    }
    $out.Add(("Raw lines that look active : {0}" -f $rawActive))
    $out.Add(("Active crons parsed        : {0}" -f $activeParsed))
    if ($rawActive -eq $activeParsed) {
        $out.Add('Active cron cross-check    : OK')
    } else {
        $out.Add('Active cron cross-check    : MISMATCH - check UNPARSED lines below')
    }

    $bad = @($lineTypes.GetEnumerator() | Where-Object { $_.Value -like 'UNPARSED*' } | Sort-Object Name)
    $out.Add('')
    if ($bad.Count -gt 0) {
        $out.Add('Lines NOT understood (check manually):')
        foreach ($b in $bad) { $out.Add(("  line {0,4} [{1}] {2}" -f $b.Name, $b.Value, $lines[$b.Name - 1].Trim())) }
    } else {
        $out.Add('Lines NOT understood       : none')
    }

    # Every cron with its result, for spot-checking
    $rows = foreach ($e in $entries) {
        $runs = $null
        if (-not $e.Reboot) { $runs = Get-CronRunTimes $e $Start $End }
        $count = 0
        if ($runs) { $count = $runs.Count }
        $status = 'ACTIVE'
        if ($e.Commented) { $status = 'COMMENTED' }
        $first = ''; $all = ''
        if ($count -gt 0) {
            $first = $runs[0].ToString($fmt)
            $all = Format-RunTimes $runs $fmt
        }
        $inWin = 'NO'
        if ($e.Reboot) { $inWin = 'REBOOT' } elseif ($count -gt 0) { $inWin = 'YES' }
        [pscustomobject]@{
            LineNo   = $e.LineNo
            Status   = $status
            InWindow = $inWin
            FirstRun = $first
            RunCount = $count
            Runs     = $all
            Name     = ($e.Tags -join ' | ')
            Cron     = $e.Raw
        }
    }
    $out.Add('')
    $out.Add("ALL $($entries.Count) CRONS:")
    # Wide fixed width so long commands are not cut off in the file
    $table = $rows | Format-Table LineNo, Status, InWindow, FirstRun, RunCount, Name, Cron -AutoSize | Out-String -Width 4096
    foreach ($l in ($table.Trim("`r", "`n") -split "`r?`n")) { $out.Add($l.TrimEnd()) }
}

# ---------------------------------------------------------------- print + save
$text = $out -join [Environment]::NewLine
Write-Output $text

if (-not $OutFile) { $OutFile = Join-Path (Get-Location) ("maintenance_crons_{0}.txt" -f $Start.ToString('yyyyMMdd')) }
if (-not $ScreenOnly) {
    $text | Out-File -LiteralPath $OutFile -Encoding utf8
    Write-Host ""
    Write-Host "Saved to: $OutFile" -ForegroundColor Green
    if ($Csv -and $Audit) {
        $csvFile = [IO.Path]::ChangeExtension($OutFile, '.audit.csv')
        $rows | Export-Csv -LiteralPath $csvFile -NoTypeInformation -Encoding UTF8
        Write-Host "Audit CSV: $csvFile" -ForegroundColor Green
    }
}

if (-not $NoClipboard -and -not $ScreenOnly) {
    try { $text | Set-Clipboard; Write-Host 'Copied to clipboard.' -ForegroundColor Green } catch { }
}
