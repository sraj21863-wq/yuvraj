# OB Stage 2 - unattended setup.
#
# Finds OB_Stage_2.zip wherever the browser put it, extracts it to C:\OB, and
# starts the dashboard. Safe to run more than once: a workbook that is already
# there is never overwritten, because by then it holds pasted data the zip
# does not have.

$ErrorActionPreference = 'Stop'
$dest   = 'C:\OB'
$folder = Join-Path $dest 'OB Stage 2'
$wbName = 'Order_Booking_Master_STAGE2.xlsx'

function Say($t, $c) { if ($c) { Write-Host $t -ForegroundColor $c } else { Write-Host $t } }
function Stop-Here($code) { Say ''; Read-Host '  Press Enter to close' | Out-Null; exit $code }

Say ''
Say '  OB STAGE 2 - SETUP' 'Cyan'
Say '  =================='
Say ''

# ---- 1. find the zip ------------------------------------------------------
# Named folders one at a time rather than a recursive sweep of the whole
# profile: AppData alone would take longer than the rest of this script.
$spots = New-Object System.Collections.ArrayList
foreach ($base in @($env:USERPROFILE, $env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer)) {
  if (-not $base) { continue }
  foreach ($leaf in @('Downloads', 'Desktop', 'Documents')) {
    [void]$spots.Add((Join-Path $base $leaf))
  }
}
[void]$spots.Add($env:USERPROFILE)
[void]$spots.Add((Get-Location).Path)

Say '  Looking for OB_Stage_2.zip ...'
$zip = $null
foreach ($d in ($spots | Select-Object -Unique)) {
  if (-not (Test-Path -LiteralPath $d)) { continue }
  $hit = Get-ChildItem -LiteralPath $d -Filter 'OB_Stage_2*.zip' -File -ErrorAction SilentlyContinue |
         Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($hit) { $zip = $hit; break }
}

if (-not $zip) {
  Say ''
  Say '  Could not find OB_Stage_2.zip.' 'Red'
  Say ''
  Say '  I looked in Downloads, Desktop and Documents, including the OneDrive'
  Say '  versions of those. Download the zip from the chat - it saves to'
  Say '  Downloads - and run this file again.'
  Stop-Here 1
}
Say ('  found    ' + $zip.FullName) 'Green'
Say ('  size     ' + [math]::Round($zip.Length / 1MB, 1) + ' MB')

# Windows marks downloaded files as blocked, which makes the .cmd launchers
# refuse to run later with no useful message.
try { Unblock-File -LiteralPath $zip.FullName -ErrorAction SilentlyContinue } catch { }

# ---- 2. extract to a staging folder --------------------------------------
$tmp = Join-Path $env:TEMP ('obstage2_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
Say ''
Say '  Extracting ...'
try {
  Expand-Archive -LiteralPath $zip.FullName -DestinationPath $tmp -Force
} catch {
  Say ''
  Say ('  Could not extract it: ' + $_.Exception.Message) 'Red'
  Say ''
  Say '  Extract it by hand instead: right-click the zip, Extract All, and set'
  Say '  the path to  C:\OB'
  Stop-Here 1
}

$src = Join-Path $tmp 'OB Stage 2'
if (-not (Test-Path -LiteralPath $src)) { $src = $tmp }

# ---- 3. copy into place, protecting a workbook that already has data ------
$wb = Join-Path $folder $wbName
$hadWorkbook = Test-Path -LiteralPath $wb

if (-not (Test-Path -LiteralPath $folder)) {
  New-Item -ItemType Directory -Path $folder -Force | Out-Null
}

if ($hadWorkbook) {
  $bk = Join-Path $folder 'backups'
  New-Item -ItemType Directory -Path $bk -Force | Out-Null
  $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
  Copy-Item -LiteralPath $wb -Force `
            -Destination (Join-Path $bk ('Order_Booking_Master_STAGE2--yours-' + $stamp + '.xlsx'))
}

Get-ChildItem -LiteralPath $src -Force | ForEach-Object {
  if ($hadWorkbook -and $_.Name -eq $wbName) { return }      # yours stays put
  Copy-Item -LiteralPath $_.FullName -Destination $folder -Recurse -Force
}
Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue

# ---- 4. check what landed ------------------------------------------------
$missing = @()
foreach ($n in @($wbName, 'pipeline\run_service.cmd', 'dashboard\Order_Booking_Dashboard.html')) {
  if (-not (Test-Path -LiteralPath (Join-Path $folder $n))) { $missing += $n }
}
if ($missing.Count -gt 0) {
  Say ''
  Say ('  Missing after extracting: ' + ($missing -join ', ')) 'Red'
  Say '  Re-download the zip and run this again.'
  Stop-Here 1
}

Say ''
Say ('  Folder ready:  ' + $folder) 'Green'
if ($hadWorkbook) {
  Say ''
  Say '  You already had a workbook here, so it was LEFT ALONE - it holds data' 'Yellow'
  Say '  the zip does not. Everything else was updated. A dated copy of yours'  'Yellow'
  Say '  is in  backups\  in case you wanted the fresh one instead.'            'Yellow'
}

Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $folder + '"')

# ---- 5. Python, then launch ----------------------------------------------
$havePy = $false
foreach ($c in @('py', 'python')) {
  $exe = Get-Command $c -ErrorAction SilentlyContinue
  if (-not $exe) { continue }
  try {
    & $c -c 'import sys' 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { $havePy = $true; break }
  } catch { }
}

Say ''
if (-not $havePy) {
  Say '  Python is not installed yet, so the live dashboard cannot start.' 'Yellow'
  Say ''
  Say '  Install it with this one line, in this window:'
  Say ''
  Say '      winget install -e --id Python.Python.3.12' 'Cyan'
  Say ''
  Say '  Then CLOSE this window, open a new one, and run this setup file again.'
  Say '  A new window is required - this one cannot see a program installed'
  Say '  after it opened.'
  Say ''
  Say '  Meanwhile you can look at the dashboard with no Python at all:'
  Say ('      ' + (Join-Path $folder 'dashboard\Order_Booking_Dashboard_standalone.html'))
  Stop-Here 0
}

Say '  Python found. Starting the dashboard ...' 'Green'
Say ''
Say '  A second window opens and stays open - that is the data service.'
Say '  Your browser goes to  http://127.0.0.1:8787/  and the banner turns green.'
Say '  Closing that window stops the live connection to Excel.'
Start-Process -FilePath (Join-Path $folder 'pipeline\run_service.cmd') `
              -WorkingDirectory (Join-Path $folder 'pipeline')
Start-Sleep -Seconds 2
Stop-Here 0
