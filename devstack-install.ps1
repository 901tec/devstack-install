# WATCHABLE INSTALL-ONLY (run in elevated PowerShell)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::UTF8

$logRoot = "$env:ProgramData\DevStackDeploy"
$log = Join-Path $logRoot "install.log"
New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
"=== Dev Stack (WATCH / install-only) $(Get-Date -Format s) ===" | Out-File $log -Encoding utf8

# ---------- helpers ----------
function Log {
  param([string]$m)
  $m | Tee-Object -FilePath $log -Append
}
function Invoke-WinGet {
  param([string[]]$Tokens)
  $echo = $Tokens -join ' '
  Log "`n> winget $echo"
  & winget @Tokens --accept-package-agreements --accept-source-agreements --disable-interactivity
  if ($LASTEXITCODE -ne 0) { throw "WinGet failed: $echo (code $LASTEXITCODE)" }
}

# presence checks that avoid slow winget calls
function Is-AppxChatGPT {
  try {
    $pkg = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object {
      $_.Name -like '*OpenAI*' -or $_.PackageFamilyName -like '*OpenAI*'
    }
    return [bool]$pkg
  } catch { return $false }
}
function Has-Shortcut { param([Parameter(Mandatory)][string]$name) return (Test-Path -LiteralPath "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\$name") }
function Is-GitPresent  { return ($null -ne (Get-Command git  -ErrorAction SilentlyContinue)) }
function Is-NodePresent { return ($null -ne (Get-Command node -ErrorAction SilentlyContinue)) }
function Is-CursorPresent {
  $paths = @(
    "$env:LOCALAPPDATA\Programs\Cursor\Cursor.exe",
    "$env:ProgramFiles\Cursor\Cursor.exe",
    "$env:ProgramFiles(x86)\Cursor\Cursor.exe"
  )
  foreach ($p in $paths) { if (Test-Path -LiteralPath $p) { return $true } }
  return $false
}
function Is-LibreOfficePresent {
  $paths = @(
    "$env:ProgramFiles\LibreOffice\program\soffice.exe",
    "$env:ProgramFiles(x86)\LibreOffice\program\soffice.exe"
  )
  foreach ($p in $paths) { if (Test-Path -LiteralPath $p) { return $true } }
  return $false
}
function Get-Py312Cmd {
  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py) { return @('py','-3.12') }
  $candidates = @(
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files (x86)\Python312\python.exe"
  )
  try {
    $candidates += Get-ChildItem "C:\Program Files\WindowsApps" -Filter "PythonSoftwareFoundation.Python.3.12*" -ErrorAction SilentlyContinue |
      ForEach-Object { Join-Path $_.FullName "python.exe" }
  } catch {}
  foreach ($p in $candidates) { if ($p -and (Test-Path -LiteralPath $p)) { return @($p) } }
  return $null
}
function Is-Python312Present { return ($null -ne (Get-Py312Cmd)) }

# ---------- ChatGPT: best → fallback ----------
function Install-ChatGPT-BestEffort {
  if (Is-AppxChatGPT) { Log "> ChatGPT Store app present — skipping."; return }

  # If a Start Menu shortcut already exists, skip all ChatGPT work
  if (Has-Shortcut 'ChatGPT.lnk' -or Has-Shortcut 'ChatGPT.url') { Log "> ChatGPT shortcut present — skipping."; return }

  # Try via winget/msstore (official Store app)
  try {
    # Ensure msstore source is available (ignore failure if Store blocked)
    try {
      $hasMsStore = (& winget source list --disable-interactivity | Select-String -SimpleMatch 'msstore' -ErrorAction SilentlyContinue)
    } catch { $hasMsStore = $null }
    if (-not $hasMsStore) {
      Log "> Enabling winget source: msstore"
      & winget source enable msstore --disable-interactivity | Out-Null
    }
    Log "> Attempting ChatGPT via winget/msstore..."
    Invoke-WinGet @('install','-e','--id','9NT1R1C2HH7J','--source','msstore')
    Log "> ChatGPT installed from msstore."
    return
  } catch {
    Log "> ChatGPT via msstore failed or blocked. Falling back to PWA."
  }

  # Fallback: Edge app-mode shortcuts (PWA)
  $edgePaths = @(
    "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
  )
  $edge = $null
  foreach ($p in $edgePaths) { if (Test-Path -LiteralPath $p) { $edge = $p; break } }
  $startDir = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
  $deskDir  = "$env:Public\Desktop"
  foreach ($d in @($startDir, $deskDir)) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }

  if ($edge) {
    $args = '--no-first-run --no-default-browser-check --profile-directory=Default --app="https://chatgpt.com"'
    $wsh = New-Object -ComObject WScript.Shell
    foreach ($dest in @($startDir, $deskDir)) {
      $lnk = Join-Path $dest "ChatGPT.lnk"
      $sc = $wsh.CreateShortcut($lnk)
      $sc.TargetPath   = $edge
      $sc.Arguments    = $args
      $sc.IconLocation = "$edge,0"
      $sc.WorkingDirectory = Split-Path $edge
      $sc.WindowStyle  = 1
      $sc.Description  = "ChatGPT"
      $sc.Save()
      Log "Created ChatGPT app shortcut: $lnk"
    }
  } else {
    foreach ($dest in @($startDir, $deskDir)) {
      $urlPath = Join-Path $dest "ChatGPT.url"
      @(
        "[InternetShortcut]",
        "URL=https://chatgpt.com",
        "IconFile=%SystemRoot%\system32\SHELL32.dll",
        "IconIndex=220"
      ) | Out-File -FilePath $urlPath -Encoding ASCII
      Log "Created ChatGPT URL: $urlPath"
    }
  }
}

# ---------- Run ChatGPT best→fallback, then other apps via winget ----------
Install-ChatGPT-BestEffort

if (-not (Is-GitPresent))         { Invoke-WinGet @('install','-e','--id','Git.Git') }                           else { Log "Git present — skipping." }
if (-not (Is-NodePresent))        { Invoke-WinGet @('install','-e','--id','OpenJS.NodeJS.LTS') }                 else { Log "Node present — skipping." }
if (-not (Is-Python312Present))   { Invoke-WinGet @('install','-e','--id','Python.Python.3.12','--scope','machine') } else { Log "Python 3.12 present — skipping." }
if (-not (Is-CursorPresent))      { Invoke-WinGet @('install','-e','--id','Anysphere.Cursor') }                  else { Log "Cursor present — skipping." }
if (-not (Is-LibreOfficePresent)) { Invoke-WinGet @('install','-e','--id','TheDocumentFoundation.LibreOffice') } else { Log "LibreOffice present — skipping." }

# Refresh PATH in current session (PowerShell 5.1 safe)
try {
  $machinePath = [System.Environment]::GetEnvironmentVariable('Path','Machine')
  $userPath    = [System.Environment]::GetEnvironmentVariable('Path','User')
  $parts = @()
  if ($machinePath) { $parts += $machinePath }
  if ($userPath)    { $parts += $userPath }
  if ($parts.Count) { $env:Path = ($parts -join ';') }
} catch {}

# ---------- Python 3.12 pip installs (no REPL) ----------
function Get-Python312Exe {
  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py) {
    try {
      $out = & py -3.12 -c "import sys; print(sys.executable)" 2>$null
      if ($out -and (Test-Path -LiteralPath $out)) { return $out }
    } catch {}
  }
  $candidates = @(
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files (x86)\Python312\python.exe"
  )
  try {
    $candidates += Get-ChildItem "C:\Program Files\WindowsApps" -Filter "PythonSoftwareFoundation.Python.3.12*" -ErrorAction SilentlyContinue |
      ForEach-Object { Join-Path $_.FullName "python.exe" }
  } catch {}
  foreach ($p in $candidates) { if ($p -and (Test-Path -LiteralPath $p)) { return $p } }
  return $null
}

$pyExe = Get-Python312Exe
if ($pyExe) {
  Write-Host "`n> Python: upgrade pip + install pandas, openpyxl" -ForegroundColor Cyan

  $pipUpgrade = @('-m','pip','install','--upgrade','pip','--disable-pip-version-check')
  Start-Process -FilePath $pyExe -ArgumentList $pipUpgrade -NoNewWindow -Wait

  $pipPkgs = @('-m','pip','install','--disable-pip-version-check','--no-input','pandas','openpyxl')
  Start-Process -FilePath $pyExe -ArgumentList $pipPkgs -NoNewWindow -Wait
}

# ---------- Versions ----------
Write-Host "`n=== Versions ===" -ForegroundColor Green
try { git --version } catch {}
try { node -v } catch {}
try {
  $npm = Get-Command npm -ErrorAction SilentlyContinue
  if ($npm) { npm -v } else { Write-Host "(npm not found; Node 22+ ships without npm. Use LTS or corepack.)" -ForegroundColor Yellow }
} catch {}
try {
  if ($pyExe) {
    & $pyExe --version
    & $pyExe -m pip --version
  }
} catch {}

Write-Host "`nDone. Log: $log" -ForegroundColor Green