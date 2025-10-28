# ==== Verification: Winget, ChatGPT, Cursor, Python 3.12, LibreOffice ====
$ErrorActionPreference = 'SilentlyContinue'

function Get-Py312Cmd {
  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py) { return @('py','-3.12') }
  $python312 = (Get-Command python3.12 -ErrorAction SilentlyContinue).Source
  $candidates = @(
    $python312,
    "C:\Program Files\Python312\python.exe",
    "C:\Program Files (x86)\Python312\python.exe"
  ) | Where-Object { $_ -and (Test-Path $_) }
  if ($candidates) { return @($candidates[0]) }
  # Uncomment only if needed (can be slow):
  # $wa = Get-ChildItem "C:\Program Files\WindowsApps" -Filter "PythonSoftwareFoundation.Python.3.12*" -ErrorAction SilentlyContinue |
  #       ForEach-Object { Join-Path $_.FullName "python.exe" } | Where-Object { Test-Path $_ } | Select-Object -First 1
  # if ($wa) { return @($wa) }
  return $null
}

# --- Winget ---
$wingetCmd     = Get-Command winget -ErrorAction SilentlyContinue
$wingetVersion = $null
$wingetDetail  = 'winget not on PATH'
if ($wingetCmd) {
  try { $wingetVersion = (& winget --version) -join ' ' } catch {}
  $wingetDetail = $wingetCmd.Source
}

# --- ChatGPT (Store app OR PWA shortcuts) ---
$chatgptAppx = Get-AppxPackage -Name "*OpenAI*" -ErrorAction SilentlyContinue
$startDir = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
$chatgptLnk = Join-Path $startDir "ChatGPT.lnk"
$chatgptUrl = Join-Path $startDir "ChatGPT.url"
$chatgptStatus = if ($chatgptAppx) {
  'StoreApp'
} elseif (Test-Path $chatgptLnk) {
  'PWA-Edge-Shortcut'
} elseif (Test-Path $chatgptUrl) {
  'URL-Shortcut'
} else {
  'NotFound'
}

# --- Cursor ---
$cursorPath = @(
  "$env:LOCALAPPDATA\Programs\Cursor\Cursor.exe",
  "$env:ProgramFiles\Cursor\Cursor.exe",
  "$env:ProgramFiles(x86)\Cursor\Cursor.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
$cursorPresent = [bool]$cursorPath
$cursorDetail = if ($cursorPresent) { $cursorPath } else { 'NotFound' }

# --- Python 3.12 + packages (fast spec check) ---
$pycmd     = Get-Py312Cmd
$pyVersion = $null
$pandasOK = $false
$openpyxlOK = $false
if ($pycmd) {
  try { $pyVersion = (& $pycmd '--version') -join ' ' } catch {}
  try {
    $o = & $pycmd '-c' "import importlib.util as u; print('1' if u.find_spec('pandas') else '0')"
    $pandasOK = ($o -eq '1')
  } catch {}
  try {
    $o = & $pycmd '-c' "import importlib.util as u; print('1' if u.find_spec('openpyxl') else '0')"
    $openpyxlOK = ($o -eq '1')
  } catch {}
}
$pyDetail = if ($pycmd) { ($pycmd -join ' ') } else { 'NotFound' }
$pandasDetail = if ($pandasOK) { 'installed' } else { 'missing' }
$openpyxlDetail = if ($openpyxlOK) { 'installed' } else { 'missing' }

# --- LibreOffice (no-launch, no prompts) ---
$soffice = @(
  "$env:ProgramFiles\LibreOffice\program\soffice.exe",
  "$env:ProgramFiles(x86)\LibreOffice\program\soffice.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
$loPresent = [bool]$soffice
$loVersion = $null
$loDetail  = 'NotFound'
if ($soffice) {
  try {
    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($soffice)
    if ($vi -and $vi.ProductVersion) { $loVersion = $vi.ProductVersion } elseif ($vi -and $vi.FileVersion) { $loVersion = $vi.FileVersion }
  } catch {}
  $loDetail = $soffice
}

# --- Output summary ---
$rows = @()
$rows += [pscustomobject]@{ Item='Winget';        Present=[bool]$wingetCmd; Version=$wingetVersion; Detail=$wingetDetail }
$rows += [pscustomobject]@{ Item='ChatGPT';       Present=($chatgptStatus -ne 'NotFound'); Version=$null; Detail=$chatgptStatus }
$rows += [pscustomobject]@{ Item='Cursor';        Present=$cursorPresent; Version=$null; Detail=$cursorDetail }
$rows += [pscustomobject]@{ Item='Python 3.12';   Present=[bool]$pycmd; Version=$pyVersion; Detail=$pyDetail }
$rows += [pscustomobject]@{ Item='  └─ pandas';   Present=$pandasOK; Version=$null; Detail=$pandasDetail }
$rows += [pscustomobject]@{ Item='  └─ openpyxl'; Present=$openpyxlOK; Version=$null; Detail=$openpyxlDetail }
$rows += [pscustomobject]@{ Item='LibreOffice';   Present=$loPresent; Version=$loVersion; Detail=$loDetail }

$rows | Format-Table -AutoSize

# Optional quick extras
Write-Host "`nWinget sources:" -ForegroundColor Cyan
try { winget source list --disable-interactivity } catch {}
Write-Host "`nNode/Git quick versions (optional):" -ForegroundColor Cyan
try { (node -v) } catch {}
try { (git --version) } catch {}

Write-Host "`nDone." -ForegroundColor Green