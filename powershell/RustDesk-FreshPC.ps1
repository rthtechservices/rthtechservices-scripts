<# 
RTH RustDesk One-Click Installer
- Uninstalls any existing RustDesk & service
- Installs latest RustDesk with winget
- Installs/starts Windows Service (runs before login)
- Applies server config (Host/Relay/API/Key) via --config
- Sets a permanent password
- Sets device ID = sanitized hostname (must start with letter)
- Creates Scheduled Tasks:
    * Daily 03:00 restart RustDesk service
    * Weekly 03:10 silent upgrade via winget
#>

# ------------------------- SCRIPT PARAMETERS  -------------------------

# 1) Encrypted Configuration String from RTH Rustdesk Server
     $ConfigString = '0nIt92YuMXZjlmdyV2coNWZ0hGdy5CZy9yL6MHc0RHaiojIpBXYiwiI9UUS3s0cYN2SrdTU39EePdFeBZ1c5ZzbIFzK08WMNtkZyczYHZmNmVGdhBjI6ISeltmIsISbvNmLzV2YpZnclNHajVGdoRncuQmciojI0N3boJye'

# 2) Encrypted Master Password to be set on Client.
     $b64 = 'RnVudGltZVJ1c3RkZXNrMSE='

# 3) RTH Rustdesk Server Public Key for client configuration
     $ServerKey = '0atef6fGc72fKM1o4+1Ho6ysVAxWOxOwQ7kKcXsK7IE='

# 4) Daily Scheduled Task to restart Rustdesk Service to keep service running optimally.
     $DailyRestartTime = '03:00'
     #Weekly Schedule Task to check internet for updates to client. 
     $WeeklyUpdateDay = 'Sunday'
     $WeeklyUpdateTime = '03:10'  

#5)  Error Action Preference for script
     $ErrorActionPreference = 'Stop'
     $ErrorActionPreference = 'SilentlyContinue'
# --------------------------------------------------------------
     
# ------------------------- SCRIPT FUNCTIONS  -------------------------
function fn_Log($m){ Write-Host "🠊 $m" -ForegroundColor Cyan }
function fn_Ok($m){ Write-Host "✔ $m" -ForegroundColor Green }
function fn_Warn($m){ Write-Host "✱ $m" -ForegroundColor Yellow }
function fn_Err($m){ Write-Host "❌ $m" -ForegroundColor Red }
function fn_Blank{Write-Host " "}

function fn_Assert-Admin {
  fn_Log "Checking if current context is running as Admin"
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = New-Object Security.Principal.WindowsPrincipal($id)
  if (-not $p.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    fn_Err "Please run this script in an elevated PowerShell (Run as Administrator)." -ErrorAction Stop
    fn_Blank
    Write-Host "Press Enter key to exit" -ForegroundColor Red 
    Read-Host
    exit 1
  }
}

function fn_Stop-Remove-Service {
  param([string]$Name)
  try {
     $svc = Get-Service -Name $Name -ErrorAction Stop
    if ($svc.Status -ne 'Stopped') { 
        fn_Warn "Service found running as:'$Name'. Stopping $Name and removing it"
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue }
    # Delete the service entry
    sc.exe delete "$Name" | Out-Null
    Start-Sleep -Seconds 2
  } catch {}
}

function fn_Uninstall-RustDesk {
  fn_Log "Uninstalling any existing RustDesk…"
  # 1) Stop/remove service if present
  fn_Stop-Remove-Service -Name 'RustDesk Service'   # common display name
  fn_Stop-Remove-Service -Name 'RustDesk'           # alternate service name, just in case

  # 2) Winget uninstall if it was installed that way
  try {
    fn_log "Attemping to remove using the Uninstaller"
    $rd = winget list --id RustDesk.RustDesk -e 2>$null
    if ($LASTEXITCODE -eq 0 -and $rd) {
      winget uninstall --id RustDesk.RustDesk -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
    }
  } catch {}

  # 3) Best-effort cleanup of leftover folders
 fn_log "Looking for files in User Profile or Program Files"
  $paths = @(
    "$env:ProgramFiles\RustDesk",
    "$env:ProgramData\RustDesk",
    "$env:LOCALAPPDATA\RustDesk",
    "$env:APPDATA\RustDesk"
  )
  foreach ($p in $paths) {
   if (Test-Path $p) { try { Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue } catch {} } }
}

function fn_Install-Latest-RustDesk {
  fn_Log "Installing latest RustDesk via winget…"
  # Winget install (silent)
  winget install --id RustDesk.RustDesk -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
}

function fn_Install-As-Service {
  $exe = Join-Path $env:ProgramFiles 'RustDesk\rustdesk.exe'
  if (-not (Test-Path $exe)) {
    fn_warn "Rustdesk executable not found at $exe"
    throw  "RustDesk executable not found at $exe"
  }
  fn_log "Installing Windows service…"
  # Install service so it starts before login
  Start-Process -FilePath $exe -ArgumentList '--install-service'   # requires admin
  Start-Sleep -Seconds 15
  # Make sure it's running
  try {
    fn_log "Checking if service is running"
    Start-Service -Name 'RustDesk Service' -ErrorAction SilentlyContinue
  } catch {}
  return $exe
}

function fn_Apply-Config {
  param([string]$Exe, [string]$CfgString)
  if ([string]::IsNullOrWhiteSpace($CfgString)) {
    fn_warn "No ConfigString provided. The client may still work if only ID Server is set manually later, but to push Host/API/Key centrally please paste the encrypted config string from your console into `$ConfigString`."
    return
  }
  fn_log "Applying server configuration…"
  # RustDesk CLI prints nothing by default—piping forces output; failures raise $LASTEXITCODE
  $null = Start-Process -FilePath $Exe -ArgumentList @('--config', $CfgString)  -PassThru  
  start-sleep -Seconds 10}

function fn_Set-MasterPassword {
  param([string]$Exe, [string]$Pwd)
  if ([string]::IsNullOrWhiteSpace($Pwd)) { return }
  fn_log "Setting permanent password…"
  $null = Start-Process -FilePath $Exe -ArgumentList @('--password', $Pwd) -PassThru
  start-sleep -seconds 8
}

function fn_Set-Id-To-Hostname {
  param([string]$Exe)
  # IDs must start with a letter per docs; sanitize non-alphanumerics
  $hn = $env:COMPUTERNAME
  $id = ($hn -replace '[^A-Za-z0-9\-]', '')
  if ($id.Length -eq 0 -or $id[0] -notmatch '[A-Za-z]') { $id = 'R-' + $id }
  fn_log "Setting Device ID to '$id'…"
  $null = Start-Process -FilePath $Exe -ArgumentList @('--set-id', $id) -PassThru
  start-sleep -Seconds 5
  }

function fn_Create-Scheduled-Tasks {
  # Daily restart at 3:00
  fn_log "Creating Daily Restart Scheduled Task"
  $action1   = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -WindowStyle Hidden -Command "Restart-Service -Name ''RustDesk Service'' -Force"'
  $trigger1  = New-ScheduledTaskTrigger -Daily -At ([DateTime]::Parse($DailyRestartTime))
  $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
  Register-ScheduledTask -TaskName 'RustDesk-Daily-Restart' -Action $action1 -Trigger $trigger1 -Principal $principal -Force | Out-Null

  # Weekly update via winget (silent)
  # We drop a tiny updater script to ProgramData and call it as SYSTEM
  fn_log "Creating Weekly Scheduled Task to update client software"
  $updDir = 'C:\ProgramData\RustDesk'
  New-Item -ItemType Directory -Force -Path $updDir | Out-Null
  $updScript = Join-Path $updDir 'Update-RustDesk.ps1'
  @'
try {
  winget upgrade --id RustDesk.RustDesk -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
} catch {}
'@ | Set-Content -Path $updScript -Encoding UTF8 -Force

  $action2  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\ProgramData\RustDesk\Update-RustDesk.ps1"'
  $trigger2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $WeeklyUpdateDay -At ([DateTime]::Parse($WeeklyUpdateTime))
  Register-ScheduledTask -TaskName 'RustDesk-Weekly-Update' -Action $action2 -Trigger $trigger2 -Principal $principal -Force | Out-Null
}


# ------------------------- RUN -------------------------
$date = get-date -format g
CLS
Write-Host "__________________ RTH TECH SERVICES INC. __________________" -ForegroundColor cyan
Write-Host "                                                            "
Write-Host "                         ROHAN HARE                         "-ForegroundColor Yellow
Write-Host "                  rohan@rthtechservices.com                 "-ForegroundColor Yellow
Write-Host "                                                            "
Write-Host "____________________________________________________________"-ForegroundColor Cyan
Write-Host "                                                            "
Write-Host "  This script will first remove any existing installations  " -ForegroundColor Green
Write-Host "  of the Rustdesk client, before downloading the latest     "-ForegroundColor Green
Write-Host "  version found online, installing it, and cofiguring       " -ForegroundColor Green
Write-Host "  it to point to the RTH Rustdesk Server, along with        "-ForegroundColor Green
Write-Host "  configuring the client with RTH policies and settings.    "-ForegroundColor Green
Write-Host "                                                            "
Write-Host "____________________________________________________________"-ForegroundColor Cyan
Write-Host "                                                            "
Write-Host "                     $date                                  "-ForegroundColor Yellow
Write-Host "               The script is ready to deploy.               "-ForegroundColor Magenta
Write-Host "                    Press Enter to start                    "-ForegroundColor Magenta
Write-Host "                                                            "
Write-Host "____________________________________________________________"-ForegroundColor Cyan
Read-Host
$byte = [System.Convert]::FromBase64String($b64)
$decode = [System.Text.Encoding]::UTF8.GetString($byte)
fn_Assert-Admin
#fn_Uninstall-RustDesk
fn_Install-Latest-RustDesk
$exe = fn_Install-As-Service
fn_Apply-Config -Exe $exe -CfgString $ConfigString
fn_Set-MasterPassword -Exe $exe -Pwd $decode
fn_Set-Id-To-Hostname -Exe $exe

# Restart service to apply settings cleanly
try { Restart-Service -Name 'RustDesk Service' -Force -ErrorAction SilentlyContinue } catch {}

fn_Create-Scheduled-Tasks

Write-Host "`nAll done. RustDesk is installed as a service, configured, and maintenance tasks were created."
Write-Host "Press any key to exit"
Pause
exit
