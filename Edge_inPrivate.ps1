# =====================================================================
# Launch Microsoft Edge InPrivate to www.compton.edu
# Runs for every user that logs onto the computer
# Must be run once as Administrator
# =====================================================================

# Verify running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] `
    [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host "ERROR: Please run this script as Administrator." -ForegroundColor Red
    exit 1
}

# Locate Microsoft Edge
$Edge = Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"

if (!(Test-Path $Edge)) {
    $Edge = Join-Path ${env:ProgramFiles} "Microsoft\Edge\Application\msedge.exe"
}

if (!(Test-Path $Edge)) {
    Write-Host "Microsoft Edge was not found." -ForegroundColor Red
    exit 1
}

# Registry location
$RunKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"

# Command to execute
$Command = "`"$Edge`" --inprivate https://www.compton.edu"

# Create or update the Run entry
New-ItemProperty `
    -Path $RunKey `
    -Name "LaunchComptonEdge" `
    -Value $Command `
    -PropertyType String `
    -Force | Out-Null

Write-Host ""
Write-Host "SUCCESS!" -ForegroundColor Green
Write-Host "Microsoft Edge will now launch automatically for every user at logon."
Write-Host ""
Write-Host "Command:"
Write-Host $Command