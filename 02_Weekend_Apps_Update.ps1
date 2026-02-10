# ─────────────────────────────────────────────────────────────────────────────
# Prep: TLS + connection limits
# ─────────────────────────────────────────────────────────────────────────────
[Net.ServicePointManager]::SecurityProtocol = `
    [Net.SecurityProtocolType]::Tls12 -bor `
    [Net.SecurityProtocolType]::Tls13
[Net.ServicePointManager]::DefaultConnectionLimit = 64

# ─────────────────────────────────────────────────────────────────────────────
# Winget source configuration
# ─────────────────────────────────────────────────────────────────────────────
$sourceList = winget source list
if ($sourceList -notmatch "msstore") {
    winget source add --name msstore --arg https://storeedgefd.dsx.mp.microsoft.com/v9.0 --accept-source-agreements
}
winget source update

# ─────────────────────────────────────────────────────────────────────────────
# Perform Winget Upgrades
# ─────────────────────────────────────────────────────────────────────────────
winget upgrade --all --silent --accept-source-agreements --accept-package-agreements --disable-interactivity --verbose-logs --include-unknown --include-pinned --force --scope machine
winget upgrade --all --silent --accept-source-agreements --accept-package-agreements --disable-interactivity --verbose-logs --include-unknown --include-pinned --force --scope user

# ─────────────────────────────────────────────────────────────────────────────
# Update Office if Installed
# ─────────────────────────────────────────────────────────────────────────────
$officePath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"
if (Test-Path $officePath) {
    Start-Process $officePath -ArgumentList "/update USER", "displaylevel=True"
}

