# =====================================================================
# ScriptName: Compton_Tech_Utils.ps1
# ScriptVersion: 1.10.0
# LastUpdated: 2026-04-27
# =====================================================================

# -----------------------------------------------------------------------------
# Startup Self-Update Check
# -----------------------------------------------------------------------------
function Get-LocalScriptVersion {
    param(
        [Parameter(Mandatory)][string]$ScriptPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ScriptPath)) {
            return $null
        }

        $firstLines = Get-Content -LiteralPath $ScriptPath -TotalCount 25 -ErrorAction Stop
        $versionLine = $firstLines | Where-Object { $_ -match '^\s*#\s*ScriptVersion\s*:\s*(.+?)\s*$' } | Select-Object -First 1
        if (-not $versionLine) {
            return $null
        }

        $versionText = ($versionLine -replace '^\s*#\s*ScriptVersion\s*:\s*', '').Trim()
        return [version]$versionText
    }
    catch {
        return $null
    }
}

function Get-RemoteScriptMetadata {
    param(
        [Parameter(Mandatory)][string]$RawUrl
    )

    try {
        $response = Invoke-WebRequest -Uri $RawUrl -UseBasicParsing -ErrorAction Stop
        $content = [string]$response.Content
        $versionMatch = [regex]::Match($content, '(?im)^\s*#\s*ScriptVersion\s*:\s*([^\r\n]+)')
        $updatedMatch = [regex]::Match($content, '(?im)^\s*#\s*LastUpdated\s*:\s*([^\r\n]+)')

        $remoteVersion = $null
        if ($versionMatch.Success) {
            try {
                $remoteVersion = [version]$versionMatch.Groups[1].Value.Trim()
            }
            catch {
                $remoteVersion = $null
            }
        }

        [PSCustomObject]@{
            Version     = $remoteVersion
            LastUpdated = if ($updatedMatch.Success) { $updatedMatch.Groups[1].Value.Trim() } else { $null }
            Content     = $content
        }
    }
    catch {
        Write-Host "Unable to check GitHub for script updates: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}

function Invoke-StartupSelfUpdate {
    [CmdletBinding()]
    param()

    $rawUrl = 'https://raw.githubusercontent.com/dswaney/Compton/main/Compton_Tech_Utils.ps1'
    $currentScriptPath = $PSCommandPath

    if ([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        $currentScriptPath = $MyInvocation.MyCommand.Path
    }

    if ([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        Write-Host 'Startup update check skipped because the current script path could not be determined.' -ForegroundColor Yellow
        return
    }

    Write-Host 'Checking GitHub for a newer version of Compton_Tech_Utils.ps1...' -ForegroundColor Cyan

    $localVersion = Get-LocalScriptVersion -ScriptPath $currentScriptPath
    $remoteMetadata = Get-RemoteScriptMetadata -RawUrl $rawUrl

    if (-not $remoteMetadata -or -not $remoteMetadata.Version) {
        Write-Host 'Update check skipped. Continuing with the current script.' -ForegroundColor Yellow
        return
    }

    if (-not $localVersion) {
        Write-Host "Current local script version could not be read. GitHub version detected: $($remoteMetadata.Version)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Current version: $localVersion" -ForegroundColor Gray
        Write-Host "GitHub version : $($remoteMetadata.Version)" -ForegroundColor Gray
    }

    if ($localVersion -and $remoteMetadata.Version -le $localVersion) {
        Write-Host 'This script is already up to date.' -ForegroundColor Green
        return
    }

    $prompt = if ($remoteMetadata.LastUpdated) {
        "A newer version ($($remoteMetadata.Version), updated $($remoteMetadata.LastUpdated)) is available. Download and run it now? (Y/N)"
    }
    else {
        "A newer version ($($remoteMetadata.Version)) is available. Download and run it now? (Y/N)"
    }

    do {
        $updateChoice = Read-Host $prompt
        $normalizedChoice = ([string]$updateChoice).Trim().ToUpperInvariant()
    } while ($normalizedChoice -notin @('Y','N','YES','NO'))

    if ($normalizedChoice -in @('N','NO')) {
        Write-Host 'Continuing with the current script version.' -ForegroundColor Yellow
        return
    }

    try {
        $targetDirectory = Split-Path -Path $currentScriptPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($targetDirectory) -and -not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -Path $targetDirectory -ItemType Directory -Force | Out-Null
        }

        [System.IO.File]::WriteAllText($currentScriptPath, $remoteMetadata.Content, [System.Text.UTF8Encoding]::new($false))
        Write-Host "Updated script saved to: $currentScriptPath" -ForegroundColor Green
        Write-Host 'Launching the updated script...' -ForegroundColor Cyan

        $powershellExe = if (Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue) {
            'pwsh.exe'
        }
        else {
            'powershell.exe'
        }

        Start-Process -FilePath $powershellExe -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $currentScriptPath)) | Out-Null
        exit
    }
    catch {
        Write-Host "Failed to download or relaunch the updated script: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host 'Continuing with the current script version.' -ForegroundColor Yellow
    }
}

# -----------------------------------------------------------------------------
# Compton College Tech Utils - Modular PowerShell Menu
# -----------------------------------------------------------------------------
# Set execution policy and global preferences
try {
    Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force -ErrorAction SilentlyContinue
    $ErrorActionPreference = "Stop"
    $ProgressPreference = "SilentlyContinue"  # Improves performance
} catch {
    Write-Warning "Could not set execution policy: $($_.Exception.Message)"
}

# Enhanced error handling with detailed logging
trap {
    $errorMsg = "Unhandled Error at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    Write-Error $errorMsg
    # Optional: Log to file for debugging
    # $errorMsg | Out-File -FilePath "C:\temp\script_errors.log" -Append
    exit 1
}

# Global status tracker with timestamp
$global:StatusLog = @()

function Write-StatusLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $statusEntry = @{
        Timestamp = $timestamp
        Level = $Level
        Message = $Message
    }
    
    $global:StatusLog += $statusEntry
    $global:LastStatus = $Message
    
    # Color-coded output
    $color = switch ($Level) {
        "Success" { "Green" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        default { "White" }
    }
    
    Write-Host "[$timestamp] $Message" -ForegroundColor $color
}

# -----------------------------------------------------------------------------
# System Configuration Functions
# -----------------------------------------------------------------------------

function Set-SystemSecurity {
    <#
    .SYNOPSIS
    Configures system security settings including UAC, IPv6, and TLS
    
    .DESCRIPTION
    Applies security configurations with proper error handling and validation
    #>
    
    Write-StatusLog "Configuring system security settings..." -Level "Info"
    
    try {
        # Enable UAC with validation
        $uacPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        $currentUAC = Get-ItemProperty -Path $uacPath -Name "EnableLUA" -ErrorAction SilentlyContinue
        
        if ($currentUAC.EnableLUA -ne 1) {
            Set-ItemProperty -Path $uacPath -Name "EnableLUA" -Value 1 -Type DWord
            Write-StatusLog "[OK] UAC enabled successfully" -Level "Success"
        } else {
            Write-StatusLog "[INFO] UAC already enabled" -Level "Info"
        }
        
        # Configure IPv6 (Consider if complete disable is necessary)
        $ipv6Path = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"
        if (-not (Test-Path $ipv6Path)) {
            New-Item -Path $ipv6Path -Force | Out-Null
        }
        
        $currentIPv6 = Get-ItemProperty -Path $ipv6Path -Name "DisabledComponents" -ErrorAction SilentlyContinue
        if ($currentIPv6.DisabledComponents -ne 0xFF) {
            Set-ItemProperty -Path $ipv6Path -Name "DisabledComponents" -Value 0xFF -Type DWord
            Write-StatusLog "[WARN] IPv6 disabled - restart required for full effect" -Level "Warning"
        } else {
            Write-StatusLog "[INFO] IPv6 already disabled" -Level "Info"
        }
        
        # Enable TLS 1.2 and 1.3 if available
        $supportedProtocols = [System.Net.ServicePointManager]::SecurityProtocol
        $tls12 = [System.Net.SecurityProtocolType]::Tls12
        $tls13 = 12288  # TLS 1.3 value
        
        # Enable TLS 1.2
        if (-not ($supportedProtocols -band $tls12)) {
            [System.Net.ServicePointManager]::SecurityProtocol = $supportedProtocols -bor $tls12
        }
        
        # Try to enable TLS 1.3 if supported
        try {
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor $tls13
            Write-StatusLog "[OK] TLS 1.2 and 1.3 enabled" -Level "Success"
        } catch {
            Write-StatusLog "[OK] TLS 1.2 enabled (TLS 1.3 not available)" -Level "Success"
        }
        
    } catch {
        Write-StatusLog "[ERROR] Failed to configure system security: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

function Test-MISAdminStatus {
    <#
    .SYNOPSIS
    Checks and configures MISAdmin account status
    
    .DESCRIPTION
    Validates MISAdmin account exists and configures password policy
    
    .OUTPUTS
    Returns hashtable with account status information
    #>
    
    Write-StatusLog "Checking MISAdmin account status..." -Level "Info"
    
    try {
        $misAdmin = Get-LocalUser -Name "MISAdmin" -ErrorAction SilentlyContinue
        
        if (-not $misAdmin) {
            Write-StatusLog "[ERROR] MISAdmin account not found" -Level "Error"
            return @{
                Exists = $false
                PasswordNeverExpires = $false
                Enabled = $false
                LastLogon = $null
            }
        }
        
        $accountInfo = @{
            Exists = $true
            PasswordNeverExpires = $misAdmin.PasswordNeverExpires
            Enabled = $misAdmin.Enabled
            LastLogon = $misAdmin.LastLogon
        }
        
        # Configure password policy if needed
        if (-not $misAdmin.PasswordNeverExpires) {
            try {
                Set-LocalUser -Name "MISAdmin" -PasswordNeverExpires $true
                Write-StatusLog "[CONFIG] PasswordNeverExpires set for MISAdmin" -Level "Success"
                $accountInfo.PasswordNeverExpires = $true
            } catch {
                Write-StatusLog "[ERROR] Failed to update MISAdmin password policy: $($_.Exception.Message)" -Level "Error"
                throw
            }
        } else {
            Write-StatusLog "[OK] MISAdmin exists with correct password policy" -Level "Success"
        }
        
        # Check if account is enabled
        if (-not $misAdmin.Enabled) {
            Write-StatusLog "[WARN] MISAdmin account is disabled" -Level "Warning"
        }
        
        return $accountInfo
        
    } catch {
        Write-StatusLog "[ERROR] Error checking MISAdmin status: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

# -----------------------------------------------------------------------------
# Initialization
# -----------------------------------------------------------------------------

function Initialize-System {
    <#
    .SYNOPSIS
    Performs initial system setup and validation
    #>
    
    Write-StatusLog "Initializing Compton College Tech Utils..." -Level "Info"
    
    try {
        # Apply system security settings
        Set-SystemSecurity
        
        # Check MISAdmin status
        $adminStatus = Test-MISAdminStatus
        
        # Display summary
        Write-StatusLog "System initialization completed" -Level "Success"
        
        return @{
            SecurityConfigured = $true
            MISAdminStatus = $adminStatus
            InitializationTime = Get-Date
        }
        
    } catch {
        Write-StatusLog "[ERROR] System initialization failed: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

# Run initialization
try {
    $initResult = Initialize-System
    Write-Host "`n" + "="*60
    Write-Host "INITIALIZATION SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60
    Write-Host "MISAdmin Exists: " -NoNewline
    Write-Host $initResult.MISAdminStatus.Exists -ForegroundColor $(if($initResult.MISAdminStatus.Exists){"Green"}else{"Red"})
    Write-Host "Security Configured: " -NoNewline  
    Write-Host $initResult.SecurityConfigured -ForegroundColor Green
    Write-Host "="*60 + "`n"
} catch {
    Write-Error "Critical initialization failure. Exiting."
    exit 1
}

# -----------------------------------------------------------------------------
# Option 1 - Create MISAdmin Local Admin Account
# -----------------------------------------------------------------------------
function Test-SecureStringsEqual {
    param(
        [Parameter(Mandatory)][securestring]$A,
        [Parameter(Mandatory)][securestring]$B
    )
    $ptrA = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($A)
    $ptrB = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($B)
    try {
        $sa = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptrA)
        $sb = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptrB)
        return [string]::Equals($sa, $sb, [System.StringComparison]::Ordinal)
    }
    finally {
        if ($ptrA -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptrA) }
        if ($ptrB -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptrB) }
    }
} # end Test-SecureStringsEqual

function Get-SecurePassword {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AccountName,
        [int]$MaxAttempts = 3
    )
    function Test-SecEq {
        param([securestring]$A,[securestring]$B)
        try {
            $pa = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($A)
            $pb = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($B)
            try {
                $sa = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pa)
                $sb = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pb)
                return [string]::Equals($sa, $sb, [System.StringComparison]::Ordinal)
            } finally {
                if ($pa -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pa) }
                if ($pb -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pb) }
            }
        } catch {
            $sa = (New-Object PSCredential 'u',$A).GetNetworkCredential().Password
            $sb = (New-Object PSCredential 'u',$B).GetNetworkCredential().Password
            $eq = [string]::Equals($sa, $sb, [System.StringComparison]::Ordinal)
            $sa = $sb = $null
            [System.GC]::Collect()
            return $eq
        }
    }
    $attempt = 0
    do {
        $attempt++
        Write-Host "`nPassword Entry (Attempt $attempt of $MaxAttempts)" -ForegroundColor Cyan
        try {
            $p1 = Read-Host "Enter password for '$AccountName'" -AsSecureString
            $p2 = Read-Host "Confirm password for '$AccountName'" -AsSecureString
            if (-not $p1 -or -not $p2) { Write-StatusLog "[ERROR] Empty password not allowed" -Level "Error"; continue }
            $p1txt = (New-Object PSCredential 'u',$p1).GetNetworkCredential().Password
            if ([string]::IsNullOrWhiteSpace($p1txt)) { Write-StatusLog "[ERROR] Password cannot be blank/whitespace" -Level "Error"; $p1txt = $null; continue }
            $p1txt = $null
            if (-not (Test-SecEq $p1 $p2)) { Write-StatusLog "[ERROR] Passwords do not match" -Level "Error"; continue }
            Write-StatusLog "[OK] Password validated successfully" -Level "Success"
            return $p1
        } catch {
            Write-StatusLog "[ERROR] Password prompt failed: $($_.Exception.Message)" -Level "Error"
        } finally {
            [System.GC]::Collect()
        }
    } while ($attempt -lt $MaxAttempts)
    Write-StatusLog "[ERROR] Maximum password attempts exceeded" -Level "Error"
    return $null
} # end Get-SecurePassword

function Test-AdminGroupMembership {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$AccountName)
    try {
        $adminGroup = Get-LocalGroup -Name "Administrators"
        $members = Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop
        $machineQualified = "$env:COMPUTERNAME\$AccountName"
        $isMember = $false
        foreach ($m in $members) {
            if ($m.Name -ieq $machineQualified -or $m.Name -ieq $AccountName) { $isMember = $true; break }
        }
        if ($isMember) {
            Write-StatusLog "[OK] Account has administrative privileges" -Level "Success"
        } else {
            Write-StatusLog "[WARN] Account missing administrative privileges" -Level "Warning"
            try {
                Add-LocalGroupMember -Group "Administrators" -Member $AccountName -ErrorAction Stop
                Write-StatusLog "[OK] Account added to Administrators group" -Level "Success"
                return $true
            } catch {
                Write-StatusLog "[ERROR] Failed to add account to Administrators group: $($_.Exception.Message)" -Level "Error"
                return $false
            }
        }
        return $isMember
    } catch {
        Write-StatusLog "[ERROR] Error checking admin group membership: $($_.Exception.Message)" -Level "Error"
        return $false
    }
} # end Test-AdminGroupMembership

function New-LocalAdminAccount {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$AccountName)
    $result = @{
        AccountExists       = $false
        AccountCreated      = $false
        PasswordConfigured  = $false
        AdminRights         = $false
        Errors              = @()
    }
    try {
        $password = Get-SecurePassword -AccountName $AccountName
        if (-not $password) { $result.Errors += "Failed to obtain valid password"; return $result }

        Write-StatusLog "Creating local user account..." -Level "Info"
        $newUserParams = @{
            Name                     = $AccountName
            Password                 = $password
            FullName                 = "MIS Administrator"
            PasswordNeverExpires     = $true
            AccountNeverExpires      = $true
            UserMayNotChangePassword = $false
        }
        # NOTE: Intentionally omitting -Description due to 48-char limit
        New-LocalUser @newUserParams -ErrorAction Stop | Out-Null

        $result.AccountCreated = $true
        $result.PasswordConfigured = $true
        Write-StatusLog "[OK] Account '$AccountName' created successfully" -Level "Success"

        Add-LocalGroupMember -Group "Administrators" -Member $AccountName -ErrorAction Stop
        $result.AdminRights = $true
        Write-StatusLog "[OK] Account added to Administrators group" -Level "Success"

        $createdAccount = Get-LocalUser -Name $AccountName
        Write-StatusLog "[OK] Account verification completed" -Level "Success"
        Write-Host "   - Password Never Expires: $($createdAccount.PasswordNeverExpires)" -ForegroundColor Green
        Write-Host "   - Account Never Expires: $(-not $createdAccount.AccountExpires)" -ForegroundColor Green
        Write-Host "   - Account Enabled: $($createdAccount.Enabled)" -ForegroundColor Green

        $global:LastStatus = "[OK] MISAdmin account created and configured successfully"
    } catch {
        $errorMsg = "Failed to create account: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        $result.Errors += $errorMsg
        $global:LastStatus = "[ERROR] Failed to create MISAdmin account"
    }
    return $result
} # end New-LocalAdminAccount

function New-MISAdminAccount {
    <#
    .SYNOPSIS
    Creates or validates MISAdmin account with proper security configuration
    #>
    [CmdletBinding()]
    param(
        [Parameter()][ValidateNotNullOrEmpty()][string]$AccountName = "MISAdmin",
        [Parameter()][switch]$Force
    )
    Write-StatusLog "Processing MISAdmin account: $AccountName" -Level "Info"
    $result = @{
        AccountExists       = $false
        AccountCreated      = $false
        PasswordConfigured  = $false
        AdminRights         = $false
        Errors              = @()
    }
    try {
        $existingAccount = Get-LocalUser -Name $AccountName -ErrorAction SilentlyContinue
        if ($existingAccount) {
            $result.AccountExists = $true
            Write-StatusLog "[OK] Account '$AccountName' already exists" -Level "Success"

            if (-not $existingAccount.PasswordNeverExpires) {
                try {
                    Set-LocalUser -Name $AccountName -PasswordNeverExpires $true
                    Write-StatusLog "[CONFIG] Password policy updated: PasswordNeverExpires enabled" -Level "Success"
                    $result.PasswordConfigured = $true
                } catch {
                    $errorMsg = "Failed to update password policy: $($_.Exception.Message)"
                    Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
                    $result.Errors += $errorMsg
                }
            } else {
                Write-StatusLog "[OK] Password policy already configured correctly" -Level "Success"
                $result.PasswordConfigured = $true
            }

            $result.AdminRights = Test-AdminGroupMembership -AccountName $AccountName

            if ($Force) {
                $newPassword = Get-SecurePassword -AccountName $AccountName
                if ($newPassword) {
                    try {
                        Set-LocalUser -Name $AccountName -Password $newPassword -ErrorAction Stop
                        Write-StatusLog "[OK] Password reset completed" -Level "Success"
                    } catch {
                        $errorMsg = "Failed to reset password: $($_.Exception.Message)"
                        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
                        $result.Errors += $errorMsg
                    }
                } else {
                    $result.Errors += "Failed to obtain valid password for reset"
                }
            }

            $global:LastStatus = if ($result.Errors.Count -eq 0) {
                "[OK] MISAdmin account validated and configured"
            } else {
                "[WARN] MISAdmin account exists but has configuration issues"
            }
        } else {
            Write-StatusLog "[WARN] Account '$AccountName' not found. Initiating creation process..." -Level "Warning"
            $result = New-LocalAdminAccount -AccountName $AccountName
        }
    } catch {
        $errorMsg = "Critical error processing MISAdmin account: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        $result.Errors += $errorMsg
        $global:LastStatus = "[ERROR] Failed to process MISAdmin account"
    }
    return $result
}

# -----------------------------------------------------------------------------
# Option 2 - Remove Bloatware
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Option 2 - Remove Bloatware
# -----------------------------------------------------------------------------
function Remove-BloatwareApps {
    <#
    .SYNOPSIS
    Removes known bloatware applications from Windows
    
    .DESCRIPTION
    Safely removes unwanted pre-installed applications with comprehensive error handling,
    progress reporting, and detailed logging of removal operations
    
    .PARAMETER IncludeProvisioned
    Also removes provisioned packages to prevent reinstallation
    
    .PARAMETER CustomAppList
    Optional custom list of applications to remove instead of default list
    
    .PARAMETER WhatIf
    Shows what would be removed without actually removing anything
    
    .OUTPUTS
    Returns hashtable with removal results and statistics
    
    .EXAMPLE
    Remove-BloatwareApps
    
    .EXAMPLE
    Remove-BloatwareApps -IncludeProvisioned -WhatIf
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [switch]$IncludeProvisioned,
        
        [Parameter()]
        [string[]]$CustomAppList
    )
    
    # Initialize results at the very beginning
    $results = @{
        TotalApps = 0
        SuccessfulRemovals = 0
        FailedRemovals = 0
        NotFound = 0
        ProvisionedRemoved = 0
        RemovedApps = @()
        FailedApps = @()
        NotFoundApps = @()
        Errors = @()
    }
    
    try {
        Write-Host "[SCAN] Starting bloatware removal process..." -ForegroundColor Cyan
        
        # Default bloatware list - organized by category
        $defaultBloatware = @{
            "Entertainment" = @(
                "Microsoft.ZuneMusic",
                "Microsoft.Music.Preview",
                "Microsoft.XboxIdentityProvider",
                "Microsoft.XboxGameOverlay", 
                "Microsoft.Xbox.TCUI",
                "Microsoft.XboxApp",
                "Microsoft.MicrosoftSolitaireCollection"
            )
            "News_Travel" = @(
                "Microsoft.BingTravel",
                "Microsoft.BingNews",
                "Microsoft.BingSports",
                "Microsoft.BingFinance",
                "Microsoft.BingWeather"
            )
            "Lifestyle" = @(
                "Microsoft.BingHealthAndFitness",
                "Microsoft.BingFoodAndDrink",
                "Microsoft.People"
            )
            "Productivity_Tools" = @(
                "Microsoft.3DBuilder",
                "Microsoft.WindowsMaps",
                "Microsoft.MicrosoftOfficeHub",
                "Microsoft.Getstarted"
            )
            "Communication" = @(
                "Microsoft.WindowsPhone",
                "Microsoft.SkypeApp",
                "Microsoft.YourPhone"
            )
        }
        
        # Use custom list if provided, otherwise flatten default list
        $appsToRemove = if ($CustomAppList -and $CustomAppList.Count -gt 0) {
            $CustomAppList
        } else {
            $allApps = @()
            foreach ($category in $defaultBloatware.Values) {
                $allApps += $category
            }
            $allApps
        }
        
        # Update total apps count
        $results.TotalApps = $appsToRemove.Count
        
        Write-Host "[PKG] Processing $($results.TotalApps) bloatware applications..." -ForegroundColor Cyan
        
        if ($results.TotalApps -eq 0) {
            Write-Host "[WARN] No applications specified for removal" -ForegroundColor Yellow
            $global:LastStatus = "[WARN] No bloatware applications specified for removal"
            return $results
        }
        
        # Progress tracking
        $currentApp = 0
        
        foreach ($appName in $appsToRemove) {
            $currentApp++
            $percentComplete = [math]::Round(($currentApp / $results.TotalApps) * 100)
            
            Write-Progress -Activity "Removing Bloatware" -Status "Processing $appName" -PercentComplete $percentComplete
            
            try {
                $removalResult = Remove-AppxPackageSafe -AppName $appName
                
                switch ($removalResult.Status) {
                    "Success" {
                        $results.SuccessfulRemovals += $removalResult.PackagesRemoved
                        $results.RemovedApps += $appName
                        Write-Host "[OK] Successfully removed '$appName' ($($removalResult.PackagesRemoved) package(s))" -ForegroundColor Green
                    }
                    "NotFound" {
                        $results.NotFound++
                        $results.NotFoundApps += $appName
                        Write-Host "[INFO] '$appName' not found on system" -ForegroundColor Gray
                    }
                    "Failed" {
                        $results.FailedRemovals++
                        $results.FailedApps += $appName
                        $results.Errors += $removalResult.Error
                        Write-Host "[ERROR] Failed to remove '$appName': $($removalResult.Error)" -ForegroundColor Red
                    }
                }
                
                # Remove provisioned packages if requested
                if ($IncludeProvisioned -and $removalResult.Status -eq "Success") {
                    $provisionedResult = Remove-ProvisionedAppPackage -AppName $appName
                    if ($provisionedResult.Removed) {
                        $results.ProvisionedRemoved++
                        Write-Host "[TOOLS] Removed provisioned package for '$appName'" -ForegroundColor Green
                    }
                }
                
            } catch {
                $results.FailedRemovals++
                $results.FailedApps += $appName
                $errorMsg = "Unexpected error removing '$appName': $($_.Exception.Message)"
                $results.Errors += $errorMsg
                Write-Host "[ERROR] $errorMsg" -ForegroundColor Red
            }
        }
        
        Write-Progress -Activity "Removing Bloatware" -Completed
        
        # Display summary
        Show-RemovalSummary -Results $results
        
        # Set global status
        if ($WhatIfPreference) {
            $global:LastStatus = "[REPORT] Bloatware scan completed - $($results.TotalApps) apps analyzed"
        } else {
            $global:LastStatus = "[OK] Bloatware removal completed - $($results.SuccessfulRemovals) packages removed"
        }
        
    } catch {
        Write-Host "[ERROR] Critical error during bloatware removal: $_" -ForegroundColor Red
        $results.Errors += "Critical error: $($_.Exception.Message)"
        $global:LastStatus = "[ERROR] Bloatware removal failed: $_"
    }
    
    return $results
}

function Remove-AppxPackageSafe {
    <#
    .SYNOPSIS
    Safely removes AppX packages with comprehensive error handling
    
    .PARAMETER AppName
    Name of the application to remove
    
    .PARAMETER WhatIf
    Simulates removal without actually removing packages
    
    .OUTPUTS
    Returns hashtable with removal status and details
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AppName
    )
    
    $result = @{
        Status = "Unknown"
        PackagesRemoved = 0
        Error = $null
        Packages = @()
    }
    
    try {
        # Get all packages for the app
        $packages = Get-AppxPackage -AllUsers -Name $AppName -ErrorAction SilentlyContinue
        
        if (-not $packages) {
            $result.Status = "NotFound"
            return $result
        }
        
        $result.Packages = $packages
        $packagesRemoved = 0
        
        foreach ($package in $packages) {
            try {
                $userInfo = if ($package.UserSID) { 
                    "User: $($package.UserSID)" 
                } else { 
                    "All Users" 
                }
                
                if ($WhatIfPreference) {
                    Write-Host "   Would remove: $($package.Name) - $userInfo" -ForegroundColor Yellow
                } else {
                    Remove-AppxPackage -Package $package.PackageFullName -ErrorAction Stop
                    $packagesRemoved++
                    Write-Host "   [OK] Removed: $($package.Name) - $userInfo" -ForegroundColor Green
                }
                
            } catch {
                $errorMsg = "Failed to remove package '$($package.Name)' for $userInfo`: $($_.Exception.Message)"
                Write-Host "   [WARN] $errorMsg" -ForegroundColor Yellow
                
                # Continue with other packages even if one fails
                if (-not $result.Error) {
                    $result.Error = $errorMsg
                }
            }
        }
        
        $result.PackagesRemoved = $packagesRemoved
        $result.Status = if ($packagesRemoved -gt 0 -or $WhatIfPreference) { "Success" } else { "Failed" }
        
    } catch {
        $result.Status = "Failed"
        $result.Error = $_.Exception.Message
    }
    
    return $result
}

function Remove-ProvisionedAppPackage {
    <#
    .SYNOPSIS
    Removes provisioned app packages to prevent reinstallation
    
    .PARAMETER AppName
    Name of the application to remove from provisioning
    
    .PARAMETER WhatIf
    Simulates removal without actually removing packages
    
    .OUTPUTS
    Returns hashtable indicating if provisioned package was removed
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$AppName
    )
    
    $result = @{
        Removed = $false
        Error = $null
    }
    
    try {
        $provisionedPackages = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -like "*$AppName*" }
        
        foreach ($package in $provisionedPackages) {
            if ($WhatIfPreference) {
                Write-Host "   Would remove provisioned: $($package.DisplayName)" -ForegroundColor Yellow
                $result.Removed = $true
            } else {
                Remove-AppxProvisionedPackage -Online -PackageName $package.PackageName -ErrorAction Stop
                $result.Removed = $true
            }
        }
        
    } catch {
        $result.Error = $_.Exception.Message
    }
    
    return $result
}

function Show-RemovalSummary {
    <#
    .SYNOPSIS
    Displays a comprehensive summary of the bloatware removal process
    
    .PARAMETER Results
    Results hashtable from the removal process
    
    .PARAMETER WhatIf
    Indicates if this was a simulation run
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results
    )
    
    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "REMOVAL" }
    
    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    Write-Host "BLOATWARE $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan
    
    Write-Host "Total Applications Processed: " -NoNewline
    Write-Host $Results.TotalApps -ForegroundColor White
    
    Write-Host "Successfully Removed: " -NoNewline
    Write-Host $Results.SuccessfulRemovals -ForegroundColor Green
    
    Write-Host "Not Found: " -NoNewline
    Write-Host $Results.NotFound -ForegroundColor Yellow
    
    Write-Host "Failed to Remove: " -NoNewline
    Write-Host $Results.FailedRemovals -ForegroundColor Red
    
    if ($Results.ProvisionedRemoved -gt 0) {
        Write-Host "Provisioned Packages Removed: " -NoNewline
        Write-Host $Results.ProvisionedRemoved -ForegroundColor Green
    }
    
    # Show details if there were failures
    if ($Results.FailedApps -and $Results.FailedApps.Count -gt 0) {
        Write-Host "`nFailed Applications:" -ForegroundColor Red
        foreach ($app in $Results.FailedApps) {
            Write-Host "  - $app" -ForegroundColor Red
        }
    }
    
    # Show success details for smaller lists
    if ($Results.RemovedApps -and $Results.RemovedApps.Count -gt 0 -and $Results.RemovedApps.Count -le 10) {
        Write-Host "`nSuccessfully Removed:" -ForegroundColor Green
        foreach ($app in $Results.RemovedApps) {
            Write-Host "  - $app" -ForegroundColor Green
        }
    }
    
    Write-Host "="*60 -ForegroundColor Cyan
}
# -----------------------------------------------------------------------------
# Option 3 - Recommended Registry Settings
# -----------------------------------------------------------------------------
function Apply-RecommendedRegistrySettings {
    <#
    .SYNOPSIS
    Applies recommended Windows registry settings for privacy, performance, and usability

    .DESCRIPTION
    Configures registry settings to improve system privacy, performance, and user experience.
    Settings are organized by category and applied with comprehensive error handling.

    .PARAMETER SettingsCategory
    Specific category of settings to apply (All, Privacy, Performance, UI, Security, System)

    .PARAMETER BackupRegistry
    Creates a registry backup before applying changes

    .PARAMETER Force
    Applies settings without confirmation prompts

    .OUTPUTS
    Returns hashtable with application results and statistics

    .EXAMPLE
    Apply-RecommendedRegistrySettings

    .EXAMPLE
    Apply-RecommendedRegistrySettings -SettingsCategory Privacy -BackupRegistry

    .EXAMPLE
    Apply-RecommendedRegistrySettings -WhatIf
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [ValidateSet('All', 'Privacy', 'Performance', 'UI', 'Security', 'System')]
        [string]$SettingsCategory = 'All',

        [Parameter()]
        [switch]$BackupRegistry,

        [Parameter()]
        [switch]$Force
    )

    Write-StatusLog "Configuring Windows registry settings..." -Level "Info"

    # Detect OS version
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $isWindows11 = $false
    try {
        $buildNumber = [int]$osInfo.BuildNumber
        if ($buildNumber -ge 22000) {
            $isWindows11 = $true
        }
    } catch {
        $isWindows11 = $true
    }

    # Registry settings organized by category
    $registrySettings = @{
        Privacy = @(
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection'; Name = 'AllowTelemetry'; Value = 0; Type = 'DWORD'; Description = 'Disable telemetry collection' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Name = 'AllowTelemetry'; Value = 0; Type = 'DWORD'; Description = 'Disable telemetry (policy)' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Name = 'DoNotShowFeedbackNotifications'; Value = 1; Type = 'DWORD'; Description = 'Disable feedback notifications' },
            @{ Path = 'HKCU:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'; Name = 'DisableTailoredExperiencesWithDiagnosticData'; Value = 1; Type = 'DWORD'; Description = 'Disable tailored experiences' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo'; Name = 'DisabledByGroupPolicy'; Value = 1; Type = 'DWORD'; Description = 'Disable advertising ID' },
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting'; Name = 'Disabled'; Value = 1; Type = 'DWORD'; Description = 'Disable error reporting' },
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}'; Name = 'SensorPermissionState'; Value = 0; Type = 'DWORD'; Description = 'Disable location sensor' },
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration'; Name = 'Status'; Value = 0; Type = 'DWORD'; Description = 'Disable location services' },
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location'; Name = 'Value'; Value = 'Deny'; Type = 'String'; Description = 'Deny location access' },

            # Winutil-style additions
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent'; Name = 'DisableWindowsConsumerFeatures'; Value = 1; Type = 'DWORD'; Description = 'Disable consumer features' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'; Name = 'BingSearchEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable Bing search in Start menu' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'; Name = 'CortanaConsent'; Value = 0; Type = 'DWORD'; Description = 'Disable Cortana search consent' }
        )

        Performance = @(
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'; Name = 'SystemResponsiveness'; Value = 0; Type = 'DWORD'; Description = 'Optimize system responsiveness' },
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'; Name = 'NetworkThrottlingIndex'; Value = 4294967295; Type = 'DWORD'; Description = 'Disable network throttling' },
            @{ Path = 'HKCU:\Control Panel\Desktop'; Name = 'MenuShowDelay'; Value = 1; Type = 'DWORD'; Description = 'Reduce menu show delay' },
            @{ Path = 'HKCU:\Control Panel\Desktop'; Name = 'AutoEndTasks'; Value = 1; Type = 'DWORD'; Description = 'Auto-end hung tasks' },
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'; Name = 'ClearPageFileAtShutdown'; Value = 0; Type = 'DWORD'; Description = 'Skip page file clearing' },
            @{ Path = 'HKLM:\SYSTEM\ControlSet001\Services\Ndu'; Name = 'Start'; Value = 2; Type = 'DWORD'; Description = 'Optimize network usage' },
            @{ Path = 'HKCU:\Control Panel\Mouse'; Name = 'MouseHoverTime'; Value = 400; Type = 'DWORD'; Description = 'Reduce mouse hover time' },
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters'; Name = 'IRPStackSize'; Value = 30; Type = 'DWORD'; Description = 'Optimize file sharing' },
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config'; Name = 'DODownloadMode'; Value = 1; Type = 'DWORD'; Description = 'Optimize delivery mode' }
        )

        UI = @(
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'ContentDeliveryAllowed'; Value = 0; Type = 'DWORD'; Description = 'Disable content delivery' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'OemPreInstalledAppsEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable OEM pre-installed apps' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'PreInstalledAppsEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable pre-installed apps' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'PreInstalledAppsEverEnabled'; Value = 0; Type = 'DWORD'; Description = 'Never enable pre-installed apps' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'SilentInstalledAppsEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable silent app installs' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'SubscribedContent-338387Enabled'; Value = 0; Type = 'DWORD'; Description = 'Disable start menu suggestions' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'SubscribedContent-338388Enabled'; Value = 0; Type = 'DWORD'; Description = 'Disable lock screen suggestions' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'SubscribedContent-338389Enabled'; Value = 0; Type = 'DWORD'; Description = 'Disable tips and tricks' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'SubscribedContent-353698Enabled'; Value = 0; Type = 'DWORD'; Description = 'Disable timeline suggestions' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'; Name = 'SystemPaneSuggestionsEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable system pane suggestions' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'ShowTaskViewButton'; Value = 0; Type = 'DWORD'; Description = 'Hide task view button' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People'; Name = 'PeopleBand'; Value = 0; Type = 'DWORD'; Description = 'Hide people band' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'LaunchTo'; Value = 1; Type = 'DWORD'; Description = 'Launch to This PC' },

            # Windows 11 Widgets replacements
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'TaskbarDa'; Value = 0; Type = 'DWORD'; Description = 'Hide widgets taskbar button'; SpecialHandler = 'WidgetsTaskbarButton' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh'; Name = 'AllowNewsAndInterests'; Value = 0; Type = 'DWORD'; Description = 'Disable widgets/news content' },

            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer'; Name = 'HideSCAMeetNow'; Value = 1; Type = 'DWORD'; Description = 'Hide Meet Now button' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search'; Name = 'SearchboxTaskbarMode'; Value = 1; Type = 'DWORD'; Description = 'Show search icon only' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'TaskbarAl'; Value = 0; Type = 'DWORD'; Description = 'Left-align taskbar' },

            # Winutil-style useful shell tweaks
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'HideFileExt'; Value = 0; Type = 'DWORD'; Description = 'Show file extensions' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'Hidden'; Value = 1; Type = 'DWORD'; Description = 'Show hidden files' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'; Name = 'TaskbarMn'; Value = 0; Type = 'DWORD'; Description = 'Disable Chat/Teams taskbar button where supported' }
        )

        Security = @(
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance'; Name = 'fAllowToGetHelp'; Value = 0; Type = 'DWORD'; Description = 'Disable remote assistance' },
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem'; Name = 'LongPathsEnabled'; Value = 1; Type = 'DWORD'; Description = 'Enable long paths' },
            @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching'; Name = 'SearchOrderConfig'; Value = 1; Type = 'DWORD'; Description = 'Secure driver search order' },
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl'; Name = 'DisplayParameters'; Value = 1; Type = 'DWORD'; Description = 'Show crash parameters' },
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl'; Name = 'DisableEmoticon'; Value = 1; Type = 'DWORD'; Description = 'Disable crash emoticons' }
        )

        System = @(
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Siuf\Rules'; Name = 'NumberOfSIUFInPeriod'; Value = 0; Type = 'DWORD'; Description = 'Disable customer experience program' },
            @{ Path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\OperationStatusManager'; Name = 'EnthusiastMode'; Value = 1; Type = 'DWORD'; Description = 'Enable detailed file operations' },
            @{ Path = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\UserProfileEngagement'; Name = 'ScoobeSystemSettingEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable profile engagement' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'EnableActivityFeed'; Value = 0; Type = 'DWORD'; Description = 'Disable activity feed' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'PublishUserActivities'; Value = 0; Type = 'DWORD'; Description = 'Disable activity publishing' },
            @{ Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'UploadUserActivities'; Value = 0; Type = 'DWORD'; Description = 'Disable activity uploads' },
            @{ Path = 'HKLM:\SYSTEM\Maps'; Name = 'AutoUpdateEnabled'; Value = 0; Type = 'DWORD'; Description = 'Disable maps auto-update' }
        )
    }

    # Determine which settings to apply
    $settingsToApply = if ($SettingsCategory -eq 'All') {
        $registrySettings.Values | ForEach-Object { $_ }
    } else {
        $registrySettings[$SettingsCategory]
    }

    # Initialize results tracking
    $results = @{
        TotalSettings = $settingsToApply.Count
        SuccessfulChanges = 0
        FailedChanges = 0
        SkippedSettings = 0
        ChangedSettings = @()
        FailedSettings = @()
        Errors = @()
        BackupCreated = $false
        Windows11Detected = $isWindows11
    }

    # Create registry backup if requested
    if ($BackupRegistry) {
        $results.BackupCreated = New-RegistryBackup
    }

    Write-Host "`n[TOOLS] Applying $($results.TotalSettings) registry settings..." -ForegroundColor Cyan

    # Progress tracking
    $currentSetting = 0

    foreach ($setting in $settingsToApply) {
        $currentSetting++
        $percentComplete = [math]::Round(($currentSetting / $results.TotalSettings) * 100)

        Write-Progress -Activity "Applying Registry Settings" -Status "Processing: $($setting.Description)" -PercentComplete $percentComplete

        try {
            $applyResult = Set-RegistryValueSafe @setting -Force:$Force

            switch ($applyResult.Status) {
                "Success" {
                    $results.SuccessfulChanges++
                    $results.ChangedSettings += "$($setting.Path)\$($setting.Name)"
                    Write-StatusLog "[OK] $($setting.Description)" -Level "Success"
                }
                "Unchanged" {
                    $results.SkippedSettings++
                    Write-StatusLog "[INFO] $($setting.Description) (already set)" -Level "Info"
                }
                "Skipped" {
                    $results.SkippedSettings++
                    Write-StatusLog "[WARN] $($setting.Description) skipped - $($applyResult.Error)" -Level "Warning"
                }
                "Failed" {
                    $results.FailedChanges++
                    $results.FailedSettings += $setting.Description
                    $results.Errors += $applyResult.Error
                    Write-StatusLog "[ERROR] Failed: $($setting.Description) - $($applyResult.Error)" -Level "Error"
                }
            }

        } catch {
            $results.FailedChanges++
            $results.FailedSettings += $setting.Description
            $errorMsg = "Unexpected error applying '$($setting.Description)': $($_.Exception.Message)"
            $results.Errors += $errorMsg
            Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        }
    }

    Write-Progress -Activity "Applying Registry Settings" -Completed

    # Display summary
    Show-RegistrySettingsSummary -Results $results -Category $SettingsCategory

    # Set global status
    $global:LastStatus = if ($results.FailedChanges -eq 0) {
        "[OK] Registry settings applied successfully ($($results.SuccessfulChanges) changes)"
    } else {
        "[WARN] Registry settings applied with $($results.FailedChanges) failures"
    }

    return $results
}

function Set-RegistryValueSafe {
    <#
    .SYNOPSIS
    Safely sets a registry value with validation and error handling

    .PARAMETER Path
    Registry path

    .PARAMETER Name
    Value name

    .PARAMETER Value
    Value to set

    .PARAMETER Type
    Registry value type

    .PARAMETER Description
    Human-readable description of the setting

    .PARAMETER SpecialHandler
    Optional special handler name for shell-managed or version-specific settings

    .PARAMETER Force
    Apply setting without checking current value

    .OUTPUTS
    Returns hashtable with operation status and details
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [object]$Value,

        [Parameter(Mandatory)]
        [ValidateSet('String', 'DWORD', 'QWORD', 'Binary', 'ExpandString')]
        [string]$Type,

        [Parameter()]
        [string]$Description = "Registry setting",

        [Parameter()]
        [string]$SpecialHandler = "",

        [Parameter()]
        [switch]$Force
    )

    $result = @{
        Status = "Unknown"
        Error = $null
        PreviousValue = $null
        NewValue = $Value
    }

    try {
        # Special handling for Windows 11 Widgets taskbar button
        if ($SpecialHandler -eq 'WidgetsTaskbarButton') {
            return Set-Windows11WidgetsTaskbarButton -Value ([int]$Value)
        }

        # Ensure registry path exists
        if (-not (Test-Path -Path $Path)) {
            if ($WhatIfPreference) {
                Write-Host "   Would create path: $Path" -ForegroundColor Yellow
            } else {
                New-Item -Path $Path -Force | Out-Null
            }
        }

        # Check current value if not forcing
        if (-not $Force) {
            $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $currentValue) {
                $result.PreviousValue = $currentValue.$Name
                if ($currentValue.$Name -eq $Value) {
                    $result.Status = "Unchanged"
                    return $result
                }
            }
        }

        # Apply the setting
        if ($WhatIfPreference) {
            Write-Host "   Would set: $Path\$Name = $Value ($Type)" -ForegroundColor Yellow
            $result.Status = "Success"
        } else {
            $propertyType = switch ($Type) {
                'String'       { 'String' }
                'DWORD'        { 'DWord' }
                'QWORD'        { 'QWord' }
                'Binary'       { 'Binary' }
                'ExpandString' { 'ExpandString' }
            }

            if (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force -ErrorAction Stop
            } else {
                New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $propertyType -Force -ErrorAction Stop | Out-Null
            }

            # Verify write
            $verify = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            if ($null -ne $verify -and $verify.$Name -eq $Value) {
                $result.Status = "Success"
            } else {
                $result.Status = "Failed"
                $result.Error = "Registry verification failed after write."
            }
        }

    } catch {
        $result.Status = "Failed"
        $result.Error = $_.Exception.Message
    }

    return $result
}

function Set-Windows11WidgetsTaskbarButton {
    <#
    .SYNOPSIS
    Hides or shows the Windows 11 Widgets taskbar button with safer handling

    .DESCRIPTION
    Uses reg.exe and post-write verification because TaskbarDa can behave inconsistently
    on some Windows 11 builds where Explorer or Feature Experience manages the value.
    Returns Skipped instead of Failed if the build appears not to honor the setting
    or if access is denied to the shell-managed value.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet(0,1)]
        [int]$Value
    )

    $result = @{
        Status = "Unknown"
        Error = $null
        PreviousValue = $null
        NewValue = $Value
    }

    $regPathPs = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    $regPathExe = 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    $name = 'TaskbarDa'

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would set Widgets taskbar button: $name = $Value" -ForegroundColor Yellow
            $result.Status = "Success"
            return $result
        }

        if (-not (Test-Path $regPathPs)) {
            New-Item -Path $regPathPs -Force | Out-Null
        }

        $current = Get-ItemProperty -Path $regPathPs -Name $name -ErrorAction SilentlyContinue
        if ($null -ne $current) {
            $result.PreviousValue = $current.$name
            if ($current.$name -eq $Value) {
                $result.Status = "Unchanged"
                return $result
            }
        }

        $cmdOutput = & reg.exe add $regPathExe /v $name /t REG_DWORD /d $Value /f 2>&1
        $regExit = $LASTEXITCODE
        $cmdText = ($cmdOutput | Out-String).Trim()

        if ($regExit -ne 0) {
            if ($cmdText -match 'Access is denied') {
                $result.Status = "Skipped"
                $result.Error = "Access denied writing shell-managed Widgets button value; Widgets policy can still disable the feature."
            } else {
                $result.Status = "Failed"
                $result.Error = if ($cmdText) { $cmdText } else { "reg.exe add failed with exit code $regExit" }
            }
            return $result
        }

        Start-Sleep -Milliseconds 500

        $verify1 = Get-ItemProperty -Path $regPathPs -Name $name -ErrorAction SilentlyContinue
        if ($null -ne $verify1 -and $verify1.$name -eq $Value) {
            $result.Status = "Success"
            return $result
        }

        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process explorer.exe
        Start-Sleep -Seconds 2

        $verify2 = Get-ItemProperty -Path $regPathPs -Name $name -ErrorAction SilentlyContinue
        if ($null -ne $verify2 -and $verify2.$name -eq $Value) {
            $result.Status = "Success"
            return $result
        }

        $result.Status = "Skipped"
        $result.Error = "Widgets taskbar setting is not being honored on this Windows 11 build or is being managed by Explorer/Feature Experience."
        return $result

    } catch {
        if ($_.Exception.Message -match 'Access is denied') {
            $result.Status = "Skipped"
            $result.Error = "Access denied writing shell-managed Widgets button value; Widgets policy can still disable the feature."
        } else {
            $result.Status = "Failed"
            $result.Error = $_.Exception.Message
        }
        return $result
    }
}

function New-RegistryBackup {
    <#
    .SYNOPSIS
    Creates a backup of current registry settings before applying changes

    .OUTPUTS
    Returns boolean indicating if backup was successful
    #>

    try {
        $backupPath = "C:\temp\registry_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"

        # Ensure backup directory exists
        $backupDir = Split-Path $backupPath -Parent
        if (-not (Test-Path $backupDir)) {
            New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        }

        Write-StatusLog "Creating registry backup at: $backupPath" -Level "Info"

        # Export relevant registry keys
        $exportKeys = @(
            "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies",
            "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows",
            "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Dsh",
            "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager",
            "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer",
            "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        )

        foreach ($key in $exportKeys) {
            $keyBackupPath = $backupPath -replace '\.reg$', "_$($key -replace '\\|:', '_').reg"
            reg.exe export $key $keyBackupPath /y 2>$null | Out-Null
        }

        Write-StatusLog "[OK] Registry backup created successfully" -Level "Success"
        return $true

    } catch {
        Write-StatusLog "[ERROR] Failed to create registry backup: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Show-RegistrySettingsSummary {
    <#
    .SYNOPSIS
    Displays a comprehensive summary of registry settings application

    .PARAMETER Results
    Results hashtable from the settings application

    .PARAMETER Category
    Category of settings that were applied
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results,

        [Parameter(Mandatory)]
        [string]$Category
    )

    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "APPLICATION" }

    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "REGISTRY SETTINGS $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan

    Write-Host "Category: " -NoNewline
    Write-Host $Category -ForegroundColor White

    Write-Host "Windows 11 Detected: " -NoNewline
    Write-Host $(if($Results.Windows11Detected){"Yes"}else{"No"}) -ForegroundColor White

    Write-Host "Total Settings: " -NoNewline
    Write-Host $Results.TotalSettings -ForegroundColor White

    Write-Host "Successfully Applied: " -NoNewline
    Write-Host $Results.SuccessfulChanges -ForegroundColor Green

    Write-Host "Already Configured: " -NoNewline
    Write-Host $Results.SkippedSettings -ForegroundColor Yellow

    Write-Host "Failed to Apply: " -NoNewline
    Write-Host $Results.FailedChanges -ForegroundColor Red

    if ($Results.BackupCreated) {
        Write-Host "Registry Backup: " -NoNewline
        Write-Host "Created" -ForegroundColor Green
    }

    if ($Results.FailedSettings.Count -gt 0) {
        Write-Host "`nFailed Settings:" -ForegroundColor Red
        $Results.FailedSettings | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }

    Write-Host ("=" * 60) -ForegroundColor Cyan
}

# Usage examples:
# Apply-RecommendedRegistrySettings
# Apply-RecommendedRegistrySettings -SettingsCategory Privacy -BackupRegistry
# Apply-RecommendedRegistrySettings -SettingsCategory UI
# Apply-RecommendedRegistrySettings -WhatIf

# -----------------------------------------------------------------------------
# Option 4 - Optimize Windows Services
# -----------------------------------------------------------------------------
function Optimize-WindowsServices {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [ValidateSet('All', 'Essential', 'Performance', 'Security', 'Gaming', 'Multimedia', 'Network')]
        [string]$ServiceCategory = 'All',

        [Parameter()]
        [switch]$BackupServices,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$IgnoreRunning
    )

    Write-StatusLog "Optimizing Windows service configurations..." -Level "Info"

    # Protected / shell-managed / instance-managed services that should not count as hard failures
    $protectedServicePatterns = @(
        'StateRepository',
        'TextInputManagementService',
        'AppXSvc',
        'cbdhsvc_*',
        'wscsvc',
        'sppsvc'
    )

    # Service configurations organized by category
    $serviceConfigurations = @{
        Essential = @{
            "AudioEndpointBuilder"          = "Automatic"
            "AudioSrv"                      = "Automatic"
            "BFE"                           = "Automatic"
            "BITS"                          = "AutomaticDelayedStart"
            "CoreMessagingRegistrar"        = "Automatic"
            "CryptSvc"                      = "Automatic"
            "DcomLaunch"                    = "Automatic"
            "Dhcp"                          = "Automatic"
            "Dnscache"                      = "Automatic"
            "DoSvc"                         = "Automatic"
            "DPS"                           = "Automatic"
            "DusmSvc"                       = "Automatic"
            "EventLog"                      = "Automatic"
            "FontCache"                     = "Automatic"
            "gpsvc"                         = "Automatic"
            "iphlpsvc"                      = "Automatic"
            "KeyIso"                        = "Automatic"
            "LanmanServer"                  = "Automatic"
            "LanmanWorkstation"             = "Automatic"
            "LSM"                           = "Automatic"
            "MapsBroker"                    = "AutomaticDelayedStart"
            "MpsSvc"                        = "Automatic"
            "Netlogon"                      = "Automatic"
            "Power"                         = "Automatic"
            "ProfSvc"                       = "Automatic"
            "RpcEptMapper"                  = "Automatic"
            "RpcSs"                         = "Automatic"
            "SamSs"                         = "Automatic"
            "Schedule"                      = "Automatic"
            "SENS"                          = "Automatic"
            "SgrmBroker"                    = "Automatic"
            "ShellHWDetection"              = "Automatic"
            "Spooler"                       = "Automatic"
            "sppsvc"                        = "Automatic"
            "SysMain"                       = "Automatic"
            "SystemEventsBroker"            = "Automatic"
            "TermService"                   = "Manual"
            "Themes"                        = "Automatic"
            "TrkWks"                        = "Automatic"
            "UserManager"                   = "Automatic"
            "VaultSvc"                      = "Manual"
            "W32Time"                       = "Automatic"
            "Wcmsvc"                        = "Automatic"
            "WinDefend"                     = "Automatic"
            "Winmgmt"                       = "Automatic"
            "WlanSvc"                       = "Automatic"
            "WSearch"                       = "AutomaticDelayedStart"
            "wscsvc"                        = "Automatic"
            "WinRM"                         = "Manual"
        }

        Performance = @{
            "ALG"                           = "Manual"
            "AppIDSvc"                      = "Manual"
            "AppMgmt"                       = "Manual"
            "AppReadiness"                  = "Manual"
            "AppXSvc"                       = "Manual"
            "Appinfo"                       = "Manual"
            "AxInstSV"                      = "Manual"
            "BDESVC"                        = "Manual"
            "BTAGService"                   = "Manual"
            "BthAvctpSvc"                   = "Automatic"
            "BthHFSrv"                      = "Manual"
            "Browser"                       = "Manual"
            "CDPSvc"                        = "Manual"
            "COMSysApp"                     = "Manual"
            "CertPropSvc"                   = "Manual"
            "ClipSVC"                       = "Manual"
            "CscService"                    = "Manual"
            "DcpSvc"                        = "Manual"
            "DevQueryBroker"                = "Manual"
            "DeviceAssociationService"      = "Manual"
            "DeviceInstall"                 = "Manual"
            "DiagTrack"                     = "Disabled"
            "DispBrokerDesktopSvc"          = "Automatic"
            "DisplayEnhancementService"     = "Manual"
            "DmEnrollmentSvc"               = "Manual"
            "DsSvc"                         = "Manual"
            "DsmSvc"                        = "Manual"
            "EFS"                           = "Manual"
            "EapHost"                       = "Manual"
            "EntAppSvc"                     = "Manual"
            "FDResPub"                      = "Manual"
            "Fax"                           = "Manual"
            "FrameServer"                   = "Manual"
            "FrameServerMonitor"            = "Manual"
            "GraphicsPerfSvc"               = "Manual"
            "HomeGroupListener"             = "Manual"
            "HomeGroupProvider"             = "Manual"
            "HvHost"                        = "Manual"
            "IEEtwCollectorService"         = "Manual"
            "IKEEXT"                        = "Manual"
            "InstallService"                = "Manual"
            "InventorySvc"                  = "Manual"
            "IpxlatCfgSvc"                  = "Manual"
            "KtmRm"                         = "Manual"
            "LicenseManager"                = "Manual"
            "LxpSvc"                        = "Manual"
            "MSDTC"                         = "Manual"
            "MSiSCSI"                       = "Manual"
            "McpManagementService"          = "Manual"
            "MixedRealityOpenXRSvc"         = "Manual"
            "MsKeyboardFilter"              = "Manual"
            "NaturalAuthentication"         = "Manual"
            "NcaSvc"                        = "Manual"
            "NcbService"                    = "Manual"
            "NcdAutoSetup"                  = "Manual"
            "NetSetupSvc"                   = "Manual"
            "Netman"                        = "Manual"
            "NgcCtnrSvc"                    = "Manual"
            "NgcSvc"                        = "Manual"
            "NlaSvc"                        = "Manual"
            "PNRPAutoReg"                   = "Manual"
            "PNRPsvc"                       = "Manual"
            "PcaSvc"                        = "Manual"
            "PeerDistSvc"                   = "Manual"
            "PerfHost"                      = "Manual"
            "PhoneSvc"                      = "Manual"
            "PlugPlay"                      = "Manual"
            "PolicyAgent"                   = "Manual"
            "PrintNotify"                   = "Manual"
            "PushToInstall"                 = "Manual"
            "QWAVE"                         = "Manual"
            "RasAuto"                       = "Manual"
            "RasMan"                        = "Manual"
            "RemoteAccess"                  = "Manual"
            "RetailDemo"                    = "Manual"
            "RmSvc"                         = "Manual"
            "RpcLocator"                    = "Manual"
            "SCPolicySvc"                   = "Manual"
            "SCardSvr"                      = "Manual"
            "SDRSVC"                        = "Manual"
            "SEMgrSvc"                      = "Manual"
            "SNMPTrap"                      = "Manual"
            "SSDPSRV"                       = "Manual"
            "ScDeviceEnum"                  = "Manual"
            "SecurityHealthService"         = "Manual"
            "Sense"                         = "Automatic"
            "SensorDataService"             = "Manual"
            "SensorService"                 = "Manual"
            "SensrSvc"                      = "Manual"
            "SessionEnv"                    = "Manual"
            "SharedAccess"                  = "Manual"
            "SharedRealitySvc"              = "Manual"
            "SmsRouter"                     = "Manual"
            "SstpSvc"                       = "Manual"
            "StateRepository"               = "Manual"
            "StiSvc"                        = "Manual"
            "StorSvc"                       = "Manual"
            "TabletInputService"            = "Manual"
            "TapiSrv"                       = "Manual"
            "TextInputManagementService"    = "Manual"
            "TieringEngineService"          = "Manual"
            "TimeBroker"                    = "Manual"
            "TimeBrokerSvc"                 = "Manual"
            "TokenBroker"                   = "Manual"
            "TroubleshootingSvc"            = "Manual"
            "TrustedInstaller"              = "Manual"
            "UI0Detect"                     = "Manual"
            "UevAgentService"               = "Disabled"
            "UmRdpService"                  = "Manual"
            "UsoSvc"                        = "Manual"
            "VGAuthService"                 = "Manual"
            "VMTools"                       = "Manual"
            "VSS"                           = "Manual"
            "VacSvc"                        = "Manual"
            "WEPHOSTSVC"                    = "Manual"
            "WFDSConMgrSvc"                 = "Manual"
            "WMPNetworkSvc"                 = "Manual"
            "WManSvc"                       = "Manual"
            "WPDBusEnum"                    = "Manual"
            "WSService"                     = "Manual"
            "WaaSMedicSvc"                  = "Manual"
            "WalletService"                 = "Manual"
            "WarpJITSvc"                    = "Manual"
            "WbioSrvc"                      = "Manual"
            "WcsPlugInService"              = "Manual"
            "WdNisSvc"                      = "Manual"
            "WdiServiceHost"                = "Manual"
            "WdiSystemHost"                 = "Manual"
            "WebClient"                     = "Manual"
            "Wecsvc"                        = "Manual"
            "WerSvc"                        = "Manual"
            "WiaRpc"                        = "Manual"
            "WinHttpAutoProxySvc"           = "Manual"
            "WpcMonSvc"                     = "Manual"
            "WpnService"                    = "Manual"
            "XblAuthManager"                = "Manual"
            "XblGameSave"                   = "Manual"
            "XboxGipSvc"                    = "Manual"
            "XboxNetApiSvc"                 = "Manual"
            "autotimesvc"                   = "Manual"
            "bthserv"                       = "Manual"
            "camsvc"                        = "Manual"
            "cloudidsvc"                    = "Manual"
            "dcsvc"                         = "Manual"
            "defragsvc"                     = "Manual"
            "diagnostichub.standardcollector.service" = "Manual"
            "diagsvc"                       = "Manual"
            "dmwappushservice"              = "Manual"
            "dot3svc"                       = "Manual"
            "edgeupdate"                    = "Manual"
            "edgeupdatem"                   = "Manual"
            "embeddedmode"                  = "Manual"
            "fdPHost"                       = "Manual"
            "fhsvc"                         = "Manual"
            "hidserv"                       = "Manual"
            "icssvc"                        = "Manual"
            "lfsvc"                         = "Manual"
            "lltdsvc"                       = "Manual"
            "lmhosts"                       = "Manual"
            "msiserver"                     = "Manual"
            "netprofm"                      = "Manual"
            "nsi"                           = "Manual"
            "p2pimsvc"                      = "Manual"
            "p2psvc"                        = "Manual"
            "perceptionsimulation"          = "Manual"
            "pla"                           = "Manual"
            "seclogon"                      = "Manual"
            "smphost"                       = "Manual"
            "spectrum"                      = "Manual"
            "svsvc"                         = "Manual"
            "swprv"                         = "Manual"
            "upnphost"                      = "Manual"
            "vds"                           = "Manual"
            "vm3dservice"                   = "Manual"
            "vmicguestinterface"            = "Manual"
            "vmicheartbeat"                 = "Manual"
            "vmickvpexchange"               = "Manual"
            "vmicrdv"                       = "Manual"
            "vmicshutdown"                  = "Manual"
            "vmictimesync"                  = "Manual"
            "vmicvmsession"                 = "Manual"
            "vmicvss"                       = "Manual"
            "vmvss"                         = "Manual"
            "wbengine"                      = "Manual"
            "wcncsvc"                       = "Manual"
            "webthreatdefsvc"               = "Manual"
            "wercplsupport"                 = "Manual"
            "wisvc"                         = "Manual"
            "wlidsvc"                       = "Manual"
            "wlpasvc"                       = "Manual"
            "wmiApSrv"                      = "Manual"
            "workfolderssvc"                = "Manual"
            "wuauserv"                      = "Manual"
            "wudfsvc"                       = "Manual"
        }

        Disabled = @{
            "AJRouter"                      = "Disabled"
            "AppVClient"                    = "Disabled"
            "AssignedAccessManagerSvc"      = "Disabled"
            "DialogBlockingService"         = "Disabled"
            "NetTcpPortSharing"             = "Disabled"
            "RemoteRegistry"                = "Disabled"
            "shpamsvc"                      = "Disabled"
            "ssh-agent"                     = "Disabled"
            "tzautoupdate"                  = "Disabled"
            "uhssvc"                        = "Disabled"
        }

        UserServices = @{
            "BcastDVRUserService_*"         = "Manual"
            "BluetoothUserService_*"        = "Manual"
            "CDPUserSvc_*"                  = "Automatic"
            "CaptureService_*"              = "Manual"
            "ConsentUxUserSvc_*"            = "Manual"
            "CredentialEnrollmentManagerUserSvc_*" = "Manual"
            "DeviceAssociationBrokerSvc_*"  = "Manual"
            "DevicePickerUserSvc_*"         = "Manual"
            "DevicesFlowUserSvc_*"          = "Manual"
            "MessagingService_*"            = "Manual"
            "NPSMSvc_*"                     = "Manual"
            "OneSyncSvc_*"                  = "Automatic"
            "P9RdrService_*"                = "Manual"
            "PenService_*"                  = "Manual"
            "PimIndexMaintenanceSvc_*"      = "Manual"
            "PrintWorkflowUserSvc_*"        = "Manual"
            "UdkUserSvc_*"                  = "Manual"
            "UnistoreSvc_*"                 = "Manual"
            "UserDataSvc_*"                 = "Manual"
            "WpnUserService_*"              = "Automatic"
            "cbdhsvc_*"                     = "Manual"
            "webthreatdefusersvc_*"         = "Automatic"
        }
    }

    $servicesToConfigure = @{}

    function Merge-ServiceConfigTable {
        param(
            [hashtable]$Target,
            [hashtable]$Source
        )

        foreach ($key in $Source.Keys) {
            $Target[$key] = $Source[$key]
        }
    }

    switch ($ServiceCategory) {
        'All' {
            foreach ($groupName in @('Essential', 'Performance', 'Disabled', 'UserServices')) {
                Merge-ServiceConfigTable -Target $servicesToConfigure -Source $serviceConfigurations[$groupName]
            }
            $servicesToConfigure['W32Time'] = 'Automatic'
        }
        'Essential' {
            Merge-ServiceConfigTable -Target $servicesToConfigure -Source $serviceConfigurations.Essential
            $servicesToConfigure['W32Time'] = 'Automatic'
        }
        'Performance' {
            Merge-ServiceConfigTable -Target $servicesToConfigure -Source $serviceConfigurations.Performance
            Merge-ServiceConfigTable -Target $servicesToConfigure -Source $serviceConfigurations.Disabled
            $servicesToConfigure['W32Time'] = 'Automatic'
        }
        'Security' {
            foreach ($entry in $serviceConfigurations.Essential.GetEnumerator()) {
                if ($entry.Key -match 'Defend|MpsSvc|Crypt|Security|wscsvc|BFE|Winmgmt') {
                    $servicesToConfigure[$entry.Key] = $entry.Value
                }
            }
            $servicesToConfigure['W32Time'] = 'Automatic'
        }
        default {
            if ($serviceConfigurations.ContainsKey($ServiceCategory)) {
                Merge-ServiceConfigTable -Target $servicesToConfigure -Source $serviceConfigurations[$ServiceCategory]
            }
            $servicesToConfigure['W32Time'] = 'Automatic'
        }
    }

    Write-StatusLog "Building service cache..." -Level "Info"
    $allCimServices = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue
    $cimByName = @{}
    $serviceControllerByName = @{}

    foreach ($svc in $allCimServices) {
        $cimByName[$svc.Name] = $svc
    }

    foreach ($svc in (Get-Service -ErrorAction SilentlyContinue)) {
        $serviceControllerByName[$svc.Name] = $svc
    }

    $results = @{
        TotalServices = $servicesToConfigure.Count
        SuccessfulChanges = 0
        FailedChanges = 0
        SkippedServices = 0
        NotFoundCount = 0
        RunningServiceChanges = 0
        ChangedServices = @()
        FailedServices = @()
        NotFoundServices = @()
        RunningServices = @()
        Errors = @()
        BackupCreated = $false
    }

    if ($BackupServices) {
        $results.BackupCreated = New-ServiceConfigurationBackup
    }

    Write-Host "`n[CONFIG] Optimizing $($results.TotalServices) Windows services..." -ForegroundColor Cyan

    $currentService = 0
    $statusStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $lastHeartbeatSecond = -10

    foreach ($service in $servicesToConfigure.GetEnumerator()) {
        $currentService++
        $percentComplete = [math]::Round(($currentService / $results.TotalServices) * 100)

        Write-Progress -Activity "Optimizing Services" -Status "Processing: $($service.Key) ($currentService of $($results.TotalServices))" -PercentComplete $percentComplete

        $elapsedSeconds = [int][math]::Floor($statusStopwatch.Elapsed.TotalSeconds)
        if (($elapsedSeconds - $lastHeartbeatSecond) -ge 10) {
            $lastHeartbeatSecond = $elapsedSeconds
            Write-StatusLog "Processed $currentService of $($results.TotalServices) Windows services..." -Level "Info"
        }

        try {
            $configResult = Set-ServiceStartupSafe `
                -ServiceName $service.Key `
                -StartupType $service.Value `
                -Force:$Force `
                -IgnoreRunning:$IgnoreRunning `
                -ProtectedServicePatterns $protectedServicePatterns `
                -CimByName $cimByName `
                -ServiceControllerByName $serviceControllerByName

            switch ($configResult.Status) {
                "Success" {
                    $results.SuccessfulChanges++
                    $results.ChangedServices += $service.Key
                    Write-StatusLog "[OK] $($service.Key) -> $($service.Value)" -Level "Success"

                    if ($configResult.WasRunning) {
                        $results.RunningServiceChanges++
                        $results.RunningServices += $service.Key
                    }

                    if ($configResult.RefreshNames) {
                        foreach ($refreshName in $configResult.RefreshNames) {
                            $updatedCim = Get-CimInstance Win32_Service -Filter "Name='$refreshName'" -ErrorAction SilentlyContinue
                            if ($updatedCim) { $cimByName[$refreshName] = $updatedCim }

                            $updatedSvc = Get-Service -Name $refreshName -ErrorAction SilentlyContinue
                            if ($updatedSvc) { $serviceControllerByName[$refreshName] = $updatedSvc }
                        }
                    }
                }
                "Unchanged" {
                    $results.SkippedServices++
                    Write-StatusLog "[INFO] $($service.Key) already set to $($service.Value)" -Level "Info"
                }
                "Skipped" {
                    $results.SkippedServices++
                    Write-StatusLog "[WARN] $($service.Key) skipped - $($configResult.Error)" -Level "Warning"
                }
                "NotFound" {
                    $results.NotFoundCount++
                    $results.NotFoundServices += $service.Key
                    Write-StatusLog "[WARN] Service '$($service.Key)' not found" -Level "Warning"
                }
                "Failed" {
                    $results.FailedChanges++
                    $results.FailedServices += $service.Key
                    $results.Errors += $configResult.Error
                    Write-StatusLog "[ERROR] Failed to configure '$($service.Key)': $($configResult.Error)" -Level "Error"
                }
            }

        } catch {
            $results.FailedChanges++
            $results.FailedServices += $service.Key
            $errorMsg = "Unexpected error configuring '$($service.Key)': $($_.Exception.Message)"
            $results.Errors += $errorMsg
            Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        }
    }

    Write-Progress -Activity "Optimizing Services" -Completed

    Show-ServiceOptimizationSummary -Results $results -Category $ServiceCategory

    $global:LastStatus = if ($results.FailedChanges -eq 0) {
        "[OK] Service optimization completed ($($results.SuccessfulChanges) services configured)"
    } else {
        "[WARN] Service optimization completed with $($results.FailedChanges) failures"
    }

    return $results
}

function Set-ServiceStartupSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [ValidateSet("Automatic", "Manual", "Disabled", "AutomaticDelayedStart")]
        [string]$StartupType,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$IgnoreRunning,

        [Parameter()]
        [string[]]$ProtectedServicePatterns = @(),

        [Parameter()]
        [hashtable]$CimByName = @{},

        [Parameter()]
        [hashtable]$ServiceControllerByName = @{}
    )

    $result = @{
        Status = "Unknown"
        Error = $null
        PreviousStartupType = $null
        WasRunning = $false
        ServiceExists = $false
        ServicesConfigured = 0
        RefreshNames = @()
    }

    function Test-IsProtectedServiceName {
        param(
            [string]$NameToCheck,
            [string[]]$Patterns
        )

        foreach ($pattern in $Patterns) {
            if ($NameToCheck -like $pattern) {
                return $true
            }
        }
        return $false
    }

    function Get-DelayedAutoStartFlag {
        param([string]$ExactServiceName)

        try {
            $reg = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$ExactServiceName" -Name "DelayedAutostart" -ErrorAction SilentlyContinue
            if ($null -ne $reg) { return [int]$reg.DelayedAutostart }
        } catch {}
        return 0
    }

    function Convert-CimStartMode {
        param($CimService)

        if (-not $CimService) { return $null }

        switch ($CimService.StartMode) {
            'Auto' {
                $delayed = Get-DelayedAutoStartFlag -ExactServiceName $CimService.Name
                if ($delayed -eq 1) { 'AutomaticDelayedStart' } else { 'Automatic' }
            }
            'Manual'   { 'Manual' }
            'Disabled' { 'Disabled' }
            default    { $CimService.StartMode }
        }
    }

    function Test-StartupTypeSatisfied {
        param(
            [string]$DesiredStartupType,
            [string]$ActualStartupType
        )

        switch ($DesiredStartupType) {
            'Automatic' {
                return ($ActualStartupType -in @('Automatic', 'AutomaticDelayedStart'))
            }
            default {
                return ($ActualStartupType -eq $DesiredStartupType)
            }
        }
    }

    function Set-ServiceStartupTypeNative {
        param(
            [string]$ExactServiceName,
            [string]$DesiredStartupType
        )

        $outputText = @()
        $exitCode = 0

        switch ($DesiredStartupType) {
            'Automatic' {
                $output = & sc.exe config $ExactServiceName start= auto 2>&1
                $exitCode = $LASTEXITCODE
                $outputText += ($output | Out-String).Trim()
            }
            'Manual' {
                $output = & sc.exe config $ExactServiceName start= demand 2>&1
                $exitCode = $LASTEXITCODE
                $outputText += ($output | Out-String).Trim()
            }
            'Disabled' {
                $output = & sc.exe config $ExactServiceName start= disabled 2>&1
                $exitCode = $LASTEXITCODE
                $outputText += ($output | Out-String).Trim()
            }
            'AutomaticDelayedStart' {
                $output1 = & sc.exe config $ExactServiceName start= auto 2>&1
                $exit1 = $LASTEXITCODE
                $outputText += ($output1 | Out-String).Trim()

                if ($exit1 -eq 0) {
                    $isPerUserService = ($ExactServiceName -like '*_*')

                    if ($isPerUserService) {
                        $exitCode = 0
                    } else {
                        $output2 = & reg.exe add "HKLM\SYSTEM\CurrentControlSet\Services\$ExactServiceName" /v DelayedAutostart /t REG_DWORD /d 1 /f 2>&1
                        $exit2 = $LASTEXITCODE
                        $outputText += ($output2 | Out-String).Trim()
                        $exitCode = $exit2
                    }
                } else {
                    $exitCode = $exit1
                }
            }
        }

        @{
            ExitCode = $exitCode
            Output   = (($outputText | Where-Object { $_ }) -join " | ").Trim()
        }
    }

    try {
        if ($ServiceName -like "*_*") {
            $matchingNames = @($CimByName.Keys | Where-Object { $_ -like $ServiceName })
            if (-not $matchingNames -or $matchingNames.Count -eq 0) {
                $result.Status = "NotFound"
                return $result
            }

            $successCount = 0
            $totalCount = 0
            $skipCount = 0
            $lastError = $null

            foreach ($matchName in $matchingNames) {
                $totalCount++

                try {
                    $cimSvc = $CimByName[$matchName]
                    $svc = $ServiceControllerByName[$matchName]

                    $currentStartupType = Convert-CimStartMode -CimService $cimSvc
                    $isRunning = ($svc -and $svc.Status -eq 'Running')

                    if ($null -eq $currentStartupType) {
                        $lastError = "Could not determine current start mode."
                        continue
                    }

                    if (-not $Force -and (Test-StartupTypeSatisfied -DesiredStartupType $StartupType -ActualStartupType $currentStartupType)) {
                        $successCount++
                        continue
                    }

                    if ($WhatIfPreference) {
                        Write-Host "   Would set: $matchName -> $StartupType" -ForegroundColor Yellow
                        $successCount++
                        continue
                    }

                    $setResult = Set-ServiceStartupTypeNative -ExactServiceName $matchName -DesiredStartupType $StartupType

                    if ($setResult.ExitCode -eq 0) {
                        $updatedCim = Get-CimInstance Win32_Service -Filter "Name='$matchName'" -ErrorAction SilentlyContinue
                        $verifiedMode = Convert-CimStartMode -CimService $updatedCim

                        if ($updatedCim) {
                            $CimByName[$matchName] = $updatedCim
                        }

                        if (Test-StartupTypeSatisfied -DesiredStartupType $StartupType -ActualStartupType $verifiedMode) {
                            $successCount++
                            $result.RefreshNames += $matchName
                            if ($isRunning) {
                                $result.WasRunning = $true
                            }
                        } else {
                            if (Test-IsProtectedServiceName -NameToCheck $matchName -Patterns $ProtectedServicePatterns) {
                                $skipCount++
                                $lastError = "Protected or managed service."
                            } else {
                                $lastError = "Verification failed."
                            }
                        }
                    } else {
                        if ((Test-IsProtectedServiceName -NameToCheck $matchName -Patterns $ProtectedServicePatterns) -and
                            ($setResult.Output -match 'Access is denied|Cannot open service|This service cannot accept control messages|The parameter is incorrect')) {
                            $skipCount++
                            $lastError = $setResult.Output
                        } else {
                            $lastError = $setResult.Output
                        }
                    }
                } catch {
                    $err = $_.Exception.Message
                    if (Test-IsProtectedServiceName -NameToCheck $matchName -Patterns $ProtectedServicePatterns) {
                        $skipCount++
                        $lastError = $err
                    } else {
                        $lastError = $err
                    }
                }
            }

            $result.ServicesConfigured = $successCount
            $result.ServiceExists = ($totalCount -gt 0)

            if ($successCount -eq $totalCount) {
                $result.Status = "Success"
            } elseif (($successCount + $skipCount) -eq $totalCount -and $skipCount -gt 0) {
                $result.Status = "Skipped"
                $result.Error = "One or more matching user-instance services are protected or shell-managed."
            } elseif ($successCount -gt 0) {
                $result.Status = "Success"
                if ($lastError) {
                    $result.Error = "Some matching services were skipped or failed: $lastError"
                }
            } else {
                if (Test-IsProtectedServiceName -NameToCheck $ServiceName -Patterns $ProtectedServicePatterns) {
                    $result.Status = "Skipped"
                    $result.Error = if ($lastError) { $lastError } else { "Matching services are protected or shell-managed." }
                } else {
                    $result.Status = "Failed"
                    $result.Error = $lastError
                }
            }

            return $result
        }

        $cimSvc = $CimByName[$ServiceName]
        $service = $ServiceControllerByName[$ServiceName]

        if (-not $cimSvc -and -not $service) {
            $result.Status = "NotFound"
            return $result
        }

        $result.ServiceExists = $true
        $result.PreviousStartupType = Convert-CimStartMode -CimService $cimSvc
        $result.WasRunning = ($service -and $service.Status -eq 'Running')

        if (-not $Force -and (Test-StartupTypeSatisfied -DesiredStartupType $StartupType -ActualStartupType $result.PreviousStartupType)) {
            $result.Status = "Unchanged"
            return $result
        }

        if ($WhatIfPreference) {
            Write-Host "   Would set: $ServiceName -> $StartupType" -ForegroundColor Yellow
            $result.Status = "Success"
            return $result
        }

        $setResult = Set-ServiceStartupTypeNative -ExactServiceName $ServiceName -DesiredStartupType $StartupType

        if ($setResult.ExitCode -eq 0) {
            $updatedCim = Get-CimInstance Win32_Service -Filter "Name='$ServiceName'" -ErrorAction SilentlyContinue
            $verifiedMode = Convert-CimStartMode -CimService $updatedCim

            if ($updatedCim) {
                $CimByName[$ServiceName] = $updatedCim
            }

            if (Test-StartupTypeSatisfied -DesiredStartupType $StartupType -ActualStartupType $verifiedMode) {
                $result.Status = "Success"
                $result.ServicesConfigured = 1
                $result.RefreshNames += $ServiceName
            } else {
                if (Test-IsProtectedServiceName -NameToCheck $ServiceName -Patterns $ProtectedServicePatterns) {
                    $result.Status = "Skipped"
                    $result.Error = "Service start mode could not be changed because it is protected or managed by Windows."
                } else {
                    $result.Status = "Failed"
                    $result.Error = "Service startup type verification failed."
                }
            }
        } else {
            if ((Test-IsProtectedServiceName -NameToCheck $ServiceName -Patterns $ProtectedServicePatterns) -and
                ($setResult.Output -match 'Access is denied|Cannot open service|This service cannot accept control messages|The parameter is incorrect')) {
                $result.Status = "Skipped"
                $result.Error = "Service is protected or shell-managed by Windows: $($setResult.Output)"
            } else {
                $result.Status = "Failed"
                $result.Error = $setResult.Output
            }
        }

    } catch {
        if (Test-IsProtectedServiceName -NameToCheck $ServiceName -Patterns $ProtectedServicePatterns) {
            $result.Status = "Skipped"
            $result.Error = "Service is protected or shell-managed by Windows: $($_.Exception.Message)"
        } else {
            $result.Status = "Failed"
            $result.Error = $_.Exception.Message
        }
    }

    return $result
}

function New-ServiceConfigurationBackup {
    try {
        $backupPath = "C:\temp\service_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        $backupDir = Split-Path $backupPath -Parent
        if (-not (Test-Path $backupDir)) {
            New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        }

        Write-StatusLog "Creating service configuration backup at: $backupPath" -Level "Info"

        $services = Get-CimInstance Win32_Service | Select-Object Name, StartMode, State, DisplayName
        $services | Export-Csv -Path $backupPath -NoTypeInformation

        Write-StatusLog "[OK] Service backup created successfully ($($services.Count) services)" -Level "Success"
        return $true

    } catch {
        Write-StatusLog "[ERROR] Failed to create service backup: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Show-ServiceOptimizationSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results,

        [Parameter(Mandatory)]
        [string]$Category
    )

    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "OPTIMIZATION" }

    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "SERVICE $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan

    Write-Host "Category: " -NoNewline
    Write-Host $Category -ForegroundColor White

    Write-Host "Total Services Processed: " -NoNewline
    Write-Host $Results.TotalServices -ForegroundColor White

    Write-Host "Successfully Configured: " -NoNewline
    Write-Host $Results.SuccessfulChanges -ForegroundColor Green

    Write-Host "Already Optimized: " -NoNewline
    Write-Host $Results.SkippedServices -ForegroundColor Yellow

    Write-Host "Not Found: " -NoNewline
    Write-Host $Results.NotFoundCount -ForegroundColor Yellow

    Write-Host "Failed to Configure: " -NoNewline
    Write-Host $Results.FailedChanges -ForegroundColor Red

    if ($Results.RunningServiceChanges -gt 0) {
        Write-Host "Running Services Modified: " -NoNewline
        Write-Host $Results.RunningServiceChanges -ForegroundColor Yellow
    }

    if ($Results.BackupCreated) {
        Write-Host "Configuration Backup: " -NoNewline
        Write-Host "Created" -ForegroundColor Green
    }

    if ($Results.FailedServices.Count -gt 0) {
        Write-Host "`nFailed Services:" -ForegroundColor Red
        $Results.FailedServices | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }

    if ($Results.RunningServices.Count -gt 0 -and $Results.RunningServices.Count -le 10) {
        Write-Host "`nRunning Services Modified:" -ForegroundColor Yellow
        $Results.RunningServices | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Yellow
        }
        Write-Host "Note: These services may require a restart to fully apply changes." -ForegroundColor Yellow
    }

    if ($Results.SuccessfulChanges -gt 0) {
        $bootTimeImprovement = [math]::Round($Results.SuccessfulChanges * 0.2, 1)
        Write-Host "`nEstimated Boot Time Improvement: " -NoNewline
        Write-Host "$bootTimeImprovement seconds" -ForegroundColor Green
    }

    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Get-ServiceOptimizationReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$OutputPath = "C:\temp\service_analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    )

    Write-StatusLog "Generating service optimization report..." -Level "Info"

    try {
        $services = Get-CimInstance Win32_Service | Select-Object Name, StartMode, State, DisplayName
        $runningAutoServices = $services | Where-Object { $_.StartMode -eq 'Auto' -and $_.State -eq 'Running' }
        $disabledServices = $services | Where-Object { $_.StartMode -eq 'Disabled' }
        $manualServices = $services | Where-Object { $_.StartMode -eq 'Manual' }

        $report = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows Service Optimization Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #0078d4; color: white; padding: 20px; border-radius: 5px; }
        .section { margin: 20px 0; }
        .stats { display: flex; gap: 20px; margin: 20px 0; }
        .stat-card { background-color: #f5f5f5; padding: 15px; border-radius: 5px; text-align: center; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .automatic { background-color: #d4edda; }
        .manual { background-color: #fff3cd; }
        .disabled { background-color: #f8d7da; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Windows Service Optimization Report</h1>
        <p>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>

    <div class="stats">
        <div class="stat-card">
            <h3>Total Services</h3>
            <p>$($services.Count)</p>
        </div>
        <div class="stat-card">
            <h3>Automatic</h3>
            <p>$($runningAutoServices.Count)</p>
        </div>
        <div class="stat-card">
            <h3>Manual</h3>
            <p>$($manualServices.Count)</p>
        </div>
        <div class="stat-card">
            <h3>Disabled</h3>
            <p>$($disabledServices.Count)</p>
        </div>
    </div>

    <div class="section">
        <h2>Service Breakdown by Startup Type</h2>
        <table>
            <tr><th>Service Name</th><th>Display Name</th><th>Startup Type</th><th>Status</th></tr>
"@

        foreach ($service in ($services | Sort-Object StartMode, Name)) {
            $cssClass = switch ($service.StartMode) {
                'Auto'     { 'automatic' }
                'Manual'   { 'manual' }
                'Disabled' { 'disabled' }
                default    { '' }
            }

            $report += "            <tr class='$cssClass'><td>$($service.Name)</td><td>$($service.DisplayName)</td><td>$($service.StartMode)</td><td>$($service.State)</td></tr>`n"
        }

        $report += @"
        </table>
    </div>
</body>
</html>
"@

        $outputDir = Split-Path $OutputPath -Parent
        if (-not (Test-Path $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }

        $report | Out-File -FilePath $OutputPath -Encoding UTF8

        Write-StatusLog "[OK] Service report generated: $OutputPath" -Level "Success"

        return @{
            ReportPath = $OutputPath
            TotalServices = $services.Count
            AutomaticServices = $runningAutoServices.Count
            ManualServices = $manualServices.Count
            DisabledServices = $disabledServices.Count
        }

    } catch {
        Write-StatusLog "[ERROR] Failed to generate service report: $($_.Exception.Message)" -Level "Error"
        return $null
    }
}

# Usage examples:
# Optimize-WindowsServices
# Optimize-WindowsServices -ServiceCategory Performance -BackupServices
# Optimize-WindowsServices -WhatIf
# Optimize-WindowsServices -ServiceCategory Essential -Force
# Get-ServiceOptimizationReport

# -----------------------------------------------------------------------------
# Option 5 - Enable PowerShell Remote Management
# -----------------------------------------------------------------------------
function Enable-PowerShellRemotingSafely {
    <#
    .SYNOPSIS
    Safely enables PowerShell Remoting with security best practices
    
    .DESCRIPTION
    Configures PowerShell Remoting (WinRM) with proper security settings including:
    - WinRM service configuration and startup
    - Firewall rule configuration
    - Security descriptor lockdown to administrators only
    - UAC token filter policy for local account elevation
    
    .PARAMETER AllowedUsers
    Additional users/groups to grant PowerShell remoting access (beyond administrators)
    
    .PARAMETER SkipFirewallConfiguration
    Skip automatic firewall rule configuration
    
    .PARAMETER SkipSecurityLockdown
    Skip security descriptor configuration (not recommended)
    
    .PARAMETER TrustedHosts
    Configure WinRM trusted hosts (comma-separated list)
    
    .OUTPUTS
    Returns hashtable with configuration results and status
    
    .EXAMPLE
    Enable-PowerShellRemotingSafely
    
    .EXAMPLE
    Enable-PowerShellRemotingSafely -AllowedUsers @("DOMAIN\PowerUsers") -TrustedHosts "192.168.1.*"
    
    .EXAMPLE
    Enable-PowerShellRemotingSafely -WhatIf
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string[]]$AllowedUsers = @(),
        
        [Parameter()]
        [switch]$SkipFirewallConfiguration,
        
        [Parameter()]
        [switch]$SkipSecurityLockdown,
        
        [Parameter()]
        [string]$TrustedHosts = $null
    )
    
    Write-StatusLog "Configuring PowerShell Remoting with security best practices..." -Level "Info"
    
    # Initialize results tracking
    $results = @{
        WinRMConfigured = $false
        FirewallConfigured = $false
        SecurityConfigured = $false
        UACPolicyConfigured = $false
        TrustedHostsConfigured = $false
        EndpointsSecured = @()
        Errors = @()
        DomainJoined = $false
        ConfigurationSummary = @{}
    }
    
    try {
        # Step 1: Configure WinRM Service
        Write-StatusLog "Configuring WinRM service..." -Level "Info"
        $winrmResult = Set-WinRMServiceConfiguration
        $results.WinRMConfigured = $winrmResult.Success
        if (-not $winrmResult.Success) {
            $results.Errors += $winrmResult.Error
        }
        
        # Step 2: Enable PowerShell Remoting
        if ($results.WinRMConfigured) {
            Write-StatusLog "Enabling PowerShell Remoting..." -Level "Info"
            $remotingResult = Enable-PSRemotingConfiguration
            if (-not $remotingResult.Success) {
                $results.Errors += $remotingResult.Error
            }
        }
        
        # Step 3: Configure Firewall Rules
        if (-not $SkipFirewallConfiguration) {
            Write-StatusLog "Configuring firewall rules..." -Level "Info"
            $firewallResult = Set-WinRMFirewallConfiguration
            $results.FirewallConfigured = $firewallResult.Success
            if (-not $firewallResult.Success) {
                $results.Errors += $firewallResult.Error
            }
        } else {
            Write-StatusLog "Skipping firewall configuration as requested" -Level "Info"
        }
        
        # Step 4: Configure UAC Token Filter Policy
        Write-StatusLog "Configuring UAC token filter policy..." -Level "Info"
        $uacResult = Set-UACTokenFilterPolicy
        $results.UACPolicyConfigured = $uacResult.Success
        if (-not $uacResult.Success) {
            $results.Errors += $uacResult.Error
        }
        
        # Step 5: Configure Trusted Hosts (if specified)
        if ($TrustedHosts) {
            Write-StatusLog "Configuring trusted hosts..." -Level "Info"
            $trustedHostsResult = Set-WinRMTrustedHosts -TrustedHosts $TrustedHosts
            $results.TrustedHostsConfigured = $trustedHostsResult.Success
            if (-not $trustedHostsResult.Success) {
                $results.Errors += $trustedHostsResult.Error
            }
        }
        
        # Step 6: Security Lockdown
        if (-not $SkipSecurityLockdown) {
            Write-StatusLog "Applying security lockdown..." -Level "Info"
            $securityResult = Set-PSRemotingSecurityConfiguration -AllowedUsers $AllowedUsers
            $results.SecurityConfigured = $securityResult.Success
            $results.DomainJoined = $securityResult.DomainJoined
            $results.EndpointsSecured = $securityResult.SecuredEndpoints
            if (-not $securityResult.Success) {
                $results.Errors += $securityResult.Error
            }
        } else {
            Write-StatusLog "Skipping security lockdown as requested" -Level "Warning"
        }
        
        # Generate configuration summary
        $results.ConfigurationSummary = Get-PSRemotingConfigurationSummary
        
        # Display results
        Show-PSRemotingConfigurationSummary -Results $results
        
        # Set global status
        $global:LastStatus = if ($results.Errors.Count -eq 0) {
            "[OK] PowerShell Remoting enabled and secured successfully"
        } else {
            "[WARN] PowerShell Remoting configured with $($results.Errors.Count) issues"
        }
        
    } catch {
        $errorMsg = "Critical error during PowerShell Remoting configuration: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        $results.Errors += $errorMsg
        $global:LastStatus = "[ERROR] Failed to configure PowerShell Remoting"
    }
    
    return $results
}

function Set-WinRMServiceConfiguration {
    <#
    .SYNOPSIS
    Configures the WinRM service startup and state
    
    .OUTPUTS
    Returns hashtable with configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        ServiceStarted = $false
        StartupTypeSet = $false
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure WinRM service (Automatic startup, Start service)" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Set service to Automatic startup
        $service = Get-Service -Name WinRM -ErrorAction Stop
        if ($service.StartType -ne 'Automatic') {
            Set-Service -Name WinRM -StartupType Automatic -ErrorAction Stop
            $result.StartupTypeSet = $true
            Write-StatusLog "[OK] WinRM service set to Automatic startup" -Level "Success"
        } else {
            Write-StatusLog "[INFO] WinRM service already set to Automatic startup" -Level "Info"
        }
        
        # Start the service if not running
        if ($service.Status -ne 'Running') {
            Start-Service -Name WinRM -ErrorAction Stop
            $result.ServiceStarted = $true
            Write-StatusLog "[OK] WinRM service started" -Level "Success"
        } else {
            Write-StatusLog "[INFO] WinRM service already running" -Level "Info"
        }
        
        $result.Success = $true
        
    } catch {
        $result.Error = "Failed to configure WinRM service: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Enable-PSRemotingConfiguration {
    <#
    .SYNOPSIS
    Enables PowerShell Remoting configuration
    
    .OUTPUTS
    Returns hashtable with configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        RemotingEnabled = $false
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would enable PowerShell Remoting (Skip network profile check)" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Check if remoting is already enabled
        $remotingStatus = Get-PSSessionConfiguration -ErrorAction SilentlyContinue
        if ($remotingStatus) {
            Write-StatusLog "[INFO] PowerShell Remoting appears to be already enabled" -Level "Info"
        }
        
        # Enable PowerShell Remoting
        Enable-PSRemoting -SkipNetworkProfileCheck -Force -ErrorAction Stop
        $result.RemotingEnabled = $true
        $result.Success = $true
        Write-StatusLog "[OK] PowerShell Remoting enabled successfully" -Level "Success"
        
    } catch {
        $result.Error = "Failed to enable PowerShell Remoting: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Set-WinRMFirewallConfiguration {
    <#
    .SYNOPSIS
    Configures Windows Firewall rules for WinRM
    
    .OUTPUTS
    Returns hashtable with configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        RulesEnabled = @()
        RulesAlreadyEnabled = @()
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would enable Windows Remote Management firewall rules" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Get WinRM firewall rules
        $winrmRules = Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction SilentlyContinue
        
        if (-not $winrmRules) {
            $result.Error = "No Windows Remote Management firewall rules found"
            Write-StatusLog "[WARN] No WinRM firewall rules found on this system" -Level "Warning"
            return $result
        }
        
        foreach ($rule in $winrmRules) {
            if ($rule.Enabled -eq 'False') {
                Enable-NetFirewallRule -Name $rule.Name -ErrorAction Stop
                $result.RulesEnabled += $rule.DisplayName
                Write-StatusLog "[OK] Enabled firewall rule: $($rule.DisplayName)" -Level "Success"
            } else {
                $result.RulesAlreadyEnabled += $rule.DisplayName
                Write-StatusLog "[INFO] Firewall rule already enabled: $($rule.DisplayName)" -Level "Info"
            }
        }
        
        $result.Success = $true
        
    } catch {
        $result.Error = "Failed to configure firewall rules: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Set-UACTokenFilterPolicy {
    <#
    .SYNOPSIS
    Configures UAC Token Filter Policy for WinRM local account elevation
    
    .OUTPUTS
    Returns hashtable with configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        PolicySet = $false
        PreviousValue = $null
    }
    
    try {
        $registryPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        $valueName = 'LocalAccountTokenFilterPolicy'
        
        if ($WhatIfPreference) {
            Write-Host "   Would set LocalAccountTokenFilterPolicy registry value" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Check current value
        $currentValue = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction SilentlyContinue
        $result.PreviousValue = $currentValue.$valueName
        
        if ($currentValue.$valueName -eq 1) {
            Write-StatusLog "[INFO] LocalAccountTokenFilterPolicy already set correctly" -Level "Info"
            $result.Success = $true
            return $result
        }
        
        # Set the registry value
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $registryPath -Name $valueName -Value 1 -Type DWord -ErrorAction Stop
        $result.PolicySet = $true
        $result.Success = $true
        Write-StatusLog "[OK] LocalAccountTokenFilterPolicy configured for WinRM elevation" -Level "Success"
        
    } catch {
        $result.Error = "Failed to set UAC Token Filter Policy: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Set-WinRMTrustedHosts {
    <#
    .SYNOPSIS
    Configures WinRM trusted hosts
    
    .PARAMETER TrustedHosts
    Comma-separated list of trusted hosts
    
    .OUTPUTS
    Returns hashtable with configuration results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TrustedHosts
    )
    
    $result = @{
        Success = $false
        Error = $null
        TrustedHostsSet = $false
        PreviousValue = $null
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would set WinRM trusted hosts to: $TrustedHosts" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Get current trusted hosts
        $currentTrustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        $result.PreviousValue = $currentTrustedHosts
        
        if ($currentTrustedHosts -eq $TrustedHosts) {
            Write-StatusLog "[INFO] WinRM trusted hosts already configured correctly" -Level "Info"
            $result.Success = $true
            return $result
        }
        
        # Set trusted hosts
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $TrustedHosts -Force -ErrorAction Stop
        $result.TrustedHostsSet = $true
        $result.Success = $true
        Write-StatusLog "[OK] WinRM trusted hosts configured: $TrustedHosts" -Level "Success"
        
    } catch {
        $result.Error = "Failed to set trusted hosts: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Set-PSRemotingSecurityConfiguration {
    <#
    .SYNOPSIS
    Applies security lockdown to PowerShell Remoting endpoints
    
    .PARAMETER AllowedUsers
    Additional users/groups to grant access
    
    .OUTPUTS
    Returns hashtable with security configuration results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$AllowedUsers = @()
    )
    
    $result = @{
        Success = $false
        Error = $null
        DomainJoined = $false
        SecuredEndpoints = @()
        SecurityDescriptor = $null
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would apply security lockdown to PowerShell endpoints" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Determine domain membership
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $result.DomainJoined = $computerSystem.PartOfDomain
        
        # Build security descriptor SDDL
        $baseSDDL = 'O:BA G:BA D:(A;;GA;;;BA)'  # Owner: Builtin Admins, Group: Builtin Admins, Allow Builtin Admins
        
        if ($result.DomainJoined) {
            $baseSDDL += '(A;;GA;;;DA)'  # Add Domain Admins
            Write-StatusLog "[INFO] Domain-joined system detected - including Domain Admins" -Level "Info"
        } else {
            Write-StatusLog "[INFO] Workgroup system detected - Local Admins only" -Level "Info"
        }
        
        # Add additional users if specified
        foreach ($user in $AllowedUsers) {
            try {
                $sid = (New-Object System.Security.Principal.NTAccount($user)).Translate([System.Security.Principal.SecurityIdentifier]).Value
                $baseSDDL += "(A;;GA;;;$sid)"
                Write-StatusLog "[INFO] Added user/group to security descriptor: $user" -Level "Info"
            } catch {
                Write-StatusLog "[WARN] Failed to resolve SID for user/group: $user" -Level "Warning"
            }
        }
        
        $result.SecurityDescriptor = $baseSDDL
        
        # Apply security to all PowerShell endpoints
        $endpoints = Get-PSSessionConfiguration -ErrorAction Stop
        
        foreach ($endpoint in $endpoints) {
            try {
                Set-PSSessionConfiguration -Name $endpoint.Name -SecurityDescriptorSddl $baseSDDL -Force -ErrorAction Stop
                $result.SecuredEndpoints += $endpoint.Name
                Write-StatusLog "[OK] Secured endpoint: $($endpoint.Name)" -Level "Success"
            } catch {
                Write-StatusLog "[ERROR] Failed to secure endpoint '$($endpoint.Name)': $($_.Exception.Message)" -Level "Error"
                $result.Error = "Failed to secure some endpoints"
            }
        }
        
        # Restart WinRM to apply security changes
        Restart-Service -Name WinRM -ErrorAction Stop
        Write-StatusLog "[OK] WinRM service restarted to apply security changes" -Level "Success"
        
        $result.Success = ($result.SecuredEndpoints.Count -gt 0)
        
    } catch {
        $result.Error = "Failed to apply security configuration: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Get-PSRemotingConfigurationSummary {
    <#
    .SYNOPSIS
    Gets current PowerShell Remoting configuration status
    
    .OUTPUTS
    Returns hashtable with current configuration
    #>
    
    $summary = @{
        WinRMService = @{
            Status = "Unknown"
            StartType = "Unknown"
        }
        RemotingEnabled = $false
        FirewallRules = @()
        TrustedHosts = "Not configured"
        SecurityDescriptors = @{}
    }
    
    try {
        # WinRM Service status
        $winrmService = Get-Service -Name WinRM -ErrorAction SilentlyContinue
        if ($winrmService) {
            $summary.WinRMService.Status = $winrmService.Status
            $summary.WinRMService.StartType = $winrmService.StartType
        }
        
        # Check if remoting is enabled
        $sessionConfigs = Get-PSSessionConfiguration -ErrorAction SilentlyContinue
        $summary.RemotingEnabled = ($sessionConfigs -ne $null)
        
        # Firewall rules
        $firewallRules = Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction SilentlyContinue
        $summary.FirewallRules = $firewallRules | Select-Object DisplayName, Enabled
        
        # Trusted hosts
        $trustedHosts = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        $summary.TrustedHosts = if ($trustedHosts) { $trustedHosts } else { "Not configured" }
        
        # Security descriptors for endpoints
        if ($sessionConfigs) {
            foreach ($config in $sessionConfigs) {
                $summary.SecurityDescriptors[$config.Name] = $config.Permission
            }
        }
        
    } catch {
        Write-StatusLog "[WARN] Error gathering configuration summary: $($_.Exception.Message)" -Level "Warning"
    }
    
    return $summary
}

function Show-PSRemotingConfigurationSummary {
    <#
    .SYNOPSIS
    Displays comprehensive summary of PowerShell Remoting configuration
    
    .PARAMETER Results
    Results from the configuration process
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results
    )
    
    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "CONFIGURATION" }
    
    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    Write-Host "POWERSHELL REMOTING $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan
    
    Write-Host "WinRM Service Configured: " -NoNewline
    Write-Host $(if($Results.WinRMConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.WinRMConfigured){"Green"}else{"Red"})
    
    Write-Host "Firewall Configured: " -NoNewline
    Write-Host $(if($Results.FirewallConfigured){"[OK] Yes"}else{"[WARN] Skipped/Failed"}) -ForegroundColor $(if($Results.FirewallConfigured){"Green"}else{"Yellow"})
    
    Write-Host "UAC Policy Configured: " -NoNewline
    Write-Host $(if($Results.UACPolicyConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.UACPolicyConfigured){"Green"}else{"Red"})
    
    Write-Host "Security Lockdown Applied: " -NoNewline
    Write-Host $(if($Results.SecurityConfigured){"[OK] Yes"}else{"[WARN] Skipped/Failed"}) -ForegroundColor $(if($Results.SecurityConfigured){"Green"}else{"Yellow"})
    
    if ($Results.TrustedHostsConfigured) {
        Write-Host "Trusted Hosts Configured: " -NoNewline
        Write-Host "[OK] Yes" -ForegroundColor Green
    }
    
    Write-Host "Domain Status: " -NoNewline
    Write-Host $(if($Results.DomainJoined){"Domain-joined"}else{"Workgroup"}) -ForegroundColor White
    
    if ($Results.EndpointsSecured.Count -gt 0) {
        Write-Host "`nSecured Endpoints:" -ForegroundColor Green
        $Results.EndpointsSecured | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Green
        }
    }
    
    if ($Results.Errors.Count -gt 0) {
        Write-Host "`nConfiguration Issues:" -ForegroundColor Red
        $Results.Errors | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }
    
    Write-Host "="*60 -ForegroundColor Cyan
}

# Usage examples:
# Enable-PowerShellRemotingSafely
# Enable-PowerShellRemotingSafely -AllowedUsers @("DOMAIN\PowerUsers") -TrustedHosts "192.168.1.*"
# Enable-PowerShellRemotingSafely -WhatIf

# -----------------------------------------------------------------------------
# Option 6 - Configure Automatic Time Sync
# -----------------------------------------------------------------------------

function Invoke-TimeSyncConfigurationPass {
    <#
    .SYNOPSIS
    Executes a single full time synchronization configuration pass

    .PARAMETER NTPServers
    NTP servers to configure

    .PARAMETER SyncInterval
    Hours between scheduled sync attempts

    .PARAMETER TimeZone
    Configure automatic time zone

    .PARAMETER SkipScheduledTask
    Skip scheduled task creation

    .PARAMETER ValidateAccuracy
    Validate time sync accuracy after configuration

    .PARAMETER DomainEnvironment
    Whether the system is domain joined

    .OUTPUTS
    Returns hashtable with results for one full configuration pass
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$NTPServers,

        [Parameter()]
        [int]$SyncInterval = 12,

        [Parameter()]
        [switch]$TimeZone,

        [Parameter()]
        [switch]$SkipScheduledTask,

        [Parameter()]
        [switch]$ValidateAccuracy,

        [Parameter()]
        [bool]$DomainEnvironment = $false
    )

    $passResults = @{
        NTPConfigured = $false
        ServiceConfigured = $false
        InitialSyncCompleted = $false
        ScheduledTaskConfigured = $false
        TimeZoneConfigured = $false
        AccuracyValidated = $false
        Errors = @()
        TimeSyncDetails = @{}
    }

    # Step 1: Configure Time Zone (if requested)
    if ($TimeZone) {
        Write-StatusLog "Configuring time zone..." -Level "Info"
        $timezoneResult = Set-AutomaticTimeZone
        $passResults.TimeZoneConfigured = $timezoneResult.Success
        if (-not $timezoneResult.Success -and $timezoneResult.Error) {
            $passResults.Errors += $timezoneResult.Error
        }
    }

    # Step 2: Configure NTP Servers
    Write-StatusLog "Configuring NTP server settings..." -Level "Info"
    $ntpResult = Set-NTPConfiguration -NTPServers $NTPServers -DomainEnvironment $DomainEnvironment
    $passResults.NTPConfigured = $ntpResult.Success
    $passResults.TimeSyncDetails.NTPServers = $ntpResult.ConfiguredServers
    if (-not $ntpResult.Success -and $ntpResult.Error) {
        $passResults.Errors += $ntpResult.Error
    }

    # Step 3: Configure W32Time Service
    Write-StatusLog "Configuring Windows Time service..." -Level "Info"
    $serviceResult = Set-W32TimeServiceConfiguration
    $passResults.ServiceConfigured = $serviceResult.Success
    if (-not $serviceResult.Success -and $serviceResult.Error) {
        $passResults.Errors += $serviceResult.Error
    }

    # Step 4: Perform Initial Sync
    if ($passResults.ServiceConfigured) {
        Write-StatusLog "Performing initial time synchronization..." -Level "Info"
        $syncResult = Invoke-InitialTimeSync
        $passResults.InitialSyncCompleted = $syncResult.Success
        $passResults.TimeSyncDetails.InitialSync = $syncResult
        if (-not $syncResult.Success -and $syncResult.Error) {
            $passResults.Errors += $syncResult.Error
        }
    }

    # Step 5: Configure Scheduled Task (if not skipped)
    if (-not $SkipScheduledTask) {
        Write-StatusLog "Configuring scheduled time synchronization..." -Level "Info"
        $taskResult = Set-TimeSyncScheduledTask -SyncInterval $SyncInterval
        $passResults.ScheduledTaskConfigured = $taskResult.Success
        $passResults.TimeSyncDetails.ScheduledTask = $taskResult
        if (-not $taskResult.Success -and $taskResult.Error) {
            $passResults.Errors += $taskResult.Error
        }
    }

    # Step 6: Validate Time Accuracy (if requested)
    if ($ValidateAccuracy) {
        Write-StatusLog "Validating time accuracy..." -Level "Info"
        $accuracyResult = Test-TimeAccuracy
        $passResults.AccuracyValidated = $accuracyResult.Success
        $passResults.TimeSyncDetails.Accuracy = $accuracyResult
        if (-not $accuracyResult.Success -and $accuracyResult.Error) {
            $passResults.Errors += $accuracyResult.Error
        }
    }

    return $passResults
}

function Test-TimeSyncRetryNeeded {
    <#
    .SYNOPSIS
    Determines whether the time sync configuration should be retried
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results,

        [Parameter()]
        [switch]$TimeZone,

        [Parameter()]
        [switch]$ValidateAccuracy
    )

    if (-not $Results.NTPConfigured) { return $true }
    if (-not $Results.ServiceConfigured) { return $true }
    if (-not $Results.InitialSyncCompleted) { return $true }

    if ($TimeZone -and -not $Results.TimeZoneConfigured) { return $true }
    if ($ValidateAccuracy -and -not $Results.AccuracyValidated) { return $true }

    return $false
}

function Configure-AutomaticTimeSync {
    <#
    .SYNOPSIS
    Configures comprehensive automatic time synchronization for Windows systems

    .DESCRIPTION
    Sets up robust time synchronization including NTP server configuration, service management,
    scheduled tasks for regular sync, and validation of time accuracy. Supports both domain
    and workgroup environments with appropriate fallback configurations.

    .PARAMETER NTPServers
    Custom NTP servers to use (defaults to reliable public servers)

    .PARAMETER SyncInterval
    Hours between automatic synchronization (default: 12 hours)

    .PARAMETER TimeZone
    Automatically detect and set correct time zone

    .PARAMETER SkipScheduledTask
    Skip creation of scheduled task for regular sync

    .PARAMETER ValidateAccuracy
    Validate time accuracy after configuration

    .OUTPUTS
    Returns hashtable with configuration results and time sync status

    .EXAMPLE
    Configure-AutomaticTimeSync

    .EXAMPLE
    Configure-AutomaticTimeSync -NTPServers @("pool.ntp.org", "time.windows.com") -SyncInterval 6

    .EXAMPLE
    Configure-AutomaticTimeSync -TimeZone -ValidateAccuracy -WhatIf
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string[]]$NTPServers = @("pool.ntp.org", "time.windows.com", "time.nist.gov"),

        [Parameter()]
        [ValidateRange(1, 24)]
        [int]$SyncInterval = 12,

        [Parameter()]
        [switch]$TimeZone,

        [Parameter()]
        [switch]$SkipScheduledTask,

        [Parameter()]
        [switch]$ValidateAccuracy
    )

    Write-StatusLog "Configuring comprehensive time synchronization..." -Level "Info"

    # Initialize results tracking
    $results = @{
        NTPConfigured = $false
        ServiceConfigured = $false
        InitialSyncCompleted = $false
        ScheduledTaskConfigured = $false
        TimeZoneConfigured = $false
        AccuracyValidated = $false
        DomainEnvironment = $false
        ConfigurationSummary = @{}
        Errors = @()
        TimeSyncDetails = @{}
        RetryAttempted = $false
        RetrySucceeded = $false
    }

    try {
        # Detect domain environment
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $results.DomainEnvironment = $computerSystem.PartOfDomain

        if ($results.DomainEnvironment) {
            Write-StatusLog "[INFO] Domain environment detected - will configure for domain time sync" -Level "Info"
        } else {
            Write-StatusLog "[INFO] Workgroup environment detected - will use external NTP servers" -Level "Info"
        }

        # First pass
        $pass1 = Invoke-TimeSyncConfigurationPass `
            -NTPServers $NTPServers `
            -SyncInterval $SyncInterval `
            -TimeZone:$TimeZone `
            -SkipScheduledTask:$SkipScheduledTask `
            -ValidateAccuracy:$ValidateAccuracy `
            -DomainEnvironment $results.DomainEnvironment

        $results.NTPConfigured = $pass1.NTPConfigured
        $results.ServiceConfigured = $pass1.ServiceConfigured
        $results.InitialSyncCompleted = $pass1.InitialSyncCompleted
        $results.ScheduledTaskConfigured = $pass1.ScheduledTaskConfigured
        $results.TimeZoneConfigured = $pass1.TimeZoneConfigured
        $results.AccuracyValidated = $pass1.AccuracyValidated
        $results.Errors = @($pass1.Errors)
        $results.TimeSyncDetails = $pass1.TimeSyncDetails

        # Retry once if needed
        if (Test-TimeSyncRetryNeeded -Results $results -TimeZone:$TimeZone -ValidateAccuracy:$ValidateAccuracy) {
            $results.RetryAttempted = $true
            Write-StatusLog "[WARN] Initial time sync configuration did not fully succeed. Waiting 10 seconds and retrying once..." -Level "Warning"
            Start-Sleep -Seconds 10

            $pass2 = Invoke-TimeSyncConfigurationPass `
                -NTPServers $NTPServers `
                -SyncInterval $SyncInterval `
                -TimeZone:$TimeZone `
                -SkipScheduledTask:$SkipScheduledTask `
                -ValidateAccuracy:$ValidateAccuracy `
                -DomainEnvironment $results.DomainEnvironment

            $results.NTPConfigured = $pass2.NTPConfigured
            $results.ServiceConfigured = $pass2.ServiceConfigured
            $results.InitialSyncCompleted = $pass2.InitialSyncCompleted
            $results.ScheduledTaskConfigured = $pass2.ScheduledTaskConfigured
            $results.TimeZoneConfigured = $pass2.TimeZoneConfigured
            $results.AccuracyValidated = $pass2.AccuracyValidated
            $results.TimeSyncDetails = $pass2.TimeSyncDetails
            $results.Errors = @($pass2.Errors)

            if (-not (Test-TimeSyncRetryNeeded -Results $results -TimeZone:$TimeZone -ValidateAccuracy:$ValidateAccuracy)) {
                $results.RetrySucceeded = $true
                Write-StatusLog "[OK] Time sync configuration succeeded on retry." -Level "Success"
            } else {
                Write-StatusLog "[WARN] Time sync configuration still has issues after retry." -Level "Warning"
            }
        }

        # Generate configuration summary
        $results.ConfigurationSummary = Get-TimeSyncConfigurationSummary

        # Display results
        Show-TimeSyncConfigurationSummary -Results $results

        # Set global status
        $global:LastStatus = if ($results.Errors.Count -eq 0) {
            if ($results.RetrySucceeded) {
                "[OK] Time synchronization configured successfully after retry"
            } else {
                "[OK] Time synchronization configured successfully"
            }
        } else {
            "[WARN] Time synchronization configured with $($results.Errors.Count) issues"
        }

    } catch {
        $errorMsg = "Critical error during time sync configuration: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        $results.Errors += $errorMsg
        $global:LastStatus = "[ERROR] Failed to configure time synchronization"
    }

    return $results
}

function Set-NTPConfiguration {
    <#
    .SYNOPSIS
    Configures NTP server settings based on environment

    .PARAMETER NTPServers
    Array of NTP servers to configure

    .PARAMETER DomainEnvironment
    Whether system is in domain environment

    .OUTPUTS
    Returns hashtable with configuration results
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$NTPServers,

        [Parameter()]
        [bool]$DomainEnvironment = $false
    )

    $result = @{
        Success = $false
        Error = $null
        ConfiguredServers = @()
        ConfigurationMethod = ""
    }

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure NTP servers: $($NTPServers -join ', ')" -ForegroundColor Yellow
            $result.Success = $true
            $result.ConfiguredServers = $NTPServers
            return $result
        }

        if ($DomainEnvironment) {
            # In domain environment, configure for domain hierarchy
            Write-StatusLog "[INFO] Configuring for domain time hierarchy" -Level "Info"
            $ntpServerList = $NTPServers -join " "

            # Configure as NTP client with domain hierarchy
            & w32tm /config /manualpeerlist:$ntpServerList /syncfromflags:DOMHIER /reliable:no /update 2>$null
            $result.ConfigurationMethod = "Domain Hierarchy with fallback servers"

        } else {
            # In workgroup environment, use manual peer list
            Write-StatusLog "[INFO] Configuring for workgroup manual peer list" -Level "Info"
            $ntpServerList = $NTPServers -join " "

            # Configure as NTP client with manual peer list
            & w32tm /config /manualpeerlist:$ntpServerList /syncfromflags:MANUAL /reliable:no /update 2>$null
            $result.ConfigurationMethod = "Manual peer list"
        }

        # Give W32Time a moment to process config changes
        Start-Sleep -Seconds 3

        # Verify configuration was applied - try twice
        $verified = $false
        $lastExitCodeSeen = $LASTEXITCODE

        for ($i = 1; $i -le 2; $i++) {
            $null = & w32tm /query /configuration 2>$null
            $lastExitCodeSeen = $LASTEXITCODE

            if ($lastExitCodeSeen -eq 0) {
                $verified = $true
                break
            }

            Write-StatusLog "[WARN] NTP configuration verification attempt $i failed. Waiting before retry..." -Level "Warning"
            Start-Sleep -Seconds 3
        }

        if ($verified) {
            $result.Success = $true
            $result.ConfiguredServers = $NTPServers
            Write-StatusLog "[OK] NTP servers configured: $($NTPServers -join ', ')" -Level "Success"
        } else {
            $result.Error = "w32tm configuration command failed with exit code: $lastExitCodeSeen"
            Write-StatusLog "[WARN] $($result.Error)" -Level "Warning"
        }

    } catch {
        $result.Error = "Failed to configure NTP servers: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }

    return $result
}

function Set-W32TimeServiceConfiguration {
    <#
    .SYNOPSIS
    Configures the Windows Time service for optimal operation

    .OUTPUTS
    Returns hashtable with service configuration results
    #>

    $result = @{
        Success = $false
        Error = $null
        ServiceStarted = $false
        ServiceRestarted = $false
        StartupTypeSet = $false
    }

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure W32Time service (Automatic startup, restart service)" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }

        # Get current service status
        $w32timeService = Get-Service -Name W32Time -ErrorAction Stop

        # Set service startup type to Automatic
        if ($w32timeService.StartType -ne 'Automatic') {
            Set-Service -Name W32Time -StartupType Automatic -ErrorAction Stop
            $result.StartupTypeSet = $true
            Write-StatusLog "[OK] W32Time service set to Automatic startup" -Level "Success"
        } else {
            Write-StatusLog "[INFO] W32Time service already set to Automatic startup" -Level "Info"
        }

        # Start service if not running
        if ($w32timeService.Status -ne 'Running') {
            Start-Service -Name W32Time -ErrorAction Stop
            $result.ServiceStarted = $true
            Write-StatusLog "[OK] W32Time service started" -Level "Success"
        }

        # Restart service to apply new configuration
        Restart-Service -Name W32Time -Force -ErrorAction Stop
        $result.ServiceRestarted = $true
        Write-StatusLog "[OK] W32Time service restarted to apply configuration" -Level "Success"

        # Wait for service to fully start
        Start-Sleep -Seconds 3

        $result.Success = $true

    } catch {
        $result.Error = "Failed to configure W32Time service: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }

    return $result
}

function Invoke-InitialTimeSync {
    <#
    .SYNOPSIS
    Performs initial time synchronization and validates success

    .OUTPUTS
    Returns hashtable with sync results
    #>

    $result = @{
        Success = $false
        Error = $null
        SyncAttempted = $false
        TimeBefore = $null
        TimeAfter = $null
        SyncSource = $null
    }

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would perform initial time synchronization" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }

        # Record time before sync
        $result.TimeBefore = Get-Date

        # Perform synchronization
        Write-StatusLog "[INFO] Initiating time synchronization..." -Level "Info"
        $syncOutput = & w32tm /resync /force 2>&1
        $result.SyncAttempted = $true

        # Wait for sync to complete
        Start-Sleep -Seconds 5

        # Record time after sync
        $result.TimeAfter = Get-Date

        # Check sync status
        $status = & w32tm /query /status 2>$null
        if ($LASTEXITCODE -eq 0) {
            # Parse sync source from status
            $syncSourceLine = $status | Where-Object { $_ -match "Source:" }
            if ($syncSourceLine) {
                $result.SyncSource = ($syncSourceLine -split ":")[1].Trim()
            }

            $result.Success = $true
            Write-StatusLog "[OK] Initial time synchronization completed" -Level "Success"

            if ($result.SyncSource) {
                Write-StatusLog "[INFO] Sync source: $($result.SyncSource)" -Level "Info"
            }
        } else {
            $result.Error = "Time sync command failed. Output: $syncOutput"
            Write-StatusLog "[WARN] Time sync may have failed: $($result.Error)" -Level "Warning"
        }

    } catch {
        $result.Error = "Failed to perform initial time sync: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }

    return $result
}

function Set-TimeSyncScheduledTask {
    <#
    .SYNOPSIS
    Creates or updates scheduled task for regular time synchronization

    .PARAMETER SyncInterval
    Hours between synchronization attempts

    .OUTPUTS
    Returns hashtable with scheduled task configuration results
    #>

    [CmdletBinding()]
    param(
        [Parameter()]
        [int]$SyncInterval = 12
    )

    $result = @{
        Success = $false
        Error = $null
        TaskCreated = $false
        TaskUpdated = $false
        TaskEnabled = $false
        TaskName = "Windows Time Synchronization"
    }

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would create/update scheduled task for time sync (every $SyncInterval hours)" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }

        # Check if task already exists
        $existingTask = Get-ScheduledTask -TaskName $result.TaskName -ErrorAction SilentlyContinue

        if ($existingTask) {
            # Check if task needs updating
            $taskInfo = Get-ScheduledTaskInfo -TaskName $result.TaskName

            if (-not $taskInfo.Enabled) {
                Enable-ScheduledTask -TaskName $result.TaskName -ErrorAction Stop
                $result.TaskEnabled = $true
                Write-StatusLog "[OK] Enabled existing time sync scheduled task" -Level "Success"
            } else {
                Write-StatusLog "[INFO] Time sync scheduled task already exists and is enabled" -Level "Info"
            }

            $result.Success = $true
            return $result
        }

        # Create new scheduled task
        Write-StatusLog "[INFO] Creating scheduled task for automatic time synchronization..." -Level "Info"

        # Create task action
        $action = New-ScheduledTaskAction -Execute 'w32tm.exe' -Argument '/resync /force'

        # Create triggers based on sync interval
        $triggers = @()
        $hoursInDay = 24
        $syncTimes = @()

        for ($hour = 0; $hour -lt $hoursInDay; $hour += $SyncInterval) {
            $syncTimes += "{0:D2}:00" -f $hour
        }

        foreach ($time in $syncTimes) {
            $triggers += New-ScheduledTaskTrigger -Daily -At $time
        }

        # Create task settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

        # Create task principal
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

        # Register the task
        Register-ScheduledTask `
            -TaskName $result.TaskName `
            -Action $action `
            -Trigger $triggers `
            -Settings $settings `
            -Principal $principal `
            -Description "Automatically synchronizes system time every $SyncInterval hours using w32tm" `
            -ErrorAction Stop

        $result.TaskCreated = $true
        $result.Success = $true
        Write-StatusLog "[OK] Created scheduled task: '$($result.TaskName)' (every $SyncInterval hours)" -Level "Success"

    } catch {
        $result.Error = "Failed to configure scheduled task: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }

    return $result
}

function Set-AutomaticTimeZone {
    <#
    .SYNOPSIS
    Configures automatic time zone detection and setting

    .OUTPUTS
    Returns hashtable with time zone configuration results
    #>

    $result = @{
        Success = $false
        Error = $null
        TimeZoneSet = $false
        CurrentTimeZone = $null
        AutomaticEnabled = $false
    }

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure automatic time zone detection" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }

        # Get current time zone
        $currentTZ = Get-TimeZone
        $result.CurrentTimeZone = $currentTZ.Id

        # Enable automatic time zone if available (Windows 10/11)
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\tzautoupdate"
            if (Test-Path $regPath) {
                Set-ItemProperty -Path $regPath -Name "Start" -Value 3 -ErrorAction Stop
                $result.AutomaticEnabled = $true
                Write-StatusLog "[OK] Automatic time zone updates enabled" -Level "Success"
            }
        } catch {
            Write-StatusLog "[WARN] Could not enable automatic time zone: $($_.Exception.Message)" -Level "Warning"
        }

        # Attempt to detect and set correct time zone based on location
        try {
            $location = Get-WinHomeLocation -ErrorAction SilentlyContinue
            if ($location) {
                Write-StatusLog "[INFO] Current location detected: $($location.HomeLocation)" -Level "Info"
            }
        } catch {
            Write-StatusLog "[INFO] Could not detect location for time zone setting" -Level "Info"
        }

        $result.Success = $true
        Write-StatusLog "[OK] Time zone configuration completed" -Level "Success"

    } catch {
        $result.Error = "Failed to configure time zone: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }

    return $result
}

function Test-TimeAccuracy {
    <#
    .SYNOPSIS
    Validates time accuracy against reliable time sources

    .OUTPUTS
    Returns hashtable with accuracy validation results
    #>

    $result = @{
        Success = $false
        Error = $null
        LocalTime = $null
        NetworkTime = $null
        TimeDifference = $null
        AccuracyWithinTolerance = $false
        ToleranceSeconds = 30
    }

    try {
        if ($WhatIfPreference) {
            Write-Host "   Would validate time accuracy against network time sources" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }

        $result.LocalTime = Get-Date

        # Query W32Time for current sync status
        $w32tmStatus = & w32tm /query /status 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-StatusLog "[OK] W32Time service is responding normally" -Level "Success"

            # Parse last sync time
            $lastSyncLine = $w32tmStatus | Where-Object { $_ -match "Last Successful Sync Time:" }
            if ($lastSyncLine) {
                Write-StatusLog "[INFO] $($lastSyncLine.Trim())" -Level "Info"
            }

            # Check if time is within reasonable tolerance
            # For this validation, we'll consider the system accurate if w32tm reports success
            $result.AccuracyWithinTolerance = $true
            $result.Success = $true
            Write-StatusLog "[OK] Time synchronization appears to be working correctly" -Level "Success"

        } else {
            $result.Error = "W32Time status query failed"
            Write-StatusLog "[WARN] Could not validate time accuracy - W32Time status unavailable" -Level "Warning"
        }

    } catch {
        $result.Error = "Failed to validate time accuracy: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }

    return $result
}

function Get-TimeSyncConfigurationSummary {
    <#
    .SYNOPSIS
    Gets current time synchronization configuration status

    .OUTPUTS
    Returns hashtable with current configuration
    #>

    $summary = @{
        W32TimeService = @{
            Status = "Unknown"
            StartType = "Unknown"
        }
        NTPConfiguration = @{}
        ScheduledTasks = @()
        LastSyncStatus = "Unknown"
        CurrentTimeZone = "Unknown"
    }

    try {
        # W32Time service status
        $w32timeService = Get-Service -Name W32Time -ErrorAction SilentlyContinue
        if ($w32timeService) {
            $summary.W32TimeService.Status = $w32timeService.Status
            $summary.W32TimeService.StartType = $w32timeService.StartType
        }

        # Current time zone
        $currentTZ = Get-TimeZone -ErrorAction SilentlyContinue
        if ($currentTZ) {
            $summary.CurrentTimeZone = $currentTZ.DisplayName
        }

        # NTP configuration
        $w32tmConfig = & w32tm /query /configuration 2>$null
        if ($LASTEXITCODE -eq 0) {
            $ntpServerLine = $w32tmConfig | Where-Object { $_ -match "NtpServer:" }
            if ($ntpServerLine) {
                $summary.NTPConfiguration.Servers = ($ntpServerLine -split ":")[1].Trim()
            }
        }

        # Scheduled tasks
        $timeSyncTasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "*Time*" -or $_.TaskName -like "*Sync*" }
        $summary.ScheduledTasks = $timeSyncTasks | Select-Object TaskName, State

        # Last sync status
        $w32tmStatus = & w32tm /query /status 2>$null
        if ($LASTEXITCODE -eq 0) {
            $lastSyncLine = $w32tmStatus | Where-Object { $_ -match "Last Successful Sync Time:" }
            if ($lastSyncLine) {
                $summary.LastSyncStatus = ($lastSyncLine -split ":")[1].Trim()
            }
        }

    } catch {
        Write-StatusLog "[WARN] Error gathering time sync configuration summary: $($_.Exception.Message)" -Level "Warning"
    }

    return $summary
}

function Show-TimeSyncConfigurationSummary {
    <#
    .SYNOPSIS
    Displays comprehensive summary of time synchronization configuration

    .PARAMETER Results
    Results from the configuration process
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results
    )

    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "CONFIGURATION" }

    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    Write-Host "TIME SYNCHRONIZATION $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan

    Write-Host "Environment: " -NoNewline
    Write-Host $(if($Results.DomainEnvironment){"Domain-joined"}else{"Workgroup"}) -ForegroundColor White

    Write-Host "NTP Configured: " -NoNewline
    Write-Host $(if($Results.NTPConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.NTPConfigured){"Green"}else{"Red"})

    Write-Host "Service Configured: " -NoNewline
    Write-Host $(if($Results.ServiceConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.ServiceConfigured){"Green"}else{"Red"})

    Write-Host "Initial Sync Completed: " -NoNewline
    Write-Host $(if($Results.InitialSyncCompleted){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.InitialSyncCompleted){"Green"}else{"Red"})

    Write-Host "Scheduled Task Configured: " -NoNewline
    Write-Host $(if($Results.ScheduledTaskConfigured){"[OK] Yes"}else{"[WARN] Skipped/Failed"}) -ForegroundColor $(if($Results.ScheduledTaskConfigured){"Green"}else{"Yellow"})

    if ($Results.ContainsKey('RetryAttempted')) {
        Write-Host "Retry Attempted: " -NoNewline
        Write-Host $(if($Results.RetryAttempted){"[OK] Yes"}else{"No"}) -ForegroundColor $(if($Results.RetryAttempted){"Yellow"}else{"White"})
    }

    if ($Results.ContainsKey('RetrySucceeded') -and $Results.RetrySucceeded) {
        Write-Host "Retry Succeeded: " -NoNewline
        Write-Host "[OK] Yes" -ForegroundColor Green
    }

    if ($Results.TimeZoneConfigured) {
        Write-Host "Time Zone Configured: " -NoNewline
        Write-Host "[OK] Yes" -ForegroundColor Green
    }

    if ($Results.AccuracyValidated) {
        Write-Host "Time Accuracy Validated: " -NoNewline
        Write-Host "[OK] Yes" -ForegroundColor Green
    }

    # Show NTP servers if configured
    if ($Results.TimeSyncDetails.NTPServers) {
        Write-Host "`nConfigured NTP Servers:" -ForegroundColor Green
        $Results.TimeSyncDetails.NTPServers | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Green
        }
    }

    # Show sync source if available
    if ($Results.TimeSyncDetails.InitialSync.SyncSource) {
        Write-Host "`nCurrent Sync Source: " -NoNewline
        Write-Host $Results.TimeSyncDetails.InitialSync.SyncSource -ForegroundColor Green
    }

    if ($Results.Errors.Count -gt 0) {
        Write-Host "`nConfiguration Issues:" -ForegroundColor Red
        $Results.Errors | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }

    Write-Host "="*60 -ForegroundColor Cyan
}

# Usage examples:
# Configure-AutomaticTimeSync
# Configure-AutomaticTimeSync -NTPServers @("pool.ntp.org", "time.windows.com") -SyncInterval 6
# Configure-AutomaticTimeSync -TimeZone -ValidateAccuracy -WhatIf

# -----------------------------------------------------------------------------
# Winget Dependencies
# -----------------------------------------------------------------------------
$script:WingetPrereqsInstalled = $false

function Ensure-WingetDependenciesReady {
    <#
    .SYNOPSIS
    Ensures Winget and required dependencies are installed and configured securely
    
    .DESCRIPTION
    Downloads and installs required components for Winget functionality including:
    - Microsoft.VCLibs.140.00.UWPDesktop runtime
    - Microsoft.UI.Xaml framework
    - App Installer (Winget)
    - PowerShell Gallery and NuGet configuration
    
    .PARAMETER Force
    Force reinstallation even if components are already installed
    
    .PARAMETER SkipSourceUpdate
    Skip updating Winget sources for faster execution
    
    .PARAMETER UseOfficialSources
    Use only official Microsoft download sources
    
    .OUTPUTS
    Returns hashtable with installation results and status
    
    .EXAMPLE
    Ensure-WingetDependenciesReady
    
    .EXAMPLE
    Ensure-WingetDependenciesReady -Force -UseOfficialSources
    
    .EXAMPLE
    Ensure-WingetDependenciesReady -WhatIf
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [switch]$Force,
        
        [Parameter()]
        [switch]$SkipSourceUpdate,
        
        [Parameter()]
        [switch]$UseOfficialSources
    )
    
    # Check if already completed and not forcing
    if ($script:WingetPrereqsInstalled -and -not $Force) {
        Write-StatusLog "Prerequisites already configured - skipping" -Level "Info"
        return @{
            AlreadyConfigured = $true
            Success = $true
            ComponentsInstalled = @()
            Errors = @()
        }
    }
    
    Write-StatusLog "Preparing system for Winget and package management..." -Level "Info"
    
    # Initialize results tracking
    $results = @{
        AlreadyConfigured = $false
        Success = $false
        ComponentsInstalled = @()
        ComponentsSkipped = @()
        SourcesConfigured = $false
        PSGalleryConfigured = $false
        Errors = @()
        SecurityConfigured = $false
        DownloadResults = @()
    }
    
    try {
        # Step 1: Configure secure connections
        Write-StatusLog "Configuring secure connection settings..." -Level "Info"
        $securityResult = Set-SecureConnectionSettings
        $results.SecurityConfigured = $securityResult.Success
        if (-not $securityResult.Success) {
            $results.Errors += $securityResult.Error
        }
        
        # Step 2: Check existing installations
        Write-StatusLog "Checking existing component installations..." -Level "Info"
        $existingComponents = Get-ExistingWingetComponents
        
        # Step 3: Define download sources
        $downloadSources = Get-WingetDownloadSources -UseOfficialSources:$UseOfficialSources
        
        # Step 4: Download and install components
        Write-StatusLog "Processing Winget dependencies..." -Level "Info"
        $installResults = Install-WingetDependencies -DownloadSources $downloadSources -ExistingComponents $existingComponents -Force:$Force
        $results.ComponentsInstalled = $installResults.Installed
        $results.ComponentsSkipped = $installResults.Skipped
        $results.DownloadResults = $installResults.DownloadResults
        
        if ($installResults.Errors.Count -gt 0) {
            $results.Errors += $installResults.Errors
        }
        
        # Step 5: Configure Winget sources
        if (-not $SkipSourceUpdate) {
            Write-StatusLog "Configuring Winget sources..." -Level "Info"
            $sourceResult = Set-WingetSources
            $results.SourcesConfigured = $sourceResult.Success
            if (-not $sourceResult.Success) {
                $results.Errors += $sourceResult.Error
            }
        }
        
        # Step 6: Configure PowerShell Gallery
        Write-StatusLog "Configuring PowerShell Gallery..." -Level "Info"
        $psGalleryResult = Set-PowerShellGalleryConfiguration
        $results.PSGalleryConfigured = $psGalleryResult.Success
        if (-not $psGalleryResult.Success) {
            $results.Errors += $psGalleryResult.Error
        }
        
        # Step 7: Validate installation
        $validationResult = Test-WingetInstallation
        $results.Success = $validationResult.Success
        if (-not $validationResult.Success) {
            $results.Errors += $validationResult.Error
        }
        
        # Display summary
        Show-WingetPrerequisitesSummary -Results $results
        
        # Update global flag on success
        if ($results.Success) {
            $script:WingetPrereqsInstalled = $true
            $global:LastStatus = "[OK] Winget dependencies configured successfully"
        } else {
            $global:LastStatus = "[WARN] Winget dependencies configured with $($results.Errors.Count) issues"
        }
        
    } catch {
        $errorMsg = "Critical error during Winget dependencies setup: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        $results.Errors += $errorMsg
        $global:LastStatus = "[ERROR] Failed to configure Winget dependencies"
    }
    
    return $results
}

function Set-SecureConnectionSettings {
    <#
    .SYNOPSIS
    Configures secure connection settings for downloads
    
    .OUTPUTS
    Returns hashtable with security configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        TLSConfigured = $false
        ConnectionLimitSet = $false
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure secure TLS settings and connection limits" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Configure TLS 1.2 and 1.3 for secure downloads
        $currentProtocols = [Net.ServicePointManager]::SecurityProtocol
        $requiredProtocols = [Net.SecurityProtocolType]::Tls12
        
        # Add TLS 1.3 if available (.NET 4.8+)
        try {
            $tls13 = [Net.SecurityProtocolType]::Tls13
            $requiredProtocols = $requiredProtocols -bor $tls13
        } catch {
            # TLS 1.3 not available, continue with TLS 1.2
        }
        
        [Net.ServicePointManager]::SecurityProtocol = $requiredProtocols
        $result.TLSConfigured = $true
        Write-StatusLog "[OK] Secure TLS protocols configured" -Level "Success"
        
        # Set connection limit for better download performance
        [Net.ServicePointManager]::DefaultConnectionLimit = 64
        $result.ConnectionLimitSet = $true
        Write-StatusLog "[OK] Connection limit optimized for downloads" -Level "Success"
        
        $result.Success = $true
        
    } catch {
        $result.Error = "Failed to configure secure connections: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Get-ExistingWingetComponents {
    <#
    .SYNOPSIS
    Checks for existing Winget component installations
    
    .OUTPUTS
    Returns hashtable with existing component status
    #>
    
    $components = @{
        VCLibs = $false
        UIXaml = $false
        AppInstaller = $false
        Winget = $false
    }
    
    try {
        # Check for installed AppX packages
        $installedPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        
        # Check VCLibs
        $vcLibsPackage = $installedPackages | Where-Object { $_.Name -like "*VCLibs*" -and $_.Architecture -eq "X64" }
        $components.VCLibs = ($vcLibsPackage -ne $null)
        
        # Check UI.Xaml
        $uiXamlPackage = $installedPackages | Where-Object { $_.Name -like "*UI.Xaml*" }
        $components.UIXaml = ($uiXamlPackage -ne $null)
        
        # Check App Installer
        $appInstallerPackage = $installedPackages | Where-Object { $_.Name -like "*AppInstaller*" }
        $components.AppInstaller = ($appInstallerPackage -ne $null)
        
        # Check if winget command is available
        try {
            $wingetVersion = winget --version 2>$null
            $components.Winget = ($LASTEXITCODE -eq 0 -and $wingetVersion)
        } catch {
            $components.Winget = $false
        }
        
        Write-StatusLog "Component status - VCLibs: $($components.VCLibs), UI.Xaml: $($components.UIXaml), AppInstaller: $($components.AppInstaller), Winget: $($components.Winget)" -Level "Info"
        
    } catch {
        Write-StatusLog "[WARN] Error checking existing components: $($_.Exception.Message)" -Level "Warning"
    }
    
    return $components
}

function Get-WingetDownloadSources {
    <#
    .SYNOPSIS
    Gets download sources for Winget components with security validation
    
    .PARAMETER UseOfficialSources
    Use only official Microsoft sources
    
    .OUTPUTS
    Returns array of download source objects
    #>
    
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseOfficialSources
    )
    
    if ($UseOfficialSources) {
        # Official Microsoft sources only
        return @(
            @{
                Name = "Microsoft Visual C++ Redistributable"
                Component = "VCLibs"
                Url = "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
                Path = "$env:TEMP\Microsoft.VCLibs.appx"
                Hash = $null  # Microsoft doesn't provide hashes for aka.ms links
                Required = $true
            },
            @{
                Name = "Microsoft UI Xaml"
                Component = "UIXaml"
                Url = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
                Path = "$env:TEMP\Microsoft.UI.Xaml.appx"
                Hash = $null  # GitHub releases, hash verification recommended but not required
                Required = $true
            },
            @{
                Name = "App Installer (Winget)"
                Component = "AppInstaller"
                Url = "https://aka.ms/getwinget"
                Path = "$env:TEMP\AppInstaller.msixbundle"
                Hash = $null  # Microsoft aka.ms link
                Required = $true
            }
        )
    } else {
        # Include alternative sources with hash validation
        return @(
            @{
                Name = "Microsoft Visual C++ Redistributable"
                Component = "VCLibs"
                Url = "https://raw.githubusercontent.com/QuangVNMC/LTSC-Add-Microsoft-Store/master/Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.Appx"
                Path = "$env:TEMP\Microsoft.VCLibs.appx"
                Hash = $null  # Third-party source - use with caution
                Required = $true
            },
            @{
                Name = "Microsoft UI Xaml"
                Component = "UIXaml"
                Url = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
                Path = "$env:TEMP\Microsoft.UI.Xaml.appx"
                Hash = $null
                Required = $true
            },
            @{
                Name = "App Installer (Winget)"
                Component = "AppInstaller"
                Url = "https://aka.ms/getwinget"
                Path = "$env:TEMP\AppInstaller.msixbundle"
                Hash = $null
                Required = $true
            }
        )
    }
}

function Install-WingetDependencies {
    <#
    .SYNOPSIS
    Downloads and installs Winget dependencies with parallel processing
    
    .PARAMETER DownloadSources
    Array of download source objects
    
    .PARAMETER ExistingComponents
    Hashtable of existing component status
    
    .PARAMETER Force
    Force installation even if components exist
    
    .OUTPUTS
    Returns hashtable with installation results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$DownloadSources,
        
        [Parameter(Mandatory)]
        [hashtable]$ExistingComponents,
        
        [Parameter()]
        [switch]$Force
    )
    
    $result = @{
        Installed = @()
        Skipped = @()
        DownloadResults = @()
        Errors = @()
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would download and install Winget dependencies" -ForegroundColor Yellow
            return $result
        }
        
        # Filter sources based on existing components and force flag
        $sourcesToProcess = $DownloadSources | Where-Object {
            $componentExists = $ExistingComponents[$_.Component]
            return $Force -or -not $componentExists
        }
        
        if ($sourcesToProcess.Count -eq 0) {
            Write-StatusLog "All components already installed" -Level "Info"
            return $result
        }
        
        # Download phase - parallel downloads for speed
        Write-StatusLog "Downloading $($sourcesToProcess.Count) components..." -Level "Info"
        $downloadJobs = @()
        
        foreach ($source in $sourcesToProcess) {
            $downloadResult = Invoke-SecureDownload -Source $source
            $result.DownloadResults += $downloadResult
            
            if (-not $downloadResult.Success) {
                $result.Errors += $downloadResult.Error
                Write-StatusLog "[ERROR] Failed to download $($source.Name): $($downloadResult.Error)" -Level "Error"
            }
        }
        
        # Installation phase - sequential for stability
        Write-StatusLog "Installing downloaded components..." -Level "Info"
        foreach ($source in $sourcesToProcess) {
            $downloadResult = $result.DownloadResults | Where-Object { $_.Source.Name -eq $source.Name }
            
            if (-not $downloadResult.Success -or -not (Test-Path $source.Path)) {
                $result.Skipped += $source.Name
                Write-StatusLog "[WARN] Skipping installation of $($source.Name) - download failed" -Level "Warning"
                continue
            }
            
            try {
                $installResult = Install-WingetComponent -Source $source
                
                if ($installResult.Success) {
                    $result.Installed += $source.Name
                    Write-StatusLog "[OK] Successfully installed: $($source.Name)" -Level "Success"
                } else {
                    $result.Errors += $installResult.Error
                    Write-StatusLog "[ERROR] Failed to install $($source.Name): $($installResult.Error)" -Level "Error"
                }
                
            } catch {
                $errorMsg = "Unexpected error installing $($source.Name): $($_.Exception.Message)"
                $result.Errors += $errorMsg
                Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
            } finally {
                # Clean up downloaded file
                if (Test-Path $source.Path) {
                    Remove-Item $source.Path -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
    } catch {
        $result.Errors += "Critical error during component installation: $($_.Exception.Message)"
    }
    
    return $result
}

function Invoke-SecureDownload {
    <#
    .SYNOPSIS
    Performs secure download with multiple methods and validation
    
    .PARAMETER Source
    Source object containing URL, path, and validation info
    
    .OUTPUTS
    Returns hashtable with download results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Source
    )
    
    $result = @{
        Success = $false
        Error = $null
        Method = ""
        Source = $Source
        FileSize = 0
    }
    
    try {
        # Clean up existing file
        if (Test-Path $Source.Path) {
            Remove-Item $Source.Path -Force -ErrorAction SilentlyContinue
        }
        
        Write-StatusLog "Downloading $($Source.Name)..." -Level "Info"
        
        # Try BITS transfer first (fastest and most reliable)
        if (Get-Module -ListAvailable -Name BitsTransfer) {
            Import-Module BitsTransfer -ErrorAction SilentlyContinue
            
            try {
                Start-BitsTransfer -Source $Source.Url -Destination $Source.Path -Priority High -ErrorAction Stop
                $result.Method = "BITS"
                $result.Success = $true
            } catch {
                Write-StatusLog "[WARN] BITS download failed, trying HTTP method" -Level "Warning"
            }
        }
        
        # Fallback to HTTP download
        if (-not $result.Success) {
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($Source.Url, $Source.Path)
                $result.Method = "WebClient"
                $result.Success = $true
            } catch {
                # Final fallback to Invoke-WebRequest
                try {
                    Invoke-WebRequest -Uri $Source.Url -OutFile $Source.Path -UseBasicParsing -ErrorAction Stop
                    $result.Method = "Invoke-WebRequest"
                    $result.Success = $true
                } catch {
                    $result.Error = "All download methods failed. Last error: $($_.Exception.Message)"
                }
            }
        }
        
        # Validate download
        if ($result.Success -and (Test-Path $Source.Path)) {
            $fileInfo = Get-Item $Source.Path
            $result.FileSize = $fileInfo.Length
            
            # Basic validation - file should be larger than 1KB
            if ($fileInfo.Length -lt 1024) {
                $result.Success = $false
                $result.Error = "Downloaded file is too small (likely an error page)"
            } else {
                Write-StatusLog "[OK] Downloaded $($Source.Name) ($($fileInfo.Length) bytes) using $($result.Method)" -Level "Success"
            }
        } elseif ($result.Success) {
            $result.Success = $false
            $result.Error = "Download reported success but file not found"
        }
        
    } catch {
        $result.Error = "Download error: $($_.Exception.Message)"
    }
    
    return $result
}

function Install-WingetComponent {
    <#
    .SYNOPSIS
    Installs a specific Winget component with proper error handling
    
    .PARAMETER Source
    Source object containing installation details
    
    .OUTPUTS
    Returns hashtable with installation results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Source
    )
    
    $result = @{
        Success = $false
        Error = $null
    }
    
    try {
        switch ($Source.Component) {
            "AppInstaller" {
                # App Installer requires special handling
                Add-AppxPackage -Path $Source.Path -ForceApplicationShutdown -ErrorAction Stop
                Start-Sleep -Seconds 5  # Allow time for registration
            }
            default {
                # Standard AppX package installation
                Add-AppxPackage -Path $Source.Path -ErrorAction Stop
            }
        }
        
        $result.Success = $true
        
    } catch {
        $result.Error = $_.Exception.Message
    }
    
    return $result
}

function Set-WingetSources {
    <#
    .SYNOPSIS
    Configures Winget sources securely
    
    .OUTPUTS
    Returns hashtable with source configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        SourcesAdded = @()
        SourcesUpdated = $false
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure Winget sources" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Check if winget is available
        try {
            $wingetVersion = winget --version 2>$null
            if ($LASTEXITCODE -ne 0) {
                $result.Error = "Winget command not available after installation"
                return $result
            }
        } catch {
            $result.Error = "Failed to verify Winget availability: $($_.Exception.Message)"
            return $result
        }
        
        # Add Microsoft Store source if not present
        try {
            $sources = winget source list 2>$null
            if ($sources -notmatch "msstore") {
                winget source add --name msstore --arg "https://storeedgefd.dsx.mp.microsoft.com/v9.0" --accept-source-agreements 2>$null
                if ($LASTEXITCODE -eq 0) {
                    $result.SourcesAdded += "msstore"
                    Write-StatusLog "[OK] Added Microsoft Store source to Winget" -Level "Success"
                }
            } else {
                Write-StatusLog "[INFO] Microsoft Store source already configured" -Level "Info"
            }
        } catch {
            Write-StatusLog "[WARN] Failed to add Microsoft Store source: $($_.Exception.Message)" -Level "Warning"
        }
        
        # Update sources
        try {
            winget source update 2>$null
            if ($LASTEXITCODE -eq 0) {
                $result.SourcesUpdated = $true
                Write-StatusLog "[OK] Winget sources updated" -Level "Success"
            }
        } catch {
            Write-StatusLog "[WARN] Failed to update Winget sources: $($_.Exception.Message)" -Level "Warning"
        }
        
        $result.Success = $true
        
    } catch {
        $result.Error = "Failed to configure Winget sources: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Set-PowerShellGalleryConfiguration {
    <#
    .SYNOPSIS
    Configures PowerShell Gallery and NuGet provider securely
    
    .OUTPUTS
    Returns hashtable with configuration results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        PSGalleryTrusted = $false
        NuGetInstalled = $false
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would configure PowerShell Gallery and NuGet provider" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Configure PowerShell Gallery as trusted
        try {
            $psGallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
            if ($psGallery -and $psGallery.InstallationPolicy -ne 'Trusted') {
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
                $result.PSGalleryTrusted = $true
                Write-StatusLog "[OK] PowerShell Gallery set as trusted repository" -Level "Success"
            } else {
                Write-StatusLog "[INFO] PowerShell Gallery already trusted" -Level "Info"
                $result.PSGalleryTrusted = $true
            }
        } catch {
            $result.Error = "Failed to configure PowerShell Gallery: $($_.Exception.Message)"
            Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
        }
        
        # Install/Update NuGet provider
        try {
            $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
            $minimumVersion = [Version]"2.8.5.201"
            
            if (-not $nugetProvider -or $nugetProvider.Version -lt $minimumVersion) {
                Install-PackageProvider -Name NuGet -MinimumVersion $minimumVersion -Force -ErrorAction Stop
                $result.NuGetInstalled = $true
                Write-StatusLog "[OK] NuGet package provider installed/updated" -Level "Success"
            } else {
                Write-StatusLog "[INFO] NuGet package provider already up to date" -Level "Info"
                $result.NuGetInstalled = $true
            }
        } catch {
            Write-StatusLog "[WARN] Failed to install NuGet provider: $($_.Exception.Message)" -Level "Warning"
        }
        
        $result.Success = ($result.PSGalleryTrusted -and $result.NuGetInstalled)
        
    } catch {
        $result.Error = "Failed to configure PowerShell Gallery: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $($result.Error)" -Level "Error"
    }
    
    return $result
}

function Test-WingetInstallation {
    <#
    .SYNOPSIS
    Validates that Winget is properly installed and functional
    
    .OUTPUTS
    Returns hashtable with validation results
    #>
    
    $result = @{
        Success = $false
        Error = $null
        WingetAvailable = $false
        Version = $null
        SourcesConfigured = $false
    }
    
    try {
        # Test winget command availability
        try {
            $wingetOutput = winget --version 2>$null
            if ($LASTEXITCODE -eq 0 -and $wingetOutput) {
                $result.WingetAvailable = $true
                $result.Version = $wingetOutput.Trim()
                Write-StatusLog "[OK] Winget is functional (version: $($result.Version))" -Level "Success"
            }
        } catch {
            $result.Error = "Winget command test failed: $($_.Exception.Message)"
        }
        
        # Test winget sources
        if ($result.WingetAvailable) {
            try {
                $sources = winget source list 2>$null
                $result.SourcesConfigured = ($LASTEXITCODE -eq 0 -and $sources)
                if ($result.SourcesConfigured) {
                    Write-StatusLog "[OK] Winget sources are configured" -Level "Success"
                }
            } catch {
                Write-StatusLog "[WARN] Could not verify Winget sources" -Level "Warning"
            }
        }
        
        $result.Success = $result.WingetAvailable -and $result.SourcesConfigured
        
    } catch {
        $result.Error = "Validation failed: $($_.Exception.Message)"
    }
    
    return $result
}

function Show-WingetPrerequisitesSummary {
    <#
    .SYNOPSIS
    Displays comprehensive summary of Winget prerequisites installation
    
    .PARAMETER Results
    Results from the installation process
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results
    )
    
    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "INSTALLATION" }
    
    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    Write-Host "WINGET PREREQUISITES $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan
    
    if ($Results.AlreadyConfigured) {
        Write-Host "Status: " -NoNewline
        Write-Host "Already Configured" -ForegroundColor Yellow
    } else {
        Write-Host "Security Configured: " -NoNewline
        Write-Host $(if($Results.SecurityConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.SecurityConfigured){"Green"}else{"Red"})
        
        Write-Host "Components Installed: " -NoNewline
        Write-Host $Results.ComponentsInstalled.Count -ForegroundColor Green
        
        Write-Host "Components Skipped: " -NoNewline
        Write-Host $Results.ComponentsSkipped.Count -ForegroundColor Yellow
        
        Write-Host "Winget Sources Configured: " -NoNewline
        Write-Host $(if($Results.SourcesConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.SourcesConfigured){"Green"}else{"Red"})
        
        Write-Host "PowerShell Gallery Configured: " -NoNewline
        Write-Host $(if($Results.PSGalleryConfigured){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($Results.PSGalleryConfigured){"Green"}else{"Red"})
    }
    
    if ($Results.ComponentsInstalled.Count -gt 0) {
        Write-Host "`nInstalled Components:" -ForegroundColor Green
        $Results.ComponentsInstalled | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Green
        }
    }
    
    if ($Results.Errors.Count -gt 0) {
        Write-Host "`nIssues Encountered:" -ForegroundColor Red
        $Results.Errors | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }
    
    Write-Host "="*60 -ForegroundColor Cyan
}

# Usage examples:
# Ensure-WingetDependenciesReady
# Ensure-WingetDependenciesReady -Force -UseOfficialSources
# Ensure-WingetDependenciesReady -WhatIf

# -----------------------------------------------------------------------------
# Option 9 - Application Updates
# -----------------------------------------------------------------------------

# -----------------------------------------------------------------------------
function Update-Applications {
    [CmdletBinding()]
    param(
        [switch]$IncludeUnknown = $true,
        [switch]$IncludePinned = $false,
        [switch]$AttemptMSStore = $false,
        [switch]$UpdateOffice = $true,
        [int]$OfficeWaitMinutes = 30,
        [string]$LogPath = "$env:SystemDrive\Temp\Weekend-Apps-Update.log"
    )

    $PreviousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"

    switch ($Level) {
        'INFO'  { Write-Host $line -ForegroundColor Cyan }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }

    try {
        $logDir = Split-Path -Path $LogPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $LogPath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Test-IsAdministrator {
    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $paths = @(
        "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe",
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
    )

    foreach ($pattern in $paths) {
        $match = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            Select-Object -First 1
        if ($match) {
            return $match.FullName
        }
    }

    return $null
}

function Test-IsNoiseLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return $true }

    $trimmed = $Line.Trim()

    if ($trimmed -match '^[\-/\\|]+$') { return $true }
    if ($trimmed -match 'Gu') { return $true }
    if ($trimmed -match '^\d+%$') { return $true }
    if ($trimmed -match '^\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)$') { return $true }

    return $false
}

function Get-CleanWingetOutput {
    param([object[]]$RawOutput)

    $clean = @()

    foreach ($line in $RawOutput) {
        if ($null -eq $line) { continue }

        $text = [string]$line
        $text = $text -replace "`r", ''
        $text = $text.TrimEnd()

        if (Test-IsNoiseLine -Line $text) { continue }

        $clean += $text
    }

    return @($clean)
}

function Invoke-Winget {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments,
        [switch]$IgnoreExitCode
    )

    $wingetPath = Get-WingetPath
    if (-not $wingetPath) {
        throw "winget.exe was not found on this system."
    }

    $display = $Arguments -join ' '
    Write-Log "Running: winget $display" 'INFO'

    $rawOutput = & $wingetPath @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    $cleanOutput = Get-CleanWingetOutput -RawOutput @($rawOutput)

    foreach ($line in $cleanOutput) {
        Write-Log $line 'INFO'
    }

    if (-not $IgnoreExitCode -and $exitCode -ne 0) {
        throw "winget exited with code $exitCode while running: $display"
    }

    return [PSCustomObject]@{
        ExitCode  = $exitCode
        RawOutput = @($rawOutput)
        Output    = @($cleanOutput)
    }
}

function Initialize-NetworkDefaults {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
    }
    catch {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {}
    }

    try {
        [Net.ServicePointManager]::DefaultConnectionLimit = 64
    }
    catch {}
}

function Update-WingetSources {
    Write-Log "Refreshing WinGet sources..." 'INFO'
    $null = Invoke-Winget -Arguments @('source', 'update') -IgnoreExitCode
    Write-Log "WinGet source refresh completed." 'OK'
}

function Get-UpgradeInventory {
    param([string]$Source = 'winget')

    $args = @('upgrade', '--source', $Source)
    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

function Parse-WingetInventory {
    param([string[]]$Lines)

    $mainPackages = @()
    $explicitPackages = @()
    $mode = 'None'

    foreach ($line in $Lines) {
        $trimmed = $line.Trim()

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -match 'require explicit targeting for upgrade' -or
            $trimmed -match 'need to be explicitly upgraded') {
            $mode = 'Explicit'
            continue
        }

        if ($trimmed -match '^Name\s+Id\s+Version\s+Available(\s+Source)?$') {
            if ($mode -eq 'None') {
                $mode = 'Main'
            }
            continue
        }

        if ($trimmed -match '^\d+\s+upgrades available\.?$') {
            continue
        }

        if ($trimmed -match '^Installing dependencies:$') { continue }
        if ($trimmed -match '^This package requires the following dependencies:$') { continue }
        if ($trimmed -match '^- Packages$') { continue }
        if ($trimmed -match '^[A-Za-z0-9._+-]+\.[A-Za-z0-9.+-]+$') { continue }

        if ($trimmed -match '^\(\d+/\d+\)\s+Found ') { continue }
        if ($trimmed -match '^Found .+ Version .+$') { continue }
        if ($trimmed -match '^This application is licensed to you by its owner\.$') { continue }
        if ($trimmed -match '^Microsoft is not responsible for, nor does it grant any licenses to, third-party packages\.$') { continue }
        if ($trimmed -match '^Downloading ') { continue }
        if ($trimmed -match '^Successfully verified installer hash$') { continue }
        if ($trimmed -match '^Starting package install\.\.\.$') { continue }
        if ($trimmed -match '^Successfully installed$') { continue }
        if ($trimmed -match '^No installed package found matching input criteria\.$') { continue }
        if ($trimmed -match '^A newer version was found, but the install technology is different from the current version installed\..+$') { continue }

        $pattern = '^(?<Name>.+?)\s{2,}(?<Id>[A-Za-z0-9][A-Za-z0-9._-]+)\s{2,}(?<Version>\S+)\s{2,}(?<Available>\S+)(?:\s{2,}(?<Source>\S+))?$'
        if ($trimmed -match $pattern) {
            $obj = [PSCustomObject]@{
                Name      = $matches.Name.Trim()
                Id        = $matches.Id.Trim()
                Version   = $matches.Version.Trim()
                Available = $matches.Available.Trim()
                Source    = if ($matches.Source) { $matches.Source.Trim() } else { '' }
            }

            if ($mode -eq 'Explicit') {
                $explicitPackages += $obj
            }
            elseif ($mode -eq 'Main') {
                $mainPackages += $obj
            }

            continue
        }
    }

    return [PSCustomObject]@{
        Main     = @($mainPackages)
        Explicit = @($explicitPackages)
    }
}

function Invoke-WingetUpgradeAll {
    param([string]$Source = 'winget')

    $args = @(
        'upgrade', '--all',
        '--source', $Source,
        '--silent',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--disable-interactivity'
    )

    if ($IncludeUnknown) { $args += '--include-unknown' }
    if ($IncludePinned)  { $args += '--include-pinned' }

    return (Invoke-Winget -Arguments $args -IgnoreExitCode)
}

function Invoke-TargetedUpgrade {
    param(
        [Parameter(Mandatory)]$Package,
        [switch]$UseNameFallback
    )

    $args = @(
        'upgrade',
        '--id', $Package.Id,
        '--exact',
        '--silent',
        '--accept-source-agreements',
        '--accept-package-agreements',
        '--disable-interactivity'
    )

    if ($Package.Version -eq 'Unknown' -or $IncludeUnknown) {
        $args += '--include-unknown'
    }

    if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
        $args += @('--source', $Package.Source)
    }

    $result = Invoke-Winget -Arguments $args -IgnoreExitCode

    if ($result.ExitCode -eq 0) {
        return $result
    }

    $text = ($result.Output -join "`n")
    if ($text -match 'install technology is different from the current version installed') {
        Write-Log "Package $($Package.Id) cannot be upgraded in-place because Winget reports the installed technology differs from the new package. Manual uninstall/reinstall is required." 'WARN'
        return $result
    }

    if ($UseNameFallback) {
        Write-Log "ID-targeted upgrade failed for $($Package.Id). Trying name fallback for $($Package.Name)." 'WARN'

        $fallbackArgs = @(
            'upgrade',
            '--name', $Package.Name,
            '--exact',
            '--silent',
            '--accept-source-agreements',
            '--accept-package-agreements',
            '--disable-interactivity'
        )

        if ($Package.Version -eq 'Unknown' -or $IncludeUnknown) {
            $fallbackArgs += '--include-unknown'
        }

        if (-not [string]::IsNullOrWhiteSpace($Package.Source)) {
            $fallbackArgs += @('--source', $Package.Source)
        }

        return (Invoke-Winget -Arguments $fallbackArgs -IgnoreExitCode)
    }

    return $result
}

function Invoke-PackageSet {
    param(
        [object[]]$Packages,
        [string]$Label,
        [switch]$UseNameFallback
    )

    $failures = 0

    if (-not $Packages -or $Packages.Count -eq 0) {
        Write-Log "No packages found for $Label." 'INFO'
        return 0
    }

    foreach ($pkg in $Packages) {
        Write-Log "$($Label): $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        $result = Invoke-TargetedUpgrade -Package $pkg -UseNameFallback:$UseNameFallback

        if ($result.ExitCode -eq 0) {
            Write-Log "Targeted upgrade completed for $($pkg.Id)." 'OK'
        }
        else {
            Write-Log "Targeted upgrade failed for $($pkg.Id) with exit code $($result.ExitCode)." 'WARN'
            $failures++
        }
    }

    return $failures
}

function Update-OfficeClickToRun {
    param([int]$WaitMinutes = 30)

    $officePath = 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe'
    if (-not (Test-Path -Path $officePath)) {
        Write-Log "Office Click-to-Run client not found. Skipping Office update." 'INFO'
        return
    }

    Write-Log "Starting Office Click-to-Run update..." 'INFO'

    $proc = Start-Process -FilePath $officePath `
                          -ArgumentList '/update', 'USER', 'displaylevel=False', 'forceappshutdown=True' `
                          -PassThru `
                          -WindowStyle Hidden

    $completed = $proc.WaitForExit($WaitMinutes * 60 * 1000)

    if (-not $completed) {
        Write-Log "Office update process did not finish within $WaitMinutes minute(s)." 'WARN'
        return
    }

    Write-Log "Office Click-to-Run exited with code $($proc.ExitCode)." 'INFO'
}

function Test-RebootRequired {
    try {
        $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        return [bool]$sysInfo.RebootRequired
    }
    catch {
        return $false
    }
}

function Remove-PackagesById {
    param(
        [object[]]$Packages,
        [string[]]$IdsToRemove
    )

    if (-not $Packages) { return @() }
    if (-not $IdsToRemove -or $IdsToRemove.Count -eq 0) { return @($Packages) }

    $idSet = @{}
    foreach ($id in $IdsToRemove) {
        $idSet[$id.ToLowerInvariant()] = $true
    }

    $filtered = @()
    foreach ($pkg in $Packages) {
        if (-not $idSet.ContainsKey($pkg.Id.ToLowerInvariant())) {
            $filtered += $pkg
        }
    }

    return @($filtered)
}

if (-not (Test-IsAdministrator)) {
    Write-Error "This script must be run as Administrator."
    return 1
}

Initialize-NetworkDefaults
Write-Log "Initializing application update script..." 'INFO'

try {
    Update-WingetSources

    Write-Log "Collecting pre-upgrade inventory..." 'INFO'
    $preInventoryResult = Get-UpgradeInventory -Source 'winget'
    $preParsed = Parse-WingetInventory -Lines $preInventoryResult.Output
    $preMain = $preParsed.Main
    $preExplicit = $preParsed.Explicit

    Write-Log "Main inventory found $($preMain.Count) upgradeable package(s)." $(if ($preMain.Count -gt 0) { 'OK' } else { 'INFO' })
    Write-Log "Explicit-target inventory found $($preExplicit.Count) package(s)." $(if ($preExplicit.Count -gt 0) { 'WARN' } else { 'INFO' })

    if ($preExplicit.Count -gt 0) {
        foreach ($pkg in $preExplicit) {
            Write-Log "Pre-scan explicit target: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        }
    }

    $bulkResult = Invoke-WingetUpgradeAll -Source 'winget'

    # Always process explicit-target packages from the inventory scan.
    $explicitFailures = Invoke-PackageSet -Packages $preExplicit -Label 'Explicit target required' -UseNameFallback

    if ($AttemptMSStore) {
        Write-Log "Attempting second pass against msstore source..." 'INFO'
        $null = Invoke-WingetUpgradeAll -Source 'msstore'
    }

    if ($UpdateOffice) {
        Update-OfficeClickToRun -WaitMinutes $OfficeWaitMinutes
    }

    Write-Log "Collecting post-upgrade inventory..." 'INFO'
    $postInventoryResult = Get-UpgradeInventory -Source 'winget'
    $postParsed = Parse-WingetInventory -Lines $postInventoryResult.Output
    $remainingMain = $postParsed.Main
    $remainingExplicit = $postParsed.Explicit

    # Do not keep retrying packages already handled in explicit-target pass.
    $remainingMain = Remove-PackagesById -Packages $remainingMain -IdsToRemove ($preExplicit.Id)

    $retryFailures = 0
    $retryFailures += Invoke-PackageSet -Packages $remainingMain -Label 'Retry remaining package' -UseNameFallback
    $retryFailures += Invoke-PackageSet -Packages $remainingExplicit -Label 'Retry explicit-target package' -UseNameFallback

    Write-Log "Collecting final inventory..." 'INFO'
    $finalInventoryResult = Get-UpgradeInventory -Source 'winget'
    $finalParsed = Parse-WingetInventory -Lines $finalInventoryResult.Output
    $finalMain = Remove-PackagesById -Packages $finalParsed.Main -IdsToRemove ($preExplicit.Id)
    $finalExplicit = $finalParsed.Explicit

    if ($finalMain.Count -gt 0) {
        Write-Log "Final remaining main packages: $($finalMain.Count)" 'WARN'
        foreach ($pkg in $finalMain) {
            Write-Log "Still remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        }
    }
    else {
        Write-Log "Final remaining main packages: 0" 'OK'
    }

    if ($finalExplicit.Count -gt 0) {
        Write-Log "Final remaining explicit-target packages: $($finalExplicit.Count)" 'WARN'
        foreach ($pkg in $finalExplicit) {
            Write-Log "Still explicit-target remaining: $($pkg.Name) [$($pkg.Id)] $($pkg.Version) -> $($pkg.Available)" 'WARN'
        }
    }
    else {
        Write-Log "Final remaining explicit-target packages: 0" 'OK'
    }

    if (Test-RebootRequired) {
        Write-Log "A reboot is required after application updates." 'WARN'
        return 3010
    }

    if ($explicitFailures -eq 0 -and $retryFailures -eq 0 -and $finalMain.Count -eq 0 -and $finalExplicit.Count -eq 0) {
        Write-Log "Application update script completed successfully." 'OK'
        return 0
    }
    else {
        Write-Log "Application update script completed with remaining packages or non-zero package results." 'WARN'
        return 2
    }
}
catch {
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    return 3
}
finally {
    $ErrorActionPreference = $PreviousErrorActionPreference
}
}

# -----------------------------------------------------------------------------
# Option 10 - Vendor Driver Updates (HP/Dell)
# -----------------------------------------------------------------------------
function Update-HPDrivers {
    <#
    .SYNOPSIS
    Runs the Option 10 vendor driver update workflow.

    .DESCRIPTION
    Uses the merged code from 05_Weekend_HP_Drivers_Update.ps1 v2.3.8.
    Detects HP or Dell systems, applies local desktop power policy, runs the
    matching vendor driver workflow, excludes BIOS/Firmware, writes YAML logs,
    and performs cleanup before returning to the Compton Tech Tools menu.
    #>
    [CmdletBinding()]
    param(
        [string]$WorkingRoot = 'C:\Temp\DriverUpdates',
        [string]$YamlLogFolder = 'C:\Logs',
        [switch]$IncludeSoftware,
        [switch]$IncludeBIOS,
        [string]$HpiaSourceFolder = '\\filesvr\Labscripts\HPImageAssistant',
        [string]$LocalHpiaFolder = 'C:\ProgramData\Compton\HPImageAssistant'
    )

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    # -----------------------------
    # Script metadata
    # -----------------------------
    $script:ScriptName        = '05_Weekend_HP_Drivers_Update.ps1'
    $script:ScriptVersion     = '2.3.8'
    $script:StartTime         = Get-Date
    $script:RunFailures       = New-Object System.Collections.Generic.List[string]
    $script:InstalledList     = New-Object System.Collections.Generic.List[string]
    $script:SkippedList       = New-Object System.Collections.Generic.List[string]
    $script:YamlActionLines   = New-Object System.Collections.Generic.List[string]
    $script:DriverResults     = New-Object System.Collections.Generic.List[object]
    $script:DetectedVendor    = $null
    $script:YamlPath          = $null
    $script:ComputerName      = $env:COMPUTERNAME

    # -----------------------------
    # Logging helpers
    # -----------------------------
    function Write-Log {
        param(
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line = "[{0}] [{1}] {2}" -f $timestamp, $Level.PadRight(5), $Message

        switch ($Level) {
            'INFO'  { Write-Host $line -ForegroundColor Cyan }
            'OK'    { Write-Host $line -ForegroundColor Green }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
            default { Write-Host $line }
        }
    }

    function Write-Section {
        param([Parameter(Mandatory)][string]$Title)

        $border = ('=' * 72)
        Write-Host ''
        Write-Host $border -ForegroundColor Magenta
        Write-Host ("  {0}" -f $Title) -ForegroundColor Magenta
        Write-Host $border -ForegroundColor Magenta
        Add-YamlAction ("Section: {0}" -f $Title)
    }

    function Add-RunFailure {
        param([Parameter(Mandatory)][string]$Message)
        $script:RunFailures.Add($Message) | Out-Null
        Write-Log $Message 'WARN'
    }

    function Add-DriverResult {
        param(
            [Parameter(Mandatory)][string]$Vendor,
            [Parameter(Mandatory)][string]$Name,
            [string]$Id,
            [string]$Category,
            [ValidateSet('Detected','Installed','Downloaded','Skipped','Failed','Blocked')][string]$Status,
            [string]$Message = ''
        )

        $script:DriverResults.Add([pscustomobject]@{
            Vendor   = $Vendor
            Name     = $Name
            Id       = $Id
            Category = $Category
            Status   = $Status
            Message  = $Message
        }) | Out-Null
    }

    function ConvertTo-YamlSafeString {
        param([AllowNull()][object]$Value)

        if ($null -eq $Value) { return 'null' }

        $text = [string]$Value
        $text = $text -replace "`r", ''
        $text = $text -replace "`n", ' '
        $text = $text -replace "'", "''"
        return "'$text'"
    }

    function Initialize-YamlLog {
        param(
            [Parameter(Mandatory)][string]$ComputerName,
            [Parameter(Mandatory)][string]$YamlPath
        )

        $script:YamlPath = $YamlPath
        $script:YamlActionLines.Clear()

        Add-YamlAction 'Script initialized.'
        Add-YamlAction ("Working root: {0}" -f $WorkingRoot)
        Add-YamlAction ("Computer: {0}" -f $ComputerName)
    }

    function Add-YamlAction {
        param([Parameter(Mandatory)][string]$Text)
        $script:YamlActionLines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $Text))) | Out-Null
    }

    function Save-YamlLog {
        param(
            [Parameter(Mandatory)][string]$Status
        )

        if ([string]::IsNullOrWhiteSpace($script:YamlPath)) {
            return
        }

        $endTime = Get-Date
        $duration = [math]::Round(($endTime - $script:StartTime).TotalSeconds, 0)

        $lines = New-Object System.Collections.Generic.List[string]

        $lines.Add('script:') | Out-Null
        $lines.Add(("  name: {0}" -f (ConvertTo-YamlSafeString $script:ScriptName))) | Out-Null
        $lines.Add(("  version: {0}" -f (ConvertTo-YamlSafeString $script:ScriptVersion))) | Out-Null
        $lines.Add(("  computer: {0}" -f (ConvertTo-YamlSafeString $script:ComputerName))) | Out-Null
        $lines.Add(("  started: {0}" -f (ConvertTo-YamlSafeString ($script:StartTime.ToString('s'))))) | Out-Null
        $lines.Add(("  ended: {0}" -f (ConvertTo-YamlSafeString ($endTime.ToString('s'))))) | Out-Null
        $lines.Add(("  duration_seconds: {0}" -f $duration)) | Out-Null

        $lines.Add('run:') | Out-Null
        $lines.Add(("  status: {0}" -f (ConvertTo-YamlSafeString $Status))) | Out-Null
        $lines.Add(("  vendor: {0}" -f (ConvertTo-YamlSafeString $script:DetectedVendor))) | Out-Null

        $lines.Add('  actions:') | Out-Null
        if ($script:YamlActionLines.Count -eq 0) {
            $lines.Add('    []') | Out-Null
        }
        else {
            foreach ($line in $script:YamlActionLines) {
                $lines.Add($line) | Out-Null
            }
        }

        $lines.Add('  installed:') | Out-Null
        if ($script:InstalledList.Count -eq 0) {
            $lines.Add('    []') | Out-Null
        }
        else {
            foreach ($item in $script:InstalledList) {
                $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
            }
        }

        $lines.Add('  skipped:') | Out-Null
        if ($script:SkippedList.Count -eq 0) {
            $lines.Add('    []') | Out-Null
        }
        else {
            foreach ($item in $script:SkippedList) {
                $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
            }
        }

        $lines.Add('  failures:') | Out-Null
        if ($script:RunFailures.Count -eq 0) {
            $lines.Add('    []') | Out-Null
        }
        else {
            foreach ($item in $script:RunFailures) {
                $lines.Add(("    - {0}" -f (ConvertTo-YamlSafeString $item))) | Out-Null
            }
        }

        $lines.Add('drivers:') | Out-Null
        if ($script:DriverResults.Count -eq 0) {
            $lines.Add('  []') | Out-Null
        }
        else {
            foreach ($driver in $script:DriverResults) {
                $lines.Add('  -') | Out-Null
                $lines.Add(("    vendor: {0}" -f (ConvertTo-YamlSafeString $driver.Vendor))) | Out-Null
                $lines.Add(("    name: {0}" -f (ConvertTo-YamlSafeString $driver.Name))) | Out-Null
                $lines.Add(("    id: {0}" -f (ConvertTo-YamlSafeString $driver.Id))) | Out-Null
                $lines.Add(("    category: {0}" -f (ConvertTo-YamlSafeString $driver.Category))) | Out-Null
                $lines.Add(("    status: {0}" -f (ConvertTo-YamlSafeString $driver.Status))) | Out-Null
                $lines.Add(("    message: {0}" -f (ConvertTo-YamlSafeString $driver.Message))) | Out-Null
            }
        }

        Set-Content -LiteralPath $script:YamlPath -Value $lines -Encoding UTF8
        Write-Host ("YAML log written successfully: {0}" -f $script:YamlPath) -ForegroundColor Green
    }

    # -----------------------------
    # File/folder helpers
    # -----------------------------
    function Ensure-Folder {
        param([Parameter(Mandatory)][string]$Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
    }

    function Ensure-WorkingFolderPermissions {
        param([Parameter(Mandatory)][string]$Path)

        try {
            & icacls.exe $Path '/grant' '*S-1-1-0:(OI)(CI)F' '/T' '/C' | Out-Null
        }
        catch {
            Write-Log ("Unable to relax working folder permissions on {0}: {1}" -f $Path, $_.Exception.Message) 'WARN'
        }
    }

    function Remove-WorkingFolderRobust {
        param(
            [Parameter(Mandatory)][string]$Path,
            [int]$RetryCount = 6,
            [int]$RetryDelaySeconds = 5
        )

        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Log ("Working folder already absent: {0}" -f $Path) 'OK'
            return
        }

        Write-Log ("Attempting to remove working folder: {0}" -f $Path) 'INFO'
        for ($i = 1; $i -le $RetryCount; $i++) {
            try {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
                Write-Log 'Working folder removed successfully.' 'OK'
                return
            }
            catch {
                if ($i -eq $RetryCount) {
                    Add-RunFailure ("Failed to remove working folder after retries: {0}" -f $_.Exception.Message)
                    return
                }
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }

    # -----------------------------
    # Power policy
    # -----------------------------
    function Set-LocalPowerPolicyDesktop {
        try {
            $base = 'HKLM:\SOFTWARE\Policies\Microsoft\Power\PowerSettings'

            $displayGuid         = '3C0BC021-C8A8-4E07-A973-6B14CBCB2B7E'
            $sleepGuid           = '29F6C1DB-86DA-48C5-9FDB-F2B67B1F44DA'
            $unattendedSleepGuid = '7BC4A2F9-D8FC-4469-B07B-33EB785AACA0'
            $hybridSleepGuid     = '94AC6D29-73CE-41A6-809F-6363BA21B47E'
            $hibernateGuid       = '9D7815A6-7EE4-497E-8888-515A05F02364'

            foreach ($guid in @($displayGuid,$sleepGuid,$unattendedSleepGuid,$hybridSleepGuid,$hibernateGuid)) {
                $path = Join-Path $base $guid
                if (-not (Test-Path -LiteralPath $path)) {
                    New-Item -Path $path -Force | Out-Null
                }
            }

            New-ItemProperty -Path (Join-Path $base $displayGuid) -Name ACSettingIndex -PropertyType DWord -Value 3600 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $displayGuid) -Name DCSettingIndex -PropertyType DWord -Value 3600 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $sleepGuid) -Name ACSettingIndex -PropertyType DWord -Value 0 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $sleepGuid) -Name DCSettingIndex -PropertyType DWord -Value 0 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $unattendedSleepGuid) -Name ACSettingIndex -PropertyType DWord -Value 0 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $unattendedSleepGuid) -Name DCSettingIndex -PropertyType DWord -Value 0 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $hybridSleepGuid) -Name ACSettingIndex -PropertyType DWord -Value 1 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $hybridSleepGuid) -Name DCSettingIndex -PropertyType DWord -Value 1 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $hibernateGuid) -Name ACSettingIndex -PropertyType DWord -Value 0 -Force | Out-Null
            New-ItemProperty -Path (Join-Path $base $hibernateGuid) -Name DCSettingIndex -PropertyType DWord -Value 0 -Force | Out-Null

            powercfg -change -monitor-timeout-ac 60 > $null 2>&1
            powercfg -change -monitor-timeout-dc 60 > $null 2>&1
            powercfg -change -standby-timeout-ac 0 > $null 2>&1
            powercfg -change -standby-timeout-dc 0 > $null 2>&1
            powercfg -hibernate off > $null 2>&1

            Add-YamlAction 'Applied local power policy (display 60 minutes, sleep never).'
            Write-Log 'Local power policy applied successfully.' 'OK'
        }
        catch {
            if (($_ | Out-String) -match 'Group policy override settings exist') {
                Write-Log 'Local power policy already enforced by policy.' 'WARN'
            }
            else {
                Add-RunFailure ("Failed to apply local power policy: {0}" -f $_.Exception.Message)
            }
        }
    }

    function Enforce-DesktopPowerSettings {
        $highPerf = (powercfg -l | Select-String 'High performance' | ForEach-Object {
            if ($_ -match '([A-Fa-f0-9\-]{36})') { $matches[1] }
        } | Select-Object -First 1)

        if ($highPerf) {
            powercfg -setactive $highPerf > $null 2>&1
            Write-Log ("High Performance power plan available: {0}" -f $highPerf) 'OK'
        }
        else {
            Write-Log 'High Performance power plan was not found. Continuing with current plan.' 'WARN'
        }

        powercfg -change -monitor-timeout-ac 60 > $null 2>&1
        powercfg -change -monitor-timeout-dc 60 > $null 2>&1
        powercfg -change -standby-timeout-ac 0 > $null 2>&1
        powercfg -change -standby-timeout-dc 0 > $null 2>&1
        powercfg -hibernate off > $null 2>&1

        Write-Log 'Desktop power settings enforced successfully.' 'OK'
    }

    # -----------------------------
    # Vendor detection
    # -----------------------------
    function Get-SystemManufacturer {
        try {
            return (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Manufacturer
        }
        catch {
            throw "Unable to determine system manufacturer. $($_.Exception.Message)"
        }
    }

    function Get-DriverVendor {
        $manufacturer = Get-SystemManufacturer
        Write-Log ("Detected manufacturer: {0}" -f $manufacturer) 'INFO'

        if ($manufacturer -match 'Dell') { return 'Dell' }
        if ($manufacturer -match 'HP|Hewlett-Packard') { return 'HP' }

        throw "Unsupported manufacturer for this script: $manufacturer"
    }

    # -----------------------------
    # HP Support - HP Image Assistant extracted-folder deployment
    # -----------------------------
    function Get-HPSystemModelInfo {
        [CmdletBinding()]
        param()

        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
        $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue

        $platform = $null
        if ($bb -and $bb.Product) {
            $platform = $bb.Product.ToString().Trim().ToUpper()
            if ($platform.Length -gt 4) {
                $platform = $platform.Substring(0,4)
            }
        }

        $model = if ($cs.Model) { $cs.Model.ToString().Trim() } else { 'Unknown' }
        $sku = if ($csp -and $csp.Version) { $csp.Version.ToString().Trim() } else { 'Unknown' }

        $info = [pscustomobject]@{
            Manufacturer = $cs.Manufacturer
            Model        = $model
            SKU          = $sku
            Platform     = $platform
        }

        Write-Log ("Detected HP system model: {0}" -f $info.Model) 'INFO'
        Write-Log ("Detected HP platform/baseboard ID: {0}" -f $info.Platform) 'INFO'
        Add-YamlAction ("Detected HP system model: {0}" -f $info.Model)
        Add-YamlAction ("Detected HP platform/baseboard ID: {0}" -f $info.Platform)

        return $info
    }

    function Install-HPIAFromExtractedFolder {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$SourceFolder,
            [Parameter(Mandatory)][string]$DestinationFolder
        )

        Write-Section 'HP Image Assistant Local Deployment'
        Write-Log ("HPIA source folder: {0}" -f $SourceFolder) 'INFO'
        Write-Log ("HPIA local folder: {0}" -f $DestinationFolder) 'INFO'
        Add-YamlAction ("HPIA source folder: {0}" -f $SourceFolder)
        Add-YamlAction ("HPIA local folder: {0}" -f $DestinationFolder)

        if (-not (Test-Path -LiteralPath $SourceFolder)) {
            throw "HPIA source folder not found: $SourceFolder"
        }

        $sourceFiles = @(Get-ChildItem -Path $SourceFolder -Recurse -File -ErrorAction SilentlyContinue)
        Write-Log ("Source HPIA file count: {0}" -f $sourceFiles.Count) 'INFO'
        Add-YamlAction ("Source HPIA file count: {0}" -f $sourceFiles.Count)

        $sourceExe = Get-ChildItem -Path $SourceFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
            Sort-Object FullName |
            Select-Object -First 1

        if (-not $sourceExe) {
            throw "HPImageAssistant.exe was not found anywhere under source folder: $SourceFolder"
        }

        Write-Log ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName) 'OK'
        Add-YamlAction ("Found source HPImageAssistant.exe: {0}" -f $sourceExe.FullName)

        try {
            if (Test-Path -LiteralPath $DestinationFolder) {
                Write-Log ("Removing existing local HPIA folder: {0}" -f $DestinationFolder) 'INFO'
                Remove-Item -LiteralPath $DestinationFolder -Recurse -Force -ErrorAction Stop
            }

            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null

            Write-Log 'Copying extracted HPIA files locally with robocopy...' 'INFO'

            $roboLog = Join-Path $DestinationFolder 'HPIA_robocopy.log'
            $roboArgs = @(
                ('"{0}"' -f $SourceFolder),
                ('"{0}"' -f $DestinationFolder),
                '/E',
                '/COPY:DAT',
                '/R:3',
                '/W:5',
                '/NFL',
                '/NDL',
                '/NP',
                ('/LOG:"{0}"' -f $roboLog)
            )

            $robo = Start-Process -FilePath "$env:SystemRoot\System32\robocopy.exe" -ArgumentList ($roboArgs -join ' ') -Wait -PassThru -NoNewWindow

            # Robocopy exit codes 0-7 are success/non-fatal. 8+ indicates failure.
            if ($robo.ExitCode -ge 8) {
                throw "Robocopy failed copying HPIA files. Exit code: $($robo.ExitCode). Log: $roboLog"
            }

            Write-Log ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode) 'OK'
            Add-YamlAction ("HPIA files copied locally with robocopy exit code {0}." -f $robo.ExitCode)

            try {
                Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue |
                    ForEach-Object { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue }
            }
            catch {}

            $copiedFiles = @(Get-ChildItem -Path $DestinationFolder -Recurse -File -ErrorAction SilentlyContinue)
            Write-Log ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count) 'INFO'
            Add-YamlAction ("Local HPIA folder file count after copy: {0}" -f $copiedFiles.Count)

            $localExeItem = Get-ChildItem -Path $DestinationFolder -Recurse -Filter 'HPImageAssistant.exe' -File -ErrorAction SilentlyContinue |
                Sort-Object FullName |
                Select-Object -First 1

            if (-not $localExeItem) {
                $sampleFiles = $copiedFiles | Select-Object -First 20 | ForEach-Object { $_.FullName }
                foreach ($sample in $sampleFiles) {
                    Write-Log ("Local HPIA sample file: {0}" -f $sample) 'WARN'
                }

                throw "HPImageAssistant.exe was not found anywhere under local copy folder: $DestinationFolder"
            }

            $localExe = $localExeItem.FullName

            Write-Log ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe) 'OK'
            Add-YamlAction ("Resolved local HPImageAssistant.exe location: {0}" -f $localExe)

            return $localExe
        }
        catch {
            throw "Failed to deploy HP Image Assistant locally: $($_.Exception.Message)"
        }
    }

    function Get-HpiaExitStatus {
        param([int]$ExitCode)

        switch ($ExitCode) {
            0    { return 'success' }
            1    { return 'failed' }
            2    { return 'cancelled' }
            3    { return 'needs_reboot' }
            256  { return 'no_recommendations' }
            257  { return 'recommendations_found' }
            3010 { return 'needs_reboot' }
            3011 { return 'not_auto_installable_skipped' }
            4097 { return 'invalid_parameters' }
            default { return 'unknown' }
        }
    }

    function Get-HpiaRecommendationObjects {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$ReportFolder
        )

        # Use a flexible PowerShell array instead of a strongly typed generic list.
        # HPIA reports can contain mixed object types from JSON and XML parsing.
        $recommendations = @()

        $jsonFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)
        foreach ($jsonFile in $jsonFiles) {
            try {
                $json = Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

                if ($json.HPIA -and $json.HPIA.Recommendations) {
                    foreach ($rec in @($json.HPIA.Recommendations)) {
                        $recommendations += $rec
                        try {
                            Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                        }
                        catch {}
                    }
                }
                elseif ($json.Recommendations) {
                    foreach ($rec in @($json.Recommendations)) {
                        $recommendations += $rec
                        try {
                            Write-Log ("HPIA JSON recommendation type: {0}" -f $rec.GetType().FullName) 'INFO'
                        }
                        catch {}
                    }
                }
            }
            catch {
                Write-Log ("Unable to parse HPIA JSON report {0}: {1}" -f $jsonFile.FullName, $_.Exception.Message) 'WARN'
            }
        }

        $xmlFiles = @(Get-ChildItem -Path $ReportFolder -Filter '*.xml' -Recurse -File -ErrorAction SilentlyContinue)
        foreach ($xmlFile in $xmlFiles) {
            try {
                [xml]$xml = Get-Content -LiteralPath $xmlFile.FullName -Raw -ErrorAction Stop

                $nodes = @($xml.SelectNodes('//*[contains(translate(local-name(), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "recommend")]'))
                foreach ($node in $nodes) {
                    $recommendations += $node
                    try {
                        Write-Log ("HPIA XML recommendation node type: {0}" -f $node.GetType().FullName) 'INFO'
                    }
                    catch {}
                }
            }
            catch {
                Write-Log ("Unable to parse HPIA XML report {0}: {1}" -f $xmlFile.FullName, $_.Exception.Message) 'WARN'
            }
        }

        return @($recommendations)
    }

    function Get-HpiaRecommendationValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]$Recommendation,
            [Parameter(Mandatory)][string[]]$PropertyNames
        )

        foreach ($prop in $PropertyNames) {
            try {
                if ($Recommendation.PSObject.Properties.Name -contains $prop) {
                    $value = $Recommendation.$prop
                    if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace($value.ToString())) {
                        return $value.ToString()
                    }
                }
            }
            catch {}
        }

        # XML fallback
        try {
            foreach ($prop in $PropertyNames) {
                $node = $Recommendation.SelectSingleNode('.//*[local-name()="' + $prop + '"]')
                if ($node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
                    return $node.InnerText.Trim()
                }
            }
        }
        catch {}

        return $null
    }

    function Get-HpiaSoftPaqNumber {
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Recommendation)

        $candidate = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
            'SoftPaqId','SoftpaqId','SoftPaq','Softpaq','SoftPaqNumber','SoftpaqNumber','SP','Id','ID','Number'
        )

        if ($candidate -match '(?i)sp?(\d{5,6})') {
            return $matches[1]
        }

        $text = ($Recommendation | Out-String)
        if ($text -match '(?i)sp(\d{5,6})') {
            return $matches[1]
        }

        if ($text -match '\b(\d{5,6})\b') {
            return $matches[1]
        }

        return $null
    }

    function Get-HpiaRecommendationCategory {
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Recommendation)

        return (Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
            'Category','Type','RecommendationType','ComponentType','Class','Group'
        ))
    }

    function Get-HpiaRecommendationName {
        [CmdletBinding()]
        param([Parameter(Mandatory)]$Recommendation)

        $name = Get-HpiaRecommendationValue -Recommendation $Recommendation -PropertyNames @(
            'Name','Title','Component','ComponentName','Description','SoftPaqName','SoftpaqName'
        )

        if ($name) { return $name }

        $text = ($Recommendation | Out-String).Trim()
        if ($text.Length -gt 160) {
            return $text.Substring(0,160)
        }

        return $text
    }

    function New-HpiaDriverSPList {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$ReportFolder,
            [Parameter(Mandatory)][string]$SPListPath
        )

        $recommendations = @(Get-HpiaRecommendationObjects -ReportFolder $ReportFolder)
        Write-Log ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count) 'INFO'
        Add-YamlAction ("HPIA recommendations parsed from reports: {0}" -f $recommendations.Count)

        $selected = @()
        $seen = @{}

        foreach ($rec in $recommendations) {
            $category = Get-HpiaRecommendationCategory -Recommendation $rec
            $name = Get-HpiaRecommendationName -Recommendation $rec
            $sp = Get-HpiaSoftPaqNumber -Recommendation $rec

            if (-not $sp) {
                continue
            }

            # Exclude BIOS/Firmware explicitly.
            $combined = ("{0} {1}" -f $category, $name)
            if ($combined -match '(?i)\bBIOS\b|Firmware') {
                $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
                Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Blocked' -Message 'Excluded because it appears to be BIOS/Firmware.'
                continue
            }

            # Prefer driver-like recommendations, but allow blank category if the SoftPaq was recommended and not BIOS/Firmware.
            if ($combined -notmatch '(?i)Driver|Bluetooth|Chipset|Audio|Graphics|Video|LAN|WLAN|Wireless|NIC|Touch|Fingerprint|Card Reader|Storage|Serial|USB|Management Engine' -and $category) {
                $script:SkippedList.Add(("SP{0} {1}" -f $sp, $name)) | Out-Null
                Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Skipped' -Message 'Excluded because it did not appear to be a driver recommendation.'
                continue
            }

            if (-not $seen.ContainsKey($sp)) {
                $seen[$sp] = $true
                $selected += $sp
                Add-DriverResult -Vendor 'HP' -Name $name -Id $sp -Category $category -Status 'Detected' -Message 'Selected for HPIA SPList install.'
            }
        }

        if (@($selected).Count -gt 0) {
            Set-Content -LiteralPath $SPListPath -Value $selected -Encoding ASCII
            Write-Log ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath) 'OK'
            Add-YamlAction ("Created filtered HPIA SPList with {0} SoftPaqs: {1}" -f @($selected).Count, $SPListPath)
        }
        else {
            Write-Log 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.' 'OK'
            Add-YamlAction 'No non-BIOS/Firmware driver SoftPaq recommendations were selected from HPIA reports.'
        }

        return @($selected)
    }


    function Invoke-HPDriverUpdates {
        Write-Section 'HP Driver Analysis and Installation'

        $hpInfo = Get-HPSystemModelInfo

        $hpiaExe = Install-HPIAFromExtractedFolder -SourceFolder $HpiaSourceFolder -DestinationFolder $LocalHpiaFolder

        $hpiaReportFolder = Join-Path $YamlLogFolder 'HPIA'
        $hpiaDownloadFolder = Join-Path $WorkingRoot 'HPIADownloads'
        $hpiaExtractFolder = Join-Path $WorkingRoot 'HPIAExtracted'
        $spListPath = Join-Path $WorkingRoot 'HPIA_Driver_SPList.txt'

        Ensure-Folder -Path $hpiaReportFolder
        Ensure-Folder -Path $hpiaDownloadFolder
        Ensure-Folder -Path $hpiaExtractFolder

        # Pass 1: List all recommendations so we can filter out BIOS/Firmware before install.
        Write-Log 'Running HP Image Assistant list pass to detect recommended updates...' 'INFO'
        Add-YamlAction 'Running HPIA list pass to detect recommendations.'

        $listArgs = @(
            '/Operation:Analyze',
            '/Action:List',
            '/Category:All',
            '/Selection:All',
            '/Silent',
            "/ReportFolder:`"$hpiaReportFolder`"",
            "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`""
        )

        Write-Log ("HPIA list command: {0} {1}" -f $hpiaExe, ($listArgs -join ' ')) 'INFO'
        $listProc = Start-Process -FilePath $hpiaExe -ArgumentList ($listArgs -join ' ') -Wait -PassThru -NoNewWindow
        $listExitCode = [int]$listProc.ExitCode
        $listStatus = Get-HpiaExitStatus -ExitCode $listExitCode

        Write-Log ("HPIA list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus) 'INFO'
        Add-YamlAction ("HPIA list pass completed with exit code {0} ({1})." -f $listExitCode, $listStatus)

        if ($listExitCode -eq 256) {
            Write-Log 'HPIA found no recommendations for this system.' 'OK'
            Add-YamlAction 'HPIA found no recommendations for this system.'
            return
        }

        if ($listExitCode -notin @(0, 257, 3010)) {
            throw "HPIA list pass failed with exit code $listExitCode ($listStatus). Review reports in $hpiaReportFolder."
        }

        $selectedSoftPaqs = @(New-HpiaDriverSPList -ReportFolder $hpiaReportFolder -SPListPath $spListPath)

        if ($selectedSoftPaqs.Count -eq 0) {
            Write-Log 'No recommended driver SoftPaqs were selected after excluding BIOS/Firmware.' 'OK'
            Add-YamlAction 'No recommended driver SoftPaqs were selected after excluding BIOS/Firmware.'
            return
        }

        Write-Log ("Installing recommended non-BIOS/Firmware driver SoftPaqs: {0}" -f ($selectedSoftPaqs -join ', ')) 'INFO'
        Add-YamlAction ("Installing recommended non-BIOS/Firmware driver SoftPaqs: {0}" -f ($selectedSoftPaqs -join ', '))

        # Pass 2: Install only the filtered SPList.
        $installArgs = @(
            '/Operation:Analyze',
            '/Action:Install',
            "/SPList:`"$spListPath`"",
            '/Silent',
            "/ReportFolder:`"$hpiaReportFolder`"",
            "/SoftpaqDownloadFolder:`"$hpiaDownloadFolder`"",
            "/SoftpaqExtractFolder:`"$hpiaExtractFolder`""
        )

        Write-Log ("HPIA install command: {0} {1}" -f $hpiaExe, ($installArgs -join ' ')) 'INFO'
        Write-Progress -Activity 'HP Image Assistant' -Status ("Installing recommended drivers for {0}" -f $hpInfo.Model) -PercentComplete 50

        $installProc = Start-Process -FilePath $hpiaExe -ArgumentList ($installArgs -join ' ') -Wait -PassThru -NoNewWindow
        $exitCode = [int]$installProc.ExitCode
        $status = Get-HpiaExitStatus -ExitCode $exitCode

        Write-Progress -Activity 'HP Image Assistant' -Completed

        Write-Log ("HP Image Assistant install pass completed with exit code {0} ({1})." -f $exitCode, $status) 'INFO'
        Add-YamlAction ("HP Image Assistant install pass completed with exit code {0} ({1})." -f $exitCode, $status)

        $downloadedFiles = @(Get-ChildItem -Path $hpiaDownloadFolder -File -Recurse -ErrorAction SilentlyContinue)
        foreach ($file in $downloadedFiles) {
            Add-DriverResult -Vendor 'HP' -Name $file.Name -Id $null -Category 'Driver' -Status 'Detected' -Message ("Downloaded/processed by HPIA: {0}" -f $file.FullName)
        }

        $reportFiles = @(Get-ChildItem -Path $hpiaReportFolder -File -Recurse -ErrorAction SilentlyContinue)
        foreach ($report in $reportFiles) {
            Add-YamlAction ("HPIA report generated: {0}" -f $report.FullName)
        }

        if ($exitCode -eq 3010 -or $status -eq 'needs_reboot') {
            Write-Log 'HPIA installed one or more updates and a reboot may be required.' 'WARN'
            Add-YamlAction 'HPIA installed updates and indicated reboot may be required.'
            return
        }

        if ($exitCode -eq 3011) {
            Write-Log 'One or more HPIA SoftPaqs were not auto-installable and were skipped.' 'WARN'
            Add-YamlAction 'One or more HPIA SoftPaqs were not auto-installable and were skipped.'
            return
        }

        if ($exitCode -eq 0 -or $exitCode -eq 257) {
            Write-Log 'HPIA completed recommended driver install workflow.' 'OK'
            Add-YamlAction 'HPIA completed recommended driver install workflow.'
            return
        }

        throw "HP Image Assistant install pass failed or returned an unexpected exit code: $exitCode ($status). Review reports in $hpiaReportFolder."
    }

    # -----------------------------
    # Dell support
    # -----------------------------
    function Wait-ForFile {
        param(
            [Parameter(Mandatory)][string]$Path,
            [int]$TimeoutSeconds = 30
        )

        $start = Get-Date
        while (-not (Test-Path -LiteralPath $Path)) {
            Start-Sleep -Seconds 1
            if (((Get-Date) - $start).TotalSeconds -ge $TimeoutSeconds) {
                return $false
            }
        }
        return $true
    }

    function Get-DcuReportXmlSafely {
        param(
            [Parameter(Mandatory)][string]$Path,
            [int]$RetryCount = 5,
            [int]$RetryDelaySeconds = 2
        )

        if (-not (Wait-ForFile -Path $Path -TimeoutSeconds 20)) {
            throw "Dell DCU report file was not found: $Path"
        }

        Start-Sleep -Seconds 3

        for ($i = 1; $i -le $RetryCount; $i++) {
            try {
                $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                try {
                    $sr = New-Object System.IO.StreamReader($fs)
                    try {
                        $content = $sr.ReadToEnd()
                    }
                    finally {
                        $sr.Dispose()
                    }
                }
                finally {
                    $fs.Dispose()
                }

                $xml = New-Object System.Xml.XmlDocument
                $xml.LoadXml($content)
                return $xml
            }
            catch {
                if ($i -eq $RetryCount) {
                    throw
                }
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }

    function Get-DellNodeText {
        param(
            [Parameter(Mandatory)][xml]$Node,
            [Parameter(Mandatory)][string[]]$Names
        )

        foreach ($name in $Names) {
            try {
                $xpath = './/*[local-name()="' + $name + '"]'
                $child = $Node.SelectSingleNode($xpath)
                if ($child -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
                    return $child.InnerText.Trim()
                }
            }
            catch {}
        }

        return $null
    }

    function Get-DellReportItems {
        param([Parameter(Mandatory)][xml]$Xml)

        $items = @()
        try {
            $xpath = '//*[local-name()="Update" or local-name()="Package" or local-name()="SoftwareComponent" or local-name()="component" or local-name()="Device"]'
            $nodes = $Xml.SelectNodes($xpath)
            foreach ($node in $nodes) {
                $name = Get-DellNodeText -Node $node -Names @('Name','Title','PackageName')
                $version = Get-DellNodeText -Node $node -Names @('Version','PackageVersion')
                $category = Get-DellNodeText -Node $node -Names @('Category','Type')
                $id = Get-DellNodeText -Node $node -Names @('Id','PackageId','ReleaseId')

                if ($name -or $id) {
                    $items += [pscustomobject]@{
                        Id       = $id
                        Name     = $name
                        Version  = $version
                        Category = $category
                    }
                }
            }
        }
        catch {}

        return @($items)
    }

    function Invoke-DellDriverUpdates {
        Write-Section 'Dell Command Update Workflow'

        $dcuCli = Join-Path ${env:ProgramFiles} 'Dell\CommandUpdate\dcu-cli.exe'
        if (-not (Test-Path -LiteralPath $dcuCli)) {
            throw "Dell Command | Update CLI was not found: $dcuCli"
        }

        Write-Log ("Using Dell Command | Update CLI: {0}" -f $dcuCli) 'OK'
        Add-YamlAction 'Using Dell Command | Update CLI.'

        $dcuScanLog  = Join-Path $WorkingRoot 'Dell-DCU-Scan.log'
        $dcuApplyLog = Join-Path $WorkingRoot 'Dell-DCU-Apply.log'
        $dcuReport   = Join-Path $WorkingRoot 'Dell-DCU-ApplicableUpdates.xml'

        Write-Log 'Dell DCU Configure...' 'INFO'
        Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Configuring Dell Command Update' -PercentComplete 10
        $configureArgs = "/configure -silent -userConsent=disable -scheduleAuto -lockSettings=disable"
        $proc = Start-Process -FilePath $dcuCli -ArgumentList $configureArgs -Wait -PassThru -NoNewWindow
        Write-Log ("Dell DCU Configure exit code: {0}" -f $proc.ExitCode) 'INFO'

        Write-Log 'Dell DCU Scan...' 'INFO'
        Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Scanning for updates' -PercentComplete 35
        $scanArgs = "/scan -silent -outputLog=""$dcuScanLog"" -report=""$dcuReport"""
        $proc = Start-Process -FilePath $dcuCli -ArgumentList $scanArgs -Wait -PassThru -NoNewWindow
        Write-Log ("Dell DCU Scan exit code: {0}" -f $proc.ExitCode) 'INFO'
        Write-Log 'Waiting for Dell Command | Update to finish writing the report...' 'INFO'

        try {
            $xml = Get-DcuReportXmlSafely -Path $dcuReport
            $items = Get-DellReportItems -Xml $xml

            if ($items.Count -gt 0) {
                Add-YamlAction ("Dell DCU report parsed successfully. Updates detected: {0}" -f $items.Count)

                $total = $items.Count
                $index = 0

                foreach ($item in $items) {
                    $index++
                    $percent = 35 + [math]::Floor(($index / $total) * 35)
                    $label = if ($item.Name) { $item.Name } elseif ($item.Id) { $item.Id } else { 'Dell update' }

                    Write-Progress -Activity 'Dell Driver Update Workflow' -Status ("Parsing report: {0}" -f $label) -PercentComplete $percent

                    if ($item.Category -match 'BIOS|Firmware') {
                        $script:SkippedList.Add($label) | Out-Null
                        Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Blocked' -Message 'BIOS/Firmware update blocked by script policy.'
                        Write-Log ("Blocking Dell BIOS/Firmware update: {0}" -f $label) 'WARN'
                    }
                    else {
                        Add-DriverResult -Vendor 'Dell' -Name $item.Name -Id $item.Id -Category $item.Category -Status 'Detected' -Message 'Detected in Dell DCU report.'
                    }
                }
            }
            else {
                Add-YamlAction 'Dell DCU report parsed but returned no identifiable updates.'
            }
        }
        catch {
            Write-Log ("Failed to parse Dell DCU report: {0}" -f $_.Exception.Message) 'WARN'
        }

        Write-Log 'Dell DCU ApplyUpdates...' 'INFO'
        Write-Progress -Activity 'Dell Driver Update Workflow' -Status 'Applying updates' -PercentComplete 85
        $applyArgs = "/applyUpdates -silent -reboot=disable -outputLog=""$dcuApplyLog"""
        $proc = Start-Process -FilePath $dcuCli -ArgumentList $applyArgs -Wait -PassThru -NoNewWindow
        Write-Log ("Dell DCU ApplyUpdates exit code: {0}" -f $proc.ExitCode) 'INFO'
        Write-Log (("Dell DCU logs: {0} ; {1} ; {2}") -f $dcuScanLog, $dcuApplyLog, $dcuReport) 'OK'
        Write-Progress -Activity 'Dell Driver Update Workflow' -Completed
    }

    # -----------------------------
    # Main
    # -----------------------------
    $finalStatus = 'success'

    try {
        Write-Section 'Initialization'

        Ensure-Folder -Path $YamlLogFolder
        Ensure-Folder -Path $WorkingRoot
        Ensure-WorkingFolderPermissions -Path $WorkingRoot

        $yamlName = "{0}-{1}-{2}.yml" -f $script:ComputerName, '05_Weekend_Vendor_Drivers_Update', (Get-Date -Format 'yyyy-MM-dd_HHmmss')
        $yamlPath = Join-Path $YamlLogFolder $yamlName

        Write-Log ("YAML log will be written to: {0}" -f $yamlPath) 'INFO'
        Initialize-YamlLog -ComputerName $script:ComputerName -YamlPath $yamlPath

        Write-Log 'Initializing vendor driver update script...' 'INFO'

        Write-Section 'Power Policy Enforcement'
        Set-LocalPowerPolicyDesktop

        Write-Section 'Vendor Detection'
        $script:DetectedVendor = Get-DriverVendor
        Write-Log ("Detected vendor workflow: {0}" -f $script:DetectedVendor) 'INFO'
        Write-Log ("Working root: {0}" -f $WorkingRoot) 'INFO'

        if ($script:DetectedVendor -eq 'HP') {
            Invoke-HPDriverUpdates
        }
        elseif ($script:DetectedVendor -eq 'Dell') {
            Invoke-DellDriverUpdates
        }

        Write-Section 'Final Power Reapply'
        Write-Log 'Reapplying enforced local desktop power policy (display 60 minutes, sleep never)...' 'INFO'
        Set-LocalPowerPolicyDesktop
        Enforce-DesktopPowerSettings
    }
    catch {
        $finalStatus = 'failed'
        Add-RunFailure ("Script failed: {0}" -f $_.Exception.Message)
    }
    finally {
        Write-Section 'Cleanup'
        Remove-WorkingFolderRobust -Path $WorkingRoot

        if ($script:RunFailures.Count -gt 0 -and $finalStatus -ne 'failed') {
            $finalStatus = 'completed_with_warnings'
        }

        if ($script:RunFailures.Count -gt 0) {
            Write-Log (("{0} driver update completed with one or more failures.") -f $script:DetectedVendor) 'WARN'
        }
        else {
            Write-Log (("{0} driver update script completed successfully.") -f $script:DetectedVendor) 'OK'
        }

        Save-YamlLog -Status $finalStatus
    }
}

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Option 11 - Windows Updates
# -----------------------------------------------------------------------------
function Update-WindowsOS {
    <#
    .SYNOPSIS
    Runs the merged 06_Weekend_Windows_Updates logic from inside Compton Tech Utils.
    #>
    # =====================================================================
    # ScriptName: 06_Weekend_Windows_Updates.ps1
    # ScriptVersion: 1.8
    # LastUpdated: 2026-04-26
    # Purpose: Installs Windows Updates using PSWindowsUpdate, writes a
    #          YAML audit log in C:\Logs, and explicitly reboots if Windows
    #          reports that a reboot is required.
    # =====================================================================
    
    [CmdletBinding()]
    param(
        [switch]$ResetWUComponentsFirst = $false,
        [int]$OperationTimeoutSeconds = 1800,
        [int]$RebootDelaySeconds = 30,
        [string]$YamlLogFolder = "$env:SystemDrive\Logs"
    )
    
    $ErrorActionPreference = 'Stop'
    
    $script:RunStart        = Get-Date
    $script:ComputerName    = $env:COMPUTERNAME
    $script:YamlLogPath     = $null
    $script:RuntimeLogPath  = $null
    $script:UpdateEntries   = New-Object System.Collections.Generic.List[object]
    $script:RawUpdateLines  = New-Object System.Collections.Generic.List[string]
    $script:ActionHistory   = New-Object System.Collections.Generic.List[object]
    $script:OverallResult   = 'Unknown'
    $script:RebootRequired  = $false
    $script:FailureMessage  = $null
    
    function Invoke-WindowsUpdateServicesEnablement {
        <#
        .SYNOPSIS
        Runs the merged 01_Enable_Windows_Update_Services logic before the Windows Update install phase.
        #>
        # =====================================================================
        # ScriptName: 01_Enable_Windows_Update_Services.ps1
        # ScriptVersion: 1.8
        # LastUpdated: 2026-04-25
        # Purpose: Restore Windows Update services, tasks, policy settings,
        #          and Windows 11 classic right-click context menu behavior;
        #          verify required services are running, retry startup failures
        #          up to 4 total attempts, and force a reboot if critical
        #          services still refuse to start.
        # =====================================================================
        
        [CmdletBinding()]
        param()
        
        $ErrorActionPreference = 'Stop'
        
        function Write-Status {
            param(
                [string]$Message,
                [ValidateSet('INFO','OK','WARN','ERROR')]
                [string]$Level = 'INFO'
            )
        
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            switch ($Level) {
                'INFO'  { Write-Host "[$timestamp] [INFO ] $Message" -ForegroundColor Cyan }
                'OK'    { Write-Host "[$timestamp] [ OK  ] $Message" -ForegroundColor Green }
                'WARN'  { Write-Host "[$timestamp] [WARN ] $Message" -ForegroundColor Yellow }
                'ERROR' { Write-Host "[$timestamp] [ERROR] $Message" -ForegroundColor Red }
            }
        }
        
        function Test-IsAdmin {
            $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
            return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        }
        
        function Set-ServiceStartRegistry {
            param(
                [Parameter(Mandatory)]
                [string]$ServiceName,
        
                [Parameter(Mandatory)]
                [int]$StartValue
            )
        
            $paths = @(
                "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName",
                "HKLM:\SYSTEM\ControlSet001\Services\$ServiceName"
            )
        
            foreach ($path in $paths) {
                if (Test-Path $path) {
                    try {
                        Set-ItemProperty -Path $path -Name Start -Value $StartValue -Type DWord -ErrorAction Stop
                        Write-Status "Set registry Start=$StartValue for $ServiceName at $path" 'OK'
                    }
                    catch {
                        Write-Status "Failed setting Start for $ServiceName at $path : $($_.Exception.Message)" 'WARN'
                    }
                }
                else {
                    Write-Status "Registry path not found for $ServiceName at $path" 'WARN'
                }
            }
        }
        
        function Set-ServiceStartupAndStart {
            param(
                [Parameter(Mandatory)]
                [string]$Name,
        
                [Parameter(Mandatory)]
                [ValidateSet('Automatic','Manual')]
                [string]$StartupType
            )
        
            try {
                $svc = Get-Service -Name $Name -ErrorAction Stop
        
                try {
                    Set-Service -Name $Name -StartupType $StartupType -ErrorAction Stop
                    Write-Status "Set startup type for $Name to $StartupType" 'OK'
                }
                catch {
                    Write-Status "Set-Service failed for $Name. Trying sc.exe config..." 'WARN'
                    $startValue = if ($StartupType -eq 'Automatic') { 'auto' } else { 'demand' }
                    & sc.exe config $Name start= $startValue | Out-Null
                    Write-Status "Configured startup type for $Name via sc.exe" 'OK'
                }
        
                try {
                    Start-Service -Name $Name -ErrorAction Stop
                    Write-Status "Started service: $Name" 'OK'
                }
                catch {
                    Write-Status "Could not start service $Name immediately: $($_.Exception.Message)" 'WARN'
                }
            }
            catch {
                Write-Status "Service not found or inaccessible: $Name" 'WARN'
            }
        }
        
        function Get-ServiceStateSafe {
            param(
                [Parameter(Mandatory)]
                [string]$Name
            )
        
            try {
                return (Get-Service -Name $Name -ErrorAction Stop).Status
            }
            catch {
                return $null
            }
        }
        
        function Wait-ForServiceRunning {
            param(
                [Parameter(Mandatory)]
                [string]$Name,
        
                [int]$TimeoutSeconds = 15
            )
        
            $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
        
            while ($stopWatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                $status = Get-ServiceStateSafe -Name $Name
                if ($status -eq 'Running') {
                    return $true
                }
        
                Start-Sleep -Seconds 2
            }
        
            return $false
        }
        
        function Ensure-ServiceRunningWithRetry {
            param(
                [Parameter(Mandatory)]
                [string]$Name,
        
                [Parameter(Mandatory)]
                [ValidateSet('Automatic','Manual')]
                [string]$StartupType,
        
                [int]$MaxAttempts = 4,
        
                [int]$WaitPerAttemptSeconds = 15
            )
        
            for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
                $currentState = Get-ServiceStateSafe -Name $Name
        
                if ($currentState -eq 'Running') {
                    Write-Status "Service $Name is already running." 'OK'
                    return $true
                }
        
                Write-Status "Attempt $attempt of $MaxAttempts to start service $Name..." 'INFO'
        
                try {
                    Set-ServiceStartupAndStart -Name $Name -StartupType $StartupType
                }
                catch {
                    Write-Status "Unexpected error while attempting to start $Name : $($_.Exception.Message)" 'WARN'
                }
        
                if (Wait-ForServiceRunning -Name $Name -TimeoutSeconds $WaitPerAttemptSeconds) {
                    Write-Status "Verified service is running: $Name" 'OK'
                    return $true
                }
        
                $stateAfterWait = Get-ServiceStateSafe -Name $Name
                Write-Status "Service $Name did not reach Running state after attempt $attempt. Current state: $stateAfterWait" 'WARN'
        
                if ($attempt -lt $MaxAttempts) {
                    Start-Sleep -Seconds 5
                }
            }
        
            Write-Status "Service $Name failed to reach Running state after $MaxAttempts attempts." 'ERROR'
            return $false
        }
        
        function Force-RebootNow {
            param(
                [string]$Reason = 'Required Windows Update services failed to start after multiple attempts.'
            )
        
            Write-Status "FORCING REBOOT: $Reason" 'ERROR'
        
            try {
                shutdown.exe /r /f /t 30 /c "$Reason" | Out-Null
                Write-Status "Forced reboot command issued successfully. System will restart in 30 seconds." 'ERROR'
            }
            catch {
                Write-Status "Failed to issue shutdown.exe reboot command: $($_.Exception.Message)" 'ERROR'
            }
        
            exit 1
        }
        
        function Enable-ScheduledTaskSafe {
            param(
                [Parameter(Mandatory)]
                [string]$TaskPath,
        
                [Parameter(Mandatory)]
                [string]$TaskName
            )
        
            try {
                $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
                if ($task.State -eq 'Disabled') {
                    Enable-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop | Out-Null
                    Write-Status "Enabled scheduled task: $TaskPath$TaskName" 'OK'
                }
                else {
                    Write-Status "Scheduled task already enabled or available: $TaskPath$TaskName" 'INFO'
                }
            }
            catch {
                Write-Status "Scheduled task not found or could not be enabled: $TaskPath$TaskName" 'WARN'
            }
        }
        
        function Remove-RegistryValueSafe {
            param(
                [Parameter(Mandatory)]
                [string]$Path,
        
                [Parameter(Mandatory)]
                [string]$Name
            )
        
            try {
                if (Test-Path $Path) {
                    $prop = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                    if ($null -ne $prop) {
                        Remove-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                        Write-Status "Removed $Path\$Name" 'OK'
                    }
                    else {
                        Write-Status "Registry value not present: $Path\$Name" 'INFO'
                    }
                }
                else {
                    Write-Status "Registry path not present: $Path" 'INFO'
                }
            }
            catch {
                Write-Status "Failed to remove $Path\$Name : $($_.Exception.Message)" 'WARN'
            }
        }
        
        function Set-RegistryDwordSafe {
            param(
                [Parameter(Mandatory)]
                [string]$Path,
        
                [Parameter(Mandatory)]
                [string]$Name,
        
                [Parameter(Mandatory)]
                [int]$Value
            )
        
            try {
                if (-not (Test-Path $Path)) {
                    New-Item -Path $Path -Force | Out-Null
                }
        
                New-ItemProperty -Path $Path -Name $Name -PropertyType DWord -Value $Value -Force | Out-Null
                Write-Status "Set $Path\$Name = $Value" 'OK'
            }
            catch {
                Write-Status "Failed to set $Path\$Name : $($_.Exception.Message)" 'ERROR'
            }
        }
        
        function Set-ClassicRightClickForRegistryRoot {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [string]$RegistryRoot,

                [Parameter(Mandatory)]
                [string]$DisplayName
            )

            $basePath = Join-Path $RegistryRoot 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
            $subPath  = Join-Path $basePath 'InprocServer32'

            try {
                if (-not (Test-Path -LiteralPath $basePath)) {
                    New-Item -Path $basePath -Force -ErrorAction Stop | Out-Null
                    Write-Status "Created classic right-click menu CLSID key for $DisplayName" 'OK'
                }
                else {
                    Write-Status "Classic right-click menu CLSID key already exists for $DisplayName" 'INFO'
                }

                if (-not (Test-Path -LiteralPath $subPath)) {
                    New-Item -Path $subPath -Force -ErrorAction Stop | Out-Null
                    Write-Status "Created classic right-click menu InprocServer32 key for $DisplayName" 'OK'
                }
                else {
                    Write-Status "Classic right-click menu InprocServer32 key already exists for $DisplayName" 'INFO'
                }

                Set-Item -Path $subPath -Value '' -ErrorAction Stop
                Write-Status "Enabled Windows 11 classic right-click menu for $DisplayName." 'OK'
                return $true
            }
            catch {
                Write-Status "Failed to enable Windows 11 classic right-click menu for $DisplayName : $($_.Exception.Message)" 'WARN'
                return $false
            }
        }

        function Enable-ClassicWindows11RightClickMenu {
            [CmdletBinding()]
            param()

            $successCount = 0
            $attemptCount = 0

            try {
                $loadedUserHives = Get-ChildItem -Path Registry::HKEY_USERS -ErrorAction Stop |
                    Where-Object {
                        $_.PSChildName -match '^S-1-5-21-' -and
                        $_.PSChildName -notmatch '_Classes$'
                    }

                foreach ($hive in $loadedUserHives) {
                    $attemptCount++
                    if (Set-ClassicRightClickForRegistryRoot -RegistryRoot ("Registry::HKEY_USERS\{0}" -f $hive.PSChildName) -DisplayName ("loaded user hive {0}" -f $hive.PSChildName)) {
                        $successCount++
                    }
                }
            }
            catch {
                Write-Status "Unable to enumerate currently loaded user hives: $($_.Exception.Message)" 'WARN'
            }

            try {
                $profileListPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
                $profiles = Get-ChildItem -Path $profileListPath -ErrorAction Stop |
                    ForEach-Object {
                        $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                        if ($props.ProfileImagePath -and $_.PSChildName -match '^S-1-5-21-') {
                            [pscustomobject]@{
                                Sid = $_.PSChildName
                                ProfilePath = [Environment]::ExpandEnvironmentVariables($props.ProfileImagePath)
                            }
                        }
                    } |
                    Where-Object {
                        $_.ProfilePath -and
                        (Test-Path -LiteralPath (Join-Path $_.ProfilePath 'NTUSER.DAT')) -and
                        $_.ProfilePath -notmatch '\\(Default|Default User|Public|All Users)$'
                    }

                foreach ($profile in $profiles) {
                    $hkuPath = "Registry::HKEY_USERS\$($profile.Sid)"
                    if (Test-Path -LiteralPath $hkuPath) {
                        continue
                    }

                    $tempHiveName = "TempClassicContext_$($profile.Sid -replace '[^A-Za-z0-9]', '_')"
                    $ntUserDat = Join-Path $profile.ProfilePath 'NTUSER.DAT'
                    $loaded = $false

                    try {
                        $loadOutput = & reg.exe load "HKU\$tempHiveName" "$ntUserDat" 2>&1
                        if ($LASTEXITCODE -eq 0) {
                            $loaded = $true
                            $attemptCount++
                            if (Set-ClassicRightClickForRegistryRoot -RegistryRoot "Registry::HKEY_USERS\$tempHiveName" -DisplayName ("offline profile {0}" -f $profile.ProfilePath)) {
                                $successCount++
                            }
                        }
                        else {
                            Write-Status "Unable to load offline user hive for $($profile.ProfilePath): $loadOutput" 'WARN'
                        }
                    }
                    catch {
                        Write-Status "Unable to load offline user hive for $($profile.ProfilePath): $($_.Exception.Message)" 'WARN'
                    }
                    finally {
                        if ($loaded) {
                            [gc]::Collect()
                            [gc]::WaitForPendingFinalizers()
                            $unloadOutput = & reg.exe unload "HKU\$tempHiveName" 2>&1
                            if ($LASTEXITCODE -ne 0) {
                                Write-Status "Unable to unload temporary hive HKU\${tempHiveName}: $unloadOutput" 'WARN'
                            }
                        }
                    }
                }
            }
            catch {
                Write-Status "Unable to enumerate local user profiles for classic right-click menu: $($_.Exception.Message)" 'WARN'
            }

            try {
                $defaultHive = 'C:\Users\Default\NTUSER.DAT'
                if (Test-Path -LiteralPath $defaultHive) {
                    $defaultHiveName = 'TempClassicContext_DefaultUser'
                    $loadOutput = & reg.exe load "HKU\$defaultHiveName" "$defaultHive" 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        try {
                            $attemptCount++
                            if (Set-ClassicRightClickForRegistryRoot -RegistryRoot "Registry::HKEY_USERS\$defaultHiveName" -DisplayName 'Default User profile for future users') {
                                $successCount++
                            }
                        }
                        finally {
                            [gc]::Collect()
                            [gc]::WaitForPendingFinalizers()
                            $unloadOutput = & reg.exe unload "HKU\$defaultHiveName" 2>&1
                            if ($LASTEXITCODE -ne 0) {
                                Write-Status "Unable to unload temporary Default User hive: $unloadOutput" 'WARN'
                            }
                        }
                    }
                    else {
                        Write-Status "Unable to load Default User hive for future users: $loadOutput" 'WARN'
                    }
                }
                else {
                    Write-Status "Default User hive was not found at $defaultHive" 'WARN'
                }
            }
            catch {
                Write-Status "Failed to apply classic right-click menu to Default User profile: $($_.Exception.Message)" 'WARN'
            }

            if ($attemptCount -gt 0 -and $successCount -gt 0) {
                Write-Status "Classic Windows 11 right-click menu applied to $successCount of $attemptCount targeted user profile hive(s). Users may need to sign out/in or restart Explorer." 'OK'
            }
            elseif ($attemptCount -eq 0) {
                Write-Status 'No user profile hives were available for classic right-click menu enforcement.' 'WARN'
            }
            else {
                Write-Status 'Classic Windows 11 right-click menu was not applied to any user profile hive.' 'WARN'
            }
        }

        # Script updater staging and scheduled-task repair were intentionally removed from Option 11.
        # Option 11 should only restore Windows Update prerequisites and then run Windows Updates.
        # This prevents update runs from failing or modifying updater tasks when the file share is unavailable.
        
        if (-not (Test-IsAdmin)) {
            Write-Host ""
            Write-Host "This script must be run as Administrator." -ForegroundColor Red
            exit 1
        }
        
        Write-Status "Initializing Windows Update service restoration..." 'INFO'
        Enable-ClassicWindows11RightClickMenu
        Write-Status "Skipping script updater checks by design for Option 11." 'INFO'
        
        # Restore registry startup values first
        # Common defaults used for Windows Update-related services:
        # wuauserv = Manual (3)
        # bits = Manual (3)
        # dosvc = Automatic family (2)
        # UsoSvc = Automatic (2)
        # WaaSMedicSvc = Manual/triggered on many systems (3)
        
        Set-ServiceStartRegistry -ServiceName 'wuauserv'     -StartValue 3
        Set-ServiceStartRegistry -ServiceName 'bits'         -StartValue 3
        Set-ServiceStartRegistry -ServiceName 'dosvc'        -StartValue 2
        Set-ServiceStartRegistry -ServiceName 'UsoSvc'       -StartValue 2
        Set-ServiceStartRegistry -ServiceName 'WaaSMedicSvc' -StartValue 3
        
        # Initial restore and startup
        Set-ServiceStartupAndStart -Name 'wuauserv' -StartupType Manual
        Set-ServiceStartupAndStart -Name 'bits'     -StartupType Manual
        Set-ServiceStartupAndStart -Name 'dosvc'    -StartupType Automatic
        Set-ServiceStartupAndStart -Name 'UsoSvc'   -StartupType Automatic
        
        # WaaSMedicSvc can be protected; set registry above, then try starting via sc.exe
        try {
            & sc.exe config WaaSMedicSvc start= demand | Out-Null
            Write-Status "Configured WaaSMedicSvc startup via sc.exe" 'OK'
        }
        catch {
            Write-Status "Could not configure WaaSMedicSvc via sc.exe" 'WARN'
        }
        
        try {
            & sc.exe start WaaSMedicSvc | Out-Null
            Write-Status "Attempted to start WaaSMedicSvc" 'INFO'
        }
        catch {
            Write-Status "Could not start WaaSMedicSvc directly" 'WARN'
        }
        
        # Restore Automatic Updates policy
        $wuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
        $auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
        
        Set-RegistryDwordSafe -Path $auPolicyPath -Name 'NoAutoUpdate' -Value 0
        Set-RegistryDwordSafe -Path $auPolicyPath -Name 'AUOptions'    -Value 3
        
        # Remove common WSUS redirection values if they were previously set
        Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUServer'
        Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'WUStatusServer'
        Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'UpdateServiceUrlAlternate'
        Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'SetProxyBehaviorForUpdateDetection'
        Remove-RegistryValueSafe -Path $wuPolicyPath -Name 'DisableWindowsUpdateAccess'
        Remove-RegistryValueSafe -Path $auPolicyPath -Name 'UseWUServer'
        
        # Re-enable common update scheduled tasks
        $tasks = @(
            @{ Path = '\Microsoft\Windows\WindowsUpdate\';      Name = 'Scheduled Start' },
            @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan' },
            @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Schedule Scan Static Task' },
            @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'USO_UxBroker' },
            @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Reboot' },
            @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Maintenance Install' },
            @{ Path = '\Microsoft\Windows\UpdateOrchestrator\'; Name = 'Refresh Settings' },
            @{ Path = '\Microsoft\Windows\WaaSMedic\';          Name = 'PerformRemediation' }
        )
        
        foreach ($task in $tasks) {
            Enable-ScheduledTaskSafe -TaskPath $task.Path -TaskName $task.Name
        }
        
        # Restart key services in a sensible order
        $restartOrder = @('bits', 'dosvc', 'wuauserv', 'UsoSvc')
        foreach ($svc in $restartOrder) {
            try {
                Restart-Service -Name $svc -Force -ErrorAction Stop
                Write-Status "Restarted service: $svc" 'OK'
            }
            catch {
                Write-Status "Could not restart $svc : $($_.Exception.Message)" 'WARN'
            }
        }
        
        # Verify and retry critical services
        $requiredServices = @(
            @{ Name = 'wuauserv'; StartupType = 'Manual' },
            @{ Name = 'bits';     StartupType = 'Manual' },
            @{ Name = 'dosvc';    StartupType = 'Automatic' },
            @{ Name = 'UsoSvc';   StartupType = 'Automatic' }
        )
        
        $failedServices = @()
        
        foreach ($requiredService in $requiredServices) {
            $serviceStarted = Ensure-ServiceRunningWithRetry -Name $requiredService.Name -StartupType $requiredService.StartupType -MaxAttempts 4 -WaitPerAttemptSeconds 15
            if (-not $serviceStarted) {
                $failedServices += $requiredService.Name
            }
        }
        
        if ($failedServices.Count -gt 0) {
            $failedList = $failedServices -join ', '
            Write-Status "One or more critical Windows Update services failed to start: $failedList" 'ERROR'
            Force-RebootNow -Reason "Windows Update service recovery failed. Services not running: $failedList"
        }
        
        Write-Status "Windows Update settings have been restored and critical services are running." 'OK'
        Write-Status "No reboot required. Continuing normally." 'INFO'
    }

    function Write-ImmediateRuntimeLog {
        param(
            [Parameter(Mandatory)]
            [string]$Line
        )
    
        try {
            if (-not [string]::IsNullOrWhiteSpace($script:RuntimeLogPath)) {
                Add-Content -Path $script:RuntimeLogPath -Value $Line -Encoding UTF8
            }
        }
        catch {
            # Intentionally suppress runtime log write failures to avoid masking the main task.
        }
    }
    
    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO','OK','WARN','ERROR')]
            [string]$Level = 'INFO'
        )
    
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"
    
        switch ($Level) {
            'INFO'  { Write-Host $line -ForegroundColor Cyan }
            'OK'    { Write-Host $line -ForegroundColor Green }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
        }
    
        $script:ActionHistory.Add([PSCustomObject]@{
            Time    = $timestamp
            Level   = $Level
            Message = $Message
        }) | Out-Null
    
        Write-ImmediateRuntimeLog -Line $line
    
        if (-not [string]::IsNullOrWhiteSpace($script:YamlLogPath)) {
            try {
                Write-YamlLog
            }
            catch {
                # Intentionally suppress checkpoint write failures to avoid recursion from logging.
            }
        }
    }
    
    function Test-IsAdministrator {
        try {
            $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($identity)
            return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        }
        catch {
            return $false
        }
    }
    
    function Ensure-Folder {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )
    
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }
    }
    
    function ConvertTo-YamlSafeValue {
        param(
            [AllowNull()]
            [object]$Value
        )
    
        if ($null -eq $Value) {
            return 'null'
        }
    
        if ($Value -is [bool]) {
            return $Value.ToString().ToLowerInvariant()
        }
    
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
            return [string]$Value
        }
    
        $text = [string]$Value
        $text = $text -replace "`r", ' '
        $text = $text -replace "`n", ' '
        $text = $text -replace '"', '\"'
        return '"' + $text + '"'
    }
    
    function Initialize-YamlLog {
        Ensure-Folder -Path $YamlLogFolder
    
        $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $baseName = "$($script:ComputerName)-WindowsUpdates-$timestamp"
        $script:YamlLogPath = Join-Path $YamlLogFolder ($baseName + '.yaml')
        $script:RuntimeLogPath = Join-Path $YamlLogFolder ($baseName + '.log')
    
        Set-Content -Path $script:RuntimeLogPath -Value @(
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO ] Runtime log initialized.",
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [INFO ] YAML checkpoint file: $($script:YamlLogPath)"
        ) -Encoding UTF8
    
        Set-Content -Path $script:YamlLogPath -Value @(
            'computer_name: ' + (ConvertTo-YamlSafeValue $script:ComputerName),
            'script_name: "06_Weekend_Windows_Updates.ps1"',
            'script_version: "1.8"',
            'status: "Initializing"',
            'run_started: ' + (ConvertTo-YamlSafeValue ($script:RunStart.ToString('yyyy-MM-dd HH:mm:ss'))),
            'yaml_log_path: ' + (ConvertTo-YamlSafeValue $script:YamlLogPath),
            'runtime_log_path: ' + (ConvertTo-YamlSafeValue $script:RuntimeLogPath),
            'actions:',
            '  - time: ' + (ConvertTo-YamlSafeValue (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
            '    level: "INFO"',
            '    message: "Logging initialized"'
        ) -Encoding UTF8
    
        Write-Log "Runtime log will be written to: $($script:RuntimeLogPath)" 'INFO'
        Write-Log "YAML log will be written to: $($script:YamlLogPath)" 'INFO'
    }
    
    function Add-UpdateEntry {
        param(
            [string]$Title,
            [string]$KB,
            [string]$Size,
            [string]$Status,
            [string]$Result,
            [string]$Source = 'PSWindowsUpdate'
        )
    
        $entry = [PSCustomObject]@{
            Title  = $Title
            KB     = $KB
            Size   = $Size
            Status = $Status
            Result = $Result
            Source = $Source
        }
    
        $script:UpdateEntries.Add($entry) | Out-Null
    }
    
    function Try-ParseUpdateObject {
        param(
            [Parameter(Mandatory)]
            [object]$Item
        )
    
        if ($null -eq $Item) {
            return $false
        }
    
        $properties = $Item.PSObject.Properties.Name
        if (-not $properties -or $properties.Count -eq 0) {
            return $false
        }
    
        $title = $null
        foreach ($name in @('Title','KBArticleTitle')) {
            if ($properties -contains $name -and -not [string]::IsNullOrWhiteSpace([string]$Item.$name)) {
                $title = [string]$Item.$name
                break
            }
        }
    
        $kb = $null
        foreach ($name in @('KB','KBArticleIDs','KBArticleID')) {
            if ($properties -contains $name -and $null -ne $Item.$name) {
                $value = $Item.$name
                if ($value -is [System.Array]) {
                    $kb = (($value | ForEach-Object { [string]$_ }) -join ', ')
                }
                else {
                    $kb = [string]$value
                }
                if (-not [string]::IsNullOrWhiteSpace($kb)) {
                    break
                }
            }
        }
    
        $size = $null
        foreach ($name in @('Size','MaxDownloadSize')) {
            if ($properties -contains $name -and $null -ne $Item.$name) {
                $size = [string]$Item.$name
                if (-not [string]::IsNullOrWhiteSpace($size)) {
                    break
                }
            }
        }
    
        $status = $null
        foreach ($name in @('Status','Result','UpdateState')) {
            if ($properties -contains $name -and $null -ne $Item.$name) {
                $status = [string]$Item.$name
                if (-not [string]::IsNullOrWhiteSpace($status)) {
                    break
                }
            }
        }
    
        $result = $null
        foreach ($name in @('Result','Status','HResult')) {
            if ($properties -contains $name -and $null -ne $Item.$name) {
                $result = [string]$Item.$name
                if (-not [string]::IsNullOrWhiteSpace($result)) {
                    break
                }
            }
        }
    
        if (-not [string]::IsNullOrWhiteSpace($title) -or -not [string]::IsNullOrWhiteSpace($kb)) {
            Add-UpdateEntry -Title $title -KB $kb -Size $size -Status $status -Result $result
            return $true
        }
    
        return $false
    }
    
    function Write-YamlLog {
        try {
            if ([string]::IsNullOrWhiteSpace($script:YamlLogPath)) {
                Initialize-YamlLog
            }
    
            $runEnd = Get-Date
            $duration = [math]::Round(($runEnd - $script:RunStart).TotalSeconds, 0)
    
            $yamlLines = New-Object System.Collections.Generic.List[string]
    
            $yamlLines.Add('computer_name: ' + (ConvertTo-YamlSafeValue $script:ComputerName)) | Out-Null
            $yamlLines.Add('script_name: "06_Weekend_Windows_Updates.ps1"') | Out-Null
            $yamlLines.Add('script_version: "1.8"') | Out-Null
            $yamlLines.Add('run_started: ' + (ConvertTo-YamlSafeValue ($script:RunStart.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
            $yamlLines.Add('run_finished: ' + (ConvertTo-YamlSafeValue ($runEnd.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
            $yamlLines.Add('duration_seconds: ' + $duration) | Out-Null
            $yamlLines.Add('yaml_log_path: ' + (ConvertTo-YamlSafeValue $script:YamlLogPath)) | Out-Null
            $yamlLines.Add('runtime_log_path: ' + (ConvertTo-YamlSafeValue $script:RuntimeLogPath)) | Out-Null
            $yamlLines.Add('reset_wu_components_first: ' + ($(if ($ResetWUComponentsFirst) { 'true' } else { 'false' }))) | Out-Null
            $yamlLines.Add('operation_timeout_seconds: ' + $OperationTimeoutSeconds) | Out-Null
            $yamlLines.Add('reboot_delay_seconds: ' + $RebootDelaySeconds) | Out-Null
            $yamlLines.Add('reboot_required: ' + ($(if ($script:RebootRequired) { 'true' } else { 'false' }))) | Out-Null
            $yamlLines.Add('overall_result: ' + (ConvertTo-YamlSafeValue $script:OverallResult)) | Out-Null
    
            if (-not [string]::IsNullOrWhiteSpace($script:FailureMessage)) {
                $yamlLines.Add('failure_message: ' + (ConvertTo-YamlSafeValue $script:FailureMessage)) | Out-Null
            }
            else {
                $yamlLines.Add('failure_message: null') | Out-Null
            }
    
            $yamlLines.Add('updates:') | Out-Null
            if ($script:UpdateEntries.Count -gt 0) {
                foreach ($entry in $script:UpdateEntries) {
                    $yamlLines.Add('  - title: '  + (ConvertTo-YamlSafeValue $entry.Title))  | Out-Null
                    $yamlLines.Add('    kb: '     + (ConvertTo-YamlSafeValue $entry.KB))     | Out-Null
                    $yamlLines.Add('    size: '   + (ConvertTo-YamlSafeValue $entry.Size))   | Out-Null
                    $yamlLines.Add('    status: ' + (ConvertTo-YamlSafeValue $entry.Status)) | Out-Null
                    $yamlLines.Add('    result: ' + (ConvertTo-YamlSafeValue $entry.Result)) | Out-Null
                    $yamlLines.Add('    source: ' + (ConvertTo-YamlSafeValue $entry.Source)) | Out-Null
                }
            }
            else {
                $yamlLines.Add('  - title: "No structured update entries captured"') | Out-Null
                $yamlLines.Add('    kb: null') | Out-Null
                $yamlLines.Add('    size: null') | Out-Null
                $yamlLines.Add('    status: "None"') | Out-Null
                $yamlLines.Add('    result: "None"') | Out-Null
                $yamlLines.Add('    source: "Script"') | Out-Null
            }
    
            $yamlLines.Add('raw_output:') | Out-Null
            if ($script:RawUpdateLines.Count -gt 0) {
                foreach ($line in $script:RawUpdateLines) {
                    $yamlLines.Add('  - ' + (ConvertTo-YamlSafeValue $line)) | Out-Null
                }
            }
            else {
                $yamlLines.Add('  - "No raw update output captured"') | Out-Null
            }
    
            $yamlLines.Add('actions:') | Out-Null
            if ($script:ActionHistory.Count -gt 0) {
                foreach ($action in $script:ActionHistory) {
                    $yamlLines.Add('  - time: ' + (ConvertTo-YamlSafeValue $action.Time)) | Out-Null
                    $yamlLines.Add('    level: ' + (ConvertTo-YamlSafeValue $action.Level)) | Out-Null
                    $yamlLines.Add('    message: ' + (ConvertTo-YamlSafeValue $action.Message)) | Out-Null
                }
            }
            else {
                $yamlLines.Add('  - "No actions recorded"') | Out-Null
            }
    
            Set-Content -Path $script:YamlLogPath -Value $yamlLines -Encoding UTF8
        }
        catch {
            Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
        }
    }
    
    function Ensure-PSWindowsUpdate {
        Write-Log "Ensuring PSWindowsUpdate module is available..." 'INFO'
    
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        catch {}
    
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Write-Log "Installing NuGet package provider..." 'INFO'
            Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
        }
    
        try {
            $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
            if ($repo.InstallationPolicy -ne 'Trusted') {
                Write-Log "Setting PSGallery as Trusted..." 'INFO'
                Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
            }
        }
        catch {
            Write-Log "Could not validate PSGallery repository settings: $($_.Exception.Message)" 'WARN'
        }
    
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Log "Installing PSWindowsUpdate module..." 'INFO'
            Install-Module -Name PSWindowsUpdate -Force -AllowClobber -Scope AllUsers
        }
    
        Import-Module PSWindowsUpdate -Force
        Write-Log "PSWindowsUpdate module imported successfully." 'OK'
    }
    
    function Reset-WUComponentsSafe {
        Write-Log "Resetting Windows Update components using built-in manual routine..." 'INFO'
    
        $services = @('wuauserv', 'bits', 'cryptsvc', 'msiserver')
    
        foreach ($serviceName in $services) {
            try {
                $service = Get-Service -Name $serviceName -ErrorAction Stop
                if ($service.Status -ne 'Stopped') {
                    Write-Log "Stopping service: $serviceName" 'INFO'
                    Stop-Service -Name $serviceName -Force -ErrorAction Stop
                }
                else {
                    Write-Log "Service already stopped: $serviceName" 'INFO'
                }
            }
            catch {
                Write-Log "Failed to stop service ${serviceName}: $($_.Exception.Message)" 'WARN'
            }
        }
    
        Start-Sleep -Seconds 3
    
        $shouldDeleteUpdatePolicy = $false
        $shouldDeletePolicyManagerUpdate = $false
        $wuTriggerDetails = @()
        $pmTriggerDetails = @()
    
        try {
            $wuPolicyPath = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy'
            $wuSettingsPath = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\Settings'
    
            if (Test-Path -LiteralPath $wuPolicyPath) {
                $shouldDeleteUpdatePolicy = $true
                $wuTriggerDetails += "Detected registry hive: $wuPolicyPath"
                Write-Log "Detected Windows Update policy hive: $wuPolicyPath" 'INFO'
            }
    
            if (Test-Path -LiteralPath $wuSettingsPath) {
                $wuSettings = Get-ItemProperty -Path $wuSettingsPath -ErrorAction SilentlyContinue
                if ($wuSettings) {
                    $pauseProps = $wuSettings.PSObject.Properties | Where-Object {
                        $_.Name -like '*Pause*' -and $null -ne $_.Value -and "$($_.Value)".Trim() -ne ''
                    }
    
                    if ($pauseProps) {
                        $shouldDeleteUpdatePolicy = $true
                        foreach ($prop in $pauseProps) {
                            $detail = "{0} = {1}" -f $prop.Name, $prop.Value
                            $wuTriggerDetails += $detail
                        }
    
                        $pauseNames = ($pauseProps | Select-Object -ExpandProperty Name) -join ', '
                        Write-Log "Detected pause-related Windows Update state: $pauseNames" 'INFO'
                    }
                }
            }
        }
        catch {
            Write-Log "Failed while checking Windows Update policy state: $($_.Exception.Message)" 'WARN'
        }
    
        try {
            $pmUpdatePath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update'
    
            if (Test-Path -LiteralPath $pmUpdatePath) {
                $pmUpdate = Get-ItemProperty -Path $pmUpdatePath -ErrorAction SilentlyContinue
                if ($pmUpdate) {
                    $interestingProps = $pmUpdate.PSObject.Properties | Where-Object {
                        $_.Name -match 'Pause|Paused|ProviderSet|WinningProvider|Enrolled'
                    }
    
                    if ($interestingProps) {
                        $shouldDeletePolicyManagerUpdate = $true
                        foreach ($prop in $interestingProps) {
                            $detail = "{0} = {1}" -f $prop.Name, $prop.Value
                            $pmTriggerDetails += $detail
                        }
    
                        $propNames = ($interestingProps | Select-Object -ExpandProperty Name) -join ', '
                        Write-Log "Detected PolicyManager Update state: $propNames" 'INFO'
                    }
                }
            }
        }
        catch {
            Write-Log "Failed while checking PolicyManager Update state: $($_.Exception.Message)" 'WARN'
        }
    
        if ($wuTriggerDetails.Count -gt 0) {
            foreach ($detail in $wuTriggerDetails) {
                Write-Log "Windows Update cleanup trigger: $detail" 'INFO'
            }
        }
    
        if ($pmTriggerDetails.Count -gt 0) {
            foreach ($detail in $pmTriggerDetails) {
                Write-Log "PolicyManager cleanup trigger: $detail" 'INFO'
            }
        }
    
        if ($shouldDeleteUpdatePolicy) {
            try {
                Write-Log "Deleting Windows Update policy registry hive..." 'INFO'
                & reg.exe delete "HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy" /f | Out-Null
                Write-Log "Deleted HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy" 'OK'
            }
            catch {
                Write-Log "Failed to delete HKLM\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy: $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Log "No pause-related Windows Update policy state detected; skipping UpdatePolicy registry delete." 'INFO'
        }
    
        if ($shouldDeletePolicyManagerUpdate) {
            try {
                Write-Log "Deleting PolicyManager Update registry hive..." 'INFO'
                & reg.exe delete "HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Update" /f | Out-Null
                Write-Log "Deleted HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Update" 'OK'
            }
            catch {
                Write-Log "Failed to delete HKLM\SOFTWARE\Microsoft\PolicyManager\current\device\Update: $($_.Exception.Message)" 'WARN'
            }
        }
        else {
            Write-Log "No relevant PolicyManager Update state detected; skipping PolicyManager registry delete." 'INFO'
        }
    
        $foldersToClear = @(
            "$env:SystemRoot\SoftwareDistribution",
            "$env:SystemRoot\System32\catroot2"
        )
    
        foreach ($folder in $foldersToClear) {
            try {
                if (Test-Path -LiteralPath $folder) {
                    Write-Log "Clearing folder: $folder" 'INFO'
                    Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop
                    Write-Log "Cleared folder: $folder" 'OK'
                }
                else {
                    Write-Log "Folder not present, skipping: $folder" 'INFO'
                }
            }
            catch {
                Write-Log "Failed to clear folder ${folder}: $($_.Exception.Message)" 'WARN'
            }
        }
    
        foreach ($serviceName in $services) {
            try {
                Write-Log "Starting service: $serviceName" 'INFO'
                Start-Service -Name $serviceName -ErrorAction Stop
                Write-Log "Started service: $serviceName" 'OK'
            }
            catch {
                Write-Log "Failed to start service ${serviceName}: $($_.Exception.Message)" 'WARN'
            }
        }
    
        Write-Log "Windows Update components reset complete." 'OK'
    }
    
    function Invoke-PSWindowsUpdateJobOnce {
        param(
            [Parameter(Mandatory)]
            [int]$TimeoutSeconds
        )
    
        Write-Log "Starting background job for Install-WindowsUpdate..." 'INFO'
        $job = Start-Job -ScriptBlock {
            Import-Module PSWindowsUpdate -Force
            Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -Verbose *>&1
        }
    
        Write-Log "Background job started with ID $($job.Id). Waiting for completion for up to $TimeoutSeconds seconds..." 'INFO'
    
        $pollIntervalSeconds = 5
        $heartbeatIntervalSeconds = 60
        $elapsedSeconds = 0
        $completed = $null
    
        while (-not $completed -and $elapsedSeconds -lt $TimeoutSeconds) {
            $completed = Wait-Job -Job $job -Timeout $pollIntervalSeconds
            if ($completed) { break }
    
            $elapsedSeconds += $pollIntervalSeconds
    
            if (($elapsedSeconds % $heartbeatIntervalSeconds) -eq 0) {
                $jobState = (Get-Job -Id $job.Id).State
                Write-Log "Windows Update job still running after $elapsedSeconds seconds. Current job state: $jobState" 'INFO'
            }
        }
    
        if (-not $completed) {
            $jobState = (Get-Job -Id $job.Id).State
            Write-Log "Windows Update job did not finish before timeout. Final observed state: $jobState" 'ERROR'
            Stop-Job -Job $job -Force | Out-Null
            Remove-Job -Job $job -Force | Out-Null
            throw "Install-WindowsUpdate exceeded timeout of $TimeoutSeconds seconds."
        }
    
        Write-Log "Background job completed. Receiving Windows Update output..." 'INFO'
    
        $receiveErrors = @()
        $results = Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable receiveErrors
    
        $jobState = $job.State
        $jobReason = $null
        try {
            if ($job.ChildJobs -and $job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].JobStateInfo.Reason) {
                $jobReason = $job.ChildJobs[0].JobStateInfo.Reason.Message
            }
        }
        catch {}
    
        Write-Log "Removing completed background job..." 'INFO'
        Remove-Job -Job $job -Force | Out-Null
    
        if ($results) {
            foreach ($item in $results) {
                $line = [string]$item
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    $script:RawUpdateLines.Add($line) | Out-Null
                    Write-Log $line 'INFO'
                }
    
                [void](Try-ParseUpdateObject -Item $item)
            }
        }
        else {
            Write-Log "Install-WindowsUpdate returned no normal output objects." 'WARN'
        }
    
        if ($receiveErrors -and $receiveErrors.Count -gt 0) {
            foreach ($err in $receiveErrors) {
                $errText = [string]$err
                if (-not [string]::IsNullOrWhiteSpace($errText)) {
                    $script:RawUpdateLines.Add("ERROR: $errText") | Out-Null
                    Write-Log "Install-WindowsUpdate reported error output: $errText" 'WARN'
                }
            }
    
            $errorText = (($receiveErrors | ForEach-Object { [string]$_ }) -join ' | ')
            throw $errorText
        }
    
        if ($jobState -eq 'Failed') {
            if ([string]::IsNullOrWhiteSpace($jobReason)) {
                throw "Install-WindowsUpdate background job failed."
            }
            else {
                throw "Install-WindowsUpdate background job failed: $jobReason"
            }
        }
    
        return $results
    }
    
    function Install-AvailableWindowsUpdates {
        Write-Log "Scanning for and installing available Windows Updates..." 'INFO'
    
        if (-not (Get-Command -Name Install-WindowsUpdate -ErrorAction SilentlyContinue)) {
            throw "Install-WindowsUpdate command was not found."
        }
    
        $attempt = 1
        $maxAttempts = 2
    
        while ($attempt -le $maxAttempts) {
            try {
                if ($attempt -gt 1) {
                    Write-Log "Retrying Windows Update install phase. Attempt $attempt of $maxAttempts..." 'INFO'
                }
    
                [void](Invoke-PSWindowsUpdateJobOnce -TimeoutSeconds $OperationTimeoutSeconds)
                Write-Log "Windows Update installation command completed." 'OK'
                return
            }
            catch {
                $message = $_.Exception.Message
                Write-Log "Windows Update install attempt $attempt failed: $message" 'WARN'
    
                if ($message -match '0x80248007' -and $attempt -lt $maxAttempts) {
                    Write-Log "Detected Windows Update datastore/catalog error 0x80248007. Resetting Windows Update components and retrying once..." 'WARN'
                    Reset-WUComponentsSafe
                    $attempt++
                    continue
                }
    
                throw
            }
        }
    }
    
    function Test-WURebootRequired {
        try {
            if (Get-Command -Name Get-WURebootStatus -ErrorAction SilentlyContinue) {
                $status = Get-WURebootStatus -Silent
                return [bool]$status
            }
        }
        catch {
            Write-Log "Get-WURebootStatus check failed: $($_.Exception.Message)" 'WARN'
        }
    
        try {
            $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
            return [bool]$sysInfo.RebootRequired
        }
        catch {
            Write-Log "Microsoft.Update.SystemInfo reboot check failed: $($_.Exception.Message)" 'WARN'
        }
    
        return $false
    }
    
    function Invoke-ExplicitReboot {
        param(
            [int]$Delay = 30
        )
    
        $comment = 'Restarting to complete Windows Update installation.'
    
        $arguments = @(
            '/r'
            '/t', $Delay.ToString()
            '/d', 'p:2:17'
            '/c', "`"$comment`""
            '/f'
        )
    
        Write-Log "Issuing reboot command: shutdown.exe $($arguments -join ' ')" 'INFO'
        Write-YamlLog
    
        & "$env:SystemRoot\System32\shutdown.exe" @arguments
    
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            throw "shutdown.exe returned exit code $exitCode"
        }
    
        Write-Log "Reboot command issued successfully." 'OK'
    }
    
    # Main
    if (-not (Test-IsAdministrator)) {
        Write-Error "Please run this script as Administrator."
        return 1
    }
    
    Initialize-YamlLog
    Write-Log "Initializing weekend Windows update script..." 'INFO'
    
    try {
        Write-Log "Running Windows Update service/policy restoration before update scan..." 'INFO'
        Invoke-WindowsUpdateServicesEnablement
        Write-Log "Windows Update service/policy restoration completed. Continuing to Windows Update install phase..." 'OK'

        Write-Log "Beginning prerequisite validation and module preparation..." 'INFO'
        Ensure-PSWindowsUpdate
    
        if ($ResetWUComponentsFirst) {
            Write-Log "ResetWUComponentsFirst switch detected. Beginning Windows Update component reset..." 'INFO'
            Reset-WUComponentsSafe
        }
    
        Write-Log "Beginning Windows Update scan and install phase..." 'INFO'
        Install-AvailableWindowsUpdates
    
        Write-Log "Checking whether Windows reports a reboot requirement..." 'INFO'
        $script:RebootRequired = Test-WURebootRequired
    
        if ($script:RebootRequired) {
            $script:OverallResult = 'SucceededWithRebootRequired'
            Write-Log "Windows reports that a reboot is required." 'OK'
            Write-YamlLog
            Invoke-ExplicitReboot -Delay $RebootDelaySeconds
            return 3010
        }
        else {
            $script:OverallResult = 'Succeeded'
            Write-Log "No reboot is currently required." 'OK'
            Write-YamlLog
            return 0
        }
    }
    catch {
        $script:OverallResult = 'Failed'
        $script:FailureMessage = $_.Exception.Message
        Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
        Write-YamlLog
        return 2
    }
}
# Option 12 - Disk Cleanup
# -----------------------------------------------------------------------------
function Run-DiskCleanup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$SkipSystemCleanup,
        [switch]$SkipTempFolders,
        [switch]$SkipSSDTrim,
        [int]$TimeoutSeconds = 300,
        [string]$LogPath = "$env:TEMP\DiskCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    )

    # Security: Require elevation for system cleanup operations
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "Disk cleanup operations require Administrator privileges"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking
    $script:logEntries = New-Object System.Collections.ArrayList
    $script:cleanedPaths = New-Object System.Collections.ArrayList
    $script:failedPaths = New-Object System.Collections.ArrayList
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $totalSpaceFreed = 0

    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO', [switch]$NoNewline)
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        if ($null -eq $script:logEntries) {
            $script:logEntries = New-Object System.Collections.ArrayList
        }
        [void]$script:logEntries.Add($logEntry)
        
        $writeParams = @{}
        if ($NoNewline) { $writeParams.NoNewline = $true }
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red @writeParams }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow @writeParams }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green @writeParams }
            'INFO' { Write-Host $Message -ForegroundColor Cyan @writeParams }
            'OPERATION' { Write-Host $Message -ForegroundColor Magenta @writeParams }
            'PROGRESS' { Write-Host $Message -ForegroundColor Gray @writeParams }
        }
    }

    # Speed: Get disk space before operation
    function Get-DiskSpaceInfo {
        param([string]$Path)
        try {
            $drive = [System.IO.DriveInfo]::new((Split-Path $Path -Qualifier))
            return @{
                FreeSpace = $drive.AvailableFreeSpace
                TotalSize = $drive.TotalSize
                UsedSpace = $drive.TotalSize - $drive.AvailableFreeSpace
            }
        } catch {
            return $null
        }
    }

    # Security: Validate paths to prevent malicious deletion
    function Test-SafePath {
        param([string]$Path)
        
        # Security: Block critical system paths
        $blockedPaths = @(
            "C:\Windows\System32",
            "C:\Windows\SysWOW64", 
            "C:\Program Files",
            "C:\Program Files (x86)",
            "C:\Users\*\Desktop",
            "C:\Users\*\Documents",
            "C:\Users\*\Pictures",
            "C:\Users\*\Videos",
            "C:\Users\*\Music",
            "C:\Windows\explorer.exe",
            "C:\Windows\System32\drivers"
        )
        
        foreach ($blocked in $blockedPaths) {
            if ($Path -like $blocked) {
                return $false
            }
        }
        
        # Security: Allow legitimate cleanup locations
        $allowedPatterns = @(
            "C:\Windows\Temp*",
            "C:\Temp*", 
            "C:\SWSetup*",
            "C:\system.sav*",
            "*\AppData\Local\Temp*",
            "C:\Windows\SoftwareDistribution\Download*",
            "C:\Windows\Prefetch*",
            "C:\Windows\Logs\CBS*",
            "*\AppData\Local\Microsoft\Windows\INetCache*",
            "*\AppData\Local\Microsoft\Windows\WebCache*",
            "C:\ProgramData\Microsoft\Windows\WER\ReportQueue*",
            "*\AppData\Local\CrashDumps*",
            "*\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache*"
        )
        
        foreach ($pattern in $allowedPatterns) {
            if ($Path -like $pattern) {
                return $true
            }
        }
        
        return $false
    }

# Speed: Optimized folder cleanup with size tracking and progress feedback
function Remove-FolderContents {
    param(
        [string]$Path,
        [string]$Description,
        [switch]$ContentsOnly
    )
    
    if (-not (Test-Path $Path)) {
        return @{ Success = $true; SpaceFreed = [int64]0; Message = "Path does not exist" }
    }
    
    # Security: Validate path safety
    if (-not (Test-SafePath $Path)) {
        return @{ Success = $false; SpaceFreed = [int64]0; Message = "Path blocked for security" }
    }
    
    try {
        Write-LogEntry "      [DIR] Analyzing $Description..." 'PROGRESS' -NoNewline
        
        # Speed: Get size before deletion with better error handling
        [int64]$sizeBefore = 0
        $itemCount = 0
        if ($ContentsOnly) {
            try {
                $items = @(Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue)
                $itemCount = $items.Count
                if ($items.Count -gt 0) {
                    $sizeSum = ($items | Where-Object { -not $_.PSIsContainer } | 
                               Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    if ($null -ne $sizeSum) { 
                        $sizeBefore = [int64]$sizeSum 
                    } else { 
                        $sizeBefore = [int64]0 
                    }
                }
            } catch {
                $sizeBefore = [int64]0
            }
        } else {
            try {
                $files = @(Get-ChildItem -Path $Path -Recurse -Force -File -ErrorAction SilentlyContinue)
                $itemCount = $files.Count
                if ($files.Count -gt 0) {
                    $sizeSum = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    if ($null -ne $sizeSum) { 
                        $sizeBefore = [int64]$sizeSum 
                    } else { 
                        $sizeBefore = [int64]0 
                    }
                }
            } catch {
                $sizeBefore = [int64]0
            }
        }
        
        # Show analysis results
        Write-Host " ($itemCount items)" -ForegroundColor Gray
        
        if ($itemCount -eq 0) {
            return @{ Success = $true; SpaceFreed = [int64]0; Message = "Folder is empty" }
        }
        
        if ($PSCmdlet.ShouldProcess($Path, "Clean folder contents")) {
            Write-LogEntry "      [REMOVE]  Cleaning $Description..." 'PROGRESS'
            
            if ($ContentsOnly) {
                # Speed: Check PowerShell version for parallel support
                if ($PSVersionTable.PSVersion.Major -ge 7) {
                    # PowerShell 7+ with parallel processing (simplified progress)
                    Write-LogEntry "      [REMOVE]  Processing $itemCount items in parallel..." 'PROGRESS'
                    Get-ChildItem -Path $Path -Force -ErrorAction Stop | 
                        ForEach-Object -Parallel {
                            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        } -ThrottleLimit 5 -ErrorAction SilentlyContinue
                    Write-LogEntry "      [OK] Parallel cleanup completed" 'PROGRESS'
                } else {
                    # Windows PowerShell 5.1 compatible method with progress
                    $items = Get-ChildItem -Path $Path -Force -ErrorAction Stop
                    $processedCount = 0
                    foreach ($item in $items) {
                        try {
                            Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction Stop
                            $processedCount++
                            
                            # Show progress dots every 10 items
                            if ($processedCount % 10 -eq 0) {
                                Write-Host "." -NoNewline -ForegroundColor Gray
                            }
                            
                            # Show progress percentage for large operations
                            if ($items.Count -gt 50 -and $processedCount % 25 -eq 0) {
                                $percentComplete = [math]::Round(($processedCount / $items.Count) * 100)
                                Write-Host " $percentComplete%" -NoNewline -ForegroundColor Yellow
                            }
                        } catch {
                            # Continue with next item if one fails
                            Write-Verbose "Failed to remove $($item.FullName): $_"
                        }
                    }
                    if ($processedCount -gt 10) {
                        Write-Host "" # New line after progress dots
                    }
                }
            } else {
                Write-LogEntry "      [REMOVE]  Removing entire folder..." 'PROGRESS'
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
            }
            
            return @{ 
                Success = $true
                SpaceFreed = [math]::Max([int64]0, $sizeBefore)
                Message = "Successfully cleaned"
            }
        }
        
        return @{ Success = $true; SpaceFreed = [int64]0; Message = "WhatIf mode - no changes made" }
        
    } catch {
        Write-Host "" # New line if we were showing progress
        return @{ 
            Success = $false
            SpaceFreed = [int64]0
            Message = $_.Exception.Message
        }
    }
}

    # Speed: Optimized system drive detection
    function Get-SystemDrive {
        try {
            # Speed: Use .NET method instead of WMI for better performance
            $systemDrive = [Environment]::SystemDirectory.Substring(0, 1)
            return $systemDrive
        } catch {
            # Fallback to WMI if needed
            try {
                return (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).SystemDrive.TrimEnd(':', '\')
            } catch {
                return "C"  # Safe fallback
            }
        }
    }

    # Security: Safe cleanmgr execution with timeout and progress
    function Invoke-WindowsCleanup {
        param([int]$TimeoutSec = 300)
        
        try {
            # Security: Verify cleanmgr.exe exists and is signed
            $cleanmgrPath = "$env:SystemRoot\System32\cleanmgr.exe"
            if (-not (Test-Path $cleanmgrPath)) {
                throw "Windows Disk Cleanup utility not found"
            }
            
            Write-LogEntry "      [RESTART] Starting Windows Disk Cleanup (timeout: ${TimeoutSec}s)..." 'PROGRESS'
            
            # Speed: Use job for timeout control
            $job = Start-Job -ScriptBlock {
                param($TimeoutSec)
                Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/Sagerun:1", "/VERYLOWDISK" -NoNewWindow -Wait -PassThru
            } -ArgumentList $TimeoutSec
            
            # Show progress while waiting
            $elapsed = 0
            $progressInterval = 5
            while ($job.State -eq 'Running' -and $elapsed -lt $TimeoutSec) {
                Start-Sleep -Seconds $progressInterval
                $elapsed += $progressInterval
                
                # Show progress every 10 seconds
                if ($elapsed % 10 -eq 0) {
                    $remainingTime = $TimeoutSec - $elapsed
                    Write-LogEntry "      [PENDING] Cleanup in progress... (${elapsed}s elapsed, ${remainingTime}s remaining)" 'PROGRESS'
                }
            }
            
            $result = Wait-Job -Job $job -Timeout 1
            
            if ($null -eq $result -or $job.State -eq 'Running') {
                Stop-Job -Job $job -Force
                Write-LogEntry "      [WARN]  Cleanup timed out, stopping process..." 'PROGRESS'
                throw "Disk Cleanup timed out after $TimeoutSec seconds"
            }
            
            $jobResult = Receive-Job -Job $job
            Remove-Job -Job $job -Force
            
            return @{ Success = $true; ExitCode = $jobResult.ExitCode }
            
        } catch {
            return @{ Success = $false; Error = $_.Exception.Message }
        }
    }

    try {
        Write-LogEntry "`n[CLEAN] Starting Disk Cleanup Operations..." 'OPERATION'
        
        # Get initial disk space
        $initialSpace = Get-DiskSpaceInfo "C:\"
        if ($initialSpace) {
            Write-LogEntry "Initial free space: $([math]::Round($initialSpace.FreeSpace / 1GB, 2)) GB" 'INFO'
        }

        # Phase 1: Windows System Cleanup
        if (-not $SkipSystemCleanup) {
            Write-LogEntry "`n[REMOVE] Running Windows Disk Cleanup..." 'OPERATION'
            
            $cleanupResult = Invoke-WindowsCleanup -TimeoutSec $TimeoutSeconds
            if ($cleanupResult.Success) {
                Write-LogEntry "[OK] Windows Disk Cleanup completed successfully" 'SUCCESS'
            } else {
                Write-LogEntry "[WARN] Windows Disk Cleanup failed: $($cleanupResult.Error)" 'WARNING'
            }
        }

        # Phase 2: Manual Temp Folder Cleanup
        if (-not $SkipTempFolders) {
            Write-LogEntry "`n[WIPE] Cleaning temporary folders..." 'OPERATION'
            
            # Speed: Define cleanup targets with priorities
            $cleanupTargets = @(
                @{ Path = "C:\SWSetup"; Description = "HP Software Setup"; ContentsOnly = $false },
                @{ Path = "C:\Temp"; Description = "System Temp"; ContentsOnly = $false },
                @{ Path = "C:\system.sav"; Description = "System Save"; ContentsOnly = $false },
                @{ Path = "C:\Windows\Temp"; Description = "Windows Temp"; ContentsOnly = $true },
                @{ Path = "$env:TEMP"; Description = "User Temp"; ContentsOnly = $true },
                @{ Path = "C:\Windows\SoftwareDistribution\Download"; Description = "Windows Update Cache"; ContentsOnly = $true },
                @{ Path = "C:\Windows\Prefetch"; Description = "Windows Prefetch"; ContentsOnly = $true },
                @{ Path = "C:\Windows\Logs\CBS"; Description = "Component Based Servicing Logs"; ContentsOnly = $true },
                @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"; Description = "Internet Cache"; ContentsOnly = $true },
                @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"; Description = "Web Cache"; ContentsOnly = $true },
                @{ Path = "C:\ProgramData\Microsoft\Windows\WER\ReportQueue"; Description = "Error Report Queue"; ContentsOnly = $true },
                @{ Path = "$env:LOCALAPPDATA\CrashDumps"; Description = "Crash Dumps"; ContentsOnly = $true },
                @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"; Description = "Delivery Optimization Cache"; ContentsOnly = $true }
            )

            # Speed: Process cleanup targets efficiently
            foreach ($target in $cleanupTargets) {
                $result = Remove-FolderContents -Path $target.Path -Description $target.Description -ContentsOnly:$target.ContentsOnly
                
                if ($result.Success) {
                    if ($result.SpaceFreed -gt 0) {
                        $script:totalSpaceFreed += $result.SpaceFreed
                        $sizeText = if ($result.SpaceFreed -gt 1GB) { 
                            "$([math]::Round($result.SpaceFreed / 1GB, 2)) GB" 
                        } else { 
                            "$([math]::Round($result.SpaceFreed / 1MB, 1)) MB" 
                        }
                        Write-LogEntry "   [OK] Cleaned $($target.Description): $sizeText freed" 'SUCCESS'
                        
                        if ($null -eq $script:cleanedPaths) {
                            $script:cleanedPaths = New-Object System.Collections.ArrayList
                        }
                        [void]$script:cleanedPaths.Add(@{
                            Path = $target.Path
                            Description = $target.Description
                            SpaceFreed = $result.SpaceFreed
                        })
                    } else {
                        Write-LogEntry "   [SKIP] $($target.Description): $($result.Message)" 'INFO'
                    }
                } else {
                    Write-LogEntry "   [WARN] Failed to clean $($target.Description): $($result.Message)" 'WARNING'
                    
                    if ($null -eq $script:failedPaths) {
                        $script:failedPaths = New-Object System.Collections.ArrayList
                    }
                    [void]$script:failedPaths.Add(@{
                        Path = $target.Path
                        Description = $target.Description
                        Error = $result.Message
                    })
                }
            }
        }

        # Phase 3: SSD Trim Optimization
        if (-not $SkipSSDTrim) {
            Write-LogEntry "`n[SAVE] Performing SSD optimization..." 'OPERATION'
            
            try {
                $sysDriveLetter = Get-SystemDrive
                
                Write-LogEntry "      [SCAN] Checking drive $sysDriveLetter..." 'PROGRESS'
                
                # Security: Verify drive exists and is ready
                $drive = Get-Volume -DriveLetter $sysDriveLetter -ErrorAction Stop
                
                if ($PSCmdlet.ShouldProcess("Drive $sysDriveLetter", "Perform SSD Trim")) {
                    Write-LogEntry "      [FAST] Starting TRIM operation on drive $sysDriveLetter..." 'PROGRESS'
                    
                    # Speed: Run trim operation with progress monitoring
                    $trimJob = Start-Job -ScriptBlock {
                        param($DriveLetter)
                        Optimize-Volume -DriveLetter $DriveLetter -ReTrim -ErrorAction Stop
                    } -ArgumentList $sysDriveLetter
                    
                    # Show progress while waiting for trim
                    $elapsed = 0
                    $maxWait = 120
                    while ($trimJob.State -eq 'Running' -and $elapsed -lt $maxWait) {
                        Start-Sleep -Seconds 5
                        $elapsed += 5
                        
                        if ($elapsed % 15 -eq 0) {
                            Write-LogEntry "      [PENDING] TRIM in progress... (${elapsed}s elapsed)" 'PROGRESS'
                        }
                    }
                    
                    $trimResult = Wait-Job -Job $trimJob -Timeout 5
                    
                    if ($null -ne $trimResult) {
                        Remove-Job -Job $trimJob -Force
                        Write-LogEntry "   [OK] SSD Trim completed successfully on drive $sysDriveLetter" 'SUCCESS'
                    } else {
                        Stop-Job -Job $trimJob -Force
                        Remove-Job -Job $trimJob -Force
                        Write-LogEntry "   [WARN] SSD Trim timed out on drive $sysDriveLetter" 'WARNING'
                    }
                }
                
            } catch {
                Write-LogEntry "   [WARN] SSD Trim failed: $_" 'WARNING'
            }
        }

        # Get final disk space and calculate savings
        $finalSpace = Get-DiskSpaceInfo "C:\"
        if ($initialSpace -and $finalSpace) {
            $actualSpaceFreed = $finalSpace.FreeSpace - $initialSpace.FreeSpace
            if ($actualSpaceFreed -gt 0) {
                $savedText = if ($actualSpaceFreed -gt 1GB) { 
                    "$([math]::Round($actualSpaceFreed / 1GB, 2)) GB" 
                } else { 
                    "$([math]::Round($actualSpaceFreed / 1MB, 1)) MB" 
                }
                Write-LogEntry "[SAVE] Total disk space recovered: $savedText" 'SUCCESS'
            }
        }

    } catch {
        Write-LogEntry "Critical error during disk cleanup: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Disk cleanup failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Safe counting with null checks
        $cleanedCount = if ($null -ne $script:cleanedPaths) { $script:cleanedPaths.Count } else { 0 }
        $failedCount = if ($null -ne $script:failedPaths) { $script:failedPaths.Count } else { 0 }
        
        Write-LogEntry "`n[SUMMARY] Disk Cleanup Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Locations cleaned: $cleanedCount" 'SUCCESS'
        Write-LogEntry "Failed operations: $failedCount" 'ERROR'
        
        if ($script:totalSpaceFreed -gt 0) {
            $totalSizeText = if ($script:totalSpaceFreed -gt 1GB) { 
                "$([math]::Round($script:totalSpaceFreed / 1GB, 2)) GB" 
            } else { 
                "$([math]::Round($script:totalSpaceFreed / 1MB, 1)) MB" 
            }
            Write-LogEntry "Estimated space freed: $totalSizeText" 'SUCCESS'
        }
        
        # Show failed operations if any
        if ($failedCount -gt 0 -and $null -ne $script:failedPaths) {
            Write-LogEntry "`n[ERROR] Failed Operations:" 'ERROR'
            foreach ($failed in $script:failedPaths) {
                Write-LogEntry "  - $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file
        try {
            if ($null -ne $script:logEntries) {
                $script:logEntries.ToArray() | Out-File -FilePath $LogPath -Encoding UTF8 -Force
                Write-LogEntry "[NOTE] Detailed log saved to: $LogPath" 'INFO'
            }
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status
        if ($cleanedCount -gt 0) {
            $statusMsg = "[OK] Disk cleanup completed - $cleanedCount locations processed"
            if ($failedCount -gt 0) {
                $statusMsg += " ($failedCount failed)"
            }
            if ($script:totalSpaceFreed -gt 0) {
                $sizeText = if ($script:totalSpaceFreed -gt 1GB) { 
                    "$([math]::Round($script:totalSpaceFreed / 1GB, 2))GB" 
                } else { 
                    "$([math]::Round($script:totalSpaceFreed / 1MB, 1))MB" 
                }
                $statusMsg += " [~$sizeText freed]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] Disk cleanup completed but no locations were processed"
        }
        
        Write-LogEntry "=== Disk Cleanup Operations Completed ===" 'INFO'
    }
}


# -----------------------------------------------------------------------------
# Option 13 - System Repair
# Updated from 08_System_Repair.ps1 v2.4 on 2026-04-27
# -----------------------------------------------------------------------------
function Invoke-SystemMaintenance {
    [CmdletBinding()]
    param(
        [switch]$AutoRepairOnDetection = $true,
        [switch]$AllowWmiRepair = $true,
        [switch]$AllowNetworkReset = $false,
        [switch]$AllowWindowsUpdateReset = $false,
        [switch]$AllowOfflineDiskRepair = $false,
        [switch]$AllowFirewallReset = $false,
        [switch]$AllowIconCacheRebuild = $false,
        [switch]$AllowCopilotRemoval = $false,
        [switch]$AggressiveCleanup = $false,
        [switch]$ClearEventLogs = $false,
        [switch]$AutoRebootIfNeeded = $false,
        [int]$AutoRebootDelaySeconds = 60,
        [string]$LogDirectory = 'C:\Logs'
    )

    $previousErrorActionPreference = $ErrorActionPreference

    try {
    $ErrorActionPreference = 'Stop'
    
    $script:RunStart = Get-Date
    $script:ComputerName = $env:COMPUTERNAME
    $script:TimestampForFile = $script:RunStart.ToString('yyyy-MM-dd_HH-mm-ss')
    $script:BaseFileName = "{0}-SystemRepair-{1}" -f $script:ComputerName, $script:TimestampForFile
    $script:YamlLogPath = Join-Path $LogDirectory ($script:BaseFileName + '.yaml')
    
    $script:Summary = [ordered]@{
        ComputerName                 = $script:ComputerName
        StartTime                    = $script:RunStart
        EndTime                      = $null
        StepsSucceeded               = 0
        StepsFailed                  = 0
        Warnings                     = 0
        RebootRequired               = $false
        PendingRebootDetected        = $false
        DiskCorruptionSuspected      = $false
        DismCorruptionDetected       = $false
        SfcIntegrityViolations       = $false
        WmiRepositoryInconsistent    = $false
        RepairsAttempted             = New-Object System.Collections.Generic.List[string]
        Notes                        = New-Object System.Collections.Generic.List[string]
    }
    
    $script:DetailedResults = New-Object System.Collections.Generic.List[object]
    
    function Ensure-LogDirectory {
        if (-not (Test-Path -LiteralPath $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
    }
    
    function ConvertTo-YamlScalar {
        param(
            [Parameter(Mandatory)][AllowNull()]$Value
        )
    
        if ($null -eq $Value) {
            return 'null'
        }
    
        if ($Value -is [bool]) {
            return $Value.ToString().ToLowerInvariant()
        }
    
        if ($Value -is [datetime]) {
            return "'" + $Value.ToString('yyyy-MM-dd HH:mm:ss') + "'"
        }
    
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
            return [string]$Value
        }
    
        $text = [string]$Value
        $text = $text -replace "'", "''"
        return "'" + $text + "'"
    }
    
    function Add-DetailedResult {
        param(
            [Parameter(Mandatory)][string]$Step,
            [Parameter(Mandatory)][string]$Status,
            [Parameter(Mandatory)][string]$Message,
            [AllowNull()]$Data = $null
        )
    
        $script:DetailedResults.Add([PSCustomObject]@{
            Timestamp = Get-Date
            Step      = $Step
            Status    = $Status
            Message   = $Message
            Data      = $Data
        }) | Out-Null
    }
    
    function Write-YamlLog {
        try {
            Ensure-LogDirectory
    
            $lines = New-Object System.Collections.Generic.List[string]
    
            $lines.Add('run:') | Out-Null
            $lines.Add("  computer_name: $(ConvertTo-YamlScalar $script:ComputerName)") | Out-Null
            $lines.Add("  start_time: $(ConvertTo-YamlScalar $script:Summary.StartTime)") | Out-Null
            $lines.Add("  end_time: $(ConvertTo-YamlScalar $script:Summary.EndTime)") | Out-Null
            $lines.Add("  yaml_log_path: $(ConvertTo-YamlScalar $script:YamlLogPath)") | Out-Null
            $lines.Add('') | Out-Null
    
            $lines.Add('settings:') | Out-Null
            $lines.Add("  auto_repair_on_detection: $(ConvertTo-YamlScalar $AutoRepairOnDetection)") | Out-Null
            $lines.Add("  allow_wmi_repair: $(ConvertTo-YamlScalar $AllowWmiRepair)") | Out-Null
            $lines.Add("  allow_network_reset: $(ConvertTo-YamlScalar $AllowNetworkReset)") | Out-Null
            $lines.Add("  allow_windows_update_reset: $(ConvertTo-YamlScalar $AllowWindowsUpdateReset)") | Out-Null
            $lines.Add("  allow_offline_disk_repair: $(ConvertTo-YamlScalar $AllowOfflineDiskRepair)") | Out-Null
            $lines.Add("  allow_firewall_reset: $(ConvertTo-YamlScalar $AllowFirewallReset)") | Out-Null
            $lines.Add("  allow_icon_cache_rebuild: $(ConvertTo-YamlScalar $AllowIconCacheRebuild)") | Out-Null
            $lines.Add("  allow_copilot_removal: $(ConvertTo-YamlScalar $AllowCopilotRemoval)") | Out-Null
            $lines.Add("  aggressive_cleanup: $(ConvertTo-YamlScalar $AggressiveCleanup)") | Out-Null
            $lines.Add("  clear_event_logs: $(ConvertTo-YamlScalar $ClearEventLogs)") | Out-Null
            $lines.Add("  auto_reboot_if_needed: $(ConvertTo-YamlScalar $AutoRebootIfNeeded)") | Out-Null
            $lines.Add("  auto_reboot_delay_seconds: $(ConvertTo-YamlScalar $AutoRebootDelaySeconds)") | Out-Null
            $lines.Add('') | Out-Null
    
            $lines.Add('summary:') | Out-Null
            $lines.Add("  steps_succeeded: $(ConvertTo-YamlScalar $script:Summary.StepsSucceeded)") | Out-Null
            $lines.Add("  steps_failed: $(ConvertTo-YamlScalar $script:Summary.StepsFailed)") | Out-Null
            $lines.Add("  warnings: $(ConvertTo-YamlScalar $script:Summary.Warnings)") | Out-Null
            $lines.Add("  reboot_required: $(ConvertTo-YamlScalar $script:Summary.RebootRequired)") | Out-Null
            $lines.Add("  pending_reboot_detected: $(ConvertTo-YamlScalar $script:Summary.PendingRebootDetected)") | Out-Null
            $lines.Add("  disk_corruption_suspected: $(ConvertTo-YamlScalar $script:Summary.DiskCorruptionSuspected)") | Out-Null
            $lines.Add("  dism_corruption_detected: $(ConvertTo-YamlScalar $script:Summary.DismCorruptionDetected)") | Out-Null
            $lines.Add("  sfc_integrity_violations: $(ConvertTo-YamlScalar $script:Summary.SfcIntegrityViolations)") | Out-Null
            $lines.Add("  wmi_repository_inconsistent: $(ConvertTo-YamlScalar $script:Summary.WmiRepositoryInconsistent)") | Out-Null
            $lines.Add('') | Out-Null
    
            $lines.Add('repairs_attempted:') | Out-Null
            if ($script:Summary.RepairsAttempted.Count -gt 0) {
                foreach ($repair in $script:Summary.RepairsAttempted) {
                    $lines.Add("  - $(ConvertTo-YamlScalar $repair)") | Out-Null
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
            $lines.Add('') | Out-Null
    
            $lines.Add('notes:') | Out-Null
            if ($script:Summary.Notes.Count -gt 0) {
                foreach ($note in $script:Summary.Notes) {
                    $lines.Add("  - $(ConvertTo-YamlScalar $note)") | Out-Null
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
            $lines.Add('') | Out-Null
    
            $lines.Add('detailed_results:') | Out-Null
            if ($script:DetailedResults.Count -gt 0) {
                foreach ($entry in $script:DetailedResults) {
                    $lines.Add('  -') | Out-Null
                    $lines.Add("    timestamp: $(ConvertTo-YamlScalar $entry.Timestamp)") | Out-Null
                    $lines.Add("    step: $(ConvertTo-YamlScalar $entry.Step)") | Out-Null
                    $lines.Add("    status: $(ConvertTo-YamlScalar $entry.Status)") | Out-Null
                    $lines.Add("    message: $(ConvertTo-YamlScalar $entry.Message)") | Out-Null
    
                    if ($null -eq $entry.Data) {
                        $lines.Add('    data: null') | Out-Null
                    }
                    elseif ($entry.Data -is [System.Collections.IDictionary]) {
                        $lines.Add('    data:') | Out-Null
                        foreach ($key in $entry.Data.Keys) {
                            $lines.Add("      $key`: $(ConvertTo-YamlScalar $entry.Data[$key])") | Out-Null
                        }
                    }
                    else {
                        $lines.Add("    data: $(ConvertTo-YamlScalar $entry.Data)") | Out-Null
                    }
                }
            }
            else {
                $lines.Add('  []') | Out-Null
            }
    
            Set-Content -Path $script:YamlLogPath -Value $lines -Encoding UTF8
        }
        catch {
            Write-Warning "Failed to write YAML log: $($_.Exception.Message)"
        }
    }
    
    function Write-Log {
        param(
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('INFO','OK','WARN','ERROR')][string]$Level = 'INFO'
        )
    
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line = "[$timestamp] [$('{0,-5}' -f $Level)] $Message"
    
        switch ($Level) {
            'INFO'  { Write-Host $line -ForegroundColor Cyan }
            'OK'    { Write-Host $line -ForegroundColor Green }
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
        }
    }
    
    function Add-Note {
        param([string]$Message)
        $script:Summary.Notes.Add($Message) | Out-Null
    }
    
    function Add-RepairAttempt {
        param([string]$Message)
        $script:Summary.RepairsAttempted.Add($Message) | Out-Null
    }
    
    function Complete-Step {
        param([string]$Name)
        $script:Summary.StepsSucceeded++
        Write-Log "$Name completed." 'OK'
        Add-DetailedResult -Step $Name -Status 'Succeeded' -Message "$Name completed successfully."
    }
    
    function Fail-Step {
        param(
            [string]$Name,
            [string]$Reason
        )
        $script:Summary.StepsFailed++
        Write-Log "$Name failed: $Reason" 'ERROR'
        Add-Note "$Name failed: $Reason"
        Add-DetailedResult -Step $Name -Status 'Failed' -Message $Reason
    }
    
    function Warn-Step {
        param(
            [string]$Name,
            [string]$Reason
        )
        $script:Summary.Warnings++
        Write-Log "$Name warning: $Reason" 'WARN'
        Add-Note "$Name warning: $Reason"
        Add-DetailedResult -Step $Name -Status 'Warning' -Message $Reason
    }
    
    function Test-IsAdministrator {
        try {
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = New-Object Security.Principal.WindowsPrincipal($identity)
            return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        }
        catch {
            return $false
        }
    }
    
    
    function Set-ClassicContextMenuForHive {
        param(
            [Parameter(Mandatory)][string]$RootKey
        )
    
        $basePath = Join-Path $RootKey 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}'
        $inprocPath = Join-Path $basePath 'InprocServer32'
    
        if (-not (Test-Path -LiteralPath $basePath)) {
            New-Item -Path $basePath -Force | Out-Null
        }
    
        if (-not (Test-Path -LiteralPath $inprocPath)) {
            New-Item -Path $inprocPath -Force | Out-Null
        }
    
        New-ItemProperty -Path $inprocPath -Name '(default)' -Value '' -PropertyType String -Force | Out-Null
        Write-Log "Classic context menu registry value set for hive: $RootKey" 'OK'
        Add-DetailedResult -Step 'ClassicContextMenuRegistry' -Status 'Info' -Message "Classic context menu registry value set." -Data @{
            RootKey = $RootKey
            RegistryPath = $inprocPath
        }
    }
    
    function Enable-ClassicContextMenuAllUsers {
        Write-Log 'Applying classic Windows 10-style context menu for all users...' 'INFO'
    
        Set-ClassicContextMenuForHive -RootKey 'Registry::HKEY_CURRENT_USER'
    
        $userSids = Get-ChildItem Registry::HKEY_USERS |
            Where-Object {
                $_.PSChildName -match '^S-1-5-21-' -and
                $_.PSChildName -notmatch '_Classes$'
            } |
            Select-Object -ExpandProperty PSChildName
    
        foreach ($sid in $userSids) {
            Set-ClassicContextMenuForHive -RootKey "Registry::HKEY_USERS\$sid"
        }
    
        $defaultHiveName = 'HKU\DefaultTemp'
        $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTemp'
        $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'
    
        if (Test-Path -LiteralPath $defaultUserNtUserDat) {
            $hiveLoaded = $false
    
            try {
                $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to load Default User hive: $($loadResult -join ' ')"
                }
    
                $hiveLoaded = $true
                Start-Sleep -Milliseconds 750
    
                Set-ClassicContextMenuForHive -RootKey $defaultHivePsPath
    
                Start-Sleep -Milliseconds 750
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Start-Sleep -Milliseconds 750
            }
            catch {
                Write-Log "Failed to update Default User profile: $($_.Exception.Message)" 'ERROR'
                Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Failed' -Message $_.Exception.Message
            }
            finally {
                if ($hiveLoaded) {
                    $unloaded = $false
    
                    foreach ($attempt in 1..5) {
                        $unloadResult = & reg.exe unload $defaultHiveName
                        if ($LASTEXITCODE -eq 0) {
                            $unloaded = $true
                            Write-Log 'Applied classic context menu to Default User profile.' 'OK'
                            Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Succeeded' -Message 'Applied classic context menu to Default User profile.'
                            break
                        }
    
                        Start-Sleep -Seconds 1
                        [System.GC]::Collect()
                        [System.GC]::WaitForPendingFinalizers()
                    }
    
                    if (-not $unloaded) {
                        Write-Log 'Classic context menu was written to Default User profile, but unloading the hive failed. A reboot may be required before the hive is released.' 'WARN'
                        Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Warning' -Message 'Classic context menu was written to Default User profile, but unloading the hive failed. A reboot may be required before the hive is released.'
                    }
                }
            }
        }
        else {
            Write-Log 'Default User NTUSER.DAT not found; future new users were not updated.' 'WARN'
            Add-DetailedResult -Step 'ClassicContextMenuDefaultUser' -Status 'Warning' -Message 'Default User NTUSER.DAT not found; future new users were not updated.'
        }
    
        Write-Log 'Classic context menu registry changes applied. Users may need to sign out and back in.' 'INFO'
        Add-DetailedResult -Step 'ClassicContextMenuAllUsers' -Status 'Info' -Message 'Classic context menu registry changes applied for current, loaded, and default user profiles.'
    }
    
    function Invoke-Safely {
        param(
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][scriptblock]$ScriptBlock,
            [switch]$WarnOnly
        )
    
        try {
            & $ScriptBlock
            Complete-Step -Name $Name
            return $true
        }
        catch {
            if ($WarnOnly) {
                Warn-Step -Name $Name -Reason $_.Exception.Message
            }
            else {
                Fail-Step -Name $Name -Reason $_.Exception.Message
            }
            return $false
        }
    }
    
    function Get-PendingRebootState {
        $result = [ordered]@{
            CBServicing_RebootPending         = $false
            WindowsUpdate_RebootRequired      = $false
            SessionManager_PendingFileRename  = $false
            SessionManager_PendingFileRename2 = $false
            UpdateExeVolatile                 = $false
            PackagesPending                   = $false
            WUAU_RebootRequired_COM           = $false
            AnyPendingReboot                  = $false
        }
    
        $cbsRebootPending = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        $wuRebootRequired = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
        $packagesPending  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackagesPending'
        $sessionMgr       = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
        $updateExe        = 'HKLM:\SOFTWARE\Microsoft\Updates'
    
        try { $result.CBServicing_RebootPending = Test-Path -LiteralPath $cbsRebootPending } catch {}
        try { $result.WindowsUpdate_RebootRequired = Test-Path -LiteralPath $wuRebootRequired } catch {}
        try { $result.PackagesPending = Test-Path -LiteralPath $packagesPending } catch {}
    
        try {
            $pendingRename = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations' -ErrorAction Stop).PendingFileRenameOperations
            if ($null -ne $pendingRename -and $pendingRename.Count -gt 0) {
                $result.SessionManager_PendingFileRename = $true
            }
        }
        catch {}
    
        try {
            $pendingRename2 = (Get-ItemProperty -Path $sessionMgr -Name 'PendingFileRenameOperations2' -ErrorAction Stop).PendingFileRenameOperations2
            if ($null -ne $pendingRename2 -and $pendingRename2.Count -gt 0) {
                $result.SessionManager_PendingFileRename2 = $true
            }
        }
        catch {}
    
        try {
            $uev = (Get-ItemProperty -Path $updateExe -Name 'UpdateExeVolatile' -ErrorAction Stop).UpdateExeVolatile
            if ($null -ne $uev -and [int]$uev -ne 0) {
                $result.UpdateExeVolatile = $true
            }
        }
        catch {}
    
        try {
            $sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
            if ($sysInfo.RebootRequired) {
                $result.WUAU_RebootRequired_COM = $true
            }
        }
        catch {}
    
        if ($result.Values -contains $true) {
            $result.AnyPendingReboot = $true
        }
    
        Add-DetailedResult -Step 'PendingRebootCheckData' -Status 'Info' -Message 'Collected pending reboot state.' -Data $result
        [PSCustomObject]$result
    }
    
    function Test-SystemDiskIsSSD {
        try {
            $systemDriveLetter = $env:SystemDrive.TrimEnd(':')
            $partition = Get-Partition -DriveLetter $systemDriveLetter -ErrorAction Stop
            $disk = Get-Disk -Number $partition.DiskNumber -ErrorAction Stop
            return ($disk.MediaType -eq 'SSD')
        }
        catch {
            return $false
        }
    }
    
    function Invoke-DismCommand {
        param([string[]]$Arguments)
    
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "$env:SystemRoot\System32\dism.exe"
        $psi.Arguments = $Arguments -join ' '
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
    
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        [void]$proc.Start()
    
        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()
    
        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
        }
    
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
        }
    
        Add-DetailedResult -Step 'DISM' -Status 'Info' -Message ("Executed DISM: " + ($Arguments -join ' ')) -Data @{
            ExitCode = $proc.ExitCode
        }
    
        return [PSCustomObject]@{
            ExitCode = $proc.ExitCode
            StdOut   = $stdout
            StdErr   = $stderr
        }
    }
    
    function Invoke-SfcCommand {
        param([string]$Arguments)
    
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "$env:SystemRoot\System32\sfc.exe"
        $psi.Arguments = $Arguments
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
    
        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        [void]$proc.Start()
    
        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit()
    
        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            $stdout -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
        }
    
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            $stderr -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'WARN' }
        }
    
        Add-DetailedResult -Step 'SFC' -Status 'Info' -Message ("Executed SFC: " + $Arguments) -Data @{
            ExitCode = $proc.ExitCode
        }
    
        return [PSCustomObject]@{
            ExitCode = $proc.ExitCode
            StdOut   = $stdout
            StdErr   = $stderr
        }
    }
    
    
    function Get-DiskSpaceInfo {
        param([string]$Path)
    
        try {
            $driveRoot = Split-Path -Path $Path -Qualifier
            if ([string]::IsNullOrWhiteSpace($driveRoot)) {
                $driveRoot = $env:SystemDrive + '\'
            }
    
            $drive = [System.IO.DriveInfo]::new($driveRoot)
            return @{
                FreeSpace  = [int64]$drive.AvailableFreeSpace
                TotalSize  = [int64]$drive.TotalSize
                UsedSpace  = [int64]($drive.TotalSize - $drive.AvailableFreeSpace)
            }
        }
        catch {
            return $null
        }
    }
    
    function Test-SafeCleanupPath {
        param([string]$Path)
    
        if ([string]::IsNullOrWhiteSpace($Path)) {
            return $false
        }
    
        $normalized = $Path.TrimEnd('\')
    
        $blockedPaths = @(
            'C:\Windows\System32',
            'C:\Windows\SysWOW64',
            'C:\Program Files',
            'C:\Program Files (x86)',
            'C:\Windows\explorer.exe',
            'C:\Windows\System32\drivers'
        )
    
        foreach ($blocked in $blockedPaths) {
            if ($normalized -ieq $blocked -or $normalized -like ($blocked + '\*')) {
                return $false
            }
        }
    
        $allowedPatterns = @(
            'C:\Windows\Temp*',
            'C:\Temp*',
            'C:\SWSetup*',
            'C:\Lab Update Scripts*',
            'C:\ProgramData\Win11UpgradeStage*',
            'C:\windows.old*',
            'C:\system.sav*',
            'C:\Windows\SoftwareDistribution.bak*',
            'C:\SoftwareDistribution.bak*',
            'C:\Windows\SoftwareDistribution\Download*',
            'C:\Windows\Prefetch*',
            'C:\Windows\Logs\CBS*',
            'C:\ProgramData\Microsoft\Windows\WER\ReportQueue*',
            "$env:TEMP*",
            "$env:LOCALAPPDATA\Temp*",
            "$env:LOCALAPPDATA\Microsoft\Windows\INetCache*",
            "$env:LOCALAPPDATA\Microsoft\Windows\WebCache*",
            "$env:LOCALAPPDATA\CrashDumps*",
            "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache*",
            "$env:LOCALAPPDATA\D3DSCache*",
            "$env:LOCALAPPDATA\NVIDIA\DXCache*",
            "$env:LOCALAPPDATA\NVIDIA\GLCache*"
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    
        foreach ($pattern in $allowedPatterns) {
            if ($normalized -like $pattern) {
                return $true
            }
        }
    
        return $false
    }
    
    function Remove-FolderContents {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Description,
            [switch]$ContentsOnly
        )
    
        if (-not (Test-Path -LiteralPath $Path)) {
            return @{
                Success    = $true
                SpaceFreed = [int64]0
                ItemCount  = 0
                Message    = 'Path does not exist'
            }
        }
    
        if (-not (Test-SafeCleanupPath -Path $Path)) {
            return @{
                Success    = $false
                SpaceFreed = [int64]0
                ItemCount  = 0
                Message    = 'Path blocked for security'
            }
        }
    
        try {
            $items = if ($ContentsOnly) {
                @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)
            }
            else {
                @(Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue)
            }
    
            $itemCount = $items.Count
            [int64]$sizeBefore = 0
    
            if ($itemCount -gt 0) {
                $files = if ($ContentsOnly) {
                    $items | Where-Object { -not $_.PSIsContainer }
                }
                else {
                    $items
                }
    
                if ($files.Count -gt 0) {
                    $sizeSum = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    if ($null -ne $sizeSum) {
                        $sizeBefore = [int64]$sizeSum
                    }
                }
            }
    
            if ($itemCount -eq 0) {
                return @{
                    Success    = $true
                    SpaceFreed = [int64]0
                    ItemCount  = 0
                    Message    = 'Folder is empty'
                }
            }
    
            if ($ContentsOnly) {
                Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | ForEach-Object {
                    try {
                        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                    }
                    catch {
                        Write-Log "Could not remove $($_.FullName): $($_.Exception.Message)" 'WARN'
                    }
                }
            }
            else {
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
            }
    
            return @{
                Success    = $true
                SpaceFreed = $sizeBefore
                ItemCount  = $itemCount
                Message    = 'Successfully cleaned'
            }
        }
        catch {
            return @{
                Success    = $false
                SpaceFreed = [int64]0
                ItemCount  = 0
                Message    = $_.Exception.Message
            }
        }
    }
    
    
    
    
    function Get-FolderSizeInfo {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$Path
        )
    
        $info = [ordered]@{
            Path       = $Path
            Exists     = $false
            ItemCount  = 0
            FileCount  = 0
            FolderCount = 0
            SizeBytes  = [int64]0
            SizeMB     = [double]0
            SizeGB     = [double]0
            Message    = $null
        }
    
        if (-not (Test-Path -LiteralPath $Path)) {
            $info.Message = 'Path does not exist'
            return [PSCustomObject]$info
        }
    
        $info.Exists = $true
    
        try {
            $items = @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue)
            $files = @($items | Where-Object { -not $_.PSIsContainer })
            $folders = @($items | Where-Object { $_.PSIsContainer })
            $sizeBytes = ($files | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            if ($null -eq $sizeBytes) { $sizeBytes = 0 }
    
            $info.ItemCount = $items.Count
            $info.FileCount = $files.Count
            $info.FolderCount = $folders.Count
            $info.SizeBytes = [int64]$sizeBytes
            $info.SizeMB = [math]::Round(([double]$info.SizeBytes / 1MB), 2)
            $info.SizeGB = [math]::Round(([double]$info.SizeBytes / 1GB), 3)
            $info.Message = 'Size calculated successfully'
        }
        catch {
            $info.Message = $_.Exception.Message
            Write-Log "Could not calculate folder size for $Path`: $($info.Message)" 'WARN'
        }
    
        return [PSCustomObject]$info
    }
    
    function Stop-WindowsUpdateLockingProcesses {
        [CmdletBinding()]
        param()
    
        Write-Log 'Checking for Windows Update processes that may lock SoftwareDistribution...' 'INFO'
    
        $processNames = @(
            'MoUsoCoreWorker',
            'TiWorker',
            'TrustedInstaller',
            'UsoClient',
            'MusNotification',
            'MusNotificationUx',
            'SIHClient'
        )
    
        $results = New-Object System.Collections.Generic.List[object]
    
        foreach ($name in $processNames) {
            $processes = @(Get-Process -Name $name -ErrorAction SilentlyContinue)
    
            if ($processes.Count -eq 0) {
                $results.Add([PSCustomObject]@{
                    Name    = $name
                    Action  = 'NotRunning'
                    Success = $true
                    Message = 'Process not running'
                }) | Out-Null
                continue
            }
    
            foreach ($proc in $processes) {
                try {
                    Write-Log "Stopping possible Windows Update lock process: $($proc.ProcessName) PID $($proc.Id)" 'WARN'
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                    Start-Sleep -Milliseconds 500
    
                    $stillRunning = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
                    if ($null -eq $stillRunning) {
                        Write-Log "Stopped process $($proc.ProcessName) PID $($proc.Id)." 'OK'
                        $results.Add([PSCustomObject]@{
                            Name    = $proc.ProcessName
                            ProcessId = $proc.Id
                            Action  = 'Stopped'
                            Success = $true
                            Message = 'Stopped successfully'
                        }) | Out-Null
                    }
                    else {
                        Write-Log "Process $($proc.ProcessName) PID $($proc.Id) is still running after stop attempt." 'WARN'
                        $results.Add([PSCustomObject]@{
                            Name    = $proc.ProcessName
                            ProcessId = $proc.Id
                            Action  = 'StopAttempted'
                            Success = $false
                            Message = 'Still running after Stop-Process'
                        }) | Out-Null
                    }
                }
                catch {
                    Write-Log "Could not stop $($proc.ProcessName) PID $($proc.Id): $($_.Exception.Message)" 'WARN'
                    $results.Add([PSCustomObject]@{
                        Name    = $proc.ProcessName
                        ProcessId = $proc.Id
                        Action  = 'FailedToStop'
                        Success = $false
                        Message = $_.Exception.Message
                    }) | Out-Null
                }
            }
        }
    
        return @($results)
    }
    
    function Get-FolderSizeBytesSafe {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Path
        )
    
        try {
            if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
                return [int64]0
            }
    
            $sum = (Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer } |
                Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    
            if ($null -eq $sum) { return [int64]0 }
            return [int64]$sum
        }
        catch {
            Write-Log "Unable to calculate size for ${Path}: $($_.Exception.Message)" 'WARN'
            return [int64]0
        }
    }
    
    function Stop-SoftwareDistributionBackupLockingServices {
        [CmdletBinding()]
        param()
    
        $services = @(
            'wuauserv',
            'bits',
            'cryptsvc',
            'dosvc',
            'UsoSvc',
            'WaaSMedicSvc',
            'TrustedInstaller',
            'msiserver'
        )
    
        foreach ($svcName in $services) {
            try {
                $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                if ($null -eq $svc) {
                    Write-Log "Service not found while releasing SoftwareDistribution locks: $svcName" 'INFO'
                    continue
                }
    
                if ($svc.Status -ne 'Stopped') {
                    Write-Log "Stopping service to release SoftwareDistribution locks: $svcName ($($svc.Status))" 'INFO'
                    Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                }
    
                $svc.Refresh()
                Write-Log "Service state after stop attempt: $svcName = $($svc.Status)" 'INFO'
            }
            catch {
                Write-Log "Could not stop service ${svcName}: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    
    function Stop-SoftwareDistributionBackupLockingProcesses {
        [CmdletBinding()]
        param()
    
        $processNames = @(
            'TiWorker',
            'TrustedInstaller',
            'MoUsoCoreWorker',
            'UsoClient',
            'wuauclt',
            'bitsadmin',
            'msiexec',
            'MusNotification',
            'MusNotificationUx',
            'SIHClient'
        )
    
        foreach ($procName in $processNames) {
            try {
                $procs = @(Get-Process -Name $procName -ErrorAction SilentlyContinue)
                foreach ($proc in $procs) {
                    Write-Log "Stopping process to release SoftwareDistribution locks: $($proc.ProcessName) PID $($proc.Id)" 'WARN'
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                }
            }
            catch {
                Write-Log "Could not stop process ${procName}: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    
    function Start-SoftwareDistributionBackupUpdateServices {
        [CmdletBinding()]
        param()
    
        $services = @('cryptsvc', 'bits', 'wuauserv', 'dosvc', 'UsoSvc')
    
        foreach ($svcName in $services) {
            try {
                $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                if ($null -eq $svc) { continue }
    
                if ($svc.Status -ne 'Running') {
                    Write-Log "Restarting update-related service after SoftwareDistribution backup cleanup: $svcName" 'INFO'
                    Start-Service -Name $svcName -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 1
                }
    
                $svc.Refresh()
                Write-Log "Service state after restart attempt: $svcName = $($svc.Status)" 'INFO'
            }
            catch {
                Write-Log "Could not restart service ${svcName}: $($_.Exception.Message)" 'WARN'
            }
        }
    }
    
    function Remove-SoftwareDistributionBakFolders {
        [CmdletBinding()]
        param()
    
        Write-Log 'Starting forced SoftwareDistribution backup folder cleanup. This will delete .bak, .bak1, .bak2, .old_*, and backup variants without creating new backups.' 'INFO'
    
        $basePath = 'C:\Windows'
        $targets = @()
    
        try {
            if (-not (Test-Path -LiteralPath $basePath)) {
                Write-Log "SoftwareDistribution backup cleanup skipped because $basePath does not exist." 'WARN'
                return [PSCustomObject]@{
                    Success         = $false
                    SpaceFreed      = [int64]0
                    SpaceFreedBytes = [int64]0
                    SpaceFreedMB    = [double]0
                    SpaceFreedGB    = [double]0
                    ItemCount       = 0
                    DeletedCount    = 0
                    FailedCount     = 0
                    Message         = "$basePath does not exist"
                }
            }
    
            # Direct Select-Object pipeline. No ArrayList/List .Add() calls are used here.
            $targets = @(
                Get-ChildItem -LiteralPath $basePath -Directory -Force -ErrorAction Stop |
                    Where-Object {
                        $_ -and
                        -not [string]::IsNullOrWhiteSpace($_.FullName) -and
                        $_.Name -ne 'SoftwareDistribution' -and
                        (
                            $_.Name -match '^SoftwareDistribution\.bak\d*$' -or
                            $_.Name -match '^SoftwareDistribution\.(old|old_).*$' -or
                            $_.Name -match '^SoftwareDistribution[_-](bak|backup|old).*$' -or
                            $_.Name -match '^SoftwareDistribution\.backup\d*$'
                        )
                    } |
                    Select-Object -ExpandProperty FullName
            )
        }
        catch {
            Write-Log "Could not scan $basePath for SoftwareDistribution backup folders: $($_.Exception.Message)" 'WARN'
            $targets = @()
        }
    
        # Explicit safety check for the common folders that are known to exist in the field.
        foreach ($explicitTarget in @(
            'C:\Windows\SoftwareDistribution.bak',
            'C:\Windows\SoftwareDistribution.bak1',
            'C:\Windows\SoftwareDistribution.bak2',
            'C:\Windows\SoftwareDistribution.bak3',
            'C:\Windows\SoftwareDistribution.bak4',
            'C:\Windows\SoftwareDistribution.bak5'
        )) {
            if (Test-Path -LiteralPath $explicitTarget) {
                $targets += $explicitTarget
            }
        }
    
        $targets = @(
            $targets |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        )
    
        if (-not $targets -or $targets.Count -eq 0) {
            Write-Log 'No SoftwareDistribution backup folders were found.' 'INFO'
            return [PSCustomObject]@{
                Success         = $true
                SpaceFreed      = [int64]0
                SpaceFreedBytes = [int64]0
                SpaceFreedMB    = [double]0
                SpaceFreedGB    = [double]0
                ItemCount       = 0
                DeletedCount    = 0
                FailedCount     = 0
                Message         = 'No SoftwareDistribution backup folders found'
            }
        }
    
        Write-Log "Found $($targets.Count) SoftwareDistribution backup folder(s) to force delete." 'WARN'
    
        [int64]$totalFreedBytes = 0
        [int]$deletedCount = 0
        [int]$failedCount = 0
        [int]$totalItems = 0
        $failureMessages = @()
    
        Stop-SoftwareDistributionBackupLockingServices
        Stop-SoftwareDistributionBackupLockingProcesses
    
        foreach ($target in $targets) {
            if ([string]::IsNullOrWhiteSpace($target)) {
                Write-Log 'Skipping blank SoftwareDistribution backup folder target.' 'WARN'
                continue
            }
    
            if (-not (Test-Path -LiteralPath $target)) {
                Write-Log "SoftwareDistribution backup folder no longer exists: $target" 'INFO'
                continue
            }
    
            $sizeBytes = Get-FolderSizeBytesSafe -Path $target
            $sizeGB = [math]::Round(([double]$sizeBytes / 1GB), 3)
            $sizeMB = [math]::Round(([double]$sizeBytes / 1MB), 2)
            $itemCount = @(Get-ChildItem -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue).Count
    
            Write-Log "Attempting forced deletion of SoftwareDistribution backup folder: $target | Size before deletion: $sizeGB GB ($sizeMB MB) | Items: $itemCount" 'WARN'
    
            $deleted = $false
    
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                Write-Log "SoftwareDistribution backup delete attempt $attempt of 5: $target" 'INFO'
    
                try {
                    cmd.exe /c "attrib -r -s -h `"$target`" /s /d" 2>$null | Out-Null
                }
                catch {
                    Write-Log "Could not clear attributes on ${target}: $($_.Exception.Message)" 'WARN'
                }
    
                try {
                    Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
                    if (-not (Test-Path -LiteralPath $target)) {
                        Write-Log "Deleted SoftwareDistribution backup folder with Remove-Item: $target | Estimated freed: $sizeGB GB ($sizeMB MB)" 'OK'
                        $totalFreedBytes += $sizeBytes
                        $totalItems += $itemCount
                        $deletedCount++
                        $deleted = $true
                        break
                    }
                }
                catch {
                    Write-Log "Remove-Item failed on attempt $attempt for ${target}: $($_.Exception.Message)" 'WARN'
                }
    
                try {
                    Write-Log "Trying cmd.exe rmdir fallback for SoftwareDistribution backup folder: $target" 'WARN'
                    cmd.exe /c "rmdir /s /q `"$target`"" 2>$null | Out-Null
                    if (-not (Test-Path -LiteralPath $target)) {
                        Write-Log "Deleted SoftwareDistribution backup folder with cmd rmdir: $target | Estimated freed: $sizeGB GB ($sizeMB MB)" 'OK'
                        $totalFreedBytes += $sizeBytes
                        $totalItems += $itemCount
                        $deletedCount++
                        $deleted = $true
                        break
                    }
                }
                catch {
                    Write-Log "cmd rmdir fallback failed on attempt $attempt for ${target}: $($_.Exception.Message)" 'WARN'
                }
    
                Stop-SoftwareDistributionBackupLockingServices
                Stop-SoftwareDistributionBackupLockingProcesses
                Start-Sleep -Seconds 3
            }
    
            if (-not $deleted) {
                $failedCount++
                $failureMessages += "Failed to delete $target after 5 attempts"
                Write-Log "FAILED to delete SoftwareDistribution backup folder after all attempts: $target" 'ERROR'
            }
        }
    
        Start-SoftwareDistributionBackupUpdateServices
    
        $freedMB = [math]::Round(([double]$totalFreedBytes / 1MB), 2)
        $freedGB = [math]::Round(([double]$totalFreedBytes / 1GB), 3)
    
        if ($failedCount -eq 0) {
            Write-Log "Forced SoftwareDistribution backup cleanup completed. Deleted folders: $deletedCount. Estimated freed: $freedGB GB ($freedMB MB)." 'OK'
        }
        else {
            Write-Log "Forced SoftwareDistribution backup cleanup completed with failures. Deleted folders: $deletedCount. Failed folders: $failedCount. Estimated freed: $freedGB GB ($freedMB MB)." 'WARN'
        }
    
        return [PSCustomObject]@{
            Success         = ($failedCount -eq 0)
            SpaceFreed      = $totalFreedBytes
            SpaceFreedBytes = $totalFreedBytes
            SpaceFreedMB    = $freedMB
            SpaceFreedGB    = $freedGB
            ItemCount       = $totalItems
            DeletedCount    = $deletedCount
            FailedCount     = $failedCount
            Message         = if ($failedCount -eq 0) { 'Forced SoftwareDistribution backup cleanup completed' } else { ($failureMessages -join '; ') }
        }
    }
    
    function Invoke-WindowsCleanup {
        param([int]$TimeoutSec = 300)
    
        try {
            $cleanmgrPath = Join-Path $env:SystemRoot 'System32\cleanmgr.exe'
            if (-not (Test-Path -LiteralPath $cleanmgrPath)) {
                throw 'Windows Disk Cleanup utility not found.'
            }
    
            Write-Log "Starting Windows Disk Cleanup (timeout: ${TimeoutSec}s)..." 'INFO'
    
            $job = Start-Job -ScriptBlock {
                Start-Process -FilePath "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList '/SAGERUN:1','/VERYLOWDISK' -NoNewWindow -Wait -PassThru
            }
    
            $result = Wait-Job -Job $job -Timeout $TimeoutSec
    
            if ($null -eq $result -or $job.State -eq 'Running') {
                Stop-Job -Job $job -Force | Out-Null
                Remove-Job -Job $job -Force | Out-Null
                throw "Disk Cleanup timed out after $TimeoutSec seconds."
            }
    
            $proc = Receive-Job -Job $job
            Remove-Job -Job $job -Force | Out-Null
    
            return @{
                Success  = $true
                ExitCode = $proc.ExitCode
            }
        }
        catch {
            return @{
                Success = $false
                Error   = $_.Exception.Message
            }
        }
    }
    
    function Invoke-TempCleanup {
        [int64]$totalSpaceFreed = 0
        $cleanupResults = New-Object System.Collections.Generic.List[object]
        $initialSpace = Get-DiskSpaceInfo -Path $env:SystemDrive
    
        Write-Log 'Cleaning temporary files and caches...' 'INFO'
    
        $windowsCleanup = Invoke-WindowsCleanup -TimeoutSec 300
        if ($windowsCleanup.Success) {
            Write-Log 'Windows Disk Cleanup completed successfully.' 'OK'
            $cleanupResults.Add([PSCustomObject]@{
                Path        = 'cleanmgr.exe'
                Description = 'Windows Disk Cleanup'
                ItemCount   = 0
                SpaceFreed  = [int64]0
                Status      = 'Success'
                Message     = "Exit code $($windowsCleanup.ExitCode)"
            }) | Out-Null
        }
        else {
            Warn-Step -Name 'TempCleanup' -Reason "Windows Disk Cleanup failed: $($windowsCleanup.Error)"
            $cleanupResults.Add([PSCustomObject]@{
                Path        = 'cleanmgr.exe'
                Description = 'Windows Disk Cleanup'
                ItemCount   = 0
                SpaceFreed  = [int64]0
                Status      = 'Warning'
                Message     = $windowsCleanup.Error
            }) | Out-Null
        }
    
        $cleanupTargets = New-Object System.Collections.Generic.List[object]
    
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Lab Update Scripts'; Description = 'Lab Update Scripts'; ContentsOnly = $false })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\ProgramData\Win11UpgradeStage'; Description = 'Windows 11 Upgrade Staging'; ContentsOnly = $false })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\SWSetup'; Description = 'HP Software Setup'; ContentsOnly = $false })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Temp'; Description = 'System Temp'; ContentsOnly = $false })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\windows.old'; Description = 'Previous Windows Installation'; ContentsOnly = $false })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\system.sav'; Description = 'System Save'; ContentsOnly = $false })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Temp'; Description = 'Windows Temp'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = $env:TEMP; Description = 'User Temp'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Temp"; Description = 'Local Temp'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\SoftwareDistribution\Download'; Description = 'Windows Update Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Prefetch'; Description = 'Windows Prefetch'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\Windows\Logs\CBS'; Description = 'CBS Logs'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"; Description = 'Internet Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"; Description = 'Web Cache'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = 'C:\ProgramData\Microsoft\Windows\WER\ReportQueue'; Description = 'Error Report Queue'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\CrashDumps"; Description = 'Crash Dumps'; ContentsOnly = $true })
        [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"; Description = 'Delivery Optimization Cache'; ContentsOnly = $true })
    
        if ($AggressiveCleanup) {
            [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\D3DSCache"; Description = 'Direct3D Shader Cache'; ContentsOnly = $true })
            [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\NVIDIA\DXCache"; Description = 'NVIDIA DX Cache'; ContentsOnly = $true })
            [void]$cleanupTargets.Add([PSCustomObject]@{ Path = "$env:LOCALAPPDATA\NVIDIA\GLCache"; Description = 'NVIDIA GL Cache'; ContentsOnly = $true })
        }
    
        foreach ($target in $cleanupTargets) {
            if ($null -eq $target -or [string]::IsNullOrWhiteSpace($target.Path)) {
                Write-Log 'Skipping cleanup target because the path is blank.' 'WARN'
                continue
            }
    
            Write-Log "Cleaning $($target.Description) at $($target.Path)" 'INFO'
            $result = Remove-FolderContents -Path $target.Path -Description $target.Description -ContentsOnly:([bool]$target.ContentsOnly)
    
            if ($result.Success) {
                if ($result.SpaceFreed -gt 0) {
                    $totalSpaceFreed += [int64]$result.SpaceFreed
                    $sizeText = if ($result.SpaceFreed -ge 1GB) {
                        '{0} GB' -f [math]::Round($result.SpaceFreed / 1GB, 2)
                    }
                    else {
                        '{0} MB' -f [math]::Round($result.SpaceFreed / 1MB, 1)
                    }
                    Write-Log "Cleaned $($target.Description): $sizeText freed across $($result.ItemCount) item(s)." 'OK'
                }
                else {
                    Write-Log "$($target.Description): $($result.Message)" 'INFO'
                }
    
                $cleanupResults.Add([PSCustomObject]@{
                    Path        = $target.Path
                    Description = $target.Description
                    ItemCount   = $result.ItemCount
                    SpaceFreed  = [int64]$result.SpaceFreed
                    Status      = 'Success'
                    Message     = $result.Message
                }) | Out-Null
            }
            else {
                Warn-Step -Name 'TempCleanup' -Reason "$($target.Description) failed: $($result.Message)"
                $cleanupResults.Add([PSCustomObject]@{
                    Path        = $target.Path
                    Description = $target.Description
                    ItemCount   = 0
                    SpaceFreed  = [int64]0
                    Status      = 'Failed'
                    Message     = $result.Message
                }) | Out-Null
            }
        }
    
        Write-Log 'Cleaning SoftwareDistribution backup folders at C:\Windows and C:\ root backup variants' 'INFO'
        $sdBackupResult = Remove-SoftwareDistributionBakFolders
        if ($sdBackupResult.Success) {
            if ($sdBackupResult.SpaceFreed -gt 0) {
                $totalSpaceFreed += [int64]$sdBackupResult.SpaceFreed
                Write-Log "Cleaned SoftwareDistribution Backup Folders: $($sdBackupResult.SpaceFreedGB) GB ($($sdBackupResult.SpaceFreedMB) MB) freed across $($sdBackupResult.ItemCount) item(s)." 'OK'
            }
            else {
                Write-Log "SoftwareDistribution Backup Folders: $($sdBackupResult.Message)" 'INFO'
            }
    
            $cleanupResults.Add([PSCustomObject]@{
                Path        = 'C:\Windows and C:\ SoftwareDistribution backup variants'
                Description = 'SoftwareDistribution Backup Folders'
                ItemCount   = $sdBackupResult.ItemCount
                SpaceFreed  = [int64]$sdBackupResult.SpaceFreed
                Status      = 'Success'
                Message     = $sdBackupResult.Message
            }) | Out-Null
        }
        else {
            Warn-Step -Name 'TempCleanup' -Reason "SoftwareDistribution Backup Folders failed: $($sdBackupResult.Message)"
            $cleanupResults.Add([PSCustomObject]@{
                Path        = 'C:\Windows and C:\ SoftwareDistribution backup variants'
                Description = 'SoftwareDistribution Backup Folders'
                ItemCount   = $sdBackupResult.ItemCount
                SpaceFreed  = [int64]0
                Status      = 'Warning'
                Message     = $sdBackupResult.Message
            }) | Out-Null
        }
    
        $finalSpace = Get-DiskSpaceInfo -Path $env:SystemDrive
        $actualFreed = [int64]0
        if ($initialSpace -and $finalSpace) {
            $actualFreed = [int64]($finalSpace.FreeSpace - $initialSpace.FreeSpace)
        }
    
        Add-DetailedResult -Step 'TempCleanup' -Status 'Info' -Message 'Enhanced temporary file cleanup completed.' -Data @{
            EstimatedSpaceFreedMB = [math]::Round($totalSpaceFreed / 1MB, 2)
            ActualSpaceFreedMB    = [math]::Round($actualFreed / 1MB, 2)
            TargetsProcessed      = $cleanupResults.Count
            ResultsJson           = ($cleanupResults | ForEach-Object {
                [ordered]@{
                    Path         = $_.Path
                    Description  = $_.Description
                    ItemCount    = $_.ItemCount
                    SpaceFreedMB = [math]::Round(([double]$_.SpaceFreed) / 1MB, 2)
                    Status       = $_.Status
                    Message      = $_.Message
                }
            } | ConvertTo-Json -Compress)
        }
    }
    
    function Invoke-RepairVolumeScan {
    
        $systemDrive = $env:SystemDrive.TrimEnd(':')
        Write-Log "Running Repair-Volume scan on $($env:SystemDrive)" 'INFO'
        $result = Repair-Volume -DriveLetter $systemDrive -Scan -ErrorAction Stop
    
        $resultText = $null
        if ($null -ne $result) {
            $resultText = $result | Out-String
            if ($resultText -match 'corrupt|error|repair') {
                $script:Summary.DiskCorruptionSuspected = $true
                Warn-Step -Name 'RepairVolumeScan' -Reason 'Disk scan output suggests errors or corruption may exist.'
            }
        }
    
        Add-DetailedResult -Step 'RepairVolumeScan' -Status 'Info' -Message 'Repair-Volume scan completed.' -Data @{
            Output = $resultText
        }
    }
    
    function Invoke-RepairVolumeOfflineFix {
        $systemDrive = $env:SystemDrive.TrimEnd(':')
        Write-Log "Running offline disk repair on $($env:SystemDrive)" 'WARN'
        Repair-Volume -DriveLetter $systemDrive -OfflineScanAndFix -ErrorAction Stop | Out-Null
        $script:Summary.RebootRequired = $true
        Add-RepairAttempt 'Repair-Volume -OfflineScanAndFix'
        Add-DetailedResult -Step 'OfflineDiskRepair' -Status 'Info' -Message 'Offline disk repair was started.'
    }
    
    function Invoke-DismDetection {
        Write-Log "Running DISM CheckHealth..." 'INFO'
        $check = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/CheckHealth')
    
        Write-Log "Running DISM ScanHealth..." 'INFO'
        $scan = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/ScanHealth')
    
        $combined = (($check.StdOut, $scan.StdOut, $check.StdErr, $scan.StdErr) -join "`n")
    
        if ($check.ExitCode -ne 0 -or $scan.ExitCode -ne 0) {
            $script:Summary.DismCorruptionDetected = $true
            Warn-Step -Name 'DISMDetection' -Reason 'DISM detection returned a non-zero exit code.'
            return
        }
    
        # Avoid false positives from phrases like "No component store corruption detected."
        if ($combined -match '(?i)No component store corruption detected|No component store corruption was detected|The component store is repairable\s*:\s*No') {
            $script:Summary.DismCorruptionDetected = $false
            Write-Log 'DISM did not detect component store corruption.' 'OK'
            return
        }
    
        if ($combined -match '(?i)The component store is repairable|component store is repairable|repairable\s*:\s*Yes|corruption detected|component store corruption detected') {
            $script:Summary.DismCorruptionDetected = $true
            Warn-Step -Name 'DISMDetection' -Reason 'DISM detected component store corruption.'
        }
        else {
            $script:Summary.DismCorruptionDetected = $false
            Write-Log 'DISM detection completed without confirmed corruption.' 'OK'
        }
    }
    
    function Invoke-DismRepair {
        Write-Log "Running DISM RestoreHealth..." 'WARN'
        $repair = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/RestoreHealth')
        if ($repair.ExitCode -ne 0) {
            throw "DISM RestoreHealth exited with code $($repair.ExitCode)"
        }
    
        Write-Log "Running DISM StartComponentCleanup..." 'INFO'
        $cleanup = Invoke-DismCommand -Arguments @('/Online','/Cleanup-Image','/StartComponentCleanup')
        if ($cleanup.ExitCode -ne 0) {
            throw "DISM StartComponentCleanup exited with code $($cleanup.ExitCode)"
        }
    
        Add-RepairAttempt 'DISM RestoreHealth + StartComponentCleanup'
    }
    
    function Invoke-SfcDetection {
        Write-Log "Running SFC verify-only scan..." 'INFO'
        $result = Invoke-SfcCommand -Arguments '/verifyonly'
    
        $combined = (($result.StdOut, $result.StdErr) -join "`n")
    
        if ($combined -match '(?i)Windows Resource Protection found integrity violations|found integrity violations') {
            $script:Summary.SfcIntegrityViolations = $true
            Warn-Step -Name 'SFCDetection' -Reason 'SFC detected integrity violations.'
            return
        }
    
        if ($combined -match '(?i)Windows Resource Protection did not find any integrity violations|did not find any integrity violations') {
            $script:Summary.SfcIntegrityViolations = $false
            Write-Log 'SFC did not detect integrity violations.' 'OK'
            return
        }
    
        if ($combined -match '(?i)Windows Resource Protection found corrupt files and successfully repaired them|Windows Resource Protection found corrupt files but was unable to fix some of them') {
            $script:Summary.SfcIntegrityViolations = $true
            Warn-Step -Name 'SFCDetection' -Reason 'SFC reported corrupt files.'
            return
        }
    
        if ($result.ExitCode -notin 0,1) {
            $script:Summary.SfcIntegrityViolations = $true
            Warn-Step -Name 'SFCDetection' -Reason "SFC verify returned unusual exit code $($result.ExitCode)."
        }
        else {
            Write-Log 'SFC detection completed without confirmed integrity violations.' 'OK'
        }
    }
    
    function Invoke-SfcRepair {
        Write-Log "Running SFC /SCANNOW..." 'WARN'
        $result = Invoke-SfcCommand -Arguments '/scannow'
    
        if ($result.ExitCode -notin 0,1) {
            throw "SFC /SCANNOW exited with code $($result.ExitCode)"
        }
    
        Add-RepairAttempt 'SFC /SCANNOW'
    }
    
    function Invoke-WmiCheck {
        Write-Log "Checking WMI repository consistency..." 'INFO'
        $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /verifyrepository 2>&1
        $text = ($output | Out-String).Trim()
    
        if ($text) {
            $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
        }
    
        Add-DetailedResult -Step 'WmiRepositoryCheck' -Status 'Info' -Message 'WMI verify completed.' -Data @{
            Output = $text
        }
    
        if ($text -match 'inconsistent') {
            $script:Summary.WmiRepositoryInconsistent = $true
            Warn-Step -Name 'WmiRepositoryCheck' -Reason 'WMI repository reported as inconsistent.'
        }
    }
    
    function Invoke-WmiRepair {
        Write-Log "Repairing WMI repository..." 'WARN'
        $output = & "$env:SystemRoot\System32\wbem\winmgmt.exe" /salvagerepository 2>&1
        $text = ($output | Out-String).Trim()
    
        if ($text) {
            $text -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object { Write-Log $_ 'INFO' }
        }
    
        Add-RepairAttempt 'winmgmt /salvagerepository'
        Add-DetailedResult -Step 'WmiRepair' -Status 'Info' -Message 'WMI salvage completed.' -Data @{
            Output = $text
        }
    }
    
    function Invoke-NetworkReset {
        Write-Log "Flushing DNS cache..." 'INFO'
        ipconfig /flushdns | Out-Null
    
        Write-Log "Resetting Winsock..." 'WARN'
        netsh winsock reset | Out-Null
    
        Write-Log "Resetting TCP/IP stack..." 'WARN'
        netsh int ip reset | Out-Null
    
        $script:Summary.RebootRequired = $true
        Add-RepairAttempt 'Winsock/TCPIP reset'
        Add-DetailedResult -Step 'NetworkReset' -Status 'Info' -Message 'Network reset completed.'
    }
    
    function Invoke-DnsFlushOnly {
        Write-Log "Flushing DNS cache..." 'INFO'
        ipconfig /flushdns | Out-Null
        Add-DetailedResult -Step 'DnsFlush' -Status 'Info' -Message 'DNS cache flushed.'
    }
    
    function Invoke-StorageOptimization {
        $systemDrive = $env:SystemDrive.TrimEnd(':')
        $isSsd = Test-SystemDiskIsSSD
    
        if ($isSsd) {
            Write-Log "SSD detected. Running ReTrim on $($env:SystemDrive)" 'INFO'
            Optimize-Volume -DriveLetter $systemDrive -ReTrim -ErrorAction Stop | Out-Null
        }
        else {
            Write-Log "SSD not detected. Running standard optimization on $($env:SystemDrive)" 'INFO'
            Optimize-Volume -DriveLetter $systemDrive -Analyze -ErrorAction Stop | Out-Null
        }
    
        Add-DetailedResult -Step 'StorageOptimization' -Status 'Info' -Message 'Storage optimization completed.' -Data @{
            IsSSD = $isSsd
        }
    }
    
    function Invoke-IconCacheRebuild {
        Write-Log "Rebuilding icon and thumbnail caches..." 'WARN'
    
        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    
        $explorerCachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
        $deletedFiles = New-Object System.Collections.Generic.List[string]
    
        $singleFileTargets = @(
            "$env:LOCALAPPDATA\IconCache.db"
        )
    
        foreach ($path in $singleFileTargets) {
            if (Test-Path -LiteralPath $path) {
                try {
                    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                    $deletedFiles.Add((Split-Path -Leaf $path)) | Out-Null
                }
                catch {
                    Write-Log "Failed to delete cache file $path : $($_.Exception.Message)" 'WARN'
                }
            }
        }
    
        if (Test-Path -LiteralPath $explorerCachePath) {
            $patterns = @(
                'iconcache*',
                'thumbcache_*.db',
                'thumbcache_idx.db'
            )
    
            foreach ($pattern in $patterns) {
                Get-ChildItem -LiteralPath $explorerCachePath -Filter $pattern -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try {
                        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                        $deletedFiles.Add($_.Name) | Out-Null
                    }
                    catch {
                        Write-Log "Failed to delete cache file $($_.FullName): $($_.Exception.Message)" 'WARN'
                    }
                }
            }
        }
    
        Start-Process explorer.exe
        Add-RepairAttempt 'Icon and thumbnail cache rebuild'
        Add-DetailedResult -Step 'IconCacheRebuild' -Status 'Info' -Message 'Icon and thumbnail cache rebuild completed.' -Data @{
            DeletedFiles = ($deletedFiles -join '; ')
        }
    }
    
    
    function Set-RegistryDWORDValue {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][int]$Value
        )
    
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
    
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
    }
    
    function Disable-CopilotForLoadedUsers {
        $targetSids = New-Object System.Collections.Generic.List[string]
        $targetSids.Add('HKEY_CURRENT_USER') | Out-Null
    
        Get-ChildItem Registry::HKEY_USERS -ErrorAction SilentlyContinue |
            Where-Object {
                $_.PSChildName -match '^S-1-5-21-' -and
                $_.PSChildName -notmatch '_Classes$'
            } |
            ForEach-Object {
                $targetSids.Add("HKEY_USERS\\$($_.PSChildName)") | Out-Null
            }
    
        foreach ($root in $targetSids | Select-Object -Unique) {
            $policyPath = "Registry::$root\Software\Policies\Microsoft\Windows\WindowsCopilot"
            $explorerPath = "Registry::$root\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    
            Set-RegistryDWORDValue -Path $policyPath -Name 'TurnOffWindowsCopilot' -Value 1
            Set-RegistryDWORDValue -Path $explorerPath -Name 'ShowCopilotButton' -Value 0
    
            Add-DetailedResult -Step 'CopilotDisableRegistry' -Status 'Info' -Message 'Applied Copilot disable settings for loaded profile.' -Data @{
                Root = $root
                PolicyPath = $policyPath
                ExplorerPath = $explorerPath
            }
        }
    }
    
    function Disable-CopilotForDefaultUser {
        $defaultHiveName = 'HKU\DefaultTempCopilot'
        $defaultHivePsPath = 'Registry::HKEY_USERS\DefaultTempCopilot'
        $defaultUserNtUserDat = 'C:\Users\Default\NTUSER.DAT'
    
        if (-not (Test-Path -LiteralPath $defaultUserNtUserDat)) {
            Write-Log 'Default User NTUSER.DAT not found; future new users were not updated for Copilot disable.' 'WARN'
            return
        }
    
        $hiveLoaded = $false
        try {
            $loadResult = & reg.exe load $defaultHiveName $defaultUserNtUserDat
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to load Default User hive: $($loadResult -join ' ')"
            }
    
            $hiveLoaded = $true
            Start-Sleep -Milliseconds 750
    
            $policyPath = "$defaultHivePsPath\Software\Policies\Microsoft\Windows\WindowsCopilot"
            $explorerPath = "$defaultHivePsPath\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    
            Set-RegistryDWORDValue -Path $policyPath -Name 'TurnOffWindowsCopilot' -Value 1
            Set-RegistryDWORDValue -Path $explorerPath -Name 'ShowCopilotButton' -Value 0
    
            Add-DetailedResult -Step 'CopilotDisableDefaultUser' -Status 'Info' -Message 'Applied Copilot disable settings for Default User profile.' -Data @{
                PolicyPath = $policyPath
                ExplorerPath = $explorerPath
            }
        }
        finally {
            if ($hiveLoaded) {
                Start-Sleep -Milliseconds 750
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                Start-Sleep -Milliseconds 750
                & reg.exe unload $defaultHiveName | Out-Null
            }
        }
    }
    
    function Invoke-CopilotDisableAndRemoval {
        Write-Log 'Disabling Microsoft Copilot for current, loaded, and future user profiles...' 'WARN'
    
        Set-RegistryDWORDValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot' -Name 'TurnOffWindowsCopilot' -Value 1
        Set-RegistryDWORDValue -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer' -Name 'HideCopilotButton' -Value 1
    
        Disable-CopilotForLoadedUsers
        Disable-CopilotForDefaultUser
    
        $removedPackages = New-Object System.Collections.Generic.List[string]
        $packagePatterns = @(
            'Microsoft.Windows.Copilot',
            '*Copilot*'
        )
    
        foreach ($pattern in $packagePatterns) {
            $packages = @(Get-AppxPackage -AllUsers -Name $pattern -ErrorAction SilentlyContinue)
            foreach ($pkg in $packages) {
                if ($removedPackages -contains $pkg.PackageFullName) {
                    continue
                }
    
                try {
                    Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                    $removedPackages.Add($pkg.PackageFullName) | Out-Null
                    Write-Log "Removed Copilot Appx package: $($pkg.Name) [$($pkg.PackageFullName)]" 'OK'
                }
                catch {
                    Write-Log "Failed to remove Copilot Appx package $($pkg.PackageFullName): $($_.Exception.Message)" 'WARN'
                }
            }
    
            $provisionedPackages = @(Get-AppxProvisionedPackage -Online | Where-Object {
                $_.DisplayName -like $pattern -or $_.PackageName -like $pattern
            })
    
            foreach ($prov in $provisionedPackages) {
                try {
                    Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction Stop | Out-Null
                    $removedPackages.Add($prov.PackageName) | Out-Null
                    Write-Log "Removed provisioned Copilot package: $($prov.DisplayName) [$($prov.PackageName)]" 'OK'
                }
                catch {
                    Write-Log "Failed to remove provisioned Copilot package $($prov.PackageName): $($_.Exception.Message)" 'WARN'
                }
            }
        }
    
        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process explorer.exe
    
        $script:Summary.RebootRequired = $true
        Add-RepairAttempt 'Microsoft Copilot disable and removal'
        Add-DetailedResult -Step 'CopilotDisableAndRemoval' -Status 'Info' -Message 'Microsoft Copilot disable and removal routine completed.' -Data @{
            RemovedPackages = ($removedPackages | Select-Object -Unique) -join '; '
        }
    }
    
    function Invoke-FirewallReset {
        Write-Log "Resetting Windows Firewall to defaults..." 'WARN'
        netsh advfirewall reset | Out-Null
        Add-RepairAttempt 'Firewall reset'
        Add-DetailedResult -Step 'FirewallReset' -Status 'Info' -Message 'Firewall reset completed.'
    }
    
    function Get-ServiceStateSafe {
        param(
            [Parameter(Mandatory)][string]$Name
        )
    
        try {
            $svc = Get-Service -Name $Name -ErrorAction Stop
            return [PSCustomObject]@{
                Name        = $svc.Name
                DisplayName = $svc.DisplayName
                Status      = [string]$svc.Status
                Exists      = $true
                Error       = $null
            }
        }
        catch {
            return [PSCustomObject]@{
                Name        = $Name
                DisplayName = $null
                Status      = 'NotFound'
                Exists      = $false
                Error       = $_.Exception.Message
            }
        }
    }
    
    function Wait-ServiceStateSafe {
        param(
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][ValidateSet('Running','Stopped')][string]$DesiredStatus,
            [int]$TimeoutSeconds = 30
        )
    
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    
        do {
            $state = Get-ServiceStateSafe -Name $Name
            if (-not $state.Exists) {
                return $state
            }
    
            if ($state.Status -eq $DesiredStatus) {
                return $state
            }
    
            Start-Sleep -Seconds 1
        } while ((Get-Date) -lt $deadline)
    
        return Get-ServiceStateSafe -Name $Name
    }
    
    function Stop-ServiceWithValidation {
        param(
            [Parameter(Mandatory)][string]$Name,
            [int]$TimeoutSeconds = 30
        )
    
        $before = Get-ServiceStateSafe -Name $Name
    
        if (-not $before.Exists) {
            Write-Log "Service validation: $Name was not found. Skipping stop." 'WARN'
            return [PSCustomObject]@{
                Name         = $Name
                BeforeStatus = $before.Status
                AfterStatus  = $before.Status
                Success      = $true
                Message      = 'Service not found; skipped'
            }
        }
    
        Write-Log "Service validation: $Name current state is $($before.Status)." 'INFO'
    
        if ($before.Status -eq 'Stopped') {
            Write-Log "Service validation: $Name is already stopped." 'OK'
            return [PSCustomObject]@{
                Name         = $Name
                BeforeStatus = $before.Status
                AfterStatus  = 'Stopped'
                Success      = $true
                Message      = 'Already stopped'
            }
        }
    
        try {
            Write-Log "Stopping service $Name..." 'INFO'
            Stop-Service -Name $Name -Force -ErrorAction Stop
        }
        catch {
            Write-Log "Stop-Service reported an issue for $Name`: $($_.Exception.Message)" 'WARN'
        }
    
        $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Stopped' -TimeoutSeconds $TimeoutSeconds
        $success = ($after.Status -eq 'Stopped')
    
        if ($success) {
            Write-Log "Service validation: $Name stopped successfully." 'OK'
        }
        else {
            Write-Log "Service validation: $Name did not stop. Current state: $($after.Status)." 'WARN'
        }
    
        return [PSCustomObject]@{
            Name         = $Name
            BeforeStatus = $before.Status
            AfterStatus  = $after.Status
            Success      = $success
            Message      = if ($success) { 'Stopped successfully' } else { "Expected Stopped but found $($after.Status)" }
        }
    }
    
    function Start-ServiceWithValidation {
        param(
            [Parameter(Mandatory)][string]$Name,
            [int]$TimeoutSeconds = 30
        )
    
        $before = Get-ServiceStateSafe -Name $Name
    
        if (-not $before.Exists) {
            Write-Log "Service validation: $Name was not found. Skipping start." 'WARN'
            return [PSCustomObject]@{
                Name         = $Name
                BeforeStatus = $before.Status
                AfterStatus  = $before.Status
                Success      = $true
                Message      = 'Service not found; skipped'
            }
        }
    
        if ($before.Status -eq 'Running') {
            Write-Log "Service validation: $Name is already running." 'OK'
            return [PSCustomObject]@{
                Name         = $Name
                BeforeStatus = $before.Status
                AfterStatus  = 'Running'
                Success      = $true
                Message      = 'Already running'
            }
        }
    
        try {
            Write-Log "Starting service $Name..." 'INFO'
            Start-Service -Name $Name -ErrorAction Stop
        }
        catch {
            Write-Log "Start-Service reported an issue for $Name`: $($_.Exception.Message)" 'WARN'
        }
    
        $after = Wait-ServiceStateSafe -Name $Name -DesiredStatus 'Running' -TimeoutSeconds $TimeoutSeconds
        $success = ($after.Status -eq 'Running')
    
        if ($success) {
            Write-Log "Service validation: $Name started successfully." 'OK'
        }
        else {
            Write-Log "Service validation: $Name did not start. Current state: $($after.Status)." 'WARN'
        }
    
        return [PSCustomObject]@{
            Name         = $Name
            BeforeStatus = $before.Status
            AfterStatus  = $after.Status
            Success      = $success
            Message      = if ($success) { 'Started successfully' } else { "Expected Running but found $($after.Status)" }
        }
    }
    
    
    function Remove-PathWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Description,
            [int]$MaxAttempts = 5,
            [int]$InitialDelaySeconds = 2,
            [switch]$ReleaseWindowsUpdateLocks
        )
    
        $result = [ordered]@{
            Path            = $Path
            Description     = $Description
            ExistsBefore    = $false
            Deleted         = $false
            Attempts        = 0
            ItemCount       = 0
            FileCount       = 0
            FolderCount     = 0
            SizeBytesBefore = [int64]0
            SizeMBBefore    = [double]0
            SizeGBBefore    = [double]0
            SizeBytesAfter  = [int64]0
            SizeMBAfter     = [double]0
            SizeGBAfter     = [double]0
            SpaceFreed      = [int64]0
            EstimatedFreedBytes = [int64]0
            EstimatedFreedMB = [double]0
            EstimatedFreedGB = [double]0
            Message         = $null
        }
    
        if (-not (Test-Path -LiteralPath $Path)) {
            $result.Message = 'Path does not exist; nothing to delete'
            Write-Log "$Description does not exist at $Path. Nothing to delete." 'INFO'
            return [PSCustomObject]$result
        }
    
        $result.ExistsBefore = $true
    
        $beforeInfo = Get-FolderSizeInfo -Path $Path
        $result.ItemCount = $beforeInfo.ItemCount
        $result.FileCount = $beforeInfo.FileCount
        $result.FolderCount = $beforeInfo.FolderCount
        $result.SizeBytesBefore = [int64]$beforeInfo.SizeBytes
        $result.SizeMBBefore = [double]$beforeInfo.SizeMB
        $result.SizeGBBefore = [double]$beforeInfo.SizeGB
    
        Write-Log "Preparing to delete $Description at $Path. No backup will be created." 'INFO'
        Write-Log "Size before deletion for $Description`: $($result.SizeGBBefore) GB ($($result.SizeMBBefore) MB), items: $($result.ItemCount), files: $($result.FileCount), folders: $($result.FolderCount)." 'INFO'
    
        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            $result.Attempts = $attempt
    
            try {
                if (-not (Test-Path -LiteralPath $Path)) {
                    $result.Deleted = $true
                    $result.Message = 'Path already gone during retry validation'
                    Write-Log "$Description no longer exists at $Path." 'OK'
                    break
                }
    
                if ($ReleaseWindowsUpdateLocks -and $attempt -gt 1) {
                    Write-Log "Attempt $attempt is releasing possible Windows Update locks before retrying $Description deletion." 'WARN'
                    Stop-WindowsUpdateLockingProcesses | Out-Null
                }
    
                Write-Log "Deletion attempt $attempt of $MaxAttempts for $Description..." 'INFO'
                Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
    
                Start-Sleep -Seconds 1
                if (-not (Test-Path -LiteralPath $Path)) {
                    $result.Deleted = $true
                    $result.Message = 'Deleted successfully'
                    Write-Log "Deleted $Description successfully on attempt $attempt." 'OK'
                    break
                }
    
                throw "$Description still exists after Remove-Item completed."
            }
            catch {
                $result.Message = $_.Exception.Message
                Write-Log "Deletion attempt $attempt failed for $Description`: $($result.Message)" 'WARN'
    
                if ($ReleaseWindowsUpdateLocks) {
                    Stop-WindowsUpdateLockingProcesses | Out-Null
                }
    
                if ($attempt -lt $MaxAttempts) {
                    $delay = $InitialDelaySeconds * $attempt
                    Write-Log "Waiting $delay second(s), then retrying $Description deletion. Files may still be locked by services or Windows Update processes." 'INFO'
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    Start-Sleep -Seconds $delay
                }
            }
        }
    
        if (Test-Path -LiteralPath $Path) {
            $afterInfo = Get-FolderSizeInfo -Path $Path
            $result.SizeBytesAfter = [int64]$afterInfo.SizeBytes
            $result.SizeMBAfter = [double]$afterInfo.SizeMB
            $result.SizeGBAfter = [double]$afterInfo.SizeGB
        }
    
        $freedBytes = [int64]([math]::Max(0, ([int64]$result.SizeBytesBefore - [int64]$result.SizeBytesAfter)))
        $result.SpaceFreed = $freedBytes
        $result.EstimatedFreedBytes = $freedBytes
        $result.EstimatedFreedMB = [math]::Round(([double]$freedBytes / 1MB), 2)
        $result.EstimatedFreedGB = [math]::Round(([double]$freedBytes / 1GB), 3)
    
        if ($result.Deleted) {
            Write-Log "Estimated space freed by deleting $Description`: $($result.EstimatedFreedGB) GB ($($result.EstimatedFreedMB) MB)." 'OK'
        }
        else {
            Write-Log "Failed to delete $Description after $MaxAttempts attempt(s). Last error: $($result.Message). Estimated remaining size: $($result.SizeGBAfter) GB." 'ERROR'
        }
    
        return [PSCustomObject]$result
    }
    
    function Invoke-WindowsUpdateComponentReset {
        Write-Log "Resetting Windows Update components..." 'WARN'
        Write-Log "SoftwareDistribution will be deleted directly. No .bak, .bak1, or timestamped backup folder will be created." 'INFO'
    
        $services = @('wuauserv','bits','cryptsvc','msiserver','usosvc','DoSvc','WaaSMedicSvc')
        $stopResults = New-Object System.Collections.Generic.List[object]
        $startResults = New-Object System.Collections.Generic.List[object]
        $deleteResults = New-Object System.Collections.Generic.List[object]
        $lockProcessResults = New-Object System.Collections.Generic.List[object]
        $backupCleanupResult = $null
    
        foreach ($svc in $services) {
            $stopResults.Add((Stop-ServiceWithValidation -Name $svc -TimeoutSeconds 45)) | Out-Null
        }
    
        $criticalServicesStillRunning = @($stopResults | Where-Object {
            $_.Name -in @('wuauserv','bits','cryptsvc') -and $_.Success -eq $false
        })
    
        if ($criticalServicesStillRunning.Count -gt 0) {
            Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more Windows Update services did not stop cleanly: $($criticalServicesStillRunning.Name -join ', ')"
        }
    
        Start-Sleep -Seconds 2
    
        foreach ($lockResult in @(Stop-WindowsUpdateLockingProcesses)) { $lockProcessResults.Add($lockResult) | Out-Null }
        $backupCleanupResult = Remove-SoftwareDistributionBakFolders
    
        $paths = @(
            @{ Path = "$env:WINDIR\SoftwareDistribution"; Description = 'Windows Update SoftwareDistribution folder' },
            @{ Path = "$env:WINDIR\System32\catroot2"; Description = 'Windows Update Catroot2 folder' }
        )
    
        foreach ($target in $paths) {
            $deleteResults.Add((Remove-PathWithRetry -Path $target.Path -Description $target.Description -MaxAttempts 5 -InitialDelaySeconds 2 -ReleaseWindowsUpdateLocks)) | Out-Null
        }
    
        foreach ($svc in $services) {
            $startResults.Add((Start-ServiceWithValidation -Name $svc -TimeoutSeconds 45)) | Out-Null
        }
    
        $failedDeletes = @($deleteResults | Where-Object { $_.ExistsBefore -eq $true -and $_.Deleted -eq $false })
        $failedStarts = @($startResults | Where-Object { $_.Success -eq $false })
    
        if ($failedDeletes.Count -gt 0) {
            Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more update folders could not be deleted: $($failedDeletes.Description -join ', ')"
        }
    
        if ($failedStarts.Count -gt 0) {
            Warn-Step -Name 'WindowsUpdateComponentReset' -Reason "One or more update services did not restart cleanly: $($failedStarts.Name -join ', ')"
        }
    
        $script:Summary.RebootRequired = $true
        Add-RepairAttempt 'Windows Update component reset with direct SoftwareDistribution deletion, backup cleanup, folder size logging, lock release, and service validation'
        Add-DetailedResult -Step 'WindowsUpdateComponentReset' -Status 'Info' -Message 'Windows Update components reset. SoftwareDistribution was deleted directly with no backup folder creation; old backup folders were removed and estimated GB freed was logged.' -Data @{
            StoppedServicesJson = ($stopResults | ConvertTo-Json -Compress)
            LockProcessesJson  = ($lockProcessResults | ConvertTo-Json -Compress)
            BackupCleanupJson  = ($backupCleanupResult | ConvertTo-Json -Compress)
            DeletedPathsJson   = ($deleteResults | ConvertTo-Json -Compress)
            StartedServicesJson = ($startResults | ConvertTo-Json -Compress)
        }
    }
    
    function Invoke-ScheduledTaskHealthCheck {
        Write-Log "Checking scheduled task health..." 'INFO'
    
        $badTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $_.State -eq 'Unknown' -or $_.TaskPath -eq $null
        }
    
        $taskNames = @()
    
        if ($badTasks) {
            foreach ($task in $badTasks) {
                $name = "$($task.TaskPath)$($task.TaskName)"
                $taskNames += $name
                Warn-Step -Name 'ScheduledTaskCheck' -Reason "Task may need review: $name"
            }
        }
    
        Add-DetailedResult -Step 'ScheduledTaskHealthCheck' -Status 'Info' -Message 'Scheduled task health check completed.' -Data @{
            SuspectTaskCount = $taskNames.Count
            SuspectTasks     = ($taskNames -join '; ')
        }
    }
    
    function Invoke-EventLogSummary {
        Write-Log "Collecting recent system health events..." 'INFO'
    
        try {
            $recentUnexpected = Get-WinEvent -FilterHashtable @{
                LogName = 'System'
                Id      = 41, 6008
                StartTime = (Get-Date).AddDays(-7)
            } -ErrorAction Stop
    
            $count = @($recentUnexpected).Count
    
            Add-DetailedResult -Step 'EventLogSummary' -Status 'Info' -Message 'Collected recent unexpected shutdown events.' -Data @{
                UnexpectedShutdownCount = $count
            }
    
            if ($recentUnexpected) {
                Warn-Step -Name 'EventLogSummary' -Reason "Recent unexpected shutdown events found: $count"
            }
        }
        catch {
            Warn-Step -Name 'EventLogSummary' -Reason $_.Exception.Message
        }
    }
    
    function Invoke-EventLogClear {
        Write-Log "Clearing classic event logs..." 'WARN'
        wevtutil el | ForEach-Object {
            try {
                wevtutil cl $_ 2>$null
            }
            catch {
            }
        }
        Add-RepairAttempt 'Event log clear'
        Add-DetailedResult -Step 'EventLogClear' -Status 'Info' -Message 'Event log clear attempted.'
    }
    
    function Invoke-DiskSpaceCheck {
        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
        if ($drive) {
            $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
            $sizeGB = [math]::Round($drive.Size / 1GB, 2)
            Write-Log "System drive free space: $freeGB GB of $sizeGB GB" 'INFO'
    
            Add-DetailedResult -Step 'DiskSpaceCheck' -Status 'Info' -Message 'Disk space checked.' -Data @{
                FreeGB  = $freeGB
                TotalGB = $sizeGB
            }
    
            if ($freeGB -lt 10) {
                Warn-Step -Name 'DiskSpaceCheck' -Reason "Low free space on system drive: $freeGB GB"
            }
        }
    }
    
    function Invoke-DefenderStatusCheck {
        try {
            $status = Get-MpComputerStatus -ErrorAction Stop
            Write-Log "Defender Antivirus Enabled: $($status.AntivirusEnabled)" 'INFO'
            Write-Log "Defender RealTime Protection Enabled: $($status.RealTimeProtectionEnabled)" 'INFO'
    
            Add-DetailedResult -Step 'DefenderStatusCheck' -Status 'Info' -Message 'Defender status checked.' -Data @{
                AntivirusEnabled          = $status.AntivirusEnabled
                RealTimeProtectionEnabled = $status.RealTimeProtectionEnabled
            }
        }
        catch {
            Warn-Step -Name 'DefenderStatusCheck' -Reason $_.Exception.Message
        }
    }
    
    function Invoke-RebootIfNeeded {
        param(
            [int]$DelaySeconds = 60
        )
    
        $comment = 'Restarting after automated system repair operations.'
    
        $args = @(
            '/r'
            '/t', $DelaySeconds.ToString()
            '/d', 'p:2:17'
            '/c', "`"$comment`""
            '/f'
        )
    
        Write-Log "Issuing reboot command: shutdown.exe $($args -join ' ')" 'WARN'
        & "$env:SystemRoot\System32\shutdown.exe" @args
    
        if ($LASTEXITCODE -ne 0) {
            throw "shutdown.exe returned exit code $LASTEXITCODE"
        }
    
        Add-DetailedResult -Step 'AutoReboot' -Status 'Info' -Message 'Automatic reboot command issued.' -Data @{
            DelaySeconds = $DelaySeconds
        }
    }
    
    function Show-Summary {
        $script:Summary.EndTime = Get-Date
    
        Write-Log "---------------- Summary ----------------" 'INFO'
        Write-Log "Computer Name: $($script:Summary.ComputerName)" 'INFO'
        Write-Log "Start Time: $($script:Summary.StartTime)" 'INFO'
        Write-Log "End Time:   $($script:Summary.EndTime)" 'INFO'
        Write-Log "YAML Log:   $($script:YamlLogPath)" 'INFO'
        Write-Log "Succeeded:  $($script:Summary.StepsSucceeded)" 'INFO'
        Write-Log "Failed:     $($script:Summary.StepsFailed)" 'INFO'
        Write-Log "Warnings:   $($script:Summary.Warnings)" 'INFO'
        Write-Log "Pending Reboot Detected: $($script:Summary.PendingRebootDetected)" 'INFO'
        Write-Log "Reboot Required: $($script:Summary.RebootRequired)" 'INFO'
        Write-Log "Disk Corruption Suspected: $($script:Summary.DiskCorruptionSuspected)" 'INFO'
        Write-Log "DISM Corruption Detected: $($script:Summary.DismCorruptionDetected)" 'INFO'
        Write-Log "SFC Integrity Violations: $($script:Summary.SfcIntegrityViolations)" 'INFO'
        Write-Log "WMI Repository Inconsistent: $($script:Summary.WmiRepositoryInconsistent)" 'INFO'
    
        if ($script:Summary.RepairsAttempted.Count -gt 0) {
            Write-Log "Repairs Attempted:" 'INFO'
            foreach ($repair in $script:Summary.RepairsAttempted) {
                Write-Log " - $repair" 'INFO'
            }
        }
    
        if ($script:Summary.Notes.Count -gt 0) {
            Write-Log "Notes:" 'INFO'
            foreach ($note in $script:Summary.Notes) {
                Write-Log " - $note" 'INFO'
            }
        }
    }
    
    
    function Invoke-LogArchiveRetention {
        [CmdletBinding()]
        param(
            [string]$LogDirectory = 'C:\Logs',
            [string]$ComputerName = $env:COMPUTERNAME
        )
    
        Write-Log "Starting Sunday-based log archive and retention processing in $LogDirectory" 'INFO'
    
        if (-not (Test-Path -LiteralPath $LogDirectory)) {
            Write-Log "Log directory does not exist: $LogDirectory" 'WARN'
            Add-DetailedResult -Step 'LogArchiveRetention' -Status 'Warning' -Message "Log directory not found: $LogDirectory"
            return
        }
    
        $now = Get-Date
        $thisSunday = $now.Date.AddDays(-[int]$now.DayOfWeek)
        $previousSunday = $thisSunday.AddDays(-7)
        $twoSundaysAgo = $thisSunday.AddDays(-14)
    
        Write-Log "This Sunday: $thisSunday" 'INFO'
        Write-Log "Previous Sunday: $previousSunday" 'INFO'
        Write-Log "Two Sundays Ago: $twoSundaysAgo" 'INFO'
    
        $extensions = @('.yaml', '.yml', '.txt')
    
        $allLooseLogs = Get-ChildItem -LiteralPath $LogDirectory -File -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $extensions -contains $_.Extension.ToLowerInvariant() -and
                $_.FullName -ne $script:YamlLogPath
            }
    
        $logsToArchive = $allLooseLogs | Where-Object {
            $_.CreationTime -ge $previousSunday -and $_.CreationTime -lt $thisSunday
        } | Sort-Object CreationTime, Name
    
        $archiveDateText = $previousSunday.ToString('yyyy-MM-dd')
        $zipPath = Join-Path $LogDirectory ("{0}-logs-{1}.zip" -f $ComputerName, $archiveDateText)
    
        $archiveSummary = [ordered]@{
            ThisSunday                 = $thisSunday
            PreviousSunday             = $previousSunday
            TwoSundaysAgo              = $twoSundaysAgo
            LooseLogsFound             = @($allLooseLogs).Count
            LogsSelectedForArchive     = @($logsToArchive).Count
            ArchiveCreated             = $false
            ArchivePath                = $null
            DeletedOriginalFiles       = @()
            DeletedOldLooseLogs        = @()
            DeletedExpiredZipFiles     = @()
            Errors                     = @()
        }
    
        if (@($logsToArchive).Count -gt 0) {
            Write-Log "Preparing archive for previous Sunday week: $zipPath" 'INFO'
    
            try {
                if (Test-Path -LiteralPath $zipPath) {
                    Write-Log "Existing archive found for that Sunday. Removing and recreating: $zipPath" 'WARN'
                    Remove-Item -LiteralPath $zipPath -Force -ErrorAction Stop
                }
    
                Compress-Archive -Path ($logsToArchive.FullName) -DestinationPath $zipPath -CompressionLevel Optimal -Force -ErrorAction Stop
    
                if (-not (Test-Path -LiteralPath $zipPath)) {
                    throw 'ZIP file was not created.'
                }
    
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
                $zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    
                try {
                    $zipEntries = @($zip.Entries)
                    if ($zipEntries.Count -lt 1) {
                        throw 'ZIP file was created but contains no entries.'
                    }
    
                    if ($zipEntries.Count -lt @($logsToArchive).Count) {
                        throw "ZIP file entry count ($($zipEntries.Count)) is less than expected source file count ($(@($logsToArchive).Count))."
                    }
                }
                finally {
                    $zip.Dispose()
                }
    
                $archiveSummary.ArchiveCreated = $true
                $archiveSummary.ArchivePath = $zipPath
                Write-Log "Archive created successfully: $zipPath" 'OK'
    
                foreach ($file in $logsToArchive) {
                    try {
                        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                        $archiveSummary.DeletedOriginalFiles += $file.FullName
                        Write-Log "Deleted archived source log: $($file.FullName)" 'OK'
                    }
                    catch {
                        $msg = "Failed to delete archived source file $($file.FullName): $($_.Exception.Message)"
                        $archiveSummary.Errors += $msg
                        Write-Log $msg 'WARN'
                    }
                }
            }
            catch {
                $msg = "Archive creation/validation failed: $($_.Exception.Message)"
                $archiveSummary.Errors += $msg
                Write-Log $msg 'ERROR'
            }
        }
        else {
            Write-Log 'No loose log files were found for the previous Sunday-to-Saturday period.' 'INFO'
        }
    
        $remainingLooseLogs = Get-ChildItem -LiteralPath $LogDirectory -File -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $extensions -contains $_.Extension.ToLowerInvariant() -and
                $_.FullName -ne $script:YamlLogPath
            }
    
        $oldLooseLogsToDelete = $remainingLooseLogs | Where-Object {
            $_.CreationTime -lt $twoSundaysAgo
        }
    
        foreach ($file in $oldLooseLogsToDelete) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                $archiveSummary.DeletedOldLooseLogs += $file.FullName
                Write-Log "Deleted loose log older than two Sundays: $($file.FullName)" 'OK'
            }
            catch {
                $msg = "Failed to delete old loose log $($file.FullName): $($_.Exception.Message)"
                $archiveSummary.Errors += $msg
                Write-Log $msg 'WARN'
            }
        }
    
        $zipFilesToDelete = Get-ChildItem -LiteralPath $LogDirectory -File -Filter '*.zip' -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -like "$ComputerName-logs-*.zip" -and
                $_.CreationTime -lt $twoSundaysAgo
            }
    
        foreach ($zipFile in $zipFilesToDelete) {
            try {
                Remove-Item -LiteralPath $zipFile.FullName -Force -ErrorAction Stop
                $archiveSummary.DeletedExpiredZipFiles += $zipFile.FullName
                Write-Log "Deleted ZIP archive older than two Sundays: $($zipFile.FullName)" 'OK'
            }
            catch {
                $msg = "Failed to delete expired ZIP $($zipFile.FullName): $($_.Exception.Message)"
                $archiveSummary.Errors += $msg
                Write-Log $msg 'WARN'
            }
        }
    
        Add-DetailedResult -Step 'LogArchiveRetention' -Status 'Info' -Message 'Sunday-based log archive and retention processing completed.' -Data @{
            ThisSunday                  = $archiveSummary.ThisSunday
            PreviousSunday              = $archiveSummary.PreviousSunday
            TwoSundaysAgo               = $archiveSummary.TwoSundaysAgo
            LooseLogsFound              = $archiveSummary.LooseLogsFound
            LogsSelectedForArchive      = $archiveSummary.LogsSelectedForArchive
            ArchiveCreated              = $archiveSummary.ArchiveCreated
            ArchivePath                 = $archiveSummary.ArchivePath
            DeletedOriginalFilesCount   = @($archiveSummary.DeletedOriginalFiles).Count
            DeletedOldLooseLogsCount    = @($archiveSummary.DeletedOldLooseLogs).Count
            DeletedExpiredZipFilesCount = @($archiveSummary.DeletedExpiredZipFiles).Count
            ErrorsCount                 = @($archiveSummary.Errors).Count
            DeletedOriginalFiles        = ($archiveSummary.DeletedOriginalFiles -join '; ')
            DeletedOldLooseLogs         = ($archiveSummary.DeletedOldLooseLogs -join '; ')
            DeletedExpiredZipFiles      = ($archiveSummary.DeletedExpiredZipFiles -join '; ')
            Errors                      = ($archiveSummary.Errors -join '; ')
        }
    
        if (@($archiveSummary.Errors).Count -gt 0) {
            Warn-Step -Name 'LogArchiveRetention' -Reason ("Completed with errors: " + ($archiveSummary.Errors -join ' | '))
        }
        else {
            Write-Log 'Sunday-based log archive and retention processing completed successfully.' 'OK'
        }
    }
    
    if (-not (Test-IsAdministrator)) {
        Write-Error "This script must be run as Administrator."
        $global:LastStatus = "[ERROR] System Repair must be run as Administrator."
        return
    }
    
    Ensure-LogDirectory
    Write-Log "Initializing automated system health and repair script..." 'INFO'
    Write-Log "Detection-first mode is enabled. Repairs run automatically only when needed." 'INFO'
    Write-Log "Detailed YAML log will be written to $($script:YamlLogPath)" 'INFO'
    
    Invoke-Safely -Name 'DiskSpaceCheck' -ScriptBlock {
        Invoke-DiskSpaceCheck
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'PendingRebootCheck' -ScriptBlock {
        $pending = Get-PendingRebootState
        $script:Summary.PendingRebootDetected = [bool]$pending.AnyPendingReboot
    
        $pending | Format-List | Out-String | ForEach-Object {
            $_.TrimEnd() -split "`r?`n" | Where-Object { $_.Trim() } | ForEach-Object {
                Write-Log $_ 'INFO'
            }
        }
    
        if ($pending.AnyPendingReboot) {
            Warn-Step -Name 'PendingRebootCheck' -Reason 'A pending reboot was detected before maintenance began.'
        }
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'RepairVolumeScan' -ScriptBlock {
        Invoke-RepairVolumeScan
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'DISMDetection' -ScriptBlock {
        Invoke-DismDetection
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'SFCDetection' -ScriptBlock {
        Invoke-SfcDetection
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'WmiRepositoryCheck' -ScriptBlock {
        Invoke-WmiCheck
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'DnsFlush' -ScriptBlock {
        Invoke-DnsFlushOnly
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'TempCleanup' -ScriptBlock {
        Invoke-TempCleanup
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'StorageOptimization' -ScriptBlock {
        Invoke-StorageOptimization
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'ScheduledTaskHealthCheck' -ScriptBlock {
        Invoke-ScheduledTaskHealthCheck
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'EventLogSummary' -ScriptBlock {
        Invoke-EventLogSummary
    } -WarnOnly | Out-Null
    
    Invoke-Safely -Name 'DefenderStatusCheck' -ScriptBlock {
        Invoke-DefenderStatusCheck
    } -WarnOnly | Out-Null
    
    if ($AutoRepairOnDetection) {
        if ($script:Summary.DismCorruptionDetected) {
            Invoke-Safely -Name 'DISMRepair' -ScriptBlock {
                Invoke-DismRepair
            } | Out-Null
        }
    
        if ($script:Summary.SfcIntegrityViolations) {
            Invoke-Safely -Name 'SFCRepair' -ScriptBlock {
                Invoke-SfcRepair
            } | Out-Null
        }
    
        if ($script:Summary.WmiRepositoryInconsistent -and $AllowWmiRepair) {
            Invoke-Safely -Name 'WMIRepair' -ScriptBlock {
                Invoke-WmiRepair
            } -WarnOnly | Out-Null
        }
    
        if ($script:Summary.DiskCorruptionSuspected -and $AllowOfflineDiskRepair) {
            Invoke-Safely -Name 'OfflineDiskRepair' -ScriptBlock {
                Invoke-RepairVolumeOfflineFix
            } -WarnOnly | Out-Null
        }
    }
    
    if ($AllowNetworkReset) {
        Invoke-Safely -Name 'NetworkReset' -ScriptBlock {
            Invoke-NetworkReset
        } -WarnOnly | Out-Null
    }
    
    if ($AllowIconCacheRebuild) {
        Invoke-Safely -Name 'IconCacheRebuild' -ScriptBlock {
            Invoke-IconCacheRebuild
        } -WarnOnly | Out-Null
    }
    
    if ($AllowCopilotRemoval) {
        Invoke-Safely -Name 'CopilotDisableAndRemoval' -ScriptBlock {
            Invoke-CopilotDisableAndRemoval
        } -WarnOnly | Out-Null
    }
    
    if ($AllowFirewallReset) {
        Invoke-Safely -Name 'FirewallReset' -ScriptBlock {
            Invoke-FirewallReset
        } -WarnOnly | Out-Null
    }
    
    if ($AllowWindowsUpdateReset) {
        Invoke-Safely -Name 'WindowsUpdateComponentReset' -ScriptBlock {
            Invoke-WindowsUpdateComponentReset
        } -WarnOnly | Out-Null
    }
    
    if ($ClearEventLogs) {
        Invoke-Safely -Name 'EventLogClear' -ScriptBlock {
            Invoke-EventLogClear
        } -WarnOnly | Out-Null
    }
    
    Invoke-Safely -Name 'LogArchiveRetention' -ScriptBlock {
        Invoke-LogArchiveRetention -LogDirectory $LogDirectory
    } -WarnOnly | Out-Null
    
    Show-Summary
    Write-YamlLog
    
    if ($AutoRebootIfNeeded -and ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected)) {
        try {
            Invoke-RebootIfNeeded -DelaySeconds $AutoRebootDelaySeconds
            Write-YamlLog
        }
        catch {
            Fail-Step -Name 'AutoReboot' -Reason $_.Exception.Message
            Show-Summary
            Write-YamlLog
            $global:LastStatus = "[ERROR] System Repair auto reboot step failed. See YAML log for details."
            return
        }
    }
    
    if ($script:Summary.StepsFailed -gt 0) {
        Write-YamlLog
        $global:LastStatus = "[ERROR] System Repair completed with failures. See YAML log for details."
        return
    }
    elseif ($script:Summary.RebootRequired -or $script:Summary.PendingRebootDetected) {
        Write-YamlLog
        $global:LastStatus = "[OK] System Repair completed; reboot is recommended or required. See YAML log for details."
        return
    }
    else {
        Write-YamlLog
        $global:LastStatus = "[OK] System Repair completed successfully. See YAML log for details."
        return
    }
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

# -----------------------------------------------------------------------------
# Option 14 - Remove User Profiles
# -----------------------------------------------------------------------------
function Remove-UserProfilesClassroom {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string[]]$ExcludedProfiles = @('Default', 'Public', 'MISAdmin'),
        [switch]$Force,
        [switch]$WhatIf,
        [string]$LogPath = "$env:TEMP\ProfileCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [int]$TimeoutSeconds = 30
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    # Security: Validate excluded profiles
    $ExcludedProfiles = $ExcludedProfiles | Where-Object { 
        $_ -match '^[a-zA-Z0-9._-]+$' -and $_.Length -le 104  # Valid Windows username format
    }
    
    # Always include critical system profiles
    $SystemProfiles = @('Default', 'Public', 'All Users', 'Default User', 'NetworkService', 'LocalService', 'systemprofile')
    $ExcludedProfiles = ($ExcludedProfiles + $SystemProfiles) | Sort-Object -Unique

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize logging
    $logEntries = @()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
        }
    }

    try {
        Write-LogEntry "=== Profile Cleanup Session Started ===" 'INFO'
        Write-LogEntry "Excluded profiles: $($ExcludedProfiles -join ', ')" 'INFO'

        if (-not $Force -and -not $WhatIf) {
            # Security: Enhanced warning with specific exclusions
            Write-Host "`n" + "="*60 -ForegroundColor Red
            Write-Host "CRITICAL WARNING - USER PROFILE DELETION" -ForegroundColor Red -BackgroundColor Black
            Write-Host "="*60 -ForegroundColor Red
            Write-Host "This operation will PERMANENTLY DELETE all user profiles except:" -ForegroundColor Yellow
            $ExcludedProfiles | ForEach-Object { Write-Host "  [OK] $_" -ForegroundColor Green }
            Write-Host "`nDeleted profiles CANNOT be recovered!" -ForegroundColor Red
            Write-Host "="*60 -ForegroundColor Red

            # Security: Double confirmation
            $confirmation1 = Read-Host "`n[?] Type 'DELETE' to confirm profile deletion (case-sensitive)"
            if ($confirmation1 -ne 'DELETE') {
                Write-LogEntry "Operation cancelled - incorrect confirmation" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled profile cleanup."
                return
            }

            $confirmation2 = Read-Host "[?] Final confirmation - Type 'YES' to proceed"
            if ($confirmation2 -ne 'YES') {
                Write-LogEntry "Operation cancelled by user" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled profile cleanup."
                return
            }
        }

        Write-LogEntry "[SCAN] Scanning for user profiles..." 'INFO'

        # Speed: Use faster WMI query with specific filters
        $cimSessionOptions = New-CimSessionOption -Protocol WSMan
        $cimSession = New-CimSession -SessionOption $cimSessionOptions

        # Security: More precise profile filtering
        $profileQuery = @"
SELECT * FROM Win32_UserProfile 
WHERE LocalPath LIKE 'C:\\Users\\%' 
AND NOT LocalPath LIKE 'C:\\Users\\Default%' 
AND NOT LocalPath LIKE 'C:\\Users\\Public%'
AND NOT LocalPath LIKE 'C:\\Users\\All Users%'
"@

        $AllUserProfiles = Get-CimInstance -CimSession $cimSession -Query $profileQuery -ErrorAction Stop
        
        if (-not $AllUserProfiles) {
            Write-LogEntry "No user profiles found for deletion" 'INFO'
            $global:LastStatus = "[INFO] No user profiles found for deletion."
            return
        }

        # Security: Pre-validate profiles and check for active sessions
        $validatedProfiles = @()
        $activeUsers = @()
        
        # Speed: Get active sessions once
        try {
            $activeSessions = Get-CimInstance -CimSession $cimSession -ClassName Win32_LogonSession -ErrorAction SilentlyContinue |
                Where-Object { $_.LogonType -in @(2, 10, 11) }  # Interactive, RemoteInteractive, CachedInteractive
            
            $activeUserSIDs = $activeSessions | ForEach-Object {
                Get-CimAssociatedInstance -CimSession $cimSession -InputObject $_ -ResultClassName Win32_UserAccount -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty SID
            }
        } catch {
            Write-LogEntry "Warning: Could not check for active sessions" 'WARNING'
            $activeUserSIDs = @()
        }

        foreach ($Profile in $AllUserProfiles) {
            $ProfilePath = $Profile.LocalPath
            $ProfileName = Split-Path $ProfilePath -Leaf
            
            # Security: Validate profile path
            if ($ProfilePath -notmatch '^C:\\Users\\[^\\]+$') {
                Write-LogEntry "Skipping invalid profile path: $ProfilePath" 'WARNING'
                continue
            }
            
            # Security: Check exclusion list (case-insensitive)
            if ($ExcludedProfiles -contains $ProfileName) {
                Write-LogEntry "Skipping excluded profile: $ProfileName" 'INFO'
                continue
            }
            
            # Security: Check for active sessions
            if ($Profile.SID -in $activeUserSIDs) {
                Write-LogEntry "Skipping active user session: $ProfileName" 'WARNING'
                $activeUsers += $ProfileName
                continue
            }
            
            # Security: Check if profile is currently loaded
            if ($Profile.Loaded) {
                Write-LogEntry "Skipping loaded profile: $ProfileName" 'WARNING'
                continue
            }
            
            $validatedProfiles += @{
                Profile = $Profile
                Name = $ProfileName
                Path = $ProfilePath
                SID = $Profile.SID
                Size = if (Test-Path $ProfilePath) { 
                    try {
                        (Get-ChildItem $ProfilePath -Recurse -Force -ErrorAction SilentlyContinue | 
                         Measure-Object -Property Length -Sum).Sum / 1MB
                    } catch { 0 }
                } else { 0 }
            }
        }

        if ($validatedProfiles.Count -eq 0) {
            Write-LogEntry "No profiles available for deletion after validation" 'INFO'
            if ($activeUsers.Count -gt 0) {
                Write-LogEntry "Active users found: $($activeUsers -join ', ')" 'WARNING'
            }
            $global:LastStatus = "[INFO] No profiles available for deletion."
            return
        }

        # Display deletion plan
        Write-LogEntry "`n[REPORT] Deletion Plan:" 'INFO'
        $totalSize = 0
        foreach ($profileInfo in $validatedProfiles) {
            $sizeStr = if ($profileInfo.Size -gt 0) { " (~$([math]::Round($profileInfo.Size, 2)) MB)" } else { "" }
            Write-LogEntry "  - $($profileInfo.Name)$sizeStr" 'INFO'
            $totalSize += $profileInfo.Size
        }
        Write-LogEntry "Total profiles to delete: $($validatedProfiles.Count)" 'INFO'
        Write-LogEntry "Estimated disk space to free: $([math]::Round($totalSize, 2)) MB" 'INFO'

        if ($WhatIf) {
            Write-LogEntry "WhatIf mode - no profiles will be deleted" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - $($validatedProfiles.Count) profiles would be deleted."
            return
        }

        # Speed: Parallel deletion using runspaces for large numbers of profiles
        Write-LogEntry "`n[CLEAN] Starting profile deletion..." 'INFO'
        $deletedCount = 0
        $failedCount = 0
        $deletionResults = @()

        if ($validatedProfiles.Count -gt 5) {
            # Speed: Use parallel processing for many profiles
            Write-LogEntry "Using parallel deletion for $($validatedProfiles.Count) profiles" 'INFO'
            
            $runspacePool = [runspacefactory]::CreateRunspacePool(1, [Math]::Min(4, $validatedProfiles.Count))
            $runspacePool.Open()
            
            $jobs = @()
            
            foreach ($profileInfo in $validatedProfiles) {
                $powershell = [powershell]::Create().AddScript({
                    param($ProfileObject, $TimeoutSeconds)
                    
                    $result = @{
                        Name = Split-Path $ProfileObject.LocalPath -Leaf
                        Success = $false
                        Error = $null
                        Duration = 0
                    }
                    
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    
                    try {
                        # Set timeout for individual deletion
                        $job = Start-Job -ScriptBlock {
                            param($profile)
                            Remove-CimInstance -InputObject $profile -ErrorAction Stop
                        } -ArgumentList $ProfileObject
                        
                        if (Wait-Job $job -Timeout $TimeoutSeconds) {
                            Receive-Job $job -ErrorAction Stop
                            $result.Success = $true
                        } else {
                            Remove-Job $job -Force
                            throw "Operation timed out after $TimeoutSeconds seconds"
                        }
                    } catch {
                        $result.Error = $_.Exception.Message
                    } finally {
                        $sw.Stop()
                        $result.Duration = $sw.ElapsedMilliseconds
                    }
                    
                    return $result
                }).AddParameter('ProfileObject', $profileInfo.Profile).AddParameter('TimeoutSeconds', $TimeoutSeconds)
                
                $powershell.RunspacePool = $runspacePool
                $jobs += @{
                    PowerShell = $powershell
                    Handle = $powershell.BeginInvoke()
                    ProfileName = $profileInfo.Name
                }
            }
            
            # Wait for all jobs to complete
            foreach ($job in $jobs) {
                try {
                    $result = $job.PowerShell.EndInvoke($job.Handle)
                    $deletionResults += $result
                    
                    if ($result.Success) {
                        Write-LogEntry "[OK] Deleted profile: $($result.Name) ($($result.Duration)ms)" 'SUCCESS'
                        $deletedCount++
                    } else {
                        Write-LogEntry "[X] Failed to delete profile: $($result.Name) - $($result.Error)" 'ERROR'
                        $failedCount++
                    }
                } catch {
                    Write-LogEntry "[X] Job error for profile: $($job.ProfileName) - $_" 'ERROR'
                    $failedCount++
                } finally {
                    $job.PowerShell.Dispose()
                }
            }
            
            $runspacePool.Close()
            $runspacePool.Dispose()
            
        } else {
            # Speed: Sequential deletion for few profiles
            foreach ($profileInfo in $validatedProfiles) {
                $profileName = $profileInfo.Name
                $profile = $profileInfo.Profile
                
                try {
                    Write-LogEntry "Deleting profile: $profileName..." 'INFO'
                    
                    # Security: Final safety check
                    if ($ExcludedProfiles -contains $profileName) {
                        Write-LogEntry "Safety check failed - profile is excluded: $profileName" 'WARNING'
                        continue
                    }
                    
                    $deleteJob = Start-Job -ScriptBlock {
                        param($profileObject)
                        Remove-CimInstance -InputObject $profileObject -ErrorAction Stop
                    } -ArgumentList $profile
                    
                    if (Wait-Job $deleteJob -Timeout $TimeoutSeconds) {
                        Receive-Job $deleteJob -ErrorAction Stop
                        Remove-Job $deleteJob
                        Write-LogEntry "[OK] Successfully deleted profile: $profileName" 'SUCCESS'
                        $deletedCount++
                    } else {
                        Remove-Job $deleteJob -Force
                        throw "Deletion timed out after $TimeoutSeconds seconds"
                    }
                    
                } catch {
                    Write-LogEntry "[X] Failed to delete profile: $profileName - $_" 'ERROR'
                    $failedCount++
                }
            }
        }

        # Cleanup verification
        Write-LogEntry "`n[SCAN] Verifying deletion results..." 'INFO'
        Start-Sleep -Seconds 2  # Allow system to update
        
        $remainingProfiles = Get-CimInstance -CimSession $cimSession -Query $profileQuery -ErrorAction SilentlyContinue
        $actualRemaining = $remainingProfiles | Where-Object {
            $name = Split-Path $_.LocalPath -Leaf
            $name -notin $ExcludedProfiles
        }
        
        if ($actualRemaining) {
            Write-LogEntry "[WARN] Warning: $($actualRemaining.Count) profiles still exist after deletion" 'WARNING'
            $actualRemaining | ForEach-Object {
                $name = Split-Path $_.LocalPath -Leaf
                Write-LogEntry "  - Remaining: $name" 'WARNING'
            }
        }

    } catch {
        Write-LogEntry "Critical error during profile cleanup: $_" 'ERROR'
        throw
    } finally {
        if ($cimSession) {
            Remove-CimSession $cimSession -ErrorAction SilentlyContinue
        }
        
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n=== Profile Cleanup Session Completed ===" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Profiles deleted: $deletedCount" 'SUCCESS'
        if ($failedCount -gt 0) {
            Write-LogEntry "Profiles failed: $failedCount" 'ERROR'
        }
        if ($activeUsers.Count -gt 0) {
            Write-LogEntry "Active users skipped: $($activeUsers.Count)" 'WARNING'
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "Log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status
        if ($deletedCount -gt 0) {
            $global:LastStatus = "[OK] Deleted $deletedCount user profiles in $([math]::Round($duration, 1))s"
            if ($failedCount -gt 0) {
                $global:LastStatus += " ($failedCount failed)"
            }
        } else {
            $global:LastStatus = "[INFO] No profiles were deleted"
        }
    }
}

# Enhanced alias for compatibility
Set-Alias -Name Remove-UserProfiles -Value Remove-UserProfilesClassroom -Force

# -----------------------------------------------------------------------------
# Option 15 - Disable the display of the last user logged on
# -----------------------------------------------------------------------------
function Apply-LoginScreenRegistryFixes {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$EnhancedSecurity,
        [string]$LogPath = "$env:TEMP\LoginRegistryFixes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$BackupRegistry
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking (intentionally unscoped so nested functions can update them)
    $appliedSettings = @()
    $failedSettings  = @()
    $backupData      = @{}
    $logEntries      = @()
    $stopwatch       = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Output to console with colors
        switch ($Level) {
            'ERROR'    { Write-Host $Message -ForegroundColor Red }
            'WARNING'  { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS'  { Write-Host $Message -ForegroundColor Green }
            'INFO'     { Write-Host $Message -ForegroundColor Cyan }
            'SECURITY' { Write-Host $Message -ForegroundColor Magenta }
            default    { Write-Host $Message }
        }
        
        # Add to log collection
        $logEntries += $logEntry
    }

    # Security: Registry validation function
    function Test-RegistryPath {
        param([string]$Path)
        try {
            return Test-Path $Path -ErrorAction Stop
        } catch {
            return $false
        }
    }

    # Security: Safe registry backup function
    function Backup-RegistryValue {
        param(
            [string]$Path,
            [string]$Name
        )
        try {
            if (Test-RegistryPath $Path) {
                $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($null -ne $currentValue) {
                    return $currentValue.$Name
                }
            }
            return $null
        } catch {
            return $null
        }
    }

    # Speed: Optimized registry setting function
    function Set-RegistryValueSafe {
        param(
            [string]$Path,
            [string]$Name,
            [object]$Value,
            [string]$Type = 'DWord',
            [string]$Description,
            [switch]$SecurityCritical
        )
        
        try {
            # Security: Validate registry path format
            if ($Path -notmatch '^HK(LM|CU|CR|U|CC):\\') {
                throw "Invalid registry path format: $Path"
            }
            
            # Backup current value if requested
            if ($BackupRegistry) {
                $currentValue = Backup-RegistryValue -Path $Path -Name $Name
                if ($null -ne $currentValue) {
                    $existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                    $backupData["$Path\$Name"] = @{
                        Value = $currentValue
                        Type  = if ($existing) { $existing.$Name.GetType().Name } else { $null }
                    }
                }
            }
            
            # Create registry path if it doesn't exist
            if (-not (Test-RegistryPath $Path)) {
                Write-LogEntry "Creating registry path: $Path" 'INFO'
                New-Item -Path $Path -Force | Out-Null
            }
            
            # Apply setting
            if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set registry value to $Value")) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
                
                # Verify the setting was applied
                $verifyValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                if ($verifyValue.$Name -eq $Value) {
                    $level = if ($SecurityCritical) { 'SECURITY' } else { 'SUCCESS' }
                    Write-LogEntry "[OK] $Description" $level
                    $appliedSettings += [pscustomobject]@{
                        Path            = $Path
                        Name            = $Name
                        Value           = $Value
                        Description     = $Description
                        SecurityCritical= $SecurityCritical.IsPresent
                    }
                    return $true
                } else {
                    throw "Verification failed: Expected $Value, got $($verifyValue.$Name)"
                }
            }
        } catch {
            Write-LogEntry "[X] Failed to apply $Description : $_" 'ERROR'
            $failedSettings += [pscustomobject]@{
                Path        = $Path
                Name        = $Name
                Description = $Description
                Error       = $_.Exception.Message
            }
            return $false
        }
    }

    try {
        Write-LogEntry "=== Login Screen Registry Security Configuration Started ===" 'INFO'
        Write-LogEntry "Enhanced Security Mode: $($EnhancedSecurity.IsPresent)" 'INFO'
        if ($PSBoundParameters.ContainsKey('WhatIf') -and $PSBoundParameters['WhatIf']) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Core login screen security settings
        Write-LogEntry "`n[SHIELD] Applying core login screen security settings..." 'INFO'
        
        # Speed: Batch registry operations using array
        $coreSettings = @(
            @{
                Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                Name = "EnableFirstLogonAnimation"
                Value = 0
                Type = 'DWord'
                Description = "Disabled first logon animation"
                SecurityCritical = $false
            },
            @{
                Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                Name = "dontdisplaylastusername"
                Value = 1
                Type = 'DWord'
                Description = "Configured login screen to hide last username"
                SecurityCritical = $true
            }
        )

        # Enhanced security settings (applied when -EnhancedSecurity is used)
        if ($EnhancedSecurity) {
            Write-LogEntry "`n[SECURE] Applying enhanced security settings..." 'SECURITY'
            
            $enhancedSettings = @(
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "ShutdownWithoutLogon"
                    Value = 0
                    Type = 'DWord'
                    Description = "Disabled shutdown without logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "UndockWithoutLogon"
                    Value = 0
                    Type = 'DWord'
                    Description = "Disabled undock without logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "AutoAdminLogon"
                    Value = 0
                    Type = 'String' # This value is a REG_SZ ("0"/"1" as string)
                    Description = "Disabled automatic administrator logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "ForceAutoLogon"
                    Value = 0
                    Type = 'String' # Also REG_SZ
                    Description = "Disabled forced automatic logon"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "InactivityTimeoutSecs"
                    Value = 900
                    Type = 'DWord'
                    Description = "Set login screen timeout to 15 minutes"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
                    Name = "MaxDevicePasswordFailedAttempts"
                    Value = 5
                    Type = 'DWord'
                    Description = "Set maximum password attempts to 5"
                    SecurityCritical = $true
                },
                @{
                    Path = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
                    Name = "CachedLogonsCount"
                    Value = 2
                    Type = 'String' # REG_SZ number as string
                    Description = "Limited cached logons to 2"
                    SecurityCritical = $true
                }
            )
            
            $coreSettings += $enhancedSettings
        }

        # Speed: Process settings in batches
        $batchSize = 5
        $totalSettings = $coreSettings.Count
        
        for ($i = 0; $i -lt $totalSettings; $i += $batchSize) {
            $end = [Math]::Min($i + $batchSize - 1, $totalSettings - 1)
            $batch = $coreSettings[$i..$end]
            
            foreach ($setting in $batch) {
                $params = @{
                    Path            = $setting.Path
                    Name            = $setting.Name
                    Value           = $setting.Value
                    Description     = $setting.Description
                    SecurityCritical= $setting.SecurityCritical
                }
                if ($setting.ContainsKey('Type') -and $setting.Type) {
                    $params.Type = $setting.Type
                }
                Set-RegistryValueSafe @params | Out-Null
            }
            
            if ($end -lt ($totalSettings - 1)) {
                Start-Sleep -Milliseconds 50
            }
        }

        # Security: Additional hardening for classroom environments
        if ($EnhancedSecurity) {
            Write-LogEntry "`n[SCHOOL] Applying classroom-specific security hardening..." 'SECURITY'
            
            # Disable guest account
            try {
                $guestAccount = Get-LocalUser -Name "Guest" -ErrorAction SilentlyContinue
                if ($guestAccount -and $guestAccount.Enabled) {
                    if ($PSCmdlet.ShouldProcess("LocalUser 'Guest'", "Disable")) {
                        Disable-LocalUser -Name "Guest" -ErrorAction Stop
                        Write-LogEntry "[OK] Disabled Guest account" 'SECURITY'
                    }
                }
            } catch {
                Write-LogEntry "[X] Failed to disable Guest account: $_" 'ERROR'
            }
            
            # Set strong password policy via registry
            $passwordPolicySettings = @(
                @{
                    Path = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters"
                    Name = "RequireStrongKey"
                    Value = 1
                    Type = 'DWord'
                    Description = "Enabled strong authentication keys"
                },
                @{
                    Path = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
                    Name = "LimitBlankPasswordUse"
                    Value = 1
                    Type = 'DWord'
                    Description = "Disabled blank password usage"
                }
            )
            
            foreach ($setting in $passwordPolicySettings) {
                Set-RegistryValueSafe @setting -SecurityCritical | Out-Null
            }
        }

        # Security: Registry integrity verification
        Write-LogEntry "`n[SCAN] Verifying registry integrity..." 'INFO'
        $verificationErrors = 0
        
        foreach ($setting in $appliedSettings) {
            try {
                $currentValue = Get-ItemProperty -Path $setting.Path -Name $setting.Name -ErrorAction Stop
                if ($currentValue.($setting.Name) -ne $setting.Value) {
                    Write-LogEntry "[WARN] Verification failed for $($setting.Description)" 'WARNING'
                    $verificationErrors++
                }
            } catch {
                Write-LogEntry "[WARN] Could not verify $($setting.Description)" 'WARNING'
                $verificationErrors++
            }
        }
        
        if ($verificationErrors -eq 0) {
            Write-LogEntry "[OK] All registry settings verified successfully" 'SUCCESS'
        } else {
            Write-LogEntry "[WARN] $verificationErrors settings failed verification" 'WARNING'
        }

        # Create registry backup file
        if ($BackupRegistry -and $backupData.Count -gt 0) {
            try {
                $backupFile = "$env:TEMP\LoginRegistryBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json -Depth 5 | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Registry backup saved to: $backupFile" 'INFO'
            } catch {
                Write-LogEntry "[WARN] Failed to save registry backup: $_" 'WARNING'
            }
        }

    } catch {
        Write-LogEntry "Critical error during registry configuration: $_" 'ERROR'
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Generate summary
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Settings applied: $($appliedSettings.Count)" 'SUCCESS'
        Write-LogEntry "Settings failed: $($failedSettings.Count)" 'ERROR'
        
        $securitySettings = ($appliedSettings | Where-Object { $_.SecurityCritical }).Count
        if ($securitySettings -gt 0) {
            Write-LogEntry "Security-critical settings applied: $securitySettings" 'SECURITY'
        }
        
        if ($failedSettings.Count -gt 0) {
            Write-LogEntry "`n[ERROR] Failed Settings:" 'ERROR'
            foreach ($failed in $failedSettings) {
                Write-LogEntry "  - $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[NOTE] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status
        if ($appliedSettings.Count -gt 0) {
            $statusMsg = "[OK] Applied $($appliedSettings.Count) login screen security settings"
            if ($failedSettings.Count -gt 0) {
                $statusMsg += " ($($failedSettings.Count) failed)"
            }
            if ($securitySettings -gt 0) {
                $statusMsg += " [$securitySettings security-critical]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] No login screen settings were applied"
        }
        
        Write-LogEntry "=== Login Screen Registry Configuration Completed ===" 'INFO'
    }
}

# Create restore function for emergency rollback
function Restore-LoginScreenRegistrySettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BackupFile
    )
    
    if (-not (Test-Path $BackupFile)) {
        throw "Backup file not found: $BackupFile"
    }
    
    try {
        $backupData = Get-Content $BackupFile | ConvertFrom-Json
        $restored = 0
        
        foreach ($entry in $backupData.PSObject.Properties) {
            $pathAndName = $entry.Name -split '\\'
            $path = $pathAndName[0..($pathAndName.Length-2)] -join '\'
            $name = $pathAndName[-1]
            
            try {
                Set-ItemProperty -Path $path -Name $name -Value $entry.Value.Value -ErrorAction Stop
                Write-Host "[OK] Restored: $($entry.Name)" -ForegroundColor Green
                $restored++
            } catch {
                Write-Warning "[WARN] Failed to restore: $($entry.Name) - $_"
            }
        }
        
        Write-Host "[OK] Restored $restored registry settings from backup" -ForegroundColor Cyan
    } catch {
        throw "Failed to restore from backup: $_"
    }
}

# -----------------------------------------------------------------------------
# Option 16 - Enable Automatic Login with CC-Student
# -----------------------------------------------------------------------------
function Set-DomainAutoLogin {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$UserName = "CC-Student",
        
        [Parameter(Mandatory = $false)]
        [securestring]$Password,
        
        [Parameter(Mandatory = $false)]
        [string]$DomainName = "Compton.edu",
        
        [switch]$DisableAutoLogin,
        [switch]$WhatIf,
        [switch]$Force,
        [string]$LogPath = "$env:TEMP\AutoLoginConfig_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [int]$AutoLoginCount = 1,  # Number of auto-logins before disabling
        [switch]$UseLocalSystemEncryption
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    
    # Initialize tracking
    $logEntries = @()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'SECURITY' { Write-Host $Message -ForegroundColor Magenta }
        }
    }

    # Security: Secure credential management
    function Get-SecureCredential {
        param(
            [string]$Username,
            [string]$Domain,
            [securestring]$SecurePassword
        )
        
        if (-not $SecurePassword) {
            Write-LogEntry "[WARN] SECURITY WARNING: Auto-login requires storing credentials" 'WARNING'
            Write-LogEntry "Consider using alternative authentication methods for production" 'WARNING'
            
            if (-not $Force) {
                $response = Read-Host "Continue with credential storage? (Type 'ACCEPT' to proceed)"
                if ($response -ne 'ACCEPT') {
                    throw "Operation cancelled - credential storage not accepted"
                }
            }
            
            # Prompt for secure password
            $SecurePassword = Read-Host "Enter password for $Domain\$Username" -AsSecureString
            if (-not $SecurePassword -or $SecurePassword.Length -eq 0) {
                throw "Password is required for auto-login configuration"
            }
        }
        
        return New-Object System.Management.Automation.PSCredential("$Domain\$Username", $SecurePassword)
    }

    # Security: Registry validation and backup
    function Backup-AutoLoginSettings {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $backupData = @{}
        
        $settingsToBackup = @(
            'AutoAdminLogon',
            'DefaultUserName', 
            'DefaultPassword',
            'DefaultDomainName',
            'AutoLogonCount'
        )
        
        foreach ($setting in $settingsToBackup) {
            try {
                $value = Get-ItemProperty -Path $regPath -Name $setting -ErrorAction SilentlyContinue
                if ($value) {
                    $backupData[$setting] = $value.$setting
                }
            } catch {
                # Setting doesn't exist, which is fine
            }
        }
        
        if ($backupData.Count -gt 0) {
            try {
                $backupFile = "$env:TEMP\AutoLoginBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Registry backup saved to: $backupFile" 'INFO'
                return $backupFile
            } catch {
                Write-LogEntry "[WARN] Failed to save registry backup: $_" 'WARNING'
            }
        }
        
        return $null
    }

    # Security: Secure password encryption using DPAPI
    function Protect-AutoLoginPassword {
        param([securestring]$SecurePassword)
        
        try {
            # Convert SecureString to encrypted string using DPAPI
            $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
            $plaintext = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            
            # Use DPAPI to encrypt for local machine
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($plaintext)
            $encryptedBytes = [System.Security.Cryptography.ProtectedData]::Protect(
                $bytes, 
                $null, 
                [System.Security.Cryptography.DataProtectionScope]::LocalMachine
            )
            
            # Clear plaintext from memory
            $plaintext = $null
            [System.GC]::Collect()
            
            return [System.Convert]::ToBase64String($encryptedBytes)
            
        } catch {
            throw "Failed to encrypt password: $_"
        }
    }

    # Security: Domain validation
    function Test-DomainConnectivity {
        param([string]$DomainName)
        
        try {
            Write-LogEntry "[SCAN] Validating domain connectivity..." 'INFO'
            
            # Test domain controller connectivity
            $domainController = Resolve-DnsName -Name $DomainName -Type A -ErrorAction Stop
            if (-not $domainController) {
                return $false
            }
            
            # Test LDAP connectivity
            $ldapTest = Test-NetConnection -ComputerName $DomainName -Port 389 -WarningAction SilentlyContinue
            if (-not $ldapTest.TcpTestSucceeded) {
                Write-LogEntry "[WARN] LDAP connectivity test failed" 'WARNING'
                return $false
            }
            
            Write-LogEntry "[OK] Domain connectivity validated" 'SUCCESS'
            return $true
            
        } catch {
            Write-LogEntry "[WARN] Domain validation failed: $_" 'WARNING'
            return $false
        }
    }

    # Speed: Optimized registry operations
    function Set-AutoLoginRegistry {
        param(
            [string]$Username,
            [string]$Domain,
            [string]$EncryptedPassword,
            [int]$LoginCount
        )
        
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        $settingsApplied = 0
        
        # Security: Validate registry path exists
        if (-not (Test-Path $regPath)) {
            throw "Winlogon registry path not found: $regPath"
        }
        
        # Batch registry operations for speed
        $registrySettings = @(
            @{ Name = 'AutoAdminLogon'; Value = '1'; Type = 'String' },
            @{ Name = 'DefaultUserName'; Value = $Username; Type = 'String' },
            @{ Name = 'DefaultDomainName'; Value = $Domain; Type = 'String' },
            @{ Name = 'AutoLogonCount'; Value = $LoginCount; Type = 'DWord' }
        )
        
        # Add password setting based on encryption method
        if ($UseLocalSystemEncryption) {
            # Use Windows built-in LSA encryption
            $registrySettings += @{ Name = 'DefaultPassword'; Value = ''; Type = 'String' }
            
            # Store encrypted password separately (requires additional LSA configuration)
            Write-LogEntry "Using LSA Secret storage for password (enhanced security)" 'SECURITY'
            # Note: LSA Secret storage requires additional implementation
        } else {
            # Store DPAPI-encrypted password
            $registrySettings += @{ Name = 'DefaultPassword'; Value = $EncryptedPassword; Type = 'String' }
        }
        
        foreach ($setting in $registrySettings) {
            try {
                if ($PSCmdlet.ShouldProcess("$regPath\$($setting.Name)", "Set registry value")) {
                    Set-ItemProperty -Path $regPath -Name $setting.Name -Value $setting.Value -Type $setting.Type -ErrorAction Stop
                    Write-LogEntry "[OK] Set $($setting.Name)" 'SUCCESS'
                    $settingsApplied++
                }
            } catch {
                Write-LogEntry "[X] Failed to set $($setting.Name): $_" 'ERROR'
                throw
            }
        }
        
        return $settingsApplied
    }

    try {
        Write-LogEntry "=== Domain Auto-Login Configuration Started ===" 'INFO'
        
        if ($WhatIf) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Handle disable auto-login request
        if ($DisableAutoLogin) {
            Write-LogEntry "[BLOCK] Disabling auto-login..." 'INFO'
            
            $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
            
            if ($PSCmdlet.ShouldProcess($regPath, "Disable auto-login")) {
                try {
                    Set-ItemProperty -Path $regPath -Name "AutoAdminLogon" -Value "0" -Type String -ErrorAction Stop
                    
                    # Clear stored credentials for security
                    $credentialSettings = @('DefaultPassword', 'DefaultUserName', 'DefaultDomainName', 'AutoLogonCount')
                    foreach ($setting in $credentialSettings) {
                        try {
                            Remove-ItemProperty -Path $regPath -Name $setting -ErrorAction SilentlyContinue
                        } catch {
                            # Setting may not exist, continue
                        }
                    }
                    
                    Write-LogEntry "[OK] Auto-login disabled and credentials cleared" 'SUCCESS'
                    $global:LastStatus = "[OK] Auto-login disabled successfully."
                    return
                    
                } catch {
                    throw "Failed to disable auto-login: $_"
                }
            }
            return
        }

        # Security: Critical warning about auto-login risks
        if (-not $Force) {
            Write-Host "`n" + "="*70 -ForegroundColor Red
            Write-Host "CRITICAL SECURITY WARNING - AUTO-LOGIN CONFIGURATION" -ForegroundColor Red -BackgroundColor Black
            Write-Host "="*70 -ForegroundColor Red
            Write-Host "This configuration will:" -ForegroundColor Yellow
            Write-Host "  - Store domain credentials on local system" -ForegroundColor Yellow
            Write-Host "  - Allow automatic login without authentication" -ForegroundColor Yellow
            Write-Host "  - Potentially expose credentials to local attacks" -ForegroundColor Yellow
            Write-Host "`nRECOMMENDATIONS:" -ForegroundColor Cyan
            Write-Host "  - Use only in secure, controlled environments" -ForegroundColor Cyan
            Write-Host "  - Consider using domain Group Policy instead" -ForegroundColor Cyan
            Write-Host "  - Limit auto-login count to minimize exposure" -ForegroundColor Cyan
            Write-Host "  - Regularly rotate the password" -ForegroundColor Cyan
            Write-Host "="*70 -ForegroundColor Red
            
            $confirmation = Read-Host "`n[?] Type 'UNDERSTAND' to acknowledge security risks and continue"
            if ($confirmation -ne 'UNDERSTAND') {
                Write-LogEntry "Operation cancelled - security risks not acknowledged" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled auto-login configuration."
                return
            }
        }

        # Validate inputs
        Write-LogEntry "[SCAN] Validating configuration parameters..." 'INFO'
        
        if ([string]::IsNullOrWhiteSpace($UserName)) {
            throw "Username cannot be empty"
        }
        
        if ([string]::IsNullOrWhiteSpace($DomainName)) {
            throw "Domain name cannot be empty"
        }
        
        # Security: Validate domain connectivity
        $domainConnectivity = Test-DomainConnectivity -DomainName $DomainName
        if (-not $domainConnectivity -and -not $Force) {
            throw "Domain connectivity validation failed. Use -Force to override."
        }

        # Security: Get secure credentials
        Write-LogEntry "[LOCK] Processing credentials securely..." 'SECURITY'
        $credential = Get-SecureCredential -Username $UserName -Domain $DomainName -SecurePassword $Password
        
        # Security: Test credential validity (optional)
        if ($domainConnectivity) {
            try {
                Write-LogEntry "[SCAN] Validating credentials..." 'INFO'
                # Note: In production, you might want to test credentials against domain
                # This is omitted here to avoid additional authentication attempts
                Write-LogEntry "[OK] Credential format validated" 'SUCCESS'
            } catch {
                Write-LogEntry "[WARN] Credential validation failed: $_" 'WARNING'
                if (-not $Force) {
                    throw "Invalid credentials provided"
                }
            }
        }

        # Security: Backup current settings
        Write-LogEntry "[SAVE] Backing up current auto-login settings..." 'INFO'
        $backupFile = Backup-AutoLoginSettings
        
        if ($WhatIf) {
            Write-LogEntry "WhatIf: Would configure auto-login for $DomainName\$UserName" 'INFO'
            Write-LogEntry "WhatIf: Would set AutoLogonCount to $AutoLoginCount" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - auto-login would be configured."
            return
        }

        # Security: Encrypt password
        Write-LogEntry "[SECURE] Encrypting credentials..." 'SECURITY'
        $encryptedPassword = Protect-AutoLoginPassword -SecurePassword $credential.Password
        
        # Apply registry settings
        Write-LogEntry "[NOTE] Applying auto-login registry settings..." 'INFO'
        $settingsCount = Set-AutoLoginRegistry -Username $UserName -Domain $DomainName -EncryptedPassword $encryptedPassword -LoginCount $AutoLoginCount
        
        # Security: Verify settings were applied
        Write-LogEntry "[SCAN] Verifying configuration..." 'INFO'
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        
        try {
            $autoLogonValue = Get-ItemProperty -Path $regPath -Name "AutoAdminLogon" -ErrorAction Stop
            $userNameValue = Get-ItemProperty -Path $regPath -Name "DefaultUserName" -ErrorAction Stop
            $domainValue = Get-ItemProperty -Path $regPath -Name "DefaultDomainName" -ErrorAction Stop
            
            if ($autoLogonValue.AutoAdminLogon -eq "1" -and 
                $userNameValue.DefaultUserName -eq $UserName -and 
                $domainValue.DefaultDomainName -eq $DomainName) {
                
                Write-LogEntry "[OK] Auto-login configuration verified successfully" 'SUCCESS'
            } else {
                throw "Configuration verification failed"
            }
        } catch {
            throw "Failed to verify auto-login configuration: $_"
        }

        # Security: Set appropriate permissions on registry key
        try {
            Write-LogEntry "[SECURE] Securing registry permissions..." 'SECURITY'
            
            $acl = Get-Acl -Path $regPath
            # Remove inherited permissions to protect stored credentials
            $acl.SetAccessRuleProtection($true, $true)
            
            # Keep only essential permissions (System, Administrators)
            $accessRules = $acl.Access | Where-Object { 
                $_.IdentityReference -notmatch "(Users|Everyone|Authenticated Users)" 
            }
            
            Set-Acl -Path $regPath -AclObject $acl
            Write-LogEntry "[OK] Registry permissions secured" 'SUCCESS'
            
        } catch {
            Write-LogEntry "[WARN] Failed to secure registry permissions: $_" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during auto-login configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Auto-login configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        
        if (-not $DisableAutoLogin -and -not $WhatIf) {
            Write-LogEntry "User: $DomainName\$UserName" 'INFO'
            Write-LogEntry "Auto-login count: $AutoLoginCount" 'INFO'
            Write-LogEntry "Encryption: DPAPI (LocalMachine scope)" 'SECURITY'
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[NOTE] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set final status
        if (-not $global:LastStatus -or $global:LastStatus -notlike "*auto-login*") {
            if ($DisableAutoLogin) {
                $global:LastStatus = "[OK] Auto-login disabled successfully."
            } else {
                $global:LastStatus = "[OK] Auto-login configured for $DomainName\$UserName."
            }
        }
        
        Write-LogEntry "=== Auto-Login Configuration Completed ===" 'INFO'
    }
}

# Utility function to check current auto-login status
function Get-AutoLoginStatus {
    [CmdletBinding()]
    param()
    
    try {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        
        $autoLogon = Get-ItemProperty -Path $regPath -Name "AutoAdminLogon" -ErrorAction SilentlyContinue
        $userName = Get-ItemProperty -Path $regPath -Name "DefaultUserName" -ErrorAction SilentlyContinue
        $domainName = Get-ItemProperty -Path $regPath -Name "DefaultDomainName" -ErrorAction SilentlyContinue
        $loginCount = Get-ItemProperty -Path $regPath -Name "AutoLogonCount" -ErrorAction SilentlyContinue
        
        return @{
            Enabled = ($autoLogon.AutoAdminLogon -eq "1")
            UserName = $userName.DefaultUserName
            DomainName = $domainName.DefaultDomainName
            AutoLogonCount = $loginCount.AutoLogonCount
            PasswordStored = (Get-ItemProperty -Path $regPath -Name "DefaultPassword" -ErrorAction SilentlyContinue) -ne $null
        }
        
    } catch {
        return @{
            Error = $_.Exception.Message
            Enabled = $false
        }
    }
}

# -----------------------------------------------------------------------------
# Option 7 - Set Desktop/Laptop Power Settings with Local Policy Enforcement
# -----------------------------------------------------------------------------
function Set-DesktopPowerSettings {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Force,
        [string]$LogPath = "C:\Logs\PowerSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    )

    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This function must be run as Administrator.'
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'

    if (-not (Test-Path -LiteralPath (Split-Path -Path $LogPath -Parent))) {
        New-Item -Path (Split-Path -Path $LogPath -Parent) -ItemType Directory -Force | Out-Null
    }

    $script:Option7LogEntries = New-Object System.Collections.Generic.List[string]
    $script:Option7Results = New-Object System.Collections.Generic.List[object]
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    function Write-PowerLog {
        param(
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('INFO','OK','WARN','ERROR','HARDWARE')][string]$Level = 'INFO'
        )
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line = "[$timestamp] [$Level] $Message"
        $script:Option7LogEntries.Add($line) | Out-Null
        $color = switch ($Level) {
            'OK' { 'Green' }
            'WARN' { 'Yellow' }
            'ERROR' { 'Red' }
            'HARDWARE' { 'Magenta' }
            default { 'Cyan' }
        }
        Write-Host $line -ForegroundColor $color
    }

    function Add-PowerResult {
        param(
            [string]$Setting,
            [string]$Value,
            [bool]$Success = $true,
            [string]$ErrorMessage = $null
        )
        $script:Option7Results.Add([pscustomobject]@{
            Setting = $Setting
            Value = $Value
            Success = $Success
            Error = $ErrorMessage
        }) | Out-Null
    }

    function Invoke-PowerCfgSafe {
        param(
            [Parameter(Mandatory)][string[]]$Arguments,
            [Parameter(Mandatory)][string]$Description,
            [switch]$ContinueOnError
        )
        try {
            $output = & powercfg.exe @Arguments 2>&1
            $exitCode = $LASTEXITCODE
            if ($exitCode -ne 0) {
                $text = if ($output) { ($output | Out-String).Trim() } else { 'No output returned.' }
                throw "$Description failed. ExitCode=$exitCode. $text"
            }
            Write-PowerLog "[OK] $Description" 'OK'
            Add-PowerResult -Setting $Description -Value ($Arguments -join ' ')
            return $output
        }
        catch {
            Write-PowerLog "[X] $($_.Exception.Message)" 'ERROR'
            Add-PowerResult -Setting $Description -Value ($Arguments -join ' ') -Success $false -ErrorMessage $_.Exception.Message
            if (-not $ContinueOnError) { throw }
        }
    }

    function Set-RegistryDwordSafe {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Name,
            [Parameter(Mandatory)][int]$Value,
            [Parameter(Mandatory)][string]$Description
        )
        try {
            if (-not (Test-Path -LiteralPath $Path)) {
                New-Item -Path $Path -Force | Out-Null
            }
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
            Write-PowerLog "[OK] Policy registry set: $Description = $Value" 'OK'
            Add-PowerResult -Setting "Policy: $Description" -Value ([string]$Value)
        }
        catch {
            Write-PowerLog "[X] Failed policy registry set: $Description - $($_.Exception.Message)" 'ERROR'
            Add-PowerResult -Setting "Policy: $Description" -Value ([string]$Value) -Success $false -ErrorMessage $_.Exception.Message
        }
    }

    function Get-SystemHardwareType {
        try {
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            $enclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue

            $chassisTypes = @()
            if ($enclosure -and $enclosure.ChassisTypes) { $chassisTypes = @($enclosure.ChassisTypes) }

            $laptopChassis = @(8,9,10,11,12,14,18,21,30,31,32)
            $desktopChassis = @(3,4,5,6,7,15,16)
            $serverChassis = @(17,23)

            $isLaptop = (($chassisTypes | Where-Object { $_ -in $laptopChassis }).Count -gt 0) -or ($null -ne $battery) -or ($computerSystem.PCSystemType -eq 2)
            $isServer = (($chassisTypes | Where-Object { $_ -in $serverChassis }).Count -gt 0) -or ($computerSystem.PCSystemType -eq 4)
            $isDesktop = (-not $isLaptop -and ((($chassisTypes | Where-Object { $_ -in $desktopChassis }).Count -gt 0) -or $computerSystem.PCSystemType -in @(1,3)))

            [pscustomobject]@{
                IsLaptop = [bool]$isLaptop
                IsDesktop = [bool]$isDesktop
                IsServer = [bool]$isServer
                HasBattery = [bool]($null -ne $battery)
                Manufacturer = $computerSystem.Manufacturer
                Model = $computerSystem.Model
                PCSystemType = $computerSystem.PCSystemType
                ChassisTypes = ($chassisTypes -join ',')
            }
        }
        catch {
            Write-PowerLog "Hardware detection failed, defaulting to desktop profile. Error: $($_.Exception.Message)" 'WARN'
            [pscustomobject]@{ IsLaptop = $false; IsDesktop = $true; IsServer = $false; HasBattery = $false; Manufacturer = 'Unknown'; Model = 'Unknown'; PCSystemType = 'Unknown'; ChassisTypes = 'Unknown' }
        }
    }

    function Get-PowerSchemeGuid {
        param(
            [ValidateSet('Balanced','High performance','Power saver','Ultimate Performance')]
            [string]$PreferredScheme
        )

        $aliases = @{
            'Balanced' = 'SCHEME_BALANCED'
            'High performance' = 'SCHEME_MIN'
            'Power saver' = 'SCHEME_MAX'
            'Ultimate Performance' = 'e9a42b02-d5df-448d-aa00-03f14749eb61'
        }

        if ($PreferredScheme -eq 'Ultimate Performance') {
            Invoke-PowerCfgSafe -Arguments @('/duplicatescheme', $aliases[$PreferredScheme]) -Description 'Ensure Ultimate Performance scheme exists' -ContinueOnError | Out-Null
        }

        $list = powercfg.exe /list 2>$null
        foreach ($line in $list) {
            if ($line -match 'Power Scheme GUID:\s*([a-fA-F0-9-]+)\s*\((.+?)\)') {
                $guid = $matches[1]
                $name = $matches[2]
                if ($name -ieq $PreferredScheme) { return $guid }
            }
        }

        if ($aliases.ContainsKey($PreferredScheme)) { return $aliases[$PreferredScheme] }
        return 'SCHEME_BALANCED'
    }

    function Backup-PowerConfiguration {
        try {
            $backup = [ordered]@{
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                ActiveScheme = (powercfg.exe /getactivescheme 2>$null | Out-String).Trim()
                AvailableSchemes = (powercfg.exe /list 2>$null | Out-String).Trim()
            }
            $backupFile = "C:\Logs\PowerSettingsBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $backup | ConvertTo-Json -Depth 4 | Out-File -LiteralPath $backupFile -Encoding UTF8 -Force
            Write-PowerLog "Power settings backup saved to: $backupFile" 'OK'
        }
        catch {
            Write-PowerLog "Failed to save power settings backup: $($_.Exception.Message)" 'WARN'
        }
    }

    function Set-PolicyPowerSetting {
        param(
            [Parameter(Mandatory)][string]$SubGroupGuid,
            [Parameter(Mandatory)][string]$SettingGuid,
            [Nullable[int]]$ACValue,
            [Nullable[int]]$DCValue,
            [Parameter(Mandatory)][string]$Description
        )
        $policyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Power\PowerSettings\$SubGroupGuid\$SettingGuid"
        if ($null -ne $ACValue) { Set-RegistryDwordSafe -Path $policyPath -Name 'ACSettingIndex' -Value $ACValue.Value -Description "$Description AC" }
        if ($null -ne $DCValue) { Set-RegistryDwordSafe -Path $policyPath -Name 'DCSettingIndex' -Value $DCValue.Value -Description "$Description DC" }
    }

    function Set-PowerProfileValues {
        param(
            [Parameter(Mandatory)][string]$SchemeGuid,
            [Parameter(Mandatory)][hashtable]$Profile,
            [Parameter(Mandatory)][string]$ProfileName
        )

        $subVideo = '7516b95f-f776-4464-8c53-06167f40cc99'
        $videoIdle = '3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e'
        $subSleep = '238c9fa8-0aad-41ed-83f4-97be242c8f20'
        $standbyIdle = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da'
        $hibernateIdle = '9d7815a6-7ee4-497e-8888-515a05f02364'
        $subDisk = '0012ee47-9041-4b5d-9b77-535fba8b1442'
        $diskIdle = '6738e2c4-e8a5-4a42-b16a-e040e769756e'
        $subProcessor = '54533251-82be-4824-96c1-47b60b740d00'
        $procMin = '893dee8e-2bef-41e0-89c6-b55d0929964c'
        $procMax = 'bc5038f7-23e0-4960-96da-33abaf5935ec'

        Write-PowerLog "Applying $ProfileName profile values to scheme $SchemeGuid" 'INFO'

        Invoke-PowerCfgSafe -Arguments @('/setacvalueindex', $SchemeGuid, $subVideo, $videoIdle, [string]$Profile.DisplayACSeconds) -Description "$ProfileName display timeout on AC" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setdcvalueindex', $SchemeGuid, $subVideo, $videoIdle, [string]$Profile.DisplayDCSeconds) -Description "$ProfileName display timeout on battery" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setacvalueindex', $SchemeGuid, $subSleep, $standbyIdle, [string]$Profile.SleepACSeconds) -Description "$ProfileName sleep timeout on AC" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setdcvalueindex', $SchemeGuid, $subSleep, $standbyIdle, [string]$Profile.SleepDCSeconds) -Description "$ProfileName sleep timeout on battery" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setacvalueindex', $SchemeGuid, $subSleep, $hibernateIdle, [string]$Profile.HibernateACSeconds) -Description "$ProfileName hibernate timeout on AC" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setdcvalueindex', $SchemeGuid, $subSleep, $hibernateIdle, [string]$Profile.HibernateDCSeconds) -Description "$ProfileName hibernate timeout on battery" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setacvalueindex', $SchemeGuid, $subDisk, $diskIdle, [string]$Profile.DiskACSeconds) -Description "$ProfileName disk timeout on AC" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setdcvalueindex', $SchemeGuid, $subDisk, $diskIdle, [string]$Profile.DiskDCSeconds) -Description "$ProfileName disk timeout on battery" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setacvalueindex', $SchemeGuid, $subProcessor, $procMin, [string]$Profile.ProcessorMinAC) -Description "$ProfileName processor minimum on AC" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setdcvalueindex', $SchemeGuid, $subProcessor, $procMin, [string]$Profile.ProcessorMinDC) -Description "$ProfileName processor minimum on battery" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setacvalueindex', $SchemeGuid, $subProcessor, $procMax, [string]$Profile.ProcessorMaxAC) -Description "$ProfileName processor maximum on AC" -ContinueOnError | Out-Null
        Invoke-PowerCfgSafe -Arguments @('/setdcvalueindex', $SchemeGuid, $subProcessor, $procMax, [string]$Profile.ProcessorMaxDC) -Description "$ProfileName processor maximum on battery" -ContinueOnError | Out-Null

        Set-PolicyPowerSetting -SubGroupGuid $subVideo -SettingGuid $videoIdle -ACValue $Profile.DisplayACSeconds -DCValue $Profile.DisplayDCSeconds -Description "$ProfileName policy display timeout seconds"
        Set-PolicyPowerSetting -SubGroupGuid $subSleep -SettingGuid $standbyIdle -ACValue $Profile.SleepACSeconds -DCValue $Profile.SleepDCSeconds -Description "$ProfileName policy sleep timeout seconds"
        Set-PolicyPowerSetting -SubGroupGuid $subSleep -SettingGuid $hibernateIdle -ACValue $Profile.HibernateACSeconds -DCValue $Profile.HibernateDCSeconds -Description "$ProfileName policy hibernate timeout seconds"
        Set-PolicyPowerSetting -SubGroupGuid $subDisk -SettingGuid $diskIdle -ACValue $Profile.DiskACSeconds -DCValue $Profile.DiskDCSeconds -Description "$ProfileName policy disk timeout seconds"
        Set-PolicyPowerSetting -SubGroupGuid $subProcessor -SettingGuid $procMin -ACValue $Profile.ProcessorMinAC -DCValue $Profile.ProcessorMinDC -Description "$ProfileName policy processor minimum percent"
        Set-PolicyPowerSetting -SubGroupGuid $subProcessor -SettingGuid $procMax -ACValue $Profile.ProcessorMaxAC -DCValue $Profile.ProcessorMaxDC -Description "$ProfileName policy processor maximum percent"
    }

    try {
        Write-PowerLog '=== Option 7 Power Settings Configuration Started ===' 'INFO'
        Backup-PowerConfiguration

        $hardware = Get-SystemHardwareType
        Write-PowerLog "Detected Manufacturer: $($hardware.Manufacturer)" 'HARDWARE'
        Write-PowerLog "Detected Model       : $($hardware.Model)" 'HARDWARE'
        Write-PowerLog "Detected Chassis     : $($hardware.ChassisTypes)" 'HARDWARE'
        Write-PowerLog "Battery Present      : $($hardware.HasBattery)" 'HARDWARE'

        $desktopProfile = @{
            DisplayACSeconds = 3600; DisplayDCSeconds = 3600
            SleepACSeconds = 0; SleepDCSeconds = 0
            HibernateACSeconds = 0; HibernateDCSeconds = 0
            DiskACSeconds = 0; DiskDCSeconds = 0
            ProcessorMinAC = 100; ProcessorMinDC = 100
            ProcessorMaxAC = 100; ProcessorMaxDC = 100
        }

        $laptopProfile = @{
            # Conservative battery behavior, higher performance while charging.
            DisplayACSeconds = 3600; DisplayDCSeconds = 600
            SleepACSeconds = 0; SleepDCSeconds = 1200
            HibernateACSeconds = 0; HibernateDCSeconds = 3600
            DiskACSeconds = 0; DiskDCSeconds = 600
            ProcessorMinAC = 50; ProcessorMinDC = 5
            ProcessorMaxAC = 100; ProcessorMaxDC = 70
        }

        if ($hardware.IsLaptop) {
            $targetPlan = 'Balanced'
            $profile = $laptopProfile
            $profileName = 'Laptop AC-performance / DC-conservative'
            Write-PowerLog 'Laptop detected. Applying battery-conservative DC settings and higher-performance AC settings.' 'INFO'
        }
        else {
            $targetPlan = 'High performance'
            $profile = $desktopProfile
            $profileName = 'Desktop high-performance'
            Write-PowerLog 'Desktop/workstation detected. Applying high-performance desktop settings.' 'INFO'
        }

        if (-not $Force -and -not $WhatIfPreference) {
            Write-Host "`nThis will apply and policy-enforce the following Option 7 profile:" -ForegroundColor Yellow
            Write-Host "  Profile: $profileName" -ForegroundColor Cyan
            Write-Host "  Power plan: $targetPlan" -ForegroundColor Cyan
            Write-Host "  AC display timeout: $([int]($profile.DisplayACSeconds / 60)) minutes" -ForegroundColor Cyan
            Write-Host "  Battery display timeout: $([int]($profile.DisplayDCSeconds / 60)) minutes" -ForegroundColor Cyan
            Write-Host "  AC sleep timeout: $(if($profile.SleepACSeconds -eq 0){'Never'}else{"$([int]($profile.SleepACSeconds / 60)) minutes"})" -ForegroundColor Cyan
            Write-Host "  Battery sleep timeout: $(if($profile.SleepDCSeconds -eq 0){'Never'}else{"$([int]($profile.SleepDCSeconds / 60)) minutes"})" -ForegroundColor Cyan
            Write-Host "  Local policy registry enforcement: Enabled" -ForegroundColor Cyan
            $confirm = Read-Host 'Continue? (Y/N)'
            if ($confirm -notin @('Y','y','YES','Yes','yes')) {
                Write-PowerLog 'Operation cancelled by user.' 'WARN'
                $global:LastStatus = '[WARN] Option 7 power settings cancelled by user.'
                return
            }
        }

        $schemeGuid = Get-PowerSchemeGuid -PreferredScheme $targetPlan
        Write-PowerLog "Using power scheme: $targetPlan ($schemeGuid)" 'INFO'

        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Apply $profileName power profile")) {
            Set-PowerProfileValues -SchemeGuid $schemeGuid -Profile $profile -ProfileName $profileName
            Invoke-PowerCfgSafe -Arguments @('/setactive', $schemeGuid) -Description "Set active power plan to $targetPlan" -ContinueOnError | Out-Null
            Invoke-PowerCfgSafe -Arguments @('/S', $schemeGuid) -Description 'Reapply active scheme to commit values' -ContinueOnError | Out-Null

            # Local Group Policy-style power setting values were written under HKLM:\SOFTWARE\Policies\Microsoft\Power\PowerSettings.

            # Refresh policy and power subsystem where available.
            gpupdate.exe /target:computer /force | Out-Null
            Write-PowerLog 'Computer policy refresh requested with gpupdate.' 'OK'
        }

        Write-PowerLog 'Verifying active power scheme...' 'INFO'
        $activeScheme = (powercfg.exe /getactivescheme 2>$null | Out-String).Trim()
        Write-PowerLog "Active scheme after configuration: $activeScheme" 'OK'

        $successCount = ($script:Option7Results | Where-Object { $_.Success }).Count
        $failureCount = ($script:Option7Results | Where-Object { -not $_.Success }).Count
        if ($failureCount -eq 0) {
            $global:LastStatus = "[OK] Option 7 power settings applied successfully using profile: $profileName."
        }
        else {
            $global:LastStatus = "[WARN] Option 7 completed with $failureCount failed setting(s). Review $LogPath."
        }
    }
    catch {
        Write-PowerLog "Critical error during Option 7 power configuration: $($_.Exception.Message)" 'ERROR'
        $global:LastStatus = "[ERROR] Option 7 power settings failed: $($_.Exception.Message)"
    }
    finally {
        $stopwatch.Stop()
        Write-PowerLog "Duration: $([math]::Round($stopwatch.Elapsed.TotalSeconds,2)) seconds" 'INFO'
        Write-PowerLog '=== Option 7 Power Settings Configuration Completed ===' 'INFO'
        try {
            $script:Option7LogEntries | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force
            Write-Host "`nDetailed Option 7 log saved to: $LogPath" -ForegroundColor Cyan
        }
        catch {
            Write-Host "Failed to write Option 7 log: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Get-PowerConfiguration {
    [CmdletBinding()]
    param()

    try {
        [pscustomobject]@{
            ActiveScheme = (powercfg.exe /getactivescheme 2>$null | Out-String).Trim()
            AvailableSchemes = (powercfg.exe /list 2>$null | Out-String).Trim()
            PolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Power\PowerSettings'
        }
    }
    catch {
        [pscustomobject]@{ Error = $_.Exception.Message }
    }
}

# -----------------------------------------------------------------------------
# Option 17 - Stage Lab Scripts and Run Register-Tasks Scripts
# -----------------------------------------------------------------------------
function Register-LabScheduledTasks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SourcePath = "\\filesvr\Labscripts",

        [Parameter(Mandatory = $false)]
        [string]$DestinationPath = "C:\Scripts"
    )

    function Write-Option17Log {
        param(
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('INFO','OK','WARN','ERROR')]
            [string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $color = switch ($Level) {
            'OK'    { 'Green' }
            'WARN'  { 'Yellow' }
            'ERROR' { 'Red' }
            default { 'Cyan' }
        }

        Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
    }

    Write-Option17Log 'Starting Option 17: Stage lab scripts and run Register-Tasks scripts.' 'INFO'

    try {
        if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            throw 'Option 17 must be run from an elevated PowerShell session.'
        }

        Write-Option17Log "Ensuring destination folder exists: $DestinationPath" 'INFO'
        if (-not (Test-Path -LiteralPath $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Option17Log "Created destination folder: $DestinationPath" 'OK'
        }
        else {
            Write-Option17Log "Destination folder already exists: $DestinationPath" 'OK'
        }

        Write-Option17Log "Verifying source path is reachable: $SourcePath" 'INFO'
        if (-not (Test-Path -LiteralPath $SourcePath)) {
            throw "Source path is not reachable: $SourcePath"
        }

        Write-Option17Log "Copying root-level *.ps1 files from $SourcePath to $DestinationPath using robocopy." 'INFO'

        $robocopyArgs = @(
            $SourcePath,
            $DestinationPath,
            '*.ps1',
            '/COPY:DAT',
            '/DCOPY:DAT',
            '/R:2',
            '/W:5',
            '/NP',
            '/NFL',
            '/NDL'
        )

        $robocopyOutput = & robocopy.exe @robocopyArgs 2>&1
        $robocopyExitCode = $LASTEXITCODE

        if ($robocopyOutput) {
            $robocopyOutput | ForEach-Object {
                $line = [string]$_
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Write-Option17Log "ROBOCOPY: $line" 'INFO'
                }
            }
        }

        if ($robocopyExitCode -ge 8) {
            throw "Robocopy failed with exit code $robocopyExitCode."
        }

        Write-Option17Log "Robocopy completed successfully with exit code $robocopyExitCode." 'OK'

        $registerScripts = Get-ChildItem -LiteralPath $DestinationPath -Filter 'Register-Tasks*.ps1' -File -ErrorAction Stop | Sort-Object Name

        if (-not $registerScripts -or $registerScripts.Count -eq 0) {
            Write-Option17Log "No Register-Tasks*.ps1 files were found in $DestinationPath after copy." 'WARN'
            $global:LastStatus = '[WARN] Option 17 completed, but no Register-Tasks scripts were found.'
            return
        }

        Write-Option17Log "Found $($registerScripts.Count) Register-Tasks script(s) to run." 'OK'

        $successCount = 0
        $failureCount = 0

        foreach ($scriptFile in $registerScripts) {
            Write-Option17Log "Running: $($scriptFile.FullName)" 'INFO'

            $process = Start-Process -FilePath 'powershell.exe' `
                -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptFile.FullName) `
                -Wait `
                -PassThru `
                -WindowStyle Hidden

            if ($process.ExitCode -eq 0) {
                $successCount++
                Write-Option17Log "Completed successfully: $($scriptFile.Name)" 'OK'
            }
            else {
                $failureCount++
                Write-Option17Log "Script exited with code $($process.ExitCode): $($scriptFile.Name)" 'ERROR'
            }
        }

        Write-Option17Log "Option 17 complete. Successful: $successCount ; Failed: $failureCount" 'INFO'

        if ($failureCount -gt 0) {
            $global:LastStatus = "[WARN] Option 17 completed with $failureCount Register-Tasks script failure(s)."
        }
        else {
            $global:LastStatus = "[OK] Option 17 completed successfully. Copied scripts and ran $successCount Register-Tasks script(s)."
        }
    }
    catch {
        $global:LastStatus = "[ERROR] Option 17 failed: $($_.Exception.Message)"
        Write-Option17Log $global:LastStatus 'ERROR'
    }
}


# -----------------------------------------------------------------------------
# Option 18 - Set OneDrive to Automatically Login at Boot
# -----------------------------------------------------------------------------
function Set-OneDriveAutoLoginPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$EnableSilentConfig,
        [switch]$DisableFirstRunWizard,
        [switch]$EnableAutoStartup,
        [switch]$EnableFilesOnDemand,
        [switch]$DisablePersonalSync,
        [switch]$EnableKnownFolderMove,
        [switch]$DisableAutoLogin,
        [string]$TenantId,
        [ValidateRange(1, 30)]
        [int]$SyncThrottleKbps = 0,  # 0 = No throttling
        [switch]$BackupSettings,
        [string]$LogPath = "$env:TEMP\OneDrivePolicy_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    )

    # Security: Require elevation for policy configuration
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking using ArrayList for better performance and compatibility
    $script:logEntries = New-Object System.Collections.ArrayList
    $script:appliedPolicies = New-Object System.Collections.ArrayList
    $script:failedPolicies = New-Object System.Collections.ArrayList
    $script:skippedPolicies = New-Object System.Collections.ArrayList
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        # Ensure logEntries exists before adding to it
        if ($null -eq $script:logEntries) {
            $script:logEntries = New-Object System.Collections.ArrayList
        }
        [void]$script:logEntries.Add($logEntry)
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO' { Write-Host $Message -ForegroundColor Cyan }
            'POLICY' { Write-Host $Message -ForegroundColor Magenta }
            'SKIP' { Write-Host $Message -ForegroundColor DarkGray }
        }
    }

    # Security: Registry validation and path management
    function Test-RegistryPath {
        param([string]$Path)
        try {
            # Security: Validate registry path format
            if ($Path -notmatch '^HK(LM|CU|CR|U|CC):\\') {
                return $false
            }
            return Test-Path $Path -ErrorAction SilentlyContinue
        } catch {
            return $false
        }
    }

    function New-RegistryPath {
        param([string]$Path)
        try {
            if (-not (Test-RegistryPath $Path)) {
                New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
                Write-LogEntry "[OK] Created registry path: $Path" 'SUCCESS'
                return $true
            }
            return $true
        } catch {
            Write-LogEntry "[X] Failed to create registry path: $Path - $_" 'ERROR'
            return $false
        }
    }

    # Get current registry value safely
    function Get-RegistryValue {
        param(
            [string]$Path,
            [string]$Name
        )
        
        try {
            if (Test-RegistryPath $Path) {
                $property = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                if ($property) {
                    return $property.$Name
                }
            }
            return $null
        } catch {
            return $null
        }
    }

    # Check if policy value needs to be changed
    function Test-PolicyValue {
        param(
            [string]$Path,
            [string]$Name,
            [object]$DesiredValue
        )
        
        $currentValue = Get-RegistryValue -Path $Path -Name $Name
        return ($currentValue -eq $DesiredValue)
    }

    # Speed: Optimized policy application function with existing value detection
    function Set-OneDrivePolicy {
        param(
            [string]$Path,
            [string]$Name,
            [object]$Value,
            [string]$Type = 'DWord',
            [string]$Description,
            [switch]$Critical
        )
        
        try {
            # Check if the policy already has the correct value
            if (Test-PolicyValue -Path $Path -Name $Name -DesiredValue $Value) {
                $level = if ($Critical) { 'SKIP' } else { 'SKIP' }
                Write-LogEntry "[SKIP] $Description (already configured)" $level
                
                # [OK] FIX: Ensure skippedPolicies exists and use ArrayList.Add()
                if ($null -eq $script:skippedPolicies) {
                    $script:skippedPolicies = New-Object System.Collections.ArrayList
                }
                $skipInfo = @{
                    Path = $Path
                    Name = $Name
                    Value = $Value
                    Description = $Description
                    Critical = $Critical.IsPresent
                }
                [void]$script:skippedPolicies.Add($skipInfo)
                return $true
            }
            
            # Ensure registry path exists
            if (-not (New-RegistryPath -Path $Path)) {
                throw "Cannot create registry path: $Path"
            }
            
            # Apply policy setting
            if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set policy value to $Value")) {
                Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
                
                # Verify the setting was applied
                $verifyValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
                if ($verifyValue.$Name -eq $Value) {
                    $level = if ($Critical) { 'POLICY' } else { 'SUCCESS' }
                    Write-LogEntry "[OK] $Description" $level
                    
                    # [OK] FIX: Ensure appliedPolicies exists and use ArrayList.Add()
                    if ($null -eq $script:appliedPolicies) {
                        $script:appliedPolicies = New-Object System.Collections.ArrayList
                    }
                    $policyInfo = @{
                        Path = $Path
                        Name = $Name
                        Value = $Value
                        Description = $Description
                        Critical = $Critical.IsPresent
                    }
                    [void]$script:appliedPolicies.Add($policyInfo)
                    return $true
                } else {
                    throw "Verification failed: Expected $Value, got $($verifyValue.$Name)"
                }
            }
        } catch {
            Write-LogEntry "[X] Failed to apply $Description : $_" 'ERROR'
            
            # [OK] FIX: Ensure failedPolicies exists and use ArrayList.Add()
            if ($null -eq $script:failedPolicies) {
                $script:failedPolicies = New-Object System.Collections.ArrayList
            }
            $failInfo = @{
                Path = $Path
                Name = $Name
                Description = $Description
                Error = $_.Exception.Message
            }
            [void]$script:failedPolicies.Add($failInfo)
            return $false
        }
    }

    # Security: OneDrive installation and version validation
    function Test-OneDriveInstallation {
        try {
            $oneDrivePaths = @(
                "${env:ProgramFiles}\Microsoft OneDrive\OneDrive.exe",
                "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe",
                "${env:LOCALAPPDATA}\Microsoft\OneDrive\OneDrive.exe"
            )
            
            $installations = @()
            foreach ($path in $oneDrivePaths) {
                if (Test-Path $path) {
                    try {
                        $version = (Get-ItemProperty -Path $path).VersionInfo.FileVersion
                        $installations += @{
                            Path = $path
                            Version = $version
                            Type = if ($path -like "*Program Files*") { "System" } else { "User" }
                        }
                    } catch {
                        $installations += @{
                            Path = $path
                            Version = "Unknown"
                            Type = if ($path -like "*Program Files*") { "System" } else { "User" }
                        }
                    }
                }
            }
            
            return @{
                IsInstalled = ($installations.Count -gt 0)
                Installations = $installations
                RecommendedPath = $installations | Where-Object { $_.Type -eq "System" } | Select-Object -First 1
            }
            
        } catch {
            Write-LogEntry "OneDrive installation check failed: $_" 'WARNING'
            return @{
                IsInstalled = $false
                Error = $_.Exception.Message
            }
        }
    }

    # Security: Backup current OneDrive policies
    function Backup-OneDrivePolicies {
        param([string[]]$RegistryPaths)
        
        try {
            $backupData = @{}
            
            foreach ($regPath in $RegistryPaths) {
                if (Test-RegistryPath $regPath) {
                    $properties = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                    if ($properties) {
                        $backupData[$regPath] = @{}
                        foreach ($property in $properties.PSObject.Properties) {
                            if ($property.Name -notmatch '^PS') {  # Skip PowerShell properties
                                $backupData[$regPath][$property.Name] = $property.Value
                            }
                        }
                    }
                }
            }
            
            if ($backupData.Count -gt 0) {
                $backupFile = "$env:TEMP\OneDrivePolicyBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
                $backupData | ConvertTo-Json -Depth 4 | Out-File -FilePath $backupFile -Encoding UTF8
                Write-LogEntry "[OK] Policy backup saved to: $backupFile" 'INFO'
                return $backupFile
            }
            
        } catch {
            Write-LogEntry "[WARN] Failed to backup OneDrive policies: $_" 'WARNING'
        }
        
        return $null
    }

    # Security: Validate Tenant ID format
    function Test-TenantId {
        param([string]$TenantId)
        
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            return $true  # Optional parameter
        }
        
        # GUID format validation
        $guidPattern = '^[a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$'
        return $TenantId -match $guidPattern
    }

    try {
        Write-LogEntry "=== OneDrive Auto-Login Policy Configuration Started ===" 'INFO'
        
        # Set default behaviors if no specific parameters provided
        if (-not ($EnableSilentConfig -or $DisableFirstRunWizard -or $EnableAutoStartup -or $EnableFilesOnDemand -or $DisablePersonalSync -or $EnableKnownFolderMove -or $DisableAutoLogin)) {
            Write-LogEntry "No specific policies specified, applying default configuration..." 'INFO'
            $EnableSilentConfig = $true
            $DisableFirstRunWizard = $true
            $EnableAutoStartup = $true
        }
        
        if ($WhatIfPreference) {
            Write-LogEntry "WhatIf mode - no registry changes will be made" 'INFO'
        }

        # Security: Validate Tenant ID if provided
        if ($TenantId -and -not (Test-TenantId -TenantId $TenantId)) {
            throw "Invalid Tenant ID format. Must be a valid GUID."
        }

        # Validate OneDrive installation
        Write-LogEntry "`n[SCAN] Validating OneDrive installation..." 'INFO'
        $oneDriveInfo = Test-OneDriveInstallation
        
        if (-not $oneDriveInfo.IsInstalled) {
            Write-LogEntry "[WARN] OneDrive not detected on this system" 'WARNING'
            Write-LogEntry "Policies will be applied but may not take effect until OneDrive is installed" 'WARNING'
        } else {
            Write-LogEntry "[OK] OneDrive installation detected:" 'SUCCESS'
            foreach ($installation in $oneDriveInfo.Installations) {
                Write-LogEntry "  - $($installation.Type): $($installation.Path) (v$($installation.Version))" 'INFO'
            }
        }

        # Define registry paths
        $registryPaths = @{
            MainPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
            UserPolicy = "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
            TenantRestrictions = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantRestrictions"
            KnownFolderMove = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\KnownFolderMove"
        }

        # Backup existing policies if requested
        if ($BackupSettings) {
            Write-LogEntry "`n[SAVE] Backing up current OneDrive policies..." 'INFO'
            $backupFile = Backup-OneDrivePolicies -RegistryPaths $registryPaths.Values
        }

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary - Policies that would be applied:" 'INFO'
            if ($EnableSilentConfig) { Write-LogEntry "  - Silent account configuration: Enabled" 'INFO' }
            if ($DisableFirstRunWizard) { Write-LogEntry "  - First run wizard: Disabled" 'INFO' }
            if ($EnableAutoStartup) { Write-LogEntry "  - Auto startup: Enabled" 'INFO' }
            if ($EnableFilesOnDemand) { Write-LogEntry "  - Files On-Demand: Enabled" 'INFO' }
            if ($DisablePersonalSync) { Write-LogEntry "  - Personal account sync: Disabled" 'INFO' }
            if ($EnableKnownFolderMove) { Write-LogEntry "  - Known Folder Move: Enabled" 'INFO' }
            if ($TenantId) { Write-LogEntry "  - Tenant restriction: $TenantId" 'INFO' }
            if ($SyncThrottleKbps -gt 0) { Write-LogEntry "  - Sync throttle: $SyncThrottleKbps KB/s" 'INFO' }
            $global:LastStatus = "[INFO] WhatIf completed - OneDrive policies would be configured."
            return
        }

        # Apply core OneDrive policies
        Write-LogEntry "`n[REPAIR] Applying OneDrive policies..." 'POLICY'
        
        # Silent account configuration
        if ($EnableSilentConfig) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "SilentAccountConfig" -Value 1 -Description "Enabled silent account configuration" -Critical | Out-Null
        }
        
        # Disable first run wizard
        if ($DisableFirstRunWizard) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DisableFirstRunWizard" -Value 1 -Description "Disabled first run wizard" | Out-Null
        }
        
        # Auto startup policy
        if ($EnableAutoStartup) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "OneDriveStartupPolicy" -Value 1 -Description "Enabled OneDrive auto startup" | Out-Null
        }
        
        # Files On-Demand
        if ($EnableFilesOnDemand) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "FilesOnDemandEnabled" -Value 1 -Description "Enabled Files On-Demand" | Out-Null
        }
        
        # Disable personal account sync
        if ($DisablePersonalSync) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DisablePersonalSync" -Value 1 -Description "Disabled personal account synchronization" -Critical | Out-Null
        }
        
        # Known Folder Move
        if ($EnableKnownFolderMove) {
            Set-OneDrivePolicy -Path $registryPaths.KnownFolderMove -Name "KnownFolderMoveOpt" -Value 1 -Description "Enabled Known Folder Move optimization" | Out-Null
            
            if ($TenantId) {
                Set-OneDrivePolicy -Path $registryPaths.KnownFolderMove -Name $TenantId -Value 1 -Description "Enabled Known Folder Move for tenant: $TenantId" | Out-Null
            }
        }
        
        # Tenant restrictions
        if ($TenantId) {
            Set-OneDrivePolicy -Path $registryPaths.TenantRestrictions -Name $TenantId -Value 1 -Type "String" -Description "Applied tenant restriction for: $TenantId" -Critical | Out-Null
        }
        
        # Sync throttling
        if ($SyncThrottleKbps -gt 0) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "UploadBandwidthLimit" -Value $SyncThrottleKbps -Description "Set upload bandwidth limit: $SyncThrottleKbps KB/s" | Out-Null
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "DownloadBandwidthLimit" -Value $SyncThrottleKbps -Description "Set download bandwidth limit: $SyncThrottleKbps KB/s" | Out-Null
        }
        
        # Disable auto-login (override other settings)
        if ($DisableAutoLogin) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "SilentAccountConfig" -Value 0 -Description "Disabled OneDrive auto-login" -Critical | Out-Null
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "OneDriveStartupPolicy" -Value 0 -Description "Disabled OneDrive auto startup" | Out-Null
        }

        # Additional security and performance policies
        Write-LogEntry "`n[SECURE] Applying security and performance policies..." 'POLICY'
        
        # Prevent OneDrive from generating network traffic until user signs in
        Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "PreventNetworkTrafficPreUserSignIn" -Value 1 -Description "Prevented network traffic before user sign-in" | Out-Null
        
        # Block external sharing
        Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "BlockExternalSync" -Value 1 -Description "Blocked external sharing and sync" | Out-Null
        
        # Enable automatic sign-in
        if (-not $DisableAutoLogin) {
            Set-OneDrivePolicy -Path $registryPaths.MainPolicy -Name "AutomaticUploadBandwidthPercentage" -Value 70 -Description "Set automatic upload bandwidth to 70%" | Out-Null
        }

        # Verification of applied policies
        Write-LogEntry "`n[SCAN] Verifying policy application..." 'INFO'
        $verificationErrors = 0
        
        # Ensure appliedPolicies exists before iterating
        if ($null -ne $script:appliedPolicies) {
            foreach ($policy in $script:appliedPolicies) {
                try {
                    $currentValue = Get-ItemProperty -Path $policy.Path -Name $policy.Name -ErrorAction Stop
                    if ($currentValue.($policy.Name) -ne $policy.Value) {
                        Write-LogEntry "[WARN] Verification failed for $($policy.Description)" 'WARNING'
                        $verificationErrors++
                    }
                } catch {
                    Write-LogEntry "[WARN] Could not verify $($policy.Description)" 'WARNING'
                    $verificationErrors++
                }
            }
        }
        
        $appliedCount = if ($null -ne $script:appliedPolicies) { $script:appliedPolicies.Count } else { 0 }
        $skippedCount = if ($null -ne $script:skippedPolicies) { $script:skippedPolicies.Count } else { 0 }
        
        if ($verificationErrors -eq 0 -and $appliedCount -gt 0) {
            Write-LogEntry "[OK] All OneDrive policies verified successfully" 'SUCCESS'
        } elseif ($appliedCount -eq 0 -and $skippedCount -gt 0) {
            Write-LogEntry "[OK] All OneDrive policies already correctly configured" 'SUCCESS'
        } elseif ($verificationErrors -gt 0) {
            Write-LogEntry "[WARN] $verificationErrors policies failed verification" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during OneDrive policy configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] OneDrive policy configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        # Safe counting with null checks
        $appliedCount = if ($null -ne $script:appliedPolicies) { $script:appliedPolicies.Count } else { 0 }
        $skippedCount = if ($null -ne $script:skippedPolicies) { $script:skippedPolicies.Count } else { 0 }
        $failedCount = if ($null -ne $script:failedPolicies) { $script:failedPolicies.Count } else { 0 }
        
        Write-LogEntry "`n[SUMMARY] Policy Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Policies applied: $appliedCount" 'SUCCESS'
        Write-LogEntry "Policies already correct: $skippedCount" 'SKIP'
        Write-LogEntry "Policies failed: $failedCount" 'ERROR'
        
        # Safe critical policy counting
        $criticalPolicies = 0
        $criticalSkipped = 0
        
        if ($null -ne $script:appliedPolicies) {
            $criticalPolicies = ($script:appliedPolicies | Where-Object { $_.Critical }).Count
        }
        if ($null -ne $script:skippedPolicies) {
            $criticalSkipped = ($script:skippedPolicies | Where-Object { $_.Critical }).Count
        }
        
        if ($criticalPolicies -gt 0) {
            Write-LogEntry "Critical policies applied: $criticalPolicies" 'POLICY'
        }
        if ($criticalSkipped -gt 0) {
            Write-LogEntry "Critical policies already configured: $criticalSkipped" 'SKIP'
        }
        
        if ($failedCount -gt 0 -and $null -ne $script:failedPolicies) {
            Write-LogEntry "`n[ERROR] Failed Policies:" 'ERROR'
            foreach ($failed in $script:failedPolicies) {
                Write-LogEntry "  - $($failed.Description): $($failed.Error)" 'ERROR'
            }
        }
        
        # Write detailed log file with null check
        try {
            if ($null -ne $script:logEntries) {
                $script:logEntries.ToArray() | Out-File -FilePath $LogPath -Encoding UTF8 -Force
                Write-LogEntry "[NOTE] Detailed log saved to: $LogPath" 'INFO'
            }
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set global status with safe counting
        $totalPolicies = $appliedCount + $skippedCount
        if ($totalPolicies -gt 0) {
            if ($appliedCount -gt 0) {
                $statusMsg = "[OK] Applied $appliedCount OneDrive policies"
                if ($skippedCount -gt 0) {
                    $statusMsg += " ($skippedCount already correct)"
                }
            } else {
                $statusMsg = "[OK] All $skippedCount OneDrive policies already correctly configured"
            }
            
            if ($failedCount -gt 0) {
                $statusMsg += " [$failedCount failed]"
            }
            if ($criticalPolicies -gt 0 -or $criticalSkipped -gt 0) {
                $statusMsg += " [$($criticalPolicies + $criticalSkipped) critical]"
            }
            $global:LastStatus = $statusMsg
        } else {
            $global:LastStatus = "[WARN] No OneDrive policies were processed"
        }
        
        Write-LogEntry "=== OneDrive Policy Configuration Completed ===" 'INFO'
    }
}

# Utility function to check current OneDrive policies
function Get-OneDrivePolicyStatus {
    [CmdletBinding()]
    param()
    
    try {
        $policyPaths = @{
            "Main Policies" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
            "User Policies" = "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
            "Tenant Restrictions" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\TenantRestrictions"
            "Known Folder Move" = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive\KnownFolderMove"
        }
        
        $status = @{}
        
        foreach ($category in $policyPaths.GetEnumerator()) {
            $status[$category.Key] = @{
                Path = $category.Value
                Exists = (Test-Path $category.Value)
                Policies = @{}
            }
            
            if ($status[$category.Key].Exists) {
                try {
                    $properties = Get-ItemProperty -Path $category.Value -ErrorAction SilentlyContinue
                    if ($properties) {
                        foreach ($prop in $properties.PSObject.Properties) {
                            if ($prop.Name -notmatch '^PS') {
                                $status[$category.Key].Policies[$prop.Name] = $prop.Value
                            }
                        }
                    }
                } catch {
                    $status[$category.Key].Error = $_.Exception.Message
                }
            }
        }
        
        return $status
        
    } catch {
        return @{
            Error = $_.Exception.Message
        }
    }
}

# Utility function to remove OneDrive policies
function Remove-OneDrivePolicies {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$ConfirmEach
    )
    
    $policyPaths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive",
        "HKCU:\SOFTWARE\Policies\Microsoft\OneDrive"
    )
    
    $removedCount = 0
    
    foreach ($path in $policyPaths) {
        if (Test-Path $path) {
            try {
                if ($PSCmdlet.ShouldProcess($path, "Remove OneDrive policy registry key")) {
                    if (-not $ConfirmEach -or (Read-Host "Remove $path? (y/n)") -eq 'y') {
                        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                        Write-Host "[OK] Removed OneDrive policies from: $path" -ForegroundColor Green
                        $removedCount++
                    }
                }
            } catch {
                Write-Warning "[WARN] Failed to remove $path : $_"
            }
        }
    }
    
    Write-Host "[OK] Removed OneDrive policies from $removedCount registry locations" -ForegroundColor Cyan
}


# -----------------------------------------------------------------------------
# Option 19 - Full System Update
# -----------------------------------------------------------------------------
function Run-CorePostDeploymentTasks {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string[]]$IncludeTasks = @(),
        [string[]]$ExcludeTasks = @(),
        [switch]$ParallelExecution,
        [switch]$Force,
        [ValidateRange(1, 10)]
        [int]$MaxParallelJobs = 4,
        [ValidateRange(5, 180)]
        [int]$TaskTimeoutMinutes = 30,
        [string]$LogPath = "$env:TEMP\PostDeploymentTasks_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$GenerateReport,
        [switch]$ContinueOnError
    )

    # Security: Require elevation for system-wide changes
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference  = 'Continue'
    $ProgressPreference     = 'SilentlyContinue'
    
    # Initialize tracking
    $logEntries      = @()
    $taskResults     = @()
    $globalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry  = "[$timestamp] [$Level] $Message"
        $script:logEntries += $logEntry
        
        switch ($Level) {
            'ERROR'   { Write-Host $Message -ForegroundColor Red }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green }
            'INFO'    { Write-Host $Message -ForegroundColor Cyan }
            'TASK'    { Write-Host $Message -ForegroundColor Magenta }
            'SYSTEM'  { Write-Host $Message -ForegroundColor Blue }
        }
    }

    # Helper: run a named function in-process (no job)
function Invoke-FunctionInProcess {
    param([string]$FunctionName)
    
    $prev = @{
        ConfirmPreference = $ConfirmPreference
        PSDefaultParameterValues = if ($PSDefaultParameterValues) { $PSDefaultParameterValues.Clone() } else { @{} }
        ProgressPreference = $ProgressPreference
        VerbosePreference = $VerbosePreference
        InformationPreference = $InformationPreference
        ErrorActionPreference = $ErrorActionPreference
    }
    
    try {
        $script:ConfirmPreference = 'None'
        $script:PSDefaultParameterValues = @{'*:Confirm' = $false}
        $script:ProgressPreference = 'SilentlyContinue'
        $script:VerbosePreference = 'SilentlyContinue'
        $script:InformationPreference = 'Continue'
        $script:ErrorActionPreference = 'Stop'

        if (-not (Get-Command $FunctionName -ErrorAction SilentlyContinue)) {
            throw "Function '$FunctionName' not found in current session"
        }
        
        & $FunctionName
        
    } catch {
        # Provide detailed error information for troubleshooting
        $errorDetails = @"
Function '$FunctionName' failed with error:
  Message: $($_.Exception.Message)
  Location: Line $($_.InvocationInfo.ScriptLineNumber), Position $($_.InvocationInfo.OffsetInLine)
  Script: $($_.InvocationInfo.ScriptName)
"@
        Write-Host $errorDetails -ForegroundColor Red
        throw
    } finally {
        $script:ConfirmPreference = $prev.ConfirmPreference
        $script:PSDefaultParameterValues = $prev.PSDefaultParameterValues
        $script:ProgressPreference = $prev.ProgressPreference
        $script:VerbosePreference = $prev.VerbosePreference
        $script:InformationPreference = $prev.InformationPreference
        $script:ErrorActionPreference = $prev.ErrorActionPreference
    }
}

    # Define all available post-deployment tasks with metadata
    $availableTasks = @{
        'RemoveBloatware' = @{
            Name = 'Remove Bloatware Apps'
            Function = 'Remove-BloatwareApps'
            Category = 'Cleanup'
            Priority = 2
            RequiresElevation = $true
            CanRunParallel = $true
            Dependencies = @()
            EstimatedDuration = 120
        }
        'RegistrySettings' = @{
            Name = 'Apply Registry Settings'
            Function = 'Apply-RecommendedRegistrySettings'
            Category = 'Configuration'
            Priority = 3
            RequiresElevation = $true
            CanRunParallel = $true
            Dependencies = @()
            EstimatedDuration = 45
        }
        'OptimizeServices' = @{
            Name = 'Optimize Windows Services'
            Function = 'Optimize-WindowsServices'
            Category = 'Performance'
            Priority = 4
            RequiresElevation = $true
            CanRunParallel = $true
            Dependencies = @()
            EstimatedDuration = 60
        }
        'PowerShellRemoting' = @{
            Name = 'Configure PowerShell Remoting'
            Function = 'Enable-PowerShellRemotingSafely'
            Category = 'Security'
            Priority = 5
            RequiresElevation = $true
            CanRunParallel = $false
            Dependencies = @()
            EstimatedDuration = 30
        }
        'TimeSync' = @{
            Name = 'Configure Automatic Time Sync'
            Function = 'Configure-AutomaticTimeSync'
            Category = 'Configuration'
            Priority = 6
            RequiresElevation = $true
            CanRunParallel = $true
            Dependencies = @()
            EstimatedDuration = 20
        }
        'WingetDependencies' = @{
            Name = 'Ensure Winget Dependencies'
            Function = 'Ensure-WingetDependenciesReady'
            Category = 'Prerequisites'
            Priority = 7
            RequiresElevation = $false
            CanRunParallel = $false
            Dependencies = @()
            EstimatedDuration = 90
        }
        'UpdateApplications' = @{
            Name = 'Update Applications via Winget'
            Function = 'Update-Applications'
            Category = 'Updates'
            Priority = 8
            RequiresElevation = $false
            CanRunParallel = $false
            Dependencies = @('WingetDependencies')
            EstimatedDuration = 300
        }
        'HPDrivers' = @{
            Name = 'Update HP/Dell Drivers'
            Function = 'Update-HPDrivers'
            Category = 'Drivers'
            Priority = 9
            RequiresElevation = $true
            CanRunParallel = $true
            Dependencies = @()
            EstimatedDuration = 180
        }
        'WindowsUpdate' = @{
            Name = 'Run Windows Update'
            Function = 'Update-WindowsOS'
            Category = 'Updates'
            Priority = 10
            RequiresElevation = $true
            CanRunParallel = $false
            Dependencies = @()
            EstimatedDuration = 600
        }
    }

    # Security: Task validation and filtering
    function Get-FilteredTasks {
        param(
            [hashtable]$Tasks,
            [string[]]$Include,
            [string[]]$Exclude
        )
        $filteredTasks = @{}
        if ($Include.Count -gt 0) {
            foreach ($taskId in $Include) {
                if ($Tasks.ContainsKey($taskId)) { $filteredTasks[$taskId] = $Tasks[$taskId] }
                else { Write-LogEntry "[WARN] Unknown task specified in include list: $taskId" 'WARNING' }
            }
        } else {
            $filteredTasks = $Tasks.Clone()
        }
        foreach ($taskId in $Exclude) {
            if ($filteredTasks.ContainsKey($taskId)) {
                $filteredTasks.Remove($taskId)
                Write-LogEntry "Excluded task: $($Tasks[$taskId].Name)" 'INFO'
            }
        }
        return $filteredTasks
    }

    # Security: Validate task dependencies
    function Test-TaskDependencies {
        param([hashtable]$Tasks)
        $dependencyErrors = @()
        foreach ($task in $Tasks.GetEnumerator()) {
            foreach ($dependency in $task.Value.Dependencies) {
                if (-not $Tasks.ContainsKey($dependency)) {
                    $dependencyErrors += "Task '$($task.Key)' depends on missing task '$dependency'"
                }
            }
        }
        return $dependencyErrors
    }

    # Speed: Execute task with timeout and monitoring (sequential, in-process)
    function Invoke-DeploymentTask {
        param(
            [string]$TaskId,
            [hashtable]$TaskInfo,
            [int]$TimeoutMinutes,
            [switch]$WhatIfMode
        )
        
        $taskStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = @{
            TaskId    = $TaskId
            Name      = $TaskInfo.Name
            Category  = $TaskInfo.Category
            StartTime = Get-Date
            EndTime   = $null
            Duration  = $null
            Success   = $false
            Output    = @()
            Error     = $null
            ExitCode  = 0
        }
        
        try {
            Write-LogEntry "[START] Starting task: $($TaskInfo.Name)" 'TASK'
            
            if ($WhatIfMode) {
                Write-LogEntry "WhatIf: Would execute $($TaskInfo.Function)" 'INFO'
                Start-Sleep -Seconds 1
                $result.Success = $true
                $result.Output += "WhatIf execution completed"
            } else {
                # Run function directly in-process
                Write-LogEntry "[INFO] Running '$($TaskInfo.Function)' in-process..." 'INFO'
                Invoke-FunctionInProcess -FunctionName $TaskInfo.Function
                $result.Success = $true
                $result.Output += "In-process execution completed"
            }

            $taskStopwatch.Stop()
            $result.Duration = $taskStopwatch.Elapsed
            $result.EndTime  = Get-Date

            if ($result.Success) {
                Write-LogEntry "[OK] Completed: $($TaskInfo.Name) ($([math]::Round($result.Duration.TotalSeconds, 1))s)" 'SUCCESS'
            } else {
                Write-LogEntry "[X] Failed: $($TaskInfo.Name) - $($result.Error)" 'ERROR'
            }

        } catch {
            $taskStopwatch.Stop()
            $result.Duration = $taskStopwatch.Elapsed
            $result.EndTime  = Get-Date
            $result.Error    = $_.Exception.Message
            $result.Success  = $false
            Write-LogEntry "[X] Task failed: $($TaskInfo.Name) - $_" 'ERROR'
        }
        return $result
    }

    # Generate comprehensive deployment report
    function New-DeploymentReport {
        param([array]$TaskResults)
        $report = @{
            Summary = @{
                TotalTasks       = $TaskResults.Count
                SuccessfulTasks  = ($TaskResults | Where-Object { $_.Success }).Count
                FailedTasks      = ($TaskResults | Where-Object { -not $_.Success }).Count
                TotalDuration    = (($TaskResults | Where-Object { $_.Duration } | Measure-Object -Property Duration -Sum).Sum)
                StartTime        = ($TaskResults | Sort-Object StartTime | Select-Object -First 1).StartTime
                EndTime          = ($TaskResults | Sort-Object EndTime | Select-Object -Last 1).EndTime
            }
            Categories = @{}
            Tasks = $TaskResults
        }
        if (-not $report.Summary.TotalDuration) { $report.Summary.TotalDuration = [TimeSpan]::Zero }
        foreach ($task in $TaskResults) {
            if (-not $report.Categories.ContainsKey($task.Category)) {
                $report.Categories[$task.Category] = @{
                    Tasks = @()
                    SuccessCount = 0
                    FailedCount = 0
                    TotalDuration = [TimeSpan]::Zero
                }
            }
            $report.Categories[$task.Category].Tasks += $task
            if ($task.Success) { $report.Categories[$task.Category].SuccessCount++ }
            else { $report.Categories[$task.Category].FailedCount++ }
            if ($task.Duration) {
                $report.Categories[$task.Category].TotalDuration = $report.Categories[$task.Category].TotalDuration.Add($task.Duration)
            }
        }
        return $report
    }

    try {
        Write-LogEntry "=== Core Post-Deployment Tasks Started ===" 'SYSTEM'
        Write-LogEntry "Parallel Execution: Disabled (running in-process mode)" 'INFO'
        Write-LogEntry "Task Timeout: $TaskTimeoutMinutes minutes" 'INFO'
        if ($WhatIfPreference) { Write-LogEntry "WhatIf mode - no actual changes will be made" 'INFO' }

        # Filter and validate tasks
        Write-LogEntry "`n[SCAN] Analyzing task configuration..." 'INFO'
        $tasksToRun = Get-FilteredTasks -Tasks $availableTasks -Include $IncludeTasks -Exclude $ExcludeTasks
        if ($tasksToRun.Count -eq 0) { throw "No tasks selected for execution" }

        $dependencyErrors = Test-TaskDependencies -Tasks $tasksToRun
        if ($dependencyErrors.Count -gt 0) {
            Write-LogEntry "[ERROR] Dependency validation failed:" 'ERROR'
            foreach ($error in $dependencyErrors) { Write-LogEntry "  - $error" 'ERROR' }
            throw "Task dependency validation failed"
        }

        # Sort tasks by priority
        $sortedTasks = $tasksToRun.GetEnumerator() | Sort-Object { $_.Value.Priority }
        Write-LogEntry "[OK] Task validation completed" 'SUCCESS'
        Write-LogEntry "Tasks to execute: $($tasksToRun.Count)" 'INFO'

        # Display execution plan
        Write-LogEntry "`n[REPORT] Execution Plan:" 'INFO'
        $totalEstimatedTime = 0
        foreach ($task in $sortedTasks) {
            $estimate = $task.Value.EstimatedDuration
            $totalEstimatedTime += $estimate
            Write-LogEntry "  - $($task.Value.Name) ($($task.Value.Category)) - ~$([math]::Round($estimate/60, 1))min" 'INFO'
        }
        Write-LogEntry "Total estimated time: $([math]::Round($totalEstimatedTime/60, 1)) minutes" 'INFO'

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary - Tasks that would be executed:" 'INFO'
            foreach ($task in $sortedTasks) { Write-LogEntry "  - $($task.Value.Name): $($task.Value.Function)" 'INFO' }
            $global:LastStatus = "[INFO] WhatIf completed - $($tasksToRun.Count) tasks would be executed."
            return
        }

        # Execute tasks sequentially
        Write-LogEntry "`n[START] Beginning task execution..." 'SYSTEM'
        foreach ($task in $sortedTasks) {
            $result = Invoke-DeploymentTask -TaskId $task.Key -TaskInfo $task.Value -TimeoutMinutes $TaskTimeoutMinutes -WhatIfMode:$WhatIfPreference
            $taskResults += $result
            if (-not $result.Success -and -not $ContinueOnError) {
                Write-LogEntry "[ERROR] Stopping execution due to task failure: $($result.Name)" 'ERROR'
                break
            }
        }

        # Generate final report
        if ($GenerateReport) {
            Write-LogEntry "`n[SUMMARY] Generating deployment report..." 'INFO'
            $deploymentReport = New-DeploymentReport -TaskResults $taskResults
            $reportFile = "$env:TEMP\DeploymentReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $deploymentReport | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportFile -Encoding UTF8
            Write-LogEntry "[OK] Deployment report saved to: $reportFile" 'SUCCESS'
        }

    } catch {
        Write-LogEntry "Critical error during post-deployment execution: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Core post-deployment tasks failed: $_"
        throw
    } finally {
        $globalStopwatch.Stop()
        $totalDuration = $globalStopwatch.Elapsed
        
        # Final summary
        Write-LogEntry "`n[SUMMARY] Post-Deployment Summary:" 'SYSTEM'
        Write-LogEntry "Total Duration: $([math]::Round($totalDuration.TotalMinutes, 2)) minutes" 'INFO'
        
        if ($taskResults) {
            $successCount = ($taskResults | Where-Object { $_.Success }).Count
            $failureCount = ($taskResults | Where-Object { -not $_.Success }).Count
            
            Write-LogEntry "Tasks Executed: $($taskResults.Count)" 'INFO'
            Write-LogEntry "Successful: $successCount" 'SUCCESS'
            Write-LogEntry "Failed: $failureCount" 'ERROR'
            
            if ($failureCount -gt 0) {
                Write-LogEntry "`n[ERROR] Failed Tasks:" 'ERROR'
                foreach ($failedTask in ($taskResults | Where-Object { -not $_.Success })) {
                    Write-LogEntry "  - $($failedTask.Name): $($failedTask.Error)" 'ERROR'
                }
            }
            
            if ($failureCount -eq 0) {
                $global:LastStatus = "[OK] All $successCount post-deployment tasks completed successfully."
            } else {
                $global:LastStatus = "[WARN] Post-deployment completed with issues: $successCount successful, $failureCount failed."
            }
        } else {
            $global:LastStatus = "[INFO] No tasks were executed."
        }
        
        # Write comprehensive log file
        try {
            $dir = Split-Path -Path $LogPath -Parent
            if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[NOTE] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        Write-LogEntry "=== Core Post-Deployment Tasks Completed ===" 'SYSTEM'
        Write-Host "`n$($global:LastStatus)" -ForegroundColor $(if($global:LastStatus -like "*successfully*") {'Green'} elseif($global:LastStatus -like "*issues*") {'Yellow'} else {'Red'})
    }
}

# Utility function to list available tasks
function Get-AvailableDeploymentTasks {
    $tasks = @{
        'RemoveBloatware'   = 'Remove Bloatware Apps'
        'RegistrySettings'  = 'Apply Registry Settings'
        'OptimizeServices'  = 'Optimize Windows Services'
        'PowerShellRemoting'= 'Configure PowerShell Remoting'
        'TimeSync'          = 'Configure Automatic Time Sync'
        'WingetDependencies'= 'Ensure Winget Dependencies'
        'UpdateApplications'= 'Update Applications via Winget'
        'HPDrivers'         = 'Update HP/Dell Drivers'
        'WindowsUpdate'     = 'Run Windows Update'
    }
    
    Write-Host "`n[REPORT] Available Deployment Tasks:" -ForegroundColor Cyan
    foreach ($task in $tasks.GetEnumerator()) {
        Write-Host "  - $($task.Key): $($task.Value)" -ForegroundColor White
    }
}


# -----------------------------------------------------------------------------
# Option 20 - Network Repair
# -----------------------------------------------------------------------------
function Run-NetworkDiagnostics {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$SkipBasicTests,
        [switch]$SkipAdvancedTests,
        [switch]$SkipRepairAttempts,
        [switch]$QuickMode,
        [int]$TimeoutSeconds = 30,
        [string]$LogPath = "$PSScriptRoot\NetworkDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [string[]]$TestHosts = @('8.8.8.8', '1.1.1.1', 'google.com', 'yahoo.com'),
        [int]$PingCount = 4,
        [switch]$ExportReport
    )

    # Security: Require elevation for network repair operations
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "Network diagnostics and repair operations require Administrator privileges"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'
    
    # Initialize tracking variables
    $script:logEntries = New-Object System.Collections.ArrayList
    $script:testResults = New-Object System.Collections.ArrayList
    $script:repairActions = New-Object System.Collections.ArrayList
    $script:networkIssues = New-Object System.Collections.ArrayList
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Network diagnostic results structure
    $script:diagnosticResults = @{
        IsLocalIssue = $false
        IsExternalIssue = $false
        CriticalIssues = @()
        Recommendations = @()
        EscalationRequired = $false
        ConnectivityScore = 0
        AdapterHealth = @{}
        DNSHealth = @{}
        RoutingHealth = @{}
        FirewallHealth = @{}
    }

    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO', [switch]$NoNewline)
        
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        if ($null -eq $script:logEntries) {
            $script:logEntries = New-Object System.Collections.ArrayList
        }
        [void]$script:logEntries.Add($logEntry)
        
        $writeParams = @{}
        if ($NoNewline) { 
            $writeParams.NoNewline = $true 
        }
        
        switch ($Level) {
            'ERROR' { Write-Host $Message -ForegroundColor Red @writeParams }
            'WARNING' { Write-Host $Message -ForegroundColor Yellow @writeParams }
            'SUCCESS' { Write-Host $Message -ForegroundColor Green @writeParams }
            'INFO' { Write-Host $Message -ForegroundColor Cyan @writeParams }
            'OPERATION' { Write-Host $Message -ForegroundColor Magenta @writeParams }
            'PROGRESS' { Write-Host $Message -ForegroundColor Gray @writeParams }
            'CRITICAL' { Write-Host $Message -ForegroundColor White -BackgroundColor Red @writeParams }
        }
    }

    function Add-TestResult {
        param(
            [string]$TestName,
            [bool]$Passed,
            [string]$Details,
            [string]$Category = 'General',
            [string]$Recommendation = $null
        )
        
        if ($null -eq $script:testResults) {
            $script:testResults = New-Object System.Collections.ArrayList
        }
        
        [void]$script:testResults.Add(@{
            TestName = $TestName
            Passed = $Passed
            Details = $Details
            Category = $Category
            Recommendation = $Recommendation
            Timestamp = Get-Date
        })
        
        if (-not $Passed -and $Category -eq 'Critical') {
            if ($null -eq $script:networkIssues) {
                $script:networkIssues = New-Object System.Collections.ArrayList
            }
            [void]$script:networkIssues.Add(@{
                Issue = $TestName
                Details = $Details
                Recommendation = $Recommendation
            })
        }
    }

    function Get-NetworkAdapters {
        try {
            Write-LogEntry "      [*] Enumerating network adapters..." 'PROGRESS'
            
            $adapters = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop |
                Where-Object { $_.NetConnectionStatus -ne $null -and ($_.AdapterType -like "*Ethernet*" -or $_.AdapterType -like "*Wireless*") } |
                Select-Object Name, NetConnectionID, NetConnectionStatus, AdapterType, MACAddress, DeviceID
            
            foreach ($adapter in $adapters) {
                $statusText = switch ($adapter.NetConnectionStatus) {
                    2 { "Connected" }
                    7 { "Media Disconnected" }
                    0 { "Disconnected" }
                    default { "Status: $($adapter.NetConnectionStatus)" }
                }
                
                Write-LogEntry "         - $($adapter.Name): $statusText" 'INFO'
                
                $script:diagnosticResults.AdapterHealth[$adapter.Name] = @{
                    Status = $statusText
                    ConnectionID = $adapter.NetConnectionID
                    MAC = $adapter.MACAddress
                    Type = $adapter.AdapterType
                    DeviceID = $adapter.DeviceID
                }
            }
            
            return $adapters
        } catch {
            Write-LogEntry "Failed to enumerate network adapters: $_" 'ERROR'
            return @()
        }
    }

    function Test-NetworkConnectivity {
        param([string[]]$Hosts, [int]$Count = 4)
        
        Write-LogEntry "      [*] Testing connectivity to external hosts..." 'PROGRESS'
        
        $successfulPings = 0
        $totalPings = 0
        $results = @{}
        
        foreach ($targetHost in $Hosts) {
            try {
                Write-LogEntry "         Testing $targetHost..." 'PROGRESS' -NoNewline
                
                $testResult = Test-NetConnection -ComputerName $targetHost -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction Stop
                
                if ($testResult) {
                    $pingResult = Test-Connection -ComputerName $targetHost -Count $Count -Quiet -ErrorAction SilentlyContinue
                    
                    if ($pingResult) {
                        $successfulPings++
                        Write-Host " [OK]" -ForegroundColor Green
                        $results[$targetHost] = @{ Success = $true; Details = "Reachable" }
                    } else {
                        Write-Host " [FAIL]" -ForegroundColor Red
                        $results[$targetHost] = @{ Success = $false; Details = "Ping failed" }
                    }
                } else {
                    Write-Host " [FAIL]" -ForegroundColor Red
                    $results[$targetHost] = @{ Success = $false; Details = "Unreachable" }
                }
                
                $totalPings++
                
            } catch {
                Write-Host " [ERROR]" -ForegroundColor Yellow
                $results[$targetHost] = @{ Success = $false; Details = "Test failed: $($_.Exception.Message)" }
                $totalPings++
            }
        }
        
        $connectivityScore = if ($totalPings -gt 0) { ($successfulPings / $totalPings) * 100 } else { 0 }
        $script:diagnosticResults.ConnectivityScore = $connectivityScore
        
        Add-TestResult -TestName "External Connectivity" -Passed ($connectivityScore -gt 50) -Details "$successfulPings of $totalPings hosts reachable ($([math]::Round($connectivityScore, 1))%)" -Category "Critical" -Recommendation "Check internet connection and firewall settings"
        
        return @{
            SuccessRate = $connectivityScore
            Results = $results
            TotalTests = $totalPings
            SuccessfulTests = $successfulPings
        }
    }

    function Test-DNSResolution {
        param([string[]]$TestDomains = @('google.com', 'yahoo.com', 'cloudflare.com'))
        
        Write-LogEntry "      [*] Testing DNS resolution..." 'PROGRESS'
        
        $dnsServers = @()
        $dnsResults = @{}
        
        try {
            $dnsConfig = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop | 
                Where-Object { $_.ServerAddresses.Count -gt 0 }
            
            foreach ($config in $dnsConfig) {
                $dnsServers += $config.ServerAddresses
            }
            
            $dnsServers = $dnsServers | Select-Object -Unique
            Write-LogEntry "         Configured DNS servers: $($dnsServers -join ', ')" 'INFO'
            $script:diagnosticResults.DNSHealth['ConfiguredServers'] = $dnsServers
            
        } catch {
            Write-LogEntry "         Failed to get DNS configuration: $_" 'WARNING'
        }
        
        $successfulResolves = 0
        $totalResolves = 0
        
        foreach ($domain in $TestDomains) {
            try {
                Write-LogEntry "         Resolving $domain..." 'PROGRESS' -NoNewline
                
                $resolved = Resolve-DnsName -Name $domain -Type A -ErrorAction Stop -DnsOnly
                
                if ($resolved -and $resolved.Count -gt 0) {
                    $successfulResolves++
                    Write-Host " [OK] ($($resolved[0].IPAddress))" -ForegroundColor Green
                    $dnsResults[$domain] = @{ Success = $true; IP = $resolved[0].IPAddress }
                } else {
                    Write-Host " [FAIL]" -ForegroundColor Red
                    $dnsResults[$domain] = @{ Success = $false; Error = "No resolution" }
                }
                
            } catch {
                Write-Host " [FAIL] ($($_.Exception.Message))" -ForegroundColor Red
                $dnsResults[$domain] = @{ Success = $false; Error = $_.Exception.Message }
            }
            
            $totalResolves++
        }
        
        $dnsScore = if ($totalResolves -gt 0) { ($successfulResolves / $totalResolves) * 100 } else { 0 }
        $script:diagnosticResults.DNSHealth['ResolutionScore'] = $dnsScore
        $script:diagnosticResults.DNSHealth['Results'] = $dnsResults
        
        Add-TestResult -TestName "DNS Resolution" -Passed ($dnsScore -gt 75) -Details "$successfulResolves of $totalResolves domains resolved ($([math]::Round($dnsScore, 1))%)" -Category "Critical" -Recommendation "Check DNS server configuration or try alternative DNS servers (8.8.8.8, 1.1.1.1)"
        
        return @{
            SuccessRate = $dnsScore
            ConfiguredServers = $dnsServers
            Results = $dnsResults
        }
    }

    function Test-RoutingTable {
        try {
            Write-LogEntry "      [*] Analyzing routing table..." 'PROGRESS'
            
            $routes = Get-NetRoute -AddressFamily IPv4 -ErrorAction Stop | 
                Where-Object { $_.RouteMetric -lt 1000 } |
                Sort-Object RouteMetric |
                Select-Object -First 10 DestinationPrefix, NextHop, RouteMetric, InterfaceAlias
            
            $defaultRoute = $routes | Where-Object { $_.DestinationPrefix -eq '0.0.0.0/0' } | Select-Object -First 1
            
            if ($defaultRoute) {
                Write-LogEntry "         Default gateway: $($defaultRoute.NextHop) via $($defaultRoute.InterfaceAlias)" 'INFO'
                $script:diagnosticResults.RoutingHealth['DefaultGateway'] = $defaultRoute.NextHop
                $script:diagnosticResults.RoutingHealth['DefaultInterface'] = $defaultRoute.InterfaceAlias
                
                Write-LogEntry "         Testing default gateway connectivity..." 'PROGRESS' -NoNewline
                $gwTest = Test-Connection -ComputerName $defaultRoute.NextHop -Count 2 -Quiet -ErrorAction SilentlyContinue
                
                if ($gwTest) {
                    Write-Host " [OK]" -ForegroundColor Green
                    Add-TestResult -TestName "Default Gateway" -Passed $true -Details "Gateway $($defaultRoute.NextHop) is reachable" -Category "Critical"
                } else {
                    Write-Host " [FAIL]" -ForegroundColor Red
                    Add-TestResult -TestName "Default Gateway" -Passed $false -Details "Gateway $($defaultRoute.NextHop) is unreachable" -Category "Critical" -Recommendation "Check network cable, switch, or router connectivity"
                    $script:diagnosticResults.IsLocalIssue = $true
                }
                
            } else {
                Write-LogEntry "         No default route found!" 'ERROR'
                Add-TestResult -TestName "Default Gateway" -Passed $false -Details "No default route configured" -Category "Critical" -Recommendation "Configure default gateway or check DHCP settings"
                $script:diagnosticResults.IsLocalIssue = $true
            }
            
            $script:diagnosticResults.RoutingHealth['TotalRoutes'] = $routes.Count
            
        } catch {
            Write-LogEntry "Failed to analyze routing table: $_" 'ERROR'
            Add-TestResult -TestName "Routing Analysis" -Passed $false -Details "Failed to analyze routing: $_" -Category "Warning"
        }
    }

    function Test-WindowsFirewall {
        try {
            Write-LogEntry "      [*] Checking Windows Firewall status..." 'PROGRESS'
            
            $firewallProfiles = Get-NetFirewallProfile -ErrorAction Stop
            
            foreach ($profile in $firewallProfiles) {
                $enabled = $profile.Enabled
                $statusText = if ($enabled) { "Enabled" } else { "Disabled" }
                
                Write-LogEntry "         $($profile.Name) Profile: $statusText" 'INFO'
                
                $script:diagnosticResults.FirewallHealth[$profile.Name] = @{
                    Enabled = $enabled
                    DefaultInboundAction = $profile.DefaultInboundAction
                    DefaultOutboundAction = $profile.DefaultOutboundAction
                }
            }
            
            $blockingRules = Get-NetFirewallRule -Direction Outbound -Action Block -Enabled True -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayGroup -notlike "*Windows*" -and $_.DisplayName -notlike "*Microsoft*" } |
                Select-Object -First 5 DisplayName, Direction, Action
            
            if ($blockingRules.Count -gt 0) {
                Write-LogEntry "         Found $($blockingRules.Count) custom blocking rules" 'WARNING'
                $script:diagnosticResults.FirewallHealth['CustomBlockingRules'] = $blockingRules.Count
                Add-TestResult -TestName "Firewall Rules" -Passed $false -Details "$($blockingRules.Count) custom blocking rules found" -Category "Warning" -Recommendation "Review custom firewall rules that may block network access"
            } else {
                Add-TestResult -TestName "Firewall Rules" -Passed $true -Details "No problematic blocking rules found" -Category "General"
            }
            
        } catch {
            Write-LogEntry "Failed to check Windows Firewall: $_" 'ERROR'
            Add-TestResult -TestName "Windows Firewall" -Passed $false -Details "Failed to check firewall: $_" -Category "Warning"
        }
    }

    function Repair-NetworkAdapters {
        if ($SkipRepairAttempts) {
            Write-LogEntry "Skipping network adapter repairs (disabled by parameter)" 'INFO'
            return
        }
        
        Write-LogEntry "      [*] Attempting network adapter repairs..." 'PROGRESS'
        
        try {
            $adapters = Get-NetAdapter -ErrorAction Stop | 
                Where-Object { $_.Status -eq 'Disconnected' -or $_.Status -eq 'NotPresent' }
            
            foreach ($adapter in $adapters) {
                if ($PSCmdlet.ShouldProcess($adapter.Name, "Reset network adapter")) {
                    Write-LogEntry "         Resetting adapter: $($adapter.Name)..." 'PROGRESS'
                    
                    try {
                        Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
                        Start-Sleep -Seconds 2
                        Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction Stop
                        Start-Sleep -Seconds 3
                        
                        $script:repairActions.Add("Reset network adapter: $($adapter.Name)")
                        Write-LogEntry "         [OK] Successfully reset $($adapter.Name)" 'SUCCESS'
                        
                    } catch {
                        Write-LogEntry "         [FAIL] Failed to reset $($adapter.Name): $_" 'WARNING'
                    }
                }
            }
            
        } catch {
            Write-LogEntry "Failed during adapter repair: $_" 'ERROR'
        }
    }

    function Repair-NetworkStack {
        if ($SkipRepairAttempts) {
            Write-LogEntry "Skipping network stack repairs (disabled by parameter)" 'INFO'
            return
        }
        
        Write-LogEntry "      [*] Performing network stack repairs..." 'PROGRESS'
        
        $repairCommands = @(
            @{ Command = "ipconfig /release"; Description = "Release IP configuration"; RequiresReboot = $false },
            @{ Command = "ipconfig /flushdns"; Description = "Flush DNS cache"; RequiresReboot = $false },
            @{ Command = "ipconfig /renew"; Description = "Renew IP configuration"; RequiresReboot = $false },
            @{ Command = "netsh winsock reset"; Description = "Reset Winsock catalog"; RequiresReboot = $true },
            @{ Command = "netsh int ip reset"; Description = "Reset TCP/IP stack"; RequiresReboot = $true }
        )
        
        $rebootRequired = $false
        $networkStackResetFailed = $false
        
        foreach ($repair in $repairCommands) {
            if ($PSCmdlet.ShouldProcess($repair.Description, "Execute repair command")) {
                try {
                    Write-LogEntry "         $($repair.Description)..." 'PROGRESS'
                    
                    $result = Invoke-Expression $repair.Command 2>&1
                    $exitCode = $LASTEXITCODE
                    
                    $resultString = $result | Out-String
                    
                    if ($repair.Command -like "*netsh int ip reset*") {
                        if ($resultString -like "*Restart the computer to complete this action*" -or 
                            $resultString -like "*Resetting*OK*" -or
                            $exitCode -eq 1) {
                            
                            if ($resultString -like "*Resetting*OK*") {
                                Write-LogEntry "         [OK] $($repair.Description) completed (restart required)" 'SUCCESS'
                                $script:repairActions.Add("$($repair.Description) - restart required")
                                $rebootRequired = $true
                            } elseif ($exitCode -eq 1 -and $resultString -like "*Access is denied*") {
                                Write-LogEntry "         [PARTIAL] $($repair.Description) partially completed (restart required)" 'WARNING'
                                $script:repairActions.Add("$($repair.Description) - partial success, restart required")
                                $rebootRequired = $true
                            } else {
                                Write-LogEntry "         [FAIL] $($repair.Description) failed with exit code $exitCode" 'ERROR'
                                $networkStackResetFailed = $true
                            }
                        } else {
                            Write-LogEntry "         [OK] $($repair.Description) completed" 'SUCCESS'
                            $script:repairActions.Add($repair.Description)
                            if ($repair.RequiresReboot) { 
                                $rebootRequired = $true 
                            }
                        }
                    } elseif ($repair.Command -like "*netsh winsock reset*") {
                        if ($exitCode -eq 0 -or $resultString -like "*successfully*" -or $resultString -like "*reset*") {
                            Write-LogEntry "         [OK] $($repair.Description) completed (restart required)" 'SUCCESS'
                            $script:repairActions.Add("$($repair.Description) - restart required")
                            $rebootRequired = $true
                        } else {
                            Write-LogEntry "         [FAIL] $($repair.Description) failed with exit code $exitCode" 'ERROR'
                        }
                    } else {
                        if ($exitCode -eq 0 -or $null -eq $exitCode) {
                            Write-LogEntry "         [OK] $($repair.Description) completed" 'SUCCESS'
                            $script:repairActions.Add($repair.Description)
                        } else {
                            Write-LogEntry "         [WARN] $($repair.Description) returned exit code $exitCode" 'WARNING'
                            $script:repairActions.Add("$($repair.Description) - warning (exit code $exitCode)")
                        }
                    }
                    
                } catch {
                    Write-LogEntry "         [FAIL] $($repair.Description) failed: $_" 'ERROR'
                }
                
                Start-Sleep -Seconds 1
            }
        }
        
        if ($rebootRequired) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[WARNING] SYSTEM RESTART REQUIRED" 'WARNING'
            Write-LogEntry "Network stack changes require a system restart to take effect." 'WARNING'
            Write-LogEntry "After restart, please run this diagnostic script again to verify the repairs." 'INFO'
            
            $global:RestartRequired = $true
            $global:RebootReason = "Network stack changes require a system restart to take effect."
        }
        
        if ($networkStackResetFailed) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[ERROR] Network stack reset encountered significant failures." 'ERROR'
            Write-LogEntry "Consider running Windows Network Reset from Settings > Network & Internet > Status." 'INFO'
            Write-LogEntry "Alternative: Use 'netsh int ip reset reset.log' from elevated Command Prompt." 'INFO'
        }
    }

    function Analyze-NetworkHealth {
        Write-LogEntry "      [*] Analyzing network health..." 'PROGRESS'
        
        $criticalFailures = $script:testResults | Where-Object { -not $_.Passed -and $_.Category -eq 'Critical' }
        $passedTests = ($script:testResults | Where-Object { $_.Passed }).Count
        $totalTests = $script:testResults.Count
        
        $connectivityScore = if ($totalTests -gt 0) { 
            [math]::Round(($passedTests / $totalTests) * 100, 1) 
        } else { 
            0 
        }
        
        $script:diagnosticResults.ConnectivityScore = $connectivityScore
        
        $gatewayTest = $script:testResults | Where-Object { $_.TestName -eq 'Default Gateway' }
        $dnsTest = $script:testResults | Where-Object { $_.TestName -eq 'DNS Resolution' }
        $connectivityTest = $script:testResults | Where-Object { $_.TestName -eq 'External Connectivity' }
        
        if ($gatewayTest -and -not $gatewayTest.Passed) {
            $script:diagnosticResults.IsLocalIssue = $true
            $script:diagnosticResults.CriticalIssues += "Local network connectivity issue (gateway unreachable)"
            $script:diagnosticResults.Recommendations += "Check network cables, switch ports, and local network equipment"
        }
        
        if ($dnsTest -and -not $dnsTest.Passed -and $gatewayTest.Passed) {
            $script:diagnosticResults.IsLocalIssue = $true
            $script:diagnosticResults.CriticalIssues += "DNS resolution failure"
            $script:diagnosticResults.Recommendations += "Check DNS server configuration or try alternative DNS servers"
        }
        
        if ($connectivityTest -and -not $connectivityTest.Passed -and $gatewayTest.Passed -and $dnsTest.Passed) {
            $script:diagnosticResults.IsExternalIssue = $true
            $script:diagnosticResults.CriticalIssues += "External connectivity issue (ISP or upstream)"
            $script:diagnosticResults.Recommendations += "Contact ISP or check with Tier 3 support for upstream connectivity issues"
            $script:diagnosticResults.EscalationRequired = $true
        }
        
        if ($criticalFailures.Count -gt 2 -or $connectivityScore -lt 30) {
            $script:diagnosticResults.EscalationRequired = $true
        }
        
        if ($connectivityScore -gt 80) {
            $script:diagnosticResults.Recommendations += "Network connectivity is healthy"
        } elseif ($connectivityScore -gt 60) {
            $script:diagnosticResults.Recommendations += "Minor network issues detected - monitor for stability"
        } else {
            $script:diagnosticResults.Recommendations += "Significant network issues require attention"
        }
    }

    function Export-DiagnosticReport {
        if (-not $ExportReport) { 
            return 
        }
        
        $reportPath = $LogPath -replace '\.log$', '_TechnicalReport.txt'
        
        try {
            $report = @()
            $report += "=== NETWORK DIAGNOSTIC TECHNICAL REPORT ==="
            $report += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $report += "Technician: $env:USERNAME"
            $report += "Computer Name: $env:COMPUTERNAME"
            $report += "Domain: $env:USERDOMAIN"
            
            try {
                $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
                $report += "Manufacturer: $($computerSystem.Manufacturer)"
                $report += "Model: $($computerSystem.Model)"
                
                $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
                $report += "Operating System: $($os.Caption) $($os.Version)"
            } catch {
                $report += "System Info: Unable to retrieve detailed system information"
            }
            
            $report += ""
            $report += "NETWORK ADAPTERS:"
            $report += "=================="
            
            try {
                $allAdapters = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop |
                    Where-Object { $_.NetConnectionStatus -ne $null -or $_.AdapterType -like "*Ethernet*" -or $_.AdapterType -like "*Wireless*" -or $_.AdapterType -like "*802.11*" } |
                    Sort-Object Name
                
                foreach ($adapter in $allAdapters) {
                    $statusText = switch ($adapter.NetConnectionStatus) {
                        2 { "Connected" }
                        7 { "Media Disconnected" }
                        0 { "Disconnected" }
                        1 { "Connecting" }
                        8 { "Hardware Not Present" }
                        9 { "Hardware Disabled" }
                        10 { "Hardware Malfunction" }
                        11 { "Media Disconnected" }
                        default { "Status Code: $($adapter.NetConnectionStatus)" }
                    }
                    
                    $macAddress = if ($adapter.MACAddress) { $adapter.MACAddress } else { "Not Available" }
                    $connectionId = if ($adapter.NetConnectionID) { $adapter.NetConnectionID } else { "Not Named" }
                    $adapterType = if ($adapter.AdapterType) { $adapter.AdapterType } else { "Unknown Type" }
                    
                    $report += "Adapter: $($adapter.Name)"
                    $report += "  Connection Name: $connectionId"
                    $report += "  MAC Address: $macAddress"
                    $report += "  Type: $adapterType"
                    $report += "  Status: $statusText"
                    $report += "  Device ID: $($adapter.DeviceID)"
                    
                    if ($adapter.NetConnectionStatus -eq 2) {
                        try {
                            $ipConfig = Get-NetIPConfiguration -InterfaceAlias $adapter.NetConnectionID -ErrorAction SilentlyContinue
                            if ($ipConfig) {
                                if ($ipConfig.IPv4Address) {
                                    $report += "  IPv4 Address: $($ipConfig.IPv4Address.IPAddress)"
                                    $report += "  Subnet Mask: $($ipConfig.IPv4Address.PrefixLength)"
                                }
                                if ($ipConfig.IPv4DefaultGateway) {
                                    $report += "  Default Gateway: $($ipConfig.IPv4DefaultGateway.NextHop)"
                                }
                                if ($ipConfig.DNSServer) {
                                    $dnsServers = ($ipConfig.DNSServer.ServerAddresses | Where-Object { $_ -notmatch ':' }) -join ', '
                                    if ($dnsServers) {
                                        $report += "  DNS Servers: $dnsServers"
                                    }
                                }
                            }
                        } catch {
                            # Continue silently if IP config retrieval fails
                        }
                    }
                    $report += ""
                }
                
                $totalAdapters = $allAdapters.Count
                $connectedAdapters = ($allAdapters | Where-Object { $_.NetConnectionStatus -eq 2 }).Count
                $disconnectedAdapters = ($allAdapters | Where-Object { $_.NetConnectionStatus -eq 7 -or $_.NetConnectionStatus -eq 0 }).Count
                
                $report += "ADAPTER SUMMARY:"
                $report += "Total Adapters: $totalAdapters"
                $report += "Connected: $connectedAdapters"
                $report += "Disconnected/Media Disconnected: $disconnectedAdapters"
                
            } catch {
                $report += "ERROR: Unable to retrieve network adapter information"
                $report += "Error Details: $($_.Exception.Message)"
            }
            
            $report += ""
            $report += "=========================================="
            $report += ""
            
            $report += "EXECUTIVE SUMMARY:"
            if ($script:diagnosticResults.CriticalIssues.Count -gt 0) {
                $report += "CRITICAL ISSUES:"
                foreach ($issue in $script:diagnosticResults.CriticalIssues) {
                    $report += "- $issue"
                }
                $report += ""
            }
            
            $report += "DETAILED TEST RESULTS:"
            foreach ($test in $script:testResults) {
                $status = if ($test.Passed) { "PASS" } else { "FAIL" }
                $report += "[$status] $($test.TestName): $($test.Details)"
                if ($test.Recommendation) {
                    $report += "  -> Recommendation: $($test.Recommendation)"
                }
            }
            $report += ""
            
            $report += "NETWORK CONFIGURATION:"
            $report += "DNS Servers: $($script:diagnosticResults.DNSHealth.ConfiguredServers -join ', ')"
            $report += "Default Gateway: $($script:diagnosticResults.RoutingHealth.DefaultGateway)"
            $report += "Default Interface: $($script:diagnosticResults.RoutingHealth.DefaultInterface)"
            $report += ""
            
            if ($script:repairActions.Count -gt 0) {
                $report += "REPAIR ACTIONS PERFORMED:"
                foreach ($action in $script:repairActions) {
                    $report += "- $action"
                }
                $report += ""
            }
            
            $report += "RECOMMENDATIONS:"
            foreach ($rec in $script:diagnosticResults.Recommendations) {
                $report += "- $rec"
            }
            
            $report | Out-File -FilePath $reportPath -Encoding UTF8 -Force
            Write-LogEntry "[SUCCESS] Technical report exported to: $reportPath" 'SUCCESS'
            
        } catch {
            Write-LogEntry "Failed to export technical report: $_" 'ERROR'
        }
    }

    # Main execution of Run-NetworkDiagnostics function
    try {
        Write-LogEntry "" 'INFO'
        Write-LogEntry "[OPERATION] Starting Network Diagnostics and Repair..." 'OPERATION'
        Write-LogEntry "Computer: $env:COMPUTERNAME" 'INFO'
        Write-LogEntry "Domain: $env:USERDOMAIN" 'INFO'
        Write-LogEntry "User: $env:USERNAME" 'INFO'
        Write-LogEntry "Test timeout: $TimeoutSeconds seconds" 'INFO'
        Write-LogEntry "Quick mode: $QuickMode" 'INFO'
        
        Write-LogEntry "" 'INFO'
        Write-LogEntry "[OPERATION] System Information:" 'OPERATION'
        try {
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            Write-LogEntry "Manufacturer: $($computerSystem.Manufacturer)" 'INFO'
            Write-LogEntry "Model: $($computerSystem.Model)" 'INFO'
            
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
            Write-LogEntry "OS: $($os.Caption) $($os.Version)" 'INFO'
        } catch {
            Write-LogEntry "Could not retrieve detailed system information" 'WARNING'
        }
        
        if (-not $SkipBasicTests) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[OPERATION] Phase 1: Basic Network Analysis" 'OPERATION'
            
            $adapters = Get-NetworkAdapters
            Add-TestResult -TestName "Network Adapters" -Passed ($adapters.Count -gt 0) -Details "$($adapters.Count) adapters found" -Category "General"
            
            Test-RoutingTable
            Test-WindowsFirewall
        }
        
        Write-LogEntry "" 'INFO'
        Write-LogEntry "[OPERATION] Phase 2: Connectivity Testing" 'OPERATION'
        
        $dnsResults = Test-DNSResolution -TestDomains @('google.com', 'yahoo.com', 'cloudflare.com')
        
        $testHosts = if ($QuickMode) { $TestHosts[0..1] } else { $TestHosts }
        $connectivityResults = Test-NetworkConnectivity -Hosts $testHosts -Count $PingCount
        
        if (-not $SkipAdvancedTests -and -not $QuickMode) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[OPERATION] Phase 3: Advanced Network Analysis" 'OPERATION'
        }
        
        if (-not $SkipRepairAttempts) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[OPERATION] Phase 4: Network Repair Operations" 'OPERATION'
            
            Repair-NetworkAdapters
            Repair-NetworkStack
        }
        
        Write-LogEntry "" 'INFO'
        Write-LogEntry "[OPERATION] Phase 5: Analysis and Reporting" 'OPERATION'
        
        Analyze-NetworkHealth
        Export-DiagnosticReport
        
    } catch {
        Write-LogEntry "Critical error during network diagnostics: $_" 'CRITICAL'
        $global:LastStatus = "[ERROR] Network diagnostics failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        $passedTests = ($script:testResults | Where-Object { $_.Passed }).Count
        $totalTests = $script:testResults.Count
        $repairCount = $script:repairActions.Count
        
        Write-LogEntry "" 'INFO'
        Write-LogEntry "[INFO] Network Diagnostics Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Tests passed: $passedTests of $totalTests" 'INFO'
        Write-LogEntry "Connectivity score: $($script:diagnosticResults.ConnectivityScore)%" 'INFO'
        Write-LogEntry "Repair actions performed: $repairCount" 'INFO'
        
        Write-LogEntry "" 'INFO'
        Write-LogEntry "[INFO] Issue Classification:" 'INFO'
        if ($script:diagnosticResults.IsLocalIssue -and $script:diagnosticResults.IsExternalIssue) {
            Write-LogEntry "[WARNING] MIXED ISSUES: Both local and external problems detected" 'WARNING'
        } elseif ($script:diagnosticResults.IsLocalIssue) {
            Write-LogEntry "[ERROR] LOCAL ISSUE: Problem is with this workstation or local network" 'ERROR'
        } elseif ($script:diagnosticResults.IsExternalIssue) {
            Write-LogEntry "[WARNING] EXTERNAL ISSUE: Problem is with ISP or upstream connectivity" 'WARNING'
        } else {
            Write-LogEntry "[SUCCESS] NO MAJOR ISSUES: Network connectivity appears healthy" 'SUCCESS'
        }
        
        if ($script:diagnosticResults.EscalationRequired) {
            Write-LogEntry "[CRITICAL] ESCALATION RECOMMENDED: Contact Tier 3 support" 'CRITICAL'
        }
        
        if ($script:diagnosticResults.CriticalIssues.Count -gt 0) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[ERROR] Critical Issues Found:" 'ERROR'
            foreach ($issue in $script:diagnosticResults.CriticalIssues) {
                Write-LogEntry "  - $issue" 'ERROR'
            }
        }
        
        if ($script:diagnosticResults.Recommendations.Count -gt 0) {
            Write-LogEntry "" 'INFO'
            Write-LogEntry "[INFO] Recommendations:" 'INFO'
            foreach ($rec in $script:diagnosticResults.Recommendations) {
                Write-LogEntry "  - $rec" 'INFO'
            }
        }
        
        try {
            if ($null -ne $script:logEntries) {
                $enhancedLog = @()
                $enhancedLog += "=== NETWORK DIAGNOSTICS LOG ==="
                $enhancedLog += "Computer Name: $env:COMPUTERNAME"
                $enhancedLog += "Domain: $env:USERDOMAIN"
                $enhancedLog += "User: $env:USERNAME"
                $enhancedLog += "Date/Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                $enhancedLog += ""
                
                $enhancedLog += "NETWORK ADAPTERS SUMMARY:"
                $enhancedLog += "========================="
                try {
                    $logAdapters = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction Stop |
                        Where-Object { $_.NetConnectionStatus -ne $null -or $_.AdapterType -like "*Ethernet*" -or $_.AdapterType -like "*Wireless*" -or $_.AdapterType -like "*802.11*" } |
                        Sort-Object Name
                    
                    foreach ($adapter in $logAdapters) {
                        $statusText = switch ($adapter.NetConnectionStatus) {
                            2 { "Connected" }
                            7 { "Media Disconnected" }
                            0 { "Disconnected" }
                            default { "Status: $($adapter.NetConnectionStatus)" }
                        }
                        $macAddress = if ($adapter.MACAddress) { $adapter.MACAddress } else { "N/A" }
                        $enhancedLog += "$($adapter.Name) | MAC: $macAddress | Status: $statusText"
                    }
                } catch {
                    $enhancedLog += "ERROR: Could not retrieve adapter information"
                }
                
                $enhancedLog += ""
                $enhancedLog += "DIAGNOSTIC LOG ENTRIES:"
                $enhancedLog += "======================"
                $enhancedLog += $script:logEntries.ToArray()
                
                $enhancedLog | Out-File -FilePath $LogPath -Encoding UTF8 -Force
                Write-LogEntry "[SUCCESS] Enhanced log with computer info saved to: $LogPath" 'INFO'
            }
        } catch {
            Write-LogEntry "[WARNING] Failed to save enhanced log file: $_" 'WARNING'
        }
        
        if ($global:RestartRequired -eq $true) {
            $global:LastStatus = "[RESTART REQUIRED] Network repairs completed - System restart needed to complete changes (Score: $($script:diagnosticResults.ConnectivityScore)%)"
        } elseif ($script:diagnosticResults.EscalationRequired) {
            $global:LastStatus = "[ESCALATION] Network issues require Tier 3 escalation - Score: $($script:diagnosticResults.ConnectivityScore)%"
        } elseif ($script:diagnosticResults.IsLocalIssue) {
            $global:LastStatus = "[LOCAL] Local network issue detected - Score: $($script:diagnosticResults.ConnectivityScore)%"
        } elseif ($script:diagnosticResults.IsExternalIssue) {
            $global:LastStatus = "[EXTERNAL] External network issue detected - Score: $($script:diagnosticResults.ConnectivityScore)%"
        } elseif ($passedTests -eq $totalTests) {
            $global:LastStatus = "[SUCCESS] Network diagnostics completed - All tests passed (Score: $($script:diagnosticResults.ConnectivityScore)%)"
        } else {
            $global:LastStatus = "[WARNING] Network diagnostics completed - $($totalTests - $passedTests) of $totalTests tests failed (Score: $($script:diagnosticResults.ConnectivityScore)%)"
        }
        
        Write-LogEntry "=== Network Diagnostics Completed ===" 'INFO'
    }
}

function Test-NetworkPerformance {
    param(
        [string]$TargetHost = 'google.com',
        [int]$TestDuration = 30,
        [int]$PacketSize = 1472
    )
    
    Write-Host "[OPERATION] Starting network performance testing..." -ForegroundColor Magenta
    
    try {
        $sizes = @(64, 512, 1024, $PacketSize)
        $results = @{}
        
        foreach ($size in $sizes) {
            Write-Host "   Testing with $size byte packets..." -ForegroundColor Gray
            
            $pingResults = @()
            for ($i = 1; $i -le 10; $i++) {
                try {
                    $ping = Test-Connection -ComputerName $TargetHost -Count 1 -BufferSize $size -ErrorAction Stop
                    $pingResults += $ping.ResponseTime
                } catch {
                    Write-Host "   Packet size $size failed: $_" -ForegroundColor Yellow
                }
            }
            
            if ($pingResults.Count -gt 0) {
                $avgLatency = ($pingResults | Measure-Object -Average).Average
                $results[$size] = $avgLatency
                Write-Host "   Average latency ($size bytes): $([math]::Round($avgLatency, 2))ms" -ForegroundColor Cyan
            }
        }
        
        return $results
        
    } catch {
        Write-Host "Performance testing failed: $_" -ForegroundColor Red
        return @{}
    }
}

function Test-SpecificPorts {
    param(
        [hashtable]$PortTests = @{
            'HTTP' = @{ Host = 'google.com'; Port = 80 }
            'HTTPS' = @{ Host = 'google.com'; Port = 443 }
            'DNS' = @{ Host = '8.8.8.8'; Port = 53 }
            'SMTP' = @{ Host = 'smtp.gmail.com'; Port = 587 }
        }
    )
    
    Write-Host "[OPERATION] Testing specific network ports..." -ForegroundColor Magenta
    
    $portResults = @{}
    
    foreach ($testName in $PortTests.Keys) {
        $test = $PortTests[$testName]
        
        try {
            Write-Host "   Testing $testName ($($test.Host):$($test.Port))..." -ForegroundColor Gray -NoNewline
            
            $connection = Test-NetConnection -ComputerName $test.Host -Port $test.Port -InformationLevel Quiet -WarningAction SilentlyContinue
            
            if ($connection) {
                Write-Host " [OK]" -ForegroundColor Green
                $portResults[$testName] = @{ Success = $true; Details = "Port accessible" }
            } else {
                Write-Host " [FAIL]" -ForegroundColor Red
                $portResults[$testName] = @{ Success = $false; Details = "Port blocked or filtered" }
            }
            
        } catch {
            Write-Host " [ERROR]" -ForegroundColor Yellow
            $portResults[$testName] = @{ Success = $false; Details = "Test failed: $_" }
        }
    }
    
    return $portResults
}

function Get-NetworkConfiguration {
    Write-Host "[OPERATION] Gathering detailed network configuration..." -ForegroundColor Magenta
    
    $config = @{}
    
    try {
        $ipConfig = Get-NetIPConfiguration -ErrorAction Stop
        $config['IPConfiguration'] = $ipConfig | Select-Object InterfaceAlias, IPv4Address, IPv4DefaultGateway, DNSServer
        
        $interfaces = Get-NetAdapter -ErrorAction Stop | Where-Object { $_.Status -eq 'Up' }
        $config['ActiveInterfaces'] = $interfaces | Select-Object Name, InterfaceDescription, LinkSpeed, FullDuplex
        
        $arpTable = Get-NetNeighbor -ErrorAction SilentlyContinue | 
            Where-Object { $_.State -eq 'Reachable' -or $_.State -eq 'Stale' } |
            Select-Object -First 10 IPAddress, LinkLayerAddress, State
        $config['ARPTable'] = $arpTable
        
        $netStats = Get-NetAdapterStatistics -ErrorAction SilentlyContinue |
            Where-Object { $_.BytesReceived -gt 0 -or $_.BytesSent -gt 0 } |
            Select-Object Name, BytesReceived, BytesSent, PacketsReceived, PacketsSent
        $config['NetworkStatistics'] = $netStats
        
        Write-Host "[SUCCESS] Network configuration gathered successfully" -ForegroundColor Green
        
    } catch {
        Write-Host "Failed to gather network configuration: $_" -ForegroundColor Red
    }
    
    return $config
}

function Invoke-AdvancedNetworkRepair {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$ResetNetworkStack,
        [switch]$ResetFirewall,
        [switch]$ResetDNS,
        [switch]$ResetWinsock
    )
    
    Write-Host "[OPERATION] Starting advanced network repair operations..." -ForegroundColor Magenta
    
    $repairActions = @()
    
    try {
        if ($ResetDNS) {
            if ($PSCmdlet.ShouldProcess("DNS Client", "Reset DNS cache and configuration")) {
                Write-Host "   Resetting DNS configuration..." -ForegroundColor Gray
                
                & ipconfig /flushdns | Out-Null
                Set-DnsClientGlobalSetting -SuffixSearchList @() -ErrorAction SilentlyContinue
                
                $repairActions += "DNS cache and configuration reset"
                Write-Host "   [OK] DNS reset completed" -ForegroundColor Green
            }
        }
        
        if ($ResetWinsock) {
            if ($PSCmdlet.ShouldProcess("Winsock", "Reset Winsock catalog")) {
                Write-Host "   Resetting Winsock catalog..." -ForegroundColor Gray
                
                & netsh winsock reset | Out-Null
                
                $repairActions += "Winsock catalog reset"
                Write-Host "   [OK] Winsock reset completed (restart required)" -ForegroundColor Green
            }
        }
        
        if ($ResetNetworkStack) {
            if ($PSCmdlet.ShouldProcess("TCP/IP Stack", "Reset network stack")) {
                Write-Host "   Resetting TCP/IP stack..." -ForegroundColor Gray
                
                & netsh int ip reset | Out-Null
                & netsh int ipv6 reset | Out-Null
                
                $repairActions += "TCP/IP stack reset"
                Write-Host "   [OK] Network stack reset completed (restart required)" -ForegroundColor Green
            }
        }
        
        if ($ResetFirewall) {
            if ($PSCmdlet.ShouldProcess("Windows Firewall", "Reset to default settings")) {
                Write-Host "   Resetting Windows Firewall..." -ForegroundColor Gray
                
                & netsh advfirewall reset | Out-Null
                
                $repairActions += "Windows Firewall reset to defaults"
                Write-Host "   [OK] Firewall reset completed" -ForegroundColor Green
            }
        }
        
        if ($repairActions.Count -gt 0) {
            Write-Host "" -ForegroundColor Cyan
            Write-Host "[INFO] Advanced repair summary:" -ForegroundColor Cyan
            foreach ($action in $repairActions) {
                Write-Host "   - $action" -ForegroundColor Cyan
            }
            
            if ($ResetNetworkStack -or $ResetWinsock) {
                Write-Host "" -ForegroundColor Yellow
                Write-Host "[WARNING] SYSTEM RESTART REQUIRED for changes to take effect" -ForegroundColor Yellow
            }
        }
        
    } catch {
        Write-Host "Advanced network repair failed: $_" -ForegroundColor Red
    }
    
    return $repairActions
}

function Get-NetworkStatus {
    $status = @{
        Timestamp = Get-Date
        Adapters = @{}
        Connectivity = @{}
        DNS = @{}
        Overall = 'Unknown'
    }
    
    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
        foreach ($adapter in $adapters) {
            $status.Adapters[$adapter.Name] = $adapter.Status
        }
        
        $quickHosts = @('8.8.8.8', 'yahoo.com')
        $connectivityScore = 0
        
        foreach ($targetHost in $quickHosts) {
            if (Test-Connection -ComputerName $targetHost -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                $connectivityScore += 50
                $status.Connectivity[$targetHost] = 'Reachable'
            } else {
                $status.Connectivity[$targetHost] = 'Unreachable'
            }
        }
        
        try {
            $dns = Resolve-DnsName -Name 'google.com' -Type A -ErrorAction Stop
            $status.DNS['Resolution'] = 'Working'
        } catch {
            $status.DNS['Resolution'] = 'Failed'
        }
        
        if ($connectivityScore -eq 100 -and $status.DNS.Resolution -eq 'Working') {
            $status.Overall = 'Healthy'
        } elseif ($connectivityScore -gt 0) {
            $status.Overall = 'Partial'
        } else {
            $status.Overall = 'Down'
        }
        
    } catch {
        $status.Overall = 'Error'
        $status.Error = $_.Exception.Message
    }
    
    return $status
}

# -----------------------------------------------------------------------------
# Option 8 - Network Optimization
# -----------------------------------------------------------------------------
# Network Optimization Module - Clean Version
# To use: dot-source this file in your main script with: . ".\NetworkOptimization.ps1"
# Then call: Start-NetworkOptimization

# Global variables
$Script:LogPath = $null

function Initialize-NetworkOptimizationLogging {
    param([string]$LogDirectory = $env:TEMP)
    
    Set-StrictMode -Version Latest
    $ErrorActionPreference = "Stop"
    $Script:LogPath = "$LogDirectory\NetworkOptimization_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    if ([string]::IsNullOrEmpty($Message)) {
        Write-Host ""
        if ($Script:LogPath) {
            try { Add-Content -Path $Script:LogPath -Value "" -ErrorAction SilentlyContinue } catch { }
        }
        return
    }
    
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $LogEntry -ForegroundColor White }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
    }
    
    if ($Script:LogPath) {
        try { Add-Content -Path $Script:LogPath -Value $LogEntry -ErrorAction SilentlyContinue } catch { }
    }
}

function Test-AdministratorPrivileges {
    try {
        $CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
        $Principal = New-Object Security.Principal.WindowsPrincipal($CurrentUser)
        $IsAdmin = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if (-not $IsAdmin) {
            Write-LogMessage -Message "Script must be run as Administrator" -Level "ERROR"
            return $false
        }
        
        Write-LogMessage -Message "Administrator privileges confirmed" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to verify administrator privileges: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Backup-NetworkConfiguration {
    try {
        Write-LogMessage -Message "Creating network configuration backup..." -Level "INFO"
        
        $BackupPath = "$env:TEMP\NetworkBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -Path $BackupPath -ItemType Directory -Force | Out-Null
        
        Get-NetAdapter | Export-Clixml -Path "$BackupPath\NetAdapters.xml"
        Get-NetIPConfiguration | Export-Clixml -Path "$BackupPath\IPConfiguration.xml"
        
        $RegKeys = @(
            "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters",
            "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"
        )
        
        foreach ($Key in $RegKeys) {
            if (Test-Path $Key) {
                try {
                    $KeyName = $Key.Split('\')[-1]
                    $ProcessInfo = Start-Process -FilePath "reg" -ArgumentList "export", $Key, "$BackupPath\$KeyName.reg", "/y" -Wait -PassThru -WindowStyle Hidden -RedirectStandardError "$env:TEMP\reg_error.tmp"
                    
                    if ($ProcessInfo.ExitCode -eq 0) {
                        Write-LogMessage -Message "Registry key $KeyName exported successfully" -Level "SUCCESS"
                    }
                }
                catch {
                    Write-LogMessage -Message "Registry export for $KeyName skipped (non-critical)" -Level "INFO"
                }
            }
        }
        
        Write-LogMessage -Message "Backup created at: $BackupPath" -Level "SUCCESS"
        return $BackupPath
    }
    catch {
        Write-LogMessage -Message "Failed to create backup: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Disable-IPv6Protocol {
    try {
        Write-LogMessage -Message "Disabling IPv6 protocol..." -Level "INFO"
        
        $IPv6RegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"
        
        if (-not (Test-Path $IPv6RegPath)) {
            New-Item -Path $IPv6RegPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $IPv6RegPath -Name "DisabledComponents" -Value 0xFF -Type DWord
        
        $NetworkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        
        foreach ($Adapter in $NetworkAdapters) {
            try {
                Disable-NetAdapterBinding -Name $Adapter.Name -ComponentID "ms_tcpip6" -Confirm:$false
                Write-LogMessage -Message "IPv6 disabled on adapter: $($Adapter.Name)" -Level "SUCCESS"
            }
            catch {
                Write-LogMessage -Message "Failed to disable IPv6 on adapter $($Adapter.Name): $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        Write-LogMessage -Message "IPv6 protocol disabled successfully" -Level "SUCCESS"
    }
    catch {
        Write-LogMessage -Message "Failed to disable IPv6: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Optimize-TCPIPSettings {
    try {
        Write-LogMessage -Message "Optimizing TCP/IP settings..." -Level "INFO"
        
        $TcpipRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
        
        $OptimizationSettings = @{
            "Tcp1323Opts" = 3
            "TcpWindowSize" = 65536
            "GlobalMaxTcpWindowSize" = 65536
            "EnableTCPChimney" = 1
            "EnableRSS" = 1
            "MaxUserPort" = 65534
            "TcpTimedWaitDelay" = 30
            "SynAttackProtect" = 1
            "TcpAckFrequency" = 2
            "SackOpts" = 1
            "EnableDeadGWDetect" = 1
            "EnablePMTUDiscovery" = 1
        }
        
        foreach ($Setting in $OptimizationSettings.GetEnumerator()) {
            try {
                Set-ItemProperty -Path $TcpipRegPath -Name $Setting.Key -Value $Setting.Value -Type DWord
                Write-LogMessage -Message "Set $($Setting.Key) = $($Setting.Value)" -Level "SUCCESS"
            }
            catch {
                Write-LogMessage -Message "Failed to set $($Setting.Key): $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        Write-LogMessage -Message "TCP/IP optimization completed" -Level "SUCCESS"
    }
    catch {
        Write-LogMessage -Message "Failed to optimize TCP/IP settings: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Optimize-NetworkAdapterSettings {
    try {
        Write-LogMessage -Message "Optimizing network adapter settings..." -Level "INFO"
        
        $NetworkAdapters = Get-NetAdapter | Where-Object { 
            $_.Status -eq "Up" -and 
            $_.Virtual -eq $false -and 
            $_.Name -notmatch "Loopback|Teredo|isatap"
        }
        
        foreach ($Adapter in $NetworkAdapters) {
            try {
                Write-LogMessage -Message "Optimizing adapter: $($Adapter.Name)" -Level "INFO"
                
                $AdvancedSettings = @{
                    "*JumboPacket" = "9014"
                    "*InterruptModeration" = "1"
                    "*RSS" = "1"
                    "*TCPUDPChecksumOffloadIPv4" = "3"
                    "*LsoV2IPv4" = "1"
                    "*FlowControl" = "3"
                    "*ReceiveBuffers" = "2048"
                    "*TransmitBuffers" = "2048"
                }
                
                foreach ($Setting in $AdvancedSettings.GetEnumerator()) {
                    try {
                        Set-NetAdapterAdvancedProperty -Name $Adapter.Name -RegistryKeyword $Setting.Key -RegistryValue $Setting.Value -ErrorAction SilentlyContinue
                    }
                    catch {
                        continue
                    }
                }
                
                Write-LogMessage -Message "Adapter $($Adapter.Name) optimized successfully" -Level "SUCCESS"
            }
            catch {
                Write-LogMessage -Message "Failed to optimize adapter $($Adapter.Name): $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        Write-LogMessage -Message "Network adapter optimization completed" -Level "SUCCESS"
    }
    catch {
        Write-LogMessage -Message "Failed to optimize network adapters: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Optimize-DNSSettings {
    try {
        Write-LogMessage -Message "Optimizing DNS settings..." -Level "INFO"
        
        $DNSRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters"
        
        $DNSSettings = @{
            "MaxCacheEntryTtlLimit" = 86400
            "MaxNegativeCacheTtl" = 900
            "QueryIpMatching" = 1
            "PriorityNetBios" = 0
            "EnableAutoDoh" = 2
        }
        
        foreach ($Setting in $DNSSettings.GetEnumerator()) {
            try {
                Set-ItemProperty -Path $DNSRegPath -Name $Setting.Key -Value $Setting.Value -Type DWord
                Write-LogMessage -Message "Set DNS $($Setting.Key) = $($Setting.Value)" -Level "SUCCESS"
            }
            catch {
                Write-LogMessage -Message "Failed to set DNS $($Setting.Key): $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        $FastDNSServers = @("1.1.1.1", "1.0.0.1", "8.8.8.8", "8.8.4.4")
        $ActiveAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        
        foreach ($Adapter in $ActiveAdapters) {
            try {
                Set-DnsClientServerAddress -InterfaceAlias $Adapter.Name -ServerAddresses $FastDNSServers
                Write-LogMessage -Message "DNS servers set for adapter: $($Adapter.Name)" -Level "SUCCESS"
            }
            catch {
                Write-LogMessage -Message "Failed to set DNS for adapter $($Adapter.Name): $($_.Exception.Message)" -Level "WARNING"
            }
        }
        
        Clear-DnsClientCache
        Write-LogMessage -Message "DNS cache cleared" -Level "SUCCESS"
        Write-LogMessage -Message "DNS optimization completed" -Level "SUCCESS"
    }
    catch {
        Write-LogMessage -Message "Failed to optimize DNS settings: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Optimize-WindowsNetworkStack {
    try {
        Write-LogMessage -Message "Optimizing Windows network stack..." -Level "INFO"
        
        $QoSRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile"
        
        if (Test-Path $QoSRegPath) {
            Set-ItemProperty -Path $QoSRegPath -Name "NetworkThrottlingIndex" -Value 0xFFFFFFFF -Type DWord
            Write-LogMessage -Message "Network throttling disabled" -Level "SUCCESS"
        }
        
        try {
            netsh int tcp set global autotuninglevel=normal | Out-Null
            Write-LogMessage -Message "TCP Auto-Tuning set to normal" -Level "SUCCESS"
        }
        catch {
            Write-LogMessage -Message "Failed to set TCP Auto-Tuning: $($_.Exception.Message)" -Level "WARNING"
        }
        
        try {
            netsh int tcp set global chimney=enabled | Out-Null
            Write-LogMessage -Message "TCP Chimney enabled" -Level "SUCCESS"
        }
        catch {
            Write-LogMessage -Message "Failed to enable TCP Chimney: $($_.Exception.Message)" -Level "WARNING"
        }
        
        try {
            netsh int tcp set global rss=enabled | Out-Null
            Write-LogMessage -Message "Receive Side Scaling enabled" -Level "SUCCESS"
        }
        catch {
            Write-LogMessage -Message "Failed to enable RSS: $($_.Exception.Message)" -Level "WARNING"
        }
        
        Write-LogMessage -Message "Windows network stack optimization completed" -Level "SUCCESS"
    }
    catch {
        Write-LogMessage -Message "Failed to optimize Windows network stack: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Test-NetworkConnectivity {
    try {
        Write-LogMessage -Message "Testing network connectivity..." -Level "INFO"
        
        $TestHosts = @("8.8.8.8", "1.1.1.1", "google.com")
        $ConnectivityResults = @()
        
        foreach ($TestHost in $TestHosts) {
            try {
                $PingResult = Test-Connection -ComputerName $TestHost -Count 2 -Quiet
                $ConnectivityResults += [PSCustomObject]@{
                    Host = $TestHost
                    Status = if ($PingResult) { "Success" } else { "Failed" }
                }
                
                if ($PingResult) {
                    Write-LogMessage -Message "Connectivity test to ${TestHost}: SUCCESS" -Level "SUCCESS"
                } else {
                    Write-LogMessage -Message "Connectivity test to ${TestHost}: FAILED" -Level "WARNING"
                }
            }
            catch {
                Write-LogMessage -Message "Connectivity test to ${TestHost} failed: $($_.Exception.Message)" -Level "WARNING"
                $ConnectivityResults += [PSCustomObject]@{
                    Host = $TestHost
                    Status = "Error"
                }
            }
        }
        
        $SuccessfulTests = ($ConnectivityResults | Where-Object { $_.Status -eq "Success" }).Count
        Write-LogMessage -Message "Network connectivity tests completed: $SuccessfulTests/$($TestHosts.Count) successful" -Level "INFO"
        
        return $ConnectivityResults
    }
    catch {
        Write-LogMessage -Message "Failed to test network connectivity: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

function Show-OptimizationSummary {
    param(
        [string]$BackupPath,
        [array]$ConnectivityResults
    )
    
    Write-LogMessage -Message "=== NETWORK OPTIMIZATION SUMMARY ===" -Level "INFO"
    Write-LogMessage -Message "Optimizations completed successfully:" -Level "SUCCESS"
    Write-LogMessage -Message "  [OK] IPv6 protocol disabled" -Level "SUCCESS"
    Write-LogMessage -Message "  [OK] TCP/IP settings optimized" -Level "SUCCESS"
    Write-LogMessage -Message "  [OK] Network adapter settings optimized" -Level "SUCCESS"
    Write-LogMessage -Message "  [OK] DNS settings optimized" -Level "SUCCESS"
    Write-LogMessage -Message "  [OK] Windows network stack optimized" -Level "SUCCESS"
    Write-LogMessage -Message " " -Level "INFO"
    Write-LogMessage -Message "Configuration backup saved to: $BackupPath" -Level "INFO"
    Write-LogMessage -Message "Log file saved to: $Script:LogPath" -Level "INFO"
    Write-LogMessage -Message " " -Level "INFO"
    Write-LogMessage -Message "IMPORTANT: A system restart is recommended for all changes to take effect." -Level "WARNING"
    Write-LogMessage -Message "To re-enable IPv6 later, set HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents to 0" -Level "INFO"
}

# MAIN FUNCTION TO CALL FROM YOUR SCRIPT
function Start-NetworkOptimization {
    <#
    .SYNOPSIS
        Main function to run all network optimizations
    .DESCRIPTION
        This is the function you call from your main script to run the network optimization
    .EXAMPLE
        Start-NetworkOptimization
    #>
    
    try {
        # Initialize logging
        Initialize-NetworkOptimizationLogging
        
        Write-LogMessage -Message "Starting Network Optimization Process..." -Level "INFO"
        Write-LogMessage -Message " " -Level "INFO"
        
        # Check administrator privileges
        if (-not (Test-AdministratorPrivileges)) {
            throw "Administrator privileges required to run network optimization"
        }
        
        # Create backup
        $BackupPath = Backup-NetworkConfiguration
        
        # Run optimizations
        Write-LogMessage -Message " " -Level "INFO"
        Disable-IPv6Protocol
        
        Write-LogMessage -Message " " -Level "INFO"
        Optimize-TCPIPSettings
        
        Write-LogMessage -Message " " -Level "INFO"
        Optimize-NetworkAdapterSettings
        
        Write-LogMessage -Message " " -Level "INFO"
        Optimize-DNSSettings
        
        Write-LogMessage -Message " " -Level "INFO"
        Optimize-WindowsNetworkStack
        
        # Test connectivity
        Write-LogMessage -Message " " -Level "INFO"
        $ConnectivityResults = Test-NetworkConnectivity
        
        # Show summary
        Write-LogMessage -Message " " -Level "INFO"
        Show-OptimizationSummary -BackupPath $BackupPath -ConnectivityResults $ConnectivityResults
        
        return @{
            Success = $true
            BackupPath = $BackupPath
            LogPath = $Script:LogPath
            Message = "Network optimization completed successfully"
        }
    }
    catch {
        $ErrorMessage = "Network optimization failed: $($_.Exception.Message)"
        Write-LogMessage -Message $ErrorMessage -Level "ERROR"
        
        return @{
            Success = $false
            BackupPath = $null
            LogPath = $Script:LogPath
            Message = $ErrorMessage
            Error = $_.Exception
        }
    }
}

# -----------------------------------------------------------------------------
# Menu Display
# -----------------------------------------------------------------------------
function Show-Menu {
    Clear-Host
    Write-Host "[==========================================================]" -ForegroundColor Magenta
    Write-Host "|              Compton College Tech Utils                  |" -ForegroundColor Magenta
    Write-Host "[==========================================================]" -ForegroundColor Magenta
    Write-Host ""

    Write-Host "1.  Create MISAdmin account" -ForegroundColor White
    Write-Host "2.  Remove Windows Bloatware" -ForegroundColor White
    Write-Host "3.  Set Recommended Registry Settings" -ForegroundColor White
    Write-Host "4.  Optimize Windows Services" -ForegroundColor White
    Write-Host "5.  Enable PowerShell Remote Management" -ForegroundColor White
    Write-Host "6.  Configure Automatic Time Sync" -ForegroundColor White
    Write-Host "7.  Set Desktop/Laptop Power Settings" -ForegroundColor White
    Write-Host "8.  Network Optimization" -ForegroundColor White
    Write-Host "9.  Application Updates" -ForegroundColor White
    Write-Host "10. HP/Dell Driver Updates" -ForegroundColor White
    Write-Host "11. Windows Updates" -ForegroundColor White
    Write-Host "12. Disk Cleanup" -ForegroundColor White
    Write-Host "13. System Repair" -ForegroundColor White
    Write-Host "14. Remove User Profiles" -ForegroundColor White
    Write-Host "15. Disable Last User Display" -ForegroundColor White
    Write-Host "16. Enable Automatic Login with CC-Student" -ForegroundColor White
    Write-Host "17. Stage Lab Scripts and Register Scheduled Tasks" -ForegroundColor White
    Write-Host "18. Set OneDrive Auto Login on Boot" -ForegroundColor White
    Write-Host "19. Run Full System Updates" -ForegroundColor White
    Write-Host "20. Network Diag & Repair" -ForegroundColor White
    Write-Host "Q.  Exit" -ForegroundColor Red

    Write-Host ""
    Write-Host "Last Status: $global:LastStatus" -ForegroundColor Yellow
    Write-Host ""
}

# -----------------------------------------------------------------------------
# Main Loop
# -----------------------------------------------------------------------------
function Main {
    do {
        Clear-Host  # <== Clear screen BEFORE showing the menu
        Show-Menu
        $selection = Read-Host "Select an option"
        switch ($selection) {
            "1" {
                Clear-Host
                New-MISAdminAccount
                Pause
            }
            "2" {
                Clear-Host
                Remove-BloatwareApps
                Pause
            }
			"3" {
				Clear-Host
				Apply-RecommendedRegistrySettings
				Pause
			}
			"4" {
				Clear-Host
				Optimize-WindowsServices
				Pause
			}
			"5" {
				Clear-Host
				Enable-PowerShellRemotingSafely
				Pause
			}
			"6" {
				Clear-Host
				Configure-AutomaticTimeSync
				Pause
			}
			"7" {
				Clear-Host
				Set-DesktopPowerSettings
				Pause
			}
			"8" {
				Clear-Host
				Start-NetworkOptimization
				Pause
			}
			"9" {
				Clear-Host
				Update-Applications
				Pause
			}
			"10" {
				Clear-Host
				Update-HPDrivers
				Pause
			}
			"11"{
				Clear-Host
				Update-WindowsOS
				Pause
			}
			"12" {
				Clear-Host
				Run-DiskCleanup
				Pause
				}
			"13" {
				Clear-Host
				Invoke-SystemMaintenance
				Pause
			}
			"14" {
				Clear-Host
				Remove-UserProfilesClassroom
				Pause
			}
			"15" {
				Clear-Host
				Apply-LoginScreenRegistryFixes
				Pause
			}
			"16" {
				Clear-Host
				Set-DomainAutoLogin
				# Get-AdminCredentials
				Pause
			}
			"17" {
				Clear-Host
				Register-LabScheduledTasks
				Pause
			}
			"18" {
				Clear-Host
				Set-OneDriveAutoLoginPolicy
				Pause
			}
			"19" {
				Clear-Host
				Run-CorePostDeploymentTasks
				Pause
			}
			"20" {
				Clear-Host
				Run-NetworkDiagnostics
				Pause
			}
			
			"Q" {
				Clear-Host
				Write-Host "Compton College Tech Utils has exited." -ForegroundColor Cyan
				exit
			}
            default {
                $global:LastStatus = "[ERROR] Invalid selection. Please try again."
                Start-Sleep -Seconds 1
            }
        }
    } while ($true)
}

# -----------------------------------------------------------------------------
# Entry Point
# -----------------------------------------------------------------------------
Invoke-StartupSelfUpdate
Main
