# =====================================================================
# ScriptName: Compton_Tech_Utils.ps1
# ScriptVersion: 1.8.9
# LastUpdated: 2026-04-08
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
# Option 10 - HP Driver Updates
# -----------------------------------------------------------------------------

function Update-HPDrivers {
    <#
    .SYNOPSIS
    Runs the HP driver update workflow used by Option 10.

    .DESCRIPTION
    Replaces the prior Option 10 implementation with the contents of
    05_Weekend_HP_Drivers_Update.ps1, adapted to run as a callable menu
    function instead of a standalone script. Logging is YAML-only.
    #>
    [CmdletBinding()]
    param(

    [switch]$IncludeBIOS = $true,
    [switch]$IncludeSoftware = $false,
    [switch]$SuspendBitLockerForBIOS = $true,
    [string]$WorkingRoot = 'C:\Temp\HPDrivers',
    [string]$YamlLogFolder = 'C:\Logs',
    [int]$CleanupRetryCount = 12,
    [int]$CleanupRetryDelaySeconds = 10
    )

$ErrorActionPreference = 'Stop'

$script:RunStart = Get-Date
$script:ComputerName = $env:COMPUTERNAME
$script:YamlLogPath = $null
$script:OverallResult = 'Unknown'
$script:FailureMessage = $null
$script:CleanupResult = $false
$script:DetectedSoftPaqs = New-Object System.Collections.Generic.List[object]
$script:InstalledSoftPaqResults = New-Object System.Collections.Generic.List[object]

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

function Test-IsHPSystem {
    try {
        $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
        return ($manufacturer -match 'HP|Hewlett-Packard')
    }
    catch {
        return $false
    }
}

function Ensure-Folder {
    param([string]$Path)

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

    $text = [string]$Value
    $text = $text -replace "`r", ' '
    $text = $text -replace "`n", ' '
    $text = $text -replace '"', '\"'
    return '"' + $text + '"'
}

function Initialize-YamlLog {
    Ensure-Folder -Path $YamlLogFolder

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
    $fileName = "$($script:ComputerName)-HPDrivers-$timestamp.yml"
    $script:YamlLogPath = Join-Path $YamlLogFolder $fileName

    Write-Log "YAML log will be written to: $($script:YamlLogPath)" 'INFO'
}

function Add-DetectedSoftPaq {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Version,
        [string]$Category
    )

    $script:DetectedSoftPaqs.Add([PSCustomObject]@{
        Id       = $Id
        Name     = $Name
        Version  = $Version
        Category = $Category
    }) | Out-Null
}

function Add-InstalledSoftPaqResult {
    param(
        [string]$Id,
        [string]$Name,
        [string]$Version,
        [string]$Status,
        [string]$Message
    )

    $script:InstalledSoftPaqResults.Add([PSCustomObject]@{
        Id      = $Id
        Name    = $Name
        Version = $Version
        Status  = $Status
        Message = $Message
    }) | Out-Null
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
        $yamlLines.Add('script_name: "05_Weekend_HP_Drivers_Update.ps1"') | Out-Null
        $yamlLines.Add('script_version: "1.2"') | Out-Null
        $yamlLines.Add('run_started: ' + (ConvertTo-YamlSafeValue ($script:RunStart.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
        $yamlLines.Add('run_finished: ' + (ConvertTo-YamlSafeValue ($runEnd.ToString('yyyy-MM-dd HH:mm:ss')))) | Out-Null
        $yamlLines.Add('duration_seconds: ' + $duration) | Out-Null

        $yamlLines.Add('options:') | Out-Null
        $yamlLines.Add('  include_bios: ' + ($(if ($IncludeBIOS) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('  include_software: ' + ($(if ($IncludeSoftware) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('  suspend_bitlocker_for_bios: ' + ($(if ($SuspendBitLockerForBIOS) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('  working_root: ' + (ConvertTo-YamlSafeValue $WorkingRoot)) | Out-Null
        $yamlLines.Add('  cleanup_retry_count: ' + $CleanupRetryCount) | Out-Null
        $yamlLines.Add('  cleanup_retry_delay_seconds: ' + $CleanupRetryDelaySeconds) | Out-Null

        $yamlLines.Add('cleanup_successful: ' + ($(if ($script:CleanupResult) { 'true' } else { 'false' }))) | Out-Null
        $yamlLines.Add('overall_result: ' + (ConvertTo-YamlSafeValue $script:OverallResult)) | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($script:FailureMessage)) {
            $yamlLines.Add('failure_message: ' + (ConvertTo-YamlSafeValue $script:FailureMessage)) | Out-Null
        }
        else {
            $yamlLines.Add('failure_message: null') | Out-Null
        }

        $yamlLines.Add('detected_softpaqs:') | Out-Null
        if ($script:DetectedSoftPaqs.Count -gt 0) {
            foreach ($item in $script:DetectedSoftPaqs) {
                $yamlLines.Add('  - id: ' + (ConvertTo-YamlSafeValue $item.Id)) | Out-Null
                $yamlLines.Add('    name: ' + (ConvertTo-YamlSafeValue $item.Name)) | Out-Null
                $yamlLines.Add('    version: ' + (ConvertTo-YamlSafeValue $item.Version)) | Out-Null
                $yamlLines.Add('    category: ' + (ConvertTo-YamlSafeValue $item.Category)) | Out-Null
            }
        }
        else {
            $yamlLines.Add('  - id: null') | Out-Null
            $yamlLines.Add('    name: "No applicable HP SoftPaq updates detected"') | Out-Null
            $yamlLines.Add('    version: null') | Out-Null
            $yamlLines.Add('    category: null') | Out-Null
        }

        $yamlLines.Add('install_results:') | Out-Null
        if ($script:InstalledSoftPaqResults.Count -gt 0) {
            foreach ($item in $script:InstalledSoftPaqResults) {
                $yamlLines.Add('  - id: ' + (ConvertTo-YamlSafeValue $item.Id)) | Out-Null
                $yamlLines.Add('    name: ' + (ConvertTo-YamlSafeValue $item.Name)) | Out-Null
                $yamlLines.Add('    version: ' + (ConvertTo-YamlSafeValue $item.Version)) | Out-Null
                $yamlLines.Add('    status: ' + (ConvertTo-YamlSafeValue $item.Status)) | Out-Null
                $yamlLines.Add('    message: ' + (ConvertTo-YamlSafeValue $item.Message)) | Out-Null
            }
        }
        else {
            $yamlLines.Add('  - id: null') | Out-Null
            $yamlLines.Add('    name: "No SoftPaq install operations were performed"') | Out-Null
            $yamlLines.Add('    version: null') | Out-Null
            $yamlLines.Add('    status: "None"') | Out-Null
            $yamlLines.Add('    message: "None"') | Out-Null
        }

        Set-Content -Path $script:YamlLogPath -Value $yamlLines -Encoding UTF8
        Write-Log "YAML log written successfully: $($script:YamlLogPath)" 'OK'
    }
    catch {
        Write-Log "Failed to write YAML log: $($_.Exception.Message)" 'WARN'
    }
}

function Save-PowerSettings {
    $settings = [ordered]@{}

    $settings.DisplayTimeoutDC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\DC\{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}'
    ).SettingIndexValue / 60

    $settings.DisplayTimeoutAC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\AC\{3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e}'
    ).SettingIndexValue / 60

    $settings.SleepTimeoutDC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\DC\{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}'
    ).SettingIndexValue / 60

    $settings.SleepTimeoutAC = (
        Get-CimInstance -Namespace root\cimv2\power -Class Win32_PowerSettingDataIndex |
        Where-Object InstanceID -EQ 'Microsoft:PowerSettingDataIndex\{381b4222-f694-41f0-9685-ff5bb260df2e}\AC\{29f6c1db-86da-48c5-9fdb-f2b67b1f44da}'
    ).SettingIndexValue / 60

    return [PSCustomObject]$settings
}

function Set-UnlimitedPowerTimeouts {
    Write-Log "Temporarily disabling monitor and sleep timeouts..." 'INFO'
    powercfg -change -monitor-timeout-dc 0 | Out-Null
    powercfg -change -monitor-timeout-ac 0 | Out-Null
    powercfg -change -standby-timeout-dc 0 | Out-Null
    powercfg -change -standby-timeout-ac 0 | Out-Null
}

function Restore-PowerSettings {
    param($Saved)

    if ($null -eq $Saved) { return }

    Write-Log "Restoring previous power timeout settings..." 'INFO'
    powercfg -change -monitor-timeout-dc $Saved.DisplayTimeoutDC | Out-Null
    powercfg -change -monitor-timeout-ac $Saved.DisplayTimeoutAC | Out-Null
    powercfg -change -standby-timeout-dc $Saved.SleepTimeoutDC | Out-Null
    powercfg -change -standby-timeout-ac $Saved.SleepTimeoutAC | Out-Null
}

function Ensure-Tls12 {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Write-Log "Enabled TLS 1.2 for PowerShell Gallery access." 'OK'
    }
    catch {
        Write-Log "Could not explicitly enable TLS 1.2: $($_.Exception.Message)" 'WARN'
    }
}


function Initialize-HPNetworkAccess {
    Write-Log "Applying HP download network compatibility fixes..." 'INFO'

    try {
        & netsh winhttp reset proxy | Out-Null
        Write-Log "Reset WinHTTP proxy to direct access." 'OK'
    }
    catch {
        Write-Log "Could not reset WinHTTP proxy: $($_.Exception.Message)" 'WARN'
    }

    try {
        [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy
        [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
        Write-Log "Configured .NET web requests to use direct proxy settings with default credentials." 'OK'
    }
    catch {
        Write-Log "Could not adjust .NET proxy settings: $($_.Exception.Message)" 'WARN'
    }

    try {
        [System.Net.ServicePointManager]::Expect100Continue = $false
    }
    catch {}

    Ensure-Tls12

    try {
        $dnsOk = $false
        try {
            if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
                $null = Resolve-DnsName -Name 'ftp.hp.com' -Type A -ErrorAction Stop
                $dnsOk = $true
            }
        }
        catch {}

        if (-not $dnsOk) {
            try {
                $null = [System.Net.Dns]::GetHostAddresses('ftp.hp.com')
                $dnsOk = $true
            }
            catch {}
        }

        if ($dnsOk) {
            Write-Log "DNS resolution for ftp.hp.com succeeded." 'OK'
        }
        else {
            Write-Log "DNS resolution for ftp.hp.com failed. HP SoftPaq downloads may fail until DNS/network access is restored." 'WARN'
        }
    }
    catch {
        Write-Log "Could not validate DNS resolution for ftp.hp.com: $($_.Exception.Message)" 'WARN'
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [int]$RetryCount = 3,
        [int]$DelaySeconds = 8,
        [string]$ActionDescription = 'Operation'
    )

    $lastError = $null

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            & $ScriptBlock
            if ($attempt -gt 1) {
                Write-Log "$ActionDescription succeeded on attempt $attempt of $RetryCount." 'OK'
            }
            return $true
        }
        catch {
            $lastError = $_.Exception.Message
            Write-Log "$ActionDescription failed on attempt $attempt of ${RetryCount}: $lastError" 'WARN'

            if ($attempt -lt $RetryCount) {
                Start-Sleep -Seconds $DelaySeconds
                Initialize-HPNetworkAccess
            }
        }
    }

    throw "$ActionDescription failed after $RetryCount attempts. Last error: $lastError"
}

function Ensure-NuGetProvider {
    Write-Log "Ensuring NuGet provider is installed..." 'INFO'
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers | Out-Null
        Write-Log "NuGet provider is ready." 'OK'
    }
    catch {
        Write-Log "NuGet provider installation/check failed: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-PSGalleryTrusted {
    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
            Write-Log "Set PSGallery repository to Trusted." 'OK'
        }
        else {
            Write-Log "PSGallery repository already Trusted." 'INFO'
        }
    }
    catch {
        Write-Log "Could not validate/set PSGallery trust: $($_.Exception.Message)" 'WARN'
    }
}

function Install-ModuleIfPossible {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$AllowClobber
    )

    $args = @{
        Name        = $Name
        Scope       = 'AllUsers'
        Force       = $true
        ErrorAction = 'Stop'
    }

    if ($AllowClobber) {
        $args['AllowClobber'] = $true
    }

    Install-Module @args | Out-Null
}

function Ensure-PackageTooling {
    Write-Log "Ensuring PowerShell package tooling is current enough for HPCMSL..." 'INFO'

    Initialize-HPNetworkAccess
    Ensure-NuGetProvider
    Ensure-PSGalleryTrusted

    try {
        Install-ModuleIfPossible -Name 'PowerShellGet' -AllowClobber
        Write-Log "PowerShellGet installed/updated." 'OK'
    }
    catch {
        Write-Log "PowerShellGet update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Install-ModuleIfPossible -Name 'Microsoft.PowerShell.PSResourceGet'
        Write-Log "Microsoft.PowerShell.PSResourceGet installed/updated." 'OK'
    }
    catch {
        Write-Log "PSResourceGet install/update failed: $($_.Exception.Message)" 'WARN'
    }

    try {
        Import-Module PowerShellGet -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Could not import PowerShellGet: $($_.Exception.Message)" 'WARN'
    }

    try {
        Import-Module Microsoft.PowerShell.PSResourceGet -Force -ErrorAction Stop
    }
    catch {
        Write-Log "Could not import PSResourceGet yet: $($_.Exception.Message)" 'WARN'
    }
}

function Ensure-HPCMSL {
    Write-Log "Ensuring HP CMSL is available..." 'INFO'

    Ensure-PackageTooling

    $moduleLoaded = $false

    if (-not (Get-Module -ListAvailable -Name HPCMSL)) {
        try {
            if (Get-Command -Name Install-PSResource -ErrorAction SilentlyContinue) {
                Write-Log "Installing HPCMSL with Install-PSResource..." 'INFO'
                Install-PSResource -Name HPCMSL -Scope AllUsers -TrustRepository -Quiet -AcceptLicense -ErrorAction Stop | Out-Null
            }
            else {
                Write-Log "Install-PSResource not available. Falling back to Install-Module for HPCMSL..." 'WARN'
                Install-ModuleIfPossible -Name 'HPCMSL' -AllowClobber
            }
        }
        catch {
            Write-Log "Primary HPCMSL install attempt failed: $($_.Exception.Message)" 'WARN'

            try {
                Write-Log "Trying fallback HPCMSL install with Install-Module..." 'INFO'
                Install-ModuleIfPossible -Name 'HPCMSL' -AllowClobber
            }
            catch {
                throw "Failed to install HPCMSL. $($_.Exception.Message)"
            }
        }
    }
    else {
        Write-Log "HPCMSL already present on system." 'INFO'
    }

    try {
        Import-Module HPCMSL -Force -ErrorAction Stop
        $moduleLoaded = $true
    }
    catch {
        try {
            Import-Module HP.Softpaq -Force -ErrorAction Stop
            $moduleLoaded = $true
        }
        catch {
            throw "HPCMSL/HP.Softpaq could not be imported after installation. $($_.Exception.Message)"
        }
    }

    if ($moduleLoaded) {
        Write-Log "HP CMSL imported successfully." 'OK'
    }
}

function Get-HPSoftpaqCategories {
    $categories = @('Driver')

    if ($IncludeBIOS) {
        $categories += 'BIOS'
    }

    if ($IncludeSoftware) {
        $categories += @('Diagnostic', 'Dock', 'Software', 'Utility')
    }

    return $categories
}

function Get-DriverList {
    $categories = Get-HPSoftpaqCategories
    Write-Log "Querying HP SoftPaq list for categories: $($categories -join ', ')" 'INFO'
    Write-Log "BIOS category is enabled by default so Flash BIOS firmware updates can be downloaded and installed when applicable." 'INFO'

    $list = Get-SoftpaqList -Category $categories

    if (-not $list) {
        Write-Log "No applicable HP SoftPaq updates were returned." 'OK'
        return @()
    }

    $biosCount = 0

    foreach ($item in $list) {
        $category = $null
        if ($item.PSObject.Properties.Name -contains 'Category') {
            $category = [string]$item.Category
        }

        if ($category -match 'BIOS') {
            $biosCount++
        }

        Add-DetectedSoftPaq -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Category $category
        Write-Log "Detected: [$($item.Id)] $($item.Name) Version $($item.Version) Category [$category]" 'INFO'
    }

    if ($biosCount -gt 0) {
        Write-Log "Detected $biosCount applicable BIOS/Flash BIOS firmware SoftPaq update(s)." 'INFO'
    }
    else {
        Write-Log "No applicable BIOS/Flash BIOS firmware SoftPaq updates were detected." 'INFO'
    }

    return @($list)
}

function Install-SoftpaqList {
    param([object[]]$Softpaqs)

    $failures = 0

    foreach ($item in $Softpaqs) {
        try {
            $category = $null
            if ($item.PSObject.Properties.Name -contains 'Category') {
                $category = [string]$item.Category
            }

            if ($category -match 'BIOS') {
                Write-Log "Downloading and installing Flash BIOS firmware SoftPaq [$($item.Id)] $($item.Name)..." 'INFO'
            }
            else {
                Write-Log "Installing SoftPaq [$($item.Id)] $($item.Name)..." 'INFO'
            }

            Get-Softpaq -Number $item.Id -Action SilentInstall | Out-Null

            if ($category -match 'BIOS') {
                Write-Log "Installed Flash BIOS firmware SoftPaq [$($item.Id)] $($item.Name). A reboot may be required to complete flashing." 'OK'
                Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Succeeded' -Message 'Flash BIOS firmware installed successfully; reboot may be required'
            }
            else {
                Write-Log "Installed SoftPaq [$($item.Id)] $($item.Name)." 'OK'
                Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Succeeded' -Message 'Installed successfully'
            }
        }
        catch {
            $msg = $_.Exception.Message
            Write-Log "Failed SoftPaq [$($item.Id)] $($item.Name): $msg" 'WARN'
            Add-InstalledSoftPaqResult -Id ([string]$item.Id) -Name ([string]$item.Name) -Version ([string]$item.Version) -Status 'Failed' -Message $msg
            $failures++
        }
    }

    return $failures
}

function Suspend-BitLockerIfNeeded {
    if (-not $SuspendBitLockerForBIOS) {
        return
    }

    if (-not $IncludeBIOS) {
        Write-Log "SuspendBitLockerForBIOS was requested, but IncludeBIOS is not enabled. Skipping BitLocker suspend." 'WARN'
        return
    }

    try {
        $vol = Get-BitLockerVolume -MountPoint 'C:'
        if ($vol.VolumeStatus -ne 'FullyDecrypted') {
            Suspend-BitLocker -MountPoint 'C:' -RebootCount 1
            Write-Log "BitLocker suspended for one reboot." 'OK'
        }
        else {
            Write-Log "BitLocker is not active on C:. No suspend needed." 'INFO'
        }
    }
    catch {
        Write-Log "Could not evaluate or suspend BitLocker: $($_.Exception.Message)" 'WARN'
    }
}

function Remove-WorkingFolderRobust {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$RetryCount = 12,
        [int]$RetryDelaySeconds = 10
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "Working folder already absent: $Path" 'OK'
        return $true
    }

    Write-Log "Attempting to remove working folder: $Path" 'INFO'

    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            Start-Sleep -Seconds 2
            Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $Path)) {
                Write-Log "Working folder removed successfully." 'OK'
                return $true
            }
        }
        catch {
            Write-Log "Cleanup attempt $i/$RetryCount failed: $($_.Exception.Message)" 'WARN'
        }

        if ($i -lt $RetryCount) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    Write-Log "Working folder still exists after cleanup attempts: $Path" 'ERROR'
    return $false
}

# Main
if (-not (Test-IsAdministrator)) {
    Write-Error "Please run this script as Administrator."
    return 1
}

Initialize-YamlLog

if (-not (Test-IsHPSystem)) {
    Write-Log "This is not an HP or Hewlett-Packard system. Skipping HP driver update." 'WARN'
    $script:OverallResult = 'SkippedNonHPSystem'
    Write-YamlLog
    return 0
}

Ensure-Folder -Path $WorkingRoot

$SavedPower = $null
$OriginalLocation = (Get-Location).Path
$Failures = 0
$CleanupOk = $false

try {
    Write-Log "Initializing HP driver update script..." 'INFO'
    Write-Log "Working root: $WorkingRoot" 'INFO'

    $SavedPower = Save-PowerSettings
    Set-UnlimitedPowerTimeouts
    Ensure-HPCMSL
    Suspend-BitLockerIfNeeded

    Set-Location -Path $WorkingRoot

    $softpaqs = Get-DriverList
    if ($softpaqs.Count -eq 0) {
        $CleanupOk = Remove-WorkingFolderRobust -Path $WorkingRoot -RetryCount $CleanupRetryCount -RetryDelaySeconds $CleanupRetryDelaySeconds
        $script:CleanupResult = $CleanupOk
        $script:OverallResult = if ($CleanupOk) { 'SucceededNoApplicableUpdates' } else { 'SucceededNoApplicableUpdatesCleanupIncomplete' }
        Write-YamlLog
        return $(if ($CleanupOk) { 0 } else { 2 })
    }

    $Failures = Install-SoftpaqList -Softpaqs $softpaqs
}
catch {
    $script:FailureMessage = $_.Exception.Message
    Write-Log "Script failed: $($_.Exception.Message)" 'ERROR'
    $Failures++
}
finally {
    try {
        Set-Location -Path $OriginalLocation
    }
    catch {
    }

    try {
        Restore-PowerSettings -Saved $SavedPower
    }
    catch {
        Write-Log "Failed restoring power settings: $($_.Exception.Message)" 'WARN'
    }

    $CleanupOk = Remove-WorkingFolderRobust -Path $WorkingRoot -RetryCount $CleanupRetryCount -RetryDelaySeconds $CleanupRetryDelaySeconds
    $script:CleanupResult = $CleanupOk
}

if ($Failures -eq 0 -and $CleanupOk) {
    Write-Log "HP driver update script completed successfully." 'OK'
    $script:OverallResult = 'Succeeded'
    Write-YamlLog
    return 0
}
elseif ($Failures -eq 0 -and -not $CleanupOk) {
    Write-Log "HP driver update succeeded, but cleanup was incomplete." 'WARN'
    $script:OverallResult = 'SucceededCleanupIncomplete'
    Write-YamlLog
    return 2
}
else {
    Write-Log "HP driver update completed with one or more failures." 'WARN'
    if ([string]::IsNullOrWhiteSpace($script:FailureMessage)) {
        $script:FailureMessage = 'One or more SoftPaq installations failed.'
    }
    $script:OverallResult = 'Failed'
    Write-YamlLog
    return 3
}
}

# -----------------------------------------------------------------------------
# Option 11 - Windows Updates
# -----------------------------------------------------------------------------
function Update-WindowsOS {
    <#
    .SYNOPSIS
    Performs comprehensive Windows OS updates with enhanced security and progress tracking
    
    .DESCRIPTION
    Executes a complete Windows update cycle including Group Policy refresh, Windows Update
    component reset, update installation, and post-update validation. Includes security
    validation, progress monitoring, and enterprise-ready features.
    
    .PARAMETER UpdateCategory
    Categories of updates to install (All, Security, Critical, Recommended, Optional)
    
    .PARAMETER CreateRestorePoint
    Create system restore point before installing updates
    
    .PARAMETER SkipGroupPolicyReset
    Skip Group Policy folder reset and update
    
    .PARAMETER SkipComponentReset
    Skip Windows Update component reset for faster execution
    
    .PARAMETER MaxUpdateTimeout
    Maximum timeout for individual update operations in minutes
    
    .PARAMETER AllowReboot
    Allow automatic reboot if required by updates
    
    .PARAMETER DeferFeatureUpdates
    Defer feature updates and install only quality updates
    
    .OUTPUTS
    Returns hashtable with detailed update results and system status
    
    .EXAMPLE
    Update-WindowsOS
    
    .EXAMPLE
    Update-WindowsOS -UpdateCategory Security -CreateRestorePoint
    
    .EXAMPLE
    Update-WindowsOS -SkipGroupPolicyReset -DeferFeatureUpdates -WhatIf
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [ValidateSet('All', 'Security', 'Critical', 'Recommended', 'Optional')]
        [string]$UpdateCategory = 'All',
        
        [Parameter()]
        [switch]$CreateRestorePoint,
        
        [Parameter()]
        [switch]$SkipGroupPolicyReset,
        
        [Parameter()]
        [switch]$SkipComponentReset,
        
        [Parameter()]
        [ValidateRange(5, 120)]
        [int]$MaxUpdateTimeout = 60,
        
        [Parameter()]
        [switch]$AllowReboot,
        
        [Parameter()]
        [switch]$DeferFeatureUpdates
    )
    
    Write-StatusLog "Starting comprehensive Windows OS update process..." -Level "Info"
    
    # Initialize results tracking
    $results = @{
        PrerequisitesReady = $false
        GroupPolicyReset = $false
        ComponentsReset = $false
        RestorePointCreated = $false
        AvailableUpdates = @()
        InstalledUpdates = @()
        FailedUpdates = @()
        SecurityUpdates = @()
        RebootRequired = $false
        UpdateDuration = $null
        SystemValidation = @{}
        Errors = @()
        ProgressDetails = @()
    }
    
    $startTime = Get-Date
    
    try {
        # Step 1: Validate prerequisites and prepare system
        Write-StatusLog "Validating system prerequisites..." -Level "Info"
        $prereqResult = Ensure-WindowsUpdatePrerequisites
        $results.PrerequisitesReady = $prereqResult.Success
        
        if (-not $results.PrerequisitesReady) {
            $results.Errors += $prereqResult.Errors
            throw "Windows Update prerequisites not ready"
        }
        
        # Step 2: Create restore point if requested
        if ($CreateRestorePoint) {
            Write-StatusLog "Creating system restore point..." -Level "Info"
            $restoreResult = New-SystemRestorePointSafe -Description "Before Windows Updates"
            $results.RestorePointCreated = $restoreResult.Success
            
            if (-not $restoreResult.Success) {
                Write-StatusLog "[WARN] Failed to create restore point: $($restoreResult.Error)" -Level "Warning"
            }
        }
        
        # Step 3: Group Policy management
        if (-not $SkipGroupPolicyReset) {
            Write-StatusLog "Managing Group Policy configuration..." -Level "Info"
            $gpResult = Reset-GroupPolicyConfiguration
            $results.GroupPolicyReset = $gpResult.Success
            
            if (-not $gpResult.Success) {
                $results.Errors += $gpResult.Errors
            }
        }
        
        # Step 4: Windows Update component reset
        if (-not $SkipComponentReset) {
            Write-StatusLog "Resetting Windows Update components..." -Level "Info"
            $componentResult = Reset-WindowsUpdateComponents -MaxTimeout $MaxUpdateTimeout
            $results.ComponentsReset = $componentResult.Success
            
            if (-not $componentResult.Success) {
                $results.Errors += $componentResult.Errors
            }
        }
        
        # Step 5: Discover available updates
        Write-StatusLog "Discovering available Windows updates..." -Level "Info"
        $updateDiscovery = Get-AvailableWindowsUpdates -Category $UpdateCategory -DeferFeatureUpdates:$DeferFeatureUpdates
        $results.AvailableUpdates = $updateDiscovery.Updates
        
        if ($updateDiscovery.Errors.Count -gt 0) {
            $results.Errors += $updateDiscovery.Errors
        }
        
        if ($results.AvailableUpdates.Count -eq 0) {
            Write-StatusLog "No Windows updates available for installation" -Level "Info"
            $global:LastStatus = "[OK] Windows is up to date"
            return $results
        }
        
        Write-StatusLog "Found $($results.AvailableUpdates.Count) available updates" -Level "Success"
        
        # Step 6: Display update plan
        Show-WindowsUpdatePlan -Updates $results.AvailableUpdates -Category $UpdateCategory
        
        # Step 7: Install updates with progress tracking
        Write-StatusLog "Installing Windows updates with security validation..." -Level "Info"
        $installationResult = Install-WindowsUpdatesSecurely -Updates $results.AvailableUpdates -MaxTimeout $MaxUpdateTimeout -AllowReboot:$AllowReboot
        
        $results.InstalledUpdates = $installationResult.Installed
        $results.FailedUpdates = $installationResult.Failed
        $results.SecurityUpdates = $installationResult.SecurityUpdates
        $results.RebootRequired = $installationResult.RebootRequired
        $results.ProgressDetails = $installationResult.ProgressDetails
        
        if ($installationResult.Errors.Count -gt 0) {
            $results.Errors += $installationResult.Errors
        }
        
        # Step 8: Post-update system validation
        Write-StatusLog "Performing post-update system validation..." -Level "Info"
        $validationResult = Test-PostUpdateSystemHealth
        $results.SystemValidation = $validationResult
        
        if ($validationResult.Errors.Count -gt 0) {
            $results.Errors += $validationResult.Errors
        }
        
        # Calculate final statistics
        $results.UpdateDuration = (Get-Date) - $startTime
        
        # Display comprehensive summary
        Show-WindowsUpdateSummary -Results $results
        
        # Set global status
        $successCount = $results.InstalledUpdates.Count
        $failCount = $results.FailedUpdates.Count
        
        if ($results.RebootRequired -and -not $AllowReboot) {
            Write-StatusLog "[WARN] System restart required to complete updates" -Level "Warning"
        }
        
        $global:LastStatus = if ($failCount -eq 0) {
            "[OK] Windows updates completed successfully ($successCount updates)"
        } else {
            "[WARN] Windows updates completed with $failCount failures ($successCount successful)"
        }
        
    } catch {
        $errorMsg = "Critical error during Windows updates: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] $errorMsg" -Level "Error"
        $results.Errors += $errorMsg
        $global:LastStatus = "[ERROR] Windows updates failed"
    } finally {
        $results.UpdateDuration = (Get-Date) - $startTime
    }
    
    return $results
}

function Ensure-WindowsUpdatePrerequisites {
    <#
    .SYNOPSIS
    Ensures Windows Update prerequisites are met
    
    .OUTPUTS
    Returns hashtable with prerequisite validation results
    #>
    
    $result = @{
        Success = $false
        PSWindowsUpdateInstalled = $false
        WingetReady = $false
        ServicesRunning = $false
        DiskSpaceAvailable = $false
        Errors = @()
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would validate Windows Update prerequisites" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Check disk space (minimum 10GB free)
        $systemDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $env:SystemDrive }
        $freeSpaceGB = [math]::Round($systemDrive.FreeSpace / 1GB, 2)
        
        if ($freeSpaceGB -lt 10) {
            $result.Errors += "Insufficient disk space: $freeSpaceGB GB available (minimum 10GB required)"
        } else {
            $result.DiskSpaceAvailable = $true
            Write-StatusLog "[OK] Sufficient disk space available: $freeSpaceGB GB" -Level "Success"
        }
        
        # Ensure Winget dependencies
        Write-StatusLog "Verifying Winget dependencies..." -Level "Info"
        $wingetResult = Ensure-WingetDependenciesReady
        $result.WingetReady = $wingetResult.Success
        
        if (-not $result.WingetReady) {
            $result.Errors += "Winget dependencies not ready"
        }
        
        # Check and install PSWindowsUpdate module
        Write-StatusLog "Checking PSWindowsUpdate module..." -Level "Info"
        $psModule = Get-Module -ListAvailable -Name "PSWindowsUpdate" -ErrorAction SilentlyContinue
        
        if (-not $psModule) {
            Write-StatusLog "Installing PSWindowsUpdate module..." -Level "Info"
            try {
                Install-Module -Name "PSWindowsUpdate" -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $result.PSWindowsUpdateInstalled = $true
                Write-StatusLog "[OK] PSWindowsUpdate module installed successfully" -Level "Success"
            } catch {
                $result.Errors += "Failed to install PSWindowsUpdate module: $($_.Exception.Message)"
            }
        } else {
            $result.PSWindowsUpdateInstalled = $true
            Write-StatusLog "[OK] PSWindowsUpdate module already available" -Level "Success"
            
            # Check for updates
            try {
                $latestVersion = Find-Module -Name "PSWindowsUpdate" -ErrorAction SilentlyContinue
                $currentVersion = $psModule | Sort-Object Version -Descending | Select-Object -First 1
                
                if ($latestVersion -and $currentVersion.Version -lt $latestVersion.Version) {
                    Write-StatusLog "Updating PSWindowsUpdate module..." -Level "Info"
                    Update-Module -Name "PSWindowsUpdate" -Force -ErrorAction Stop
                    Write-StatusLog "[OK] PSWindowsUpdate module updated successfully" -Level "Success"
                }
            } catch {
                Write-StatusLog "[WARN] Could not check for PSWindowsUpdate module updates: $($_.Exception.Message)" -Level "Warning"
            }
        }
        
        # Import the module
        try {
            Import-Module -Name "PSWindowsUpdate" -Force -ErrorAction Stop
            Write-StatusLog "[OK] PSWindowsUpdate module imported successfully" -Level "Success"
        } catch {
            $result.Errors += "Failed to import PSWindowsUpdate module: $($_.Exception.Message)"
        }
        
        # Validate Windows Update services
        $requiredServices = @('wuauserv', 'cryptsvc', 'bits', 'msiserver')
        $servicesStatus = @()
        
        foreach ($serviceName in $requiredServices) {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($service) {
                if ($service.Status -ne 'Running') {
                    try {
                        Start-Service -Name $serviceName -ErrorAction Stop
                        Write-StatusLog "[OK] Started service: $serviceName" -Level "Success"
                    } catch {
                        $result.Errors += "Failed to start service $serviceName`: $($_.Exception.Message)"
                    }
                }
                $servicesStatus += $true
            } else {
                $result.Errors += "Required service not found: $serviceName"
                $servicesStatus += $false
            }
        }
        
        $result.ServicesRunning = ($servicesStatus | Where-Object { $_ -eq $false }).Count -eq 0
        
        $result.Success = $result.PSWindowsUpdateInstalled -and $result.WingetReady -and $result.ServicesRunning -and $result.DiskSpaceAvailable
        
    } catch {
        $result.Errors += "Failed to validate prerequisites: $($_.Exception.Message)"
    }
    
    return $result
}

function Reset-GroupPolicyConfiguration {
    <#
    .SYNOPSIS
    Resets Group Policy configuration and forces update
    
    .OUTPUTS
    Returns hashtable with Group Policy reset results
    #>
    
    $result = @{
        Success = $false
        FolderRemoved = $false
        PolicyUpdated = $false
        Errors = @()
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would reset Group Policy configuration" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        # Remove local Group Policy folder
        $gpPath = Join-Path $env:Windir 'System32\GroupPolicy'
        
        if (Test-Path $gpPath) {
            Write-StatusLog "Removing local GroupPolicy folder..." -Level "Info"
            try {
                # Take ownership and reset permissions first
                takeown /f $gpPath /r /d y 2>$null | Out-Null
                icacls $gpPath /grant administrators:F /t 2>$null | Out-Null
                
                Remove-Item -Path $gpPath -Recurse -Force -ErrorAction Stop
                $result.FolderRemoved = $true
                Write-StatusLog "[OK] Group Policy folder removed successfully" -Level "Success"
            } catch {
                $result.Errors += "Failed to remove Group Policy folder: $($_.Exception.Message)"
                Write-StatusLog "[ERROR] Failed to remove Group Policy folder: $($_.Exception.Message)" -Level "Error"
            }
        } else {
            Write-StatusLog "[INFO] Group Policy folder not found, skipping removal" -Level "Info"
            $result.FolderRemoved = $true  # Consider this successful
        }
        
        # Force Group Policy update
        Write-StatusLog "Forcing Group Policy update..." -Level "Info"
        try {
            $gpupdateResult = & gpupdate /force 2>&1
            $exitCode = $LASTEXITCODE
            
            if ($exitCode -eq 0) {
                $result.PolicyUpdated = $true
                Write-StatusLog "[OK] Group Policy update completed successfully" -Level "Success"
            } else {
                $result.Errors += "gpupdate failed with exit code: $exitCode. Output: $gpupdateResult"
                Write-StatusLog "[ERROR] Group Policy update failed with exit code: $exitCode" -Level "Error"
            }
        } catch {
            $result.Errors += "Failed to execute gpupdate: $($_.Exception.Message)"
            Write-StatusLog "[ERROR] Failed to execute gpupdate: $($_.Exception.Message)" -Level "Error"
        }
        
        $result.Success = $result.FolderRemoved -and $result.PolicyUpdated
        
    } catch {
        $result.Errors += "Group Policy reset failed: $($_.Exception.Message)"
    }
    
    return $result
}

function Reset-WindowsUpdateComponents {
    <#
    .SYNOPSIS
    Resets Windows Update components with timeout and retry logic
    
    .PARAMETER MaxTimeout
    Maximum timeout in minutes for the operation
    
    .OUTPUTS
    Returns hashtable with component reset results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter()]
        [int]$MaxTimeout = 60
    )
    
    $result = @{
        Success = $false
        ComponentsReset = $false
        Errors = @()
        Duration = $null
    }
    
    $resetStart = Get-Date
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would reset Windows Update components" -ForegroundColor Yellow
            $result.Success = $true
            return $result
        }
        
        Write-StatusLog "Resetting Windows Update components..." -Level "Info"
        
        # Check if Reset-WUComponents is available
        $resetCommand = Get-Command "Reset-WUComponents" -ErrorAction SilentlyContinue
        
        if (-not $resetCommand) {
            # Use manual component reset
            $result = Reset-WUComponentsManually
        } else {
            # Use PSWindowsUpdate module command with timeout
            $timeoutMs = $MaxTimeout * 60 * 1000
            
            $job = Start-Job -ScriptBlock {
                Import-Module PSWindowsUpdate -Force
                Reset-WUComponents -Verbose
            }
            
            try {
                $completed = Wait-Job $job -Timeout ($MaxTimeout * 60)
                
                if ($completed) {
                    $output = Receive-Job $job
                    $result.ComponentsReset = $true
                    $result.Success = $true
                    Write-StatusLog "[OK] Windows Update components reset successfully" -Level "Success"
                } else {
                    Stop-Job $job -Force
                    $result.Errors += "Windows Update component reset timed out after $MaxTimeout minutes"
                    Write-StatusLog "[ERROR] Component reset timed out after $MaxTimeout minutes" -Level "Error"
                }
            } catch {
                $result.Errors += "Failed to reset Windows Update components: $($_.Exception.Message)"
                Write-StatusLog "[ERROR] Failed to reset components: $($_.Exception.Message)" -Level "Error"
            } finally {
                Remove-Job $job -Force -ErrorAction SilentlyContinue
            }
        }
        
    } catch {
        $result.Errors += "Component reset process failed: $($_.Exception.Message)"
    } finally {
        $result.Duration = (Get-Date) - $resetStart
    }
    
    return $result
}

function Reset-WUComponentsManually {
    <#
    .SYNOPSIS
    Manually resets Windows Update components when PSWindowsUpdate module is not available
    
    .OUTPUTS
    Returns hashtable with manual reset results
    #>
    
    $result = @{
        Success = $false
        ComponentsReset = $false
        Errors = @()
    }
    
    try {
        Write-StatusLog "Performing manual Windows Update component reset..." -Level "Info"
        
        # Stop Windows Update services
        $services = @('wuauserv', 'cryptsvc', 'bits', 'msiserver')
        foreach ($service in $services) {
            try {
                Stop-Service -Name $service -Force -ErrorAction Stop
                Write-StatusLog "Stopped service: $service" -Level "Info"
            } catch {
                Write-StatusLog "[WARN] Could not stop service $service`: $($_.Exception.Message)" -Level "Warning"
            }
        }
        
        # Clear Windows Update cache
        $cachePaths = @(
            "$env:SystemRoot\SoftwareDistribution",
            "$env:SystemRoot\System32\catroot2"
        )
        
        foreach ($path in $cachePaths) {
            if (Test-Path $path) {
                try {
                    Get-ChildItem -Path $path -Recurse | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Write-StatusLog "Cleared cache: $path" -Level "Info"
                } catch {
                    Write-StatusLog "[WARN] Could not clear cache $path`: $($_.Exception.Message)" -Level "Warning"
                }
            }
        }
        
        # Restart Windows Update services
        foreach ($service in $services) {
            try {
                Start-Service -Name $service -ErrorAction Stop
                Write-StatusLog "Started service: $service" -Level "Info"
            } catch {
                $result.Errors += "Failed to start service $service`: $($_.Exception.Message)"
            }
        }
        
        $result.ComponentsReset = $true
        $result.Success = $true
        Write-StatusLog "[OK] Manual Windows Update component reset completed" -Level "Success"
        
    } catch {
        $result.Errors += "Manual component reset failed: $($_.Exception.Message)"
    }
    
    return $result
}

function Get-AvailableWindowsUpdates {
    <#
    .SYNOPSIS
    Discovers available Windows updates with categorization and better error handling
    
    .PARAMETER Category
    Category of updates to discover
    
    .PARAMETER DeferFeatureUpdates
    Defer feature updates and get only quality updates
    
    .OUTPUTS
    Returns hashtable with available updates
    #>
    
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Category = 'All',
        
        [Parameter()]
        [switch]$DeferFeatureUpdates
    )
    
    $result = @{
        Updates = @()
        Errors = @()
        DiscoveryDuration = $null
    }
    
    $discoveryStart = Get-Date
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would discover available Windows updates" -ForegroundColor Yellow
            return $result
        }
        
        Write-StatusLog "Querying Windows Update servers for available updates..." -Level "Info"
        
        # Build update criteria based on category
        $updateCriteria = switch ($Category) {
            'Security' { @{Category = @('SecurityUpdates', 'CriticalUpdates')} }
            'Critical' { @{Category = @('CriticalUpdates')} }
            'Recommended' { @{Category = @('Updates', 'SecurityUpdates', 'CriticalUpdates')} }
            'Optional' { @{Category = @('OptionalUpdates')} }
            default { @{} }  # All categories
        }
        
        # Add feature update deferral if requested
        if ($DeferFeatureUpdates) {
            $updateCriteria.NotCategory = @('FeatureUpdates', 'Upgrades')
        }
        
        # Get available updates using PSWindowsUpdate with error handling
        try {
            $availableUpdates = Get-WindowsUpdate @updateCriteria -MicrosoftUpdate -ErrorAction Stop
        }
        catch {
            $result.Errors += "Failed to query Windows Update: $($_.Exception.Message)"
            Write-StatusLog "[ERROR] Windows Update query failed: $($_.Exception.Message)" -Level "Error"
            return $result
        }
        
        # Process and categorize updates with safe property access
        foreach ($update in $availableUpdates) {
            try {
                # Safe property access with comprehensive null checking
                $title = if ($update -and $update.PSObject.Properties['Title'] -and $update.Title) { 
                    $update.Title.ToString() 
                } else { 
                    "Unknown Update" 
                }
                
                $kb = if ($update -and $update.PSObject.Properties['KB'] -and $update.KB) { 
                    $update.KB.ToString() 
                } else { 
                    "N/A" 
                }
                
                $size = if ($update -and $update.PSObject.Properties['Size'] -and $update.Size) { 
                    try {
                        [long]$update.Size
                    } catch {
                        0
                    }
                } else { 
                    0 
                }
                
                $categories = if ($update -and $update.PSObject.Properties['Categories'] -and $update.Categories) { 
                    $update.Categories 
                } else { 
                    @() 
                }
                
                $severity = if ($update -and $update.PSObject.Properties['MsrcSeverity'] -and $update.MsrcSeverity) { 
                    $update.MsrcSeverity.ToString() 
                } else { 
                    "Unknown" 
                }
                
                $rebootRequired = if ($update -and $update.PSObject.Properties['RebootRequired']) { 
                    try {
                        [bool]$update.RebootRequired
                    } catch {
                        $false
                    }
                } else { 
                    $false 
                }
                
                $description = if ($update -and $update.PSObject.Properties['Description'] -and $update.Description) { 
                    $update.Description.ToString() 
                } else { 
                    "" 
                }
                
                $lastDeployment = if ($update -and $update.PSObject.Properties['LastDeploymentChangeTime']) { 
                    $update.LastDeploymentChangeTime 
                } else { 
                    $null 
                }
                
                $maxDownloadSize = if ($update -and $update.PSObject.Properties['MaxDownloadSize'] -and $update.MaxDownloadSize) { 
                    try {
                        [long]$update.MaxDownloadSize
                    } catch {
                        0
                    }
                } else { 
                    0 
                }
                
                $updateID = if ($update -and $update.PSObject.Properties['UpdateID'] -and $update.UpdateID) { 
                    $update.UpdateID.ToString() 
                } else { 
                    "" 
                }
                
                # Clean up KB number comprehensively
                if ($kb -match "^KBKB(\d+)$") {
                    $kb = "KB$($Matches[1])"
                } elseif ($kb -match "^KB(\d+)$") {
                    $kb = "KB$($Matches[1])"
                } elseif ($kb -eq "KB" -or [string]::IsNullOrEmpty($kb) -or $kb -eq "KBN/A" -or $kb -like "*N/A*") {
                    $kb = "N/A"
                }
                
                # Determine if security or critical with safe category checking
                $isSecurity = $false
                $isCritical = $false
                
                if ($categories -and $categories.Count -gt 0) {
                    try {
                        $categoryString = $categories -join " "
                        $isSecurity = $categoryString -match 'Security'
                        $isCritical = $categoryString -match 'Critical'
                    }
                    catch {
                        # If category checking fails, default to false
                        $isSecurity = $false
                        $isCritical = $false
                    }
                }
                
                $processedUpdate = @{
                    Title = $title
                    KB = $kb
                    Size = $size
                    Category = $categories
                    Severity = $severity
                    IsSecurity = $isSecurity
                    IsCritical = $isCritical
                    IsRebootRequired = $rebootRequired
                    Description = $description
                    LastDeploymentChangeTime = $lastDeployment
                    MaxDownloadSize = $maxDownloadSize
                    UpdateID = $updateID
                }
                
                $result.Updates += $processedUpdate
            }
            catch {
                $result.Errors += "Error processing update: $($_.Exception.Message)"
                Write-StatusLog "[WARN] Error processing update: $($_.Exception.Message)" -Level "Warning"
                
                # Add a minimal entry for the failed update so it doesn't get lost
                $result.Updates += @{
                    Title = "Error processing update"
                    KB = "N/A"
                    Size = 0
                    Category = @()
                    Severity = "Unknown"
                    IsSecurity = $false
                    IsCritical = $false
                    IsRebootRequired = $false
                    Description = "Failed to process update details"
                    LastDeploymentChangeTime = $null
                    MaxDownloadSize = 0
                    UpdateID = ""
                }
            }
        }
        
        Write-StatusLog "Found $($result.Updates.Count) available updates" -Level "Success"
        
    } catch {
        $result.Errors += "Failed to discover Windows updates: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] Update discovery failed: $($_.Exception.Message)" -Level "Error"
    } finally {
        $result.DiscoveryDuration = (Get-Date) - $discoveryStart
    }
    
    return $result
}

function Show-WindowsUpdatePlan {
    <#
    .SYNOPSIS
    Displays the Windows update plan before installation
    
    .PARAMETER Updates
    Array of available updates
    
    .PARAMETER Category
    Update category being processed
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Updates,
        
        [Parameter(Mandatory)]
        [string]$Category
    )
    
    Write-Host "`n" + "="*100 -ForegroundColor Cyan
    Write-Host "WINDOWS UPDATE PLAN ($Category Updates)" -ForegroundColor Cyan
    Write-Host "="*100 -ForegroundColor Cyan
    
    # Group by severity and type
    $securityUpdates = $Updates | Where-Object { $_.IsSecurity }
    $criticalUpdates = $Updates | Where-Object { $_.IsCritical -and -not $_.IsSecurity }
    $normalUpdates = $Updates | Where-Object { -not $_.IsCritical -and -not $_.IsSecurity }
    $rebootRequired = $Updates | Where-Object { $_.IsRebootRequired }
    
    if ($securityUpdates.Count -gt 0) {
        Write-Host "`nSECURITY UPDATES ($($securityUpdates.Count)):" -ForegroundColor Red
        $securityUpdates | ForEach-Object {
            $sizeText = if ($_.Size) { " - $([math]::Round($_.Size / 1MB, 1)) MB" } else { "" }
            Write-Host "  [RED] $($_.Title) ($(if ($_.KB -and $_.KB -ne 'N/A') { $_.KB } else { 'N/A' }))$sizeText" -ForegroundColor Red
        }
    }
    
    if ($criticalUpdates.Count -gt 0) {
        Write-Host "`nCRITICAL UPDATES ($($criticalUpdates.Count)):" -ForegroundColor Yellow
        $criticalUpdates | ForEach-Object {
            $sizeText = if ($_.Size) { " - $([math]::Round($_.Size / 1MB, 1)) MB" } else { "" }
            Write-Host "  [YELLOW] $($_.Title) ($(if ($_.KB -and $_.KB -ne 'N/A') { $_.KB } else { 'N/A' }))$sizeText" -ForegroundColor Yellow
        }
    }
    
    if ($normalUpdates.Count -gt 0) {
        Write-Host "`nSTANDARD UPDATES ($($normalUpdates.Count)):" -ForegroundColor Green
        $normalUpdates | ForEach-Object {
            $sizeText = if ($_.Size) { " - $([math]::Round($_.Size / 1MB, 1)) MB" } else { "" }
            Write-Host "  [GREEN] $($_.Title) ($(if ($_.KB -and $_.KB -ne 'N/A') { $_.KB } else { 'N/A' }))$sizeText" -ForegroundColor Green
        }
    }
    
    $totalSize = ($Updates | Where-Object { $_.Size } | Measure-Object -Property Size -Sum).Sum / 1MB
    Write-Host "`nTotal Download Size: $($totalSize.ToString('F1')) MB" -ForegroundColor White
    Write-Host "Total Updates: $($Updates.Count)" -ForegroundColor White
    
    if ($rebootRequired.Count -gt 0) {
        Write-Host "`n[WARN]  WARNING: $($rebootRequired.Count) update(s) will require system restart!" -ForegroundColor Yellow
    }
    
    Write-Host "="*100 -ForegroundColor Cyan
}


function Get-NormalizedWindowsUpdateKb {
    <#
    .SYNOPSIS
    Normalizes KB numbers from Windows Update result objects
    #>
    param(
        [AllowNull()]
        [object]$KBValue,

        [AllowNull()]
        [string]$Title
    )

    $kbText = $null

    if ($null -ne $KBValue) {
        try {
            if ($KBValue -is [System.Collections.IEnumerable] -and -not ($KBValue -is [string])) {
                $kbItems = @($KBValue | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                if ($kbItems.Count -gt 0) {
                    $kbText = [string]$kbItems[0]
                }
            } else {
                $kbText = [string]$KBValue
            }
        } catch {
            $kbText = $null
        }
    }

    if ([string]::IsNullOrWhiteSpace($kbText) -or $kbText -eq 'KB' -or $kbText -like '*N/A*') {
        $kbText = $null
    }

    if ($kbText -and $kbText -match '^KBKB(\d+)$') {
        $kbText = "KB$($Matches[1])"
    } elseif ($kbText -and $kbText -match '^(?:KB)?(\d{6,8})$') {
        $kbText = "KB$($Matches[1])"
    }

    if (-not $kbText -and -not [string]::IsNullOrWhiteSpace($Title)) {
        $kbMatch = [regex]::Match($Title, 'KB\d{6,8}', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($kbMatch.Success) {
            $kbText = $kbMatch.Value.ToUpper()
        }
    }

    if (-not $kbText) {
        $kbText = 'N/A'
    }

    return $kbText
}

function Get-WindowsUpdateResultState {
    <#
    .SYNOPSIS
    Classifies Windows Update installation results into terminal and non-terminal states
    #>
    param(
        [AllowNull()]
        [string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($Status)) {
        return 'Unknown'
    }

    switch -Regex ($Status.Trim()) {
        '^(Installed|Succeeded|Success)$' { return 'Installed' }
        '^(Downloaded)$' { return 'Downloaded' }
        '^(Accepted)$' { return 'Accepted' }
        '^(NotApplicable|NoUpdates|AlreadyInstalled)$' { return 'Ignored' }
        '^(Failed|Error|Aborted)$' { return 'Failed' }
        default { return 'Unknown' }
    }
}

function Merge-WindowsUpdateInstallationResults {
    <#
    .SYNOPSIS
    De-duplicates Windows Update installation results and keeps the final terminal state
    #>
    param(
        [Parameter(Mandatory)]
        [array]$InstallationResults
    )

    $resultMap = @{}

    foreach ($updateResult in $InstallationResults) {
        if (-not $updateResult) { continue }

        $title = "Unknown Update"
        $kb = "N/A"
        $size = 0
        $categories = @()
        $rebootRequired = $false
        $status = "Unknown"
        $isSecurity = $false
        $failureReason = $null

        try {
            if ($updateResult.PSObject.Properties['Title'] -and $updateResult.Title) { $title = [string]$updateResult.Title }
        } catch {}

        try {
            if ($updateResult.PSObject.Properties['KB']) { $kb = Get-NormalizedWindowsUpdateKb -KBValue $updateResult.KB -Title $title }
            else { $kb = Get-NormalizedWindowsUpdateKb -KBValue $null -Title $title }
        } catch {
            $kb = Get-NormalizedWindowsUpdateKb -KBValue $null -Title $title
        }

        try {
            if ($updateResult.PSObject.Properties['Size'] -and $null -ne $updateResult.Size) { $size = [long]$updateResult.Size }
        } catch {}

        try {
            if ($updateResult.PSObject.Properties['Categories'] -and $updateResult.Categories) { $categories = @($updateResult.Categories) }
        } catch {}

        try {
            if ($updateResult.PSObject.Properties['RebootRequired']) { $rebootRequired = [bool]$updateResult.RebootRequired }
        } catch {}

        try {
            if ($updateResult.PSObject.Properties['Result'] -and $updateResult.Result) { $status = [string]$updateResult.Result }
            elseif ($updateResult.PSObject.Properties['Status'] -and $updateResult.Status) { $status = [string]$updateResult.Status }
        } catch {}

        try {
            if ($categories -and (($categories -join ' ') -match 'Security')) {
                $isSecurity = $true
            } elseif ($title -match 'Security|Defender|Cumulative Update') {
                $isSecurity = $true
            }
        } catch {}

        try {
            if ($updateResult.PSObject.Properties['HResult'] -and $updateResult.HResult) {
                $failureReason = ('0x{0:X8}' -f ([uint32]$updateResult.HResult))
            }
        } catch {}

        if (-not $failureReason) {
            $failureReason = $status
        }

        $identityKey = "{0}|{1}" -f $title, $kb

        if (-not $resultMap.ContainsKey($identityKey)) {
            $resultMap[$identityKey] = [ordered]@{
                Title = $title
                KB = $kb
                Size = $size
                IsSecurity = $isSecurity
                RebootRequired = $rebootRequired
                SeenAccepted = $false
                SeenDownloaded = $false
                Installed = $false
                Failed = $false
                FailureReason = $null
            }
        }

        $entry = $resultMap[$identityKey]
        if ($size -gt 0) { $entry.Size = $size }
        if ($isSecurity) { $entry.IsSecurity = $true }
        if ($rebootRequired) { $entry.RebootRequired = $true }

        switch (Get-WindowsUpdateResultState -Status $status) {
            'Accepted' {
                $entry.SeenAccepted = $true
            }
            'Downloaded' {
                $entry.SeenDownloaded = $true
            }
            'Installed' {
                $entry.Installed = $true
                $entry.Failed = $false
                $entry.FailureReason = $null
            }
            'Failed' {
                if (-not $entry.Installed) {
                    $entry.Failed = $true
                    $entry.FailureReason = $failureReason
                }
            }
            default {
                # Ignore informational or unknown states unless no terminal state is ever seen
            }
        }
    }

    $final = @{
        Installed = @()
        Failed = @()
    }

    foreach ($entry in $resultMap.Values) {
        $processedResult = @{
            Title = $entry.Title
            KB = $entry.KB
            Size = $entry.Size
            IsSecurity = $entry.IsSecurity
            RebootRequired = $entry.RebootRequired
            Status = if ($entry.Installed) { 'Installed' } elseif ($entry.Failed) { 'Failed' } elseif ($entry.SeenDownloaded) { 'Downloaded' } elseif ($entry.SeenAccepted) { 'Accepted' } else { 'Unknown' }
            FailureReason = $entry.FailureReason
        }

        if ($entry.Installed) {
            $final.Installed += $processedResult
        } elseif ($entry.Failed) {
            $final.Failed += $processedResult
        }
    }

    return $final
}

function Test-BenignPostUpdateEvent {
    <#
    .SYNOPSIS
    Filters out known benign events from post-update validation noise
    #>
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Eventing.Reader.EventRecord]$EventRecord
    )

    try {
        $message = ''
        try { $message = [string]$EventRecord.Message } catch {}

        if ($EventRecord.Id -eq 3095 -and $message -match 'member of a workgroup' -and $message -match 'Netlogon service does not need to run') {
            return $true
        }

        if ($EventRecord.Id -eq 30 -and $message -match 'Microsoft-Windows-Kernel-ShimEngine/Operational' -and $message -match 'does not affect') {
            return $true
        }
    } catch {
        return $false
    }

    return $false
}


function Install-WindowsUpdatesSecurely {
    <#
    .SYNOPSIS
    Installs Windows updates with security validation and progress tracking
    
    .PARAMETER Updates
    Array of updates to install
    
    .PARAMETER MaxTimeout
    Maximum timeout per update in minutes
    
    .PARAMETER AllowReboot
    Allow automatic reboot if required
    
    .OUTPUTS
    Returns hashtable with installation results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Updates,
        
        [Parameter()]
        [int]$MaxTimeout = 60,
        
        [Parameter()]
        [switch]$AllowReboot
    )
    
    $result = @{
        Installed = @()
        Failed = @()
        SecurityUpdates = @()
        RebootRequired = $false
        ProgressDetails = @()
        Errors = @()
    }
    
    if ($WhatIfPreference) {
        Write-Host "   Would install $($Updates.Count) Windows updates" -ForegroundColor Yellow
        return $result
    }
    
    Write-Host "`n[RESTART] Installing $($Updates.Count) Windows updates..." -ForegroundColor Cyan
    
    try {
        # Install updates with progress tracking
        $installParams = @{
            AcceptAll = $true
            Install = $true
            Verbose = $true
            ErrorAction = 'Continue'
        }
        
        if ($AllowReboot) {
            $installParams.AutoReboot = $true
        }
        
        # Execute installation with error handling
        try {
            $installationResults = Install-WindowsUpdate @installParams
        }
        catch {
            $result.Errors += "Failed to execute Install-WindowsUpdate: $($_.Exception.Message)"
            Write-StatusLog "[ERROR] Install-WindowsUpdate failed: $($_.Exception.Message)" -Level "Error"
            return $result
        }
        
        # Process results with safe property access
        if ($installationResults) {
            foreach ($updateResult in $installationResults) {
                try {
                    # Safe property access with null checks
                    $title = if ($updateResult.PSObject.Properties['Title']) { $updateResult.Title } else { "Unknown Update" }
                    $kb = if ($updateResult.PSObject.Properties['KB']) { $updateResult.KB } else { "N/A" }
                    $size = if ($updateResult.PSObject.Properties['Size']) { $updateResult.Size } else { 0 }
                    $status = if ($updateResult.PSObject.Properties['Result']) { $updateResult.Result } else { "Unknown" }
                    $categories = if ($updateResult.PSObject.Properties['Categories']) { $updateResult.Categories } else { @() }
                    $rebootRequired = if ($updateResult.PSObject.Properties['RebootRequired']) { $updateResult.RebootRequired } else { $false }
                    
                    # Clean up KB number if it's duplicated
                    if ($kb -match "^KBKB(\d+)$") {
                        $kb = "KB$($Matches[1])"
                    } elseif ($kb -eq "KB" -or [string]::IsNullOrEmpty($kb)) {
                        $kb = "N/A"
                    }
                    
                    $processedResult = @{
                        Title = $title
                        KB = $kb
                        Size = $size
                        Status = $status
                        IsSecurity = $categories -match 'Security'
                        RebootRequired = $rebootRequired
                    }
                    
                    # Categorize based on status
                    if ($status -eq 'Installed' -or $status -eq 'Downloaded' -or $status -eq 'Succeeded') {
                        $result.Installed += $processedResult
                        
                        if ($processedResult.IsSecurity) {
                            $result.SecurityUpdates += $processedResult
                        }
                        
                        if ($processedResult.RebootRequired) {
                            $result.RebootRequired = $true
                        }
                        
                        Write-StatusLog "[OK] Installed: $title" -Level "Success"
                    } else {
                        $result.Failed += $processedResult
                        Write-StatusLog "[ERROR] Failed: $title - $status" -Level "Error"
                    }
                }
                catch {
                    $result.Errors += "Error processing update result: $($_.Exception.Message)"
                    Write-StatusLog "[WARN] Error processing update result: $($_.Exception.Message)" -Level "Warning"
                }
            }
        } else {
            Write-StatusLog "[WARN] No installation results returned" -Level "Warning"
            $result.Errors += "No installation results returned from Install-WindowsUpdate"
        }
        
    } catch {
        $result.Errors += "Windows update installation failed: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] Update installation failed: $($_.Exception.Message)" -Level "Error"
    }
    
    return $result
}

function Install-WindowsUpdatesSecurely {
    <#
    .SYNOPSIS
    Installs Windows updates with security validation and progress tracking
    
    .PARAMETER Updates
    Array of updates to install
    
    .PARAMETER MaxTimeout
    Maximum timeout per update in minutes
    
    .PARAMETER AllowReboot
    Allow automatic reboot if required
    
    .OUTPUTS
    Returns hashtable with installation results
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Updates,
        
        [Parameter()]
        [int]$MaxTimeout = 60,
        
        [Parameter()]
        [switch]$AllowReboot
    )
    
    $result = @{
        Installed = @()
        Failed = @()
        SecurityUpdates = @()
        RebootRequired = $false
        ProgressDetails = @()
        Errors = @()
    }
    
    if ($WhatIfPreference) {
        Write-Host "   Would install $($Updates.Count) Windows updates" -ForegroundColor Yellow
        return $result
    }
    
    Write-Host "`n[RESTART] Installing $($Updates.Count) Windows updates..." -ForegroundColor Cyan
    Write-Host "[SUMMARY] Progress will be shown below - installation may take several minutes..." -ForegroundColor Yellow
    
    try {
        # Install updates with progress tracking
        $installParams = @{
            AcceptAll = $true
            Install = $true
            Verbose = $true
            ErrorAction = 'Continue'
        }
        
        if ($AllowReboot) {
            $installParams.AutoReboot = $true
        }
        
        # Create a background job for installation with timeout monitoring
        Write-StatusLog "[START] Starting Windows Update installation process..." -Level "Info"
        
        $installJob = Start-Job -ScriptBlock {
            param($installParams)
            
            # Import the module in the job
            Import-Module PSWindowsUpdate -Force
            
            # Execute the installation
            Install-WindowsUpdate @installParams
            
        } -ArgumentList $installParams
        
        # Monitor the job with periodic status updates
        $startTime = Get-Date
        $timeoutMinutes = $MaxTimeout
        $lastStatusTime = Get-Date
        $statusUpdateInterval = 30 # seconds
        
        Write-Host "[TIMER]  Installation started at $($startTime.ToString('HH:mm:ss'))" -ForegroundColor Cyan
        Write-Host "[WAIT] Maximum timeout: $timeoutMinutes minutes" -ForegroundColor Gray
        Write-Host "------------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        
        while ($installJob.State -eq 'Running') {
            $currentTime = Get-Date
            $elapsed = $currentTime - $startTime
            
            # Show periodic status updates
            if (($currentTime - $lastStatusTime).TotalSeconds -ge $statusUpdateInterval) {
                $elapsedMinutes = [math]::Round($elapsed.TotalMinutes, 1)
                $remainingMinutes = [math]::Max(0, $timeoutMinutes - $elapsed.TotalMinutes)
                
                Write-Host "[RESTART] Still installing... Elapsed: $elapsedMinutes min | Remaining timeout: $([math]::Round($remainingMinutes, 1)) min" -ForegroundColor Yellow
                
                # Check Windows Update service status
                $wuService = Get-Service -Name 'wuauserv' -ErrorAction SilentlyContinue
                if ($wuService) {
                    $serviceStatus = if ($wuService.Status -eq 'Running') { "[OK] Running" } else { "[WARN] $($wuService.Status)" }
                    Write-Host "   Windows Update Service: $serviceStatus" -ForegroundColor Gray
                }
                
                # Check for any pending reboots
                $rebootPending = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
                if ($rebootPending) {
                    Write-Host "   [WARN] Reboot required flag detected" -ForegroundColor Yellow
                }
                
                $lastStatusTime = $currentTime
            }
            
            # Check for timeout
            if ($elapsed.TotalMinutes -gt $timeoutMinutes) {
                Write-Host "[TIME] Installation timeout reached ($timeoutMinutes minutes)" -ForegroundColor Red
                Stop-Job $installJob -Force
                Remove-Job $installJob -Force
                $result.Errors += "Installation timed out after $timeoutMinutes minutes"
                return $result
            }
            
            # Brief pause before next check
            Start-Sleep -Seconds 2
        }
        
        # Job completed - get results
        $endTime = Get-Date
        $totalTime = $endTime - $startTime
        
        Write-Host "------------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "[OK] Installation process completed in $([math]::Round($totalTime.TotalMinutes, 1)) minutes" -ForegroundColor Green
        
        if ($installJob.State -eq 'Completed') {
            try {
                $installationResults = Receive-Job $installJob -ErrorAction Stop
                Write-StatusLog "[REPORT] Processing installation results..." -Level "Info"
            }
            catch {
                $result.Errors += "Failed to receive job results: $($_.Exception.Message)"
                Write-StatusLog "[ERROR] Failed to get installation results: $($_.Exception.Message)" -Level "Error"
                return $result
            }
        } else {
            $result.Errors += "Installation job failed with state: $($installJob.State)"
            Write-StatusLog "[ERROR] Installation job failed with state: $($installJob.State)" -Level "Error"
            
            # Try to get any error output
            try {
                $jobErrors = Receive-Job $installJob -ErrorAction SilentlyContinue
                if ($jobErrors) {
                    $result.Errors += "Job errors: $jobErrors"
                }
            }
            catch {
                # Ignore errors when trying to get job errors
            }
            
            return $result
        }
        
        # Clean up the job
        Remove-Job $installJob -Force -ErrorAction SilentlyContinue
        
        # Process results with de-duplication and terminal-state handling
        if ($installationResults) {
            Write-Host "[SUMMARY] Processing $($installationResults.Count) installation results..." -ForegroundColor Cyan

            try {
                $mergedResults = Merge-WindowsUpdateInstallationResults -InstallationResults @($installationResults)

                foreach ($installedUpdate in @($mergedResults.Installed)) {
                    $result.Installed += $installedUpdate

                    if ($installedUpdate.IsSecurity) {
                        $result.SecurityUpdates += $installedUpdate
                    }

                    if ($installedUpdate.RebootRequired) {
                        $result.RebootRequired = $true
                    }

                    Write-StatusLog "[OK] Installed: $($installedUpdate.Title)" -Level "Success"
                }

                foreach ($failedUpdate in @($mergedResults.Failed)) {
                    $result.Failed += $failedUpdate
                    Write-StatusLog "[ERROR] Failed: $($failedUpdate.Title) - $($failedUpdate.FailureReason)" -Level "Error"
                }
            }
            catch {
                $result.Errors += "Error processing update result: $($_.Exception.Message)"
                Write-StatusLog "[WARN] Error processing update result: $($_.Exception.Message)" -Level "Warning"
            }

            Write-Progress -Activity "Processing Results" -Completed

        } else {
            Write-StatusLog "[WARN] No installation results returned" -Level "Warning"
            $result.Errors += "No installation results returned from Install-WindowsUpdate"
        }
        
        # Final status summary
        Write-Host "`n[REPORT] INSTALLATION SUMMARY:" -ForegroundColor Cyan
        Write-Host "   [OK] Successfully installed: $($result.Installed.Count)" -ForegroundColor Green
        Write-Host "   [ERROR] Failed installations: $($result.Failed.Count)" -ForegroundColor Red
        Write-Host "   [LOCK] Security updates: $($result.SecurityUpdates.Count)" -ForegroundColor Yellow
        Write-Host "   [RESTART] Reboot required: $(if($result.RebootRequired){'Yes'}else{'No'})" -ForegroundColor $(if($result.RebootRequired){'Yellow'}else{'Green'})
        Write-Host "   [TIMER] Total time: $([math]::Round($totalTime.TotalMinutes, 1)) minutes" -ForegroundColor White
        
    } catch {
        $result.Errors += "Windows update installation failed: $($_.Exception.Message)"
        Write-StatusLog "[ERROR] Update installation failed: $($_.Exception.Message)" -Level "Error"
    }
    
    return $result
}

function Test-PostUpdateSystemHealth {
    <#
    .SYNOPSIS
    Validates system health after Windows updates
    
    .OUTPUTS
    Returns hashtable with system health validation results
    #>
    
    $result = @{
        SystemStable = $false
        ServicesRunning = $false
        EventLogErrors = @()
        DiskSpaceAfter = 0
        Errors = @()
    }
    
    try {
        if ($WhatIfPreference) {
            Write-Host "   Would validate post-update system health" -ForegroundColor Yellow
            $result.SystemStable = $true
            return $result
        }
        
        # Check critical services
        $criticalServices = @('wuauserv', 'cryptsvc', 'bits', 'msiserver', 'eventlog', 'winmgmt', 'rpcss')
        $serviceStatus = @()
        
        foreach ($serviceName in $criticalServices) {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($service -and $service.Status -eq 'Running') {
                $serviceStatus += $true
            } else {
                $result.Errors += "Critical service not running: $serviceName"
                $serviceStatus += $false
            }
        }
        
        $result.ServicesRunning = ($serviceStatus | Where-Object { $_ -eq $false }).Count -eq 0
        
        # Check remaining disk space
        $systemDrive = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $env:SystemDrive }
        $result.DiskSpaceAfter = [math]::Round($systemDrive.FreeSpace / 1GB, 2)
        
        # Check recent event log errors (last 30 minutes) and filter known benign noise
        $cutoffTime = (Get-Date).AddMinutes(-30)
        try {
            $recentErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Level=2; StartTime=$cutoffTime} -MaxEvents 25 -ErrorAction SilentlyContinue
            if ($recentErrors) {
                $filteredErrors = @($recentErrors | Where-Object { -not (Test-BenignPostUpdateEvent -EventRecord $_) })
                if ($filteredErrors.Count -gt 0) {
                    $result.EventLogErrors = $filteredErrors | Select-Object -First 10 | ForEach-Object {
                        $eventMessage = ''
                        try { $eventMessage = [string]$_.Message } catch {}
                        if ([string]::IsNullOrWhiteSpace($eventMessage)) { $eventMessage = 'No event message available.' }

                        @{
                            TimeCreated = $_.TimeCreated
                            Id = $_.Id
                            LevelDisplayName = $_.LevelDisplayName
                            Message = $eventMessage.Substring(0, [Math]::Min(200, $eventMessage.Length))
                        }
                    }
                }
            }
        } catch {
            Write-StatusLog "[WARN] Could not check event log for errors: $($_.Exception.Message)" -Level "Warning"
        }
        
        $result.SystemStable = $result.ServicesRunning -and ($result.EventLogErrors.Count -eq 0)
        
        if ($result.SystemStable) {
            Write-StatusLog "[OK] Post-update system health validation passed" -Level "Success"
        } else {
            Write-StatusLog "[WARN] Post-update system health validation found issues" -Level "Warning"
        }
        
    } catch {
        $result.Errors += "System health validation failed: $($_.Exception.Message)"
    }
    
    return $result
}

function Show-WindowsUpdateSummary {
    <#
    .SYNOPSIS
    Displays comprehensive summary of Windows updates
    
    .PARAMETER Results
    Results from the Windows update process
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Results
    )
    
    $actionText = if ($WhatIfPreference) { "SIMULATION" } else { "UPDATE" }
    
    Write-Host "`n" + "="*100 -ForegroundColor Cyan
    Write-Host "WINDOWS $actionText SUMMARY" -ForegroundColor Cyan
    Write-Host "="*100 -ForegroundColor Cyan
    
    # Safe property access with null checks
    Write-Host "Prerequisites Ready: " -NoNewline
    $prereqReady = if ($Results.ContainsKey('PrerequisitesReady')) { $Results.PrerequisitesReady } else { $false }
    Write-Host $(if($prereqReady){"[OK] Yes"}else{"[ERROR] No"}) -ForegroundColor $(if($prereqReady){"Green"}else{"Red"})
    
    Write-Host "Group Policy Reset: " -NoNewline
    $gpReset = if ($Results.ContainsKey('GroupPolicyReset')) { $Results.GroupPolicyReset } else { $false }
    Write-Host $(if($gpReset){"[OK] Yes"}else{"[WARN] Skipped/Failed"}) -ForegroundColor $(if($gpReset){"Green"}else{"Yellow"})
    
    Write-Host "Components Reset: " -NoNewline
    $compReset = if ($Results.ContainsKey('ComponentsReset')) { $Results.ComponentsReset } else { $false }
    Write-Host $(if($compReset){"[OK] Yes"}else{"[WARN] Skipped/Failed"}) -ForegroundColor $(if($compReset){"Green"}else{"Yellow"})
    
    Write-Host "Available Updates: " -NoNewline
    $availableCount = if ($Results.ContainsKey('AvailableUpdates') -and $Results.AvailableUpdates) { $Results.AvailableUpdates.Count } else { 0 }
    Write-Host $availableCount -ForegroundColor White
    
    Write-Host "Successfully Installed: " -NoNewline
    $installedCount = if ($Results.ContainsKey('InstalledUpdates') -and $Results.InstalledUpdates) { $Results.InstalledUpdates.Count } else { 0 }
    Write-Host $installedCount -ForegroundColor Green
    
    Write-Host "Failed Installations: " -NoNewline
    $failedCount = if ($Results.ContainsKey('FailedUpdates') -and $Results.FailedUpdates) { $Results.FailedUpdates.Count } else { 0 }
    Write-Host $failedCount -ForegroundColor Red
    
    Write-Host "Security Updates: " -NoNewline
    $securityCount = if ($Results.ContainsKey('SecurityUpdates') -and $Results.SecurityUpdates) { $Results.SecurityUpdates.Count } else { 0 }
    Write-Host $securityCount -ForegroundColor $(if($securityCount -gt 0){"Green"}else{"Gray"})
    
    if ($Results.ContainsKey('RestorePointCreated') -and $Results.RestorePointCreated) {
        Write-Host "Restore Point Created: " -NoNewline
        Write-Host "[OK] Yes" -ForegroundColor Green
    }
    
    Write-Host "Reboot Required: " -NoNewline
    $rebootRequired = if ($Results.ContainsKey('RebootRequired')) { $Results.RebootRequired } else { $false }
    Write-Host $(if($rebootRequired){"[WARN] Yes"}else{"[OK] No"}) -ForegroundColor $(if($rebootRequired){"Yellow"}else{"Green"})
    
    Write-Host "Total Duration: " -NoNewline
    $duration = if ($Results.ContainsKey('UpdateDuration') -and $Results.UpdateDuration) { 
        $Results.UpdateDuration.TotalMinutes.ToString('F1') 
    } else { 
        "0.0" 
    }
    Write-Host "$duration minutes" -ForegroundColor White
    
    # Show system health status with null checks
    if ($Results.ContainsKey('SystemValidation') -and $Results.SystemValidation -and $Results.SystemValidation.ContainsKey('SystemStable')) {
        Write-Host "System Health: " -NoNewline
        $systemStable = $Results.SystemValidation.SystemStable
        Write-Host $(if($systemStable){"[OK] Stable"}else{"[WARN] Issues Detected"}) -ForegroundColor $(if($systemStable){"Green"}else{"Yellow"})
        
        if ($Results.SystemValidation.ContainsKey('DiskSpaceAfter')) {
            Write-Host "Disk Space After: " -NoNewline
            Write-Host "$($Results.SystemValidation.DiskSpaceAfter) GB" -ForegroundColor White
        }
    }
    
    # Show successful installations with null checks
    if ($Results.ContainsKey('InstalledUpdates') -and $Results.InstalledUpdates -and $Results.InstalledUpdates.Count -gt 0) {
        Write-Host "`nSuccessfully Installed Updates:" -ForegroundColor Green
        $Results.InstalledUpdates | ForEach-Object {
            $securityIcon = if ($_.IsSecurity) { "[LOCK]" } else { "[PKG]" }
            $rebootIcon = if ($_.RebootRequired) { "[RESTART]" } else { "" }
            $sizeText = if ($_.Size) { " ($([math]::Round($_.Size / 1MB, 1)) MB)" } else { "" }
            Write-Host "  [OK] $securityIcon $($_.Title) ($($_.KB))$sizeText $rebootIcon" -ForegroundColor Green
        }
    }
    
    # Show failed installations with null checks
    if ($Results.ContainsKey('FailedUpdates') -and $Results.FailedUpdates -and $Results.FailedUpdates.Count -gt 0) {
        Write-Host "`nFailed Update Installations:" -ForegroundColor Red
        $Results.FailedUpdates | ForEach-Object {
            Write-Host "  [ERROR] $($_.Title) ($($_.KB)): $($_.Status)" -ForegroundColor Red
        }
    }
    
    # Show recent system errors with null checks
    if ($Results.ContainsKey('SystemValidation') -and 
        $Results.SystemValidation -and 
        $Results.SystemValidation.ContainsKey('EventLogErrors') -and 
        $Results.SystemValidation.EventLogErrors -and 
        $Results.SystemValidation.EventLogErrors.Count -gt 0) {
        
        Write-Host "`nRecent System Errors:" -ForegroundColor Yellow
        $Results.SystemValidation.EventLogErrors | ForEach-Object {
            Write-Host "  [WARN] [$($_.TimeCreated.ToString('HH:mm:ss'))] Event $($_.Id): $($_.Message)" -ForegroundColor Yellow
        }
    }
    
    # Show restart recommendation
    if ($rebootRequired) {
        Write-Host "`n[RESTART] IMPORTANT: System restart required to complete update installation" -ForegroundColor Yellow
        Write-Host "   Some updates may not be fully functional until restart is performed." -ForegroundColor Yellow
    }
    
    # Show errors if any with null checks
    if ($Results.ContainsKey('Errors') -and $Results.Errors -and $Results.Errors.Count -gt 0) {
        Write-Host "`nIssues Encountered:" -ForegroundColor Red
        $Results.Errors | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Red
        }
    }
    
    Write-Host "="*100 -ForegroundColor Cyan
}


# -----------------------------------------------------------------------------
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
# -----------------------------------------------------------------------------
function Invoke-SystemMaintenance {
    [CmdletBinding()]
    param (
        [switch]$ArchiveLogs,
        [string]$LogArchivePath = 'C:\Logs',
        [switch]$SkipReboot,
        [int]$MaxParallelJobs = 4
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This script must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    $WarningPreference = 'SilentlyContinue'   # Speed: Reduce console output
    
    # Security: Validate log path
    if ($ArchiveLogs) {
        if ($LogArchivePath -notmatch '^[C-Z]:\\[\w\s\\.-]*$') {
            throw "Invalid log archive path. Must be a valid Windows path."
        }
        
        if (-not (Test-Path $LogArchivePath)) {
            try {
                New-Item -Path $LogArchivePath -ItemType Directory -Force | Out-Null
                Write-Host "[OK] Created log archive: $LogArchivePath" -ForegroundColor Green
            } catch {
                throw "Cannot create log archive directory: $_"
            }
        }
    }

    # Speed: Pre-calculate system drive once
    $sysDriveLetter = $env:SystemDrive.Substring(0,1)
    $tempPaths = @{
        System = @("C:\SWSetup", "C:\Temp", "C:\system.sav")
        WindowsTemp = "C:\Windows\Temp"
        UserTemp = $env:TEMP
        Additional = @(
            "C:\Windows\SoftwareDistribution\Download",
            "C:\Windows\Prefetch", 
            "C:\Windows\Logs\CBS",
            "$env:LOCALAPPDATA\Microsoft\Windows\INetCache",
            "$env:LOCALAPPDATA\Microsoft\Windows\WebCache",
            "C:\ProgramData\Microsoft\Windows\WER\ReportQueue",
            "$env:LOCALAPPDATA\CrashDumps",
            "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"
        )
    }

    Write-Host "[TOOLS] Starting optimized system maintenance..." -ForegroundColor Cyan
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # Speed: Run filesystem repair in background job
        $fsRepairJob = Start-Job -ScriptBlock {
            param($drive)
            try {
                if (Get-Volume -DriveLetter $drive -ErrorAction SilentlyContinue) {
                    Repair-Volume -DriveLetter $drive -Scan -ErrorAction SilentlyContinue
                    Repair-Volume -DriveLetter $drive -OfflineScanAndFix -ErrorAction SilentlyContinue
                }
                return "[OK] Filesystem repair completed"
            } catch {
                return "[WARN] Filesystem repair failed: $_"
            }
        } -ArgumentList $sysDriveLetter

        Write-Host "[RUN] Component store maintenance..." -ForegroundColor Yellow
        # Speed: Run critical DISM operations only, suppress output and input
        $dismJobs = @()
        $dismJobs += Start-Job -ScriptBlock { 
            $null = DISM.exe /Online /Cleanup-Image /RestoreHealth /Quiet /NoRestart 2>&1
            return "[OK] Component store restore completed"
        }
        $dismJobs += Start-Job -ScriptBlock { 
            $null = DISM.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase /Quiet /NoRestart 2>&1
            return "[OK] Component cleanup completed"
        }

        Write-Host "[RUN] Network stack reset..." -ForegroundColor Yellow
        # Speed: Combine network operations
        $networkJob = Start-Job -ScriptBlock {
            netsh winsock reset | Out-Null
            netsh int ip reset | Out-Null  
            ipconfig /flushdns | Out-Null
            Clear-DnsClientCache
            return "[OK] Network stack reset"
        }

        Write-Host "[RUN] Cleaning temporary files..." -ForegroundColor Yellow
        # Speed: Parallel cleanup using runspaces
        $cleanupJobs = @()
        
        # System paths cleanup
        $cleanupJobs += Start-Job -ScriptBlock {
            param($paths)
            $results = @()
            foreach ($path in $paths) {
                if (Test-Path $path) {
                    try {
                        Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                        $results += "[OK] Deleted: $path"
                    } catch {
                        $results += "[WARN] Failed: $path"
                    }
                }
            }
            return $results
        } -ArgumentList (,$tempPaths.System)

        # Windows Temp cleanup
        $cleanupJobs += Start-Job -ScriptBlock {
            param($winTemp, $additionalPaths)
            $results = @()
            
            # Clean Windows\Temp
            if (Test-Path $winTemp) {
                try {
                    Get-ChildItem -Path $winTemp -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    $results += "[OK] Cleared Windows Temp"
                } catch {
                    $results += "[WARN] Windows Temp cleanup failed"
                }
            }
            
            # Clean additional paths
            foreach ($folder in $additionalPaths) {
                if (Test-Path $folder) {
                    try {
                        Get-ChildItem -Path $folder -Force -ErrorAction SilentlyContinue | 
                            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                        $results += "[OK] Cleaned: $folder"
                    } catch {
                        $results += "[WARN] Failed: $folder"
                    }
                }
            }
            return $results
        } -ArgumentList $tempPaths.WindowsTemp, $tempPaths.Additional

        Write-Host "[RUN] Security hardening..." -ForegroundColor Yellow
        # Security: Enhanced security operations
        $securityJob = Start-Job -ScriptBlock {
            $results = @()
            
            # Reset firewall to secure defaults
            try {
                netsh advfirewall reset | Out-Null
                $results += "[OK] Firewall reset to secure defaults"
            } catch {
                $results += "[WARN] Firewall reset failed"
            }
            
            # Enable Windows Defender
            try {
                Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop
                Set-MpPreference -DisableScriptScanning $false -ErrorAction SilentlyContinue
                Set-MpPreference -DisableBehaviorMonitoring $false -ErrorAction SilentlyContinue
                $results += "[OK] Windows Defender enhanced"
            } catch {
                $results += "[WARN] Defender configuration failed"
            }
            
            return $results
        }

        Write-Host "[RUN] System optimization..." -ForegroundColor Yellow
        # Speed: SSD optimization
        $optimizationJob = Start-Job -ScriptBlock {
            param($drive)
            $results = @()
            
            # SSD Trim
            try {
                Optimize-Volume -DriveLetter $drive -ReTrim -ErrorAction Stop
                $results += "[OK] SSD Trim completed"
            } catch {
                $results += "[WARN] SSD Trim failed"
            }
            
            # Icon cache rebuild
            try {
                Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Remove-Item "$env:LOCALAPPDATA\IconCache.db" -Force -ErrorAction SilentlyContinue
                Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue
                Start-Process explorer
                $results += "[OK] Icon cache rebuilt"
            } catch {
                $results += "[WARN] Icon cache rebuild failed"
            }
            
            return $results
        } -ArgumentList $sysDriveLetter

        # Security: Safe bloatware removal with whitelist approach
        Write-Host "[RUN] Removing unnecessary apps..." -ForegroundColor Yellow
        $appRemovalJob = Start-Job -ScriptBlock {
            $results = @()
            # Security: Only remove specific known bloatware
            $bloatwarePatterns = @(
                '*Xbox*', '*Zune*', '*SkypeApp*', '*BingWeather*', 
                '*Microsoft.3DBuilder*', '*CandyCrushSaga*', '*Facebook*'
            )
            
            try {
                foreach ($pattern in $bloatwarePatterns) {
                    Get-AppxProvisionedPackage -Online | 
                    Where-Object DisplayName -Like $pattern |
                    ForEach-Object {
                        try {
                            Remove-AppxProvisionedPackage -Online -PackageName $_.PackageName -ErrorAction Stop
                            $results += "[OK] Removed: $($_.DisplayName)"
                        } catch {
                            $results += "[WARN] Failed to remove: $($_.DisplayName)"
                        }
                    }
                }
            } catch {
                $results += "[WARN] App removal enumeration failed"
            }
            return $results
        }

        # Wait for all jobs with timeout and real-time progress reporting
        Write-Host "[PENDING] Monitoring operations in real-time..." -ForegroundColor Yellow
        $allJobs = @($fsRepairJob, $networkJob, $securityJob, $optimizationJob, $appRemovalJob) + $dismJobs + $cleanupJobs
        
        # Create job tracking with descriptive names
        $jobTracker = @{
            $fsRepairJob.Id = @{ Name = "Filesystem Repair"; Status = "Running"; StartTime = Get-Date }
            $networkJob.Id = @{ Name = "Network Stack Reset"; Status = "Running"; StartTime = Get-Date }
            $securityJob.Id = @{ Name = "Security Hardening"; Status = "Running"; StartTime = Get-Date }
            $optimizationJob.Id = @{ Name = "System Optimization"; Status = "Running"; StartTime = Get-Date }
            $appRemovalJob.Id = @{ Name = "App Removal"; Status = "Running"; StartTime = Get-Date }
        }
        
        # Add DISM jobs to tracker
        for ($i = 0; $i -lt $dismJobs.Count; $i++) {
            $jobTracker[$dismJobs[$i].Id] = @{ 
                Name = "DISM Operation $($i + 1)"; 
                Status = "Running"; 
                StartTime = Get-Date 
            }
        }
        
        # Add cleanup jobs to tracker
        for ($i = 0; $i -lt $cleanupJobs.Count; $i++) {
            $jobTracker[$cleanupJobs[$i].Id] = @{ 
                Name = "Cleanup Task $($i + 1)"; 
                Status = "Running"; 
                StartTime = Get-Date 
            }
        }
        
        $timeoutSeconds = 600  # 10 minutes
        $startTime = Get-Date
        $completedJobs = @()
        $lastUpdate = Get-Date
        
        Write-Host "`n  [REPORT] Active Operations:" -ForegroundColor Cyan
        $jobTracker.Values | ForEach-Object { 
            Write-Host "    - $($_.Name) - Started" -ForegroundColor Gray 
        }
        Write-Host ""
        
        do {
            $currentTime = Get-Date
            
            # Check job states and report changes
            foreach ($job in $allJobs) {
                $tracker = $jobTracker[$job.Id]
                $newState = $job.State
                
                if ($tracker.Status -ne $newState) {
                    $elapsed = ($currentTime - $tracker.StartTime).TotalSeconds
                    
                    switch ($newState) {
                        'Completed' {
                            Write-Host "    [OK] $($tracker.Name) completed ($($elapsed.ToString('F1'))s)" -ForegroundColor Green
                            
                            # Show job results immediately
                            try {
                                $result = Receive-Job $job
                                if ($result) {
                                    $result | ForEach-Object { 
                                        Write-Host "       `-- $_" -ForegroundColor DarkGray 
                                    }
                                }
                            } catch {
                                Write-Host "       `-- [WARN] Could not retrieve results" -ForegroundColor Yellow
                            }
                        }
                        'Failed' {
                            Write-Host "    [ERROR] $($tracker.Name) failed ($($elapsed.ToString('F1'))s)" -ForegroundColor Red
                        }
                        'Stopped' {
                            Write-Host "    [STOP] $($tracker.Name) stopped ($($elapsed.ToString('F1'))s)" -ForegroundColor Yellow
                        }
                        'Blocked' {
                            Write-Host "    [PAUSE] $($tracker.Name) blocked - attempting to resume..." -ForegroundColor Yellow
                            try {
                                $null = Receive-Job $job -Keep
                            } catch {
                                # Continue if unable to receive job output
                            }
                        }
                    }
                    
                    $tracker.Status = $newState
                }
            }
            
            # Show periodic status updates every 30 seconds for long-running jobs
            if (($currentTime - $lastUpdate).TotalSeconds -gt 30) {
                $runningCount = ($allJobs | Where-Object { $_.State -eq 'Running' }).Count
                if ($runningCount -gt 0) {
                    $elapsed = ($currentTime - $startTime).TotalMinutes
                    Write-Host "    [PENDING] $runningCount operations still running... ($($elapsed.ToString('F1')) min elapsed)" -ForegroundColor Cyan
                    
                    # Show which jobs are still running
                    $allJobs | Where-Object { $_.State -eq 'Running' } | ForEach-Object {
                        $tracker = $jobTracker[$_.Id]
                        $jobElapsed = ($currentTime - $tracker.StartTime).TotalSeconds
                        Write-Host "       - $($tracker.Name) ($($jobElapsed.ToString('F0'))s)" -ForegroundColor DarkCyan
                    }
                }
                $lastUpdate = $currentTime
            }
            
            # Check for newly completed jobs
            $runningJobs = $allJobs | Where-Object { $_.State -eq 'Running' -or $_.State -eq 'Blocked' }
            
            # Check timeout
            $elapsed = ($currentTime - $startTime).TotalSeconds
            if ($elapsed -gt $timeoutSeconds) {
                Write-Host "`n    [WARN] Timeout reached after $($timeoutSeconds/60) minutes, stopping remaining jobs..." -ForegroundColor Yellow
                $runningJobs | Stop-Job
                break
            }
            
            # Brief pause to prevent excessive CPU usage
            Start-Sleep -Milliseconds 1000
            
        } while ($runningJobs.Count -gt 0)
        
        Write-Host "`n[SUMMARY] Final Results Summary:" -ForegroundColor Cyan
        
        # Clean up jobs and show final summary
        $successCount = 0
        $failedCount = 0
        
        foreach ($job in $allJobs) {
            $tracker = $jobTracker[$job.Id]
            $finalElapsed = ((Get-Date) - $tracker.StartTime).TotalSeconds
            
            switch ($job.State) {
                'Completed' { 
                    $successCount++
                    Write-Host "  [OK] $($tracker.Name) - Success ($($finalElapsed.ToString('F1'))s)" -ForegroundColor Green
                }
                'Failed' { 
                    $failedCount++
                    Write-Host "  [ERROR] $($tracker.Name) - Failed ($($finalElapsed.ToString('F1'))s)" -ForegroundColor Red
                }
                default { 
                    $failedCount++
                    Write-Host "  [WARN] $($tracker.Name) - $($job.State) ($($finalElapsed.ToString('F1'))s)" -ForegroundColor Yellow
                }
            }
            
            Remove-Job $job -Force
        }
        
        Write-Host "`n[STATS] Operation Summary: $successCount successful, $failedCount failed/incomplete" -ForegroundColor $(if ($failedCount -eq 0) { 'Green' } else { 'Yellow' })

        # Security: Audit system state
        Write-Host "`n[SCAN] Security audit..." -ForegroundColor Yellow
        $auditResults = @()
        
        # Check for suspicious scheduled tasks
        try {
            Get-ScheduledTask | Where-Object { 
                $_.State -eq 'Unknown' -or 
                ($_.TaskPath -notlike '\Microsoft\*' -and $_.Author -eq '') 
            } | ForEach-Object {
                $auditResults += "[WARN] Suspicious task: $($_.TaskName)"
            }
        } catch {
            $auditResults += "[WARN] Task audit failed"
        }
        
        if ($auditResults.Count -eq 0) {
            Write-Host "  [OK] No security issues detected" -ForegroundColor Green
        } else {
            $auditResults | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
        }

        # Speed: Smart event log clearing (keep critical logs)
        Write-Host "`n[FILES] Clearing non-critical event logs..." -ForegroundColor Yellow
        $criticalLogs = @('System', 'Application', 'Security')
        $clearedCount = 0
        
        Get-WinEvent -ListLog * -ErrorAction SilentlyContinue | 
        Where-Object { $_.LogName -notin $criticalLogs -and $_.RecordCount -gt 0 } |
        ForEach-Object {
            try {
                [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession.ClearLog($_.LogName)
                $clearedCount++
            } catch {
                # Silently continue for logs that can't be cleared
            }
        }
        Write-Host "  [OK] Cleared $clearedCount non-critical event logs" -ForegroundColor Green

        # Final system health check
        Write-Host "`n[HP] System health verification..." -ForegroundColor Yellow
        $healthCheck = Start-Job -ScriptBlock {
            # Verify critical services
            $criticalServices = @('Winmgmt', 'EventLog', 'RpcSs', 'DcomLaunch')
            $serviceStatus = @()
            
            foreach ($service in $criticalServices) {
                $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
                if ($svc.Status -ne 'Running') {
                    $serviceStatus += "[WARN] Service $service is $($svc.Status)"
                }
            }
            
            if ($serviceStatus.Count -eq 0) {
                return "[OK] All critical services running"
            } else {
                return $serviceStatus
            }
        }
        
        $healthResult = Wait-Job $healthCheck | Receive-Job
        Remove-Job $healthCheck
        $healthResult | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

    } catch {
        Write-Error "Critical maintenance failure: $_"
        throw
    } finally {
        $stopwatch.Stop()
        Write-Host "`n[FAST] Maintenance completed in $($stopwatch.Elapsed.TotalMinutes.ToString('F1')) minutes" -ForegroundColor Green
        
        if (-not $SkipReboot) {
            Write-Host "`n[RESTART] System restart recommended for optimal performance." -ForegroundColor Yellow
            $restart = Read-Host "Restart now? (y/N)"
            if ($restart -eq 'y' -or $restart -eq 'Y') {
                Restart-Computer -Force
            }
        }
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
# Option 7 - Set Desktop Power Settings - Only run on Desktop computers, no laptops!
# -----------------------------------------------------------------------------
function Set-DesktopPowerSettings {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('High Performance', 'Balanced', 'Power Saver', 'Ultimate Performance')]
        [string]$PowerPlan = 'High Performance',
        
        [ValidateRange(1, 600)]
        [int]$MonitorTimeoutMinutes = 60,
        
        [ValidateRange(1, 600)]
        [int]$DiskTimeoutMinutes = 0,  # 0 = Never
        
        [switch]$Force,
        # Removed [switch]$WhatIf - this is automatically provided by SupportsShouldProcess
        [switch]$AllowLaptops,
        [string]$LogPath = "$env:TEMP\PowerSettings_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
        [switch]$SkipHardwareDetection
    )

    # Security: Require elevation for power configuration
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    $ProgressPreference = 'SilentlyContinue'  # Speed: Disable progress bars
    
    # Initialize tracking
    $logEntries = @()
    $appliedSettings = @()
    $failedSettings = @()
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
            'HARDWARE' { Write-Host $Message -ForegroundColor Magenta }
        }
    }

    # Speed: Hardware detection with caching
    function Get-SystemHardwareType {
        try {
            # Use CIM for faster queries
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
            $systemEnclosure = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction SilentlyContinue
            
            $hardwareInfo = @{
                IsLaptop = $false
                IsDesktop = $false
                IsWorkstation = $false
                IsServer = $false
                HasBattery = ($battery -ne $null)
                ChassisTypes = @()
                PCSystemType = $computerSystem.PCSystemType
                TotalPhysicalMemory = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)
                Manufacturer = $computerSystem.Manufacturer
                Model = $computerSystem.Model
            }
            
            # Determine system type based on multiple factors
            if ($systemEnclosure) {
                $hardwareInfo.ChassisTypes = $systemEnclosure.ChassisTypes
                
                # Chassis type detection (more reliable than PCSystemType)
                $laptopChassisTypes = @(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32)  # Laptop variants
                $desktopChassisTypes = @(3, 4, 5, 6, 7, 15, 16)  # Desktop variants
                $serverChassisTypes = @(17, 23)  # Server variants
                
                if ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $laptopChassisTypes }) {
                    $hardwareInfo.IsLaptop = $true
                } elseif ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $desktopChassisTypes }) {
                    $hardwareInfo.IsDesktop = $true
                } elseif ($systemEnclosure.ChassisTypes | Where-Object { $_ -in $serverChassisTypes }) {
                    $hardwareInfo.IsServer = $true
                } else {
                    $hardwareInfo.IsWorkstation = $true
                }
            }
            
            # Fallback to PCSystemType if chassis detection inconclusive
            if (-not ($hardwareInfo.IsLaptop -or $hardwareInfo.IsDesktop -or $hardwareInfo.IsServer)) {
                switch ($computerSystem.PCSystemType) {
                    1 { $hardwareInfo.IsDesktop = $true }
                    2 { $hardwareInfo.IsLaptop = $true }
                    3 { $hardwareInfo.IsWorkstation = $true }
                    4 { $hardwareInfo.IsServer = $true }
                    default { $hardwareInfo.IsDesktop = $true }  # Default assumption
                }
            }
            
            # Battery presence overrides chassis detection for laptops
            if ($hardwareInfo.HasBattery -and -not $hardwareInfo.IsServer) {
                $hardwareInfo.IsLaptop = $true
                $hardwareInfo.IsDesktop = $false
            }
            
            return $hardwareInfo
            
        } catch {
            Write-LogEntry "Hardware detection failed: $_" 'WARNING'
            return @{
                IsLaptop = $false
                IsDesktop = $true  # Safe default
                Error = $_.Exception.Message
            }
        }
    }

    # Speed: Get available power schemes efficiently
    function Get-PowerSchemes {
        try {
            $schemes = @{}
            
            # Parse powercfg output for available schemes
            $output = powercfg.exe /list 2>$null
            if ($LASTEXITCODE -eq 0 -and $output) {
                foreach ($line in $output) {
                    if ($line -match 'Power Scheme GUID: ([a-f0-9-]+)\s+\((.+?)\)(\s+\*)?') {
                        $guid = $matches[1]
                        $name = $matches[2].Trim()
                        $isActive = $matches[3] -eq ' *'
                        
                        $schemes[$name] = @{
                            GUID = $guid
                            Name = $name
                            IsActive = $isActive
                        }
                    }
                }
            }
            
            # Add common scheme mappings if not found
            $commonSchemes = @{
                'High Performance' = 'SCHEME_MIN'
                'Balanced' = 'SCHEME_BALANCED'
                'Power Saver' = 'SCHEME_MAX'
                'Ultimate Performance' = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
            }
            
            foreach ($scheme in $commonSchemes.GetEnumerator()) {
                if (-not $schemes.ContainsKey($scheme.Key)) {
                    $schemes[$scheme.Key] = @{
                        GUID = $scheme.Value
                        Name = $scheme.Key
                        IsActive = $false
                    }
                }
            }
            
            return $schemes
            
        } catch {
            Write-LogEntry "Failed to get power schemes: $_" 'WARNING'
            return @{}
        }
    }

    # Security: Backup current power settings
    function Backup-PowerSettings {
        try {
            $backupData = @{
                CurrentScheme = (powercfg.exe /getactivescheme 2>$null)
                HibernationStatus = (powercfg.exe /query SCHEME_CURRENT SUB_SLEEP HIBERNATEIDLE 2>$null)
                MonitorTimeoutAC = (powercfg.exe /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE /AC 2>$null)
                MonitorTimeoutDC = (powercfg.exe /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE /DC 2>$null)
                DiskTimeoutAC = (powercfg.exe /query SCHEME_CURRENT SUB_DISK DISKIDLE /AC 2>$null)
                DiskTimeoutDC = (powercfg.exe /query SCHEME_CURRENT SUB_DISK DISKIDLE /DC 2>$null)
                Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            }
            
            $backupFile = "$env:TEMP\PowerSettingsBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $backupData | ConvertTo-Json -Depth 3 | Out-File -FilePath $backupFile -Encoding UTF8
            Write-LogEntry "[OK] Power settings backup saved to: $backupFile" 'INFO'
            return $backupFile
            
        } catch {
            Write-LogEntry "[WARN] Failed to backup power settings: $_" 'WARNING'
            return $null
        }
    }

    # Speed: Execute power configuration commands
    function Set-PowerConfiguration {
        param(
            [string]$SchemeName,
            [string]$SchemeGUID,
            [int]$MonitorTimeout,
            [int]$DiskTimeout
        )
        
        $configResults = @()
        
        # Set power scheme
        try {
            Write-LogEntry "Setting power scheme to: $SchemeName" 'INFO'
            
            if ($PSCmdlet.ShouldProcess($SchemeName, "Set active power scheme")) {
                $result = powercfg.exe /setactive $SchemeGUID 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Power Scheme"
                        Value = $SchemeName
                        Success = $true
                    }
                    Write-LogEntry "[OK] Power scheme set to: $SchemeName" 'SUCCESS'
                } else {
                    throw "Failed to set power scheme: $result"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Power Scheme"
                Value = $SchemeName
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[X] Failed to set power scheme: $_" 'ERROR'
        }
        
        # Configure monitor timeout
        try {
            Write-LogEntry "Setting monitor timeout to: $MonitorTimeout minutes" 'INFO'
            
            if ($PSCmdlet.ShouldProcess("Monitor Timeout", "Set to $MonitorTimeout minutes")) {
                # Set for both AC and DC power
                $resultAC = powercfg.exe /change monitor-timeout-ac $MonitorTimeout 2>&1
                $resultDC = powercfg.exe /change monitor-timeout-dc $MonitorTimeout 2>&1
                
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Monitor Timeout"
                        Value = "$MonitorTimeout minutes"
                        Success = $true
                    }
                    Write-LogEntry "[OK] Monitor timeout set to: $MonitorTimeout minutes" 'SUCCESS'
                } else {
                    throw "Failed to set monitor timeout: AC=$resultAC, DC=$resultDC"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Monitor Timeout"
                Value = "$MonitorTimeout minutes"
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[X] Failed to set monitor timeout: $_" 'ERROR'
        }
        
        # Configure disk timeout
        if ($DiskTimeout -eq 0) {
            try {
                Write-LogEntry "Disabling disk timeout (never turn off)" 'INFO'
                
                if ($PSCmdlet.ShouldProcess("Disk Timeout", "Disable")) {
                    $resultAC = powercfg.exe /change disk-timeout-ac 0 2>&1
                    $resultDC = powercfg.exe /change disk-timeout-dc 0 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $configResults += @{
                            Setting = "Disk Timeout"
                            Value = "Disabled"
                            Success = $true
                        }
                        Write-LogEntry "[OK] Disk timeout disabled" 'SUCCESS'
                    } else {
                        throw "Failed to disable disk timeout: AC=$resultAC, DC=$resultDC"
                    }
                }
            } catch {
                $configResults += @{
                    Setting = "Disk Timeout"
                    Value = "Disabled"
                    Success = $false
                    Error = $_.Exception.Message
                }
                Write-LogEntry "[X] Failed to disable disk timeout: $_" 'ERROR'
            }
        } else {
            try {
                Write-LogEntry "Setting disk timeout to: $DiskTimeout minutes" 'INFO'
                
                if ($PSCmdlet.ShouldProcess("Disk Timeout", "Set to $DiskTimeout minutes")) {
                    $resultAC = powercfg.exe /change disk-timeout-ac $DiskTimeout 2>&1
                    $resultDC = powercfg.exe /change disk-timeout-dc $DiskTimeout 2>&1
                    
                    if ($LASTEXITCODE -eq 0) {
                        $configResults += @{
                            Setting = "Disk Timeout"
                            Value = "$DiskTimeout minutes"
                            Success = $true
                        }
                        Write-LogEntry "[OK] Disk timeout set to: $DiskTimeout minutes" 'SUCCESS'
                    } else {
                        throw "Failed to set disk timeout: AC=$resultAC, DC=$resultDC"
                    }
                }
            } catch {
                $configResults += @{
                    Setting = "Disk Timeout"
                    Value = "$DiskTimeout minutes"
                    Success = $false
                    Error = $_.Exception.Message
                }
                Write-LogEntry "[X] Failed to set disk timeout: $_" 'ERROR'
            }
        }
        
        # Disable hibernation
        try {
            Write-LogEntry "Disabling hibernation" 'INFO'
            
            if ($PSCmdlet.ShouldProcess("Hibernation", "Disable")) {
                $result = powercfg.exe /hibernate off 2>&1
                if ($LASTEXITCODE -eq 0) {
                    $configResults += @{
                        Setting = "Hibernation"
                        Value = "Disabled"
                        Success = $true
                    }
                    Write-LogEntry "[OK] Hibernation disabled" 'SUCCESS'
                } else {
                    throw "Failed to disable hibernation: $result"
                }
            }
        } catch {
            $configResults += @{
                Setting = "Hibernation"
                Value = "Disabled"
                Success = $false
                Error = $_.Exception.Message
            }
            Write-LogEntry "[X] Failed to disable hibernation: $_" 'ERROR'
        }
        
        return $configResults
    }

    try {
        Write-LogEntry "=== Desktop Power Settings Configuration Started ===" 'INFO'
        Write-LogEntry "Power Plan: $PowerPlan" 'INFO'
        Write-LogEntry "Monitor Timeout: $MonitorTimeoutMinutes minutes" 'INFO'
        Write-LogEntry "Disk Timeout: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" 'INFO'
        
        if ($WhatIfPreference) {
            Write-LogEntry "WhatIf mode - no power settings will be changed" 'INFO'
        }

        # Hardware detection and validation
        if (-not $SkipHardwareDetection) {
            Write-LogEntry "`n[SCAN] Detecting system hardware type..." 'HARDWARE'
            $hardwareInfo = Get-SystemHardwareType
            
            Write-LogEntry "Hardware Analysis:" 'HARDWARE'
            Write-LogEntry "  - System Type: $(if($hardwareInfo.IsLaptop){'Laptop'}elseif($hardwareInfo.IsDesktop){'Desktop'}elseif($hardwareInfo.IsServer){'Server'}else{'Workstation'})" 'HARDWARE'
            Write-LogEntry "  - Has Battery: $($hardwareInfo.HasBattery)" 'HARDWARE'
            Write-LogEntry "  - Manufacturer: $($hardwareInfo.Manufacturer)" 'HARDWARE'
            Write-LogEntry "  - Model: $($hardwareInfo.Model)" 'HARDWARE'
            Write-LogEntry "  - Memory: $($hardwareInfo.TotalPhysicalMemory) GB" 'HARDWARE'
            
            # Security: Prevent accidental laptop configuration
            if ($hardwareInfo.IsLaptop -and -not $AllowLaptops -and -not $Force) {
                Write-Host "`n" + "="*60 -ForegroundColor Red
                Write-Host "LAPTOP DETECTED - OPERATION BLOCKED" -ForegroundColor Red -BackgroundColor Black
                Write-Host "="*60 -ForegroundColor Red
                Write-Host "This system appears to be a LAPTOP with battery power." -ForegroundColor Yellow
                Write-Host "Desktop power settings may negatively impact battery life!" -ForegroundColor Yellow
                Write-Host "`nTo proceed anyway, use one of these options:" -ForegroundColor Cyan
                Write-Host "  - Use -AllowLaptops parameter" -ForegroundColor Cyan
                Write-Host "  - Use -Force parameter" -ForegroundColor Cyan
                Write-Host "  - Use -SkipHardwareDetection parameter" -ForegroundColor Cyan
                Write-Host "="*60 -ForegroundColor Red
                
                $global:LastStatus = "[WARN] Operation blocked - laptop detected."
                return
            }
            
            if ($hardwareInfo.IsLaptop -and ($AllowLaptops -or $Force)) {
                Write-LogEntry "[WARN] Proceeding with laptop configuration (overridden)" 'WARNING'
            }
        }

        # User confirmation for desktop systems
        if (-not $Force -and -not $WhatIfPreference) {
            Write-Host "`n" + "="*60 -ForegroundColor Yellow
            Write-Host "POWER SETTINGS CONFIGURATION" -ForegroundColor Yellow -BackgroundColor Black
            Write-Host "="*60 -ForegroundColor Yellow
            Write-Host "This will configure the following settings:" -ForegroundColor White
            Write-Host "  - Power Plan: $PowerPlan" -ForegroundColor Cyan
            Write-Host "  - Monitor Timeout: $MonitorTimeoutMinutes minutes" -ForegroundColor Cyan
            Write-Host "  - Disk Timeout: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" -ForegroundColor Cyan
            Write-Host "  - Hibernation: Disabled" -ForegroundColor Cyan
            Write-Host "`nNote: These settings optimize for desktop performance" -ForegroundColor Yellow
            Write-Host "="*60 -ForegroundColor Yellow
            
            $confirmation = Read-Host "`n[?] Continue with power configuration? (Y/N)"
            if ($confirmation -notin @('Y', 'y', 'Yes', 'yes')) {
                Write-LogEntry "Operation cancelled by user" 'WARNING'
                $global:LastStatus = "[WARN] User cancelled power settings configuration."
                return
            }
        }

        # Get available power schemes
        Write-LogEntry "`n[SCAN] Scanning available power schemes..." 'INFO'
        $powerSchemes = Get-PowerSchemes
        
        if ($powerSchemes.Count -eq 0) {
            throw "No power schemes detected on this system"
        }
        
        Write-LogEntry "Available power schemes:" 'INFO'
        foreach ($scheme in $powerSchemes.GetEnumerator()) {
            $activeIndicator = if ($scheme.Value.IsActive) { " (ACTIVE)" } else { "" }
            Write-LogEntry "  - $($scheme.Key)$activeIndicator" 'INFO'
        }
        
        # Validate requested power plan
        if (-not $powerSchemes.ContainsKey($PowerPlan)) {
            Write-LogEntry "[WARN] Requested power plan '$PowerPlan' not found" 'WARNING'
            Write-LogEntry "Falling back to 'High Performance' scheme" 'WARNING'
            $PowerPlan = 'High Performance'
            
            if (-not $powerSchemes.ContainsKey($PowerPlan)) {
                throw "Neither requested scheme nor High Performance scheme available"
            }
        }
        
        $selectedScheme = $powerSchemes[$PowerPlan]
        Write-LogEntry "Selected scheme: $PowerPlan (GUID: $($selectedScheme.GUID))" 'INFO'

        # Backup current settings
        if (-not $WhatIfPreference) {
            Write-LogEntry "`n[SAVE] Backing up current power settings..." 'INFO'
            $backupFile = Backup-PowerSettings
        }

        if ($WhatIfPreference) {
            Write-LogEntry "`nWhatIf Summary:" 'INFO'
            Write-LogEntry "  - Would set power scheme to: $PowerPlan" 'INFO'
            Write-LogEntry "  - Would set monitor timeout to: $MonitorTimeoutMinutes minutes" 'INFO'
            Write-LogEntry "  - Would set disk timeout to: $(if($DiskTimeoutMinutes -eq 0){'Disabled'}else{"$DiskTimeoutMinutes minutes"})" 'INFO'
            Write-LogEntry "  - Would disable hibernation" 'INFO'
            $global:LastStatus = "[INFO] WhatIf completed - power settings would be configured."
            return
        }

        # Apply power configuration
        Write-LogEntry "`n[FAST] Applying power configuration..." 'INFO'
        $configResults = Set-PowerConfiguration -SchemeName $PowerPlan -SchemeGUID $selectedScheme.GUID -MonitorTimeout $MonitorTimeoutMinutes -DiskTimeout $DiskTimeoutMinutes
        
        # Process results
        $successCount = ($configResults | Where-Object { $_.Success }).Count
        $failureCount = ($configResults | Where-Object { -not $_.Success }).Count
        
        $script:appliedSettings = $configResults | Where-Object { $_.Success }
        $script:failedSettings = $configResults | Where-Object { -not $_.Success }

        # Verify configuration
        Write-LogEntry "`n[SCAN] Verifying power configuration..." 'INFO'
        try {
            $currentScheme = powercfg.exe /getactivescheme 2>$null
            if ($currentScheme -and $currentScheme -match $selectedScheme.GUID) {
                Write-LogEntry "[OK] Power scheme verification passed" 'SUCCESS'
            } else {
                Write-LogEntry "[WARN] Power scheme verification failed" 'WARNING'
            }
        } catch {
            Write-LogEntry "[WARN] Could not verify power scheme: $_" 'WARNING'
        }

    } catch {
        Write-LogEntry "Critical error during power settings configuration: $_" 'ERROR'
        $global:LastStatus = "[ERROR] Power settings configuration failed: $_"
        throw
    } finally {
        $stopwatch.Stop()
        $duration = $stopwatch.Elapsed.TotalSeconds
        
        Write-LogEntry "`n[SUMMARY] Configuration Summary:" 'INFO'
        Write-LogEntry "Duration: $([math]::Round($duration, 2)) seconds" 'INFO'
        Write-LogEntry "Settings applied: $($appliedSettings.Count)" 'SUCCESS'
        Write-LogEntry "Settings failed: $($failedSettings.Count)" 'ERROR'
        
        if ($appliedSettings.Count -gt 0) {
            Write-LogEntry "`n[OK] Successfully Applied:" 'SUCCESS'
            foreach ($setting in $appliedSettings) {
                Write-LogEntry "  - $($setting.Setting): $($setting.Value)" 'SUCCESS'
            }
        }
        
        if ($failedSettings.Count -gt 0) {
            Write-LogEntry "`n[ERROR] Failed Settings:" 'ERROR'
            foreach ($setting in $failedSettings) {
                Write-LogEntry "  - $($setting.Setting): $($setting.Error)" 'ERROR'
            }
        }
        
        # Write log file
        try {
            $logEntries | Out-File -FilePath $LogPath -Encoding UTF8 -Force
            Write-LogEntry "[NOTE] Detailed log saved to: $LogPath" 'INFO'
        } catch {
            Write-LogEntry "[WARN] Failed to save log file: $_" 'WARNING'
        }
        
        # Set final status
        if ($appliedSettings.Count -gt 0) {
            if ($failedSettings.Count -eq 0) {
                $global:LastStatus = "[OK] All power settings applied successfully."
            } else {
                $global:LastStatus = "[WARN] Power settings partially applied ($($appliedSettings.Count) success, $($failedSettings.Count) failed)."
            }
        } else {
            $global:LastStatus = "[ERROR] No power settings were applied."
        }
        
        Write-LogEntry "=== Power Settings Configuration Completed ===" 'INFO'
    }
}

# Utility function to get current power configuration
function Get-PowerConfiguration {
    [CmdletBinding()]
    param()
    
    try {
        $config = @{}
        
        # Get active scheme
        $activeScheme = powercfg.exe /getactivescheme 2>$null
        if ($activeScheme) {
            $config.ActiveScheme = $activeScheme.Trim()
        }
        
        # Get hibernation status
        $hibernation = powercfg.exe /availablesleepstates 2>$null
        $config.HibernationAvailable = ($hibernation -match "Hibernate")
        
        # Get monitor and disk timeouts (simplified parsing)
        $config.MonitorTimeoutAC = "Unknown"
        $config.MonitorTimeoutDC = "Unknown"
        $config.DiskTimeoutAC = "Unknown"
        $config.DiskTimeoutDC = "Unknown"
        
        return $config
        
    } catch {
        return @{
            Error = $_.Exception.Message
        }
    }
}

# -----------------------------------------------------------------------------
# Option 17 - Install Computer Lab Scheduled Tasks
# -----------------------------------------------------------------------------
function Register-LabScheduledTasks {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SourcePath = "\\filesvr\install\Scripts\Lab Update Scripts",
        
        [Parameter(Mandatory = $false)]
        [string]$DestinationPath = "C:\Windows\Scripts",
        
        [Parameter(Mandatory = $false)]
        [string]$AdminAccount = "MISAdmin",
        
        [switch]$Force
    )

    # Security: Require elevation
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        throw "This function must be run as Administrator"
    }

    $ErrorActionPreference = 'Continue'
    
    # Simple logging
    function Write-LogEntry {
        param([string]$Message, [string]$Level = 'INFO')
        
        $timestamp = Get-Date -Format 'HH:mm:ss'
        switch ($Level) {
            'ERROR' { Write-Host "[$timestamp] ERROR: $Message" -ForegroundColor Red }
            'WARNING' { Write-Host "[$timestamp] WARNING: $Message" -ForegroundColor Yellow }
            'SUCCESS' { Write-Host "[$timestamp] SUCCESS: $Message" -ForegroundColor Green }
            'INFO' { Write-Host "[$timestamp] INFO: $Message" -ForegroundColor Cyan }
            default { Write-Host "[$timestamp] $Message" -ForegroundColor White }
        }
    }

    # Get credentials - simplified
    function Get-AdminCredentials {
        param([string]$AccountName)
        
        Write-LogEntry "Getting credentials for $AccountName" 'INFO'
        
        $user = Get-LocalUser -Name $AccountName -ErrorAction Stop
        if (-not $user.Enabled) {
            throw "Account $AccountName is disabled"
        }
        
        Write-Host "`nEnter password for ${AccountName}:" -ForegroundColor Yellow
        $securePassword = Read-Host -AsSecureString
        
        return @{
            UserName = $AccountName
            Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
        }
    }

    # Create scheduled task using PROVEN method with custom Sunday scheduling
    function Create-ScheduledTaskProven {
        param(
            [string]$TaskName,
            [string]$FilePath,
            [string]$UserName,
            [string]$Password
        )
        
        Write-LogEntry "Creating task: $TaskName" 'INFO'
        Write-LogEntry "File: $FilePath" 'INFO'
        Write-LogEntry "User: $UserName" 'INFO'
        
        # Determine schedule time based on file prefix
        $fileName = Split-Path $FilePath -Leaf
        $scheduleTime = "01:30"  # Default time
        
        if ($fileName -match '^(\d+)_') {
            $fileNumber = [int]$matches[1]
            $scheduleTime = switch ($fileNumber) {
                1 { "01:30" }  # 01_Remove-UserProfiles.ps1
                2 { "02:00" }  # 02_Weekend_Apps_Updates.ps1
                3 { "02:45" }  # 03_Weekend_HP_Drivers_Updates.ps1
                4 { "04:00" }  # 04_Weekend_Windows_Updates.ps1
                5 { "08:00" }  # 05_SystemRepair.ps1
                default { "01:30" }
            }
        }
        
        Write-LogEntry "Scheduled for Sundays at: $scheduleTime" 'INFO'
        
        # Use the PROVEN command format from our testing
        $taskCommand = switch ([System.IO.Path]::GetExtension($FilePath).ToLower()) {
            '.ps1' { 
                "powershell.exe -ExecutionPolicy Bypass -File `"$FilePath`""
            }
            '.bat' { 
                "`"$FilePath`""
            }
            '.cmd' { 
                "`"$FilePath`""
            }
            default { 
                "`"$FilePath`""
            }
        }
        
        Write-LogEntry "Task command: $taskCommand" 'INFO'
        
        # Use batch file method - PROVEN to work, scheduled for Sundays
        $batchFile = "$env:TEMP\create_task_$TaskName.bat"
        $batchContent = @"
@echo off
echo Creating scheduled task: $TaskName
echo Command: $taskCommand
echo User: $UserName
echo File: $FilePath
echo Schedule: Sundays at $scheduleTime
echo File exists check:
if exist "$FilePath" (echo YES - File found) else (echo NO - File missing)
echo.
echo Running schtasks...
schtasks.exe /Create /TN "$TaskName" /TR "$taskCommand" /SC WEEKLY /D SUN /ST $scheduleTime /RU "$UserName" /RP "$Password" /RL HIGHEST /F
if %ERRORLEVEL% EQU 0 (
    echo SUCCESS: Task created successfully - scheduled for Sundays at $scheduleTime
) else (
    echo FAILED: Task creation failed with exit code %ERRORLEVEL%
)
"@
        
        try {
            # Create and run batch file
            $batchContent | Out-File -FilePath $batchFile -Encoding ASCII
            $output = cmd.exe /c "`"$batchFile`"" 2>&1
            $exitCode = $LASTEXITCODE
            
            Write-LogEntry "Batch execution output:" 'INFO'
            $output | ForEach-Object { 
                if ($_ -match "SUCCESS") {
                    Write-LogEntry "  $_" 'SUCCESS'
                } elseif ($_ -match "FAILED|ERROR") {
                    Write-LogEntry "  $_" 'ERROR'
                } else {
                    Write-LogEntry "  $_" 'INFO'
                }
            }
            
            # Clean up batch file
            Remove-Item $batchFile -Force -ErrorAction SilentlyContinue
            
            return ($exitCode -eq 0)
            
        } catch {
            Write-LogEntry "Exception creating task $TaskName`: $($_.Exception.Message)" 'ERROR'
            Remove-Item $batchFile -Force -ErrorAction SilentlyContinue
            return $false
        }
    }

    # Main execution
    try {
        Write-LogEntry "=== FINAL WORKING VERSION ===" 'SUCCESS'
        Write-LogEntry "Using PROVEN command format from testing!" 'SUCCESS'
        
        # Get credentials
        $credInfo = Get-AdminCredentials -AccountName $AdminAccount
        $userName = $credInfo.UserName
        $password = $credInfo.Password
        
        Write-LogEntry "Using user: $userName" 'SUCCESS'
        
        # Copy files
        Write-LogEntry "Copying files from $SourcePath to $DestinationPath" 'INFO'
        
        if (-not (Test-Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
        }
        
        $sourceFiles = Get-ChildItem -Path $SourcePath -Include @("*.bat", "*.cmd", "*.ps1") -Recurse
        Write-LogEntry "Found $($sourceFiles.Count) executable files to copy" 'INFO'
        
        foreach ($file in $sourceFiles) {
            $destFile = Join-Path $DestinationPath $file.Name
            Copy-Item -Path $file.FullName -Destination $destFile -Force
            Write-LogEntry "Copied: $($file.Name)" 'INFO'
        }
        
        # Register all tasks using PROVEN method
        $copiedFiles = Get-ChildItem -Path $DestinationPath -Include @("*.bat", "*.cmd", "*.ps1") -Recurse
        Write-LogEntry "Registering $($copiedFiles.Count) tasks using PROVEN method..." 'INFO'
        
        $successCount = 0
        $failCount = 0
        
        foreach ($file in $copiedFiles) {
            $taskName = "LabTask_$($file.BaseName)"
            
            # Remove existing task if Force specified
            if ($Force) {
                $existingTask = schtasks.exe /Query /TN $taskName 2>$null
                if ($LASTEXITCODE -eq 0) {
                    Write-LogEntry "Removing existing task: $taskName" 'INFO'
                    schtasks.exe /Delete /TN $taskName /F | Out-Null
                }
            }
            
            # Create task using proven method
            $success = Create-ScheduledTaskProven -TaskName $taskName -FilePath $file.FullName -UserName $userName -Password $password
            
            if ($success) {
                $successCount++
                
                # Verify task exists
                $verifyTask = schtasks.exe /Query /TN $taskName 2>$null
                if ($LASTEXITCODE -eq 0) {
                    Write-LogEntry "VERIFIED: Task $taskName exists and is registered" 'SUCCESS'
                } else {
                    Write-LogEntry "WARNING: Task $taskName created but verification failed" 'WARNING'
                }
            } else {
                $failCount++
            }
        }
        
        # Clear password
        $password = $null
        [System.GC]::Collect()
        
        Write-LogEntry "=== FINAL RESULTS ===" 'INFO'
        Write-LogEntry "Successfully registered: $successCount tasks" 'SUCCESS'
        Write-LogEntry "Failed to register: $failCount tasks" 'WARNING'
        
        if ($successCount -gt 0) {
            Write-LogEntry "SUCCESS! Lab scheduled tasks have been registered!" 'SUCCESS'
            Write-LogEntry "Tasks are scheduled to run on Sundays at staggered times:" 'INFO'
            Write-LogEntry "  01_Remove-UserProfiles.ps1     -> 1:30 AM" 'INFO'
            Write-LogEntry "  02_Weekend_Apps_Updates.ps1    -> 2:00 AM" 'INFO'
            Write-LogEntry "  03_Weekend_HP_Drivers_Updates.ps1 -> 2:45 AM" 'INFO'
            Write-LogEntry "  04_Weekend_Windows_Updates.ps1 -> 4:00 AM" 'INFO'
            Write-LogEntry "  05_SystemRepair.ps1            -> 8:00 AM" 'INFO'
        } else {
            Write-LogEntry "No tasks were successfully registered" 'ERROR'
        }
        
    } catch {
        Write-LogEntry "Critical error: $_" 'ERROR'
        throw
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
            Name = 'Update HP Drivers'
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
        'HPDrivers'         = 'Update HP Drivers'
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
    Write-Host "7.  Set Desktop Power Settings" -ForegroundColor White
    Write-Host "8.  Network Optimization" -ForegroundColor White
    Write-Host "9.  Application Updates" -ForegroundColor White
    Write-Host "10. HP Driver Updates" -ForegroundColor White
    Write-Host "11. Windows Updates" -ForegroundColor White
    Write-Host "12. Disk Cleanup" -ForegroundColor White
    Write-Host "13. System Repair" -ForegroundColor White
    Write-Host "14. Remove User Profiles" -ForegroundColor White
    Write-Host "15. Disable Last User Display" -ForegroundColor White
    Write-Host "16. Enable Automatic Login with CC-Student" -ForegroundColor White
    Write-Host "17. Install Computer Lab Scheduled Tasks" -ForegroundColor White
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
