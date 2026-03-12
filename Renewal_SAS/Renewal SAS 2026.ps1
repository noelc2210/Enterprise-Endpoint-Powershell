<#
.SYNOPSIS
    SAS 9.4 Teaching & Research License Renewal for Intune Deployment
.DESCRIPTION
    Renews SAS 9.4 license using packaged license file and response file.
    Designed for silent Intune deployment with robust logging and user notification.
.NOTES
    Author:  github.com/noelc2210
    Last Modified: February 2026
    Version: 1.1

    Documentation: See README.md in this repository
#>

[CmdletBinding()]
param(
    [switch]$ForceRenewal
)

#Requires -RunAsAdministrator

# ############################################################################
#    UPDATE THIS VALUE EACH YEAR - This is the ONLY change needed annually
# ############################################################################

$LicenseYear = "2026"

# ############################################################################
#    DO NOT MODIFY BELOW THIS LINE
# ############################################################################

<#
    Intune Configuration:
    - Install:    powershell.exe -ExecutionPolicy Bypass -File SAS_Renewal_2026.ps1
    - Uninstall:  (not applicable)
    - Detection:  Use SAS_Intune_Detection_2026.ps1
    - Context:    System
    - 64-bit:     Yes

    Exit Codes:
    - 0  = Success (renewal complete or license already current)
    - 1  = Failure (something broke - check logs)
    - 2  = Deferred (no active user session - Intune should retry)

    Windows Event Log:
    - Source:     SAS-Renewal
    - Log:        Application
    - Query:      Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='SAS-Renewal'}
    - Event IDs:
        1000 = Script started
        1001 = Script completed successfully
        1002 = Script failed
        1003 = Script deferred (no active session)
        1004 = Logging initialization failed
        1005 = Network logging unavailable
        1006 = Unhandled exception
#>


# ============================================================================
# CONFIGURATION - All paths and settings derive from $LicenseYear above
# ============================================================================
$Config = @{
    # License settings (built from $LicenseYear)
    LicenseYear        = $LicenseYear
    LicenseFileName    = "SAS_94_TR_${LicenseYear}_License.txt"
    ResponseFileName   = "SAS_Renewal_${LicenseYear}.properties"
    LicenseDestination = "C:\ProgramData\...\SAS"

    # SAS paths
    SASHome            = "C:\Program Files\SASHome"
    SASFoundation      = "C:\Program Files\SASHome\SASFoundation\9.4"
    SASDM              = "C:\Program Files\SASHome\SASDeploymentManager\9.4\sasdm.exe"

    # Logging
    LogsRootNetwork    = "\\...\Share\IT\...\Logs\SAS_Renewal_${LicenseYear}_Logs"
    LogsRootLocal      = "C:\Intune_Logs\SAS_Renewal_${LicenseYear}_Logs"
    LogRetentionDays   = 30

    # Process handling
    ProcessWaitSeconds = 30
    ProcessNames       = @("sas", "sasmc", "sasdm", "SASWindow", "SASFoundation", "java")

    # Registry detection key for Intune
    RegistryPath       = "HKLM:\Software\...\SAS"
    RegistryValueName  = "LicenseYear"

    # Validation - text expected in valid license file
    LicenseMarker      = "EXPIRE"

    # Windows Event Log
    EventLogSource     = "SAS-Renewal"
    EventLogName       = "Application"
}

# ============================================================================
# FORCE 64-BIT POWERSHELL
# ============================================================================
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Write-Host "Relaunching script in 64-bit PowerShell..."
    if ($myInvocation.Line) {
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File $myInvocation.MyCommand.Path @PSBoundParameters
    } else {
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -ExecutionPolicy Bypass -NoProfile -File $myInvocation.MyCommand.Definition @PSBoundParameters
    }
    exit $LASTEXITCODE
}

# ============================================================================
# SCRIPT VARIABLES - Do not edit
# ============================================================================
$Script:ExitCode       = 0
$Script:Timestamp      = (Get-Date).ToString('yyyyMMdd_HHmmss')
$Script:DeviceName     = $env:COMPUTERNAME
$Script:LogFileName    = "${Script:DeviceName}_SAS_Renewal_$Script:Timestamp.log"
$Script:LogFileNetwork = Join-Path (Join-Path $Config.LogsRootNetwork $Script:DeviceName) $Script:LogFileName
$Script:LogFileLocal   = Join-Path $Config.LogsRootLocal $Script:LogFileName
$Script:LicenseSource      = Join-Path $PSScriptRoot $Config.LicenseFileName
$Script:LicenseDestPath    = Join-Path $Config.LicenseDestination $Config.LicenseFileName
$Script:ResponseSource     = Join-Path $PSScriptRoot $Config.ResponseFileName
$Script:ResponseDestPath   = Join-Path $Config.LicenseDestination $Config.ResponseFileName
$Script:NetworkLoggingAvailable = $false
$Script:ActiveSessionId        = $null

# Use Continue globally - terminating behavior applied per-call where needed
$ErrorActionPreference = 'Continue'

# ============================================================================
# LOGGING FUNCTIONS
# ============================================================================
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $prefix = switch ($Level) {
        'INFO'    { "[INFO]   " }
        'WARN'    { "[WARN]   " }
        'ERROR'   { "[ERROR]  " }
        'SUCCESS' { "[SUCCESS]" }
    }

    $line = "$timestamp $prefix $Message"

    # Write to console (captured by transcript if active)
    Write-Host $line

    # ALSO write directly to local log file so content is never lost
    # This ensures the log has content even if transcript fails or is interrupted
    if ($Script:LogFileLocal) {
        try {
            Add-Content -Path $Script:LogFileLocal -Value $line -ErrorAction SilentlyContinue
        }
        catch { }
    }
}

# ============================================================================
# WINDOWS EVENT LOG FUNCTION
# ============================================================================
function Write-SASEvent {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [int]$EventId = 1000,

        [ValidateSet('Information', 'Warning', 'Error')]
        [string]$EntryType = 'Information'
    )

    try {
        # Register event source if it doesn't exist (requires admin/SYSTEM - always true here)
        if (-not [System.Diagnostics.EventLog]::SourceExists($Config.EventLogSource)) {
            [System.Diagnostics.EventLog]::CreateEventSource($Config.EventLogSource, $Config.EventLogName)
        }

        $fullMessage = @"
Device:       $Script:DeviceName
User Context: $env:USERNAME
License Year: $LicenseYear
Log File:     $Script:LogFileLocal

$Message
"@
        Write-EventLog -LogName $Config.EventLogName `
            -Source $Config.EventLogSource `
            -EventId $EventId `
            -EntryType $EntryType `
            -Message $fullMessage `
            -ErrorAction Stop
    }
    catch {
        # Event log failure is non-fatal - regular logging is still working
        Write-Log "Could not write to Windows Event Log: $($_.Exception.Message)" -Level WARN
    }
}

# ============================================================================
# LOGGING INITIALIZATION
# ============================================================================
function Initialize-Logging {

    # Step 1: Create local log directory
    try {
        if (-not (Test-Path $Config.LogsRootLocal)) {
            New-Item -ItemType Directory -Path $Config.LogsRootLocal -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Host "FATAL: Could not create local log directory: $($_.Exception.Message)"
        Write-SASEvent -Message "FATAL: Could not create local log directory: $($_.Exception.Message)" -EventId 1004 -EntryType Error
        return $false
    }

    # Step 2: Write log header IMMEDIATELY via direct file write
    # This happens before transcript starts so the header is never lost
    $header = @"
================================================================================
SAS 9.4 License Renewal - Version 1.1
Script Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Device:         $($env:COMPUTERNAME)
User Context:   $($env:USERNAME)
License Year:   $LicenseYear
Force Renewal:  $ForceRenewal
Script Path:    $PSScriptRoot
================================================================================
"@
    try {
        Add-Content -Path $Script:LogFileLocal -Value $header -ErrorAction Stop
    }
    catch {
        Write-Host "FATAL: Could not write to local log file: $($_.Exception.Message)"
        Write-SASEvent -Message "FATAL: Could not write to local log file: $($_.Exception.Message)" -EventId 1004 -EntryType Error
        return $false
    }

    # Step 3: Start transcript (adds extra detail like PS version, host info)
    try {
        Start-Transcript -Path $Script:LogFileLocal -Append -Force -ErrorAction Stop | Out-Null
        Write-Log "Transcript started - appending to direct log"
    }
    catch {
        # Transcript failure is non-fatal - direct file logging is already working
        Write-Log "Transcript could not start (non-fatal): $($_.Exception.Message)" -Level WARN
    }

    # Step 4: Test network logging availability
    try {
        $networkDeviceFolder = Join-Path $Config.LogsRootNetwork $Script:DeviceName
        if (-not (Test-Path $networkDeviceFolder)) {
            New-Item -ItemType Directory -Path $networkDeviceFolder -Force -ErrorAction Stop | Out-Null
            Write-Log "Created network device subfolder: $networkDeviceFolder"
        }

        # Test write access with a small temp file
        $testFile = Join-Path $networkDeviceFolder "write_test_$Script:Timestamp.tmp"
        [System.IO.File]::WriteAllText($testFile, "test") 
        Remove-Item $testFile -Force -ErrorAction SilentlyContinue

        $Script:NetworkLoggingAvailable = $true
        Write-Log "Network logging available: $networkDeviceFolder" -Level SUCCESS
    }
    catch {
        $Script:NetworkLoggingAvailable = $false
        Write-Log "Network logging unavailable (will log locally only): $($_.Exception.Message)" -Level WARN
        Write-SASEvent -Message "Network logging unavailable: $($_.Exception.Message)`nLocal log: $Script:LogFileLocal" -EventId 1005 -EntryType Warning
    }

    return $true
}

function Remove-OldLogs {
    $cutoffDate = (Get-Date).AddDays(-$Config.LogRetentionDays)

    # Clean local logs
    Get-ChildItem -Path $Config.LogsRootLocal -Filter "*.log" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoffDate } |
        ForEach-Object {
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            Write-Log "Removed old local log: $($_.Name)"
        }

    # Clean network logs (device subfolder only)
    if ($Script:NetworkLoggingAvailable) {
        try {
            $networkDeviceFolder = Join-Path $Config.LogsRootNetwork $Script:DeviceName
            if (Test-Path $networkDeviceFolder) {
                Get-ChildItem -Path $networkDeviceFolder -Filter "*.log" -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -lt $cutoffDate } |
                    ForEach-Object {
                        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
                        Write-Log "Removed old network log: $($_.Name)"
                    }
            }
        }
        catch {
            Write-Log "Could not clean network logs: $($_.Exception.Message)" -Level WARN
        }
    }
}

function Copy-LogToNetwork {
    if (-not $Script:NetworkLoggingAvailable) {
        Write-Log "Network logging unavailable - skipping network copy" -Level WARN
        return
    }

    try {
        # Stop transcript to finalize local log before copying
        try { Stop-Transcript -ErrorAction SilentlyContinue | Out-Null } catch { }
        Start-Sleep -Seconds 2

        if (-not (Test-Path $Script:LogFileLocal)) {
            Write-Log "Local log file not found - cannot copy to network" -Level WARN
            return
        }

        $localLogSize = (Get-Item $Script:LogFileLocal).Length
        if ($localLogSize -eq 0) {
            Write-Log "Local log file is empty - skipping network copy" -Level WARN
            return
        }

        Copy-Item -Path $Script:LogFileLocal -Destination $Script:LogFileNetwork -Force -ErrorAction Stop
        $networkLogSize = (Get-Item $Script:LogFileNetwork -ErrorAction SilentlyContinue).Length
        Write-Log "Log copied to network: $networkLogSize bytes -> $Script:LogFileNetwork" -Level SUCCESS
    }
    catch {
        Write-Log "Could not copy log to network: $($_.Exception.Message)" -Level WARN
    }
}

# ============================================================================
# USER SESSION FUNCTIONS
# ============================================================================
function Test-ActiveUserSession {
    try {
        # query user can throw NativeCommandError from SYSTEM context even with 2>$null
        # Capture both stdout and stderr cleanly using cmd.exe to avoid PS error stream pollution
        $queryResult = $null
        try {
            $queryResult = cmd /c "query user 2>nul" 2>$null
        } catch { }

        # Fallback: query session is more reliable for RDP sessions from SYSTEM context
        if (-not $queryResult -or $queryResult -notmatch "Active") {
            Write-Log "query user returned no active sessions - trying query session fallback"
            try {
                $queryResult = cmd /c "query session 2>nul" 2>$null
            } catch { }
        }

        if (-not $queryResult) {
            Write-Log "No user sessions found (both query user and query session returned nothing)" -Level WARN
            return $false
        }

        Write-Log "Raw query output: $($queryResult -join ' | ')"

        # Find all Active sessions and log if more than one found
        $activeSessions = @($queryResult | Where-Object { $_ -match "Active" })
        if (-not $activeSessions -or $activeSessions.Count -eq 0) {
            Write-Log "No active sessions found (all disconnected or at lock screen)" -Level WARN
            return $false
        }

        if ($activeSessions.Count -gt 1) {
            Write-Log "Multiple active sessions detected ($($activeSessions.Count)) - targeting first session" -Level WARN
            foreach ($s in $activeSessions) { Write-Log "  Session: $($s.Trim())" }
        }

        $activeSession = $activeSessions[0]

        # Parse session ID by column position
        # query user/session output format: USERNAME SESSIONNAME ID STATE IDLE LOGON
        # Username field may be prefixed with '>' for current user - strip it
        # We find the ID by locating the STATE column (Active/Disc) and taking the token before it
        $sessionId = $null
        $tokens = $activeSession.Trim() -replace '^>', '' -split '\s+'
        for ($i = 0; $i -lt $tokens.Count; $i++) {
            if ($tokens[$i] -match '^(Active|Disc)$' -and $i -gt 0) {
                if ($tokens[$i - 1] -match '^\d+$') {
                    $sessionId = $tokens[$i - 1]
                }
                break
            }
        }

        if ($sessionId) {
            $Script:ActiveSessionId = $sessionId
            Write-Log "Active session ID: $Script:ActiveSessionId (session line: $($activeSession.Trim()))"
        } else {
            $Script:ActiveSessionId = $null
            Write-Log "Could not parse session ID from query output - will fall back to broadcast" -Level WARN
        }

        # Verify explorer.exe is running (confirms interactive desktop, not just a session)
        $explorerProcess = Get-Process -Name "explorer" -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -gt 0 }
        if (-not $explorerProcess) {
            Write-Log "Explorer not running - no interactive desktop detected" -Level WARN
            return $false
        }

        Write-Log "Active user session confirmed" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Could not determine session status: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

# ============================================================================
# NOTIFICATION FUNCTIONS
# ============================================================================
function Show-MessageBox {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [int]$Timeout = 0
    )

    $shortMessage = $Message.Substring(0, [Math]::Min(60, $Message.Length))
    Write-Log "Showing message to user: $shortMessage..."

    # Target specific session ID if available, fall back to broadcast
    $msgTarget = if ($Script:ActiveSessionId) { $Script:ActiveSessionId } else { "*" }
    Write-Log "msg.exe target: $msgTarget"

    # Detect if running as 32-bit and use Sysnative to bypass WOW64 redirection
    $is32bit = [System.IntPtr]::Size -eq 4
    $msgExePath = if ($is32bit -and (Test-Path "$env:windir\Sysnative\msg.exe")) {
        "$env:windir\Sysnative\msg.exe"
    } else {
        "$env:windir\System32\msg.exe"
    }
    Write-Log "msg.exe path: $msgExePath (32-bit context: $is32bit)"

    try {
        if ($Timeout -gt 0) {
            $result = & $msgExePath $msgTarget /TIME:$Timeout $Message 2>&1
        } else {
            $result = & $msgExePath $msgTarget $Message 2>&1
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Message displayed via msg.exe (session: $msgTarget)" -Level SUCCESS
            return $true
        } else {
            Write-Log "msg.exe failed (exit code: $LASTEXITCODE, target: $msgTarget) - Output: $result" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "msg.exe exception: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================
function Test-SASInstallation {
    Write-Log "Checking SAS installation..."

    $sasExe = Join-Path $Config.SASFoundation "sas.exe"

    if (-not (Test-Path $sasExe)) {
        Write-Log "SAS executable not found at: $sasExe" -Level ERROR
        return $false
    }

    if (-not (Test-Path $Config.SASDM)) {
        Write-Log "SAS Deployment Manager not found at: $($Config.SASDM)" -Level ERROR
        return $false
    }

    Write-Log "SAS installation verified at: $($Config.SASFoundation)" -Level SUCCESS
    return $true
}

function Test-PackageFiles {
    Write-Log "Checking for packaged files..."

    Write-Log "License source: $Script:LicenseSource"
    if (-not (Test-Path $Script:LicenseSource)) {
        Write-Log "License file not found in package: $Script:LicenseSource" -Level ERROR
        return $false
    }
    Write-Log "License file found" -Level SUCCESS

    Write-Log "Response file source: $Script:ResponseSource"
    if (-not (Test-Path $Script:ResponseSource)) {
        Write-Log "Response file not found in package: $Script:ResponseSource" -Level ERROR
        return $false
    }
    Write-Log "Response file found" -Level SUCCESS

    return $true
}

function Test-LicenseFileValid {
    param([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        return $false
    }

    $content = Get-Content $FilePath -Raw -ErrorAction SilentlyContinue

    if ([string]::IsNullOrWhiteSpace($content)) {
        Write-Log "License file is empty: $FilePath" -Level ERROR
        return $false
    }

    if ($content -notmatch $Config.LicenseMarker) {
        Write-Log "License file missing expected marker '$($Config.LicenseMarker)'" -Level WARN
        return $false
    }

    if ($content -notmatch $Config.LicenseYear) {
        Write-Log "License file missing expected year '$($Config.LicenseYear)'" -Level WARN
        return $false
    }

    return $true
}

function Get-CurrentLicenseStatus {
    Write-Log "Checking current license status..."

    $status = @{
        RegistryYear  = $null
        IsCurrentYear = $false
        NeedsRenewal  = $true
    }

    try {
        if (Test-Path $Config.RegistryPath) {
            $regValue = Get-ItemProperty -Path $Config.RegistryPath -Name $Config.RegistryValueName -ErrorAction SilentlyContinue
            if ($regValue) {
                $status.RegistryYear = $regValue.$($Config.RegistryValueName)
                Write-Log "Registry shows license year: $($status.RegistryYear)"

                if ($status.RegistryYear -eq $Config.LicenseYear) {
                    $status.IsCurrentYear = $true
                    $status.NeedsRenewal  = $false
                    Write-Log "License is already current ($($Config.LicenseYear))" -Level SUCCESS
                }
            } else {
                Write-Log "Registry path exists but LicenseYear value not found"
            }
        } else {
            Write-Log "Registry path does not exist: $($Config.RegistryPath)"
        }
    }
    catch {
        Write-Log "Could not read registry: $($_.Exception.Message)" -Level WARN
    }

    return $status
}

# ============================================================================
# PROCESS MANAGEMENT
# ============================================================================
function Get-RunningSASProcesses {
    $processes = @()

    foreach ($name in $Config.ProcessNames) {
        $found = Get-Process -Name $name -ErrorAction SilentlyContinue
        if ($found) { $processes += $found }
    }

    $sasdmProcs = Get-Process -Name "sasdm" -ErrorAction SilentlyContinue
    if ($sasdmProcs) {
        foreach ($dm in $sasdmProcs) {
            try {
                $children = Get-CimInstance -ClassName Win32_Process -Filter "ParentProcessId=$($dm.Id)" -ErrorAction SilentlyContinue
                foreach ($child in $children) {
                    if ($child.Name -eq "java.exe") {
                        $javaProc = Get-Process -Id $child.ProcessId -ErrorAction SilentlyContinue
                        if ($javaProc) { $processes += $javaProc }
                    }
                }
            }
            catch { }
        }
    }

    return $processes | Sort-Object -Property Id -Unique
}

function Stop-SASProcesses {
    Write-Log "Checking for running SAS processes..."

    $processes = Get-RunningSASProcesses

    if (-not $processes -or $processes.Count -eq 0) {
        Write-Log "No SAS processes running"
        return $true
    }

    Write-Log "Found $($processes.Count) SAS-related process(es):"
    foreach ($proc in $processes) {
        Write-Log "  - $($proc.Name) (PID: $($proc.Id))"
    }

    $warningMessage = @"
Your SAS license needs to be renewed. Any open SAS applications will be closed in 30 seconds or you may click OK now.

-IT Support
"@
    # Warning popup failure is non-fatal - SAS will still be closed and renewal will proceed
    # User just won't get advance notice
    if (-not (Show-MessageBox -Message $warningMessage -Timeout 30)) {
        Write-Log "Could not display SAS closing warning to user - proceeding with process termination anyway" -Level WARN
    }

    Write-Log "Terminating SAS processes..."

    foreach ($proc in $processes) {
        try {
            $proc.Refresh()
            if ($proc.HasExited) {
                Write-Log "Process $($proc.Name) (PID: $($proc.Id)) already exited"
                continue
            }

            if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
                Write-Log "Attempting graceful close: $($proc.Name)"
                $proc.CloseMainWindow() | Out-Null
                Start-Sleep -Seconds 3
                $proc.Refresh()
            }

            if (-not $proc.HasExited) {
                Write-Log "Force terminating: $($proc.Name) (PID: $($proc.Id))"
                Stop-Process -Id $proc.Id -Force -ErrorAction Stop
            }

            Write-Log "Terminated: $($proc.Name)" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to terminate $($proc.Name): $($_.Exception.Message)" -Level WARN
        }
    }

    Start-Sleep -Seconds 2
    $remaining = Get-RunningSASProcesses

    if ($remaining -and $remaining.Count -gt 0) {
        Write-Log "Some processes still running after termination attempt" -Level WARN
        return $false
    }

    Write-Log "All SAS processes terminated" -Level SUCCESS
    return $true
}

# ============================================================================
# LICENSE RENEWAL FUNCTIONS
# ============================================================================
function Copy-LicenseFile {
    Write-Log "Copying license file to destination..."

    try {
        if (-not (Test-Path $Config.LicenseDestination)) {
            New-Item -ItemType Directory -Path $Config.LicenseDestination -Force -ErrorAction Stop | Out-Null
            Write-Log "Created destination folder: $($Config.LicenseDestination)"
        }

        Copy-Item -Path $Script:LicenseSource -Destination $Script:LicenseDestPath -Force -ErrorAction Stop

        if (-not (Test-LicenseFileValid -FilePath $Script:LicenseDestPath)) {
            Write-Log "Copied license file failed validation" -Level ERROR
            return $false
        }

        Write-Log "License file copied to: $Script:LicenseDestPath" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Failed to copy license file: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Copy-ResponseFile {
    Write-Log "Copying response file to destination..."

    try {
        if (-not (Test-Path $Config.LicenseDestination)) {
            New-Item -ItemType Directory -Path $Config.LicenseDestination -Force -ErrorAction Stop | Out-Null
            Write-Log "Created destination folder: $($Config.LicenseDestination)"
        }

        Copy-Item -Path $Script:ResponseSource -Destination $Script:ResponseDestPath -Force -ErrorAction Stop

        Write-Log "Response file copied to: $Script:ResponseDestPath" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Failed to copy response file: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Invoke-SASRenewal {
    Write-Log "Starting SAS Deployment Manager with response file..."

    try {
        $process = Start-Process -FilePath $Config.SASDM `
            -ArgumentList "-quiet", "-responsefile", "`"$Script:ResponseDestPath`"" `
            -Wait -PassThru -NoNewWindow -ErrorAction Stop

        $exitCode = $process.ExitCode
        Write-Log "SASDM completed with exit code: $exitCode"

        if ($exitCode -eq 0) {
            Write-Log "License renewal completed successfully" -Level SUCCESS
            return $true
        } elseif ($exitCode -eq 7) {
            Write-Log "SASDM exit code 7 - license already up to date (treating as success)" -Level SUCCESS
            return $true
        } else {
            Write-Log "SASDM returned non-zero exit code: $exitCode" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "Failed to run SASDM: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Set-RegistryDetectionKey {
    Write-Log "Setting registry detection key..."

    try {
        if (-not (Test-Path $Config.RegistryPath)) {
            New-Item -Path $Config.RegistryPath -Force -ErrorAction Stop | Out-Null
        }

        Set-ItemProperty -Path $Config.RegistryPath -Name $Config.RegistryValueName -Value $Config.LicenseYear -Type String -Force -ErrorAction Stop

        $verify = Get-ItemProperty -Path $Config.RegistryPath -Name $Config.RegistryValueName -ErrorAction Stop
        if ($verify.$($Config.RegistryValueName) -eq $Config.LicenseYear) {
            Write-Log "Registry key set: $($Config.RegistryPath)\$($Config.RegistryValueName) = $($Config.LicenseYear)" -Level SUCCESS
            return $true
        } else {
            Write-Log "Registry key verification failed" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "Failed to set registry key: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================
function Start-Renewal {

    # Write Event ID 1000 - script started
    Write-SASEvent -Message "SAS 9.4 License Renewal script started.`nForce Renewal: $ForceRenewal" -EventId 1000 -EntryType Information

    # Step 1: Check for active user session
    # Exit code 2 = deferred (not a failure - Intune should retry)
    Write-Log "Checking user session status..."
    if (-not (Test-ActiveUserSession)) {
        Write-Log "No active user session - deferring renewal (exit code 2)" -Level WARN
        Write-SASEvent -Message "Renewal deferred - no active user session detected." -EventId 1003 -EntryType Warning
        return 2
    }
    Write-Log "Active user session detected - proceeding with renewal" -Level SUCCESS

    # Step 2: Show initial notification
    # If this fails the user has no idea what's happening - defer so Intune retries
    $initialMessage = @"
This computer is going to be checked for a 2026 SAS license and if not found, will follow steps to renew it.

-IT Support
"@
    if (-not (Show-MessageBox -Message $initialMessage)) {
        Write-Log "Initial notification failed to display - deferring so Intune retries (exit code 2)" -Level ERROR
        Write-SASEvent -Message "Renewal deferred - could not display initial notification to user. msg.exe failed." -EventId 1003 -EntryType Warning
        return 2
    }

    # Step 3: Validate SAS installation
    if (-not (Test-SASInstallation)) {
        Write-Log "SAS is not installed on this system" -Level ERROR
        return 1
    }

    # Step 4: Check current license status
    $licenseStatus = Get-CurrentLicenseStatus

    if ($licenseStatus.IsCurrentYear -and -not $ForceRenewal) {
        Write-Log "License is already current for $($Config.LicenseYear)" -Level SUCCESS

        $alreadyActiveMessage = @"
Your SAS 2026 license is currently active. Thank you for your time.

-IT Support
"@
        # Best-effort notification - license is already good so no retry needed
        if (-not (Show-MessageBox -Message $alreadyActiveMessage)) {
            Write-Log "Could not display already-active notification - continuing anyway (license is current)" -Level WARN
        }

        Write-Log "Ensuring registry detection key is set..."
        if (Set-RegistryDetectionKey) {
            Write-Log "No renewal needed - system is compliant" -Level SUCCESS
            return 0
        } else {
            Write-Log "Failed to set registry key" -Level ERROR
            return 1
        }
    }

    # Step 5: Verify packaged files
    if (-not (Test-PackageFiles)) {
        Write-Log "Required files not found in package - aborting" -Level ERROR
        return 1
    }

    # Step 6: Copy license file
    if (-not (Copy-LicenseFile)) {
        Write-Log "Failed to copy license file - aborting" -Level ERROR
        return 1
    }

    # Step 7: Copy response file
    if (-not (Copy-ResponseFile)) {
        Write-Log "Failed to copy response file - aborting" -Level ERROR
        return 1
    }

    # Step 8: Stop SAS processes (shows 30-second warning if SAS is running)
    $processesStopped = Stop-SASProcesses
    if (-not $processesStopped) {
        Write-Log "Could not stop all SAS processes - attempting renewal anyway" -Level WARN
    }

    # Step 9: Run renewal
    if (-not (Invoke-SASRenewal)) {
        Write-Log "License renewal failed" -Level ERROR
        Write-SASEvent -Message "SAS license renewal FAILED - SASDM returned non-zero exit code. Check local log: $Script:LogFileLocal" -EventId 1002 -EntryType Error

        $failureMessage = @"
SAS did not successfully renew. Please notify ithelp@yourdomain.edu with your computer name.

-IT Support
"@
        if (-not (Show-MessageBox -Message $failureMessage)) {
            Write-Log "Could not display failure notification to user - check logs and Event Log for details" -Level WARN
        }
        return 1
    }

    # Step 10: Set registry detection key
    if (-not (Set-RegistryDetectionKey)) {
        Write-Log "Renewal succeeded but registry key was not written - Intune will retry next check-in" -Level ERROR
        Write-SASEvent -Message "SAS renewal succeeded but registry detection key failed to write. Intune will attempt renewal again on next check-in." -EventId 1002 -EntryType Error
        # Exit 1 so Intune retries - the renewal itself worked but without the key
        # the detection script will keep triggering re-deployment
        if (-not (Show-MessageBox -Message "SAS was successfully renewed with the 2026 license.`n`n-IT Support")) {
            Write-Log "Could not display success notification to user" -Level WARN
        }
        return 1
    }

    Write-Log "License file stored at: $Script:LicenseDestPath"
    Write-Log "Response file stored at: $Script:ResponseDestPath"

    $successMessage = @"
SAS was successfully renewed with the 2026 license.

-IT Support
"@
    if (-not (Show-MessageBox -Message $successMessage)) {
        Write-Log "Could not display success notification to user - renewal still completed successfully" -Level WARN
    }

    Write-Log "=============================================="
    Write-Log "License renewal completed successfully" -Level SUCCESS
    Write-Log "=============================================="
    Write-SASEvent -Message "SAS 9.4 license renewal completed successfully." -EventId 1001 -EntryType Information

    return 0
}

# ============================================================================
# ENTRY POINT
# ============================================================================
try {
    if (-not (Initialize-Logging)) {
        # Last-ditch attempt: write directly to log file even though init failed
        $failMsg = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR]   FATAL: Logging initialization failed - script cannot continue"
        try {
            if ($Script:LogFileLocal) {
                New-Item -ItemType Directory -Path (Split-Path $Script:LogFileLocal) -Force -ErrorAction SilentlyContinue | Out-Null
                Add-Content -Path $Script:LogFileLocal -Value $failMsg -ErrorAction SilentlyContinue
            }
        } catch { }
        Write-Host "FATAL: Could not initialize logging - exiting"
        Write-SASEvent -Message "FATAL: Logging initialization failed. Script cannot continue." -EventId 1004 -EntryType Error
        exit 1
    }

    $Script:ExitCode = Start-Renewal
    Remove-OldLogs
    Copy-LogToNetwork
}
catch {
    Write-Log "Unhandled exception: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    Write-SASEvent -Message "Unhandled exception: $($_.Exception.Message)`nStack trace: $($_.ScriptStackTrace)`nLocal log: $Script:LogFileLocal" -EventId 1006 -EntryType Error
    $Script:ExitCode = 1
    Remove-OldLogs
    Copy-LogToNetwork
}
finally {
    Write-Log "Script exiting with code: $Script:ExitCode"

    # One final direct write to ensure exit code is captured
    if ($Script:LogFileLocal) {
        Add-Content -Path $Script:LogFileLocal -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO]    Script exited with code: $Script:ExitCode" -ErrorAction SilentlyContinue
    }

    exit $Script:ExitCode
}
