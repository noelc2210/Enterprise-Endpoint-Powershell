<#
.SYNOPSIS
    SAS 9.4 Teaching & Research Installation Script for Intune Deployment

.DESCRIPTION
    Silently installs SAS 9.4 M9 with 2026 license from network share.
    Downloads 14GB installer, installs, cleans up.
    Uses msg.exe for all user notifications (only method that works from SYSTEM).

.NOTES
    Author: github.com/noelc2210
    Date: March 2026
    Version: 2.3.16 - Added robocopy support for long paths when copying missing files
    
    Deployment: Available (not Required) - Users initiate from Company Portal
    
    ANNUAL UPDATE - CHANGE THIS YEAR VARIABLE:
#>

#Requires -RunAsAdministrator

# ============================================================================
# ANNUAL UPDATE - CHANGE THIS YEAR VARIABLE
# ============================================================================
# This is the ONLY variable you need to change each year (e.g., "2026" to "2027")
$LicenseYear = "2026"

# ============================================================================
# DO NOT MODIFY BELOW THIS LINE
# ============================================================================

# Script Version Identifier - Updated with each change
$Script:ScriptVersion = "2.3.16"  # Added robocopy for long paths
$Script:VersionDate = "2026-03-11"
$Script:FilesAlreadyExisted = $false  # Track if files were found vs copied

# ============================================================================
# CONFIGURATION - All settings derive from $LicenseYear
# ============================================================================
$Config = @{
    # License settings
    LicenseYear = $LicenseYear
    
    # Network paths
    NetworkSourcePath = "\\...\Share\IT\...\SAS_Installer_$LicenseYear\SAS94_Install_Files"
    
    # Local paths
    LocalTempPath = "C:\Temp\SAS_Install_$LicenseYear"
    SASInstallPath = "C:\Program Files\SASHome"
    SASFoundationPath = "C:\Program Files\SASHome\SASFoundation\9.4"
    
    # Logging
    LogsRoot = "C:\Intune_Logs\SAS_Installer_${LicenseYear}_Logs"
    LogRetentionDays = 30
    
    # Installation settings
    MinDiskSpaceGB = 20
    CopyTimeoutMinutes = 5
    InstallTimeoutMinutes = 180  # 3 hours
    MaxRetryAttempts = 3
    
    # Message timeouts
    StatusMessageTimeout = 20  # Status messages auto-close after 20 seconds
    
    # Registry detection
    RegistryPath = "HKLM:\Software\...\SAS"
    RegistryValueName = "LicenseYear"
    
    # Copy progress tracking
    CopyInProgressMarker = "C:\ProgramData\...\SAS\CopyInProgress_$LicenseYear.txt"
}

# ============================================================================
# SCRIPT VARIABLES
# ============================================================================
$Script:LogFile = Join-Path $Config.LogsRoot "$env:COMPUTERNAME`_SAS_Install_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$Script:RetryAttempt = 1
$Script:ActiveSessionId = $null  # Session ID for msg.exe targeting (set by Get-ActiveUserSession)

# ============================================================================
# LOGGING FUNCTIONS
# ============================================================================
function Initialize-Logging {
    if (-not (Test-Path $Config.LogsRoot)) {
        New-Item -ItemType Directory -Path $Config.LogsRoot -Force | Out-Null
    }
    Start-Transcript -Path $Script:LogFile -Force | Out-Null
}

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
    
    Write-Host "$timestamp $prefix $Message"
}

function Remove-OldLogs {
    $cutoffDate = (Get-Date).AddDays(-$Config.LogRetentionDays)
    
    Get-ChildItem -Path $Config.LogsRoot -Filter "*.log" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoffDate } |
        ForEach-Object {
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            Write-Log "Removed old log: $($_.Name)"
        }
}

function Copy-LogToNetwork {
    <#
    .SYNOPSIS
    Copies the current log file to network share for remote access
    #>
    try {
        $networkLogPath = "\\...\Share\IT\...\Logs\SAS_Install_2026_Logs\$env:COMPUTERNAME"
        
        # Create computer-specific folder if it doesn't exist
        if (-not (Test-Path $networkLogPath)) {
            New-Item -ItemType Directory -Path $networkLogPath -Force | Out-Null
            Write-Log "Created network log folder: $networkLogPath"
        }
        
        # Copy current log file to network
        if (Test-Path $Script:LogFile) {
            $logFileName = Split-Path $Script:LogFile -Leaf
            $networkLogFile = Join-Path $networkLogPath $logFileName
            Copy-Item -Path $Script:LogFile -Destination $networkLogFile -Force
            Write-Log "Log copied to network: $networkLogFile"
        }
    }
    catch {
        Write-Log "Failed to copy log to network: $($_.Exception.Message)" -Level WARN
        # Don't fail the script if log copy fails
    }
}

# ============================================================================
# USER NOTIFICATION FUNCTIONS (msg.exe only - works from SYSTEM)
# ============================================================================
function Get-ActiveUserSession {
    <#
    .SYNOPSIS
    Detects active user session for msg.exe targeting
    PRIMARY: explorer.exe process (SYSTEM can always see this)
    FALLBACK: query commands (may not work on all systems from SYSTEM context)
    #>
    
    # Method 1: Check for explorer.exe (MOST RELIABLE from SYSTEM)
    Write-Log "Checking for active session via explorer.exe process..."
    try {
        $explorerProcess = Get-Process -Name "explorer" -ErrorAction SilentlyContinue | 
            Where-Object { $_.SessionId -gt 0 } | 
            Select-Object -First 1
        
        if ($explorerProcess) {
            $sessionId = $explorerProcess.SessionId
            $Script:ActiveSessionId = $sessionId
            Write-Log "Active session detected via explorer.exe: SessionID=$sessionId" -Level SUCCESS
            return $true
        } else {
            Write-Log "No explorer.exe process found with SessionId > 0"
        }
    }
    catch {
        Write-Log "Explorer.exe check failed: $($_.Exception.Message)" -Level WARN
    }
    
    # Method 2: Try query commands as fallback
    Write-Log "Attempting query commands as fallback..."
    try {
        $queryResult = $null
        try {
            $queryResult = cmd /c "query user 2>nul" 2>$null
        } catch { }

        if (-not $queryResult -or $queryResult -notmatch "Active") {
            Write-Log "query user returned no active sessions - trying query session fallback"
            try {
                $queryResult = cmd /c "query session 2>nul" 2>$null
            } catch { }
        }

        if (-not $queryResult) {
            Write-Log "No user sessions found via query commands" -Level WARN
            return $null
        }

        Write-Log "Raw query output: $($queryResult -join ' | ')"

        # Find all Active sessions
        $activeSessions = @($queryResult | Where-Object { $_ -match "Active" })
        if (-not $activeSessions -or $activeSessions.Count -eq 0) {
            Write-Log "No active sessions found in query output" -Level WARN
            return $null
        }

        if ($activeSessions.Count -gt 1) {
            Write-Log "Multiple active sessions detected ($($activeSessions.Count)) - targeting first session" -Level WARN
            foreach ($s in $activeSessions) { Write-Log "  Session: $($s.Trim())" }
        }

        $activeSession = $activeSessions[0]

        # Parse session ID
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
            return $true
        } else {
            $Script:ActiveSessionId = $null
            Write-Log "Could not parse session ID from query output - will fall back to broadcast" -Level WARN
            return $true  # Still return true - we'll use broadcast
        }
    }
    catch {
        Write-Log "Query commands failed: $($_.Exception.Message)" -Level WARN
        return $null
    }
}

function Show-UserMessage {
    <#
    .SYNOPSIS
    Displays message to user via msg.exe
    Falls back to broadcast if no specific session ID available
    Uses Sysnative path for 32-bit PowerShell on 64-bit systems
    
    .PARAMETER Message
    Message text to display
    
    .PARAMETER Timeout
    Seconds before auto-close (0 = never, waits for OK)
    #>
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

    try {
        # METHOD A: Detect if running as 32-bit and use Sysnative to bypass WOW64 redirection
        $is32bit = [System.IntPtr]::Size -eq 4
        $msgExePath = if ($is32bit -and (Test-Path "$env:windir\Sysnative\msg.exe")) {
            "$env:windir\Sysnative\msg.exe"
        } else {
            "$env:windir\System32\msg.exe"
        }
        
        if ($Timeout -gt 0) {
            $result = & $msgExePath $msgTarget /TIME:$Timeout $Message 2>&1
        } else {
            # /W means wait indefinitely for user to click OK
            $result = & $msgExePath $msgTarget /W $Message 2>&1
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
# SYSTEM CHECK FUNCTIONS
# ============================================================================
function Test-DiskSpace {
    Write-Log "Checking available disk space..."
    
    try {
        $drive = Get-PSDrive -Name C -ErrorAction Stop
        $freeSpaceGB = [math]::Round($drive.Free / 1GB, 2)
        
        Write-Log "Available disk space: $freeSpaceGB GB"
        
        if ($freeSpaceGB -lt $Config.MinDiskSpaceGB) {
            $diskSpaceMessage = @"
SAS Installation - System Check Failed
Insufficient disk space.
- Only $freeSpaceGB GB free out of 20 GB needed
- Please contact ithelp@yourdomain.edu for assistance.
-IT Support
"@
            $null = Show-UserMessage -Message $diskSpaceMessage
            Write-Log "Insufficient disk space: $freeSpaceGB GB (need $($Config.MinDiskSpaceGB) GB)" -Level ERROR
            return $false
        }
        
        Write-Log "Disk space check passed" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Failed to check disk space: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Test-NetworkAccess {
    Write-Log "Testing network access to: $($Config.NetworkSourcePath)"
    
    try {
        # Use -ErrorAction Stop to ensure errors are caught
        $pathExists = Test-Path $Config.NetworkSourcePath -ErrorAction Stop
        
        if (-not $pathExists) {
            Write-Log "Network share path does not exist - returning false" -Level ERROR
            
            $networkMessage = @"
SAS Installation - Network Check Failed
Cannot access your organization's network. If you are on...
- WiFi or Remote: Connect to the VPN and retry.
- Wired: Toggle WiFi off/on and retry.
-IT Support
"@
            $null = Show-UserMessage -Message $networkMessage
            Write-Log "Test-NetworkAccess returning FALSE (path not found)" -Level ERROR
            return $false
        }
        
        # Test if setup.exe exists
        $setupExe = Join-Path $Config.NetworkSourcePath "setup.exe"
        $setupExists = Test-Path $setupExe -ErrorAction Stop
        
        if (-not $setupExists) {
            Write-Log "setup.exe not found at network source - returning false" -Level ERROR
            Write-Log "Test-NetworkAccess returning FALSE (setup.exe not found)" -Level ERROR
            return $false
        }
        
        Write-Log "Network access verified" -Level SUCCESS
        Write-Log "Test-NetworkAccess returning TRUE" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "Network access test exception: $($_.Exception.Message)" -Level ERROR
        
        $networkMessage = @"
SAS Installation - Network Check Failed
Cannot access your organization's network. If you are on...
- WiFi or Remote: Connect to the VPN and retry.
- Wired: Toggle WiFi off/on and retry.
-IT Support
"@
        $null = Show-UserMessage -Message $networkMessage
        Write-Log "Test-NetworkAccess returning FALSE (exception caught)" -Level ERROR
        return $false
    }
}

function Test-SASInstalled {
    Write-Log "Checking for existing SAS installation..."
    
    $sasExe = Join-Path $Config.SASFoundationPath "sas.exe"
    
    if (Test-Path $sasExe) {
        Write-Log "SAS is already installed at: $($Config.SASFoundationPath)" -Level SUCCESS
        return $true
    }
    
    Write-Log "SAS not detected - installation needed"
    return $false
}

# ============================================================================
# INSTALLATION FUNCTIONS
# ============================================================================
function Set-CopyInProgressMarker {
    try {
        if (-not (Test-Path (Split-Path $Config.CopyInProgressMarker))) {
            New-Item -ItemType Directory -Path (Split-Path $Config.CopyInProgressMarker) -Force | Out-Null
        }
        "Copy started: $(Get-Date)" | Out-File -FilePath $Config.CopyInProgressMarker -Force
        Write-Log "Set copy-in-progress marker"
        return $true
    }
    catch {
        Write-Log "Failed to set copy marker: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Remove-CopyInProgressMarker {
    try {
        if (Test-Path $Config.CopyInProgressMarker) {
            Remove-Item $Config.CopyInProgressMarker -Force
            Write-Log "Removed copy-in-progress marker"
        }
        return $true
    }
    catch {
        Write-Log "Failed to remove copy marker: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Copy-InstallationFiles {
    Write-Log "=============================================="
    Write-Log "Checking for existing installation files"
    Write-Log "Source: $($Config.NetworkSourcePath)"
    Write-Log "Destination: $($Config.LocalTempPath)"
    Write-Log "=============================================="
    
    # Check if files already exist from previous attempt
    if (Test-Path $Config.LocalTempPath) {
        Write-Log "Found existing directory: $($Config.LocalTempPath)"
        
        # Verify if it's a complete copy
        $localFiles = Get-ChildItem -Path $Config.LocalTempPath -Recurse -File -ErrorAction SilentlyContinue
        $localFileCount = ($localFiles | Measure-Object).Count
        $localTotalBytes = ($localFiles | Measure-Object -Property Length -Sum).Sum
        $localGB = [math]::Round($localTotalBytes / 1GB, 2)
        
        Write-Log "Existing files: $localGB GB ($localFileCount files)"
        
        # Check if response file exists
        $responseFileExists = Test-Path (Join-Path $Config.LocalTempPath "sas94m9install.properties")
        
        # Check if setup.exe exists
        $setupExeExists = Test-Path (Join-Path $Config.LocalTempPath "setup.exe")
        
        # If we have a reasonable amount of data and key files exist, verify exact file count
        if ($localGB -gt 13.6 -and $localFileCount -gt 17100 -and $responseFileExists -and $setupExeExists) {
            Write-Log "Existing files appear complete - verifying exact file count..."
            
            # Count source files for comparison
            try {
                $sourceFiles = @(Get-ChildItem -Path $Config.NetworkSourcePath -Recurse -File -ErrorAction Stop)
                $sourceCount = $sourceFiles.Count
                Write-Log "Source has $sourceCount files, local has $localFileCount files"
                
                # If exact match, show "Installation Files Located" and proceed
                if ($localFileCount -eq $sourceCount) {
                    Write-Log "Exact file count match - existing files are complete"
                    
                    # Show "Installation Files Located" message
                    $filesFoundMessage = @"
SAS Installation - Installation Files Located
- Installation files found on this computer.
- Minimum system check not needed.
- Installation will now begin - this will take 60-90 minutes.
-IT Support
"@
                    $null = Show-UserMessage -Message $filesFoundMessage -Timeout $Config.StatusMessageTimeout
                    
                    $Script:FilesAlreadyExisted = $true
                    return $true
                }
                else {
                    # Files are missing - show "System Requirements Passed" message FIRST
                    Write-Log "File count mismatch: Local has $localFileCount, expected $sourceCount"
                    Write-Log "Missing files detected - will copy missing files..."
                    
                    # Show "System Requirements Passed / Copying Started" message
                    $copyStartMessage = @"
SAS Installation - System Requirement Check Passed
File Copying has started.
- This will take approximately 30 - 60 minutes depending on network speeds.
- Do not shut down or disconnect from the network during this time.
-IT Support
"@
                    $null = Show-UserMessage -Message $copyStartMessage -Timeout $Config.StatusMessageTimeout
                    
                    # NOW start copying missing files (while message is being shown/after it closes)
                    Write-Log "Attempting to copy missing files..."
                    
                    # First retry attempt
                    $copiedCount = Copy-MissingFiles -SourcePath $Config.NetworkSourcePath -DestinationPath $Config.LocalTempPath
                    
                    if ($copiedCount -gt 0) {
                        Write-Log "First retry: Copied $copiedCount missing files"
                        
                        # Recount after first retry
                        $localFiles = Get-ChildItem -Path $Config.LocalTempPath -Recurse -File
                        $localFileCount = $localFiles.Count
                        Write-Log "After first retry: $localFileCount files"
                        
                        # Check if still missing files
                        if ($localFileCount -ne $sourceCount) {
                            Write-Log "Still missing files after first retry - waiting 3 seconds..."
                            Start-Sleep -Seconds 3
                            
                            # Second retry attempt
                            $copiedCount2 = Copy-MissingFiles -SourcePath $Config.NetworkSourcePath -DestinationPath $Config.LocalTempPath
                            Write-Log "Second retry: Copied $copiedCount2 missing files"
                            
                            # Final recount
                            $localFiles = Get-ChildItem -Path $Config.LocalTempPath -Recurse -File
                            $localFileCount = $localFiles.Count
                            Write-Log "After second retry: $localFileCount files"
                            
                            # Final check
                            if ($localFileCount -eq $sourceCount) {
                                Write-Log "All files now present after second retry" -Level SUCCESS
                                $Script:FilesAlreadyExisted = $false  # Files were copied, not just found
                                return $true
                            }
                            else {
                                Write-Log "File count still mismatched after 2 retry attempts: $localFileCount / $sourceCount" -Level ERROR
                                
                                # Show error and fail
                                $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
                                $null = Show-UserMessage -Message $verifyFailMessage
                                return $false
                            }
                        }
                        else {
                            Write-Log "All files now present after first retry" -Level SUCCESS
                            $Script:FilesAlreadyExisted = $false  # Files were copied, not just found
                            return $true
                        }
                    }
                    else {
                        Write-Log "No missing files could be copied" -Level ERROR
                        
                        # Show error and fail
                        $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
                        $null = Show-UserMessage -Message $verifyFailMessage
                        return $false
                    }
                }
            }
            catch {
                Write-Log "Failed to verify source file count: $($_.Exception.Message)" -Level ERROR
                
                # Show error and fail
                $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
                $null = Show-UserMessage -Message $verifyFailMessage
                return $false
            }
        }
        else {
            Write-Log "Existing files incomplete (${localGB}GB, ${localFileCount} files, response=${responseFileExists}, setup=${setupExeExists}) - will re-copy"
            Write-Log "Removing incomplete directory..."
            Remove-Item $Config.LocalTempPath -Recurse -Force -ErrorAction Stop
        }
    }
    
    Write-Log "=============================================="
    Write-Log "Checking for existing installation files"
    Write-Log "=============================================="
    
    try {
        # Create destination directory
        if (Test-Path $Config.LocalTempPath) {
            Write-Log "Removing existing temp directory..."
            Remove-Item $Config.LocalTempPath -Recurse -Force -ErrorAction Stop
        }
        
        New-Item -ItemType Directory -Path $Config.LocalTempPath -Force | Out-Null
        Write-Log "Created temp directory: $($Config.LocalTempPath)"
        
        # Set copy-in-progress marker
        Set-CopyInProgressMarker
        
        # Calculate total size
        Write-Log "Calculating total file size..."
        $sourceItems = Get-ChildItem -Path $Config.NetworkSourcePath -Recurse -File
        $totalBytes = ($sourceItems | Measure-Object -Property Length -Sum).Sum
        $totalGB = [math]::Round($totalBytes / 1GB, 2)
        
        Write-Log "Total size to copy: $totalGB GB ($($sourceItems.Count) files)"
        
        # Calculate estimated time (assume 50 MB/s average)
        $estimatedSeconds = $totalBytes / (50 * 1024 * 1024)
        $estimatedMinutes = [math]::Ceiling($estimatedSeconds / 60)
        
        Write-Log "Estimated copy time: $estimatedMinutes minutes"
        
        # Show initial message (Message 1)
        $copyStartMessage = @"
SAS Installation - System Requirement Check Passed
File Copying has started.
- This will take approximately 30 - 60 minutes depending on network speeds.
- Do not shut down or disconnect from the network during this time.
-IT Support
"@
        $null = Show-UserMessage -Message $copyStartMessage -Timeout $Config.StatusMessageTimeout
        
        # Copy files with progress tracking
        Write-Log "Starting file copy operation..."
        
        $copiedBytes = 0
        $lastProgressTime = Get-Date
        $lastProgressBytes = 0
        $fiftyPercentShown = $false
        $almostDoneShown = $false
        
        foreach ($item in $sourceItems) {
            $relativePath = $item.FullName.Substring($Config.NetworkSourcePath.Length).TrimStart('\')
            $destPath = Join-Path $Config.LocalTempPath $relativePath
            $destDir = Split-Path $destPath
            
            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            
            Copy-Item -Path $item.FullName -Destination $destPath -Force
            
            $copiedBytes += $item.Length
            $percentComplete = [math]::Round(($copiedBytes / $totalBytes) * 100, 0)
            
            # Check for timeout (no progress)
            $timeSinceLastCheck = (Get-Date) - $lastProgressTime
            $bytesSinceLastCheck = $copiedBytes - $lastProgressBytes
            
            if ($timeSinceLastCheck.TotalMinutes -ge 1) {
                $currentSpeed = if ($timeSinceLastCheck.TotalSeconds -gt 0) { $bytesSinceLastCheck / $timeSinceLastCheck.TotalSeconds / 1MB } else { 0 }
                Write-Log "Copy progress: $percentComplete% ($copiedBytes / $totalBytes bytes, $([math]::Round($currentSpeed, 2)) MB/s)"
                
                # Copy log to network during progress updates (every minute)
                Copy-LogToNetwork
                
                $lastProgressTime = Get-Date
                $lastProgressBytes = $copiedBytes
                
                # Calculate time remaining
                $bytesRemaining = $totalBytes - $copiedBytes
                $estimatedSecondsRemaining = if ($currentSpeed -gt 0) { $bytesRemaining / ($currentSpeed * 1MB) } else { 999 }
                $estimatedMinutesRemaining = [math]::Ceiling($estimatedSecondsRemaining / 60)
                
                # Show 50% progress message (Message 5)
                if ($percentComplete -ge 50 -and -not $fiftyPercentShown) {
                    $fiftyPercentMessage = @"
SAS Installation - File Copy Progress
- Copying files: 50% complete
- Estimated time remaining: $estimatedMinutesRemaining minutes
-IT Support
"@
                    $null = Show-UserMessage -Message $fiftyPercentMessage -Timeout $Config.StatusMessageTimeout
                    $fiftyPercentShown = $true
                }
                
                # Show almost done message when 5 minutes remaining (Message 3)
                if ($estimatedMinutesRemaining -le 5 -and $percentComplete -ge 85 -and -not $almostDoneShown) {
                    $almostDoneMessage = @"
SAS Installation - File Copy Progress

Copying files: $percentComplete% complete. Almost done.
Estimated time remaining: 5 minutes

-IT Support
"@
                    $null = Show-UserMessage -Message $almostDoneMessage -Timeout $Config.StatusMessageTimeout
                    $almostDoneShown = $true
                }
            }
            
            # Check for timeout (no progress for timeout period)
            if ($bytesSinceLastCheck -eq 0 -and $timeSinceLastCheck.TotalMinutes -ge $Config.CopyTimeoutMinutes) {
                Write-Log "Copy timeout - no progress for $($Config.CopyTimeoutMinutes) minutes" -Level ERROR
                
                # Message 8: File Copy Timeout
                $copyTimeoutMessage = @"
SAS Installation - Timeout Error
- It was not able to be determined if SAS installed correctly.
- Please retry or contact ithelp@yourdomain.edu for assistance.
-IT Support
"@
                $null = Show-UserMessage -Message $copyTimeoutMessage
                Remove-CopyInProgressMarker
                return $false
            }
        }
        
        Write-Log "File copy completed successfully" -Level SUCCESS
        Write-Log "Total copied: $totalGB GB"
        
        Remove-CopyInProgressMarker
        return $true
    }
    catch {
        Write-Log "File copy failed: $($_.Exception.Message)" -Level ERROR
        Remove-CopyInProgressMarker
        return $false
    }
}

function Copy-MissingFiles {
    param (
        [string]$SourcePath,
        [string]$DestinationPath
    )
    
    Write-Log "Identifying missing files..."
    
    try {
        # Get all source files with relative paths
        $sourceFiles = Get-ChildItem -Path $SourcePath -Recurse -File -ErrorAction SilentlyContinue
        $missingCount = 0
        $copiedCount = 0
        
        foreach ($sourceFile in $sourceFiles) {
            # Calculate relative path
            $relativePath = $sourceFile.FullName.Substring($SourcePath.Length).TrimStart('\')
            $destFile = Join-Path $DestinationPath $relativePath
            
            # Check if file exists in destination
            if (-not (Test-Path $destFile)) {
                $missingCount++
                Write-Log "Missing file: $relativePath"
                
                # Create destination directory if needed
                $destDir = Split-Path $destFile -Parent
                if (-not (Test-Path $destDir)) {
                    New-Item -Path $destDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                }
                
                # Copy the missing file using robocopy for long path support
                try {
                    $sourceDir = Split-Path $sourceFile.FullName -Parent
                    $fileName = Split-Path $sourceFile.FullName -Leaf
                    
                    # Use robocopy to handle long paths (>260 characters)
                    # /R:2 = Retry 2 times, /W:3 = Wait 3 seconds between retries, /NFL /NDL /NJH /NJS = Minimal output
                    $robocopyArgs = @(
                        "`"$sourceDir`"",
                        "`"$destDir`"",
                        "`"$fileName`"",
                        "/R:2",
                        "/W:3",
                        "/NFL",
                        "/NDL",
                        "/NJH",
                        "/NJS"
                    )
                    
                    $result = & robocopy @robocopyArgs 2>&1
                    
                    # Robocopy exit codes: 0-7 are success (0-3 = success, 4-7 = some files copied)
                    # 8+ indicates errors
                    if ($LASTEXITCODE -lt 8) {
                        $copiedCount++
                    }
                    else {
                        Write-Log "Failed to copy $relativePath : Robocopy exit code $LASTEXITCODE" -Level ERROR
                    }
                }
                catch {
                    Write-Log "Failed to copy $relativePath : $($_.Exception.Message)" -Level ERROR
                }
            }
        }
        
        Write-Log "Found $missingCount missing files, successfully copied $copiedCount files"
        return $copiedCount
    }
    catch {
        Write-Log "Error identifying missing files: $($_.Exception.Message)" -Level ERROR
        return 0
    }
}

function Test-CopiedFiles {
    Write-Log "Verifying copied files..."
    
    try {
        # Check if setup.exe exists
        $setupExe = Join-Path $Config.LocalTempPath "setup.exe"
        if (-not (Test-Path $setupExe)) {
            Write-Log "setup.exe not found in copied files" -Level ERROR
            
            # Message 9: File Verification Failed
            $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
            $null = Show-UserMessage -Message $verifyFailMessage
            return $false
        }
        
        # Count source files
        Write-Log "Counting source files..."
        Write-Log "Source path: $($Config.NetworkSourcePath)"
        
        try {
            $sourceFiles = @(Get-ChildItem -Path $Config.NetworkSourcePath -Recurse -File -ErrorAction Stop)
            $sourceCount = $sourceFiles.Count
            $sourceBytes = ($sourceFiles | Measure-Object -Property Length -Sum).Sum
            $sourceGB = [math]::Round($sourceBytes / 1GB, 2)
            Write-Log "Source: $sourceGB GB ($sourceCount files)"
        }
        catch {
            Write-Log "Failed to count source files: $($_.Exception.Message)" -Level ERROR
            Write-Log "Cannot verify file copy without source file count" -Level ERROR
            return $false
        }
        
        # Count copied files
        $copiedFiles = Get-ChildItem -Path $Config.LocalTempPath -Recurse -File
        $copiedCount = $copiedFiles.Count
        $copiedBytes = ($copiedFiles | Measure-Object -Property Length -Sum).Sum
        $copiedGB = [math]::Round($copiedBytes / 1GB, 2)
        Write-Log "Copied: $copiedGB GB ($copiedCount files)"
        
        # Exact file count match required
        if ($copiedCount -ne $sourceCount) {
            Write-Log "File count mismatch: Copied $copiedCount files, expected $sourceCount files" -Level ERROR
            $missingFiles = $sourceCount - $copiedCount
            Write-Log "Missing $missingFiles files - attempting to copy missing files..."
            
            # First retry attempt
            $copiedInRetry = Copy-MissingFiles -SourcePath $Config.NetworkSourcePath -DestinationPath $Config.LocalTempPath
            
            if ($copiedInRetry -gt 0) {
                Write-Log "First retry: Copied $copiedInRetry missing files"
                
                # Recount after first retry
                $copiedFiles = Get-ChildItem -Path $Config.LocalTempPath -Recurse -File
                $copiedCount = $copiedFiles.Count
                Write-Log "After first retry: $copiedCount files"
                
                # Check if still missing files
                if ($copiedCount -ne $sourceCount) {
                    Write-Log "Still missing files after first retry - waiting 3 seconds..."
                    Start-Sleep -Seconds 3
                    
                    # Second retry attempt
                    $copiedInRetry2 = Copy-MissingFiles -SourcePath $Config.NetworkSourcePath -DestinationPath $Config.LocalTempPath
                    Write-Log "Second retry: Copied $copiedInRetry2 missing files"
                    
                    # Final recount
                    $copiedFiles = Get-ChildItem -Path $Config.LocalTempPath -Recurse -File
                    $copiedCount = $copiedFiles.Count
                    Write-Log "After second retry: $copiedCount files"
                    
                    # Final check
                    if ($copiedCount -ne $sourceCount) {
                        Write-Log "File count still mismatched after 2 retry attempts: $copiedCount / $sourceCount" -Level ERROR
                        
                        # Message 9: File Verification Failed
                        $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
                        $null = Show-UserMessage -Message $verifyFailMessage
                        return $false
                    }
                    else {
                        Write-Log "All files successfully copied after second retry" -Level SUCCESS
                    }
                }
                else {
                    Write-Log "All files successfully copied after first retry" -Level SUCCESS
                }
            }
            else {
                Write-Log "No files copied in retry attempt" -Level ERROR
                
                # Message 9: File Verification Failed
                $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
                $null = Show-UserMessage -Message $verifyFailMessage
                return $false
            }
        }
        
        # Sanity check - should be around 14GB
        if ($copiedGB -lt 10) {
            Write-Log "Copied file size too small: $copiedGB GB (expected ~14GB)" -Level ERROR
            
            # Message 9: File Verification Failed
            $verifyFailMessage = @"
SAS Installation - File Copy Error

File verification failed.
Installation files may be incomplete or corrupted.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.
"@
            $null = Show-UserMessage -Message $verifyFailMessage
            return $false
        }
        
        Write-Log "File verification passed: All $copiedCount files copied successfully" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "File verification failed: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Start-SASInstallation {
    Write-Log "=============================================="
    Write-Log "Starting SAS installation"
    Write-Log "=============================================="
    
    try {
        # Show installation starting message only if files were actually copied
        if (-not $Script:FilesAlreadyExisted) {
            $installStartMessage = @"
SAS Installation - All Files Copied Over
- Installing SAS 9.4 now (approx. 60 minutes).
- Do not shut down or restart your computer.
- A message will display when the installation is complete.
-IT Support
"@
            $null = Show-UserMessage -Message $installStartMessage -Timeout $Config.StatusMessageTimeout
        }
        
        # License file and properties file are in the depot (no copying needed)
        # License: C:\Temp\SAS_Install_2026\sid_files\SAS_94_TR_2026_SPH_License.txt
        # Properties: C:\Temp\SAS_Install_2026\sas94m9install.properties
        
        # Run setup.exe with response file from depot
        $setupExe = Join-Path $Config.LocalTempPath "setup.exe"
        $responseFile = Join-Path $Config.LocalTempPath "sas94m9install.properties"
        
        Write-Log "Launching setup.exe: $setupExe"
        Write-Log "Response file: $responseFile"
        Write-Log "License file: Located in depot sid_files folder (specified in response file)"
        Write-Log "Command: setup.exe -quiet -wait -responsefile C:\Temp\SAS_Install_2026\sas94m9install.properties -loglevel 2"
        
        $startTime = Get-Date
        $process = Start-Process -FilePath "C:\Temp\SAS_Install_2026\setup.exe" -ArgumentList "-quiet", "-wait", "-responsefile", "C:\Temp\SAS_Install_2026\sas94m9install.properties", "-loglevel", "2" -Wait -PassThru -NoNewWindow
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-Log "Setup.exe completed in $($duration.TotalMinutes) minutes"
        Write-Log "Exit code: $($process.ExitCode)"
        
        # Copy log to network after setup completes
        Copy-LogToNetwork
        
        # Check if installation took too long (Message 10)
        if ($duration.TotalMinutes -gt $Config.InstallTimeoutMinutes) {
            Write-Log "Installation exceeded timeout of $($Config.InstallTimeoutMinutes) minutes" -Level ERROR
            
            $installTimeoutMessage = @"
SAS Installation - Installation Error

SAS installation has exceeded the expected time of 180 minutes.

It is recommended to restart your computer and try again.

Contact ithelp@yourdomain.edu if this problem persists.

-IT Support
"@
            $null = Show-UserMessage -Message $installTimeoutMessage
            return $false
        }
        
        # Check exit code
        if ($process.ExitCode -eq 0) {
            Write-Log "Installation completed successfully" -Level SUCCESS
            return $true
        }
        else {
            Write-Log "Installation failed with exit code: $($process.ExitCode)" -Level ERROR
            return $false
        }
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Remove-TemporaryFiles {
    Write-Log "Cleaning up temporary files..."
    
    try {
        if (Test-Path $Config.LocalTempPath) {
            Remove-Item $Config.LocalTempPath -Recurse -Force -ErrorAction Stop
            Write-Log "Removed temp directory: $($Config.LocalTempPath)" -Level SUCCESS
        }
        return $true
    }
    catch {
        Write-Log "Failed to remove temp files: $($_.Exception.Message)" -Level WARN
        return $false
    }
}

function Set-RegistryDetectionKey {
    Write-Log "Setting registry detection key for Intune..."
    
    try {
        if (-not (Test-Path $Config.RegistryPath)) {
            New-Item -Path $Config.RegistryPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $Config.RegistryPath -Name $Config.RegistryValueName -Value $Config.LicenseYear -Type String -Force
        
        $verify = Get-ItemProperty -Path $Config.RegistryPath -Name $Config.RegistryValueName -ErrorAction Stop
        
        if ($verify.$($Config.RegistryValueName) -eq $Config.LicenseYear) {
            Write-Log "Registry key set successfully: $($Config.RegistryPath)\$($Config.RegistryValueName) = $($Config.LicenseYear)" -Level SUCCESS
            return $true
        }
        else {
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
# MAIN INSTALLATION ORCHESTRATION
# ============================================================================
function Start-Installation {
    Write-Log "=============================================="
    Write-Log "SAS 9.4 Silent Installation"
    Write-Log "Script Version: $Script:ScriptVersion ($Script:VersionDate)"
    Write-Log "License Year: $($Config.LicenseYear)"
    Write-Log "Attempt Number: $Script:RetryAttempt of $($Config.MaxRetryAttempts)"
    Write-Log "Computer: $env:COMPUTERNAME"
    Write-Log "User Context: $env:USERNAME"
    Write-Log "PowerShell: $($PSVersionTable.PSVersion) ($([System.IntPtr]::Size * 8)-bit process)"
    Write-Log "=============================================="
    
    # Check if user session is active
    $session = Get-ActiveUserSession
    if ($null -eq $session) {
        Write-Log "No active user session detected - deferring installation" -Level WARN
        return 2  # Exit code 2 = retry later
    }
    
    # Check if already installed
    if (Test-SASInstalled) {
        Write-Log "SAS is already installed - setting registry key and exiting"
        if (Set-RegistryDetectionKey) {
            Write-Log "Installation check complete - no action needed" -Level SUCCESS
            return 0
        }
        else {
            Write-Log "Failed to set registry detection key" -Level ERROR
            return 1
        }
    }
    
    # Clean up any partial installations or suspended deployments
    Write-Log "Checking for partial installations or suspended deployments..."
    
    # Remove partial SASHome installation
    $sasHomePath = "C:\Program Files\SASHome"
    if (Test-Path $sasHomePath) {
        Write-Log "Found partial SASHome installation - removing..."
        try {
            Remove-Item $sasHomePath -Recurse -Force -ErrorAction Stop
            Write-Log "Removed partial SASHome installation" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to remove partial SASHome: $($_.Exception.Message)" -Level WARN
        }
    }
    
    # Remove SAS Deployment Wizard state files to prevent resume prompts
    $deployWizardPath = "$env:USERPROFILE\AppData\Local\SAS\SASDeploymentWizard"
    if (Test-Path $deployWizardPath) {
        Write-Log "Found SAS Deployment Wizard state files - removing..."
        try {
            Remove-Item "$deployWizardPath\*" -Recurse -Force -ErrorAction Stop
            Write-Log "Removed SAS Deployment Wizard state files" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to remove deployment wizard state: $($_.Exception.Message)" -Level WARN
        }
    }
    
    # Remove temp setup directories
    $tempSetupDirs = Get-ChildItem "C:\Windows\Temp\_setup*" -Directory -ErrorAction SilentlyContinue
    if ($tempSetupDirs) {
        Write-Log "Found temp setup directories - removing..."
        foreach ($dir in $tempSetupDirs) {
            try {
                Remove-Item $dir.FullName -Recurse -Force -ErrorAction Stop
                Write-Log "Removed temp setup directory: $($dir.Name)" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to remove temp setup directory $($dir.Name): $($_.Exception.Message)" -Level WARN
            }
        }
    }
    
    Write-Log "Cleanup complete - ready for fresh installation"
    
    # Message 1: Process Starting (show BEFORE system checks)
    $processStartMessage = @"
SAS 9.4 Silent Installation - This process takes 60 - 120 minutes.
Steps that will take place:
- Check minimum system requirements
- Copy installation files over
- Silently install SAS 9.4 with 2026 license
"@
    $null = Show-UserMessage -Message $processStartMessage
    
    # System requirements check
    Write-Log "Checking system requirements..."
    
    # Check disk space (one-time check, no retry)
    if (-not (Test-DiskSpace)) {
        Write-Log "Disk space check failed - exiting" -Level ERROR
        
        # Message: Disk Space Error
        $diskSpaceMessage = @"
SAS Installation - Insufficient Disk Space

Not enough free disk space to install SAS 9.4.
At least 20 GB required on the C: drive.

-IT Support
"@
        $null = Show-UserMessage -Message $diskSpaceMessage
        return 1
    }
    
    if (-not (Test-NetworkAccess)) {
        Write-Log "Network access check failed" -Level ERROR
        
        # Message: Network Error
        $networkErrorMessage = @"
SAS Installation - Network Check Failed
Cannot access your organization's network. If you are on...
- WiFi or Remote: Connect to the VPN and retry.
- Wired: Toggle WiFi off/on and retry.
-IT Support
"@
        $null = Show-UserMessage -Message $networkErrorMessage
        return 1
    }
    
    Write-Log "All system checks passed" -Level SUCCESS
    
    # Copy files from network
    if (-not (Copy-InstallationFiles)) {
        Write-Log "File copy failed" -Level ERROR
        return 1
    }
    
    # Verify copied files
    if (-not (Test-CopiedFiles)) {
        Write-Log "File verification failed" -Level ERROR
        return 1
    }
    
    # Install SAS
    if (-not (Start-SASInstallation)) {
        Write-Log "Installation failed - keeping temp files for troubleshooting" -Level ERROR
        
        # Message 11: Installation Failed
        $failureMessage = @"
SAS Installation - Installation Error

SAS installation did not complete successfully.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if the problem persists.

-IT Support
"@
        $null = Show-UserMessage -Message $failureMessage
        
        # Keep temp files for retry/troubleshooting
        return 1
    }
    
    # Verify installation
    if (-not (Test-SASInstalled)) {
        Write-Log "Installation verification failed - keeping temp files for troubleshooting" -Level ERROR
        
        # Message 11: Installation Failed
        $failureMessage = @"
SAS Installation - Installation Error

SAS installation did not complete successfully.

Please restart your computer and try again.

Contact ithelp@yourdomain.edu if the problem persists.

-IT Support
"@
        $null = Show-UserMessage -Message $failureMessage
        
        # Keep temp files for retry/troubleshooting
        return 1
    }
    
    # Set registry detection key
    if (-not (Set-RegistryDetectionKey)) {
        Write-Log "Installation succeeded but failed to set registry key" -Level WARN
    }
    
    # Cleanup
    Remove-TemporaryFiles
    
    # Show success message (Message 7)
    $successMessage = @"
SAS Installation - Installation Complete
- SAS 9.4 with 2026 license has been successfully installed.
- You can now use SAS.
-IT Support
"@
    $null = Show-UserMessage -Message $successMessage
    
    Write-Log "=============================================="
    Write-Log "Installation completed successfully" -Level SUCCESS
    Write-Log "=============================================="
    
    return 0
}

# ============================================================================
# ENTRY POINT - NO RETRY LOOP
# ============================================================================
try {
    Initialize-Logging
    
    # Copy log to network immediately to confirm script started
    Copy-LogToNetwork
    
    $Script:RetryAttempt = 1  # Set to 1 for logging purposes
    $exitCode = Start-Installation
    
    Remove-OldLogs
    
    Write-Log "Script exiting with code: $exitCode"
}
catch {
    Write-Log "Unhandled exception: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    $exitCode = 1
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
    Copy-LogToNetwork
    exit $exitCode
}

# ============================================================================
# VERSION HISTORY
# ============================================================================
# 2.3.16 (2026-03-11) - Added robocopy support for long paths (>260 chars) when copying missing files
# 2.3.15 (2026-03-09) - Added missing file retry logic (2 attempts with 3s delay)
# 2.3.14 (2026-03-09) - Updated files located message
# 2.3.13 (2026-03-04) - Removed license copy workaround, response file reads correctly
# 2.3.12 (2026-03-03) - Fixed msg.exe /W, removed retry loop, hardcoded response path
# 2.3.11 (2026-03-02) - Copy license to _setup0, thresholds >13.6GB/>17100 files
# 2.3.10 (2026-03-02) - Check for existing files, keep temp files on failure
# 2.3.9 (2026-03-02) - License/properties in depot sid_files, reordered command parameters
# 2.3.8 (2026-02-27) - Added -loglevel 2, cleanup of partial installs and deployment wizard state
# 2.3.7 (2026-02-26) - Message 1 (Process Starting) has no timeout, user must click OK
# 2.3.6 (2026-02-26) - License and properties files copied to temp location before install
# 2.3.5 (2026-02-26) - All messages updated: ≤5 lines, hyphens, Method A (Sysnative)
# 2.3.4 (2026-02-25) - Fixed msg.exe for 32-bit PowerShell using Sysnative detection
# 2.3.3 (2026-02-24) - Fixed msg.exe full path and return value suppression
# 2.3.2 (2026-02-23) - Added 3-layer session detection fallback
# 2.3.1 (2026-02-23) - Network log copy and msg.exe improvements
# 2.3.0 (2026-02-20) - Switched to msg.exe for all notifications (VBScript removed)
