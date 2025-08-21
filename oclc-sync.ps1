[CmdletBinding()]
param (
    # The base path for all operational files.
    [Parameter(Mandatory = $false)]
    [string]$BasePath = "c:\ProgramData\clc_oclc_sync",

    # The organization ID to process. If not provided, all organizations will be processed.
    [Parameter(Mandatory = $false)]
    [string]$OrganizationID,

    # The full path to your JSON settings file.
    [Parameter(Mandatory = $false)]
    [string]$SettingsFilePath,

    # The full path to your SQL query file for ADDING records to OCLC.
    [Parameter(Mandatory = $false)]
    [string]$SqlFilePathAdd,

    # The full path to your SQL query file for DELETING records from OCLC.
    [Parameter(Mandatory = $false)]
    [string]$SqlFilePathDelete,

    # The reference date for the query.
    # For ADD operations, records created on or after this date.
    # For DELETE operations, records deleted on or after this date.
    # Defaults to 7 days before the script runs.
    [Parameter(Mandatory = $false)]
    [datetime]$QueryReferenceDate = (Get-Date).AddDays(-2).Date,

    # The SQL Server instance name (e.g., "SERVER\INSTANCE" or "SERVER"). Defaults to "(local)".
    [Parameter(Mandatory = $false)]
    [string]$ServerInstance = "(local)",

    # The name of the Polaris database. Defaults to "polaris".
    [Parameter(Mandatory = $false)]
    [string]$Database = "polaris",

    # The directory where CSV files will be saved.
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    # Full path to WinSCPnet.dll. Example: "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"
    [Parameter(Mandatory = $false)]
    [string]$WinSCPPath,

    # Switch to enable SFTP upload. If not present, upload is skipped.
    [Parameter(Mandatory = $false)]
    [switch]$EnableSftpUpload = $true,

    # The global SFTP hostname. Defaults to "filex-m1.oclc.org".
    [Parameter(Mandatory = $false)]
    [string]$SftpGlobalHostname = "filex-m1.oclc.org",

    # The SSH Host Key Fingerprint for the global SFTP server.
    # Required if -EnableSftpUpload is used. Example: "ssh-rsa 2048 xx:xx:xx..."
    [Parameter(Mandatory = $false)]
    [string]$SftpGlobalHostKeyFingerprint = "ssh-ed25519 255 jzPlRQf9nD6aJEGymaXvLKfP0fq6PFhPSleRbLpM5X0=",

    # The directory where daily log files will be created. Defaults to the ".\logs" directory in the script's location.
    [Parameter(Mandatory = $false)]
    [string]$LogDirectory,

    # Switch to enable logging of generated SQL queries. If not present, this is skipped.
    [Parameter(Mandatory = $false)]
    [switch]$EnableSqlLogging,

    # Switch to enable logging of raw data returned from SQL queries. If not present, this is skipped.
    [Parameter(Mandatory = $false)]
    [switch]$EnableRawSqlDataLogging
)

# --- Path Initializations ---
# If specific paths are not provided, construct them from the BasePath.
if (-not $PSBoundParameters.ContainsKey('SettingsFilePath')) {
    $SettingsFilePath = Join-Path $BasePath "settings.json"
}
if (-not $PSBoundParameters.ContainsKey('SqlFilePathAdd')) {
    $SqlFilePathAdd = Join-Path $BasePath "NewOCLCRecords.sql"
}
if (-not $PSBoundParameters.ContainsKey('SqlFilePathDelete')) {
    $SqlFilePathDelete = Join-Path $BasePath "DeleteOCLCRecords.sql"
}
if (-not $PSBoundParameters.ContainsKey('OutputPath')) {
    $OutputPath = Join-Path $BasePath "output"
}
if (-not $PSBoundParameters.ContainsKey('WinSCPPath')) {
    $WinSCPPath = Join-Path $BasePath "WinSCPnet.dll"
}
if (-not $PSBoundParameters.ContainsKey('LogDirectory')) {
    $LogDirectory = Join-Path $BasePath "logs"
}


# --- Error Collection Setup ---
# Initialize a global, thread-safe list to hold all error messages for a final summary.
$Global:ErrorCollection = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))

# --- Log File Setup ---
if (-not (Test-Path -Path $LogDirectory)) {
    try {
        Write-Host "INFO: Log directory not found at '$LogDirectory'. Creating it..."
        New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "FATAL: Could not create log directory '$LogDirectory'. Error: $($_.Exception.Message)"
        return
    }
}
$TodaysLogFile = Join-Path -Path $LogDirectory -ChildPath "$((Get-Date).ToString('yyyy-MM-dd')).log"


#region Functions
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [string]$Message,

        [Parameter(Mandatory = $False)]
        [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
        [String]$Level = "INFO"
    )

    $Stamp = (Get-Date).toString("yyyy-MM-dd HH:mm:ss")
    $Line = "[$Stamp] [$Level] $Message"
    
    # Write to the daily log file
    try {
        Add-Content -Path $TodaysLogFile -Value $Line -ErrorAction Stop
    }
    catch {
        Write-Host "FATAL: Could not write to log file '$TodaysLogFile'. Error: $($_.Exception.Message)"
    }

    # Add errors to the global collection for summary
    if ($Level -in "ERROR", "FATAL") {
        [void]$Global:ErrorCollection.Add($Line)
    }

    # Write to console with color for immediate feedback
    switch ($Level) {
        "INFO"  { Write-Host $Line -ForegroundColor Green }
        "WARN"  { Write-Host $Line -ForegroundColor Yellow }
        "ERROR" { Write-Host $Line -ForegroundColor Red }
        "FATAL" { Write-Host $Line -ForegroundColor DarkRed }
        "DEBUG" { Write-Host $Line -ForegroundColor Cyan }
        default { Write-Host $Line }
    }
}

Function Start-SftpUpload {
    param (
        [string]$LocalFilePath,
        [string]$RemoteFileName,
        [string]$Username,
        [string]$Password,
        [string]$GlobalHostname,
        [string]$GlobalHostKeyFingerprint,
        [string]$RemoteDirectoryPath,
        [string]$OrganizationName
    )

    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol              = [WinSCP.Protocol]::Sftp
        HostName              = $GlobalHostname
        UserName              = $Username
        Password              = $Password
        SshHostKeyFingerprint = $GlobalHostKeyFingerprint
        Timeout               = [System.TimeSpan]::FromSeconds(60) # Set a 60-second timeout
    }

    $session = New-Object WinSCP.Session
    try {
        Write-Log -Level INFO -Message "      SFTP: Connecting to $GlobalHostname for $($OrganizationName)..."
        $session.Open($sessionOptions)
        Write-Log -Level INFO -Message "      SFTP: Connected."

        if (-not $RemoteDirectoryPath.EndsWith("/")) {
            $RemoteDirectoryPath += "/"
        }
        $fullRemoteFilePath = $RemoteDirectoryPath + $RemoteFileName
        
        Write-Log -Level INFO -Message "      SFTP: Uploading data file '$LocalFilePath' to '$fullRemoteFilePath'..."
        $session.PutFiles($LocalFilePath, $fullRemoteFilePath).Check()
        Write-Log -Level INFO -Message "      SFTP: Data file uploaded successfully."

        return $true # Indicate success
    }
    catch {
        # --- MODIFIED BEHAVIOR ---
        # A failed SFTP attempt is now logged as a WARNING, not an error.
        # This prevents the script from treating it as a critical failure but still highlights the issue.
        Write-Log -Level WARN -Message "      SFTP WARNING: Upload failed on this attempt for '$($OrganizationName)'. Error: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Log -Level WARN -Message "      SFTP WARNING: Inner Exception: $($_.Exception.InnerException.Message)"
        }
        return $false # Indicate failure to allow for retry logic
    }
    finally {
        if ($session.Opened) {
            $session.Close()
            Write-Log -Level INFO -Message "      SFTP: Session closed for $($OrganizationName)."
        }
        $session.Dispose()
    }
}

Function Process-OclcOperation {
    param (
        [string]$OperationType,
        [string]$CollectionID,
        [string]$SqlQueryTemplate,
        [psobject]$Setting,
        [hashtable]$ScriptParameters
    )

    $orgID = $Setting.OrganizationID
    $orgName = $Setting.OrganizationName
    $fileNamePrefix = $Setting.FilenamePrefix
    $sftpUsername = $Setting.SftpUsername
    $sftpPassword = $Setting.SftpPassword
    $sftpRemoteDirectory = $Setting.SftpRemoteDirectory
    $runDate = $ScriptParameters.RunDate
    $FormattedQueryReferenceDate = $ScriptParameters.FormattedQueryReferenceDate
    
    if ($OperationType -eq "ADD") {
        $logAction = "new records for OCLC"
        $outputFileSuffix = "NewOCLCRecords.csv"
    }
    else { # DELETE
        $logAction = "records to delete from OCLC"
        $outputFileSuffix = "DeleteOCLCRecords.csv"
    }

    Write-Log -Level INFO -Message "Processing Org: $orgName (ID: $orgID) for OCLC $OperationType (Collection: $collectionID)"

    $sftpPrereqsMet = $true
    if ($ScriptParameters.EnableSftpUpload) {
        if ([string]::IsNullOrEmpty($sftpUsername) -or [string]::IsNullOrEmpty($sftpPassword) -or [string]::IsNullOrEmpty($sftpRemoteDirectory)) {
            Write-Log -Level WARN -Message "Skipping $OperationType SFTP for Org: $orgName (ID: $orgID) due to missing SFTP configuration."
            $sftpPrereqsMet = $false
        }
    }

    $currentSqlQuery = $SqlQueryTemplate -f $orgID, $FormattedQueryReferenceDate
    
    if ($ScriptParameters.EnableSqlLogging) {
        Write-Log -Level "DEBUG" -Message "`n-- SQL Query for $OperationType -- Org: $orgName (ID: $orgID) --`n$currentSqlQuery`n-- End of Query --`n"
    }

    $csvFileName = "{0}.{1}.{2}.{3}" -f $collectionID, $fileNamePrefix, $runDate, $outputFileSuffix
    $csvFilePath = Join-Path -Path $ScriptParameters.OutputPath -ChildPath $csvFileName

    $conn = $null
    $reader = $null
    $recordCount = 0

    try {
        $conn = New-Object System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = "Server=$($ScriptParameters.ServerInstance);Database=$($ScriptParameters.Database);Integrated Security=True;TrustServerCertificate=$true;"
        Write-Log -Level INFO -Message "      SQL: Opening connection to $($ScriptParameters.ServerInstance)..."
        $conn.Open()
        Write-Log -Level INFO -Message "      SQL: Connection opened."
        
        $cmd = New-Object System.Data.SqlClient.SqlCommand($currentSqlQuery, $conn)
        Write-Log -Level INFO -Message "      SQL: Executing query for $orgName..."
        $reader = $cmd.ExecuteReader()

        $columnNames = @()
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $columnNames += $reader.GetName($i)
        }
        $headerLine = ($columnNames -join ',')

        Out-File -FilePath $csvFilePath -Encoding UTF8 -InputObject $headerLine -Force -ErrorAction Stop

        while ($reader.Read()) {
            $recordCount++
            $lineValues = @()
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $value = $reader.GetValue($i)
                $formattedValue = ""
                if ($value -ne [System.DBNull]::Value) {
                    $formattedValue = $value.ToString()
                }
                $lineValues += $formattedValue
            }
            $csvLine = ($lineValues -join ',')
            Add-Content -Path $csvFilePath -Value $csvLine -Encoding UTF8 -ErrorAction Stop
        }

        Write-Log -Level INFO -Message "      SUCCESS: Found and exported $($recordCount.ToString("N0")) $logAction to $csvFilePath"

        # SFTP Upload with Retry Logic
        if ($ScriptParameters.EnableSftpUpload -and $sftpPrereqsMet) {
            if ($recordCount -gt 0) {
                $sftpSuccess = $false
                $maxRetries = 2
                for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
                    Write-Log -Level INFO -Message "      SFTP: Attempting upload ($attempt of $maxRetries) for '$csvFileName'..."
                    
                    $sftpSuccess = Start-SftpUpload -LocalFilePath $csvFilePath -RemoteFileName $csvFileName -Username $sftpUsername -Password $sftpPassword -GlobalHostname $ScriptParameters.SftpGlobalHostname -GlobalHostKeyFingerprint $ScriptParameters.SftpGlobalHostKeyFingerprint -RemoteDirectoryPath $sftpRemoteDirectory -OrganizationName $orgName
                    
                    if ($sftpSuccess) {
                        # If upload succeeds, break out of the retry loop
                        break
                    }
                    
                    if ($attempt -lt $maxRetries) {
                        Write-Log -Level WARN -Message "      SFTP: Upload failed on attempt $attempt. Will retry after 5 seconds..."
                        Start-Sleep -Seconds 5
                    }
                } # End of retry loop

                # Final status check after all attempts
                if ($sftpSuccess) {
                    Write-Log -Level INFO -Message "      SUCCESS: SFTP upload for '$csvFileName' completed successfully."
                } else {
                    # The final failure after all retries is logged as a critical ERROR.
                    # This ensures the script stops and flags a critical error.
                    Write-Log -Level ERROR -Message "      SFTP CRITICAL FAILURE: Upload for '$csvFileName' did not succeed after all $maxRetries attempts."
                }
            } else {
                Write-Log -Level INFO -Message "      SFTP: Skipping upload for Org: $orgName as no records were exported."
            }
        }
    } catch {
        Write-Log -Level ERROR -Message "      FAILURE: Failed to process Org: $orgName (ID: $orgID). Error: $($_.Exception.Message)"
        if ($_.Exception.InnerException) {
            Write-Log -Level ERROR -Message "      Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    finally {
        if ($reader -ne $null -and -not $reader.IsClosed) { $reader.Close() }
        if ($conn -ne $null -and $conn.State -eq [System.Data.ConnectionState]::Open) { $conn.Close() }
    }
}
#endregion

# --- Script Start ---
Write-Log -Level INFO -Message "========== Script execution started. =========="
# --- ADDED: Log the reference date being used for the entire run ---
Write-Log -Level INFO -Message "Using SQL reference date for this run: $($QueryReferenceDate.ToString('yyyy-MM-dd'))"

#region Initial Checks
# (This section remains unchanged)
Write-Log -Level INFO -Message "Performing initial checks..."
if (-not (Test-Path -Path $OutputPath)) {
    Write-Log -Level WARN -Message "Output path '$OutputPath' not found. Creating it..."
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

if (-not (Test-Path -Path $SettingsFilePath)) {
    Write-Log -Level FATAL -Message "Settings file not found at '$SettingsFilePath'. Please check the path or provide the -SettingsFilePath parameter."
    return
}

if (-not (Test-Path -Path $SqlFilePathAdd)) {
    Write-Log -Level FATAL -Message "SQL query file for ADD operations not found at '$SqlFilePathAdd'. Please check the path or provide the -SqlFilePathAdd parameter."
    return
}

if (-not (Test-Path -Path $SqlFilePathDelete)) {
    Write-Log -Level FATAL -Message "SQL query file for DELETE operations not found at '$SqlFilePathDelete'. Please check the path or provide the -SqlFilePathDelete parameter."
    return
}

if ($EnableSftpUpload) {
    if (-not (Test-Path -Path $WinSCPPath)) {
        Write-Log -Level FATAL -Message "WinSCPnet.dll not found at '$WinSCPPath'. Please check the path or provide the -WinSCPPath parameter."
        return
    }
    if (-not $SftpGlobalHostKeyFingerprint) {
        Write-Log -Level FATAL -Message "SFTP uploads are enabled, but -SftpGlobalHostKeyFingerprint was not provided. This is required."
        return
    }
    try {
        Add-Type -Path $WinSCPPath -ErrorAction Stop
        Write-Log -Level INFO -Message "WinSCP .NET assembly loaded successfully from '$WinSCPPath'."
    }
    catch {
        Write-Log -Level FATAL -Message "Failed to load WinSCP .NET assembly from '$WinSCPPath'. Error: $($_.Exception.Message)"
        return
    }
}
#endregion

#region Load SQL Query Templates
# (This section remains unchanged)
Write-Log -Level INFO -Message "Loading SQL query templates..."
try {
    $sqlQueryTemplateAdd = Get-Content -Raw -Path $SqlFilePathAdd -ErrorAction Stop
    $sqlQueryTemplateDelete = Get-Content -Raw -Path $SqlFilePathDelete -ErrorAction Stop
    Write-Log -Level INFO -Message "SQL query templates loaded successfully."
}
catch {
    Write-Log -Level FATAL -Message "Failed to read a SQL query file. Error: $($_.Exception.Message)"
    return
}
#endregion

# --- Main Processing ---
Write-Log -Level INFO -Message "Starting to process organizations from settings file..."

try {
    $orgSettings = Get-Content -Raw -Path $SettingsFilePath | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Log -Level FATAL -Message "Failed to read or parse JSON settings file '$SettingsFilePath'. Error: $($_.Exception.Message)"
    return
}

if ($null -eq $orgSettings -or $orgSettings.Count -eq 0) {
    Write-Log -Level WARN -Message "No settings found in the JSON file or the file is empty. Exiting."
    return
}

# --- ADDED: Filter for a specific Organization ID if provided ---
if (-not [string]::IsNullOrEmpty($OrganizationID)) {
    Write-Log -Level INFO -Message "Parameter -OrganizationID specified. Filtering for Org ID: $OrganizationID"
    $orgSettings = $orgSettings | Where-Object { $_.OrganizationID -eq $OrganizationID }

    if ($null -eq $orgSettings -or $orgSettings.Count -eq 0) {
        Write-Log -Level WARN -Message "No organization with ID '$OrganizationID' found in the settings file. Exiting."
        return
    }
}

$scriptParameters = @{
    RunDate                     = (Get-Date).ToString('yyyyMMdd')
    FormattedQueryReferenceDate = $QueryReferenceDate.ToString("yyyy-MM-dd")
    ServerInstance              = $ServerInstance
    Database                    = $Database
    OutputPath                  = $OutputPath
    EnableSftpUpload            = $EnableSftpUpload
    EnableSqlLogging            = $EnableSqlLogging
    EnableRawSqlDataLogging     = $EnableRawSqlDataLogging
    SftpGlobalHostname          = $SftpGlobalHostname
    SftpGlobalHostKeyFingerprint= $SftpGlobalHostKeyFingerprint
}


foreach ($setting in $orgSettings) {
    if ([string]::IsNullOrEmpty($setting.OrganizationID) -or [string]::IsNullOrEmpty($setting.FilenamePrefix) -or [string]::IsNullOrEmpty($setting.OrganizationName)) {
        Write-Log -Level WARN -Message "Skipping setting due to missing 'OrganizationID', 'FilenamePrefix', or 'OrganizationName'. Details: $($setting | Out-String)"
        continue
    }

    Write-Log -Level INFO -Message "--- Processing Org: $($setting.OrganizationName) (ID: $($setting.OrganizationID)) ---"

    if (-not [string]::IsNullOrEmpty($setting.CollectionIDOCLCAdd)) {
        Process-OclcOperation -OperationType "ADD" -CollectionID $setting.CollectionIDOCLCAdd -SqlQueryTemplate $sqlQueryTemplateAdd -Setting $setting -ScriptParameters $scriptParameters
    }
    else {
        Write-Log -Level WARN -Message "Skipping ADD operation for Org: $($setting.OrganizationName) as 'CollectionIDOCLCAdd' is not defined."
    }

    if (-not [string]::IsNullOrEmpty($setting.CollectionIDOCLCDel)) {
        Process-OclcOperation -OperationType "DELETE" -CollectionID $setting.CollectionIDOCLCDel -SqlQueryTemplate $sqlQueryTemplateDelete -Setting $setting -ScriptParameters $scriptParameters
    }
    else {
        Write-Log -Level WARN -Message "Skipping DELETE operation for Org: $($setting.OrganizationName) as 'CollectionIDOCLCDel' is not defined."
    }
}

# --- Error Summary Section ---
Write-Log -Level INFO -Message "------------------------------------------------------"
if ($Global:ErrorCollection.Count -gt 0) {
    Write-Log -Level ERROR -Message "========== SCRIPT COMPLETED WITH ERRORS =========="
    Write-Log -Level ERROR -Message "The following $($Global:ErrorCollection.Count) critical errors occurred during execution:"
    foreach ($err in $Global:ErrorCollection) {
        # The message is already formatted with timestamp and level, so we just output it.
        # We use Write-Host directly to avoid re-logging it via Write-Log.
        Write-Host $err -ForegroundColor Red
    }
    Write-Log -Level ERROR -Message "========== END OF ERROR SUMMARY =========="
}
else {
    Write-Log -Level INFO -Message "========== SCRIPT COMPLETED SUCCESSFULLY WITH 0 CRITICAL ERRORS =========="
}

Write-Log -Level INFO -Message "========== Script processing finished. =========="
