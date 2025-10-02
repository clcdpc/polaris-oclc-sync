# OCLC WorldCat Data Sync PowerShell Script

## Overview

This PowerShell script is designed to automate the process of synchronizing library holdings from a Polaris SQL database with OCLC's WorldCat. It queries the Polaris database for newly added or deleted bibliographic records, formats the data into CSV files, and securely uploads them to OCLC's SFTP server for processing by WorldShare Collection Manager.
This script is intended to be run as an automated job, typically using a SQL Server Agent.

**Compatibility:** This script has been tested with **Polaris ILS version 7.8**. 

## Features

- **Parameterized:** Key configuration details like file paths, server instances, and dates can be passed as parameters.
- **Centralized Pathing:** Uses a single `$BasePath` parameter to manage the locations of all necessary files (settings, logs, SQL queries, etc.).
- **Automated CSV Generation:** Queries the database for records to add and delete based on a reference date and exports them to separate CSV files.
- **Secure SFTP Upload:** Uses WinSCP .NET assembly to upload the generated CSV files to OCLC's SFTP server.
- **Robust Logging:** Creates daily log files and provides different log levels for clear monitoring and troubleshooting.
- **Retry Logic:** Includes built-in retry logic for SFTP uploads to handle transient network issues.
- **Consortium Support:** It is designed to work in a consortium enviornment where each member library has their own OCLC Metadata Collection/Account.
  - When you are starting the intial syncs, the script supports passing a parameter where only ONE of the orgs in the settings.json will be run. This allows you to for example rollout the sync process gradually and have some members doing a daily sync alongside those you're doing an intial sync with.

## Prerequisites

1.  **Windows Server:** A server with PowerShell 5.1 or higher.
2.  **Polaris SQL Server Access:** Read access to the Polaris database. The script uses Windows Authentication to connect.
3.  **WinSCP:** The `WinSCPnet.dll` assembly is required for SFTP functionality. You can download it from the [WinSCP website](https://winscp.net/eng/downloads.php). Place the `.dll` in the path specified by the `$WinSCPPath` parameter (defaults to `<BasePath>\WinSCPnet.dll`).
4.  **Good OCLC numbers in your Polaris MARC records**. This process does a CSV match with OCLC.
5.  An approved non-MARC Sync Collection both ADDING and DELETING holdings. Check with OCLC support as needed.
6.  OCLC file transfer credentials which you must obtain from OCLC support and are different from your WorldShare Collection Manager login.
7.  Although the SQL scripts will delete holdings if all item records are in these statuses: 7=Lost, 8=Claim Returned, 9=Claim Never Had, 10=Missing, 11=Withdrawn, 16=Unavailable, 20=Does Not Circulate, 21=Claim Missing Parts, **it is recommended that you enable the Polaris feature to RETAIN deleted item records**. The SQL script will also look for deleted item record statuses and using deleted item records means you don't have to be as careful with your timing remembering to run the script while items are in a Withdrawn status before deleting the record entirely.

## Configuration

### Script Parameters

The script's behavior is controlled by parameters. The most important one is `$BasePath`, which defaults to `c:\ProgramData\clc_oclc_sync`. All other file paths are derived from this base path unless explicitly overridden.

### `settings.json` File

This is the main configuration file, located by default at `<BasePath>\settings.json`. It contains a JSON array of objects, where each object represents an organization or library to be processed.

**Example `settings.json`:**
```json
[
  {
    "OrganizationID": "101",
    "OrganizationName": "Springfield Public Library",
    "FilenamePrefix": "springfield_pl",
    "SftpUsername": "sftp_user_springfield",
    "SftpPassword": "secure_password_1",
    "SftpRemoteDirectory": "/oclc/uploads/springfield",
    "CollectionIDOCLCAdd": "123456789",
    "CollectionIDOCLCDel": "987654321"
  }
]
```

### SQL Query Files

The script uses two SQL files (`NewOCLCRecords.sql` and `DeleteOCLCRecords.sql`) to fetch the data. These files must contain a placeholder `{0}` for the `OrganizationID` and `{1}` for the `$QueryReferenceDate`.

## Setup as a SQL Server Agent Job

To automate the script, you can schedule it as a job in the SQL Server Agent.

1.  **Open SQL Server Management Studio (SSMS)** and connect to your database engine.
2.  **Navigate to SQL Server Agent** > **Jobs**. Right-click **Jobs** and select **New Job...**.
3.  **General Tab:**
    *   Give the job a descriptive name (e.g., "OCLC Daily Sync").
    *   Set the **Owner** to a service account with appropriate permissions.
4.  **Steps Tab:**
    *   Click **New...** to create a new job step.
    *   **Step Name:** "Run OCLC Sync Script".
    *   **Type:** Select **PowerShell**.
    *   **Run as:** Choose a credential that has permissions to execute PowerShell scripts and access the Polaris database.
    *   **Command:** Enter the command to execute the script.
        ```powershell
        # Example: Run with default settings
        C:\path\to\your\scripts\oclc-sync.ps1

        # Example: Override the BasePath
        C:\path\to\your\scripts\oclc-sync.ps1 -BasePath "D:\OCLC_Sync"
        ```
5.  **Schedules Tab:**
    *   Click **New...** to create a schedule.
    *   Configure the frequency (e.g., daily) and time for the job to run.

## OCLC Documentation

For more information on OCLC's data sync process and SFTP requirements, please refer to their official documentation:

-   **About Data Sync Collections:** [OCLC Support - Data Sync Collections](https://help.oclc.org/Metadata_Services/WorldShare_Collection_Manager/Data_sync_collections)
-   **File Upload/Download Information:** [OCLC Support - Upload and Download Files](https://help.oclc.org/Metadata_Services/WorldShare_Collection_Manager/Get_started/Upload_and_download_files)
-   **SFTP Client Setup:** [OCLC Support - Upload Files with an SFTP Client](https://help.oclc.org/Librarian_Toolbox/Exchange_files_with_OCLC/Upload_files_with_SFTP/20SFTP_client)

## Common Troubleshooting Issues

-   **SQL Connection Fails:**
    *   **Cause:** The account running the PowerShell script does not have permission to access the Polaris database.
    *   **Solution:** Ensure the 'Run as' account for the SQL Server Agent job step has the necessary database roles (e.g., `db_datareader`).
-   **SFTP Upload Fails:**
    *   **Cause:** Incorrect SFTP credentials, firewall blocking the connection, or wrong host key fingerprint.
    *   **Solution:** Verify all SFTP parameters (`SftpUsername`, `SftpPassword`, `SftpGlobalHostKeyFingerprint`). Check the script's log files for detailed error messages from WinSCP.
-   **File Path Not Found:**
    *   **Cause:** The script cannot find required files like `settings.json` or `WinSCPnet.dll`.
    *   **Solution:** Ensure your `BasePath` is correct and that all required files are in their respective locations within that path. Check the log files for "FATAL" messages indicating which file is missing.
-   **Script Execution Policy:**
    *   **Cause:** PowerShell's execution policy prevents the script from running.
    *   **Solution:** You may need to set the execution policy. The SQL Server Agent job for PowerShell often bypasses this, but for manual testing, you might run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`.
---
