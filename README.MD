
# Unified Audit Log Analyzer


## Table of Contents


- [Unified Audit Log Analyzer](#unified-audit-log-analyzer)
  - [Unified Audit Log Analyzer (GUI)](#unified-audit-log-analyzer-gui)
  - [Expand-UnifiedAuditLogReport (non-GUI)](#expand-unifiedauditlogreport-non-gui)




## Unified Audit Log Analyzer (GUI)

The **Unified Audit Log Analyzer** is a powerful PowerShell-based graphical user interface (GUI) tool designed to help administrators analyze and visualize Microsoft 365 (M365) Unified Audit Logs. This tool provides a user-friendly interface for loading, filtering, and exporting audit log data in various formats (CSV, JSON). It supports drag-and-drop functionality, advanced filtering, hierarchical data exploration, and customizable themes.

## **Run the Script**:

- Open PowerShell and navigate to the directory containing the script. There are two ways of run the program/script

- Run the script using the following command without parameter

    ```powershell
    .\UnifiedAuditLogAnalyzer.ps1
    ```

- Example command to load data via `InputData` parameter:

    - The parametric version can take full file path `"C:\AuditLogs\logdata.csv"` 

        ```powershell
        .\UnifiedAuditLogAnalyzer.ps1 -InputData "C:\AuditLogs\logdata.csv"
        ```

    - The parametric version can take full file path `"C:\AuditLogs\logdata.csv"`

        ```powershell
        $log = Import-csv  "C:\AuditLogs\logdata.csv"
        .\UnifiedAuditLogAnalyzer.ps1 -InputData $log
        ```

---

## Expand-UnifiedAuditLogReport (non-GUI)

`Expand-UnifiedAuditLogReport` is a PowerShell script designed to process and analyze audit logs from Microsoft 365 Unified Audit Log data. The script can handle audit logs in two ways: This script was developed to streamline audit log analysis and reporting within Microsoft 365 environments.

1. **Using locally provided audit logs**

2. **Retrieving audit data directly from a source**

The script expands JSON-based audit log entries, extracts key parameters, and generates detailed reports.

## Features

- Supports processing of audit logs from local data files or direct retrieval from Microsoft 365.

- Converts JSON-based audit log data into an easily readable format.

- Extracts and structures detailed parameter information.

- Generates CSV reports for audit logs and modified property reports.

- Supports filtering based on audit record types, operations, date ranges, and free text searches.

- Provides an option to prioritize completeness over performance.

- Enables session-based data retrieval for handling large datasets.


## Examples

### Example 1: Using Local Data

```powershell
$Unifiedlogs = Search-UnifiedAuditLog -EndDate 3/15/2025 -StartDate 3/1/2025 -ResultSize 500 -RecordType ExchangeAdmin
Expand-UnifiedAuditLogReport -AuditLogs $Unifiedlogs -CreateAuditModifiedPropertyReport -ReportPath "C:\Reports\AuditReport.csv"
```

You are directly getting auditlog from server

### Example 2: Using Local Data

```powershell
$Unifiedlogs = Import-csv "ExistingM365UnifiedAuditfile.csv"
Expand-UnifiedAuditLogReport -AuditLogs $logs -CreateAuditModifiedPropertyReport -ReportPath "C:\Reports\AuditReport.csv"
```