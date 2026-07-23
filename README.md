# PCHH System Triage

## Description

PCHH Triage is a PowerShell script (`triage.ps1`) that collects system specifications, crash logs, and diagnostics. It compiles this data into a single compressed ZIP archive on the desktop for troubleshooting and analysis.

## Prerequisites

* Windows Operating System.


* Administrator privileges.



## Data Collected

* **System Specifications:** CPU, GPU, Motherboard, BIOS, Memory, and Operating System details.


* **Storage Diagnostics:** Drive capacity, partition data, and SMART health metrics.


* **Crash Logs:** Minidump files (`.dmp`) from the last 60 days.


* **Event Logs:** System event logs (EVTX) from the last 14 days, filtered for critical hardware, storage, and GPU errors.


* **Reliability Records:** Windows reliability history.


* **Network & Processes:** Network adapter states, Wi-Fi signal strength, and running processes.


* **Installed Software:** List of installed applications.



## Usage

1. Open PowerShell or Windows Terminal as Administrator.


2. Run `triage.ps1`.


3. Wait for the process to complete.


4. The archive is saved to `Desktop\PCHH-Triage` and automatically copied to the clipboard.



## Output

The script generates `PCHH-Triage_<random_number>.zip` containing:

* `specs-programs.txt`: Hardware and software specifications.


* `system_eventlogs.evtx`: Windows event logs.


* `reliability.csv`: Reliability history.


* `triage-report.html`: Interactive HTML report of the collected data.


* Relevant `.dmp` files.
