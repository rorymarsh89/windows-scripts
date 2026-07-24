# PCHH Triage

## Description
PCHH Triage is a PowerShell script (`triage.ps1`) that collects system specifications, crash logs, security posture, and diagnostics. It compiles everything into a single compressed ZIP archive on the desktop, alongside an interactive HTML report for fast triage.

## Prerequisites
* Windows Operating System.
* Administrator privileges.

## Data Collected

**System Specifications**
CPU (model, speed, cores/threads), GPU(s) (model, accurate VRAM, friendly driver version), Motherboard, BIOS (version, date), TPM status/version, Secure Boot, Fast Boot, Windows edition/build/install date, uptime, page file size, boot device, Active Power Plan, UAC status, default audio playback device, Device Manager errors, installed Windows updates (hotfixes).

**GPU & Displays**
Every graphics adapter with driver version (NVIDIA/AMD/Intel translated to vendor version numbers, e.g. 566.36 instead of the raw WMI string), Hardware-accelerated GPU Scheduling (HAGS) status, and every connected display mapped to its adapter with resolution, refresh rate, and bit depth.

**Memory (RAM)**
Per-stick slot, manufacturer, part number, capacity, and rated vs. configured speed (flags when XMP/EXPO isn't enabled). Physical memory usage and commit charge at time of capture.

**Storage Diagnostics**
Drive capacity, partition/volume data, dirty bit status per volume, and SMART health — both Windows' reliability counters and raw ATA attributes (reallocated sectors, pending sectors, uncorrectable sectors, UltraDMA CRC errors, command timeouts, and the drive's own failure prediction flag).

**Reliability History**
Full Windows Reliability Monitor history (~1 year), exported as CSV and shown as an interactive, scrollable timeline categorised into Critical/Warning/Informational events.

**Event Viewer (curated)**
A targeted set of System log events from the last 12 months (matching the reliability history window): Kernel-Power 41 (with bugcheck code extraction), BugCheck 1001, WHEA hardware errors, disk I/O errors, storage controller resets, filesystem corruption, GPU driver timeouts (TDR), service crashes, memory diagnostic results, and unexpected shutdowns. The full System event log is also exported in full (`system_eventlogs.evtx`) for deeper analysis.

**Crash Logs**
Minidump (`.dmp`) files from the last 60 days, listed in the report with filename, date, and size.

**Security**
Windows Defender status (real-time protection, last scan times, signature age) and threat detection history; Defender exclusions with specific flags for risky patterns (whole-drive, broad system folder, or executable-extension exclusions); Firewall status per profile; hosts file integrity (custom entry count plus flags for redirected update/security domains); startup persistence (Run keys, non-Microsoft Scheduled Tasks, Startup folder shortcuts) flagged when running from Temp or failing signature validation; services set to Automatic but not running; and a check for conflicting real-time antivirus products running simultaneously.

**Networking**
Physical network adapters (status, link speed, media type), Wi-Fi signal strength/band/channel/radio type/negotiated rates, and detection of active VPN or virtual network adapters. No IP addresses, SSIDs, MAC addresses, or connection endpoints are collected.

**Apps**
Running processes (sortable by name, instance count, or memory use), installed programs, installed Windows updates, and browser extensions across Chrome, Edge, Brave, Opera, Opera GX, Vivaldi, and Firefox.

**Known Software Detection**
Flags common anti-cheat/kernel drivers (Vanguard, Easy Anti-Cheat, BattlEye, FACEIT AC), overclocking/monitoring tools (MSI Afterburner, RTSS, Intel XTU, Ryzen Master), RGB/peripheral suites (iCUE, Synapse, G HUB, Armoury Crate, Mystic Light, Aura Sync, NZXT CAM, MSI Dragon Center), audio/overlay software (Nahimic, GeForce Experience, Xbox Game Bar, Streamlabs OBS), problematic network software (Hola VPN, Killer Network Manager), and potential bloatware/PUPs (McAfee, Norton, WildTangent, IObit utilities, Restoro, PC "cleaner" utilities, third-party driver updaters). Also fingerprints GPU driver crashes (TDR events, video-related bugcheck codes, LiveKernelEvents) so display driver instability is called out directly.

**Privacy**
No username, hostname, IP address, SSID, or MAC address is collected anywhere in the report.

## Usage
1. Open PowerShell or Windows Terminal as Administrator.
2. Run `triage.ps1`.
3. Wait for the process to complete (collecting GPU/display data via DXDIAG can take up to 30 seconds).
4. The archive is saved to `Desktop\PCHH-Triage` and automatically copied to the clipboard.

## Output
The script generates `PCHH-Triage_<random_number>.zip` containing:
* `specs-programs.txt` — Hardware and software specifications.
* `system_eventlogs.evtx` — Full Windows System event log export.
* `reliability.csv` — Reliability history.
* `triage-report.html` — Interactive report with a sidebar covering System Summary, Reliability History, Event Viewer, Drives, GPU and Display(s), Memory (RAM), Network, Security, Running Processes, Installed Apps, Browser Extensions, and Memory Dumps (when present).
* Relevant `.dmp` files.

Nothing is left loose on the desktop — everything ships inside the single zip.
