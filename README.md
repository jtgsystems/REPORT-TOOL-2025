# System Report Tool 2025

## Description

This project provides a PowerShell script designed to generate a comprehensive
system report, providing detailed information about the system's hardware,
software, and configuration.

## Features

- **Administrator Privilege Check:** Ensures the script runs with administrator
  privileges, restarting itself with elevation if necessary.
- **Comprehensive System Report Generation:** Gathers and presents a wide range
  of system information.
- **Report Sections:**
  - Windows Update Status (pending updates and last installed update)
  - System Information (manufacturer, model, OS name, OS version, processor,
    BIOS)
  - Memory Information (total RAM, memory slots)
  - Disk Information (drive information, size, free space)
  - Graphics Information (video controller information)
  - Network Information (network adapter information)
  - Power and Battery Information
  - System Uptime
  - Startup Programs
  - Installed Software
  - Finds and lists large files (larger than 0.1 GB)
- **Output:** Saves the report to a text file on the user's desktop.
- **Speech Output:** Provides speech output for key information, such as the
  report generation status and key findings.
- **User Prompt:** Prompts the user to open the generated report.

## Dependencies

- PowerShell
- System.Speech (for speech output)

## Installation

1. Ensure PowerShell is installed on your system. (It is typically pre-installed
   on Windows systems.)
2. Save the `SYSTEM-REPORT-TOOL.ps1` script to a directory of your choice.

## Usage

1. Open PowerShell as an administrator. (The script will attempt to elevate
   itself if run without admin privileges.)
2. Navigate to the directory where you saved the script.
3. Run the script using the command: `.\SYSTEM-REPORT-TOOL.ps1`.
4. The script will generate a comprehensive system report and save it to a text
   file named "COMPREHENSIVE SYSTEM REPORT - by JTG Systems.txt" on your
   desktop.
5. The script will provide speech output and prompt you to open the report.

## Configuration

The following variables can be modified within the script to customize its
behavior:

- `$outputFile`: Specifies the path to the output report file. Defaults to
  `$env:USERPROFILE\Desktop\COMPREHENSIVE SYSTEM REPORT - by JTG Systems.txt`
  (your desktop).
- `$minSizeGB`: Specifies the minimum size (in GB) for files to be considered
  "large" in the large files report. Defaults to `0.1`.

## Attribution

Credit: [https://www.jtgsystems.com](https://www.jtgsystems.com)
