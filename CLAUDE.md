# System Report Tool 2025 - Claude Code Reference

## Project Overview

**System Report Tool 2025** is a comprehensive PowerShell diagnostic utility designed for IT repair technicians and system administrators. It generates detailed system reports covering hardware specifications, software inventory, Windows Update status, and large file analysis.

**Purpose**: Streamline computer diagnostics by automating the collection of critical system information into a single, human-readable report saved to the desktop.

**Attribution**: Created by [JTG Systems](https://www.jtgsystems.com)

---

## Key Features

### Comprehensive System Analysis
- **Windows Update Status**: Pending updates count and last installation date
- **Hardware Information**: Manufacturer, model, processor, BIOS, RAM, disk drives
- **Memory Diagnostics**: Total RAM, memory slots (used/empty)
- **Graphics Information**: Video controller details and VRAM
- **Network Analysis**: Active adapters with MAC addresses and link speeds
- **Power Management**: Active power plan and battery status
- **System Uptime**: Days/hours/minutes since last boot
- **Startup Programs**: Programs configured to run at system startup
- **Installed Software**: Complete inventory via WMI
- **Large Files Scanner**: Finds files exceeding configurable size threshold (default: 0.1 GB)

### User Experience Enhancements
- **Speech Synthesis**: Audible progress updates and key findings via System.Speech
- **Auto-Elevation**: Automatically requests administrator privileges if needed
- **Progress Reporting**: Real-time feedback during file scanning operations
- **Desktop Output**: Report saved as "COMPREHENSIVE SYSTEM REPORT - by JTG Systems.txt"
- **Interactive Prompt**: Option to open report immediately after generation

---

## Directory Structure

```
REPORT-TOOL-2025/
├── SYSTEM-REPORT-TOOL.ps1    # Main PowerShell script (571 lines)
├── README.md                  # Comprehensive documentation
├── banner.png                 # Repository banner (647 KB)
├── CLAUDE.md                  # This file
└── .git/                      # Git repository metadata
```

---

## Technology Stack

### Core Technologies
- **PowerShell**: Windows automation framework
- **WMI/CIM**: Windows Management Instrumentation for system queries
- **System.Speech**: Text-to-speech synthesis (.NET Framework)

### PowerShell Modules Used
- `Get-CimInstance`: Modern replacement for WMI queries (Win32_* classes)
- `Get-NetAdapter`: Network adapter information
- `Start-Job`: Background task parallelization
- `Get-HotFix`: Windows update history

### System Requirements
- Windows OS (Windows 7/Server 2008 R2 or later)
- PowerShell 5.1 or later (recommended)
- Administrator privileges (auto-requested)
- .NET Framework 3.5+ (for System.Speech)

---

## Architecture & Design

### Execution Flow
1. **Privilege Check**: Ensures administrator rights via `Ensure-Administrator` function
2. **Task Definition**: 10 data collection tasks defined as hashtable entries
3. **Large File Scan**: Executes first (long-running) with real-time progress
4. **Parallel Jobs**: Remaining tasks run concurrently via `Start-Job`
5. **Result Aggregation**: Waits for all jobs, retrieves output
6. **Report Generation**: Formats data into human-readable text report
7. **Output**: Saves to desktop, speaks summary, prompts user

### Performance Optimizations
- **Parallel Processing**: Background jobs collect data concurrently
- **Runspace Pooling**: Large file scanner uses runspace pool for multi-threading
- **Scoped File Search**: Limited to user profile directory (`$env:USERPROFILE`)
- **Exclusion Filters**: Skips AppData and system folders to reduce scan time
- **Top 50 Limit**: Returns only 50 largest files to prevent report bloat

### Error Handling Strategy
- **Try-Catch Blocks**: Every data collection task wrapped in error handling
- **Graceful Degradation**: Returns "N/A" or placeholder if data unavailable
- **Warning Messages**: Non-fatal errors logged via `Write-Warning`
- **Continued Execution**: `$ErrorActionPreference = "Continue"` prevents crashes
- **Job State Validation**: Checks job completion before retrieving results

---

## Development Workflow

### Configuration Variables
Located at script initialization (lines 37-39):

```powershell
$outputFile = "$env:USERPROFILE\Desktop\COMPREHENSIVE SYSTEM REPORT - by JTG Systems.txt"
$minSizeGB = 0.1  # Minimum file size for large files report
```

**Customization Points**:
- `$outputFile`: Change report save location
- `$minSizeGB`: Adjust large file threshold (in GB)
- `$excludedPathsRegex`: Modify excluded directories pattern (line 272)

### Running the Script

#### Method 1: PowerShell Console
```powershell
# Navigate to script directory
cd C:\path\to\REPORT-TOOL-2025

# Execute (will auto-elevate if needed)
.\SYSTEM-REPORT-TOOL.ps1
```

#### Method 2: Right-Click Context Menu
1. Right-click `SYSTEM-REPORT-TOOL.ps1`
2. Select "Run with PowerShell"
3. Allow UAC elevation prompt

#### Method 3: Scheduled Task
Create scheduled task for automated reports:
```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\path\to\SYSTEM-REPORT-TOOL.ps1"
$trigger = New-ScheduledTaskTrigger -Daily -At 9am
Register-ScheduledTask -TaskName "Daily System Report" -Action $action -Trigger $trigger -RunLevel Highest
```

### Execution Policy Requirements
If script won't run due to execution policy:
```powershell
# Check current policy
Get-ExecutionPolicy

# Temporarily bypass (session only)
PowerShell.exe -ExecutionPolicy Bypass -File .\SYSTEM-REPORT-TOOL.ps1

# Permanently set for current user
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Technical Deep Dive

### Data Collection Tasks

#### 1. WindowsUpdateStatus
- **API**: `Microsoft.Update.Session` COM object
- **Data**: Pending updates count, last installed update date
- **Performance**: Moderate (searches entire update catalog)

#### 2. SystemInfo
- **Classes**: `Win32_ComputerSystem`, `Win32_OperatingSystem`, `Win32_Processor`, `Win32_BIOS`
- **Data**: Manufacturer, model, OS, CPU cores/threads, BIOS version/date

#### 3. MemoryInfo
- **Classes**: `Win32_PhysicalMemoryArray`, `Win32_PhysicalMemory`
- **Data**: Total RAM (GB), total/used/empty memory slots
- **Calculation**: Sums capacity of all physical memory modules

#### 4. DiskInfo
- **Class**: `Win32_LogicalDisk` (DriveType=3 = Fixed disks)
- **Data**: Drive letter, total size (GB), free space (GB)
- **Format**: Table with DeviceID, SizeGB, FreeGB columns

#### 5. GraphicsInfo
- **Class**: `Win32_VideoController`
- **Data**: GPU name, VRAM (GB), driver version
- **Note**: May show integrated + dedicated GPUs

#### 6. NetworkInfo
- **Cmdlet**: `Get-NetAdapter`
- **Filter**: Only adapters with Status = "Up"
- **Data**: Name, description, MAC address, link speed

#### 7. PowerInfo
- **Commands**: `powercfg /GETACTIVESCHEME`, `Get-CimInstance Win32_Battery`
- **Data**: Active power plan name, battery status/charge percentage
- **Logic**: Regex extracts GUID, maps to friendly name

#### 8. SystemUptime
- **Class**: `Win32_OperatingSystem`
- **Calculation**: Current time - `LastBootUpTime`
- **Format**: "X days, Y hours, Z minutes"

#### 9. StartupPrograms
- **Class**: `Win32_StartupCommand`
- **Data**: Program name, command line, registry location

#### 10. InstalledSoftware
- **Class**: `Win32_Product`
- **Data**: Name, version, vendor (sorted alphabetically)
- **Warning**: Can be slow (triggers Windows Installer consistency check)

#### 11. FindLargeFiles (Optimized Scanner)
- **Method**: Parallel runspace pool processing
- **Scope**: `$env:USERPROFILE` directory
- **Exclusions**: AppData, Local Settings, Application Data
- **Workers**: `[Environment]::ProcessorCount` (scales with CPU cores)
- **Output**: Top 50 largest files with size, path, last modified date

### Speech Synthesis Implementation

```powershell
function Speak-Text {
  param ([string]$text)
  try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $synth.Rate = 2  # 2x speed for efficiency
    $synth.Speak($text)
  }
  catch {
    Write-Warning "Speech synthesis not available: $($_.Exception.Message)"
  }
  finally {
    if ($synth) { $synth.Dispose() }
  }
}
```

**Key Points**:
- Graceful fallback if System.Speech unavailable
- Increased speech rate (2x) for faster technician workflow
- Proper disposal of synthesizer object

---

## Testing Approach

### Manual Testing Checklist
- [ ] Run on Windows 10/11 systems
- [ ] Run on Windows Server 2016/2019/2022
- [ ] Test with standard user account (verify auto-elevation)
- [ ] Test with administrator account
- [ ] Verify all report sections populate correctly
- [ ] Confirm large files scanner finds known large files
- [ ] Check speech synthesis on systems with/without audio
- [ ] Validate report saves to desktop
- [ ] Test "Open file" prompt (Y/N responses)

### Edge Cases
- Systems without battery (desktop PCs)
- Minimal installations (few installed programs)
- Systems with execution policy restrictions
- Virtual machines (may have limited hardware info)
- Systems with pending reboots (Windows Update data)

### Known Data Limitations
- **Win32_Product**: Slow and incomplete (misses per-user installs)
- **Battery Status**: May show "Unknown" on some hardware
- **Large Files**: Skips inaccessible files (permissions denied)
- **Windows Update**: Requires COM object support (may fail on Server Core)

---

## Performance Considerations

### Execution Time
- **Typical Runtime**: 30-90 seconds
- **Bottlenecks**: Large file scan (60-80% of total time), installed software query
- **Optimization**: Parallel jobs reduce sequential wait time by ~40%

### Resource Usage
- **CPU**: Multi-core utilization during file scan (runspace pool)
- **Memory**: ~50-150 MB (depends on file system size)
- **Disk I/O**: High during large file scan (sequential reads)

### Scaling Factors
- **User Profile Size**: Larger profiles increase scan time exponentially
- **Number of Files**: 100K+ files can take 2-5 minutes
- **Installed Programs**: 100+ programs adds 10-20 seconds
- **Disk Count**: More drives = longer disk info query

---

## Known Issues & Troubleshooting

### Issue: Script Won't Run (Execution Policy)
**Symptoms**: Red error about execution policy
**Solution**: Run with bypass flag or change policy
```powershell
PowerShell.exe -ExecutionPolicy Bypass -File .\SYSTEM-REPORT-TOOL.ps1
```

### Issue: "Access Denied" During File Scan
**Symptoms**: Warning messages about inaccessible folders
**Solution**: Normal behavior - script continues with accessible files

### Issue: Installed Software Section Empty
**Symptoms**: "Installed software information unavailable"
**Root Cause**: Win32_Product query failure (Windows Installer service)
**Workaround**: Check Task Manager > Services > Windows Installer (msiserver) status

### Issue: Speech Not Working
**Symptoms**: No audible feedback, warnings about System.Speech
**Solution**: Install .NET Framework 3.5 via Windows Features
```powershell
Enable-WindowsOptionalFeature -Online -FeatureName NetFx3 -All
```

### Issue: Slow Large File Scan
**Symptoms**: File scan takes 5+ minutes
**Solutions**:
1. Increase `$minSizeGB` threshold (e.g., 0.5 GB instead of 0.1 GB)
2. Modify `$excludedPathsRegex` to skip more directories
3. Reduce `Select-Object -First 50` to `First 20`

### Issue: Report Not Saving
**Symptoms**: Script completes but no file on desktop
**Check**:
1. Desktop path: `$env:USERPROFILE\Desktop` resolves correctly
2. Disk space: Desktop drive has available space
3. Permissions: User can write to desktop folder

---

## Git Repository Information

### Repository Details
- **GitHub URL**: https://github.com/jtgsystems/REPORT-TOOL-2025
- **Default Branch**: main
- **Clone Command**: `gh repo clone jtgsystems/REPORT-TOOL-2025`

### Recent Commits (Last 20)
Notable development milestones:
- **0c298ce**: Add banner to README
- **692bbd8**: Refactor file scanning logic (improved object creation)
- **2fca12f**: Enhance power plan extraction (regex improvements)
- **9e87ddc**: Exclude user profile system paths from scan
- **f0c78ca**: Add robust data acquisition methodology to README
- **22b156e**: Initial README with full documentation

### Contributing
Contributions should follow PowerShell best practices:
- Use approved verbs (`Get-`, `Set-`, `New-`, etc.)
- Include try-catch error handling
- Add comments for complex logic
- Test on multiple Windows versions
- Update README.md with new features

---

## Future Roadmap

### Planned Features
1. **HTML Report Option**: Color-coded output with charts
2. **Email Integration**: Auto-send reports via SMTP
3. **Historical Tracking**: Compare reports over time
4. **Registry Analysis**: Scan for common issues
5. **Event Log Parsing**: Recent errors/warnings
6. **Driver Version Check**: Flag outdated drivers
7. **Security Audit**: Firewall status, antivirus state
8. **Cloud Upload**: Auto-backup reports to OneDrive/SharePoint

### Code Improvements
- Replace `Win32_Product` with registry-based software detection
- Add progress bars for all long-running tasks
- Implement JSON export format for automation
- Add `-Silent` parameter (no speech output)
- Create GUI version with WPF

### Performance Enhancements
- Cache WMI queries for repeated data access
- Implement incremental file scanning (skip unchanged directories)
- Add multithreading to installed software query
- Optimize memory usage for systems with 1M+ files

---

## Quick Reference Commands

### Check PowerShell Version
```powershell
$PSVersionTable.PSVersion
```

### Run Script Silently (No Speech)
Comment out lines 42, 555-556 or modify `Speak-Text` function

### Generate Report on Remote Computer
```powershell
Invoke-Command -ComputerName REMOTE-PC -FilePath .\SYSTEM-REPORT-TOOL.ps1
```

### Export Report to Network Share
Change line 38:
```powershell
$outputFile = "\\SERVER\Share\Reports\$(env:COMPUTERNAME)-$(Get-Date -Format 'yyyyMMdd').txt"
```

### Customize Large File Threshold
Change line 39:
```powershell
$minSizeGB = 1.0  # Only files > 1 GB
```

---

## Security Considerations

### Administrator Privileges
- Script requires admin rights for complete WMI access
- Auto-elevation prompt ensures user consent via UAC
- No credentials stored or transmitted

### Data Privacy
- Report contains hardware/software inventory (non-sensitive)
- No personal files accessed (only metadata)
- Local-only operation (no network transmission)

### Safe Deployment
- Read-only operations (no system modifications)
- No registry changes or file deletions
- Can be run repeatedly without side effects

---

## Support & Resources

### Documentation
- **README.md**: Comprehensive usage guide with examples
- **CLAUDE.md**: This file - development reference
- **Code Comments**: Inline explanations throughout script

### Attribution
- **Website**: https://www.jtgsystems.com
- **GitHub**: https://github.com/jtgsystems
- **Repository**: https://github.com/jtgsystems/REPORT-TOOL-2025

### External Resources
- [PowerShell Documentation](https://docs.microsoft.com/powershell)
- [WMI Reference](https://docs.microsoft.com/windows/win32/wmisdk)
- [System.Speech Namespace](https://docs.microsoft.com/dotnet/api/system.speech.synthesis)

---

## Development Environment Setup

### Local Development
1. Clone repository:
   ```powershell
   gh repo clone jtgsystems/REPORT-TOOL-2025
   cd REPORT-TOOL-2025
   ```

2. Edit script:
   ```powershell
   code SYSTEM-REPORT-TOOL.ps1  # VS Code
   # or
   powershell_ise.exe SYSTEM-REPORT-TOOL.ps1  # PowerShell ISE
   ```

3. Test changes:
   ```powershell
   .\SYSTEM-REPORT-TOOL.ps1
   ```

4. Commit and push:
   ```powershell
   git add SYSTEM-REPORT-TOOL.ps1
   git commit -m "Description of changes"
   git push origin main
   ```

### Recommended Tools
- **VS Code**: With PowerShell extension
- **PowerShell ISE**: Built-in Windows editor
- **Git**: Version control
- **Pester**: PowerShell testing framework

---

Last Updated: 2025-10-26
Project Version: 1.0
PowerShell Version Compatibility: 5.1+
Windows Compatibility: 7/8/10/11, Server 2008 R2+
