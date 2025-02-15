# Function to play sound and speak text with increased speed
function Speak-Text {
  param (
    [string]$text
  )
  try {
    Add-Type -AssemblyName System.Speech
    $synth = New-Object System.Speech.Synthesis.SpeechSynthesizer
    $synth.Rate = 2  # Increase speech rate for faster output
    $synth.Speak($text)
  }
  catch {
    Write-Warning "Speech synthesis not available: $($_.Exception.Message)"
  }
  finally {
    if ($synth) {
      $synth.Dispose()
    }
  }
}

# Function to ensure the script runs with administrator privileges
function Ensure-Administrator {
  if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    Write-Host "Requesting administrator privileges..." -ForegroundColor Yellow
    Write-Host "Please wait..." -ForegroundColor Yellow
    # Start a new PowerShell process with elevated privileges and wait for it to finish
    Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -Wait
    # Exit the original process after the elevated process completes
    Exit
  }
}

# Invoke administrator check
Ensure-Administrator

$ErrorActionPreference = "Continue"
$outputFile = "$env:USERPROFILE\Desktop\COMPREHENSIVE SYSTEM REPORT - by JTG Systems.txt"
$minSizeGB = 0.1  # Minimum size in GB for large files

Write-Host "Starting Comprehensive System Report Generation..." -ForegroundColor Yellow
Speak-Text "Starting comprehensive system report generation."

Write-Host "Generating Comprehensive System Report and searching for files larger than $minSizeGB GB..." -ForegroundColor Cyan
Write-Host "Results will be saved to: $outputFile" -ForegroundColor Cyan

$start = Get-Date

# Define data collection tasks as hashtable entries
$tasks = @(
  @{
    Name        = "WindowsUpdateStatus"
    ScriptBlock = {
      try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        $pendingUpdates = $updateSearcher.Search("IsInstalled=0")
        $lastInstalledUpdate = (Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1).InstalledOn
        @{
          PendingUpdatesCount = $pendingUpdates.Updates.Count
          LastInstalledUpdate = $lastInstalledUpdate
        }
      }
      catch {
        Write-Warning "Error retrieving Windows Update status: $($_.Exception.Message)"
        @{
          PendingUpdatesCount = "N/A"
          LastInstalledUpdate = "N/A"
        }
      }
    }
  },
  @{
    Name        = "SystemInfo"
    ScriptBlock = {
      try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem
        $processor = Get-CimInstance -ClassName Win32_Processor
        $bios = Get-CimInstance -ClassName Win32_BIOS

        @{
          Manufacturer     = $computerSystem.Manufacturer
          Model            = $computerSystem.Model
          OSName           = $operatingSystem.Caption
          OSVersion        = $operatingSystem.Version
          ProcessorName    = $processor.Name
          ProcessorCores   = $processor.NumberOfCores
          ProcessorThreads = $processor.NumberOfLogicalProcessors
          BIOSVersion      = $bios.SMBIOSBIOSVersion
          BIOSDate         = $bios.ReleaseDate
        }
      }
      catch {
        Write-Warning "Error retrieving system information: $($_.Exception.Message)"
        @{
          Manufacturer     = "N/A"
          Model            = "N/A"
          OSName           = "N/A"
          OSVersion        = "N/A"
          ProcessorName    = "N/A"
          ProcessorCores   = "N/A"
          ProcessorThreads = "N/A"
          BIOSVersion      = "N/A"
          BIOSDate         = "N/A"
        }
      }
    }
  },
  @{
    Name        = "MemoryInfo"
    ScriptBlock = {
      try {
        $memorySlots = Get-CimInstance -ClassName Win32_PhysicalMemoryArray
        $installedMemory = Get-CimInstance -ClassName Win32_PhysicalMemory
        $totalRAM = [math]::Round(($installedMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)
        $totalSlots = $memorySlots.MemoryDevices
        $usedSlots = $installedMemory.Count
        $emptySlots = $totalSlots - $usedSlots

        @{
          TotalRAM   = $totalRAM
          TotalSlots = $totalSlots
          UsedSlots  = $usedSlots
          EmptySlots = $emptySlots
        }
      }
      catch {
        Write-Warning "Error retrieving memory information: $($_.Exception.Message)"
        @{
          TotalRAM   = "N/A"
          TotalSlots = "N/A"
          UsedSlots  = "N/A"
          EmptySlots = "N/A"
        }
      }
    }
  },
  @{
    Name        = "DiskInfo"
    ScriptBlock = {
      try {
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
        Select-Object DeviceID, @{Name = "SizeGB"; Expression = { [math]::Round($_.Size / 1GB, 2) } },
        @{Name = "FreeGB"; Expression = { [math]::Round($_.FreeSpace / 1GB, 2) } }
      }
      catch {
        Write-Warning "Error retrieving disk information: $($_.Exception.Message)"
        "N/A"
      }
    }
  },
  @{
    Name        = "GraphicsInfo"
    ScriptBlock = {
      try {
        Get-CimInstance -ClassName Win32_VideoController | Select-Object Name, @{Name = "MemoryGB"; Expression = { [math]::Round($_.AdapterRAM / 1GB, 2) } }, DriverVersion
      }
      catch {
        Write-Warning "Error retrieving graphics information: $($_.Exception.Message)"
        "N/A"
      }
    }
  },
  @{
    Name        = "NetworkInfo"
    ScriptBlock = {
      try {
        Get-NetAdapter | Where-Object Status -eq "Up" | Select-Object Name, InterfaceDescription, MacAddress, LinkSpeed
      }
      catch {
        Write-Warning "Error retrieving network information: $($_.Exception.Message)"
        "N/A"
      }
    }
  },
  @{
    Name        = "PowerInfo"
    ScriptBlock = {
      try {
        $powerPlan = powercfg /GETACTIVESCHEME
        if ($powerPlan -match '([a-f0-9-]{36})') {
          $guid = $matches[1]
          $planName = (powercfg /L | Where-Object { $_ -match $guid }) -replace '^.*\((.*)\).*$','$1'
        }
        else {
          $planName = "Unknown"
        }
        $batteryReport = if (Get-CimInstance -ClassName Win32_Battery) {
          Get-CimInstance -ClassName Win32_Battery | Select-Object @{Name = "BatteryStatus"; Expression = {
              switch ($_.BatteryStatus) {
                1 { "Discharging" }
                2 { "AC Power" }
                3 { "Fully Charged" }
                4 { "Low" }
                5 { "Critical" }
                6 { "Charging" }
                7 { "Charging and High" }
                8 { "Charging and Low" }
                9 { "Charging and Critical" }
                10 { "Undefined" }
                11 { "Partially Charged" }
              }
            }
          }, EstimatedChargeRemaining
        }
        else {
          "No battery detected"
        }
        @{
          PowerPlan   = $planName
          BatteryInfo = $batteryReport
        }
      }
      catch {
        Write-Warning "Error retrieving power information: $($_.Exception.Message)"
        @{
          PowerPlan   = "N/A"
          BatteryInfo = "N/A"
        }
      }
    }
  },
  @{
    Name        = "SystemUptime"
    ScriptBlock = {
      try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem
        $uptime = (Get-Date) - $os.LastBootUpTime
        "{0} days, {1} hours, {2} minutes" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
      }
      catch {
        Write-Warning "Error retrieving system uptime: $($_.Exception.Message)"
        "N/A"
      }
    }
  },
  @{
    Name        = "StartupPrograms"
    ScriptBlock = {
      try {
        Get-CimInstance -ClassName Win32_StartupCommand | Select-Object Name, Command, Location
      }
      catch {
        Write-Warning "Error retrieving startup programs: $($_.Exception.Message)"
        "N/A"
      }
    }
  },
  @{
    Name        = "InstalledSoftware"
    ScriptBlock = {
      try {
        Get-CimInstance -ClassName Win32_Product | Select-Object Name, Version, Vendor | Sort-Object Name
      }
      catch {
        Write-Warning "Error retrieving installed software: $($_.Exception.Message)"
        "N/A"
      }
    }
  },
  @{
    Name        = "FindLargeFiles"
    ScriptBlock = {
      param($minSizeGB)
      try {
        # Initialize variables
        $largeFiles = [System.Collections.ArrayList]::new()
        $errors = @()
        $GB = 1GB
        $progress = 0
        $excludedPathsRegex = '^.*\\(AppData|Local Settings|Application Data)\\.*$'
        
        # Create runspace pool for parallel processing
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, [Environment]::ProcessorCount)
        $runspacePool.Open()
        $jobs = @()
        
        Write-Host "`nStarting optimized file scan..." -ForegroundColor Yellow
        $scanPath = $env:USERPROFILE
        
        # Get directories first
        $directories = Get-ChildItem -Path $scanPath -Directory -ErrorAction SilentlyContinue |
                      Where-Object { $_.FullName -notmatch $excludedPathsRegex }
        
        # Process each directory in parallel
        foreach ($dir in $directories) {
            $powershell = [powershell]::Create().AddScript({
                param($path, $minSize, $excludedRegex)
                Get-ChildItem -Path $path -File -Recurse -ErrorAction SilentlyContinue |
                Where-Object { 
                    $_.Length -ge $minSize -and 
                    $_.FullName -notmatch $excludedRegex 
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        SizeGB = [math]::Round($_.Length / 1GB, 2)
                        Path = $_.FullName
                        LastModified = $_.LastWriteTime
                    }
                }
            }).AddArgument($dir.FullName).AddArgument($minSizeGB * $GB).AddArgument($excludedPathsRegex)
            
            $powershell.RunspacePool = $runspacePool
            
            $jobs += @{
                PowerShell = $powershell
                Handle = $powershell.BeginInvoke()
            }
        }
        
        # Collect results
        $progress = 0
        foreach ($job in $jobs) {
            $results = $job.PowerShell.EndInvoke($job.Handle)
            if ($results) {
                foreach ($item in $results) {
                    [void]$largeFiles.Add($item)
                }
            }
            $job.PowerShell.Dispose()
            
            # Update progress
            $progress++
            $percentComplete = [math]::Round(($progress / $jobs.Count) * 100)
            Write-Progress -Activity "Scanning directories" -Status "$percentComplete% Complete" -PercentComplete $percentComplete
        }
        
        $runspacePool.Close()
        $runspacePool.Dispose()
        
        Write-Host "`nScan Summary:" -ForegroundColor Yellow
        Write-Host "Large files found: $($largeFiles.Count)" -ForegroundColor Cyan
        
        # Return results sorted by size
        if ($largeFiles.Count -gt 0) {
            $largeFiles | Sort-Object SizeGB -Descending | Select-Object -First 50
        }
        else {
            @()
        }
      }
      catch {
        Write-Warning "Error in file scan: $($_.Exception.Message)"
        Write-Warning "Stack trace: $($_.Exception.StackTrace)"
        @()
      }
    }
  }
)

# Run FindLargeFiles task first
Write-Host "`nStarting file scan...`n" -ForegroundColor Yellow
Write-Host "Note: The scan will show real-time progress as it checks each folder." -ForegroundColor Cyan
Write-Host "Large files (>$minSizeGB GB) will be highlighted in green as they are found.`n" -ForegroundColor Cyan

$findLargeFilesTask = $tasks | Where-Object { $_.Name -eq "FindLargeFiles" }
$largeFilesResult = & $findLargeFilesTask.ScriptBlock -minSizeGB $minSizeGB
$results = @{ }
$results["FindLargeFiles"] = @($largeFilesResult)

# Start other tasks as background jobs
$jobs = @()
foreach ($task in $tasks) {
  if ($task.Name -ne "FindLargeFiles") {
    $jobs += Start-Job -Name $task.Name -ScriptBlock $task.ScriptBlock
  }
}

# Wait for all jobs to complete
Write-Host "`nGathering system information..." -ForegroundColor Yellow
Wait-Job -Job $jobs

# Retrieve job results
foreach ($job in $jobs) {
  if ($job.State -eq 'Completed') {
    try {
      $output = Receive-Job -Job $job
      $results[$job.Name] = $output
    }
    catch {
      $results[$job.Name] = "Error retrieving data."
      Write-Warning "Error receiving job output for $($job.Name): $($_.Exception.Message)"
    }
  }
  else {
    $results[$job.Name] = "Job did not complete successfully."
    Write-Warning "Job $($job.Name) did not complete successfully."
  }
  # Remove the job
  Remove-Job -Job $job
}

# Initialize report
$now = Get-Date
$dayOfWeek = $now.ToString('dddd')
$report = "===== Comprehensive System Report for Repair Technicians =====`n"
$report += "Generated on: $($now.ToString('yyyy-MM-dd    hh:mm tt'))  ($dayOfWeek)`n`n"

# Process Windows Update Status first
if ($tasks.Name -contains "WindowsUpdateStatus") {
  $wuData = $results["WindowsUpdateStatus"]
  $report += "--- Windows Update Status ---`n"
  $report += "Windows Update Status: $($wuData.PendingUpdatesCount) pending updates    "
  if ($wuData.PendingUpdatesCount -gt 0) {
    $report += "<------ Update your computer`n"
    $report += "Last Installed Update: $($wuData.LastInstalledUpdate)`n`n"
  }
  elseif ($wuData.PendingUpdatesCount -eq 0) {
    $report += "<------ Job well done!`n"
    $report += "Last Installed Update: $($wuData.LastInstalledUpdate)`n`n"
  }
  else {
    $report += "Status unavailable`n`n"
  }
}

# Process other tasks
foreach ($task in $tasks) {
  $name = $task.Name
  # Skip WindowsUpdateStatus as it's already processed
  if ($name -eq "WindowsUpdateStatus") { continue }
  $data = $results[$name]
  switch ($name) {
    "SystemInfo" {
      $report += "--- System Information ---`n"
      $report += "Manufacturer: $($data.Manufacturer)`n"
      $report += "Model: $($data.Model)`n"
      $report += "Operating System: $($data.OSName)`n"
      $report += "OS Version: $($data.OSVersion)`n"
      $report += "Processor: $($data.ProcessorName)`n"
      $report += "Processor Cores: $($data.ProcessorCores)`n"
      $report += "Processor Threads: $($data.ProcessorThreads)`n"
      $report += "BIOS Version: $($data.BIOSVersion)`n"
      $report += "BIOS Date: $($data.BIOSDate)`n"
      $report += "`n"
    }
    "MemoryInfo" {
      $report += "--- Memory Information ---`n"
      $report += "Total RAM: $($data.TotalRAM) GB`n"
      $report += "Total Memory Slots: $($data.TotalSlots)`n"
      $report += "Used Memory Slots: $($data.UsedSlots)`n"
      $report += "Empty Memory Slots: $($data.EmptySlots)`n"
      $report += "`n"
    }
    "DiskInfo" {
      $report += "--- Disk Information ---`n"
      if ($data -ne "N/A") {
        $report += ($data | Format-Table -AutoSize | Out-String)
      }
      else {
        $report += "Disk information unavailable.`n"
      }
      $report += "`n"
    }
    "GraphicsInfo" {
      $report += "--- Graphics Information ---`n"
      if ($data -ne "N/A") {
        $report += ($data | Format-Table -AutoSize | Out-String)
      }
      else {
        $report += "Graphics information unavailable.`n"
      }
      $report += "`n"
    }
    "NetworkInfo" {
      $report += "--- Network Adapter Information ---`n"
      if ($data -ne "N/A") {
        $report += ($data | Format-Table -AutoSize | Out-String)
      }
      else {
        $report += "Network adapter information unavailable.`n"
      }
      $report += "`n"
    }
    "PowerInfo" {
      $report += "--- Power and Battery Information ---`n"
      $report += "Active Power Plan: $($data.PowerPlan)`n"
      $report += "Battery Status: $($data.BatteryInfo | Out-String)"
      $report += "`n`n"
    }
    "SystemUptime" {
      $report += "--- System Uptime ---`n"
      $report += "Current Uptime: $($data)`n`n"
    }
    "StartupPrograms" {
      $report += "--- Startup Programs ---`n"
      if ($data -ne "N/A") {
        $report += ($data | Format-Table -AutoSize | Out-String)
      }
      else {
        $report += "Startup programs information unavailable.`n"
      }
      $report += "`n"
    }
    "InstalledSoftware" {
      $report += "--- Installed Software ---`n"
      if ($data -ne "N/A") {
        $report += ($data | Format-Table -AutoSize | Out-String)
      }
      else {
        $report += "Installed software information unavailable.`n"
      }
      $report += "`n"
    }
    "FindLargeFiles" {
      $report += "--- Largest Files Report (Top 50) ---`n"
      if ($data -and $data.Count -gt 0) {
        $report += "File Size (GB)  | Last Modified       | File Path`n"
        $report += "--------------- | ------------------- | ----------`n"
        foreach ($file in $data) {
          $report += "{0,-15} | {1,-19} | {2}`n" -f $file.SizeGB, $file.LastModified.ToString("yyyy-MM-dd HH:mm:ss"), $file.Path
        }
      }
      else {
        $report += "No large files found.`n"
      }
      $report += "`n"
    }
  }
}

# Save the report to the desktop
$report | Out-File -FilePath $outputFile -Encoding utf8

$end = Get-Date
$duration = $end - $start

# Prepare summary for speech
$summary = @()

# Important info: Windows Update Status, Total RAM, Number of large files found
if ($results["WindowsUpdateStatus"].PendingUpdatesCount -gt 0) {
  $summary += "You have $($results["WindowsUpdateStatus"].PendingUpdatesCount) pending Windows updates."
}
elseif ($results["WindowsUpdateStatus"].PendingUpdatesCount -eq 0) {
  $summary += "Windows is up to date."
}
else {
  $summary += "Windows Update status is unavailable."
}

$summary += "Total installed RAM is $($results["MemoryInfo"].TotalRAM) gigabytes."

if ($results["FindLargeFiles"] -ne "N/A") {
  $summary += "Found $($results["FindLargeFiles"].Count) large files exceeding $minSizeGB gigabytes."
}
else {
  $summary += "Large files information is unavailable."
}

$summary += "The report was generated in $([math]::Round($duration.TotalSeconds,2)) seconds."

# Final notifications
Speak-Text "Check the desktop for the complete report, sir."
Speak-Text ($summary -join " ")

# Inform user via Write-Host
Write-Host "Comprehensive report generated and saved to: $outputFile" -ForegroundColor Green
Write-Host "Time taken: $([math]::Round($duration.TotalSeconds,2)) seconds." -ForegroundColor Cyan
Write-Host "Check the desktop for the complete report, sir." -ForegroundColor Green

# Prompt to open the report
$openFile = Read-Host "Do you want to open the results file now? (Y/N)"
if ($openFile.Trim().ToUpper() -eq 'Y') {
  Invoke-Item $outputFile
}

Write-Host "Report generation completed successfully." -ForegroundColor Green
Read-Host "Press Enter to exit..."
