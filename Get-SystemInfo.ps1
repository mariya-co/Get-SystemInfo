<# Automated system information retrieval script
This script retrieves battery report, WLAN report, reliability history, current OS version, uptime, and processor state from a remote machine.
Requires admin credentials.
#>

Function Test-ComputerConnectivity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ComputerName
    )
    $pingResult = Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction SilentlyContinue
    if (!$pingResult) {
        Write-Host "The computer $ComputerName is not reachable. Please check the network connection." -ForegroundColor Red
        exit
    }
    try {
        Test-WSMan -ComputerName $ComputerName -Authentication Default -ErrorAction Stop | Out-Null
        Write-Host "WinRM is available on $ComputerName." -ForegroundColor Green
    }
    catch {
        Write-Host "WinRM is not available on $ComputerName. Please check the WinRM configuration." -ForegroundColor Red
    }
}

# Define the regex pattern for hostnames - this can be changed to suit your naming convention.
$pattern = '^[A-Za-z]{3}-[A-Za-z]{3}-\d{2}$'
# Prompt for valid computer name, check for null and syntax match.
while (
    [string]::IsNullOrWhiteSpace($ComputerName) -or 
    ($ComputerName -notmatch $pattern)
) {
    $ComputerName = Read-Host "Enter the computer name:"
    if ($ComputerName -notmatch $pattern) {
        Write-Host "Invalid format! Please check your syntax and try again." -ForegroundColor Red
    }
}

Test-ComputerConnectivity -ComputerName $ComputerName

# Check if running as administrator - query for credentials if not running as admin.
$credCheck = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $credCheck) {
    $cred = Get-Credential -Message "Enter local administrator credentials"
    $session = New-PSSession -ComputerName $ComputerName -Credential $cred
}
else {
    $session = New-PSSession -ComputerName $ComputerName
}

# Define paths.
$outputPath = "C:\Temp\$ComputerName-SystemReport"
$remotePath = "C:\Temp\SystemReports"

if (-not (Test-Path -Path $outputPath)) {
    Write-Host "Output path does not exist, creating $outputPath" -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $outputPath | Out-Null
}

$scriptBlock = {
    param ($ComputerName, $remotePath)

    # Define Get-ReliabilityHistory locally for this session
    Function Get-ReliabilityHistory {
        $cutoffDate = (Get-Date).AddDays(-30)
        $eventIds = 6008, 1000, 1002, 20, 1033, 1035, 19
        try {
            Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction Stop |
                Where-Object {
                    ([datetime]$_.TimeGenerated -ge $cutoffDate) -and 
                    ($eventIds -contains $_.EventIdentifier)
                } |
                Select-Object TimeGenerated, SourceName, EventIdentifier, ProductName, Message
        }
        catch {
            Write-Error "Failed to query reliability records - $($_.Exception.Message)"
            $null
        }
    }

    if (-not (Test-Path -Path "C:\Temp")) {
        New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    }
    if (-not (Test-Path -Path $remotePath)) {
        New-Item -ItemType Directory -Path $remotePath | Out-Null
    }

    # Battery report - will only run if battery is present
    Write-Host "Generating battery report..." -Foregroundcolor Yellow
    $batDevice = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    $batPath = Join-Path -Path $remotePath -ChildPath "$ComputerName-BatteryReport.html"
    if ($batDevice) {
        powercfg /batteryreport /output $batPath | Out-Null
        Write-Host "Battery report generated" -ForegroundColor Green
    }
    else {
        Write-Host "No battery detected, skipping battery report." -ForegroundColor Yellow
    }

    # WLAN report
    Write-Host "Generating WLAN report..." -Foregroundcolor Yellow
    $wlanReportName = "$ComputerName-WLANReport.html"
    $generatedPath = "C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html"
    netsh wlan show wlanreport | Out-Null
    if (Test-Path $generatedPath) {
        Copy-Item -Path $generatedPath -Destination (Join-Path -Path $remotePath -ChildPath $wlanReportName) -Force
    }
    Write-Host "WLAN report generated" -ForegroundColor Green

    # Reliability history function that is saved into a variable
    Write-Host "Gathering reliability history..." -Foregroundcolor Yellow
    $reliabilityHistory = Get-ReliabilityHistory

    # Gather data that will be used in the report
    Write-Host "Gathering system information..." -Foregroundcolor Yellow
    $uptimeObj = (Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
    $uptime = "$($uptimeObj.Days) days, $($uptimeObj.Hours) hours, $($uptimeObj.Minutes) minutes"
    $processorState = Get-CimInstance -Query "select Name, PercentProcessorTime from Win32_PerfFormattedData_PerfOS_Processor" | Select-Object Name, PercentProcessorTime
    $diskSpace = Get-CimInstance -ClassName Win32_LogicalDisk | 
        Select-Object DeviceID, @{Name='FreeSpaceGB'; Expression={[math]::round($_.FreeSpace / 1GB, 2)}}, 
        @{Name='TotalSizeGB'; Expression={[math]::round($_.Size / 1GB, 2)}}
    $osinfoObj = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' | Select-Object CurrentBuildNumber, UBR
    $osVersion = "$($osinfoObj.CurrentBuildNumber).$($osinfoObj.UBR)"

    # Define main info for the top of the report
    $mainInfo = [PSCustomObject]@{
        DateOfReport = (Get-Date).ToString("dd-MM-yyyy HH:mm:ss")
        ComputerName = $ComputerName
        OSVersion = $osVersion
        Uptime = $uptime
    }
    Write-Host "Done! System information gathered." -ForegroundColor Green
    Write-Host "Generating HTML report..." -Foregroundcolor Yellow
    # Convert each section to HTML fragments

    $mainInfoHtml = $mainInfo | ConvertTo-Html -Fragment -PreContent "<h2>Main System Information</h2>" 
    $processorHtml = $processorState | ConvertTo-Html -Fragment -PreContent "<h2>Processor Utilisation</h2>"
    $diskSpaceHtml = $diskSpace | ConvertTo-Html -Fragment -PreContent "<h2>Disk Space</h2>"
    $reliabilityHtml = if ($reliabilityHistory) {
        $reliabilityHistory | ConvertTo-Html -Fragment -PreContent "<h2>Reliability History (Last 30 Days)</h2>"
    } else {
        "<h2>Reliability History (Last 30 Days)</h2><p>No reliability history found.</p>"
    }

    $combinedHtml = @"
<html>
<head>
    <title>System Report for $ComputerName</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h2 { color: #2e6c80; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
$mainInfoHtml
$processorHtml
$diskSpaceHtml
$reliabilityHtml
</body>
</html>
"@

    $htmlReportPath = Join-Path -Path $remotePath -ChildPath "$ComputerName-SystemReport.html"
    $combinedHtml | Out-File -FilePath $htmlReportPath -Encoding utf8
    Write-Host "System report finished!" -ForegroundColor Green
}


# Run on remote machine
Invoke-Command -Session $session -ScriptBlock $scriptBlock -ArgumentList $ComputerName, $remotePath

# Copy HTML reports from remote to local test output (using session context)
Copy-Item -Path "$remotePath\*.html" -Destination $outputPath -FromSession $session -Force
Write-Host "System report generated at $outputPath" -ForegroundColor Green
Write-Host "Starting cleanup process..." -ForegroundColor Yellow

# Clean up temp of remote machine
Invoke-Command -Session $session -ScriptBlock {
    param ($remotePath)
    if (Test-Path -Path $remotePath) {
        Remove-Item -Path $remotePath -Force -Recurse
    }
} -ArgumentList $remotePath

# Clean up the remote session
Remove-PSSession -Session $session

Write-Host "Cleanup completed, session has been closed and temporary files deleted." -ForegroundColor Green