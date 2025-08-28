# ------------------------------
# Computer Information Report
# ------------------------------

# 1) CSS header with a “topLink” class for the [Top] links
$header = @"
<style>
    h1 {
        font-family: "Century Gothic", Helvetica, sans-serif;
        color: #7030A0;
        font-size: 22px;
        font-weight: bold;
    }
    h2 {
        font-family: "Segoe UI", Arial, sans-serif;
        color: #0090D0;
        font-size: 16px;
    }
  
    /* only the [Top] links get 7px text */
    p.topLink {
        font-family: "Arial", Arial, sans-serif;
        color: #0090D0;
        font-size: 9px;
        margin: 2px 0;
    }
    table {
        font-family: "Segoe UI", Arial, sans-serif;
        width: 10in;
        font-size: 12px;
        table-layout: auto;
        border: 2px solid #666666;
        border-collapse: collapse;
        background: #f7f7f7;
    }
    td {
        padding: 2px;
        border: 1px solid #666666;
    }
    th {
        background: linear-gradient(#49708f, #293f50);
        font-family: "Segoe UI", Helvetica, sans-serif;
        color: #fff;
        font-size: 15px;
        text-transform: uppercase;
        padding: 5px 7px;
    }
    #CreationDate {
        font-family: Segoe UI, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;
    }
    .StopStatus { color: #ff0000; }
    .RunningStatus { color: #008000; }
</style>
"@

# 2) “Top” anchor and page title
$TopAnchor         = '<a id="top"></a>'
$ComputerNameHtml  = "<h1>Computer name: $env:COMPUTERNAME</h1>"

# 3) Gather system info
$systemInfo = Get-ComputerInfo
$cs         = Get-CimInstance Win32_ComputerSystem
$os         = Get-CimInstance Win32_OperatingSystem
$bios       = Get-CimInstance Win32_BIOS

# 4) Build a PSCustomObject for System Info
$BasicInfo = [PSCustomObject]@{
    'Computer Name'      = $env:COMPUTERNAME
    'Domain'             = if ($cs.PartOfDomain) { $cs.Domain } else { 'Workgroup' }
    'Logged On User'     = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    'OS'                 = "$($os.Caption) ($($os.OSArchitecture))"
    'Version'            = "$($os.Version) Build $($os.BuildNumber)"
    'Last Boot Up Time'  = $os.LastBootUpTime
    'Uptime (d.hh:mm)'   = ((Get-Date) - $os.LastBootUpTime).ToString('d\.hh\:mm')
    'BIOS Version'       = $bios.SMBIOSBIOSVersion
    'BIOS Release Date'  = $bios.ReleaseDate
    'PowerShell Version' = $PSVersionTable.PSVersion
    'Time Zone'          = (Get-TimeZone).Id
}

# --- take your $BasicInfo object and turn its properties into rows ---
$BasicInfoTable = $BasicInfo.PSObject.Properties |
  Select-Object `
    @{Name='Property';Expression={$_.Name}}, `
    @{Name='Value'   ;Expression={$_.Value}}

# --- convert *that* into a real HTML table with headers "Property" / "Value" ---
$SystemHtml = $BasicInfoTable |
  ConvertTo-Html -As Table -Fragment `
    -PreContent '<h2>System Info</h2>' `
    -PostContent '<p class="topLink"><a href="#top">[Top]</a></p>' `
    -Property Property,Value



$ProcessData = Get-CimInstance Win32_Processor |
  Select-Object DeviceID,Name,Manufacturer,SocketDesignation,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed 
  

$ProcessInfoData = $ProcessData.PSObject.Properties |
    Select-Object `
       @{Name='Property';Expression={$_.Name}}, `
       @{Name='Value'   ;Expression={$_.Value}}

$ProcessInfo = $ProcessInfoData | 
  ConvertTo-Html -As Table -Fragment `
    -PreContent '<h2>Processor Info</h2>' `
    -PostContent '<p class="topLink"><a href="#top">[Top]</a></p>' `
    -Property Property,Value

$BiosInfoObj = [PSCustomObject]@{
    Version         = $bios.SMBIOSBIOSVersion
    Manufacturer    = $bios.Manufacturer
    Name            = $bios.Name
    'Serial Number' = $bios.SerialNumber
    'BIOS Language' = $systemInfo.BiosCurrentLanguage
}
$BiosData = $BiosInfoObj.PSObject.Properties | 
      Select-Object `
       @{Name='Property';Expression={$_.Name}}, `
       @{Name='Value'   ;Expression={$_.Value}}

$BiosInfoHTML = $BiosData |
  ConvertTo-Html -As Table -Fragment `
    -PreContent '<h2>BIOS Information</h2>' `
    -PostContent '<p class="topLink"><a href="#top">[Top]</a></p>' `
    -Property Property, Value

$DiscInfo = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
  Select-Object `
    DeviceID,
    ProviderName,
    VolumeName,
    @{Name='Size (GB)'; Expression = {[math]::Round($_.Size/1GB,2)}},
    @{Name='Free (GB)'; Expression = {[math]::Round($_.FreeSpace/1GB,2)}},
    @{Name='Free (%)'; Expression = {[math]::Round(($_.FreeSpace/$_.Size*100),2)}} |
  ConvertTo-Html -As Table -Fragment `
    -PreContent '<h2>Disk Information</h2>' `
    -PostContent '<p class="topLink"><a href="#top">[Top]</a></p>'

$AppInfo = Get-WmiObject -Class Win32_Product |
  Select-Object Name,Version,Vendor,IdentifyingNumber |
  ConvertTo-Html -As Table -Fragment `
    -PreContent '<h2>Installed Programs and Apps</h2>' `
    -PostContent '<p class="topLink"><a href="#top">[Top]</a></p>'

$ServicesInfo = Get-CimInstance Win32_Service |
  Select-Object Name,DisplayName,State |
  ConvertTo-Html -As Table -Fragment `
    -PreContent '<h2>Services Information</h2>' `
    -PostContent '<p class="topLink"><a href="#top">[Top]</a></p>'

# color-code Running/Stopped
$ServicesInfo = $ServicesInfo -replace '<td>Running</td>','<td class="RunningStatus">Running</td>'
$ServicesInfo = $ServicesInfo -replace '<td>Stopped</td>','<td class="StopStatus">Stopped</td>'

# 6) Combine everything into a single HTML document
$Report = ConvertTo-Html `
  -Head $header `
  -Title "Computer Information Report" `
  -Body (
      $TopAnchor +
      $ComputerNameHtml +
      $SystemHtml +
      $ProcessInfo +
      $BiosInfoHTML +
      $DiscInfo +
      $ServicesInfo +
      $AppInfo
    ) `
  -PostContent "<p id='CreationDate'>Creation Date: $(Get-Date)</p>"

# 7) Write to file

IF(!(test-path "C:\RTH TECH")) {
md "C:\RTHTECH"
}
$Report | Out-File "C:\RTH TECH\CONCIERGEPC-Information-Report.html" -Encoding UTF8
