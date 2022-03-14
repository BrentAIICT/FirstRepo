<# 
.SYNOPSIS
  This command produces a report on all computers in the domain
.DESCRIPTION
  This command produces a report on all computers in the domain
  The report will be in an HTML format and will show the following 
  information:
    Computer (as a list table)
    --------
    Computer name
    Operating system
    ?Latest sevice pack installed ?
    
    each NIC (as a table)
    --------
    InterFace Name
    Interface Index
    IP Address
    Subnet mask
    MAC Address
    Default gateway
    DNS server address 
.EXAMPLE
  Get-ComputerReport
  This will identify all computers that are active in the domain 
  and produce an HTML report
.NOTES
  General notes
    Created By: Brent Denny
    Created On: 02-Mar-2022
    Last Modified : 
  Change Log:
    V0.1.0 02-Mar-2022 Created the report script  
#>
function HTMLFragment {
  $AllADComputers = Get-ADComputer -Filter *
  $IPRegex = '(\d{1,3}\.){3}\d{1,3}'
  foreach ($Computer in $AllADComputers) {
    $CimOpt = New-CimSessionOption -Protocol Dcom
    try {$CimSes = New-CimSession -ComputerName $Computer.Name -ErrorAction Stop}
    catch {$CimSes = New-CimSession -SessionOption $CimOpt -ComputerName $Computer.Name -ErrorAction Stop}
    $OSInfo = Get-CimInstance -CimSession $CimSes -ClassName Win32_OperatingSystem
    $HTML = $OSInfo | 
        Select-Object -Property @{n='ComputerName';e={$_.CSName}},
                                @{n='OperatingSystem';e={$_.Caption}},
                                OSArchitecture,
                                InstallDate |
        ConvertTo-Html -PreContent "<br><h2>$($OSInfo.CSName)</h2>"  -As List |
        Out-String
    $HTML -replace '<tr><td>(.+:)</td><td>','<tr><th>$1</th><td>'    
    $PhysNics = Get-CimInstance -CimSession $CimSes -ClassName Win32_NetworkAdapter | 
        Where-Object {$_.PhysicalAdapter -eq $true}
    foreach ($Nic in $PhysNics) {
      Get-CimInstance -CimSession $CimSes -ClassName Win32_NetworkAdapterConfiguration | 
          Where-Object {$_.InterfaceIndex -eq $Nic.InterfaceIndex} | 
          Select-Object @{n='InterfaceName';e={$Nic.Name}},
                        @{n='InterfaceIndex';e={$_.InterfaceIndex}},
                        @{n='MacAddress';e={$_.MacAddress}},
                        @{n='IPAddress';e={$_.IPAddress | Where-Object {$_ -match $IPRegex}}},
                        @{n='SubnetMask';e={$_.IPSubnet | Where-Object {$_ -match $IPRegex}}},
                        @{n='DefaultGateway';e={$_.DefaultIPGateway | Where-Object {$_ -match $IPRegex}}},
                        @{n='DNSServerAddress';e={$_.DNSServerSearchOrder | Where-Object {$_ -match $IPRegex}}} | 
              ConvertTo-Html -Fragment -PreContent '<h4>Physical Networ Adapters</h4>' 
    } 
    $PhysMemDimms = Get-CimInstance -CimSession $CimSes -ClassName Win32_PhysicalMemory
    $PhysMemDimms | 
        Select-Object -Property Name,Manufacturer,BankLabel,DeviceLocator,Capacity,PartNumber |
        ConvertTo-Html -Fragment -PreContent '<h4>Memory DIMMs</h4>'  

    $PhysDisks = Get-PhysicalDisk -CimSession $CimSes 
    $PhysDisks | 
        Select-Object -Property  Model,SerialNumber,Size,FirmwareVersion,MediaType |
        ConvertTo-Html -Fragment -PreContent '<h4>DiskDrives</h4>' 
 
  }
}
$BorderCSS = 'border:solid 2pt black;border-collapse:collapse;background-color:white;'
$CSS = "<style> h2 {width:100%;text-align:center;background-color:black;color:yellow;} body{background-color:#f2f2f2;} table{$BorderCSS} tr,td,th{$BorderCSS;padding:3pt;} th{background-color:lightblue} </style>"
$Frag = HTMLFragment
ConvertTo-Html -Head $CSS -Body $Frag | Out-file e:\report.html