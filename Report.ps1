# Import the Active Directory module for the Get-ADComputer CmdLet
Import-Module ActiveDirectory

# Get today's date for the report
$date = Get-Date

# Array to hold results
$results = @()

# Get all AD computers
$computers = Get-ADComputer -Filter *

# Iterate over each computer
foreach ($computer in $computers) {
    try {
        $computerInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer.Name -ErrorAction Stop
        $osInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer.Name -ErrorAction Stop
        $cpuInfo = Get-WmiObject -Class Win32_Processor -ComputerName $computer.Name -ErrorAction Stop
        $userInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer.Name -ErrorAction Stop
        
        $result = New-Object PSObject
        $result | Add-Member NoteProperty "Name" $computer.Name
        $result | Add-Member NoteProperty "Model" $computerInfo.Model
        $result | Add-Member NoteProperty "Status" $computerInfo.Status
        $result | Add-Member NoteProperty "Roles" ($computerInfo.Roles -join ', ')
        $result | Add-Member NoteProperty "OperatingSystem" $osInfo.Caption
        $result | Add-Member NoteProperty "Memory" "$([math]::round($computerInfo.TotalPhysicalMemory/1GB,2)) GB"
        $result | Add-Member NoteProperty "Processor" $cpuInfo.Name
        $result | Add-Member NoteProperty "ProcessorSpeed" "$([math]::round($cpuInfo.MaxClockSpeed/1000,2)) GHz"
        $result | Add-Member NoteProperty "User" $userInfo.UserName

        $results += $result

        Write-Host "Completed for computer: $($computer.Name)"
    } 
    catch {
        Write-Host "Failed to gather info for computer: $($computer.Name)"
    }
}

# Output results to CSV
$results | Export-Csv -Path "C:\Users\sbahga\Desktop\DailyComputerReport.csv" -NoTypeInformation
