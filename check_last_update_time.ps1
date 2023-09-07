Import-Module ActiveDirectory
# Change and compatibale with your OU
$ou = "CN=Computers,DC=Domain,DC=Com"
$computers=Get-ADComputer -Filter * -SearchBase $ou
$report = @()

# Iterate through each computer and check the last Windows update date
foreach ($computer in $computers) {
    $computerName = $computer.Name
    $lastUpdate = Get-HotFix -ComputerName $computerName | Sort-Object InstalledOn -Descending | Select-Object -First 1
    $report += [PSCustomObject]@{
       "ComputerName" = $computerName
       "LastUpdate" = $lastUpdate.InstalledOn
   }
}

$report | Export-Csv -Path .\Desktop\file.csv 





