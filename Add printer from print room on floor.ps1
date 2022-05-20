#Add printer from print room on floor

#Print server
$prnServer = "printserver1"

$UserName = $UserOffice = $printer = $floor = $null
$UserName = $env:username 
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry()

#Get floor
#replace building number
[string]$UserOffice = $ADUser.physicalDeliveryOfficeName -replace "h5b3","house5-3" -replace "h5","house5-1"
if ($UserOffice.Length -eq 10) { $floor = $UserOffice[7] } elseif ($UserOffice.Length -eq 11) { $floor = $UserOffice[7] + $UserOffice[8] }

#Printer on floor 3 for empl. from 1 and 2 floor
if (($floor -eq "1") -or ($floor -eq "2")) {$floor = "3"}

$printer = Get-Printer -ComputerName $prnServer | Select-Object Name, Location, DeviceType | Where-Object {($_.DeviceType -eq "Print") -and (($_.Location -like "*$UserOffice") -or ($_.Location -like "*-f$floor"))}
foreach ($prn in $printer.name) { 
    Add-Printer -ConnectionName "\\$prnServer\$prn"
}