#-Jeff Rudolph 2024-
class ComputerComparison{
    [string] $FollettBarcode
    [string] $FollettLocation
    [int] $BadLocationFlag
    [string[]] $AssociatedADComputers
    [string] $DeviceDescription
    [string] $ActiveDirectoryPCs #variable will only be used to report either single or multiple AD computer names as one string

    #Default constructor
    ComputerComparison() { $this.Init(@{}) }
    #Convenience constructor from hashtable
    ComputerComparison([hashtable]$Properties) {$this.Init($Properties)}
    #Shared initializer method -> popoulate properties with hashtable w/ matching key names
    [void] Init([hashtable]$Properties){
        foreach ($Property in $Properties.Keys){
            $this.$Property = $Properties.$Property
        }
    }

}

function PopulateADComputers($barcodeStr){
    $filterStr = "Name -like '*" + $barcodeStr + "*'"
    $matchingPCs = Get-ADComputer -Filter $filterStr
    #$countOfPCs = $matchingPCs | Measure-Object | Select-Object -ExpandProperty Count
    $exportNameList = @()
    
    foreach($pc in $matchingPCs){
        $exportNameList += $pc.Name    
    }
    
     return $exportNameList

}

#Regex used to check the location in Follett vs the location info in the name of the computer. This will set a flag to be used in the creation of spreadsheets later
#This will only be ran on barcodes that correspond to one AD computer as barcodes tied to multiple PCs is a seperate (yet related) issue and will be exported in another sheet
#if the follett location is present in AD name this function will output 0 if this function detects a mismatch name it will output 1
function AnalyzeLocation($barcode, $folletCSV){
    $index = $folletCSV.barcode.IndexOF($barcode)

    $follettLocation = ($folletCSV[$index].'Home Location').split(" ")[0] #dropping anything past a first space as likelihood of it matching anything in AD is near 0
    
    $filterStr = "Name -like '*" + $barcode + "*'"
    $ADPCName = (Get-ADComputer -Filter $filterStr).Name

    if($ADPCName -match "^[a-zA-Z]+$([regex]::Escape($follettLocation))-.*"){
        return 0
    }
    else{
        return 1
    }
    
    
}

Write-Host "`nThis script takes a Follett spreadsheet consisting of all PC (workstation and laptops) from all buildings."
Write-Host "The barcodes from Follett will be checked with Active Directory to produce spreadsheets that show barcodes not present in AD, barcodes tied to multiple PCs in AD,`nand PCs where the Follett location and location from AD naming conventions does not match."
Write-Host "Keep in mind that the spreadsheet comparing Follett locations with AD names will have some false flags (ex in Follett a PC can have its location as M200 secretary but in AD it will only have M200)"
Write-Host "Spreadsheets will only be created if these issues are detected. All spreadsheets will be saved to the desktop with a descriptive file name."
write-host "`nNOTE: When creating the Follett spreadsheet that is used by this script, make sure it has Barcode, Description, and Home Location columns, and make sure it is a .csv`n`n"

$folletPath = Read-Host "Enter full filepath for Follett Spreadsheet"
if($folletPath.StartsWith('"')){ $folletPath = $folletPath.Substring(1, $folletPath.Length-2) }

#grab all PCs in AD (excludes servers)
$computerOUs = "LIST OF ALL OUs THAT DONT CONTAIN SERVERS HERE - EXCEPT ONE"

$allADComputers = Get-ADComputer -Filter * -Properties * -SearchBase "REMAINING OU PATH HERE"

foreach($ou in $computerOUs){
    $allADComputers += Get-ADComputer -Filter * -Properties * -SearchBase $ou
}




$folletSpreadSheet = Import-Csv -Path $folletPath






#going to make a collection of objects from the spreadsheet from follet. The associatedADcomputers param will be a flag that partially corresponds to what spreadsheet will be produced.
#also will need to make a flag that looks at the Follet location and compares to AD location. Anything that fails from there will be exported into spreadsheet.
#this could cause false flags tho as follet locations can be something like M200 assistant principle but the AD location would probably just be M200

#instantiating empty mutable non type specific list
$objectList = New-Object 'System.Collections.Generic.List``1[System.Object]'

#List of barcodes with >1 computer
#List of barcodes with no computers in AD
#List of barcodes with presumed wrong location based off AD name


foreach($row in $folletSpreadSheet){
    $newComputerComparison = [ComputerComparison]::new(@{
    FollettBarcode = $row.Barcode
    FollettLocation = $row.'Home Location'
    AssociatedADComputers = PopulateADComputers $row.Barcode
    DeviceDescription = $row.Description
    })
    if($newComputerComparison.AssociatedADComputers.Count -gt 1){
        #if multiple pcs come up for same barcode we will just set this to false
        $newComputerComparison.BadLocationFlag = 0 
    }
    elseif($newComputerComparison.AssociatedADComputers.Count -eq 0){
        $newComputerComparison.BadLocationFlag = -1
    }
    else{
        $newComputerComparison.BadLocationFlag = AnalyzeLocation $row.Barcode $folletSpreadSheet
    
    }
    

    $objectList.Add($newComputerComparison)
}


#creating unique filename to save to current users desktop
$desktopBase = [Environment]::GetFolderPath("Desktop")
$todaysDate = Get-Date -UFormat "%m_%d_%Y_At%H_%M_%S%p"

#Export csv of barcodes with no computers in AD
if(($objectList | Where-Object {$_.BadLocationFlag -eq -1}).count -ne 0){
    Write-host "Creating spreadsheet for Follett barcodes missing in AD..."
    $desktopPath = $desktopBase + "\" + "FollettBarcodesMissingInAD_" + $todaysDate + ".csv"
    
    foreach($obj in ($objectList | Where-Object {$_.BadLocationFlag -eq -1})){
        $obj | Select-Object FollettBarcode, FollettLocation, DeviceDescription | Export-Csv -Path $desktopPath -NoTypeInformation -Append
    }
}


#Export csv of barcodes with >1 computer
if(($objectList | Where-Object {$_.AssociatedADComputers.Count -gt 1}).Count -ne 0){
    Write-Host "Creating spreadsheet for Follett barcodes that relate to multiple PCs in AD..."
    $desktopPath = $desktopBase + "\" + "FollettBarcodesWithMultipleADPCs" + $todaysDate + ".csv"
    foreach($obj in ($objectList | Where-Object {$_.AssociatedADComputers.Count -gt 1})){
        $obj.ActiveDirectoryPCs = $obj.AssociatedADComputers -join ', '
        $obj | Select-Object FollettBarcode, FollettLocation, ActiveDirectoryPCs, DeviceDescription | Export-Csv -Path $desktopPath -NoTypeInformation -Append
    }
}

#Export csv of barcodes with presumed wrong location based off AD name
if(($objectList | Where-Object {$_.BadLocationFlag -eq 1}).Count -ne 0){
    Write-Host "Creating spreadsheet of potentially mismatched Follett locations/AD Names..."
    $desktopPath = $desktopBase + "\" + "FolletLocationAndADNameMismatch" + $todaysDate + ".csv"
    foreach($obj in ($objectList | Where-Object{$_.BadLocationFlag -eq 1})){
        $obj.ActiveDirectoryPCs = $obj.AssociatedADComputers -join ', '
        $obj | Select-Object FollettBarcode, FollettLocation, ActiveDirectoryPCs, DeviceDescription | Export-Csv -Path $desktopPath -NoTypeInformation -Append
    }
}