#-Jeff Rudolph 2024-

function Has-Duplicates {
    param (
        [Parameter(Mandatory=$true)]
        [array]$Collection
    )

    $duplicates = $Collection | Group-Object | Where-Object { $_.Count -gt 1 }
    return $duplicates.Count -gt 0
}


Write-Host "--------------------------------"
Write-Host "This is a tool for creating spreadsheets that organize new ChromeBooks by building for entering into inventory"
Write-Host "The spreadsheet created by this script will be saved to your desktop."
Write-Host "First you will be prompted to provide a master .csv provided by EduTech containing all devices from the SAA you're currently organizing."
Write-Host "You will be prompted to scan the Edutech tag for each device. Scanning a tag will automatically submit the number to the program."
Write-Host "No interaction with the keyboard or mouse should be required once you start scanning."
Write-Host "--------------------------------`n"
$csvPath = Read-Host "Input full filepath for master csv"
if($csvPath.StartsWith('"')){ $csvPath = $csvPath.Substring(1, $csvPath.Length-2) }

$masterSAA = Import-Csv -Path $csvPath

$eduTechTags = $masterSAA.'Tag#'
$hasDuplicates = Has-Duplicates -Collection $eduTechTags




if($hasDuplicates){
    Write-Host "Error! Duplicate EduTech Tags detected in master spreadsheet!"
    Exit
}  


$building = Read-Host "Input the building these CB's are for (1 - HS, 2 - MS, 3 - E)"
if($building -notin @("1", "2", "3")){
    while($building -notin @("1", "2", "3")){
        Write-Host "Error! Input the number corresponding to the building!"
        $building = Read-Host "Input the building these CB's are for (1 - HS, 2 - MS, 3 - E)"
        
    }
}


switch ($building)
{
    "1" {
            $buildingString = "HighSchool"
            $siteShortName = "ATTH"
        }
    
    "2" {
            $buildingString = "MiddleSchool"
            $siteShortName = "ATTM"
        }
    
    "3" {
            $buildingString = "Elementary"
            $siteShortName = "ATTE"
        }
}
$todaysDate = Get-Date -UFormat "%m_%d_%Y_At%H_%M_%S%p"
$exportPath = [Environment]::GetFolderPath("Desktop") + "\" + $buildingString + "_" + $todaysDate + ".csv"



$scannerInput = "Start"

Write-Host "Now we will begin barcode scanning, when you are done scanning barcodes simply type the word `"exit`""
while($true){
    $scannerInput = Read-Host "Scan Bardcode"

    if($scannerInput -eq "exit"){
        break
    }

    if($scannerInput -notmatch "\d+$"){
        Write-Host "Error Invalid Entry..."
        continue
    }

    $index = $masterSAA.'Tag#'.IndexOf($scannerInput)
    $CBObject = $masterSAA[$index]

    #export object with parameters that inventory is expecting for csv import
    $exportObject = [PSCustomObject]@{
        'Equipment Type' = ""
        Auditable = ""
        'Barcode' = $CBObject.'Tag#'
        DateAcquired = $CBObject.'Aquisition Date'
        Description = $CBObject.Description
        'Edutech Inventory' = "Yes"
        'Edutech SAA Project' = $CBObject.SAA
        'Equipment Category' = ""
        'Equipment Vendor' = $CBObject.Vendor
        'Home Location' = ""
        Owner = ""
        'Projected Life' = ""
        'Purchase Order' = $CBObject.PO
        'Purchase Price' = ""
        'Serial Number' = $CBObject.'Serial Number'
        SiteShortName = $siteShortName


    }

    $exportObject | Export-Csv -Path $exportPath -NoTypeInformation -Append
    
}