write-host "`n---Jeff Rudolph 2024---"

class Item {
    #properties
    [string] $Barcode
    [string] $Resource
    [string] $HomeLocation
    [string] $Status
    [string] $Custodian
    #Default constructor
    Item() { $this.Init(@{}) }
    #Convenience constructor from hashtable
    Item([hashtable]$Properties) {$this.Init($Properties)}
    #Shared initializer method -> popoulate properties with hashtable w/ matching key names
    [void] Init([hashtable]$Properties){
        foreach ($Property in $Properties.Keys){
            $this.$Property = $Properties.$Property
        }
    }
}

Write-Host "This script converts Follet XML for unaccounted devices into a .csv"
Write-Host "*Note* Excel may hide leading 0s in Barcodes, but the data is still in the file, if this is an issue for you, open the .csv with Notepad `n"

$path = Read-Host "Enter Full Filepath to XML file"

#remove quotation marks from path if copied from windows file explorer
if($path.StartsWith('"')){ $path = $path.Substring(1, $path.Length-2) }


#Get all <item> tags with their subtags from xml file
[xml]$types = Get-Content $path
$items = Select-Xml -Xml $types -XPath "//Item" 

#instantiating empty mutable non type specific list
$objectList = New-Object 'System.Collections.Generic.List``1[System.Object]'

#go through each <item> tag in document and create an instance of the Item object with the corresponding values in the child tags and add the Item object to a list
foreach($item in $items){
    $resource = if($item.Node.AssetName){$item.Node.AssetName.InnerText} else {"BLANK"}
    $homeLoaction = if($item.Node.HomeLocationName){$item.Node.HomeLocationName.InnerText} else {"BLANK"}
    $status = if($item.Node.Status){$item.Node.Status.InnerText} else {"BLANK"}
    if($item.Node.CheckedOutToName){$status = $status + " " + $item.Node.CheckedOutToName.InnerText}
    $custodian = if($item.Node.CustodianName){$item.Node.CustodianName.InnerText} else {"BLANK"}

    $newItem = [Item]::new(@{
        Barcode = $item.Node.Barcode.InnerText
        Resource = $resource
        HomeLocation = $homeLoaction
        Status = $status
        Custodian = $custodian
    })
    $objectList.Add($newItem)
    
}

#creating unique filename to save to current users desktop
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$todaysDate = Get-Date -UFormat "%m_%d_%Y_At%H_%M_%S%p"
$DesktopPath = $DesktopPath + "\" + "FolletXMLSpreadsheet_" + $todaysDate + ".csv"

#create a csv and append each Item Object to the csv
foreach($obj in $objectList){
    $obj | Export-Csv -Path $DesktopPath -NoTypeInformation -Append
}

#done
Write-Host "Saving File To Desktop..."

Write-Host "`nScript Finished - Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
