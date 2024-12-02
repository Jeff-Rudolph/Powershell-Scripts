#-Jeff Rudolph 2024-
#WILL RESTART COMPUTERS DO AT END OF DAY

#----------First run these commands to save the encrypted credentials to the desktop----------
#$credential = Get-Credential
#$credentialPath = Join-Path ([Environment]::GetFolderPath("Desktop")) SecureRemotePCNameChange.ps1.credential 
#$credential | Export-Clixml $credentialPath
#---------------------------------------------------------------------------------------------

Write-Host "This script will output a FailedAutoRename_[TODAYS DATE].csv file that can be refed into this script at a later date to try renaming these PCs again"
Write-Host "Successfully renamed PCs will be recorded in their own spreadsheet with the old name and the updated name. Both will be saved to current users Desktop."


#import secure credentials
$credentialPath = Join-Path ([Environment]::GetFolderPath("Desktop")) SecureRemotePCNameChange.ps1.credential 
$credential = Import-Clixml $credentialPath


#import csv with computer names, column names of note are NewName CurrentName
$csvPath = Read-Host "Enter Full Filepath to csv file"
#remove quotation marks from path if copied from windows file explorer
if($csvPath.StartsWith('"')){ $csvPath = $csvPath.Substring(1, $csvPath.Length-2) }

$csv = Import-Csv -Path $csvPath

$DNSIssues = @()
$connectionIssues =@()

#getting the date for unique filepath creation
$todaysDate = Get-Date -UFormat "%m_%d_%Y_At%H_%M_%S%p"
$successSpreadsheet = [Environment]::GetFolderPath("Desktop") + "\" + "SuccessfulAutoRename.csv"
$failSpreadsheet = [Environment]::GetFolderPath("Desktop") + "\" + "FailedAutoRename_" + $todaysDate + ".csv"

$count = 0
foreach($row in $csv){

    #Search DNS for DNS records matching computer name, if not in DNS most likely has not been online in a very long time 
    try{
        Resolve-DnsName -Name $row.CurrentName -ErrorAction Stop | Out-Null
    }
    catch{
        $DNSIssues += $row.CurrentName
        #$row | Export-Csv -Path $failSpreadsheet -NoTypeInformation -Append
        continue
    }
    

    $currentName = $row.currentName + ".domain.org"
    $newName = $row.NewName + ".domain.org"
    
    $passThru =  Rename-Computer -ComputerName $row.CurrentName -NewName $row.NewName -DomainCredential $credential -PassThru -force -Restart
    
    #if the computer has been successfully changed print the output
    if($passThru.HasSucceeded -eq $true ){
        $passThru
        #$passThru | Export-Csv -Path $successSpreadsheet -NoTypeInformation -Append
    }  
    else{
        #computer was unreachable for some reason, most likely it is not currently online in the network.
        $connectionIssues += $row.CurrentName
        #$row | Export-Csv -Path $failSpreadsheet -NoTypeInformation -Append
    } 
    

}


#Enable these if you want more info on failures
#$DNSIssues |Out-File -FilePath (Join-Path ([Environment]::GetFolderPath("Desktop")) FailedRename_DNSIssues.txt)
#$connectionIssues | Out-File -FilePath (Join-Path ([Environment]::GetFolderPath("Desktop")) FailedRename_ConnIssues.txt)