#Jeff Rudolph -2024-

$listOfPrefixes = @()

$usernames = Get-ADUser -Filter * -Properties * -SearchBase "STUDENTOU" | Where-Object {$_.Enabled -eq $false} | select SamAccountName

#only grab accounts where they match the grad year followed by name pattern
#collect year prefixes
foreach ($user in $usernames){
    if($user.SamAccountName -match "^\d\d.*"){
        $append = $user.SamAccountName
        $listOfPrefixes += $append.Substring(0,2)
    }
}

#get unique vals
$OUsToCreate = $listOfPrefixes | Sort-Object | Get-Unique

#make OU's skip if already exists without throwing exception
foreach ($year in $OUsToCreate){
    $OUName = "20" + [string]$year
    try{New-ADOrganizationalUnit -Name $OUName -Path "DISABLEDSTUDENTOU" -ProtectedFromAccidentalDeletion $true -ErrorAction SilentlyContinue}
    catch{}
}

#again just grabbing all disabled student accounts matching the pattern
$disabledStudentAccounts = Get-ADUser -Filter * -Properties * -SearchBase "STUDENTOU" | Where-Object {$_.Enabled -eq $false} | Where-Object {$_.SamAccountName -match "^\d\d.*"} 


#move them to their new OUs to wait for deletion
foreach($account in $disabledStudentAccounts){
    $year = "20" + [string]$account.SamAccountName.Substring(0,2)
    $targetPath = "ou=$year,DISABLEDSTUDENTOU"
    Move-ADObject -Identity $account.DistinguishedName -TargetPath $targetPath
}