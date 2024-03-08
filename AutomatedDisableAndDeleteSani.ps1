#Jeff Rudolph 2024
#Script will run daily to keep AD clean and up to date
#Accounts will be disabled after 120 days and when they are disabled the day they were disabled will be recorded in the -State attrib.
#Faculty accounts will be set to delete 7yrs after the disable date if they are still inactive, Student accounts will never be automatically deleted in this manner
#Staff accounts that never logged in after being created after 120 days will be deleted regardless if disabled or not


#defining dates
$todaysDate = Get-Date
$oneTwentyDaysAgo = (Get-Date).AddDays(-120)
$yearsAgo = (Get-Date).AddDays(-2520) #7 yrs

#make directories if not exist, else wont overwrite, preparing filepaths for csv logs
$logFolder = New-Item -Path "C:\" -Name "ActiveDirectory-AutomaticScriptingLogs" -ItemType Directory -Force
$logFolderPath = Join-Path "C:\" -ChildPath $logFolder.Name
$todaysDateFilePathString = (Get-Date).ToShortDateString().Replace("/","-")



#OUs for the script
$oldGradeYearsOU = "OU="
$staffOU = "OU="
$studentsOU = "OU="



#Check Old Grade Years OU for any accounts that might have gotten in there without being disabled and disable them + update date(State) field 
$oldGradeYearAccountCheck = Get-Aduser -SearchBase $oldGradeYearsOU -Filter * -Properties * | Where-Object {$_.Enabled -eq $true} | Where-Object {$_.Description -inotlike "*(Service Account)*"}

foreach ($account in $oldGradeYearAccountCheck){
    Set-ADUser -Identity $account.SamAccountName -State $todaysDate
    Disable-ADAccount -Identity $account.SamAccountName
}



#get all enabled accounts that have not logged in within 120 days (students then staff)
$accountsToDisable = Get-ADUser -SearchBase $studentsOU -filter * -properties * | Where-Object {$_.Enabled -eq $true} | Where-Object {$_.LastLogonDate -lt $oneTwentyDaysAgo}|Where-Object {$_.LastLogonDate -ne $null} | Where-Object {$_.State -eq $null} | Where-Object {$_.Description -inotlike "*(Service Account)*"}
$accountsToDisable += Get-ADUser -SearchBase $staffOU -filter * -properties * | Where-Object {$_.Enabled -eq $true} | Where-Object {$_.LastLogonDate -lt $oneTwentyDaysAgo}| Where-Object {$_.LastLogonDate -ne $null} | Where-Object {$_.State -eq $null} | Where-Object {$_.Description -inotlike "*(Service Account)*"}
#disable all of these accounts and enter todays date in the "State" field for them
foreach ($account in $accountsToDisable){
    Set-ADUser -Identity $account.SamAccountName -State $todaysDate 
    Set-ADUser -Identity $_.SamAccountName -Description ("Account Auto-Disabled by script on $(get-date)")
    Disable-ADAccount -Identity $account.SamAccountName
}
#log disabled account as csv
if($accountsToDisable.Count -ne 0){
    $disabledFilePath = Join-Path $logFolderPath -ChildPath "$($todaysDateFilePathString)_DisabledAccountsInactive.csv"
    $accountsToDisable | Export-Csv -Path $disabledFilePath -NoTypeInformation
}



#Disable Student Accounts that have never logged in $LastLogon == null with date created > 120 days in past
$neverLoggedIn = Get-ADUser -SearchBase $studentsOU -Filter * -Properties * | Where-Object {$_.LastLogonDate -eq $null} | Where-Object {$_.Enabled -eq $true} | Where-Object {$_.whenCreated -lt $oneTwentyDaysAgo} | Where-Object {$_.State -eq $null} | Where-Object {$_.Description -inotlike "*(Service Account)*"}
foreach ($account in $neverLoggedIn){
    Set-ADUser -Identity $_.SamAccountName -State $todaysDate
    Set-ADUser -Identity $_.SamAccountName -Description ("Account Auto-Disabled by script on $(get-date)")
    Disable-ADAccount -Identity $account.SamAccountName
}
#log disabled student accounts that never logged in as csv
if($neverLoggedIn.Count -ne 0){
    $disabledNoLogonFilePath = Join-Path $logFolderPath -ChildPath "$($todaysDateFilePathString)_DisabledStudentAccountsNeverLoggedOn.csv"
    $neverLoggedIn | Export-Csv -Path $disabledNoLogonFilePath -NoTypeInformation
}



#delete staff accounts automatically after they have been disabled for 7  years OR if they have never logged in after 120 days, not doing deletion for  students as all students are organized by gradyear in OUs, trivial to manually delete these once the 7 yrs is up
$accountsToDelete = Get-ADUser -SearchBase $staffOU -Filter * -Properties * | Where-Object {$_.Enabled -eq $false} | Where-Object {([DateTime]$_.State) -lt $yearsAgo} | Where-Object {$_.State -ne $null}| Where-Object {$_.Description -inotlike "*(Service Account)*"}
$accountsToDelete += Get-ADUser -SearchBase $staffOU -Filter * -Properties * | Where-Object {$_.LastLogonDate -eq $null} | Where-Object {$_.whenCreated -lt $oneTwentyDaysAgo} | Where-Object {$_.State -eq $null} | Where-Object {$_.Description -inotlike "*(Service Account)*"}
foreach ($account in $accountsToDelete){
    Remove-ADUser -Identity $account.SamAccountName -Confirm:$false 
}
#log deleted accounts as csv
if($accountsToDelete.Count -ne 0){
    $deletedFilePath = Join-Path $logFolderPath -ChildPath "$($todaysDateFilePathString)_DeletedStaffAccounts.csv"
    $accountsToDelete | Export-Csv -Path $deletedFilePath -NoTypeInformation
}


#Any accounts that are re-enabled  after being disabled by this script will still have a date value in the state field.
#This following code block will clear the disabled date from the users attribs.
$reenabledUsers = Get-Aduser -filter * -Properties * | Where-Object {$_.enabled -eq $true} | Where-Object {$_.State -ne $null}
foreach ($account in $reenabledUsers){
    Set-ADUser -Identity $_.SamAccountName -State $null
}
