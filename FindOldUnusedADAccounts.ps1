Write-Host "-Jeff Rudolph 2024-"
Write-Host "This program will find all user accounts in active directory that have not logged in in a certain number of days and export spreadsheets with their info."
Write-Host "Spreadsheets will be saved to the same directory that this script is saved in."
Write-Host "*Note* Student accounts with email adresses that are not in the standard form of grad year followed by their initials and last name will not be picked up by this script"
$days = Read-Host("Enter the number of days")
$date = (Get-Date).AddDays(-$days)

#Looking at staff first

$OUpath = 'ENTER STAFF OU PATH HERE'



#make directories if not exist, else wont overwrite
$staffFolder = New-Item -Path $PSScriptRoot -Name "OldStaffAccounts" -ItemType Directory -Force
$studentFolder = New-Item -Path $PSScriptRoot -Name "OldStudentAccounts" -ItemType Directory -Force




$todaysDate = Get-Date -UFormat "%m_%d_%Y_At%H_%M_%S%p"

$staffFolder = Join-Path -Path $PSScriptRoot -ChildPath $staffFolder.Name
$staffFilename = Join-Path -Path $staffFolder -ChildPath "Staff-LastLogOn_$($days)_days_from_$($todaysDate).csv"

#Regex in here is to filter out non staff emails (ie Weather, NVision, SafeSchools accounts), Actual staff accounts seem to be of the form of a last name followed by \,"
Write-Host "`nSaving spreadsheet for staff accounts at $staffFilename"
Get-ADUser -Filter {LastLogonDate -lt $date} -SearchBase $OUpath  -Properties Name, DistinguishedName, LastLogonDate | Where-Object { ($_.Enabled -eq $true)}| Where-Object {$_.DistinguishedName -match "^CN=[A-za-z]+-?[A-za-z]*\\,"  }|select Name, DistinguishedName, LastLogonDate, SamAccountName | Export-Csv -Path $staffFilename -NoTypeInformation

#Now students
$OUpath = 'STUDENT OU PATH HERE'

$studentFolder = Join-Path -Path $PSScriptRoot -ChildPath $studentFolder.Name
$studentFilename = Join-Path -Path $studentFolder -ChildPath "Student-LastLogOn_$($days)_days_from_$($todaysDate).csv"

Write-Host "`nSaving spreadsheet for student accounts at $studentFilename"

#^\d\d[A-Za-z]{3,}\d* regex for finding only actual students by looking for the pattern of student emails
Get-ADUser -Filter {LastLogonDate -lt $date} -SearchBase $OUpath  -Properties Name, DistinguishedName, LastLogonDate | Where-Object { ($_.Enabled -eq $true)}| Where-Object {$_.SamAccountName -match "^\d\d[A-Za-z]{3,}\d*"  } | select Name,DistinguishedName, LastLogonDate, SamAccountName | Export-Csv -Path $studentFilename -NoTypeInformation

Write-Host "`nScript Finished - Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")