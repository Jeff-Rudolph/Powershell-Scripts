#-2024 Jeff Rudolph-

Write-Host "This script will remove all disabled user accounts from security groups. Script will report how many users were removed from each group."

#Get all groups, ignore password policy groups for FGPP, ignore default groups in users 
$allGroups = Get-ADGroup -Filter * -Properties * | Where-Object { ($_.distinguishedName -notmatch "password") -and ($_.distinguishedName -notmatch ".*(CN=Users).*") }

foreach($group in $allGroups){
    #grab all users of a group, excluding the important MS accounts and computers  
    $allUsersInGroup = Get-ADGroupMember -Identity $group.SamAccountName | Where-Object { ($_.objectClass -eq "user") -and  ($_.distinguishedName -notmatch ".*(CN=Users).*") }
    $disabledUsersToRemove = @()
    foreach($userObj in $allUsersInGroup){
        #get the actual Active Directory User Object with all properties
        $user = Get-ADUser -Identity $userObj.distinguishedName -Properties *
        if($user.Enabled -eq $false){
            #if the user account in the group is disabled add to mutable list
            $disabledUsersToRemove += $user.SamAccountName
        }
    }

    if($disabledUsersToRemove.Count -ne 0){
        #if disabled users found, remove them, write the result to the console.
        Write-Host "Disabled users removed from $($group.cn): $($disabledUsersToRemove.Count)"
        Remove-ADGroupMember -Identity $group.SamAccountName -Members $disabledUsersToRemove -Confirm:$False #-WhatIf
    }
    
}

#end of script
Write-Host "Saving File To Desktop..."
Write-Host "`nScript Finished - Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")