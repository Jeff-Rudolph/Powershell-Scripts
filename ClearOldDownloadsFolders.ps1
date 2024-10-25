#-Jeff Rudolph 2024-

Write-Host "`n`nThis script will look through the shared drives and clear any user's downloads folder if they have not downloaded anything to it within 2 years"
Write-Host "`nThis script is assuming that the location of the shared drives will exist in D:\\Home Directories\\[Admin/Faculty/Staff/etc...]"
Write-Host "If the file structure has been changed then the script will need to be updated."
Write-Host "`nNOTE:The Admin, Tech, and _Anti Ransomware folders currently present in this directory will be excluded from this process."
Write-Host "`nNOTE:This script checks the folders LastWriteTime attribute, therefore, there is an edge case in which a user who has not saved or deleted any files to their downloads but still actively reads files from the folder gets deleted."
Write-Host "`n---Running Script - This will take a few minutes---`n"

#D drive > Home Directories > Categories by job (exclude admin and tech and anti ransomware)


$yearsAgo = (Get-Date).AddDays(-730)

$basePath = "D:\Home Directories\"

$ouFolders = (Get-ChildItem -Directory -Path $basePath | Where-Object {($_.Name -ne "_Anti Ransomware") -and ($_.Name -ne "Tech") -and ($_.Name -ne "Admin")}).Name

$oldDownloadFolders = @()

#counting the total folders that exist so that the operator can determine if the numbers seem appropriate before going ahead with deletion
$studentCount = 0
$facultyCount = 0

foreach($ou in $ouFolders){
    $ouPath = $basePath + "$ou"
    if($ou -eq "Students"){
        $oldDownloadFolders += Get-ChildItem -Directory -Path $ouPath -Recurse -ErrorAction SilentlyContinue | Where-Object {($_.FullName -match "(Students)\\\d\d\\[^\\]+\\(Downloads)$") -and ($_.LastWriteTime -lt $yearsAgo) -and ($_.LastWriteTime -ne $null)}
        #each student account will have a downloads folder so can use regex to count them (same for faculty below in else)
        $studentCount += (Get-ChildItem -Directory -Path $ouPath -Recurse -ErrorAction SilentlyContinue | Where-Object {($_.FullName -match "(Students)\\\d\d\\[^\\]+\\(Downloads)$")} ).Count
    }
    else{
        $regex = "($($ou))\\[^\\]+\\(Downloads)$"
        $oldDownloadFolders += Get-ChildItem -Directory -Path $ouPath -Recurse -ErrorAction SilentlyContinue | Where-Object {($_.FullName -match $regex) -and ($_.LastWriteTime -lt $yearsAgo) -and ($_.LastWriteTime -ne $null)}
        $facultyCount += (Get-ChildItem -Directory -Path $ouPath -Recurse -ErrorAction SilentlyContinue | Where-Object {($_.FullName -match $regex)}).Count
    }
}

$answer = Read-Host "`nThere are $($studentCount) student and $($facultyCount) faculty Downloads folders ready to be deleted. Do you wish to proceed? [Y/N]"

#if user answers Y proceed else cancel
#Will delete anything in the downloads folder including hidden files, and any subdirectories as well as their contents
if(($answer -eq "Y") -or ($answer -eq "y")){
    Write-Host "Deletion in progress...`n"
    $deletedFileSize = 0
    $oldDownloadFoldersFilePaths = $oldDownloadFolders.FullName
    foreach($filepath in $oldDownloadFoldersFilePaths){
        $deletedFileSize += (Get-ChildItem -Path $filepath -Recurse -Force | Measure-Object -Sum Length).Sum
        Get-childitem -path $filepath -Recurse -ErrorAction SilentlyContinue -Force |  Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    }
    $deletedFileSize = [math]::Round(($deletedFileSize / 1GB),2)
    write-host "`nDeleted $($deletedFileSize) GB"
}
else{
    Write-Host "`nDeletion cancelled"
}

Write-Host "`nScript Finished - Press any key to exit..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
