$FolderPath = Get-ChildItem -Directory -Path "\\ERD-DC01\ERD_Documents" -Recurse -Force
$Output = @()
ForEach ($Folder in $FolderPath) {
$Acl = Get-Acl -Path $Folder.FullName
ForEach ($Access in $Acl.Access) {
$Properties = [ordered]@{'Folder Name'=$Folder.FullName;'Group/User'=$Access.IdentityReference;'Permissions'=$Access.FileSystemRights;'Inherited'=$Access.IsInherited}
$Output += New-Object -TypeName PSObject -Property $Properties            
}
}

$OutPut | export-csv -Path "$env:USERPROFILE\desktop\folderspermissions.csv" -NoTypeInformation 

$OutPut | Out-File "$env:USERPROFILE\desktop\folderspermissions.txt"

$Output | Out-GridView