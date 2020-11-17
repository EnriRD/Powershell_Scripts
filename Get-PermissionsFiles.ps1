#directory to be scanned for files and folders permissions details
$folder = '\\ERD-DC01\ERD_Documents\'

#directory to be scanned for users, groups and folders permissions details
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

$Output | Out-GridView  -Title 'Users, groups and folders permissions details'

$folder = '\\ERD-DC01\ERD_Documents\'

(get-acl $folder).access | ft IdentityReference,FileSystemRights,AccessControlType,IsInherited,InheritanceFlags -auto

  function Get-Permissions ($folder) {
  (get-acl $folder).access | select `
		@{Label="Identity";Expression={$_.IdentityReference}}, `
		@{Label="Right";Expression={$_.FileSystemRights}}, `
		@{Label="Access";Expression={$_.AccessControlType}}, `
		@{Label="Inherited";Expression={$_.IsInherited}}, `
		@{Label="Inheritance Flags";Expression={$_.InheritanceFlags}}, `
		@{Label="Propagation Flags";Expression={$_.PropagationFlags}} | ft -auto
		}

$arrayfolders = Get-ChildItem $folder | 
       Where-Object {$_.PSIsContainer} | 
       Foreach-Object {$_.FullName}



Get-Permissions $folder

$filespermissions = dir $arrayfolders 

$filespermissions | export-csv -Path "$env:USERPROFILE\desktop\filespermissions.csv" -NoTypeInformation 

$filespermissions| Out-File "$env:USERPROFILE\desktop\filespermissions.txt"

$filespermissions | Select-Object -Property Mode, FullName, Name, LastWriteTime | Out-GridView -Title 'Files permissions details'