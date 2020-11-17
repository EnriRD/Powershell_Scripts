$GroupName = Get-ADGroup -SearchBase "OU=ERD_Folders_Permissions,OU=ERD_Office_1,DC=ERD,DC=net" -Filter * | Select-Object -ExpandProperty sAMAccountName

$Members = foreach ($GroupMember in $GroupName) {
   Get-ADGroupMember -Identity $GroupMember | Select-Object @{Name="Group";Expression={$GroupMember}},name
}
"$($Members.Group) $($Members.Name)"  | Out-File "$env:USERPROFILE\desktop\temp.txt"
#Alternatively for a .CSV
$Members | Export-CSV "$env:USERPROFILE\desktop\temp.CSV" -NoTypeInformation