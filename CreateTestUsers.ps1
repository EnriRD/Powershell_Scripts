#get OU for users
import-module activedirectory

#Get Targetted OU
$orgOU = Get-ADOrganizationalUnit "ou=Test Users,ou=Org,dc=sh,dc=loc"
$orgOU.distinguishedname

#set password length
$length = "14"

#Outs the account and password created
$results =  "c:\logs\results.txt"

#Declares Inheritance
$inherNone = [System.Security.AccessControl.InheritanceFlags]::None
$propNone = [System.Security.AccessControl.PropagationFlags]::None
$inherCnIn = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit
$propInOn = [System.Security.AccessControl.PropagationFlags]::InheritOnly
$inherObIn = [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
$propNoPr = [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit
  
#current number of users in OU
$aduE = get-aduser -filter {samaccountname -like "*"} -SearchBase $orgOU
$existing = $aduE.count

#Import list of first and surnames
$FirstN = "C:\logs\names.csv"

#imports and works out max possible users that can be created
$impName = Import-Csv -path $FirstN
$FNCT = ($impName.firstname | where {$_.trim() -ne ""}).count 
$SNCT = ($impName.surname | Where {$_.trim() -ne ""}).count
$maxUN = $FNCT * $SNCT
$total = ($maxUn.ToString()) -10

do {$enter = ([int]$NOS = (read-host "Max User accounts is "$total", how many do you need"))
}
until ($nos -le $total)

$UserLists=@{}

#Randomises first and surnames
do {

    $FName = ($impName.firstname | where {$_.trim() -ne ""})|sort {get-random} | select -First 1
    $SName = ($impName.surname | Where {$_.trim() -ne ""}) |sort {get-random} | select -First 1
      
    $UserIDs = $Fname + "." + $Sname 
    
try {$UserLists.add($UserIds,$UserIDs)} catch {}
       $UserIDs = $null

     Write-Host $UserLists.count  
            
} until ($UserLists.count -eq $nos)

    $UserLists.count
    $userlists.GetEnumerator()
    $UserLists.key
    $ADUs = $UserLists.values
    

foreach ($ADu in $ADus)
    {
        #set var for random passwords
        $Assembly = Add-Type -AssemblyName System.Web
        $RandomComplexPassword = [System.Web.Security.Membership]::GeneratePassword($Length,4)

    foreach ($pwd in $RandomComplexPassword) 
        {
        #Splits username to be used to create first and surname        
        $ADComp = get-aduser -filter {samaccountname -eq $ADU}
        $spUse = $ADu.Split('.') 
        $firstNe = $spUse[0]
        $surNe = $spUse[1]

       $pwSec = ConvertTo-SecureString "$pwd" -AsPlainText -Force 

        #Creates user accounts
        if ($ADComp -eq $null)
            {
        new-aduser -Name "$ADU" `
        -SamAccountName "$ADU" `
        -AccountPassword $pwSec `
        -GivenName "$firstNe" `
        -surname "$surNe" `
        -displayname "$FnS" `
        -description "TEST $ADu" `
        -path $orgOU `
        -enable $true `
        -ProfilePath "\\shdc1\Profiles$\$ADU" `
        -HomeDirectory "\\shdc1\Home$\$ADU" `
        -HomeDrive "H:" `

        New-Item "\\shdc1\Home$\$ADU" -ItemType Directory -force

        $gADU = Get-ADUser $ADU
 
        $H = "\\shdc1\Home$\$ADU"
        $getAcl = Get-Acl $H
 
        $fileAcc = New-Object System.Security.AccessControl.FileSystemAccessRule($gADU.sid, "MODIFY", "$inherCnIn,$inherObIn", "None", "Allow")
        $getacl.setAccessRule($fileAcc)
 
        Set-Acl $H $getacl
   


        #set Group membership
        Add-ADGroupMember -Identity "DFSAccess"-Members $ADU

        #Outs results to Results file
        $adu | out-file $results -Append
        $pwd | out-file $results  -Append
        "  " | out-file $results -Append
                }
       else {"nope exists "}
       
       write-host $ADU
               
            }
      }  
 
# Total users in OU
$aduC = get-aduser -filter {samaccountname -like "*"} -SearchBase $orgOU
$TotalU = $aduC.count

#Total users created
write-host "Total New Users"
$TotalU - $existing
