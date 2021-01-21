$ErrorActionPreference = 'SilentlyContinue'
Function Get-FolderI($initialDirectory="")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder to scan"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $initialDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
    if($foldername.ShowDialog() -eq "Annulla") { Break }
}

$path = Get-FolderI

Function Get-FolderD($destinationDirectory="")
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder to export the results as CSV file"
    $foldername.rootfolder = "MyComputer"
    $foldername.SelectedPath = $destinationDirectory

    if($foldername.ShowDialog() -eq "OK")
    {
        $destinationDirectory += $foldername.SelectedPath
    }

    return $destinationDirectory
    if($foldername.ShowDialog() -eq "Annulla") { Break }
}

$DestinationPath = Get-FolderD

Set-StrictMode -version 2


Write-Host "Starting analysis of $path"

$ErrorActionPreference = 'SilentlyContinue' 
$WarningPreference = 'SilentlyContinue'
$foldersCount = 0;
$filesCount = 0;
$folderStep = 10;
$Error.Clear()
$prefix = "$DestinationPath\" + "$(get-date -f yyyyMMdd-HHmm)";
$outERRORFile = "${prefix}_fileShareInventoryERRORS.csv"	

function LoopSubFoldersAndFiles($RootPath){
        
    $outFile = "${prefix}_fileShareInventory.csv"
    $FileName = "$outFile"

    Add-Content -Value  "Type;FolderPath;FileName;CreationTime;LastAccessTime;LastWriteTime;Size_Gigabytes(GB);Size_Megabytes(MB);Size_KiloBytes(KB);FilesQuantity;Extension;Owner;ReadOnly;Group/User;Permissions;Inherited" -Path $outFile
    $Folders = Get-ChildItem $RootPath -recurse -force -ErrorAction SilentlyContinue | where {$_.psiscontainer -eq $true}

    foreach ($Folder in $Folders) {
            $Acl = Get-Acl -Path $Folder.FullName -ErrorAction SilentlyContinue
        if ($Acl.Access -ne $null) {
        ForEach ($Access in $Acl.Access) {

            $Folder_ACL = get-acl $Folder.Fullname -ErrorAction SilentlyContinue

            $FSOFolder = $fso.GetFolder($Folder.Fullname) 

            $FolderSizeGB = "{0:N2}" -f ($FSOFolder.size / 1GB)
            $FolderSizeMB = "{0:N2}" -f ($FSOFolder.size / 1MB) 
            $FolderSizeKB = "{0:N2}" -f ($FSOFolder.Size / 1KB) 

            $FolderFileCount = $FSOFolder.Files.Count 
            $FolderOwner = $Folder_ACL.Owner    

            $Properties = "FOLDER;""" + $Folder.FullName + """;;" + $Folder.CreationTime + ";" + $Folder.LastAccessTime + ";" + $Folder.LastWriteTime  + ";" + $FolderSizeGB + ";" + $FolderSizeMB +";" + $FolderSizeKB + ";" + $FolderFileCount + ";" + "N/A" + ";"  + $FolderOwner + ";" + "N/A" + ";"  + $Access.IdentityReference + ";" + $Access.FileSystemRights + ";" + $Access.IsInherited
            Add-Content -Value $Properties -Path $outFile
        }
        }

        if($Folder -ne $null) {
            $foldersCount++
			
			if ($foldersCount%$folderStep -eq 0) {
				Write-Host "$foldersCount folders / $filesCount files analyzed ..."
			}
			
			try {
				$Folder_ACL = get-acl $Folder.Fullname -ErrorAction SilentlyContinue
				$FolderOwner = $Folder_ACL.Owner
			}
			
            catch {}
            
            $CleanFolderName = $Folder.Fullname.Replace(",","") #Remove commas in folder names
            
            try {                 
                $FSOFolder = $fso.GetFolder($Folder.Fullname) 
            }
            
            catch {}

            $FolderSizeGB = "{0:N2}" -f ($FSOFolder.size / 1GB)
            $FolderSizeMB = "{0:N2}" -f ($FSOFolder.size / 1MB) 
            $FolderSizeKB = "{0:N2}" -f ($FSOFolder.Size / 1KB) 

            $FolderFileCount = $FSOFolder.Files.Count 
            $OutInfo = "FOLDER;""" + $CleanFolderName + """;;" + $Folder.CreationTime + ";" + $Folder.LastAccessTime + ";" + $Folder.LastWriteTime  + ";" + $FolderSizeGB + ";" + $FolderSizeMB +";" + $FolderSizeKB + ";" + $FolderFileCount + ";" + "N/A" + ";"  + $FolderOwner + ";" + "N/A" + ";"  + $Access.IdentityReference + ";" + $Access.FileSystemRights + ";" + $Access.IsInherited
            
            Add-Content -Value $OutInfo -Path $outFile
            
            if($FolderFileCount -gt 0){

                $Files = Get-ChildItem $Folder.Fullname -force -ErrorAction SilentlyContinue | where {$_.psiscontainer -eq $false}

                if($Files -ne $null) {

                    Foreach ($File in $Files) {
					    
                        $filesCount = $filesCount + 1

                        $FolderSizeGB = "{0:N2}" -f ($FSOFolder.size / 1GB)
                        $FolderSizeMB = "{0:N2}" -f ($FSOFolder.size / 1MB) 
                        $FolderSizeKB = "{0:N2}" -f ($FSOFolder.Size / 1KB)

                        $File_ACL = get-acl $File.Fullname -ErrorAction SilentlyContinue
                        
                        #ForEach ($File_Access in $File_ACL.Access) {
                        
                        $FileOwner = $File_ACL.Owner
                        #$FileAccess = $File_ACL.Access
						$OutInfo = "FILE;""" + $CleanFolderName + """;""" + $File.Name  + """;" + $File.CreationTime + ";" + $File.LastAccessTime + ";" + $File.LastWriteTime + ";" + $FolderSizeGB + ";" + $FolderSizeMB + ";" + $FolderSizeKB + ";" + "N/A" + ";" + $File.Extension + ";" + $FileOwner + ";" + $File.IsReadOnly + ";" + ($File_ACL.Access).IdentityReference + ";" + ($File_ACL.Access).FileSystemRights + ";" + ($File_ACL.Access).IsInherited
                        
                        Add-Content -Value $OutInfo -Path $outFile
                        #}
                    }


                  }
            }
        }
    }

    Write-host ""
    Write-host "-------------------------------------------------------------------"
    write-host "------------------------------ Done! ------------------------------"
    Write-host "-------------------------------------------------------------------"
    Write-Host "$foldersCount folders / $filesCount files analyzed"
    Write-host ""
    Write-host "-------------------------------------------------------------------"
    Write-host "--------------------- The folder scanned is: ----------------------"
    Write-host "-------------------------------------------------------------------"
    write-host $path
    Write-host ""
    Write-host "-------------------------------------------------------------------"
    Write-host "----------------- The out-put Folder - File is: -------------------"
    Write-host "-------------------------------------------------------------------"
    write-host $FileName
    Write-host "-------------------------------------------------------------------"
}

$fso = New-Object -comobject Scripting.FileSystemObject
LoopSubFoldersAndFiles($path)