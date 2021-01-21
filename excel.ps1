<#
.SYNOPSIS
Script that will search multiple values from a text file in multiple excel files
.DESCRIPTION
** This script will take the multiple values from a text file and search in the multiple excel files.
** On the first run, it will create 'search values.txt' file unless one exists.
** This script will create new result excel files with "*_RESULT.xlsx" extension. Eg. MySheet_RESULT.xlsx		  
** The search results are tagged with Name Ranges to easily locate the values found, afterwhich you can remove
   these cell names from Formulas >> Name Manager.											  
** You can enable the Out-Grid View if Microsoft .net framework 3.5 is installed, by -Grid parameter. 
** Also, you can choose to automatically open newly saved file with -OpenFile parameter.			  
Example usage:																					  
.\Excel_Search.ps1 -Folder c:\myFolder -Recurse -Color -OpenFile			  				  
Author: phyoepaing3.142@gmail.com
Country: Myanmar(Burma)
Released: 07/01/2016
.EXAMPLE
.\Excel_Search.ps1 -Folder c:\myFolder -Recurse -Color -OpenFile
This will recursively search excel files in the given directory for the values listed in text file. It will 
Colorize each cell with dark blue color(color index=41) that is found and it will open each new result excel 
file automatically.
.EXAMPLE
.\Excel_Search.ps1 -File c:\myFolder\MySheet.xlsx -Grid
This will search the multiple values from text file in the single excel file. The search output will displayed
in Grid-View.
.PARAMETER Folder
Here, put the folder name in which multiple Microsoft Excel files exist.
.PARAMETER File
You can put the file name of the Microsoft Excel file you want to search in.
.PARAMETER Recurse
This parameter will recursively search Multiple Excel files in the folder, but it will skip
the new files created by this script
.PARAMETER Color
This parameter will colorize with the dark blue color (color index=41)the found values in result sheet.
It is useful when you filter the column by color for easy viewing.
.PARAMETER Grid
This will display the search results in Grid View. Microsoft .NET Framework 3.5 is needed to use this feature.
Otherwise, do not use this parameter.
.PARAMETER OpenFile
This will automatically open the newly saved result files, after the scipt is run.
.LINK
You can find this script and more at: https://www.sysadminplus.blogspot.com/
#>
############################################################################################################################

param( [switch]$Recurse,[switch]$Grid,[switch]$Color,[switch]$OpenFile,[String]$Folder,[String]$File )

$ErrorActionPreference='silentlycontinue'        ## Disable some built-in Errors about the run-time exception when we do excel search

################### Check if $File parameter exists or $Folder parameter exists and collect file names ###########
	if ($File)
		{
		if ((gci $File | out-null) -eq '0' -OR $File.LastIndexOf('.xlsx') -eq '-1') { $File="$File`.xlsx" }		## append the file extension if doesn't exists
		$Excel_files=$File																						## put the single file into the variable
		}
	elseif ( $Folder -AND $recurse )
		{
		$Excel_files=(gci $Folder -File *.xlsx -Exclude *SEARCH_RESULT* -Recurse).FullName
		}
	elseif ($Folder)
		{
		$Excel_files=(gci $Folder -File *.xlsx -Exclude *SEARCH_RESULT*).FullName
		}
	else
		{
		Write-Host -fore Red "No files or folder parameter is specified."
		$Excel_files=$null
		}

	if (Test-Path 'search values.txt')				
	{ 
			
		if ($Excel_files)									## Continue the operation if single/multiple files exists
			{
		$Found_Values=@()
		$Output_Table=@()
		$Properties=@{'Search Value'="";'Found Value'="";'Location'="";'Sheet Name'="";'File Name'="";'Status'="";'Name Range'="";'Full Name'="" }
		$Object = New-Object -TypeName PsObject -Prop $Properties
		$i=1
		
	$Excel_files | foreach {													## loops through each excel file to search value
			$Found_In_Excel=0
			$File_info = gci $_
			$File_Directory=$File_info.DirectoryName
			$File_Name = $File_info.Name
			$File_Full_Name = $File_info.FullName
			$File_ext = $File_info.Extension
			$File_Name_wt_ext = $File_Name.TrimEnd($File_ext)

			$Excel = New-Object -ComObject Excel.Application						## creating one excel app instance by comobject method
			$Excel.Visible = $false
			$Workbook = $Excel.Workbooks.Open($File_info)							## open the excel file to search
			
			$All_Search_Values=Get-Content 'search values.txt'
			$All_Search_Values_Trimmed= $All_Search_Values | foreach { $_.Trim()}	## Trim white spaces of search values
			$All_Search_Values_Trimmed= ($All_Search_Values_Trimmed | Group).Name	## Remove Duplicate of search values
		
			foreach ($current_find in $All_Search_Values_Trimmed)
				{ 
					$Excel.Worksheets | select -expandproperty index | foreach {	## Loops the worksheets to search the value
						
						$MySheet = $Workbook.Worksheets.Item($_)
						$Range = $MySheet.Range("A:KK")								## define the search area, possibly you can extend the area if you want
						$Target = $Range.Find($current_find)						## search the string
						$First = $Target
						
							Do
							{	
								$row=$Target.row
								$column=$Target.column
															
								if($Target)
								 {
						   		  $Found_In_Excel = 1											## Note that this excel file has found values
								  $MySheet.cells.item($row,$column).name="Found_$i"				## Tag the Name of found cell for easy location
								  $Object."Search Value"=$current_find
								  $Object."Found Value"=$Target.value2
								  $Object.Location=$Target.AddressLocal()
								  $Object."Sheet Name"=$MySheet.Name
								  $Object."File Name" = $File_Name
								  $Object.Status='Found'
								  $Object."Name Range"="Found_$i"
								  $Object."Full Name"=$File_Full_Name
								  $Output_Table+= $Object | select "Search Value","Found Value",Location,"Sheet Name","File Name",Status,"Name Range","Full Name"
								  $Found_Values+=$Target.value2									##count the number of objects found
								  $i++
								  }
								if($Color)
								{ $MySheet.Cells.Item($row,$column).Interior.ColorIndex=41}	## paint the cell's wall, you can change the color index here ;P
								$Target = $Range.FindNext($Target)								## do another search loop
							}
								While ($Target -ne $NULL -and $Target.AddressLocal() -ne $First.AddressLocal())		## do search until no more values is found -AND continue search after the first match
						}
				}
			if ($Found_In_Excel)
				{
				$Workbook.SaveAs("$File_Directory`\$File_Name_wt_ext SEARCH_RESULT$File_ext")
				$Excel.Quit()
				}
			else
				{
				$Excel.Quit()
				}
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
			if ($OpenFile -AND $Found_In_Excel)
			{
			$New_Excel=New-Object -ComObject Excel.Application
			$New_Excel.Visible = $true
			$Workbook = $New_Excel.Workbooks.Open("$File_Directory`\$File_Name_wt_ext SEARCH_RESULT$File_ext")
			}
		}
		
			$Found_Values_Unique=$Found_Values | Unique
			$All_Search_Values_Trimmed | foreach { if ( $Found_Values -NotContains $_) {[array]$NotFound_Values+=$_}}		## Extract the words which are not found
			
			if ($NotFound_Values) {
					$NotFound_Values | foreach { 
						$Object."Search Value"=$_;$Object.Status="NOT Found"						
						$Output_Table+= $Object | select "Search Value",Status								## Put Not Found values into array of  objects
									}																
						}
			Write-Host -Fore Magenta "`n`nSearch Values are: $($All_Search_Values_Trimmed -join(','))"
			If($Grid) {
				$Output_Table | select * | Out-GridView -Title "My Excel Search Results"					## Grid-view to easily categorize outputs
				}
			else {
				$Output_Table | ft
				}
				
			if ($Found_Values)
				{
				Write-Host -Fore Green "Total Number of values found: $($Found_Values.count)"
				}

			if ($NotFound_Values)
				{
				Write-Host -Fore Red "`nTotal Number of values NOT found: $($NotFound_Values.count)"
				Write-Host "`nThe following values are NOT Found: $($NotFound_Values -join (','))`n"
				}
		}
		else
			{ 
			Write-Host -fore red "No Excels are found to search"
			}
	}
	else
	{
	Write-Host -fore yellow "The search value txt file is not created. Now creating one..."
	New-Item -type file 'search values.txt' | Out-Null
	}