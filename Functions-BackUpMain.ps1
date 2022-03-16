<#
.SYNOPSIS
	The script includes several useful functions to manage backups features.
.DESCRIPTION
	Assumptions:
	1. Powershell 3.0 is required: the Parameter -file for Get-ChildItem is used and only available since version 3.0 of Windows PowerShell
	2. There are no changes to folders in scope of this script during run time
	This script includes these functions:
	* Show-Result
	* Backup-OneNote
	* Convert-FolderToZip
	* ZipCopyAndMove-Folder
	* Find-FileFolders
	* Compare-Folders
	* Backup-Folder
	* Find-DuplicateFiles
	* Get-TooLongPaths
	These settings are defined within this script and can be adapted when using it:
	* $global:CONSTANT_Path7zipExe1
	* $global:CONSTANT_Path7zipExe2
.LINK
	https://github.com/Stuxnerd/PsBuS
.NOTES
	VERSION: 0.9.8 - 2022-03-16

	AUTHOR: @Stuxnerd
		If you want to support me: bitcoin:19sbTycBKvRdyHhEyJy5QbGn6Ua68mWVwC

	LICENSE: This script is licensed under GNU General Public License version 3.0 (GPLv3).
		Find more information at http://www.gnu.org/licenses/gpl.html

	TODO: These tasks have to be implemented in the following versions:
	till version 1.0 - additional features, testing and documentation
	* Wenn Ziel ein Ordner und Quelle eine Datei ist, prüfen, ob der Ordner leer ist – wenn ja löschen sonst Fehler / Farbe: Magenta, um Fehler aufzuzeigen (sollte Funktion Find-FileFolders obsolet machen)
	* default way to use trim/add \ for folders + default way to find parent folder
	* add "Ignore" as action for ActionIfFileExists and separate it to an enummeration (prevent double code in script)
	* rename variable TestFile
	* common names for parameters like FileName1 and File1
	* check, if Test-Path is used with -Type whereever useful
	* check if a file is locked (e. g. pst archives, if outlook is still running): Ask, if wait and try again or ignore (Warning)
	* test the counters
	* show progress of long duration tasks with progress bar
	* document features and functions in SourceForge Wiki and in this scipt
	till *version 1.1 - optional features
	* enhance exception management with traps, throw, $ErrorActionPreference, etc. (e.g. if the file name length is too big)
	* handle the limitation of maximum filename and foldername length in NTFS (in destination, archive and with timestamps)
	* handle path names like %APPDATA% or "." or ".."
	* use transactions to prevent curruption during file transfers
#>

#####################################
#USAGE OF EXTERNAL FUNCTION PACKAGES#
#####################################

. ../PsBuS/Functions-BackUpSupport.ps1


##################################
#GLOBAL VARIABLES - RETURN VALUES#
##################################
#the global variables are used to save the return values of functions; they are just for usage of the funtions

#global varibale to save the file name of the zip file created by Convert-FolderToZip
[String]$global:RETURNVALUE_ZipFolder = $null


####################################
#GLOBAL VARIABLES - VARIABLE VALUES#
####################################
#these values are used during execution, but are independent from a single fuction invocation

#Counter for Errors
#$global:VARIABLE_ErrorCounter is imported from Functions-BackUpSupport.ps1

#For the function Show-Report
#Counter for files moved
[int]$global:VARIABLE_MoveCounter = 0
#Counter for files copied
[int]$global:VARIABLE_CopyCounter = 0

#############################################
#VARIABLES FOR THE SETTING - CONSTANT VALUES#
#############################################
#these values define the configuration of the script; the might be overwritten by an external script which is using included functions

#paths to 7zip.exe for the function Convert-FolderToZip
[String]$global:CONSTANT_Path7zipExe1 = "C:\PortableApps\7-ZipPortable\App\7-Zip64\7z.exe"
[String]$global:CONSTANT_Path7zipExe2 = "C:\Program Files\7-Zip\7z.exe"


<#
.SYNOPSIS
	Show the results of certain backup jobs.
.DESCRIPTION
	Show the results of certain backup jobs based on global variables used by other cmdlets in this file.
.PARAMETER $ReadHost
	if used, the user has to klick a button to go on
.DESCRIPTION
	Show the number of errors, moved and copied files
#>
function Show-Result {
	Param (
		[Parameter(Mandatory=$false, Position=0)]
		[switch]
		$ReadHost
	)
	#output for copied and moved files
	if ($global:VARIABLE_MoveCounter -eq 0 -and $global:VARIABLE_MoveCounter -eq 0) {
		Trace-LogMessage -Message "No files were copied or moved." -Indent 0 -Level 0 -MessageType Confirmation
	} else {
		Trace-LogMessage -Message "$global:VARIABLE_MoveCounter files were moved and $global:VARIABLE_CopyCounter files were copied." -Indent 0 -Level 0 -MessageType Confirmation
	}
	#output for errors
	if($global:VARIABLE_ErrorCounter -gt 0) {
		Trace-LogMessage -Message "$global:VARIABLE_ErrorCounter errors occured." -Indent 0 -Level 0 -MessageType Error
		$global:VARIABLE_ErrorCounter++
	} else {
		Trace-LogMessage -Message "No Errors occured." -Indent 0 -Level 0 -MessageType Confirmation
	}
	#if requested, wait
	if ($ReadHost) {
		Read-Host "Press a button to go on"
	}
}


<#
.SYNOPSIS
	Backup OneNote
.DESCRIPTION
	Backup notebooks of OneNote (only desktop application).
	Assumption, that no other processes manipulates the number of files in DestinationPath. This may result in a endless loop.
.PARAMETER DestinationPath
	path to save OneNote files
.PARAMETER $OneNoteExportCheckTimeMilliSecods
	The time to wait for saving of a OneNote export - after this time a cyclic check procedure will meter the success of the export. Small values might result in more TracegLoging.
	The default value is 1000 ms (1 second)
.SOURCES
	https://newyear2006.wordpress.com/2011/06/19/onenote-mit-powershell-bearbeiten/
	https://stackoverflow.com/questions/31526436/onenote-one-to-microsoft-word-doc-programatically
.TODO
	Check, if folder ends with \
#>
Function Backup-OneNote {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$DestinationPath,
		[Parameter(Mandatory=$false, Position=1)]
		[int]
		$OneNoteExportCheckTimeMilliSecods = 1000
	)
	#create path
	TestAndCreate-Path -FolderName $DestinationPath

	#first check, if the destination exists and is a folder
	if (Test-Path -Path $DestinationPath -PathType Container) {
		#preparation for the check
		#the finalization of this method will be metered in two ways: counting the files and testing, if the files have been created
		#Number of files in destination path before the export
		$FileCountInFolderStart = (Get-ChildItem -Path $DestinationPath -File).Count
		#this number is initalized and will be increased by each file, which has to be exported
		$FilesToCreate = 0
		#this ArrayList will save all filenames, that have to be checked
		[System.Collections.ArrayList]$FilesToCreateList = @()
		Trace-LogMessage -Message "There are $FileCountInFolderStart files in the folder '$DestinationPath' before OneNote export." -Indent 5 -Level 10

		#preparation for the export
		#create OneNote object to read all notebooks
		$OneNote = New-Object -ComObject OneNote.Application
		[xml]$xml = $null
		$OneNote.GetHierarchy($null,[Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsNotebooks, [ref] $xml)
		#http://msdn.microsoft.com/en-us/library/ff966473.aspx
		$Notebooks = $xml.Notebooks.Notebook

		#https://msdn.microsoft.com/en-us/library/ff966473.aspx (PublishFormat)
		$Format = [Microsoft.Office.Interop.OneNote.PublishFormat] 1 #pfOneNotePackage: .onepkg
		$FormatEnding = ".onepkg"
		#run through all notebooks
		foreach($Notebook in $Notebooks) {
			$Name = $Notebook.Name
			[String]$LastModified = $notebook.lastModifiedTime.Replace(":","-") #date of last modification in format "2015-07-20T15:03:03.000Z"
			#cut of the last characters
			$LastModified = $LastModified.Substring(0,19)
			#adapt the format, so that it is equal to the functions Add-TimeStampToFileName and Get-TimeStamp
			$LastModified = $LastModified.Replace("T","--")

			#build file name with timestamp of last modification
			$Filename = $DestinationPath + $Name + " (" + $LastModified + ")" + $FormatEnding
			if (Test-Path $Filename) {
				Trace-LogMessage -Message "No export of $Name, because '$Filename' exists already" -Indent 5 -Level 8 -MessageType Confirmation
			} else {
				Trace-LogMessage -Message "Export $Name to '$Filename'" -Indent 5 -Level 8
				#it is necessary to wait for the creation of this file - so this file has to be added to the list
				$FilesToCreate++ #wait for one more file
				$FilesToCreateList.Add($Filename) | Out-Null #save the filename for testing the file existance later

				#ID is required for export
				$Id = $Notebook.ID
				$Exporter = ""
				#the export step
				$OneNote.Publish($Id, $Filename, $Format, $Exporter)
				#https://msdn.microsoft.com/en-us/library/office/gg649853%28v=office.14%29.aspx
				Trace-LogMessage -Message "Export of $Name was initiated" -Indent 3 -Level 1
			}
		}

		#after starting all exports it is required to check, if the exports were successfull
		#a cyclic check will be performed in a defined periode of time
		#the check consists of two methods
		$WaitForFileSave1 = $true
		$WaitForFileSave2 = $true

		#a shortcut will prevent waiting, if nothing is to do
		if ($FilesToCreate -eq 0 -and $FilesToCreateList.Count -eq 0) {
			$WaitForFileSave1 = $false
			$WaitForFileSave2 = $false
		}

		#test if export was successfull
		#only if both conditions are fulfilled, the export was sucessfull
		while($WaitForFileSave1 -or $WaitForFileSave2) {
			#first criteria: count the number of files in folder
			$FileCountInFolderEnd = (Get-ChildItem -Path $DestinationPath -File).Count
			#now the number of exports files must exist additionally
			if ($FileCountInFolderStart + $FilesToCreate -eq $FileCountInFolderEnd) {
				Trace-LogMessage -Message "$FilesToCreate files were exported" -Indent 5 -Level 5
				#first criteria to end the endless loop was successfull
				$WaitForFileSave1 = $false
			} else {
				$NumberOfMissingFiles = $FileCountInFolderStart + $FilesToCreate - $FileCountInFolderEnd
				Trace-LogMessage -Message "Waiting for $NumberOfMissingFiles files to be exported ..." -Indent 5 -Level 8
			}

			#second criteria: checking the existance of all files - the list of files to check is empty
			if ($FilesToCreateList.Count -eq 0) {
				$WaitForFileSave2 = $false
				Trace-LogMessage -Message "Waiting for no files to be exported." -Indent 5 -Level 8
			} else {
				#a copy of the ArrayList with the file names is required, because manipulating the ArrayList while iterating it is not allowed
				[System.Collections.ArrayList]$FileListEditable = @()
				$FilesToCreateList | ForEach-Object {$FileListEditable.Add($_) | Out-Null}
				#iterating through the copy to check, if all files exist
				foreach ($TestFile in $FileListEditable) {
					#if the file exists, it was exported successfully
					if (Test-Path -Path $TestFile) {
						#and no further check is required - it can be deleted from the list to check
						$FilesToCreateList.Remove($TestFile)
						Trace-LogMessage -Message "Export of '$TestFile' was sucessfull" -Indent 0 -Level 1 -MessageType Confirmation
					} else {
						Trace-LogMessage -Message "'$TestFile' exists not yet." -Indent 5 -Level 8
					}
				}
			}
			#TODO: This check is waiting one cycle too long
			Trace-LogMessage -Message "Wait $OneNoteExportCheckTimeMilliSecods milliseconds before next check." -Indent 3 -Level 5
			Start-Sleep -Milliseconds $OneNoteExportCheckTimeMilliSecods
		}
	} else {
		if (-NOT (Test-Path $DestinationPath)) {
			Trace-LogMessage -Message "The folder '$DestinationPath' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $DestinationPath -PathType Container)) {
			Trace-LogMessage -Message "'$DestinationPath' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}


<#
.SYNOPSIS
	Zip the content of a folder.
.DESCRIPTION
	Zip the content of a folder.
	Return the name of the created file in global variable $global:RETURNVALUE_ZipFolder
.PARAMETER SourcePath
	Path to Folder, which has to be zipped
.PARAMETER DestinationPath
	Path to Folder, where to save the zipped file
	The default value is the parent folder of SourcePath
.PARAMETER SourcePath
	Path to Folder, which has to be zipped
	The default value is the name of SourcePath
.PARAMETER ZipMethod
	'Zip','7z'
	method to zip folder content
.PARAMETER Ultra
	is only valid for 7z method
.PARAMETER SplitSizeMB
	will split the file into chunks of X MB. x has to be an int
.PARAMETER ActionIfFileExists
	'Overwrite','Error','Ignore'
	Overwrite: delete existing file, then zip folder
	Error message, but do nothing
	Ignore: just do nothing
	The default value is Overwrite
.PARAMETER Path7ZipExe
	Path to 7zip.exe
	If not defined, the script will lookup the default paths "C:\PortableApps\7-ZipPortable\App\7-Zip64\7z.exe" or "C:\Program Files\7-Zip\7z.exe"
.PARAMETER Password7z
	Password for the file. The password will not be logged.
	If not defined, the zip file will not be encrypted
	It will only be used for 7z files, not for ZIP files
.EXAMPLE
	Convert-FolderToZip -SourcePath "C:\Foo\" -DestinationPath "D:\" -DestinationName "FooD" -ZipMethod Zip -ActionIfFileExists Error
	Will pack the folder of C:\Foo\ to D:\FooD.zip. If this file exists, it will not be packed, but a error message will be shown to the user
.EXAMPLE
	Convert-FolderToZip -SourcePath "C:\Foo\" -ZipMethod 7z -Ultra -ActionIfFileExists Overwrite
	Will pack the folder C:\Foo\ to C:\Foo.7z. The package will be sparse to storage. If the file already exists, it will be deleted and then new packed.
.EXAMPLE
	$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
	Convert-FolderToZip -SourcePath "C:\Foo\" -ZipMethod 7z -Password7z $SecurePassword
	Will pack the folder C:\Foo\ to C:\Foo.7z, and encrypt the file with the given password.
.EXAMPLE
	$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
	Convert-FolderToZip -SourcePath "C:\Foo\" -Password7z $SecurePassword
	Will pack the folder C:\Foo\ to C:\Foo.zip, but not encrypt the file
#>
Function Convert-FolderToZip {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$SourcePath,
		[Parameter(Mandatory=$false, Position=1)]
		[String]
		$DestinationPath = "",
		[Parameter(Mandatory=$false, Position=2)]
		[String]
		$DestinationName = "",
		[Parameter(Mandatory=$false, Position=3)]
		[ValidateSet('Zip','7z')]
		[String]$ZipMethod = 'Zip',
		[Parameter(Mandatory=$false, Position=4)]
		[Switch]
		$Ultra,
		[Parameter(Mandatory=$false, Position=5)]
		[int]
		$SplitSizeMB,
		[Parameter(Mandatory=$false, Position=6)]
		[ValidateSet('Overwrite','Error','Ignore')]
		[String]$ActionIfFileExists = 'Overwrite',
		[Parameter(Mandatory=$false, Position=7)]
		[String]
		$Path7ZipExe = "",
		[Parameter(Mandatory=$false, Position=8)]
		[SecureString]
		$Password7z = $null
	)
	#reset global variable
	$global:RETURNVALUE_ZipFolder = $null

	#check if a passowrd is set, but will be ignored
	if ($Password7z -ne $null -and $ZipMethod -eq 'Zip') {
		Trace-LogMessage -Message "A password was set, but will not be used for ZIP files" -Indent 0 -Level 0 -MessageType Warning
	}

	#check if SourcePath exists and is a folder
	if(Test-Path -Path $SourcePath -PathType Container) {
		#define default paths and names, if necessary
		if ($DestinationPath -eq "") {
			#if no destination is defined, the parent of the source will be used
			$DestinationPath = (Get-Item -Path $SourcePath).Parent.FullName
		} else {
			#if a destination is defined, it has to be an existing path
			if (-not (Test-Path -Path $DestinationPath -PathType Container)) {
				if (-not (Test-Path -Path $DestinationPath)) {
					Trace-LogMessage -Message "'$DestinationPath' does not exist." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				} elseif (-not (Test-Path -Path $DestinationPath -PathType Container)) {
					Trace-LogMessage -Message "'$DestinationPath' is not a folder." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				}
				#otherwise we also use the default value
				$DestinationPath = (Get-Item -Path $SourcePath).Parent.FullName
				Trace-LogMessage -Message "'$DestinationPath' is the destination for compression." -Indent 0 -Level 0 -MessageType Warning
			}
		}
		if ($DestinationName -eq "") {
			#if no name is defined, the foldername of source will be used
			$DestinationName = (Get-Item -Path $SourcePath).Name
		}

		#distinguish the chosen method to check the file existance and proceed as requested
		$Destination = ""
		if ($ZipMethod -eq 'Zip') {
			$Destination = $DestinationPath.TrimEnd('\') + "\" + $DestinationName + ".zip"
		} elseif ($ZipMethod -eq '7z') {
			$Destination = $DestinationPath.TrimEnd('\') + "\" + $DestinationName + ".7z"
		}
		#check if the file exists already
		if (Test-Path -Path $Destination) {
			#handling as requested
			if ($ActionIfFileExists -eq 'Overwrite') {
				Trace-LogMessage -Message "File '$Destination' exists already - it will be deleted" -Indent 5 -Level 5
				Remove-Item -Path $Destination -Force
			} elseif ($ActionIfFileExists -eq 'Error') {
				Trace-LogMessage -Message "File '$Destination' exists already" -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
				#abort current function
				return
			} elseif ($ActionIfFileExists -eq 'Ignore') {
				Trace-LogMessage -Message "File '$Destination' exists already - it will not be zipped again" -Indent 5 -Level 5
				#abort current function
				return
			} else {
				Trace-LogMessage -Message "not allowed action for existing file" -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
			}
		}

		#distinguish the chosen method to pack folder content
		if ($ZipMethod -eq 'Zip') {
			Trace-LogMessage -Message "Start to zip content of '$SourcePath' to file '$Destination'" -Indent 3 -Level 1
			#zip the file
			Add-Type -assembly "system.io.compression.filesystem"
			[io.compression.zipfile]::CreateFromDirectory($SourcePath, $Destination)
			Trace-LogMessage -Message "The content of '$SourcePath' was zipped to file '$Destination'" -Indent 3 -Level 1 -MessageType Confirmation
			#return the filename of the zip file, which was created
			[String]$global:RETURNVALUE_ZipFolder = $Destination
			return
		} elseif ($ZipMethod -eq '7z') {
			#https://sevenzip.osdn.jp/chm/cmdline/commands/add.htm
			#https://sevenzip.osdn.jp/chm/cmdline/switches/method.htm
			#define the arguments for 7zip.exe
			#a is the argument to add files to an archive
			#”” are necessary, because the path names may include spaces
			$ArgumentList = "a `"$Destination`" `"$SourcePath`""
			$ArgumentListLog = "a `"$Destination`" `"$SourcePath`""
			#if requested, the parameter for ultra compression will be used (less storage, more computing time), this always includes multithreading
			if ($Ultra) {
				$ArgumentList = $ArgumentList + " -mx=9 mt=on"
				$ArgumentListLog = $ArgumentListLog + " -mx=9 mt=on"
			}
			#if requested, the parameter for splitting the compressed file will be used
			if ($SplitSizeMB) {
				$ArgumentList = $ArgumentList + " -v" + $SplitSizeMB + "m"
				$ArgumentListLog = $ArgumentListLog + " -v" + $SplitSizeMB + "m"
			}
			#if requested, the file will be encrpted. Encrpytion alwas includes a hidden file structure
			if ($Password7z -ne $null) {
				$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password7z)
				$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
				$ArgumentList = $ArgumentList + " -p$PlainPassword -mhe=on"
				$ArgumentListLog = $ArgumentListLog + " -pPASSWORD -mhe=on"
			}

			Trace-LogMessage -Message "Use 7zip.exe with arguments: $ArgumentListLog" -Indent 5 -Level 10

			#find the path to 7zip.exe
			$PathExe = ""
			#first priority to the argument by the user
			if (($Path7ZipExe -ne "") -and (Test-Path -Path $Path7ZipExe)) {
				$PathExe = $Path7ZipExe
			} elseif (Test-Path -Path $global:CONSTANT_Path7zipExe1) {
				$PathExe = $global:CONSTANT_Path7zipExe1
			} elseif (Test-Path -Path $global:CONSTANT_Path7zipExe2) {
				$PathExe = $global:CONSTANT_Path7zipExe2
			}
			Trace-LogMessage -Message "Found 7zip.exe at: '$PathExe'" -Indent 5 -Level 10

			Trace-LogMessage -Message "Start to zip content of '$SourcePath' to file '$Destination'" -Indent 5 -Level 1
			#execute the processing with 7zip.exe and the arguments/parameters
			$process = (Start-Process -FilePath $PathExe -ArgumentList $ArgumentList -PassThru -Wait)
			#TODO: read return code to detect errors
			Trace-LogMessage -Message "The content of '$SourcePath' was zipped to file '$Destination' with return value $process" -Indent 3 -Level 1 -MessageType Confirmation
			#return the filename of the 7z file, which was created
			[String]$global:RETURNVALUE_ZipFolder = $Destination
			return
		} else {
			Trace-LogMessage -Message "The chosen method for zipping is not supported" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	} else {
		if (-not (Test-Path -Path $SourcePath)) {
			Trace-LogMessage -Message "The folder '$SourcePath' does not exist for compression." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $SourcePath -PathType Container)) {
			Trace-LogMessage -Message "'$SourcePath' is not a folder - only folder can be compressed." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	TODO
.PARAMETER SourcePath
	TODO
.PARAMETER CopyPath
	TODO
.PARAMETER MovePath
	TODO
.EXAMPLE
#>
Function ZipCopyAndMove-Folder {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$SourcePath,
		[Parameter(Mandatory=$false, Position=1)]
		[String]
		$CopyPath = "",
		[Parameter(Mandatory=$false, Position=2)]
		[String]
		$MovePath = ""
	)
	#check if folder exists and is a folder
	if (Test-Path -Path $SourcePath -PathType Container) {
		#zip folder
		Convert-FolderToZip -SourcePath $SourcePath -ZipMethod 7z -Ultra -ActionIfFileExists Error
		#rename created zip file
		$NewFileName = Add-TimeStampToFileName -FileName $global:RETURNVALUE_ZipFolder

		#copy, if requested
		if ($CopyPath -ne "") {
			#check path
			if (Test-Path -Path $CopyPath -PathType Container) {
				Copy-Item -Path $NewFileName -Destination $CopyPath
				Trace-LogMessage -Message "Copied '$NewFileName' to '$CopyPath'" -Indent 5 -Level 3
				$global:VARIABLE_CopyCounter++
			} else {
				if (-not (Test-Path -Path $CopyPath)) {
					Trace-LogMessage -Message "'$CopyPath' does not exist." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				} elseif (-not (Test-Path -Path $CopyPath -PathType Container)) {
					Trace-LogMessage -Message "'$CopyPath' is not a folder." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				}
			}
		}
		#move, if requested
		if ($MovePath -ne "") {
			#check path
			if (Test-Path -Path $MovePath -PathType Container) {
				Move-Item -Path $NewFileName -Destination $MovePath
				Trace-LogMessage -Message "Moved '$NewFileName' to '$MovePath'" -Indent 5 -Level 3
				$global:VARIABLE_MoveCounter++
			} else {
				if (-not (Test-Path -Path $MovePath)) {
					Trace-LogMessage -Message "'$MovePath' does not exist." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				} elseif (-not (Test-Path -Path $MovePath -PathType Container)) {
					Trace-LogMessage -Message "'$MovePath' is not a folder." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				}
			}
		}
	} else {
		if (-not (Test-Path -Path $SourcePath)) {
			Trace-LogMessage -Message "'$SourcePath' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $SourcePath -PathType Container)) {
			Trace-LogMessage -Message "'$SourcePath' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	Sometimes instead of copiing files a folder with the filename will be created. This will be solved in the future (TODO). As workaround the folders will be detected.
.EXAMPLE
	TODO
#>
Function Find-FileFolders {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$FolderName,
		[Parameter(Mandatory=$false, Position=1)]
		[String[]]
		$FileTypes = (".docx",".doc",".xslx",".xsl",".pptx",".ppt",".pdf",".txt",".jpg",".png",".msg"), #only small letters to be used
		[Parameter(Mandatory=$false, Position=2)]
		[switch]
		$DeleteFolders
	)
	#inform the user
	Trace-LogMessage -Message "Start to seach folders that should be files at: '$FolderName'" -Indent 0 -Level 1 -MessageType Confirmation

	#first check if the folder exist and is a folder
	if (Test-Path -Path $FolderName -PathType Container) {
		#run through the child-items and check if these are folders
		$FolderList = Get-ChildItem -Path $FolderName -Directory -ErrorAction SilentlyContinue -Recurse
		foreach ($Folder in $FolderList) {
			#check foreach foldername, if it contains a . (fast check)
			if ($Folder.Name.Contains(".")) {
				Trace-LogMessage -Message "Folder '$($Folder.FullName)' contains a ." -Indent 5 -Level 5 -MessageType Info
				foreach ($Type in $FileTypes) {
					if ($Folder.FullName.ToLower().EndsWith($Type)) {
						Trace-LogMessage -Message "Folder '$($Folder.FullName)' ends like a file" -Indent 0 -Level 3 -MessageType Warning
						#if the folder contains items, it might be okay
						if ((Get-ChildItem -Path ($Folder.FullName) -Recurse).Count -eq 0) {
							Trace-LogMessage -Message "Folder '$($Folder.FullName)' ends like a file and does not contain childitems" -Indent 0 -Level 0 -MessageType Warning
							if ($DeleteFolders) {
								Remove-Item -Path $Folder.FullName -Force
								Trace-LogMessage -Message "Folder '$($Folder.FullName)' was deleted" -Indent 0 -Level 0 -MessageType Warning
							}
						}
					}
				}
			}
		}
	} else {
		#if one or both files does not exist (or ar not files) - they are not equal
		if (-NOT (Test-Path $FolderName)) {
			Trace-LogMessage -Message "Folder '$FolderName' does not exist" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $FolderName -PathType Container)) {
			Trace-LogMessage -Message "Folder '$FolderName' is not a folder" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}


<#
.SYNOPSIS
	Will compare the content of two folders.
.DESCRIPTION
	Will compare two folders and print/log the differences
.PARAMETER Folder1
	Path of the first folder.
.PARAMETER Folder2
	Path of the second folder.
.PARAMETER ExcludeExtensions
	use exclusions for not saving some file types.
	the syntax is an ArrayList of Extensions including the "." like @(".aspx"; ".txt")
.PARAMETER FilesOnly
	does not compare (empty) folders
.PARAMETER CompareFiles
	compares not only the existance of files, but also their content
.PARAMETER NoHash
	Uses parameter NoHash for file comparisson - much faster but less accurate
.PARAMETER NoSize
	Does not comare the file size.
.PARAMETER NoLastWriteTime
	Do not compare the LastWriteTime of files.
.EXAMPLE
	Compare-Folders "D:\Downloads\BackUpTest\Source\" "D:\Downloads\BackUpTest\Source\"
	will return True, if both folders exist, they have the same content
	Compare-Folders "D:\Downloads\BackUpTest\Source\" "D:\Downloads\BackUpTest\Source\" -CompareFiles
	will also check, if the files are the same (including hash value)
#>
Function Compare-Folders {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$Folder1,
		[Parameter(Mandatory=$true, Position=1)]
		[String]
		$Folder2,
		[Parameter(Mandatory=$false, Position=2)]
		[System.Collections.ArrayList]
		$ExcludeExtensions = @(),
		[Parameter(Mandatory=$false, Position=3)]
		[switch]
		$FilesOnly,
		[Parameter(Mandatory=$false, Position=4)]
		[switch]
		$CompareFiles,
		[Parameter(Mandatory=$false, Position=5)]
		[switch]
		$NoSize,
		[Parameter(Mandatory=$false, Position=6)]
		[switch]
		$NoHash,
		[Parameter(Mandatory=$false, Position=7)]
		[switch]
		$NoLastWriteTime
	)
	#inform the user
	Trace-LogMessage -Message "Start to compare the folders: '$Folder1' and '$Folder2'" -Indent 0 -Level 1 -MessageType Confirmation

	#first check if the folders exist and are folders
	if ((Test-Path -Path $Folder1 -PathType Container) -and (Test-Path -Path $Folder2 -PathType Container)) {
		#variable to save if differences have been found
		$FoundDifferences = $false
		
		#Add missing '\' to Foldernames
		if (-not $Folder1.EndsWith("\")) {
			$Folder1 = $Folder1 + "\"
		}
		if (-not $Folder2.EndsWith("\")) {
			$Folder2 = $Folder2 + "\"
		}

		#Read the content of both folders
		#if the search is restricted to files depends on the parameter FilesOnly
		if ($FilesOnly) {
			$FileList1 = Get-ChildItem -Path $Folder1 -File -ErrorAction SilentlyContinue -Recurse | Where-Object {$ExcludeExtensions -notcontains $_.Extension}
			$FileList2 = Get-ChildItem -Path $Folder2 -File -ErrorAction SilentlyContinue -Recurse | Where-Object {$ExcludeExtensions -notcontains $_.Extension}
		} else {
			$FileList1 = Get-ChildItem -Path $Folder1 -ErrorAction SilentlyContinue -Recurse | Where-Object {$ExcludeExtensions -notcontains $_.Extension}
			$FileList2 = Get-ChildItem -Path $Folder2 -ErrorAction SilentlyContinue -Recurse | Where-Object {$ExcludeExtensions -notcontains $_.Extension}
		}

		#Save the file and folder names to a list
		$FileNameList1 = [System.Collections.ArrayList] @()
		foreach ($File in $FileList1) {
			#The filename to save and compare has to include subfoldernames
			$FileName = $File.FullName.Replace($Folder1,"")
			$FileNameList1.Add($FileName) | Out-Null
		}
		#now for the second folder
		$FileNameList2 = [System.Collections.ArrayList] @()
		$FileList2 | ForEach-Object {$FileNameList2.Add([String]$_.FullName.Replace($Folder2,"")) | Out-Null}

		#iterate through the first list and compare it to the second
		foreach ($Name in $FileNameList1) {
			if(-not $FileNameList2.Contains($Name)) {
				Trace-LogMessage -Message "Files only in Folder 1: '$Name'" -Indent 3 -Level 0 -MessageType Warning
				#save, that differences have been found
				$FoundDifferences = $true
			}
		}

		#iterate through the second list and compare it to the first
		$FileNameList2 | Where-Object {-not $FileNameList1.Contains($_)} | ForEach-Object {Trace-LogMessage -Message "Files only in Folder 2: '$_'" -Indent 3 -Level 0 -MessageType Warning; $FoundDifferences = $true }

		#if requested, the files will be compared too
		if ($CompareFiles) {
			#iterate through the first list, to build up the filepaths
			foreach ($Name in $FileNameList1) {
				#compare only files, that are in both folders available (other cases are handled above)
				if($FileNameList2.Contains($Name)) {
					#build up the path to both files to compare
					$FileName1 = $Folder1 + $Name
					$FileName2 = $Folder2 + $Name
					#ensure, that both are files, not foders
					if ((Test-Path -Path $FileName1 -PathType Leaf) -and (Test-Path -Path $FileName2 -PathType Leaf)) {
						#compare files
						#usage of parameters NoSize, NoHash and NoLastWriteTime will be used if requested
						if ($NoSize -and $NoHash -and $NoLastWriteTime) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoSize -NoHash -NoLastWriteTime
						} elseif ($NoSize -and $NoLastWriteTime) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoSize -NoLastWriteTime
						} elseif ($NoHash -and $NoLastWriteTime) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoHash -NoLastWriteTime
						} elseif ($NoLastWriteTime) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoLastWriteTime
						} elseif ($NoSize -and $NoHash) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoSize -NoHash
						} elseif ($NoSize) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoSize
						} elseif ($NoHash) {
							Compare-Files -File1 $FileName1 -File2 $FileName2 -NoHash
						} else {
							Compare-Files -File1 $FileName1 -File2 $FileName2
						}
						#check, if the files are equal
						if (-not $global:RETURNVALUE_CompareFiles) {
							Trace-LogMessage -Message "File is different: '$Name'" -Indent 3 -Level 0 -MessageType Warning
							#save, that differences have been found
							$FoundDifferences = $true
						}
					} else {
						#the only case, which is important for the user is, if only one file is a folder
						if ((Test-Path -Path $FileName1 -PathType Leaf) -xor (Test-Path -Path $FileName2 -PathType Leaf)) {
							if (Test-Path -Path $FileName1 -PathType Container) {
								Trace-LogMessage -Message "'$FileName1' is a folder and '$FileName2' is a file" -Indent 3 -Level 0 -MessageType Warning
							} elseif (Test-Path -Path $FileName2 -PathType Container) {
								Trace-LogMessage -Message "'$FileName1' is a file and '$FileName2' is a folder" -Indent 3 -Level 0 -MessageType Warning
							}
						}
					}
				}
			}
		}
		if (-not $FoundDifferences) {
			#if no changes have been found, a message will be shown
			Trace-LogMessage -Message "No differences between folder '$Folder1' and '$Folder2' found" -Indent 0 -Level 1 -MessageType Confirmation
		}
	} else {
		#if one or both files does not exist (or ar not folders) - they are not equal
		if (-NOT (Test-Path $Folder1)) {
			Trace-LogMessage -Message "First folder '$Folder1' does not exist" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $Folder2)) {
			Trace-LogMessage -Message "Second folder '$Folder2' does not exist" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $Folder1 -PathType Container)) {
			Trace-LogMessage -Message "First folder '$Folder1' is not a folder" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $Folder2 -PathType Container)) {
			Trace-LogMessage -Message "Second folder '$Folder2' is not a folder" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}


<#
.SYNOPSIS
	Copy the content of a folder
.DESCRIPTION
	Will copy the content of a folder and if required also archive the content of the previous backup.
.PARAMETER SourcePath
	Path from where the file should be copied. Nothing will be changed.
.PARAMETER DestinationPath
	Path with a 1:1 copy of SourcePath. Prior files in this folder will be moved to ArchivePath.
.PARAMETER ArchivePath
	All files that are deprecated in DestinationPath will be moved to here. Reasons are: 1. File was renamed; 2. File was deleted; 3. File was updated (older File will be moved)
	This parameter is optional. If no valid ArchivePath is defined, the files won't be archived.
.PARAMETER FolderName
	optional Name for the subfolder in DestinationPath and ArchivePath to build the filesystenstructure of SourcePath.
	If not used, no subfolder will be created
.PARAMETER ExcludeExtensions
	use exclusions for not saving some file types.
	the syntax is an ArrayList of Extensions including the "." like @(".aspx"; ".txt")
.PARAMETER ShowResult
	Gives an overview, how many files were copied/moved at the end
.PARAMETER Compare
	Uses the function Compare-Folders for the folders SourcePath and DestinationPath after finishing
.PARAMETER NoSize
	Does not comare the file size.
.PARAMETER NoHash
	Does not use hashvalues for the comparission of files (only for the copy; for comparission they will still be used)
.PARAMETER NoLastWriteTime
	Do not compare the LastWriteTime of files.
.EXAMPLE
	Backup-Folder -SourcePath "D:\Downloads\BackUpTest\Source" -DestinationPath "D:\Downloads\BackUpTest\Destination" -ArchivePath "D:\Downloads\BackUpTest\Archive" -FolderName "Test1"
	Will copy the content from "D:\Downloads\BackUpTest\Source" to "D:\Downloads\BackUpTest\Destination".
	The old file from "D:\Downloads\BackUpTest\Destination" will be moved to "D:\Downloads\BackUpTest\Archive", if they are not equal to the files in "D:\Downloads\BackUpTest\Destination".
.TODO
	Die Funktionalität, dass die Datei im Ziel überschrieben wird, funktioniert bisher nur, wenn keine Archivierung statt findet, da die Datei sonst umbenannt und verschoben wird. Ziel sollte es sein, dass auch mit Archivierung ein Überschreiben möglich ist.
	Warnung, wenn einer der Pfade zu lang ist und es Probleme geben könnte.
	TryCatch-Block, falls kein lesender Zugriff (bspw. gesperrte PST-Datei) möglich ist und Meldung an den Nutzer
#>
Function Backup-Folder {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$SourcePath,
		[Parameter(Mandatory=$true, Position=1)]
		[String]
		$DestinationPath,
		[Parameter(Mandatory=$false, Position=2)]
		[String]
		$ArchivePath = "",
		[Parameter(Mandatory=$false, Position=3)]
		[String]
		$FolderName = "",
		[Parameter(Mandatory=$false, Position=4)]
		[System.Collections.ArrayList]
		$ExcludeExtensions = @(),
		[Parameter(Mandatory=$false, Position=5)]
		[switch]
		$ShowResult,
		[Parameter(Mandatory=$false, Position=6)]
		[switch]
		$Compare,
		[Parameter(Mandatory=$false, Position=7)]
		[switch]
		$NoSize,
		[Parameter(Mandatory=$false, Position=8)]
		[switch]
		$NoHash,
		[Parameter(Mandatory=$false, Position=9)]
		[switch]
		$NoLastWriteTime
	)
	#inform the user
	Trace-LogMessage -Message "Start to backup (and archive) the folder: '$FolderName' ('$SourcePath')" -Indent 0 -Level 1 -MessageType Confirmation

	#first check, if all folders exist and are folders
	if ((Test-Path -Path $SourcePath -PathType Container) -and (Test-Path -Path $DestinationPath -PathType Container)) {
		#preparation of counters
		[int]$CopyCounter = 0
		[int]$MoveCounter = 0
		[int]$DeleteCounter = 0

		#trim \ at the end of folder names
		$SourcePath = $SourcePath.TrimEnd('\')
		$DestinationPath = $DestinationPath.TrimEnd('\')
		$ArchivePath = $ArchivePath.TrimEnd('\')
		$FolderName = $FolderName.Trim('\')

		#check, if the archive is going to be used
		[boolean]$UseArchive = $false
		if ($ArchivePath -eq "") {
			Trace-LogMessage -Message "ArchivePath not defined: no files will be archived" -Indent 5 -Level 5
			$UseArchive = $false
		} else {
			#if the ArchivePath is defined, it has to be valid
			if (Test-Path -Path $ArchivePath -PathType Container) {
				$UseArchive = $true
				#attach the FolderName to ArchivePath
				$ArchivePath = $ArchivePath + "\" + $FolderName
				TestAndCreate-Path -FolderName $ArchivePath
				Trace-LogMessage -Message "Ensured, that folder '$ArchivePath' exists." -Indent 5 -Level 8
			} elseif (-not (Test-Path -Path $ArchivePath)) {
				Trace-LogMessage -Message "'$ArchivePath' does not exist." -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
				$UseArchive = $false
			} elseif (-not (Test-Path -Path $ArchivePath -PathType Container)) {
				Trace-LogMessage -Message "'$ArchivePath' is not a folder." -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
				$UseArchive = $false
			}
		}

		#if required attach the FolderName to DestinationPath
		if ($FolderName -ne "") {
			$DestinationPath = $DestinationPath + "\" + $FolderName
		}
		#create paths, if they do not exist yet
		TestAndCreate-Path -FolderName $DestinationPath
		Trace-LogMessage -Message "Ensured, that folder '$DestinationPath' exists." -Indent 5 -Level 8

		#Part A: Copy files to destination
		#Iterating through SourcePath
		$SourceFolderContent = Get-ChildItem -Path $SourcePath -Recurse -ErrorAction SilentlyContinue | Where-Object {$ExcludeExtensions -notcontains $_.Extension}
		foreach ($FileSystemItem in $SourceFolderContent) {
			$FileToBackUp = $FileSystemItem.FullName
			Trace-LogMessage -Message "Current file: '$FileToBackUp'" -Indent 1 -Level 6

			#STEP 1: Find files to copy, because they have changed
			#Check, if file/folder exists in target path
			#Define what tasks are to be done: copy AND/OR move prior version
			$NameInDestination = $FileToBackUp.Replace($SourcePath,$DestinationPath)
			$Status_MoveExistingFileBeforeCopy = $False
			$Status_CopyFile = $True
			if (Test-Path -Path $NameInDestination) {
				Trace-LogMessage -Message "'$FileToBackUp' exists!" -Indent 3 -Level 10
				#If the File already exisit in destination folder, it is necessary to compare both files (if they are files)
				if ((Test-Path -Path $NameInDestination -PathType Leaf) -and (Test-Path -Path $FileToBackUp -PathType Leaf)) {
					#usage of parameters NoSize, NoHash and NoLastWriteTime will be used if requested
					if ($NoSize -and $NoHash -and $NoLastWriteTime) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoSize -NoHash -NoLastWriteTime
					} elseif ($NoSize -and $NoLastWriteTime) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoSize -NoLastWriteTime
					} elseif ($NoHash -and $NoLastWriteTime) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoHash -NoLastWriteTime
					} elseif ($NoLastWriteTime) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoLastWriteTime
					} elseif ($NoSize -and $NoHash) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoSize -NoHash
					} elseif ($NoSize) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoSize
					} elseif ($NoHash) {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination -NoHash
					} else {
						Compare-Files -File1 $FileToBackUp -File2 $NameInDestination
					}
					#the returnvalue is in the global variable $global:RETURNVALUE_CompareFiles
					if ($global:RETURNVALUE_CompareFiles) {
						Trace-LogMessage -Message "'$FileToBackUp' is equal to '$NameInDestination' - Nothing to do!" -Indent 3 -Level 10
						$Status_CopyFile = $False
					} else { #Files are not equal
						Trace-LogMessage -Message "'$FileToBackUp' is NOT equal to '$NameInDestination' - Move existing copy!" -Indent 3 -Level 10
						$Status_MoveExistingFileBeforeCopy = $True
						$Status_CopyFile = $True
					}
				} else {
					if ((Test-Path -Path $NameInDestination -PathType Container) -and (Test-Path -Path $FileToBackUp -PathType Container)) {
						Trace-LogMessage -Message "'$FileToBackUp' and '$NameInDestination' are folders: no comparisson" -Indent 3 -Level 10
					} else {
						Trace-LogMessage -Message "'$FileToBackUp' or '$NameInDestination' is folder - the other one is a file" -Indent 0 -Level 0 -MessageType Error
						$global:VARIABLE_ErrorCounter++
					}
				}
			} else { #File does not exist till now in destination
				Trace-LogMessage -Message "'$FileToBackUp' does not exist in destination ($NameInDestination)!" -Indent 3 -Level 10
				$Status_MoveExistingFileBeforeCopy = $False
				$Status_CopyFile = $True
			}

			#STEP 2: Copy files if necessary
			#Now we do what is to be done: copy AND/OR move prior version
			if ((-NOT $Status_CopyFile) -AND $Status_MoveExistingFileBeforeCopy) {
				#TODO: This Case should not occur: ERROR HANDLING necessary (throw + trap)
				Trace-LogMessage -Message "ERROR: Moving old file without copying new one!" -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
			}
			if ($Status_CopyFile) {
				#STEP 2a: if the existing file in destination has to be moved to archive first, it is done right now
				if ($Status_MoveExistingFileBeforeCopy) {
					if ($UseArchive) {
						#as the renaming as file with the last modification date will not work properly of SharePoint 2010, the option KeepBoth is required
						Move-File -FileName $NameInDestination -SourcePath $DestinationPath -DestinationPath $ArchivePath -ActionIfFileExists KeepBoth -AddTimeStamp
						Trace-LogMessage -Message "Moved to archive: '$NameInDestination' because a new version exists." -Indent 5 -Level 3
						$MoveCounter++
						if ($error.Count -ge 1) {
							if ($error[0].Exception.GetType().FullName -eq "System.IO.PathTooLongException") {
								Trace-LogMessage -Message "Path contains too long item - NameInDestination: $NameInDestination" -Indent 0 -Level 1 -MessageType Warning
							}
							#reset error array
							$error.Clear() | Out-Null
						}
					} <#else { #no deletion, as the files will be overwritten - this supports the versioning on SharePoint servers
						Remove-Item -Path $NameInDestination -Force
						Trace-LogMessage -Message "Deleted file: '$NameInDestination' because a new version exists." -Indent 5 -Level 3
						$DeleteCounter++
					}#>
				} else {
					#nothing to move - nothing to do :-)
				}

				#STEP 2b: Copy file to destination
				#build the destination path - 
				#$NameInDestinationPathOnly = $FileSystemItem.Parent.FullName.Replace($SourcePath,$DestinationPath)
				#cut off the filename
				$NameInDestinationPathOnly = $NameInDestination.Substring(0,$NameInDestination.LastIndexOf('\'))
				#doesn't work, because the file $NameInDestination does not exist:
				#$NameInDestinationPathOnly = ((Get-ChildItem $NameInDestination).Directory).FullName.Replace($SourcePath,$DestinationPath)
				#create path if necessary
				TestAndCreate-Path -FolderName $NameInDestinationPathOnly
				Trace-LogMessage -Message "Ensured, that folder '$NameInDestinationPathOnly' exists." -Indent 5 -Level 8
				#if it is a file, it must be copied
				if (Test-Path -Path $FileToBackUp -PathType Leaf) {
					#the most import line in this script: Copy the file to destination ;-)
					Copy-Item -Path $FileToBackUp -Destination $NameInDestinationPathOnly -Force
					#if Copy-Item is not possible due to the file size if might be possible, that the file can also not be copied in explorer
					#in case of error 0x800700DF please refer to: https://answers.microsoft.com/en-us/ie/forum/ie8-windows_xp/error-0x800700df-the-file-size-exceeds-the-limit/d208bba6-920c-4639-bd45-f345f462934f
					#increase the FileSizeLimitInBytes at: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\WebClient\Parameters and restart the WebClient service
				} else {
					#if it is a folder, it will be created
					Trace-LogMessage -Message "'$FileToBackUp' is a path" -Indent 5 -Level 10
					TestAndCreate-Path -FolderName $NameInDestination
					Trace-LogMessage -Message "Ensured, that folder '$NameInDestination' exists." -Indent 5 -Level 8
				}
				#copied file are confirmed
				if (Test-Path -Path $NameInDestination -PathType Leaf) {
					Trace-LogMessage -Message "file copy successful: '$FileToBackUp'" -Indent 3 -Level 3 -MessageType Confirmation
					$CopyCounter++
				} elseif (Test-Path -Path $NameInDestination -PathType Container) {
					#nothing to do for folders
				} else {
					Trace-LogMessage -Message "ERROR while copy: '$FileToBackUp'" -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				}
			} else {
				#nothing to move - nothing to do :-)
			}
		}
		#Part B: Remove unneeded files or folders from destination
		#TODO: Avoid assumption, that the max filelength is not problem
		#Save the folders, which needed to be deleted later
		[System.Collections.ArrayList]$DeleteFolders = @()
		#Iterating through DestinationPath
		$DestinationFolderContent = Get-ChildItem -Path $DestinationPath -Recurse
		foreach ($FileSystemItem in $DestinationFolderContent) {
			$FileToCheck = $FileSystemItem.FullName
			Trace-LogMessage -Message "Current file to check for duplicity (reverse check): '$FileToCheck'" -Indent 3 -Level 6

			#Check, if file exists also in SourcePath
			$NameInSource = $FileToCheck.Replace($DestinationPath,$SourcePath)
			if (Test-Path -Path $NameInSource) {
				Trace-LogMessage -Message "'$FileToCheck' exists!" -Indent 5 -Level 10
				#if the File also exisit in source folder, it is the same (was checked in part A) - nothing to do
			} else {
				#files will be moved (or deleted if no archive is used), folders created and deleted if they are empty
				if (Test-Path -Path $FileToCheck -PathType Leaf) {
					if ($UseArchive) {
						#if the file does not exist in source folder, it should be moved from destination to archive
						#as the renaming as file with the last modification date will not work properly of SharePoint 2010, the option KeepBoth is required
						Move-File -FileName $FileToCheck -SourcePath $DestinationPath -DestinationPath $ArchivePath -ActionIfFileExists KeepBoth -AddTimeStamp
						Trace-LogMessage -Message "Moved to archive: '$FileToCheck' because it does not exist anymore in source folder." -Indent 5 -Level 3
						$MoveCounter++
					} else {
						#if the file does not exist in source folder, it should be deleted in destination
						Remove-Item -Path $FileToCheck -Force
						Trace-LogMessage -Message "Deleted file: '$FileToCheck' because it does not exist anymore in source folder." -Indent 5 -Level 3
						$DeleteCounter++
					}
				} else {
					#create folder in archive
					$NameInArchive = $FileToCheck.Replace($DestinationPath,$ArchivePath)
					TestAndCreate-Path -FolderName $NameInArchive
					Trace-LogMessage -Message "Ensured, that folder '$NameInArchive' exists." -Indent 5 -Level 8
					Trace-LogMessage -Message "folder created: '$NameInArchive'" -Indent 5 -Level 3 -MessageType Confirmation
					#if the folder contains something, it will not be deleted yet
					if ((Get-ChildItem -Path $FileToCheck -Recurse).Count -gt 0) {
						Trace-LogMessage -Message "folder '$FileToCheck' was not deleted, because of the content" -Indent 5 -Level 3 -MessageType Warning
						$DeleteFolders.Add($FileToCheck) | Out-Null
					} else {
						#delete folder in destination
						Remove-Item -Path $FileToCheck
						Trace-LogMessage -Message "folder deleted: '$FileToCheck'" -Indent 5 -Level 3 -MessageType Confirmation
					}
				}
			}
		}
		#PART C: Try to delete empty folder structures
		[boolean]$SomethingWasDeleted = $true
		while ($SomethingWasDeleted) {
			#this will be set to true, if a folder was deleted - to control the loop
			$SomethingWasDeleted = $false
			foreach ($folder in $DeleteFolders) {
				#check if the folder still exists (or was be deleted previously
				if (Test-Path -Path $folder -PathType Container) {
					Trace-LogMessage -Message "Try to delete: '$folder'" -Indent 5 -Level 10
					#check, if the folder contrains something
					if ((Get-ChildItem -Path $folder -Recurse).Count -gt 0) {
						Trace-LogMessage -Message "folder '$folder' was not deleted, because of the content" -Indent 5 -Level 10
					} else {
						#delte folder and check, if the deltetion was successful
						Remove-Item -Path $folder
						#the folder will not be deleted from the List in $DeleteFolders, because it is not good idea to change a list, which is iterated at the same time
						Trace-LogMessage -Message "folder deleted: '$folder'" -Indent 5 -Level 3 -MessageType Confirmation
						$SomethingWasDeleted = $true
					}
				}
			}
		}
		#PART D: Show the result (optional)
		if ($ShowResult) {
			if ($MoveCounter -gt 0 -or $CopyCounter -gt 0 -or $DeleteCounter -gt 0) {
				Trace-LogMessage -Message "$MoveCounter files moved to archive; $CopyCounter files copied to destination; $DeleteCounter files were deleted in destination." -Indent 1 -Level 1 -MessageType Confirmation
			} else {
				Trace-LogMessage -Message "No files were moved to archive or copied to destination or deleted." -Indent 1 -Level 1 -MessageType Confirmation
			}
		}
		#in each case the result of the Counters will be added up to the global counters
		$global:VARIABLE_MoveCounter += $MoveCounter
		$global:VARIABLE_CopyCounter += $CopyCounter

		#PART E; Compare Folders (optional)
		if ($Compare) {
			#the name of the subfolder in destination was already attached to DestinationPath
			#this will include the hashvalue independent from the parameter NoHash
			Compare-Folders -Folder1 $SourcePath -Folder2 $DestinationPath -CompareFiles -ExcludeExtensions $ExcludeExtensions
		}
	} else {
		if (-not (Test-Path -Path $SourcePath)) {
			Trace-LogMessage -Message "'$SourcePath' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $SourcePath -PathType Container)) {
			Trace-LogMessage -Message "'$SourcePath' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $DestinationPath)) {
			Trace-LogMessage -Message "'$DestinationPath' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $DestinationPath -PathType Container)) {
			Trace-LogMessage -Message "'$DestinationPath' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}

<#
.SYNOPSIS
	Find duplicate files in a folder
.DESCRIPTION
	Find duplicate files (e.g. in archive) to delete some of them manually to save storage space
.PARAMETER FolderName
	Path of the folder to be checked.
.EXAMPLE
	Find-DuplicateFiles -FolderName "C:\Datensicherung\Example"
.TODO
	switch param recuse (currently default)
	exclude Hash value computation to external function
	implement feature which finds duplicats independet from filesystem structure, based on size, hash and timestamp with the objective to save storage space in archive
	Add a function Find-Duplicates based on the hashvalue of file: save hash for each file to hashmap (with filepath); compare files: show filepaths to user
#>
Function Find-DuplicateFiles {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$FolderName
	)
	#check, if the path exists and is a folder
	if (Test-Path -Path $FolderName -PathType Container) {
		Trace-LogMessage -Message "Check $FolderName for duplicates files" -Indent 0 -Level 1 -MessageType Confirmation
		#data sructure to save all file names and hash values
		[System.Collections.HashTable]$FileHashList = @{}
		#data sructure to save all suplicate hash values
		[System.Collections.ArrayList]$Duplicates = @()
		#run through all files
		$FileList = Get-ChildItem -Path $FolderName -Recurse -File
		foreach ($File in $FileList) {
			#compute hash value of file
			[String]$HashOfFile = ""
			if ($version -ge 4) {
				#new version
				$HashOfFile = (Get-FileHash -Path ($File.FullName) -Algorithm MD5).Hash
			} else {
				#old version
				$md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
				$HashOfFile = [System.BitConverter]::ToString($md5.ComputeHash([System.IO.File]::ReadAllBytes($File.FullName))).Replace("-","")
			}
			Trace-LogMessage -Message "Hash value $HashOfFile for file $($File.FullName)" -Indent 5 -Level 8
			
			#compare, if hash value already exists in table
			if ($FileHashList.ContainsValue($HashOfFile)) {
				Trace-LogMessage -Message "Hash value $HashOfFile already in list." -Indent 3 -Level 5
				#save to list of duplicate hash values
				$Duplicates.Add($HashOfFile) | Out-Null
			}

			#save hash value and filename to hashtable
			$FileHashList.Add($File.FullName,$HashOfFile) | Out-Null
		}

		#run through duplicates and show result to user
		foreach ($HashValue in $Duplicates) {
			$HashEnumerator = $FileHashList.GetEnumerator() | Where-Object {$_.Value -eq $HashValue}
			#Output
			Trace-LogMessage -Message "Files with the hash value $HashValue :" -Indent 0 -Level 0 -MessageType Confirmation
			$HashEnumerator | ForEach-Object {Trace-LogMessage -Message "$($_.Key)" -Indent 2 -Level 0 -MessageType Warning}
		}
	} else {
		if (-not (Test-Path -Path $FolderName)) {
			Trace-LogMessage -Message "'$FolderName' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $FolderName -PathType Container)) {
			Trace-LogMessage -Message "'$FolderName' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}


<#
.SYNOPSIS
	Find to long paths.
.DESCRIPTION
	Sometimes too long paths lead to errors during backup (also by other tools). This cmdlet will help to search for these files and folders.
.PARAMETER PathName
	the path which is recursively checked for too long folders and files
.PARAMETER IsSubCall
	only for internal use (recursive call) - prevents output
.PARAMETER ThresholdFile
	threshold for path length of files before a warning
.PARAMETER ThresholdFolder
	threshold for path length of folders before a warning
.EXAMPLE
	Get-TooLongPaths -PathName "C:\Lorem ipsum dolor sit amet"
#>
Function Get-TooLongPaths {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$PathName,
		[Parameter(Mandatory=$false, Position=1)]
		[switch]
		$IsSubCall,
		[Parameter(Mandatory=$false, Position=2)]
		[int]
		$ThresholdFile = 260,
		[Parameter(Mandatory=$false, Position=3)]
		[int]
		$ThresholdFolder = 248
	)
	try {
		#check, if all folders exist and are folders
		if (Test-Path -Path $PathName) {
			#check if this is a recusrive call - if not prepare variables and reset output file
			if (-not $IsSubCall) {
				Trace-LogMessage -Message "Check path of '$PathName' for too long path names " -Level 3
				[System.Collections.ArrayList]$TooLongPathsList = @()
				[String]$PathsTooLongOutputPath = ".\PathsTooLong.txt"
				Out-File -FilePath $PathsTooLongOutputPath -InputObject "Paths with too long paths and files:"
				#reset error array
				$error.Clear() | Out-Null
			}
			#save the current pathname - if an error occurs we might be right here
			$PathNameError = $PathName

			#get all child item (folders will use a recusrive call)
			$FolderList = Get-ChildItem -Path $PathName -Directory -ErrorAction SilentlyContinue
			#if this call will produce a error (will happen even with SilentlyContinue) it will be found in the array $error
			if ($error.Count -ge 1) {
				if ($error[0].Exception.GetType().FullName -eq "System.IO.PathTooLongException") {
					Out-File -FilePath $PathsTooLongOutputPath -InputObject $PathNameError -Append
					Trace-LogMessage -Message "Path contains too long item: '$PathNameError'" -Indent 0 -Level 1 -MessageType Warning
					$TooLongPathsList.Add($PathNameError) | Out-Null
				}
				if ($error[0].Exception.GetType().FullName -eq "System.Management.Automation.RuntimeException") {
					Out-File -FilePath $PathsTooLongOutputPath -InputObject $PathNameError -Append
					Trace-LogMessage -Message "Path does not exist: '$PathNameError'" -Indent 0 -Level 1 -MessageType Warning
					$TooLongPathsList.Add($PathNameError) | Out-Null
				}
				#reset error array
				$error.Clear() | Out-Null
			}

			#get all child item - files
			$FileList = Get-ChildItem -Path $PathName -File -ErrorAction SilentlyContinue
			#if this call will produce a error (will happen even with SilentlyContinue) it will be found in the array $error
			if ($error.Count -ge 1) {
				if ($error[0].Exception.GetType().FullName -eq "System.IO.PathTooLongException") {
					Out-File -FilePath $PathsTooLongOutputPath -InputObject $PathNameError -Append
					Trace-LogMessage -Message "Path contains too long item: '$PathNameError'" -Indent 0 -Level 1 -MessageType Warning
					$TooLongPathsList.Add($PathNameError) | Out-Null
				}
				if ($error[0].Exception.GetType().FullName -eq "System.Management.Automation.RuntimeException") {
					Out-File -FilePath $PathsTooLongOutputPath -InputObject $PathNameError -Append
					Trace-LogMessage -Message "Path does not exist: '$PathNameError'" -Indent 0 -Level 1 -MessageType Warning
					$TooLongPathsList.Add($PathNameError) | Out-Null
				}
				#reset error array
				$error.Clear() | Out-Null
			}

			#iterate through subfolders
			foreach ($Folder in $FolderList) {
				if (Test-Path -Path $($Folder.FullName)) {
					$Length = (Get-Item -Path ($Folder.FullName) -ErrorAction SilentlyContinue).FullName.Length
					if ($error.Count -ge 1) {
						[String]$ErrorType = $error[0].Exception.GetType().FullName
						Trace-LogMessage -Message "Unexpected Error: '$ErrorType'" -Indent 0 -Level 0 -MessageType Error				
						#reset error array
						$error.Clear() | Out-Null
					}
					if ($Length -ge $ThresholdFolder) {
						Trace-LogMessage -Message "Path of folder '$Folder' is $Length long. " -Indent 4 -Level 1 -MessageType Warning
						Trace-LogMessage -Message "Full name: '$($Folder.FullName)'" -Indent 5 -Level 1 -MessageType Warning
						$TooLongPathsList.Add($Folder.FullName) | Out-Null
					} else {
						Trace-LogMessage -Message "Path of '$Folder' is $Length long, but not too long. " -Level 8
					}
					#recursive call:
					Get-TooLongPaths -PathName ($Folder.FullName) -IsSubCall
				} else {
					Trace-LogMessage -Message "Path of folder '$Folder' is $Length long. " -Indent 4 -Level 1 -MessageType Warning
					Trace-LogMessage -Message "Full name: '$($Folder.FullName)'" -Indent 5 -Level 1 -MessageType Warning
					$TooLongPathsList.Add($Folder.FullName) | Out-Null
				}
			}

			#iterate through files
			foreach ($File in $FileList) {
				if (Test-Path -Path $($File.FullName)) {
					$Length = (Get-Item -Path ($File.FullName) -ErrorAction SilentlyContinue).FullName.Length
					if ($error.Count -ge 1) {
						[String]$ErrorType = $error[0].Exception.GetType().FullName
						Trace-LogMessage -Message "Unexpected Error: '$ErrorType'" -Indent 0 -Level 0 -MessageType Error				
						#reset error array
						$error.Clear() | Out-Null
					}
					if ($Length -ge $ThresholdFile) {
						Trace-LogMessage -Message "Path of file '$File' is $Length long. " -Indent 4 -Level 1 -MessageType Warning
						Trace-LogMessage -Message "Full name: '$($File.FullName)'" -Indent 5 -Level 1 -MessageType Warning
						$TooLongPathsList.Add($File.FullName) | Out-Null
					} else {
						Trace-LogMessage -Message "Path of '$File' is $Length long, but not too long. " -Level 8
					}
				} else {
					Trace-LogMessage -Message "Path of file '$File' is $Length long. " -Indent 4 -Level 1 -MessageType Warning
					Trace-LogMessage -Message "Full name: '$($File.FullName)'" -Indent 5 -Level 1 -MessageType Warning
					$TooLongPathsList.Add($File.FullName) | Out-Null
				}
			}
		} else {
			#as Windows will work with an alias for too long paths, the requested folder might not exist (for Windows)
			#e.g. Lorem ipsum dolor sit amet\consetetur sadipscing elitr\sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat\sed diam voluptua\At vero eos et accusam et justo duo dolores et ea rebum
			# will be LOREMI~1\CONSET~1\SEDDIA~1\SEDDIA~1\ATVERO~1
			if (-not (Test-Path -Path $PathName)) {
				Trace-LogMessage -Message "Too long path '$PathName' does not exist." -Indent 0 -Level 0 -MessageType Warning
				$TooLongPathsList.Add($PathNameError) | Out-Null
				$global:VARIABLE_ErrorCounter++
			}
		}
	}
	catch {
		[String]$ErrorType = $error[0].Exception.GetType().FullName
		Trace-LogMessage -Message "Unexpected Error: '$ErrorType'" -Indent 0 -Level 0 -MessageType Error
	}
	finally {
		#at the end of all recursive calls the output will be written
		if (-not $IsSubCall) {
			Out-File -FilePath $PathsTooLongOutputPath -InputObject $TooLongPathsList -Append
		}
	}
}