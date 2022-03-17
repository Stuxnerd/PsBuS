<#
.SYNOPSIS
	The script includes several useful functions to manage backups features.
.DESCRIPTION
	Assumptions:
	 1. Powershell 3.0 is required: the Parameter -file for Get-ChildItem is used and only available since version 3.0 of Windows PowerShell
	 2. There are no changes to folders in scope of this script during run time
	This script includes these functions:
	* Compare-FileHash
	* Compare-Files
	* Move-File
	* Move-AllFiles
.LINK
	https://github.com/Stuxnerd/PsBuS
.NOTES
	VERSION: 0.9.4 - 2022-03-16

	AUTHOR: @Stuxnerd
		If you want to support me: bitcoin:19sbTycBKvRdyHhEyJy5QbGn6Ua68mWVwC

	LICENSE: 	This script is licensed under GNU General Public License version 3.0 (GPLv3).
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

. ../PsBuS/Functions-Logging.ps1
. ../PsBuS/Functions-Support.ps1


##################################
#GLOBAL VARIABLES - RETURN VALUES#
##################################
#the global variables are used to save the return values of functions; they are just for usage of the funtions

#global varibale to save result of the file hash value comparisson of the function: Compare-FileHash
[boolean]$global:RETURNVALUE_CompareFileHash = $False
#global varibale to save the return value of the function: Compare-Files
[boolean]$global:RETURNVALUE_CompareFiles = $False


####################################
#GLOBAL VARIABLES - VARIABLE VALUES#
####################################
#these values are used during execution, but are independent from a single fuction invocation

#Counter for Errors
[int]$global:VARIABLE_ErrorCounter = 0


#############################################
#VARIABLES FOR THE SETTING - CONSTANT VALUES#
#############################################
#these values define the configuration of the script; the might be overwritten by an external script which is using included functions

#not used in this script


#####################
#FUNCTION DEFINITION#
#####################

<#
.SYNOPSIS
	Compare the file has of two files.
.DESCRIPTION
	Compare the file has of two files.	
	Usage of global Variable to save return value
	[boolean]$global:RETURNVALUE_CompareFileHash will be true, if the files have the same hashvalue
.PARAMETER $FileName1
	file 1 (path + file name)
.PARAMETER $FileName2
	file 2 (path + file name)
.EXAMPLE
	Compare-FileHash -FileName1 "D:\Downloads\GamePC\1.txt" -FileName2 "D:\Downloads\GamePC\2.txt"
#>
function Compare-FileHash {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$FileName1,
		[Parameter(Mandatory=$true, Position=1)]
		[String]
		$FileName2
	)
	#test if both files exist and are files (not folders)
	if ((Test-Path -Path $FileName1 -PathType Leaf) -and (Test-Path -Path $FileName2 -PathType Leaf)) {
		#the check depends on the version of PowerShell, because the cmdlet Get-FileHash is only available in version 4 and newer
		#TODO: Make old version as parameter in Call
		[int]$PSVersion = (Get-Host).Version.Major
		$HashFile1 = $null
		$HashFile2 = $null
		if ($PSVersion -ge 4) {
			#new version
			$HashFile1 = (Get-FileHash -Path $FileName1 -Algorithm MD5).Hash
			$HashFile2 = (Get-FileHash -Path $FileName2 -Algorithm MD5).Hash
		} else {
			#old version
			$md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
			$HashFile1 = [System.BitConverter]::ToString($md5.ComputeHash([System.IO.File]::ReadAllBytes($FileName1))).Replace("-","")
			$HashFile2 = [System.BitConverter]::ToString($md5.ComputeHash([System.IO.File]::ReadAllBytes($FileName2))).Replace("-","")
			Trace-LogMessage -Message "'$FileName1' has hash value '$HashFile1'" -Indent 5 -Level 9
			Trace-LogMessage -Message "'$FileName2' has hash value '$HashFile2'" -Indent 5 -Level 9
		}
		if (($HashFile1 -eq $HashFile2) -and ($null -ne $HashFile1)) {
			#the HashValues are equal
			Trace-LogMessage -Message "'$FileName1' and '$FileName1' have the same hash value" -Indent 5 -Level 10
			$global:RETURNVALUE_CompareFileHash = $True
			Trace-LogMessage -Message "Compare-FileHash returned $global:RETURNVALUE_CompareFileHash" -Indent 1 -Level 10
			return
		} else {
			#the HashValues are not equal
			Trace-LogMessage -Message "'$File1' and '$File2' have a different hash value" -Indent 5 -Level 10
			$global:RETURNVALUE_CompareFileHash = $False
			Trace-LogMessage -Message "Compare-FileHash returned $global:RETURNVALUE_CompareFileHash" -Indent 1 -Level 10
			return
		}
	} else {
		if (-not (Test-Path -Path $FileName1)) {
			Trace-LogMessage -Message "'$FileName1' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $FileName1 -PathType Leaf)) {
			Trace-LogMessage -Message "'$FileName1' is not a file." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $FileName2)) {
			Trace-LogMessage -Message "'$FileName1' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $FileName2 -PathType Leaf)) {
			Trace-LogMessage -Message "'$FileName2' is not a file." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
		#The files should not be the same, if the comparrison can not be performed
		$global:RETURNVALUE_CompareFileHash = $False
	}
	Trace-LogMessage -Message "Compare-FileHash returned $global:RETURNVALUE_CompareFileHash" -Indent 1 -Level 10
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	Usage of global Variable to save return value
	[boolean]$global:RETURNVALUE_CompareFiles = $False
.PARAMETER File1
	TODO
.PARAMETER File2
	TODO
.PARAMETER NoSize
	Do not compare the size of files. This might be required for comparing files on different SharePoint server versions.
.PARAMETER NoLastWriteTime
	Do not compare the LastWriteTime of files.
.PARAMETER NoHash
	Do not use HashValues to compare files (sufficient in most cases, but not always; much faster). Without this parameter HashValues will be compared too.
.EXAMPLE
	Compare-Files "D:\Downloads\BackUpTest\Source\File 1.txt" "D:\Downloads\BackUpTest\Source\File 2.txt"
	will return True, if both files exist, they have the same size, ListWriteTime and HashValue. In all other cases the return value will be False
#>
Function Compare-Files {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$File1,
		[Parameter(Mandatory=$true, Position=1)]
		[String]
		$File2,
		[Parameter(Mandatory=$false, Position=2)]
		[switch]
		$NoSize,
		[Parameter(Mandatory=$false, Position=3)]
		[switch]
		$NoLastWriteTime,
		[Parameter(Mandatory=$false, Position=4)]
		[switch]
		$NoHash
	)

	#remember the status of comparisons to avoid a cascade of if-statements, but to perform comparisons in a sequence and perform operation after comparisons
	[Boolean]$StatusOfComparison = $True

	if ($NoSize -and $NoLastWriteTime -and $NoHash) {
		Trace-LogMessage -Message "'$File1' and '$File2' are compared, but without size, LastWriteTime and HashValue the result will be useless!" -Indent 1 -Level 0 -MessageType Warning
	}

	Trace-LogMessage -Message "Compare '$File1' and '$File2' " -Indent 4 -Level 9
	#First Check: Do both files exist (and are files) - this check will always be performed
	if ((Test-Path $File1 -PathType Leaf) -AND (Test-Path -Path $File2 -PathType Leaf)) {
		#the status of all comparisons will remain true 
		Trace-LogMessage -Message "Test 1: '$File1' and '$File2' exist." -Indent 4 -Level 9
	} else {
		#the result of the comparison is false and further tests will not be performed
		$StatusOfComparison = $False
		#if one or both files does not exist (or ar not files) - they are not equal
		if (-NOT (Test-Path $File1)) {
			Trace-LogMessage -Message "First file '$File1' does not exist" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $File2)) {
			Trace-LogMessage -Message "Second file '$File2' does not exist" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $File1 -PathType Leaf)) {
			Trace-LogMessage -Message "First file '$File1' is not a file" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-NOT (Test-Path $File2 -PathType Leaf)) {
			Trace-LogMessage -Message "Second file '$File2' is not a file" -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
		[boolean]$global:RETURNVALUE_CompareFiles = $False
		Trace-LogMessage -Message "Compare-Files returned $global:RETURNVALUE_CompareFiles" -Indent 1 -Level 10
		return
	}

	#Second Check: is the filesize identical?
	if ($StatusOfComparison -and (-not $NoSize)) { #test will only be performed, if the test before was sucessful and the parameter NoSize is NOT set
		if ((Get-ChildItem -Path $File1).Length -eq (Get-ChildItem -Path $File2).Length) {
			#the status of all comparisons will remain true 
			Trace-LogMessage -Message "Test 2: '$File1' and '$File2' have the same size." -Indent 4 -Level 9
		} else {
			#the result of the comparison is false and further tests will not be performed
			$StatusOfComparison = $False
			Trace-LogMessage -Message "'$File1' and '$File2' have a different size - they are not equal" -Indent 4 -Level 8
			$global:RETURNVALUE_CompareFiles = $False
			Trace-LogMessage -Message "Compare-Files returned $global:RETURNVALUE_CompareFiles" -Indent 1 -Level 10
			return
		}
	} else {
		Trace-LogMessage -Message "The size of '$File1' and '$File2' was not compared" -Indent 4 -Level 8
	}

	#Third Check: LastWriteTime
	if ($StatusOfComparison -and (-not $NoLastWriteTime)) { #test will only be performed, if the test before was sucessful and the parameter NoLastWriteTime is NOT set
		if ((Get-ItemProperty -Path $File1 -Name LastWriteTime).LastWriteTime -eq (Get-ItemProperty -Path $File2 -Name LastWriteTime).LastWriteTime) {
			#the status of all comparisons will remain true 
			Trace-LogMessage -Message "Test 3: '$File1' and '$File2' have the same LastWriteTime." -Indent 4 -Level 9
		} else {
			#the result of the comparison is false and further tests will not be performed
			$StatusOfComparison = $False
			Trace-LogMessage -Message "'$File1' and '$File2' have a different LastWriteTime - they are not equal" -Indent 4 -Level 8
			$global:RETURNVALUE_CompareFiles = $False
			Trace-LogMessage -Message "Compare-Files returned $global:RETURNVALUE_CompareFiles" -Indent 1 -Level 10
			return
		}
	} else {
		Trace-LogMessage -Message "The LastWriteTime of '$File1' and '$File2' was not compared" -Indent 4 -Level 8
	}

	#Fourth Check: HashValue
	if ($StatusOfComparison -and (-not $NoHash)) { #test will only be performed, if the test before was sucessful and the parameter NoHash is NOT set
		Compare-FileHash -FileName1 $File1 -FileName2 $File2
		#use the result in the global variable to check if the hash values are equal
		if ($global:RETURNVALUE_CompareFileHash) {
			#the status of all comparisons will remain true 
			Trace-LogMessage -Message "Test 4: '$File1' and '$File2' have the same HashValue." -Indent 4 -Level 9
		} else {
			#the result of the comparison is false and further tests will not be performed
			$StatusOfComparison = $False
			Trace-LogMessage -Message "'$File1' and '$File2' have a different hash value - they are not equal (but they have the same size and LastWriteTime)!" -Indent 1 -Level 0 -MessageType Warning
			$global:RETURNVALUE_CompareFiles = $False
			Trace-LogMessage -Message "Compare-Files returned $global:RETURNVALUE_CompareFiles" -Indent 1 -Level 10
			return
		}
	} else {
		Trace-LogMessage -Message "The HashValue of '$File1' and '$File2' was not compared" -Indent 4 -Level 8
	}

	#finally they are equal
	if ($StatusOfComparison) { #test will only be performed, if all former tests were sucessful
		Trace-LogMessage -Message "'$File1' and '$File2' are equal" -Indent 4 -Level 8
		$global:RETURNVALUE_CompareFiles = $True
		Trace-LogMessage -Message "Compare-Files returned $global:RETURNVALUE_CompareFiles" -Indent 1 -Level 10
		return
	}
	#should never occur
	$global:RETURNVALUE_CompareFiles = $False
	Trace-LogMessage -Message "Compare-Files returned $global:RETURNVALUE_CompareFiles (should not been occured)" -Indent 1 -Level 1
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	TODO
.PARAMETER FileName
	absolute file path to file that has to be moved
.PARAMETER SourcePath
	necessary to find substructure of folders
.PARAMETER DestinationPath
	TODO
.PARAMETER ActionIfFileExists
	'KeepBoth', 'Overwrite', 'Delete', 'Error'
	keep both - renaming the new file with the current timestamp
	overwrite in destination
	delete in source and remain version in destination
	error message
.PARAMETER AddTimeStamp
	TODO
.EXAMPLE
	MoveFileToArchive -File "" -DestinationPath $DestinationPath -ArchivePath $ArchivePath -WithTimeStamp
#>
Function Move-File {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$FileName,
		[Parameter(Mandatory=$true, Position=1)]
		[String]
		$SourcePath,
		[Parameter(Mandatory=$true, Position=2)]
		[String]
		$DestinationPath,
		[Parameter(Mandatory=$true, Position=3)]
		[ValidateSet('KeepBoth', 'Overwrite', 'Delete', 'Error')]
		[String]$ActionIfFileExists,
		[Parameter(Mandatory=$false, Position=4)]
		[switch]
		$AddTimeStamp
	)
	#check, if all files and folders exist and are files or folders as expected
	if ((Test-Path -Path $FileName -PathType Leaf) -and (Test-Path -Path $SourcePath -PathType Container) -and (Test-Path -Path $DestinationPath -PathType Container)) {
		# $SourcePath and $DestinationPath must not end with '\' for renaming
		$SourcePath = $SourcePath.TrimEnd('\')
		$DestinationPath = $DestinationPath.TrimEnd('\')

		#Add the timestamp to the filename, if required
		if ($AddTimeStamp) {
			Trace-LogMessage -Message "Add timestamp to file '$FileName'" -Indent 5 -Level 7
			#rename the file before moving it
			$FileNameWithTimeStamp = Add-TimeStampToFileName -FileName $FileName
			#test if renaming was successful
			if (Test-Path -Path $FileNameWithTimeStamp -PathType Leaf) {
				$FileName = $FileNameWithTimeStamp 
				Trace-LogMessage -Message "new name: '$FileName'" -Indent 5 -Level 7
			} else {
				Trace-LogMessage -Message "Adding time stamp to '$FileName' was not successful" -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
			}
		} else {
			Trace-LogMessage -Message "File '$FileName' will not get a timestamp" -Indent 5 -Level 8
		}
		#The folder name in destination (it is the same structure)
		$PathToMoveFileTo = ((Get-ChildItem $FileName).Directory).FullName.Replace($SourcePath,$DestinationPath)
		#if the destination path does not exist, it will be created
		TestAndCreate-Path -FolderName $PathToMoveFileTo
		Trace-LogMessage -Message "Ensured, that folder '$PathToMoveFileTo' exists." -Indent 5 -Level 8
		Trace-LogMessage -Message "Move file '$FileName' to '$PathToMoveFileTo'" -Indent 5 -Level 7

		#the file should not exist (with the new name) in destination
		$NewFileNameInDestination = $FileName.Replace($SourcePath,$DestinationPath)
		if (Test-Path $NewFileNameInDestination -PathType Leaf) {
			if ($ActionIfFileExists -eq "KeepBoth") {
				#rename the file to be moved using the current timestamp
				Trace-LogMessage -Message "Add additional timestamp to file '$FileName'" -Indent 5 -Level 5
				$FileNameWithTimeStamp = Add-TimeStampToFileName -FileName $FileName -UseCurrentTime -OverwriteTimestamp
				#test if renaming was successful
				if (Test-Path -Path $FileNameWithTimeStamp -PathType Leaf) {
					$FileName = $FileNameWithTimeStamp 
					Trace-LogMessage -Message "new name: '$FileName'" -Indent 5 -Level 7
				} else {
					Trace-LogMessage -Message "Adding time stamp to '$FileName' was not successful" -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				}
				#check again, if the file already exists - in that case we have a problem
				$NewFileNameInDestination2 = $FileName.Replace($SourcePath,$DestinationPath)
				#using the current time stamp - it should not happen that the file exists
				if (Test-Path $NewFileNameInDestination2 -PathType Leaf) {
					Trace-LogMessage -Message "'$FileName' has current timestamps and exists already. It will not be moved." -Indent 0 -Level 0 -MessageType Error
					$global:VARIABLE_ErrorCounter++
				} else {
					Move-Item -Path $FileName -Destination $PathToMoveFileTo -Force
					Trace-LogMessage -Message "The file '$FileName' was moved to '$PathToMoveFileTo' - using the current time for the timestamp" -Indent 5 -Level 3 -MessageType Warning
				}
			} elseif ($ActionIfFileExists -eq "Overwrite") {
				#move the file to the destination
				Move-Item -Path $FileName -Destination $PathToMoveFileTo -Force
				Trace-LogMessage -Message "The file '$FileName' was moved to '$PathToMoveFileTo' - the existing file was overwritten" -Indent 3 -Level 5
			} elseif ($ActionIfFileExists -eq "Delete") {
				#delete the file
				Remove-Item -Path $FileName -Force
				Trace-LogMessage -Message "The file '$FileName' was deleted" -Indent 3 -Level 4
			} elseif ($ActionIfFileExists -eq "Error") {
				#TODO: ERROR HANDLING necessary (throw + trap)
				Trace-LogMessage -Message "ERROR: The file '$NewFileNameInDestination' does already exist." -Indent 0 -Level 0 -MessageType Error
				$global:VARIABLE_ErrorCounter++
			}
		} else {
			#move the file to the destination
			Move-Item -Path $FileName -Destination $PathToMoveFileTo
			Trace-LogMessage -Message "The file '$FileName' was moved to '$PathToMoveFileTo'" -Indent 3 -Level 5
		}
	} else {
		if (-not (Test-Path -Path $FileName)) {
			Trace-LogMessage -Message "'$FileName' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $FileName -PathType Leaf)) {
			Trace-LogMessage -Message "'$FileName' is not a file." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $SourcePath)) {
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
	TODO
.DESCRIPTION
	TODO
.PARAMETER FolderName
	TODO
.PARAMETER DestinationPath
	TODO
.PARAMETER FileType
	all filetypes are the default value (*.*)
.PARAMETER ActionIfFileExists
	'KeepBoth', 'Overwrite', 'Delete', 'Error'
	keep both - renaming the new file with the current timestamp
	overwrite in destination
	delete in source - and remain file in destination
	error message
.EXAMPLE
	Move-AllFiles -FolderName "D:\OneNote" -FileType "*.onepkg" -DestinationPath "D:\Archive OneNote" -ActionIfFileExists Overwrite
#>
Function Move-AllFiles {
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		[String]
		$FolderName,
		[Parameter(Mandatory=$true, Position=1)]
		[String]
		$DestinationPath,
		[Parameter(Mandatory=$false, Position=2)]
		[String]
		$FileType = "*.*",
		[Parameter(Mandatory=$true, Position=3)]
		[ValidateSet('KeepBoth', 'Overwrite', 'Delete', 'Error')]
		[String]$ActionIfFileExists
	)
	#check, if all folders exist and are folders
	if ((Test-Path -Path $FolderName -PathType Container) -and (Test-Path -Path $DestinationPath -PathType Container)) {
		#Find all relevant files in source
		$FileList = Get-ChildItem -Path $FolderName -Filter $FileType -File
		Trace-LogMessage -Message "Move $($FileList.Count) file(s) '$FileType' from '$FolderName' to '$DestinationPath' (action: $ActionIfFileExists)" -Indent 3 -Level 1
		foreach ($File in $FileList) {
			#Output is included in the fuction Move-File
			Move-File -FileName $File.FullName -SourcePath $File.Directory -DestinationPath $DestinationPath -ActionIfFileExists $ActionIfFileExists
		}
	} else {
		if (-not (Test-Path -Path $FolderName)) {
			Trace-LogMessage -Message "'$FolderName' does not exist." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $FolderName -PathType Container)) {
			Trace-LogMessage -Message "'$FolderName' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		} elseif (-not (Test-Path -Path $DestinationPath)) {
			#if the destination does not exisit, this can be mitigated, by creating hte folder.
			Trace-LogMessage -Message "'$DestinationPath' does not exist." -Indent 0 -Level 0 -MessageType Warning
			TestAndCreate-Path -FolderName $DestinationPath
			Trace-LogMessage -Message "'$DestinationPath' was created." -Indent 0 -Level 0 -MessageType Info
			#TODO: Transfer this error handling to whereever applicable
			#recursive recall to ensure the files are moved in a second try
			Move-AllFiles -FolderName $FolderName -DestinationPath $DestinationPath -FileType $FileType -ActionIfFileExists $ActionIfFileExists
		} elseif (-not (Test-Path -Path $DestinationPath -PathType Container)) {
			Trace-LogMessage -Message "'$DestinationPath' is not a folder." -Indent 0 -Level 0 -MessageType Error
			$global:VARIABLE_ErrorCounter++
		}
	}
}