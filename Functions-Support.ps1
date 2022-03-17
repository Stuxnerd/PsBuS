<#
.SYNOPSIS
	The script includes several useful functions which are used by several scripts.
.DESCRIPTION
	This script includes these functions:
	* Convert-IntTo2DigitString
	* Get-TimeStamp
	* Add-TimeStampToFileName
	* TestAndCreate-Path
	* Get-NormalizedPath
	Assumptions:
	 1. Get-TimeStamp: The script needs more than one second to run, so a file with the same timestamp (exact same second) will not exist already!
.LINK
	https://github.com/Stuxnerd/PsBuS
.NOTES
	VERSION: 0.9.2 - 2022-03-16

	AUTHOR: Stuxnerd
		If you want to support me: bitcoin:19sbTycBKvRdyHhEyJy5QbGn6Ua68mWVwC

	LICENSE: 	This script is licensed under GNU General Public License version 3.0 (GPLv3).
		Find more information at http://www.gnu.org/licenses/gpl.html

	TODO: These tasks have to be implemented in the following versions:
	till version 1.0 - additional features, testing and documentation
	* Wenn Ziel ein Ordner und Quelle eine Datei ist, prüfen, ob der Ordner leer ist – wenn ja löschen sonst Fehler / Farbe: Magenta, um Fehler aufzuzeigen (sollte Funktion Find-FileFolders obsolet machen)
	* default way to find parent folder
	* add "Ignore" as action for ActionIfFileExists and separate it to an enummeration (prevent double code in script)
	* rename variable TestFile
	* common names for parameters like FileName1 and File1
	* check, if Test-Path is used with -Type whereever useful
	* check if a file is locked (e. g. pst archives, if outlook is still running): Ask, if wait and try again or ignore (Warning)
	* test the counters
	* show progress of long duration tasks with progress bar
	* get rid of $global:VARIABLE in this file
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

#not used in this script


##################################
#GLOBAL VARIABLES - RETURN VALUES#
##################################
#the global variables are used to save the return values of functions; they are just for usage of the funtions

#not used in this script


####################################
#GLOBAL VARIABLES - VARIABLE VALUES#
####################################
#these values are used during execution, but are independent from a single fuction invocation

#not used in this script


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
	TODO
.DESCRIPTION
	The functions returs a String of two digit number
.PARAMETER Number
	The imput shall be a number (like a date or time), which has one or two digits
#>
function Convert-IntTo2DigitString {
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[int]
		$Number
	)
	#distinguish, if the number needs a leading 0
	if ($Number -lt 10) {
		#Add a 0, if the number has only one digit
		return ("0" + $Number)
	}
	#transform to a String
	return ([String]$Number)
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	Assumption, that the script needs more than one second to run, so a file with the same timestamp (exact same second) will not exist already! Only if no Filename was used.
	The format " (yyyy-MM-dd--HH-mm-ss)" correspondes to the regular expression '\s[(][0-9]{4}[-]([0-9]{2}[-]){2}[-]([0-9]{2}[-]){2}[0-9]{2}[)]'.
	That is used in the function Add-TimeStampToFileName to reset the timestamp.
.PARAMETER FileName
	Optional, get the timestamp the file was written the last time
.TODO
	Write result to a global variable
#>
Function Get-TimeStamp {
	Param (
		[Parameter(Mandatory = $false, Position = 0)]
		[String]
		$FileName = ""
	)
	#current timestamp, if no file as reference (or file does not exists)
	if ($FileName -eq "" -or -not (Test-Path -Path $FileName -PathType Leaf)) {
		return '(' + (Get-Date -f yyyy-MM-dd--HH-mm-ss) + ')'
	}
 else {
		#timestamp of last time file was written
		$LastWriteTime = [datetime](Get-ItemProperty -Path $FileName -Name LastWriteTime).lastwritetime
		#for a better overview all values have two digits (but the year)
		[String]$LastWriteTimeString = "($($LastWriteTime.Year)-$(Convert-IntTo2DigitString $LastWriteTime.Month)-$(Convert-IntTo2DigitString $LastWriteTime.Day)--" +
		"$(Convert-IntTo2DigitString $LastWriteTime.Hour)-$(Convert-IntTo2DigitString $LastWriteTime.Minute)-$(Convert-IntTo2DigitString $LastWriteTime.Second))"
		return $LastWriteTimeString		
	}
}

<#
.SYNOPSIS
	TODO
.DESCRIPTION
	Assumption: A File Extension starts with a '.' and is part of every file
	If renaming is not possible the return value will be old file name
	The format of the timestamp " (yyyy-MM-dd--HH-mm-ss)" correspondes to the regular expression '\s[(][0-9]{4}[-]([0-9]{2}[-]){2}[-]([0-9]{2}[-]){2}[0-9]{2}[)]'.
.PARAMETER FileName 
	TODO
.PARAMETER UseCurrentTime 
	TODO
.PARAMETER OverwriteTimestamp 
	TODO
.EXAMPLE
	Add-TimeStampToFileName "D:\File 3.txt"
	will rename the file to "D:\File 3 (2015-12-12--22-50-49).txt" and return the new file name
#>
Function Add-TimeStampToFileName {
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[String]
		$FileName,
		[Parameter(Mandatory = $false, Position = 1)]
		[switch]
		$UseCurrentTime,
		[Parameter(Mandatory = $false, Position = 2)]
		[switch]
		$OverwriteTimestamp
	)
	#only possible, if the file exists and is a file (no folder)
	if (Test-Path -Path $FileName -PathType Leaf) {
		#distinguish, if the existing filestamp should be substituted
		#to prevent multiple renaming a temporary filename after renaming the file is used (only inside the script - the file will not be renamed in the filesystem)
		$TemporaryFileName = $FileName
		#if the timestamp has to be overwritten, the existing timestamp will be cut out and later added as if not timestamp would exists
		if ($OverwriteTimestamp) {
			#check if the file has a timestamp
			$TimeStampStart = $FileName.LastIndexOf(' ')
			$TimeStampEnd = $FileName.LastIndexOf(')')
			#cut the timestamp of the file
			#if the filename does not contains a timestamp an exception will be thrown
			try {
				$Timestamp = $FileName.Substring($TimeStampStart, $TimeStampEnd - $TimeStampStart + 1)
				#check if the timestamp matches the regular expression (e. g. " (2015-12-13--20-52-14)")
				if ($Timestamp -match '\s[(][0-9]{4}[-]([0-9]{2}[-]){2}[-]([0-9]{2}[-]){2}[0-9]{2}[)]') {
					$TemporaryFileName = $FileName.Substring(0, $TimeStampStart) + $FileName.Substring($TimeStampEnd + 1, $FileName.Length - $TimeStampEnd - 1)
				}
				#otherwise: to prevent further errors, the filename will remain and the new timestamp will be added (has been done above)
			}
			catch {
				#do nothing if the filename does not match the regular expression for a timestamp
			}
		}
		$PositionOfFileExtension = $TemporaryFileName.LastIndexOf('.')
		#distinguish, if the filestamp from the file or the current filestamp should be used
		if ($UseCurrentTime) {
			$NewFileName = $TemporaryFileName.Substring(0, $PositionOfFileExtension) + " " + (Get-TimeStamp) + $TemporaryFileName.Substring($PositionOfFileExtension)
		}
		else {
			$NewFileName = $TemporaryFileName.Substring(0, $PositionOfFileExtension) + " " + (Get-TimeStamp -FileName $FileName) + $TemporaryFileName.Substring($PositionOfFileExtension)
		}
		#ensure, that the new filename does not exist yet
		if (Test-Path -Path $NewFileName -PathType Leaf) {
			#if filename exists already: Renaming not possible (return old name)
			#overwriting the return value to a global variable with the old file name
			return $FileName
		}
		else {
			Rename-Item -Path $FileName -NewName $NewFileName -Force
			#return new file name
			return $NewFileName
		}
	}
 else {
		#Not able to add time stamp to FileName, because file does not exist (or is a folder)
		#return old name
		return $FileName
	}
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	several subdirectories will be vreated if necessary
.PARAMETER FolderName
	FullName of the folder to be tested and created if necessary
#>
Function TestAndCreate-Path {
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[String]
		$FolderName
	)
	#Check, if the folder exists, if not it will be created
	if (-not (Test-Path -Path $FolderName)) {
		New-Item $FolderName -ItemType directory | Out-Null
		Trace-LogMessage -Message "$FolderName was created" -MessageType Info -Level 8
	} else {
		Trace-LogMessage -Message "$FolderName did exist" -MessageType Info -Level 8
	}
}


<#
.SYNOPSIS
	Normalized the format of a path string to end with \ or not
.DESCRIPTION
	Ensure, that a path ends with a "\" (or even not) - independent of the users input
	Per default "\" is added.
.PARAMETER FolderName
	FullName of the folder (path)
.PARAMETER NoBackslash
	Switch, if the "\" shall be deleted instead of added.	
#>
Function Get-NormalizedPath {
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[String]
		$FolderName,
		[Parameter(Mandatory = $false, Position = 1)]
		[Switch]
		$NoBackslash
	)
	Process {
		if ($NoBackslash) {
			return $FolderName.TrimEnd("\")
		}
		elseif ($FolderName.EndsWith("\")) {
			return $FolderName
		}
		else {
			return $FolderName + "\"
		}
	}
}