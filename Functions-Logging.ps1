<#
.SYNOPSIS
	The script includes several useful functions for logging.
.DESCRIPTION
	This script includes these functions:
	* Trace-LogMessage
	* TraceLogDefaultData
	* Reset-LogMessages
	These settings are defined within this script and can be adapted when using it:
	* $global:CONSTANT_ActivateDetailedLogging
	* $global:CONSTANT_LogFilePath
	* $global:CONSTANT_LogFilePathALL
	* $global:CONSTANT_LogFilePathError
	* $global:CONSTANT_LogFileArchivePath
	* $global:CONSTANT_LogLevel
.LINK
	https://github.com/Stuxnerd/PsBuS
.NOTES
	VERSION: 0.9.1 - 2022-03-16

	AUTHOR: @Stuxnerd
		If you want to support me: bitcoin:19sbTycBKvRdyHhEyJy5QbGn6Ua68mWVwC

	LICENSE: This script is licensed under GNU General Public License version 3.0 (GPLv3).
		Find more information at http://www.gnu.org/licenses/gpl.html

	TODO: These tasks have to be implemented in the following versions:
	till version 1.0 - additional features, testing and documentation
	* transform logging format to CSV/XML (including timestamp, unique ID per message type)
	* document features and functions in SourceForge Wiki and in this scipt
	till *version 1.1 - optional features
	* enhance exception management with traps, throw, $ErrorActionPreference, etc. (e.g. if the file name length is too big)
#>


###########################################
#INTEGRATION OF EXTERNAL FUNCTION PACKAGES#
###########################################

. ../PsBuS/Functions-Support.ps1


##################################
#GLOBAL VARIABLES - RETURN VALUES#
##################################
#the global variables are used to save the return values of functions; they are just for usage of the funtions

#not used in this script


####################################
#GLOBAL VARIABLES - VARIABLE VALUES#
####################################
#these values are used during execution, but are independent from a single fuction invocation

#Counter for LogMessage - the counter is the same for all file to correlate messages (e.g. to find prior messages from the ID in the ErrorList)
#the counter will be incremented indepent wheather the message is shown or not
[int]$global:VARIABLE_LogCounter = 0


#############################################
#VARIABLES FOR THE SETTING - CONSTANT VALUES#
#############################################
#these values define the configuration of the script; the might be overwritten by an external script which is using included functions

#The creation of a detailed logging can be deactivated by setting this value to $False. This may improve the speed and reduce the required storage space)
[bool]$global:CONSTANT_ActivateDetailedLogging = $true

#path for the log file, that includes selected messages
[String]$global:CONSTANT_LogFilePath = ".\BackUpLog.txt"
#path for the log file, that includes all messages
[String]$global:CONSTANT_LogFilePathALL = ".\BackUpLogAll.txt"
#path for the log file, only for error messages
[String]$global:CONSTANT_LogFilePathError = ".\BackUpLogError.txt"
#subfolder for the log file archive
[String]$global:CONSTANT_LogFileArchivePath = ".\BackUpLog\"

#The script itself has also a certain log level. A message appears only in the log, the the number of the level for the message is lower or equal to the level of the script.
#Per default a message has least importance.
#(0 - no logging, 1 - necessary logs only, 3 - most usefull, ..., 10 - every message logged). The default value is 3.
[int]$global:CONSTANT_LogLevel = 3


#####################
#FUNCTION DEFINITION#
#####################

<#
.SYNOPSIS
	This function offers a simple way to log messages in a script.
.DESCRIPTION
	Optimize Logging. This script file should not contain any "Write-Host" statement. This function is controlled by several variables, that define its functionality.
.PARAMETER Message
	The message, that should be logged.
.PARAMETER Indent
	The indent before the message (spaces)
.PARAMETER Level
	The Level of the message. The scriptitself has also a certain level (0 - no logging, 1 - necessary logs only, ..., 9 - every message logged). The message appears only in the log, the the number of the level for the message is lower or equal to the level of the script. Per default a message has least importance.
.PARAMETER MessageType
	Indicates special types of messages. They will be presented in another color
	valid values are: 'Exception', 'Error', 'Warning', 'Confirmation', 'Info'
	The default value is 'Info' (white)
.PARAMETER NoTimeStamp
	Will deactivate the TimeStamp for this entry
.NOTES
	The global variable $global:CONSTANT_ActivateDetailedLogging will impact, which log files will be written.
.TODO
	* Parameter ExportPath for writing into an additional log file (really required?)
	* XML Logging
	* CSV Logging
#>
function Trace-LogMessage {
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[String]
		$Message,
		[Parameter(Mandatory = $false, Position = 1)]
		[int]
		$Indent = 0,
		[Parameter(Mandatory = $false, Position = 2)]
		[int]
		$Level = 10,
		[Parameter(Mandatory = $false, Position = 3)]
		[ValidateSet('Exception', 'Error', 'Warning', 'Confirmation', 'Info')]
		$MessageType = 'Info',
		[Parameter(Mandatory = $false, Position = 4)]
		[switch]
		$NoTimeStamp = $false
	)
	#the logging will only be performed if the detailed logging is activated or if the level of the log entry is important enough
	if ($global:CONSTANT_ActivateDetailedLogging -or ($global:CONSTANT_LogLevel -ge $Level)) {
		#increment the counter
		$global:VARIABLE_LogCounter++

		#Convert the indent int value to a String of spaces
		$IndentSpaces = ""
		while ($Indent-- -gt 0) {
			$IndentSpaces = $IndentSpaces + " "
		}

		#Build the final message with all additional information and spaces
		[String]$LogMessage = ""

		#add the current counter number - assumption (always under 10000)
		$counter = $global:VARIABLE_LogCounter
		switch ($counter) {
			{ $_ -lt 10 } { $LogMessage += "$counter     "; break }
			{ $_ -lt 100 } { $LogMessage += "$counter    "; break }
			{ $_ -lt 1000 } { $LogMessage += "$counter   "; break }
			{ $_ -lt 10000 } { $LogMessage += "$counter  "; break }
			{ $_ -lt 100000 } { $LogMessage += "$counter "; break }
			Default { $LogMessage += "$counter "; break }
		}

		#log the date and time - if not suppressed explicitly
		if (-not $NoTimeStamp) {
			$LogMessage = $LogMessage + "$(Get-Date -f yyyy-MM-dd--HH-mm-ss) "
		}

		#an exception or error message has an additional text EXCEPTION/ERROR (also for WARNING)
		if ($MessageType -eq 'Exception') {
			$LogMessage = $LogMessage + "EXCEPTION: "
		}
		elseif ($MessageType -eq 'Error') {
			$LogMessage = $LogMessage + "ERROR: "
		}
		elseif ($MessageType -eq 'Warning') {
			$LogMessage = $LogMessage + "WARNING: "
		}

		#concatenate the final log message
		$LogMessage = $LogMessage + $IndentSpaces + $Message

		#the output to console is only for error messages and messages with the required level
		if ($global:CONSTANT_LogLevel -ge $Level -or $MessageType -eq 'Error') {
			#Output to Console - Exception/Errors messages in red, Warnings in yellow, Confirmations in green
			if (($MessageType -eq 'Error') -or ($MessageType -eq 'Exception')) {
				Write-Host -Object $LogMessage -ForegroundColor Red
			}
			elseif ($MessageType -eq 'Warning') {
				Write-Host -Object $LogMessage -ForegroundColor Yellow
			}
			elseif ($MessageType -eq 'Confirmation') {
				Write-Host -Object $LogMessage -ForegroundColor Green
			}
			else {
				Write-Host -Object $LogMessage #(white is default)
			}
		}
		else {
			#no action for irrelevant messages
		}
		#Output to logfiles - depending on the level or the message type
		#selected messages
		if ($global:CONSTANT_LogLevel -ge $Level) {
			Out-File -FilePath $global:CONSTANT_LogFilePath -InputObject $LogMessage -Append
		}
		#special log file for error messages
		if ($MessageType -eq 'Error') {
			Out-File -FilePath $global:CONSTANT_LogFilePathError -InputObject $LogMessage -Append
		}
		#all messages (if not deactivated)
		if ($global:CONSTANT_ActivateDetailedLogging) {
			Out-File -FilePath $global:CONSTANT_LogFilePathAll -InputObject $LogMessage -Append
		}
	}
 else {
		#nothing will happen, if the logging is deactivated or the log level is too high
	}
}


<#
.SYNOPSIS
	Write general data into log file
.DESCRIPTION
	This will write some general data into the log file.
#>
function Trace-LogDefaultData {
	Trace-LogMessage -Message "New entry at $(Get-Date -f dd.MM.yyyy--HH-mm-ss)" -Level 1 -MessageType Info -NoTimeStamp
	Trace-LogMessage -Message "Client:     $env:COMPUTERNAME" -Level 1 -MessageType Info -NoTimeStamp
	Trace-LogMessage -Message "User:       $env:USERNAME" -Level 1 -MessageType Info -NoTimeStamp
	Trace-LogMessage -Message "PSVersion:  $((Get-Host).Version.Major)" -Level 1 -MessageType Info -NoTimeStamp
}


<#
.SYNOPSIS
	TODO
.DESCRIPTION
	This function adds a timestamp to the log file and moves it to a subfolder with all old log files
.PARAMETER BackUpPath
	TODO
.NOTES
	Some of the used variables are defines with the function Trace-LogMessage ()
	They might be overwritten and changed, if necessary
	- $global:CONSTANT_LogFilePath
	- $global:CONSTANT_LogFilePathALL
	- $global:CONSTANT_LogFilePathError
	- $global:CONSTANT_LogFileArchivePath
#>
function Reset-LogMessages {
	Param (
		[Parameter(Mandatory = $false, Position = 0)]
		[String]
		$BackUpPath = ""
	)

	[boolean]$LogFileExists = $false
	[boolean]$LogFileAllExists = $false
	[boolean]$LogFileErrorExists = $false
	$NewNameForLogFile = ""
	$NewNameForLogFileAll = ""
	$NewNameForLogFileError = ""

	#rename current logfiles - if they exist
	if (Test-Path -Path $global:CONSTANT_LogFilePath) {
		$NewNameForLogFile = Add-TimeStampToFileName -FileName $global:CONSTANT_LogFilePath
		$LogFileExists = $true
	}
	if (Test-Path -Path $global:CONSTANT_LogFilePathAll) {
		$NewNameForLogFileAll = Add-TimeStampToFileName -FileName $global:CONSTANT_LogFilePathAll
		$LogFileAllExists = $true
	}
	if (Test-Path -Path $global:CONSTANT_LogFilePathError) {
		$NewNameForLogFileError = Add-TimeStampToFileName -FileName $global:CONSTANT_LogFilePathError
		$LogFileErrorExists = $true
	}

	#this will create a new log file
	$message = "Old LogFiles were renamed to: "
	if ($LogFileExists) {
		$message = $message + "'$NewNameForLogFile'; "
	}
	if ($LogFileAllExists) {
		$message = $message + "'$NewNameForLogFileAll'; "
	}
	if ($LogFileErrorExists) {
		$message = $message + "'$NewNameForLogFileError'"
	}

	if ($LogFileExists -or $NewNameForLogFileAll -or $NewNameForLogFileError) {
		Trace-LogMessage -Message "$message" -MessageType Confirmation
	}
 else {
		Trace-LogMessage -Message "No old LogFiles were renamed." -MessageType Confirmation
	}

	#test, which subfolder has to be used - if nothing was defined (or the definition is wrong), the default path wil be used
	if ($BackUpPath -eq "" -or (-not(Test-Path -Path $BackUpPath -PathType Container))) {
		#show a warning
		if (-not ($BackUpPath -eq "")) {
			if (-not(Test-Path -Path $BackUpPath -PathType Container)) {
				Trace-LogMessage -Message "'$BackUpPath' is not a valid pathname. The default path '$global:CONSTANT_LogFileArchivePath' will be used" -Indent 1 -Level 0 -MessageType Warning
			}
		}
		$BackUpPath = $global:CONSTANT_LogFileArchivePath

		#if the subfolder does not exist, it has to be created
		TestAndCreate-Path -FolderName $BackUpPath
	}

	#ensure the destination path exists
	if (-not (Test-Path -Path $BackUpPath -PathType Container)) {
		Trace-LogMessage -Message "'$BackUpPath' was not created" -Indent 0 -Level 0 -MessageType Warning
	}
	#move the old log files to the archive - if they exist
	if ($LogFileExists -and (Test-Path -Path $NewNameForLogFile)) {
		#thanks to lazy evaluation the Test-Path will only be checked, if the file existed before
		Move-Item -Path $NewNameForLogFile -Destination $BackUpPath
		Trace-LogMessage -Message "'$NewNameForLogFile' was moved to '$BackUpPath'"
	}
	if ($NewNameForLogFileAll -and (Test-Path -Path $NewNameForLogFileAll)) {
		Move-Item -Path $NewNameForLogFileAll -Destination $BackUpPath
		Trace-LogMessage -Message "'$NewNameForLogFile' was moved to '$BackUpPath'"
	}
	if ($NewNameForLogFileError -and (Test-Path -Path $NewNameForLogFileError)) {
		Move-Item -Path $NewNameForLogFileError -Destination $BackUpPath
		Trace-LogMessage -Message "'$NewNameForLogFileError' was moved to '$BackUpPath'"
	}
}