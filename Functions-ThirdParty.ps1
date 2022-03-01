<#
.SYNOPSIS
    Gets the hash value of a file or string
    source: http://dbadailystuff.com/2013/03/11/get-hash-a-powershell-hash-function
.DESCRIPTION
	Gets the hash value of a file or string
	It uses System.Security.Cryptography.HashAlgorithm (http://msdn.microsoft.com/en-us/library/system.security.cryptography.hashalgorithm.aspx)
	and FileStream Class (http://msdn.microsoft.com/en-us/library/system.io.filestream.aspx)
	Based on: http://blog.brianhartsock.com/2008/12/13/using-powershell-for-md5-checksums/ and some ideas on Microsoft Online Help
	Be aware, to avoid confusions, that if you use the pipeline, the behaviour is the same as using -Text, not -File
.PARAMETER File
	File to get the hash from.
.PARAMETER Text
	Text string to get the hash from
.PARAMETER Algorithm
	Type of hash algorithm to use. Default is SHA1
.EXAMPLE
	C:\PS> Get-Hash "hello_world.txt"
	Gets the SHA1 from myFile.txt file. When there's no explicit parameter, it uses -File
.EXAMPLE
	Get-Hash -File "C:\temp\hello_world.txt"
	Gets the SHA1 from myFile.txt file
.EXAMPLE
	C:\PS> Get-Hash -Algorithm "MD5" -Text "Hello Wold!"
	Gets the MD5 from a string
.EXAMPLE
	C:\PS> "Hello Wold!" | Get-Hash
	We can pass a string throught the pipeline
.EXAMPLE
	Get-Content "c:\temp\hello_world.txt" | Get-Hash
	It gets the string from Get-Content
.EXAMPLE
	Get-ChildItem "C:\temp\*.txt" | %{ Write-Output "File: $($_)   has this hash: $(Get-Hash $_)" }
	This is a more complex example gets the hash of all "*.tmp" files
.NOTES
	DBA daily stuff (http://dbadailystuff.com) by Josep Martínez Vilà
	Licensed under a Creative Commons Attribution 3.0 Unported License
.LINK
	Original post: https://dbadailystuff.com/2013/03/11/get-hash-a-powershell-hash-function/
#>
function Get-Hash
{
	Param
	(
		[parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="set1")]
		[String]
		$Text,
		[parameter(Position=0, Mandatory=$true, 
		ValueFromPipeline=$false, ParameterSetName="set2")]
		[String]
		$File = "",
		[parameter(Mandatory=$false, ValueFromPipeline=$false)]
		[ValidateSet("MD5", "SHA", "SHA1", "SHA-256", "SHA-384", "SHA-512")]
		[String]
		$Algorithm = "SHA1"
	)
	Begin
	{
		$hashAlgorithm = [System.Security.Cryptography.HashAlgorithm]::Create($Algorithm)
	}
	Process
	{
		$md5StringBuilder = New-Object System.Text.StringBuilder 50
		$ue = New-Object System.Text.UTF8Encoding

		if ($File){
			try {
				if (!(Test-Path -literalpath $File)){
					throw "Test-Path returned false."
				}
			}
			catch {
				throw "Get-Hash - File not found or without permisions: [$File]. $_"
			}
			try {
				[System.IO.FileStream]$FileStream = [System.IO.File]::Open($File, [System.IO.FileMode]::Open);
				$hashAlgorithm.ComputeHash($FileStream) | 
					% { [void] $md5StringBuilder.Append($_.ToString("x2")) }
			}
			catch {
				throw "Get-Hash - Error reading or hashing the file: [$File]"
			}
			finally {
				$FileStream.Close()
				$FileStream.Dispose()
			}
		}
		else {
			$hashAlgorithm.ComputeHash($ue.GetBytes($Text)) | 
				% { [void] $md5StringBuilder.Append($_.ToString("x2")) }
		}

		return $md5StringBuilder.ToString()
	}
}