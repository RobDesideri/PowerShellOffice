#	Scan all docx files in specified directory and get the Content Status metadata
#	Alert: for performance purpose, this script kill all word process before and after execution!
#	Version: 1.0
#	Require .NET Framework v3.0
#	Require PS v4.0
[System.Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Word")
[String]$inPath = $inputPath
[String]$outPath = $outFilePath

Function Get-DocxMetadata {
	Param(
	[Parameter(Mandatory=$True, Position=1)]
	[string]$inputPath,
	[Parameter(Mandatory=$True, Position=2)]
	[string]$outFilePath,
	[Parameter(Mandatory=$False, Position=3)]
	[string]$outFormat="csv" #for future develop....
)
	Write-Output "Init..."
	Clear-WordProc
	Write-Output "Word precesses killed"
	Write-Output "Find all docx artifacts..."
	$docxs = Get-All-Docx -path $inPath
	Write-Output "Artifacts found"
	Write-Output "Extract all metadata..."
	If($outFormat -eq "csv"){
		$meta = Get-Metadata-As-Csv -filePath $docxs
	}
	Write-Output "Metadata extracted"
	Write-Output "Write out file in '$outPath'"
	Set-Output -outPath $outPath -outVal $meta
	Write-Output "File ready"
	Write-Output "Kill Word processes..."
	Clear-WordProc
	Write-Output "Word precesses killed"
	Write-Output "End"
}