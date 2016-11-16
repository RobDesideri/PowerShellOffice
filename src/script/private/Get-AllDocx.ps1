Function Get-AllDocx{
	param(
		[Parameter(Mandatory=$True, Position=1)]
		[String]$path
	)
	$p = Get-ChildItem -Path $path -Recurse -Name | Where Extension -eq "docx" | Select -ExpandProperty "FullName"
	Return $p
}