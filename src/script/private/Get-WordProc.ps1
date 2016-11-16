Function Get-Proc{
	param(
		[Parameter(Mandatory=$True, Position=1)]
		[string]$procName,
		[switch]$All
	)
	$p = Get-Process | Where {$_.ProcessName -like "*WINWORD*"}

	if(-not($p -eq $null))
	{
  		if($All)
		{
			$pr = Select -InputObject $p -First
		} Else {
			$pr = $p
		}
	} Else {
		$pr = $null
	}
	
	Return $pr
}