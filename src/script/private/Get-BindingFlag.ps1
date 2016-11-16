Function Get-BindingFlag{
	$binding = [System.Reflection.BindingFlags]::GetProperty
	Return $binding 
}