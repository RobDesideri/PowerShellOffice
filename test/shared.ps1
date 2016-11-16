# Some shared functions for testing purpose

Function Create-FileList([string[]]$names)
{
	$names | ForEach {
	    [PSCustomObject]@{FullName = "d:\foo\bar\$_"; Name = $_; }
	}
} 

Function Create-ProcessList([string[]]$names)
{
	$names | ForEach {
	    [PSCustomObject]@{ProcessName = $_}
	}
}

function Create-FakeProcess ([string]$prName)
{
    $p = New-Object -TypeName System.Diagnostics.Process
	$p.ProcessName = $prName
	$p.Id = Get-RandomId
}

Function Get-RandomId
{
	$r = New-Object -TypeName System.Random
	[int]$i = 100
	Return $r.Next($i)
}