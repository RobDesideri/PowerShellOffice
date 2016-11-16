# Import testing helpers dependencies
& "$PSScriptRoot\InstallTestingDepends.ps1"
Import-Module "$PSScriptRoot\lib\PowerShellTestingHelpers"
#Get all public functions
$Public = @( Get-ChildItem -Path $PSScriptRoot\..\src\script\public\*.ps1 -ErrorAction SilentlyContinue )
Foreach($import in @($Public))
{
    Try
    {
        . $import.fullname
    }
    Catch
    {
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

Describe "PowerShellOfficePublicFunctions" {
	Context "Exists" {
		It "Runs" {

		}
	}
}