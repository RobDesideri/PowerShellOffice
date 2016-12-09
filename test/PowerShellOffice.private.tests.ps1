# Import testing resources
& "$PSScriptRoot\SharedSetup.ps1"
$fp = "$PSScriptRoot\..\src\script\private\*.ps1"
$fpCollection = @( Get-ChildItem -Path $fp -ErrorAction SilentlyContinue )
Foreach($import in @($fpCollection))
{
	Try
	{
		. $import.FullName
	}
	Catch
	{
		Write-Error -Message "Failed to import function $($import.fullname): $_"
	}
}

#Test begin!

Describe "Stop-OfficeProcess" {
	$stword = [System.Diagnostics.Process]::pr
	$stexcel = "Stopped Excel"
	$stpowerpoint = "Stopped PPoint"
	$staccess = "Stopped Access"
	$stpublisher = "Stopped Publisher"
	$stonenote = "Stopped OneNote"
	$stoutlook = "Stopped Outlook"
	$stall = "Stopped all"
	Context "no process" {
		$list = ""
		It "return null with word switch" {
			Mock Stop-Process {Return $stword}
			Stop-OfficeProcess -word | Assert-MockCalled -Times 0 -Exactly
		}
		It "return null with excel switch" {
			Mock Stop-Process {Return $stexcel}
			Stop-OfficeProcess -excel | Should BeEmptyOrNull
		}
		It "return null with powerpoint switch" {
			Mock Stop-Process {Return $stpowerpoint}
			Stop-OfficeProcess  -powerpoint | Should BeEmptyOrNull
		}
		It "return null with access switch" {
			Mock Stop-Process {Return $staccess}
			Stop-OfficeProcess  -access | Should BeEmptyOrNull
		}
		It "return null with publisher switch" {
			Mock Stop-Process {Return $stpublisher}
			Stop-OfficeProcess  -publisher | Should BeEmptyOrNull
		}
		It "return null with onenote switch" {
			Mock Stop-Process {Return $stonenote}
			Stop-OfficeProcess  -onenote | Should BeEmptyOrNull
		}
		It "return null with outlook switch" {
			Mock Stop-Process {Return $stoutlook}
			Stop-OfficeProcess -outlook | Should BeEmptyOrNull
		}
		It "return null with all switch" {
			Mock Stop-Process {Return $stall}
			Stop-OfficeProcess -All | Should BeEmptyOrNull
		}
	}
	Context "only office process" {
		$list = "winword.exe", "excel.exe", "powerpnt.exe", "msaccess.exe", "mspub.exe", "onenote.exe", "outlook.exe"
	}
	Context "office and random process" {
		$list = "0.exe", "winword.exe", "excel.exe", "a.exe" ,"powerpnt.exe", "msaccess.exe", "b.exe" ,"mspub.exe", "onenote.exe", "c.exe", "outlook.exe"
	}
}

#Describe 'Get-AllDocx' {
#	It 'returns one text file when that is all there is' {
#	    Mock Get-ChildItem {
#	            [PSCustomObject]@{Name = 'a923e023.docx'}
#	    }
#	    Get-AllDocx | Should Be 'a923e023.docx'
#	}
#	It 'returns one text file when there are assorted files' {
#		$myList = 'a923e023.txt','wlke93jw3.doc'
#	    Mock Get-ChildItem {CreateFileList $myList}
#	        [PSCustomObject]@{Name = 'a923e023.docx'},
#	        [PSCustomObject]@{Name = 'wlke93jw3.doc'}
#	    }
#	    Get-AllDocx | Should Be 'a923e023.docx'
#	}
#	It 'returns multiple text files amongst assorted files' {
#	    Mock Get-ChildItem {
#	        [PSCustomObject]@{Name = 'a923e023.docx'},
#	        [PSCustomObject]@{Name = 'wlke93jw3.txt'},
#	        [PSCustomObject]@{Name = 'ke923jd.docx'},
#	        [PSCustomObject]@{Name = 'qq02000.jpg'}
#	    }
#	    Get-AllDocx | Should Be ('a923e023.docx','ke923jd.docx')
#	}
#	It 'returns nothing when there are no text files' {
#	    Mock Get-ChildItem {
#	        [PSCustomObject]@{Name = 'wlke93jw3.doc'},
#	        [PSCustomObject]@{Name = 'qq02000.doc'}
#	    }
#	    Get-AllDocx | Should BeNullOrEmpty
#	}
#}