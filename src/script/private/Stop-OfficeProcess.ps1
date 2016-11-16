function Stop-OfficeProcess (
	[switch]$word, 
	[switch]$excel, 
	[switch]$powerpoint, 
	[switch]$access,
	[switch]$publisher,
	[switch]$onenote,
	[switch]$outlook,
	[switch]$All
)
{
	If($word -or $All) {$r += $(Stop-Process "winword.exe" -PassThru).Id}
	If($excel -or $All) {$r += $(Stop-Process "excel.exe" -PassThru).Id}
	If($powerpoint -or $All) {$r += $(Stop-Process "powerpnt.exe" -PassThru).Id}
	If($access -or $All) {$r += $(Stop-Process "msaccess.exe" -PassThru).Id}
	If($publisher -or $All) {$r += $(Stop-Process "mspub.exe" -PassThru).Id}
	If($onenote -or $All) {$r += $(Stop-Process "onenote.exe" -PassThru).Id}
	If($outlook -or $All) {$r += $(Stop-Process "outlook.exe" -PassThru).Id}
	Return $r
}