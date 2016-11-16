Function New-Word-Application{
	$word = New-Object Microsoft.Office.Interop.Word.ApplicationClass
	Return $word
}