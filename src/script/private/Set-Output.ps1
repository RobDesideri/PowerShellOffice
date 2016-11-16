Function Set-Output {
	param(
		[String]$outPath,
		[String]$outVal
	)
	Set-Content -Path $outPath -Value $outVal
}