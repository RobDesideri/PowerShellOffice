Function Get-Metadata-As-Csv{
	param(
        [Parameter(Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[ValidateScript({$_.Count -gt 0})] 
        [String[]]
        $filePaths
    ) 
	$word = New-Word-Application
	$outList = New-Generic-List
	$outListLine = New-Generic-List
	$outListVal = New-Generic-List
	$outListValCom = New-Generic-List
	$binding = Get-BindingFlag
	
	$cp = @("-> COMMENT: ", "   SCOPE: ", "   DATE: ","   AUTHOR: ") #prefixed strings for comments details
	$header = "Artifact, Title, Version, Comments" #CSV header
	$lf = "`r`n" #linefeed shortcut
	$n = "" #value for empty CSV value
	
	
	#Iterate trough files
	foreach ($item in $filePaths)
	{
		$filePath = $item.ToString()
		$outListVal.Add("$filePath") #CSV[1] adding
		#Open document in only-read mode
		$document = $word.Documents.Open("$filePath", $null, $true)
		$properties = $document.BuiltInDocumentProperties
		#Iterate trough BuiltInDocumentProperties
		foreach($property in $properties) 
		{
			if($versCkr -and $titleCkr) {$versCkr = $false; $titleCkr = $false; break} #checker for perfomance purpose (we need only 2 property...)
			$pn = [System.__ComObject].InvokeMember("name", $binding, $null, $property, $null)
			try
			{
				switch ($pn)
				{
					'Content status'{$str = [System.__ComObject].InvokeMember("value", $binding, $null, $property, $null); $versCkr = $true} #CSV[2]
					'Title' {$str = [System.__ComObject].InvokeMember("value", $binding, $null, $property, $null); $titleCkr = $true} #CSV[3]
					default {$str = $null}
				}
			}
			catch [System.Exception]
			{
				$str = $n #CSV[2] / #CSV[3] empty value
			}
			finally
			{
				if(-not($str -eq $null)) {$outListVal.Add($str)} #CSV[2] / CSV[3] adding
			}
		}
		$comments = $document.Comments
		#Iterate trough Comments
		foreach ($com in $comments)
		{
			try
			{
				$outListValCom.Add([String]::Concat($cp[0], $com.Range.Text)) #comment text
			}
			catch [System.Exception]
			{
				$outListValCom.Add($cp[0] + $n)
			}
			finally
			{
				try
				{
					$outListValCom.Add([String]::Concat($cp[1], $com.Scope.Text.ToString())) #comment scope
				}
				catch [System.Exception]
				{
					$outListValCom.Add($cp[1] + $n)
				}
				finally
				{
					$outListValCom.Add([String]::Concat($cp[2], $com.Date)) #comment date
					$outListValCom.Add([String]::Concat($cp[3], $com.Contact.Name)) #comment author
				}

			}
		}
		if ($outListValCom.Count -eq 0) {$outListValCom.Add($n)} #If no comments insert empty value
		$outListValComString = [String]::Join($lf, $outListValCom.ToArray()) #CSV[4]
		$outListVal.Add($outListValComString) #CSV[4] adding
		$outListValString = [String]::Join(",", $outListVal.ToArray()) #CSV line
		$outListLine.Add($outListValString) #CSV line added
		#dispose list and document
		$outListValCom.Clear()
		$outListValComString = ""
		$outListVal.Clear()
		$outListValString = ""
		$document.Close($false, [System.Type]::Missing, [System.Type]::Missing)
	}
	$word.Quit()
	$outListString = [String]::Join($lf, $outListLine.ToArray()) #all CSV lines
	$out = [String]::Concat($header, $lf, $outListString) #final CSV
	Return $out
}