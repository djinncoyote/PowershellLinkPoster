[CmdletBinding()]
param(
	[string]$pinboardUser = "",
	[string]$dateFrom = "",
	[string]$dateTo = "",
	[string]$emailFrom = "",
	[string]$emailTo = "",
	[pscredential]$proxyCredentials,
	[string]$smtpServer = "va-mail01.dreamwidth.org",
	[switch]$TestMode
)

$envVarName = "LINKPOSTER_TIMESTAMP_$pinboardUser"

if($dateFrom.Equals("")){
	$dateFrom = [Environment]::GetEnvironmentVariable($envVarName, "User")

	if(!$dateFrom.Equals("")){
		Write-Verbose "Environment variable was found: $dateFrom"
	}
}
if(!$dateFrom.Equals("")){
	$startDateTime = [DateTime]::ParseExact($dateFrom, "yyyy-MM-ddTHH:mm:ss", $null)
}

if($dateTo.Equals("")){
	# Use the current date and time
	$endDateTime = Get-Date
	$dateTo = Get-Date -Date $endDateTime -Format "yyyy-MM-ddTHH:mm:ss"
}
else{
	$endDateTime = [DateTime]::ParseExact($dateTo, "yyyy-MM-ddTHH:mm:ss", $null)
}


$wc = New-Object System.Net.WebClient
if($proxyCredentials){
	Write-Verbose "Fetching from Pinboard using Proxy"
	$wc.Proxy.Credentials = $proxyCredentials
}

$pinboardUrl = "https://feeds.pinboard.in/rss/u:$pinboardUser/"
Write-Verbose "Fetching from $pinboardUrl"
[xml]$feed = $wc.DownloadString($pinboardUrl)

# $feed.rdf.item.count doesn't work right; it returns 1 when there are no items because XmlElement has an "Item" property.
$itemNodes = $feed.GetElementsByTagName("item")

Write-Verbose "Feed has $($itemNodes.count) entries"

if($dateFrom.Equals("")){
	Write-Verbose "Checking for links posted before: $dateTo"

	$items = $itemNodes | ? {[DateTime]::Parse($_.date) -LE $endDateTime} | sort {[DateTime]::Parse($_.date)}
}
else{
	Write-Verbose "Checking for links posted between: $dateFrom and $dateTo"

	$items = $itemNodes | ? {[DateTime]::Parse($_.date) -gt $startDateTime} | ? {[DateTime]::Parse($_.date) -LE $endDateTime} | sort {[DateTime]::Parse($_.date)}
}

$itemCount = 0
if($items){
	if($items.count){
		$itemCount = $items.Count
	}
	else{
		$itemCount = 1
	}
}

Write-Verbose "$itemCount items selected"


if($items){
	$tags = @()
	$output = "<dl class=`"links`">"
	foreach($item in $items){
		$output+="<dt class=`"link`"><a href=`"$($item.link)`" rel=`"nofollow`">$($item.Title)</a></dt>"
		$output+="<dd style=`"margin-bottom: 0.5em;`">"
		if($item.Description){
			$output += "<span class=`"link-description`">$($item.description.'#cdata-section')</span><BR/>"
		}
		if($item.subject){
			$output += "<small class=`"link-tags`">(tags:"
			foreach($tag in ($item.subject -split " ")){
				$tags += $tag
				$output += "<A href=`"https://pinboard.in/u:$pinboardUser/t:$tag`">$tag</A> "
			} 
			$output += ")</small>"
		}
		$output += "</dd>"
	}
	$output += "</dl>"
	
	$output += "`n`n--`n`nDeletionTrigger"
	
	if($tags){
		$tags += "links"
		$tagsHeader = "post-tags: " + (($tags | sort -unique) -join ", ")
		$output = $tagsHeader + "`n`n" + $output
	}

	$subjectLink = "Interesting Links for $($endDateTime.ToString("dd-MM-yyyy"))"

	$output = $output -replace "‘|’","'"
	$output = $output -replace '–',"-"
	$output = $output -replace "`“|`”",'"'
	
	if($TestMode){
		$output
	}
	else{
		Send-MailMessage -From $emailFrom -To $emailTo -Subject $subjectLink -Body $output -SmtpServer $smtpServer
		Write-Verbose "Selected links have been posted"
	}
}

if(!$TestMode){
	[Environment]::SetEnvironmentVariable($envVarName, $dateTo, "User")
	Write-Verbose "Updated environment variable to: $dateTo"
}
