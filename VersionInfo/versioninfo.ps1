$header = '<?xml version="1.0" encoding="utf-8"?>' + [Environment]::NewLine
$header += '<rss version="2.0" xmlns:sparkle="http://www.andymatuschak.org/xml-namespaces/sparkle"  xmlns:dc="http://purl.org/dc/elements/1.1/">' + [Environment]::NewLine
$header += '	<channel>' + [Environment]::NewLine
$header += '		<title>Oris4Sync Version Info</title>' + [Environment]::NewLine
$header += '		<link>{URL}/versioninfo.xml</link>' + [Environment]::NewLine
$header += '		<language>en</language>' + [Environment]::NewLine
$header += [Environment]::NewLine

$item = '		<item>' + [Environment]::NewLine
$item += '			<title>Version {VERSION}</title>' + [Environment]::NewLine
$item += '			<description><![CDATA[' + [Environment]::NewLine
$item += '				{DESCRIPTION}' + [Environment]::NewLine
$item += '			]]></description>' + [Environment]::NewLine
$item += ' 			<pubDate>{DATE}</pubDate>' + [Environment]::NewLine
$item += '			<enclosure url="{URL}/{FILENAME}"' + [Environment]::NewLine
$item += '				sparkle:version="{VERSION}"' + [Environment]::NewLine
$item += '				length="{LENGTH}"' + [Environment]::NewLine
$item += '				type="application/octet-stream"' + [Environment]::NewLine
$item += '				sparkle:dsaSignature="{SIGNATURE}"/>' + [Environment]::NewLine
$item += '		</item>' + [Environment]::NewLine
$item += [Environment]::NewLine

$footer = '	</channel>' + [Environment]::NewLine
$footer += '</rss>' + [Environment]::NewLine

$datetimeformatinfo = new-object system.globalization.datetimeformatinfo
$dateformat = $datetimeformatinfo.RFC1123Pattern

$directories = "ROOT", "ROOT\debug"

foreach ($directory in $directories)
{
	$url = "http://update.oris4.com"
	if ($directory -eq "ROOT\debug")
	{
		$url = "$url/debug"	
	}
	$body = $header -replace "{URL}", "$url"

	foreach ($file in (Get-ChildItem "$directory\*.exe" | Sort-Object CreationTime -descending))
	{
		$filename = $file.Name
		$length = $file.Length
		$version = (Get-Item $file).VersionInfo.ProductVersion
		$signature = & ..\Extras\NetSparkle-1.0.85\NetSparkleDSAHelper.exe /sign_update "$directory\$filename" NetSparkle_DSA.priv
		$dateObject = Get-Date $file.CreationTime
		$date = $dateObject.ToString($dateformat)

		$description_file = "$directory\$filename.html"
		if(Test-Path $description_file){
			$description = Get-Content $description_file
		} else {
			$description = Get-Content "template.html"
		}

		$this_item = $item -replace "{DESCRIPTION}", "$description"
		$this_item = $this_item -replace "{FILENAME}", "$filename"
		$this_item = $this_item -replace "{LENGTH}", "$length"
		$this_item = $this_item -replace "{VERSION}", "$version"
		$this_item = $this_item -replace "{DATE}", "$date"
		$this_item = $this_item -replace "{SIGNATURE}", "$signature"
		$this_item = $this_item -replace "{URL}", "$url"

		$body += $this_item
	}

	$body += $footer

	Set-Content -Path "$directory\versioninfo.xml" -Value "$body" 
}

