
$header = @"
<?xml version="1.0" encoding="utf-8"?>
<rss version="2.0" xmlns:sparkle="http://www.andymatuschak.org/xml-namespaces/sparkle"  xmlns:dc="http://purl.org/dc/elements/1.1/">
	<channel>
		<title>Oris4Sync Version Info</title>
		<link>http://update.oris4.com/versioninfo.xml</link>
		<language>en</language>

"@

$item = @"
		<item>
			<title>Version {VERSION}</title>
			<description><![CDATA[
				{DESCRIPTION}
			]]></description>
 			<pubDate>{DATE}</pubDate>
			<enclosure url="http://update.oris4.com/{FILENAME}"
				sparkle:version="{VERSION}"
				length="{LENGTH}"
				type="application/octet-stream"
				sparkle:dsaSignature="{SIGNATURE}"/>
		</item>

"@

$footer = @"
	</channel>
</rss>
"@

$header > "versioninfo.xml"

$datetimeformatinfo = new-object system.globalization.datetimeformatinfo
$dateformat = $datetimeformatinfo.RFC1123Pattern

foreach ($file in (Get-ChildItem *.exe | Sort-Object CreationTime -descending)) {
	$filename = $file.Name
	$length = $file.Length
	$version = (Get-Item $file).VersionInfo.ProductVersion
	$signature = & ..\Extras\NetSparkle-1.0.85\NetSparkleDSAHelper.exe /sign_update $filename NetSparkle_DSA.priv
	$dateObject = Get-Date $file.CreationTime
	$date = $dateObject.ToString($dateformat)

	$description_file = "$filename.html"
	if(Test-Path $description_file){
		$description = Get-Content $description_file
	} else {
		$description = Get-Content "template.html"
	}

	$this_item = $item
	$this_item = $this_item -replace "{DESCRIPTION}", "$description"
	$this_item = $this_item -replace "{FILENAME}", "$filename"
	$this_item = $this_item -replace "{LENGTH}", "$length"
	$this_item = $this_item -replace "{VERSION}", "$version"
	$this_item = $this_item -replace "{DATE}", "$date"
	$this_item = $this_item -replace "{SIGNATURE}", "$signature"
	$this_item >> "versioninfo.xml"
}

$footer >> "versioninfo.xml"

