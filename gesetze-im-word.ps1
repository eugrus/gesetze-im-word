if ($args.Length -lt 2 -or $args.Length -gt 2) {
    Write-Host "Mit diesem Tool können Sie über PowerShell Gesetzesparagraphen schnell in Microsoft Word hineinkopieren."
    Write-Host "Eingabecursor in Word an die richtige Stelle setzen und das Skript wie folgt aufrufen:"
    Write-Host ".\gesetze-im-word 109d StGB"
    Write-Host
    Write-Host "https://github.com/eugrus/gesetze-im-word - Evgeny Fishgalov - 2023"
    Write-Host
    Write-Host "Für urheberrechtlich relevante Handlungen gelten die nachstehenden Bedingungen mit der Maßgabe, dass auch entgeltliche juristische Dienstleistungen als kommerzielle Nutzung gelten: https://creativecommons.org/licenses/by-nc-sa/3.0/de/deed"
    Write-Host "Dieses Skript erstellte Evgeny Fishgalov in seiner Freizeit."
    Write-Host "Die eigene Verwendung von diesem Skript durch Evgeny Fishgalov an seinem Arbeitsplatz bei der Erfüllung seiner Arbeit begründet keine Erlaubnis für eine selbstständige lizenzlose Weiternutzung für kommerzielle Zwecke durch den Arbeitgeber."
    exit 1
}

$Paragraf = $args[0]

$Gesetz = $args[1].ToLower()

$url = "https://www.gesetze-im-internet.de/$($Gesetz)/__$($Paragraf).html"

$html = Invoke-RestMethod -Uri $url

function Decode-Html {
    param (
        [string]$inputHtml
    )

    $decodedHtml = [System.Web.HttpUtility]::HtmlDecode($inputHtml) -replace '<[^>]*>', ''
    return $decodedHtml
}

$output = $html | Select-String -Pattern '<div class="jnhtml">.*</div>' | ForEach-Object {
	$_.Matches.Value -replace '<div class="jurAbsatz">', "`n<div class='jurAbsatz'>" -replace '<dt>', "`n<dt> " -replace '<div>', "`n<div> "
} | ForEach-Object {
	Decode-Html $_
}

Write-Host $output

$word = [System.Runtime.Interopservices.Marshal]::GetActiveObject("Word.Application")

$word.Selection.TypeText($output)

Write-Host "In Word hineinkopiert."
