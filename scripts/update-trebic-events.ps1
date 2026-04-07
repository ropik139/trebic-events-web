[CmdletBinding()]
param(
    [string]$ConfigPath = "..\\trebic-events.settings.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-ProjectPath {
    param([string]$Path)
    $scriptRoot = Split-Path -Parent $PSCommandPath
    [System.IO.Path]::GetFullPath((Join-Path $scriptRoot $Path))
}

function Resolve-PathFromBase {
    param(
        [string]$BasePath,
        [string]$RelativePath
    )

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        return $RelativePath
    }

    [System.IO.Path]::GetFullPath((Join-Path $BasePath $RelativePath))
}

function Get-TimeZoneInfoSafe {
    param([string]$PreferredId)

    foreach ($id in @($PreferredId, "Europe/Prague", "Central Europe Standard Time")) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        try {
            return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
        } catch {
        }
    }

    [System.TimeZoneInfo]::Local
}

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    $raw | ConvertFrom-Json
}

function Write-JsonFile {
    param(
        [string]$Path,
        [object]$Value
    )

    Ensure-Directory -Path (Split-Path -Parent $Path)
    $Value | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Get-StringContent {
    param([string]$Url)
    (Invoke-WebRequest -UseBasicParsing -Headers @{ "User-Agent" = "CodexTrebicEvents/1.0" } -Uri $Url).Content
}

function Decode-Html {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    [System.Net.WebUtility]::HtmlDecode($Value)
}

function Normalize-Whitespace {
    param([string]$Value)
    $text = Decode-Html -Value $Value
    $text = $text -replace "(?is)<br\\s*/?>", "`n"
    $text = $text -replace "(?is)<[^>]+>", " "
    $text = $text -replace "[\u00A0\r\n\t]+", " "
    $text = $text -replace "\s{2,}", " "
    $text.Trim()
}

function Convert-HtmlFragmentToText {
    param([string]$Html)
    if ([string]::IsNullOrWhiteSpace($Html)) { return "" }
    $text = $Html -replace "(?is)<script.*?</script>", " " -replace "(?is)<style.*?</style>", " "
    $text = $text -replace "(?i)</p>", "`n`n" -replace "(?i)<br\\s*/?>", "`n" -replace "<[^>]+>", " "
    $text = Decode-Html -Value $text
    $text = $text -replace "[\u00A0\r\t]+", " " -replace " +", " " -replace "\n{3,}", "`n`n"
    $text.Trim()
}

function Resolve-AbsoluteUrl {
    param(
        [string]$Url,
        [string]$BaseUrl
    )

    if ([string]::IsNullOrWhiteSpace($Url)) { return "" }
    if ($Url -match '^https?://') { return $Url }
    [System.Uri]::new([System.Uri]::new($BaseUrl), $Url).AbsoluteUri
}

function Get-FirstMatchValue {
    param(
        [string]$Text,
        [string]$Pattern
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $match = [regex]::Match($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    if ($match.Success) { return $match.Groups["value"].Value }
    ""
}

function Get-NormalizedPlaceKey {
    param([string]$Value)
    $text = Normalize-Whitespace -Value $Value
    $text = $text.ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder
    foreach ($char in $text.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }
    (($builder.ToString().Normalize([Text.NormalizationForm]::FormC)) -replace '[^a-z0-9]+', ' ').Trim()
}

function Get-GenreDisplayHtml {
    param([string]$Genre)

    switch ($Genre) {
        "Divadlo" { return "Divadlo" }
        "Koncerty" { return "Koncerty" }
        "Vystavy" { return "V&#253;stavy" }
        "Kino" { return "Kino" }
        "Pro deti" { return "Pro d&#283;ti" }
        "Prednasky a workshopy" { return "P&#345;edn&#225;&#353;ky a workshopy" }
        "Festivaly a slavnosti" { return "Festivaly a slavnosti" }
        "Zabava a talk show" { return "Z&#225;bava a talk show" }
        default { return "Ostatn&#237;" }
    }
}

function Parse-CzechDateText {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $formats = @("d.M.yyyy", "dd.MM.yyyy", "d. M. yyyy", "dd. MM. yyyy")
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo("cs-CZ")
    foreach ($format in $formats) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value.Trim(), $format, $culture, [System.Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
            return $parsed
        }
    }

    $fallback = [datetime]::MinValue
    if ([datetime]::TryParse($Value.Trim(), $culture, [System.Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$fallback)) {
        return $fallback
    }

    $null
}

function Parse-TimeText {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $normalized = $Value.Trim().ToLowerInvariant()
    $normalized = $normalized -replace "^od\s+", "" -replace "^v\s+", "" -replace "\s+", ""
    if ($normalized -match '^(?<hour>\d{1,2})(?::(?<minute>\d{1,2}))?$') {
        $minute = 0
        if ($matches["minute"]) {
            $minute = [int]$matches["minute"]
        }
        return [timespan]::FromHours([int]$matches["hour"]) + [timespan]::FromMinutes($minute)
    }
    $null
}

function Parse-DateRange {
    param(
        [string]$Value,
        [string]$StartTimeText,
        [string]$EndTimeText
    )

    $clean = Normalize-Whitespace -Value $Value
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    $parts = $clean -split '\s*-\s*'
    $startDate = Parse-CzechDateText -Value $parts[0]
    if ($null -eq $startDate) { return $null }
    $endDate = if ($parts.Count -gt 1) { Parse-CzechDateText -Value $parts[-1] } else { $startDate }
    if ($null -eq $endDate) { $endDate = $startDate }

    $startTime = Parse-TimeText -Value $StartTimeText
    $endTime = Parse-TimeText -Value $EndTimeText
    $startAt = if ($null -ne $startTime) { $startDate.Date.Add($startTime) } else { $startDate.Date }
    $endAt = if ($null -ne $endTime) { $endDate.Date.Add($endTime) } elseif ($parts.Count -gt 1) { $endDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59) } elseif ($null -ne $startTime) { $startAt.AddHours(3) } else { $endDate.Date.AddHours(23).AddMinutes(59).AddSeconds(59) }

    [pscustomobject]@{
        startAt = $startAt
        endAt   = $endAt
    }
}

function Get-HaversineDistanceKm {
    param(
        [double]$Latitude1,
        [double]$Longitude1,
        [double]$Latitude2,
        [double]$Longitude2
    )

    $radiusKm = 6371.0
    $degToRad = [math]::PI / 180.0
    $deltaLat = ($Latitude2 - $Latitude1) * $degToRad
    $deltaLon = ($Longitude2 - $Longitude1) * $degToRad
    $a = [math]::Sin($deltaLat / 2) * [math]::Sin($deltaLat / 2) + [math]::Cos($Latitude1 * $degToRad) * [math]::Cos($Latitude2 * $degToRad) * [math]::Sin($deltaLon / 2) * [math]::Sin($deltaLon / 2)
    $c = 2 * [math]::Atan2([math]::Sqrt($a), [math]::Sqrt(1 - $a))
    $radiusKm * $c
}

function Get-CacheMap {
    param([object]$CacheObject)
    $map = @{}
    if ($null -eq $CacheObject) { return $map }
    foreach ($property in $CacheObject.PSObject.Properties) {
        $map[$property.Name] = $property.Value
    }
    $map
}

function Save-LocationCache {
    param(
        [string]$Path,
        [hashtable]$Cache
    )

    $payload = [ordered]@{}
    foreach ($key in ($Cache.Keys | Sort-Object)) {
        $payload[$key] = $Cache[$key]
    }
    Write-JsonFile -Path $Path -Value $payload
}

function Resolve-PlaceCoordinates {
    param(
        [string]$PlaceName,
        [hashtable]$KnownPlaceMap,
        [hashtable]$LocationCache
    )

    if ([string]::IsNullOrWhiteSpace($PlaceName)) { return $null }
    $key = Get-NormalizedPlaceKey -Value $PlaceName
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    if ($KnownPlaceMap.ContainsKey($key)) { return $KnownPlaceMap[$key] }
    if ($LocationCache.ContainsKey($key)) { return $LocationCache[$key] }

    foreach ($query in @("$PlaceName, okres Trebic, Czechia", "$PlaceName, Trebic, Czechia", "$PlaceName, Czechia")) {
        $url = "https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&countrycodes=cz&q=$([uri]::EscapeDataString($query))"
        try {
            $result = Invoke-RestMethod -Headers @{ "User-Agent" = "CodexTrebicEvents/1.0" } -Uri $url
            if ($null -ne $result -and $result.Count -gt 0) {
                $first = $result[0]
                $location = [pscustomobject]@{
                    latitude  = [double]$first.lat
                    longitude = [double]$first.lon
                    source    = "nominatim"
                    display   = [string]$first.display_name
                }
                $LocationCache[$key] = $location
                return $location
            }
        } catch {
            Start-Sleep -Milliseconds 300
        }
    }

    $null
}

function Get-GenreFromText {
    param(
        [string]$Title,
        [string]$Keywords,
        [string]$ShortCode,
        [string]$Summary,
        [string]$DeclaredGenre,
        [string]$Venue,
        [string]$Link
    )

    if (-not [string]::IsNullOrWhiteSpace($DeclaredGenre)) { return $DeclaredGenre }
    $haystack = Get-NormalizedPlaceKey -Value ("$Title $Keywords $ShortCode $Summary")
    $venueKey = Get-NormalizedPlaceKey -Value $Venue
    $linkKey = Get-NormalizedPlaceKey -Value $Link

    if ($venueKey -match 'hvezdarna|planetarium') {
        if ($haystack -match 'deti|rodin|animov') { return 'Pro deti' }
        if ($haystack -match 'prednask|prednasek|seminar|workshop|diln|kurz|beseda') { return 'Prednasky a workshopy' }
        return 'Ostatni'
    }

    if ($venueKey -match 'kino' -or $linkKey -match 'mkstrebic cz kino| kino ') {
        return 'Kino'
    }

    $rules = @(
        @{ Pattern = 'koncert|hudba|trio|kvartet|quartet|kapela|recital|orchestr|cello|filharmon'; Genre = 'Koncerty' },
        @{ Pattern = 'divad|predstaveni|komedie'; Genre = 'Divadlo' },
        @{ Pattern = 'vystav|expozic|galeri|vernisaz|dernisaz'; Genre = 'Vystavy' },
        @{ Pattern = 'kino|projekce|promit'; Genre = 'Kino' },
        @{ Pattern = 'deti|rodin|loutk|pohad'; Genre = 'Pro deti' },
        @{ Pattern = 'prednask|prednasek|seminar|workshop|diln|kurz|beseda'; Genre = 'Prednasky a workshopy' },
        @{ Pattern = 'festival|slavnost|jarmark|food|trh'; Genre = 'Festivaly a slavnosti' },
        @{ Pattern = 'show|stand up|talk|zabav'; Genre = 'Zabava a talk show' }
    )

    foreach ($rule in $rules) {
        if ($haystack -match $rule.Pattern) { return $rule.Genre }
    }
    "Ostatni"
}

function Get-RegionCalendarListItems {
    param([string]$ListUrl)

    $html = Get-StringContent -Url $ListUrl
    $pattern = '<a href="\?cid=(?<id>\d+)">\s*<div class="polozka_vypis">.*?<img src="(?<image>[^"]+)".*?<div class="polozka_vypis_datum_in">\s*(?<date>.*?)\s*</div>.*?<h2 class="h2_vypis">(?<title>.*?)</h2>.*?<div class="umisteni_clanku_vypis">(?<place>.*?)</div>'
    $matches = [regex]::Matches($html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($match in $matches) {
        $items.Add([pscustomobject]@{
            sourceKey   = "region_calendar"
            sourceLabel = "Regionalni kalendar Trebicsko-Moravska Vysocina"
            externalId  = [string]$match.Groups["id"].Value
            title       = Normalize-Whitespace -Value $match.Groups["title"].Value
            place       = Normalize-Whitespace -Value $match.Groups["place"].Value
            dateText    = Normalize-Whitespace -Value $match.Groups["date"].Value
            imageUrl    = Resolve-AbsoluteUrl -Url $match.Groups["image"].Value -BaseUrl $ListUrl
            detailUrl   = "https://kalendar.trebicsko-moravskavysocina.cz/iframe.php?cid=$([string]$match.Groups["id"].Value)"
        })
    }

    $items
}

function Get-DetailParameterMap {
    param([string]$Html)
    $pattern = '<div class="polozka_parametry_nazev">(?<name>.*?)</div>\s*<div class="polozka_parametry_hodnota">(?<value>.*?)</div>|<div class="polozka_parametry_cela">\s*<div class="polozka_parametry_nazev">(?<name2>.*?)</div>\s*<div class="polozka_parametry_hodnota">(?<value2>.*?)</div>'
    $matches = [regex]::Matches($Html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    $map = @{}
    foreach ($match in $matches) {
        $name = if ($match.Groups["name"].Success) { $match.Groups["name"].Value } else { $match.Groups["name2"].Value }
        $value = if ($match.Groups["value"].Success) { $match.Groups["value"].Value } else { $match.Groups["value2"].Value }
        $cleanName = Normalize-Whitespace -Value $name
        if (-not [string]::IsNullOrWhiteSpace($cleanName)) {
            $map[(Get-NormalizedPlaceKey -Value $cleanName.TrimEnd(':'))] = Convert-HtmlFragmentToText -Html $value
        }
    }
    $map
}

function Get-RegionCalendarDetailItem {
    param([object]$ListItem)

    $html = Get-StringContent -Url $ListItem.detailUrl
    $parameterMap = Get-DetailParameterMap -Html $html
    $titleMatch = [regex]::Match($html, '<h2 class="h2_clanek"[^>]*>(?<value>.*?)</h2>', [Text.RegularExpressions.RegexOptions]::Singleline)
    $title = if ($titleMatch.Success) { Normalize-Whitespace -Value $titleMatch.Groups["value"].Value } else { $ListItem.title }
    $detailImageUrl = Get-FirstMatchValue -Text $html -Pattern '<a[^>]+data-fancybox="gallery1"[^>]+href="(?<value>[^"]+)"'
    if ([string]::IsNullOrWhiteSpace($detailImageUrl)) {
        $detailImageUrl = Get-FirstMatchValue -Text $html -Pattern '<img[^>]+src="(?<value>[^"]+mini\.(?:jpg|jpeg|png|webp))"'
    }

    $dateText = $ListItem.dateText
    if ($parameterMap.ContainsKey("datum od")) {
        $endDateText = if ($parameterMap.ContainsKey("datum do")) { $parameterMap["datum do"] } else { $parameterMap["datum od"] }
        $dateText = "$($parameterMap["datum od"]) - $endDateText"
    }

    $timeFrom = if ($parameterMap.ContainsKey("cas od")) { $parameterMap["cas od"] } else { "" }
    $timeTo = if ($parameterMap.ContainsKey("cas do")) { $parameterMap["cas do"] } else { "" }
    $range = Parse-DateRange -Value $dateText -StartTimeText $timeFrom -EndTimeText $timeTo
    if ($null -eq $range) { return $null }

    $place = if ($parameterMap.ContainsKey("obec")) { $parameterMap["obec"] } else { $ListItem.place }
    $venue = if ($parameterMap.ContainsKey("misto udalosti")) { $parameterMap["misto udalosti"] } else { $place }
    $keywords = if ($parameterMap.ContainsKey("klicova slova udalosti")) { $parameterMap["klicova slova udalosti"] } else { "" }
    $shortCode = if ($parameterMap.ContainsKey("zkratka udalosti")) { $parameterMap["zkratka udalosti"] } else { "" }
    $summary = if ($parameterMap.ContainsKey("zkraceny vypis")) { $parameterMap["zkraceny vypis"] } elseif ($parameterMap.ContainsKey("text")) { $parameterMap["text"] } else { "" }
    $link = if ($parameterMap.ContainsKey("odkaz udalosti")) { $parameterMap["odkaz udalosti"] } else { $ListItem.detailUrl }

    [pscustomobject]@{
        sourceKey    = $ListItem.sourceKey
        sourceLabel  = $ListItem.sourceLabel
        title        = $title
        genre        = Get-GenreFromText -Title $title -Keywords $keywords -ShortCode $shortCode -Summary $summary -DeclaredGenre "" -Venue $venue -Link $link
        municipality = $place
        venue        = $venue
        startAt      = $range.startAt
        endAt        = $range.endAt
        startText    = $range.startAt.ToString("d. M. yyyy HH:mm")
        endText      = $range.endAt.ToString("d. M. yyyy HH:mm")
        dateLabel    = $ListItem.dateText
        timeLabel    = (($timeFrom, $timeTo | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join " - ")
        summary      = $summary
        keywords     = $keywords
        link         = $link
        detailLink   = $ListItem.detailUrl
        imageUrl     = if (-not [string]::IsNullOrWhiteSpace($detailImageUrl)) { Resolve-AbsoluteUrl -Url $detailImageUrl -BaseUrl $ListItem.detailUrl } else { $ListItem.imageUrl }
        dedupeKey    = "$($range.startAt.ToString("yyyy-MM-dd"))|$((Get-NormalizedPlaceKey -Value $title))"
    }
}

function Get-MksCategoryItems {
    param(
        [string]$SourceKey,
        [string]$SourceLabel,
        [string]$Url,
        [string]$Genre
    )

    $html = Get-StringContent -Url $Url
    $pattern = '<div class="polozka_vypis_kino">.*?<a href="(?<href>[^"]+)">.*?<img src="(?<image>[^"]+)".*?<h2 class="nazev_vypis_kino">(?<title>.*?)</h2>(?:\s*<h3 class="podnazev_vypis_kino">(?<subtitle>.*?)</h3>)?.*?ikona_vypis_kino_misto.*?<span class="polozka_data_vypis_kino_hodnota">(?<venue>.*?)</span>.*?ikona_vypis_kino_datum.*?<span class="polozka_data_vypis_kino_hodnota">\s*(?<date>.*?)\s*</span>.*?ikona_vypis_kino_cas.*?<span class="polozka_data_vypis_kino_hodnota">\s*(?<time>.*?)\s*</span>'
    $matches = [regex]::Matches($html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($match in $matches) {
        $dateText = Normalize-Whitespace -Value $match.Groups["date"].Value
        $timeText = Normalize-Whitespace -Value $match.Groups["time"].Value
        $range = Parse-DateRange -Value $dateText -StartTimeText $timeText -EndTimeText ""
        if ($null -eq $range) { continue }

        $title = Normalize-Whitespace -Value $match.Groups["title"].Value
        $subtitle = Normalize-Whitespace -Value $match.Groups["subtitle"].Value
        $items.Add([pscustomobject]@{
            sourceKey    = $SourceKey
            sourceLabel  = $SourceLabel
            title        = $title
            genre        = $Genre
            municipality = "Trebic"
            venue        = Normalize-Whitespace -Value $match.Groups["venue"].Value
            startAt      = $range.startAt
            endAt        = $range.endAt
            startText    = $range.startAt.ToString("d. M. yyyy HH:mm")
            endText      = $range.endAt.ToString("d. M. yyyy HH:mm")
            dateLabel    = $dateText
            timeLabel    = $timeText
            summary      = $subtitle
            keywords     = $Genre
            link         = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
            detailLink   = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
            imageUrl     = Resolve-AbsoluteUrl -Url $match.Groups["image"].Value -BaseUrl $Url
            dedupeKey    = "$($range.startAt.ToString("yyyy-MM-dd"))|$((Get-NormalizedPlaceKey -Value $title))"
        })
    }

    $items
}

function Get-UnescoItems {
    param(
        [string]$SourceKey,
        [string]$SourceLabel,
        [string]$Url
    )

    $html = Get-StringContent -Url $Url
    $pattern = '<div class="article">.*?<a href="(?<href>/[^"]+)"><img class="akce_foto" src="(?<image>[^"]+)".*?</div>.*?<a href="[^"]+">(?<title>.*?)</a></h4>.*?<span class="date">(?<date>.*?)</span><br/>\s*<p>\s*(?<place>.*?)<br/>\s*(?<summary>.*?)\s*</p>'
    $matches = [regex]::Matches($html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($match in $matches) {
        $dateText = Normalize-Whitespace -Value $match.Groups["date"].Value
        if ($dateText -match '(?<start>\d{1,2}\.\d{1,2}\.\d{4}(?:\s+\d{1,2}:\d{2})?).*?-\s*(?<end>\d{1,2}\.\d{1,2}\.\d{4})') {
            $range = Parse-DateRange -Value ("{0} - {1}" -f $matches["start"], $matches["end"]) -StartTimeText "" -EndTimeText ""
        } else {
            $cleanDate = ($dateText -replace '^[^\d]+', '')
            $range = Parse-DateRange -Value $cleanDate -StartTimeText "" -EndTimeText ""
        }
        if ($null -eq $range) { continue }

        $placeText = Normalize-Whitespace -Value $match.Groups["place"].Value
        $municipality = if ($placeText -match '\((?<city>[^)]+)\)') { $matches["city"] } else { "Trebic" }

        $title = Normalize-Whitespace -Value $match.Groups["title"].Value
        $summary = Normalize-Whitespace -Value $match.Groups["summary"].Value
        $detailUrl = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
        $items.Add([pscustomobject]@{
            sourceKey    = $SourceKey
            sourceLabel  = $SourceLabel
            title        = $title
            genre        = Get-GenreFromText -Title $title -Keywords "" -ShortCode "" -Summary $summary -DeclaredGenre "" -Venue $placeText -Link $detailUrl
            municipality = $municipality
            venue        = $placeText
            startAt      = $range.startAt
            endAt        = $range.endAt
            startText    = $range.startAt.ToString("d. M. yyyy HH:mm")
            endText      = $range.endAt.ToString("d. M. yyyy HH:mm")
            dateLabel    = $dateText
            timeLabel    = ""
            summary      = $summary
            keywords     = ""
            link         = $detailUrl
            detailLink   = $detailUrl
            imageUrl     = Resolve-AbsoluteUrl -Url $match.Groups["image"].Value -BaseUrl $Url
            dedupeKey    = "$($range.startAt.ToString("yyyy-MM-dd"))|$((Get-NormalizedPlaceKey -Value $title))"
        })
    }

    $items
}

function Get-DdmItems {
    param(
        [string]$SourceKey,
        [string]$SourceLabel,
        [string]$Url
    )

    $html = Get-StringContent -Url $Url
    $pattern = '<li class="js-param-search-product"[^>]*>.*?<h3 class="title" title="(?<title>[^"]+)"><a href="(?<href>[^"]+)">.*?</a></h3>.*?<div class="description">\s*<p>(?<summary>.*?)</p>.*?attr-termin-konani.*?<div class="attributes-cell">\s*(?<date>.*?)\s*</div>.*?attr-misto-konani.*?<div class="attributes-cell">\s*(?<place>.*?)\s*</div>.*?attr-zamereni.*?<div class="attributes-cell">\s*(?<focus>.*?)\s*</div>.*?<img[^>]+data-src="(?<image>[^"]+)"'
    $matches = [regex]::Matches($html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    $items = New-Object System.Collections.Generic.List[object]

    foreach ($match in $matches) {
        $dateText = Normalize-Whitespace -Value $match.Groups["date"].Value
        $range = Parse-DateRange -Value $dateText -StartTimeText "" -EndTimeText ""
        if ($null -eq $range) { continue }

        $title = Normalize-Whitespace -Value $match.Groups["title"].Value
        $summary = Normalize-Whitespace -Value $match.Groups["summary"].Value
        $focus = Normalize-Whitespace -Value $match.Groups["focus"].Value
        $place = Normalize-Whitespace -Value $match.Groups["place"].Value

        $items.Add([pscustomobject]@{
            sourceKey    = $SourceKey
            sourceLabel  = $SourceLabel
            title        = $title
            genre        = Get-GenreFromText -Title $title -Keywords $focus -ShortCode "" -Summary $summary -DeclaredGenre "" -Venue $place -Link $detailUrl
            municipality = if ($place -match 'heraltice|skripina') { $place } else { "Trebic" }
            venue        = $place
            startAt      = $range.startAt
            endAt        = $range.endAt
            startText    = $range.startAt.ToString("d. M. yyyy HH:mm")
            endText      = $range.endAt.ToString("d. M. yyyy HH:mm")
            dateLabel    = $dateText
            timeLabel    = ""
            summary      = $summary
            keywords     = $focus
            link         = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
            detailLink   = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
            imageUrl     = Resolve-AbsoluteUrl -Url ($match.Groups["image"].Value -replace '&amp;', '&') -BaseUrl $Url
            dedupeKey    = "$($range.startAt.ToString("yyyy-MM-dd"))|$((Get-NormalizedPlaceKey -Value $title))"
        })
    }

    $items
}

function Get-MuzeumScheduleItems {
    param(
        [string]$SourceKey,
        [string]$SourceLabel,
        [string]$Url
    )

    $html = Get-StringContent -Url $Url
    $entries = [regex]::Matches($html, '<(?:strong|em|span)[^>]*>(?<date>[^<]*\d{1,2}\.[^<]*)</(?:strong|em|span)>\s*-\s*(?:&nbsp;)?\s*(?:<a href="(?<href>[^"]+)">)?(?<title>[^<,]+)', [Text.RegularExpressions.RegexOptions]::Singleline)
    $slideshowImages = [regex]::Matches($html, '<img src="(?<src>/data/slideshow/22/[^"]+)" alt="(?<alt>[^"]+)"', [Text.RegularExpressions.RegexOptions]::Singleline)
    $imageMap = @{}
    foreach ($img in $slideshowImages) {
        $imageMap[(Get-NormalizedPlaceKey -Value $img.Groups["alt"].Value)] = Resolve-AbsoluteUrl -Url $img.Groups["src"].Value -BaseUrl $Url
    }

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $entries) {
        $title = Normalize-Whitespace -Value $entry.Groups["title"].Value
        $dateText = Normalize-Whitespace -Value ($entry.Groups["date"].Value -replace '[^\d\.\-\s–]', '')
        $dateText = $dateText -replace '–', '-'
        $dateText = $dateText -replace '\s+', ' '
        $range = Parse-DateRange -Value $dateText -StartTimeText "" -EndTimeText ""
        if ($null -eq $range) { continue }

        $items.Add([pscustomobject]@{
            sourceKey    = $SourceKey
            sourceLabel  = $SourceLabel
            title        = $title
            genre        = Get-GenreFromText -Title $title -Keywords "muzeum" -ShortCode "" -Summary "" -DeclaredGenre "" -Venue "Muzeum Vysociny Trebic" -Link $Url
            municipality = "Trebic"
            venue        = "Muzeum Vysociny Trebic"
            startAt      = $range.startAt
            endAt        = $range.endAt
            startText    = $range.startAt.ToString("d. M. yyyy HH:mm")
            endText      = $range.endAt.ToString("d. M. yyyy HH:mm")
            dateLabel    = $dateText
            timeLabel    = ""
            summary      = ""
            keywords     = "muzeum"
            link         = if ($entry.Groups["href"].Success) { Resolve-AbsoluteUrl -Url $entry.Groups["href"].Value -BaseUrl $Url } else { $Url }
            detailLink   = if ($entry.Groups["href"].Success) { Resolve-AbsoluteUrl -Url $entry.Groups["href"].Value -BaseUrl $Url } else { $Url }
            imageUrl     = if ($imageMap.ContainsKey((Get-NormalizedPlaceKey -Value $title))) { $imageMap[(Get-NormalizedPlaceKey -Value $title)] } else { "" }
            dedupeKey    = "$($range.startAt.ToString("yyyy-MM-dd"))|$((Get-NormalizedPlaceKey -Value $title))"
        })
    }

    $items
}

function Get-RoxyProgramItems {
    param(
        [string]$SourceKey,
        [string]$SourceLabel,
        [string]$Url
    )

    $html = Get-StringContent -Url $Url
    $pattern = '<div class="programe-grid-box[^"]*"[^>]*style="background-image:url\((?<image>[^)]+)\);">.*?<h3[^>]*>(?<title>.*?)</h3>.*?<div class="program-date[^"]*">(?<date>.*?)</div>.*?<a href="(?<href>https://www\.roxy-club\.cz/akce/[^"]+)" class="button programe-grid--more">'
    $matches = [regex]::Matches($html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    if ($matches.Count -eq 0) {
        $pattern = '<div class="programe-grid-box[^"]*"[^>]*style="background-image:url\((?<image>[^)]+)\);">.*?<div class="program-box-content.*?<h3[^>]*>(?<title>.*?)</h3>.*?<span[^>]*>(?<date>\d{1,2}\.\d{1,2}\.\d{4})</span>.*?<a href="(?<href>https://www\.roxy-club\.cz/akce/[^"]+)" class="button programe-grid--more">'
        $matches = [regex]::Matches($html, $pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
    }

    $fallbackMatches = [regex]::Matches($html, '<a href="(?<href>https://www\.roxy-club\.cz/akce/[^"]+)" class="d-flex flex-column slider-nav-image[^"]*" style="background-image:url\(''?(?<image>[^'')]+)''?\);">\s*<div class="slider-nav-title w-100">(?<title>.*?)</div>\s*<div class="slider-nav-date w-100">(?<date>\d{1,2}\.\d{1,2}\.\d{4})</div>', [Text.RegularExpressions.RegexOptions]::Singleline)

    $items = New-Object System.Collections.Generic.List[object]
    $sourceItems = if ($fallbackMatches.Count -gt $matches.Count) { $fallbackMatches } else { $matches }
    foreach ($match in $sourceItems) {
        $dateText = Normalize-Whitespace -Value $match.Groups["date"].Value
        $range = Parse-DateRange -Value $dateText -StartTimeText "" -EndTimeText ""
        if ($null -eq $range) { continue }
        $title = Normalize-Whitespace -Value $match.Groups["title"].Value

        $items.Add([pscustomobject]@{
            sourceKey    = $SourceKey
            sourceLabel  = $SourceLabel
            title        = $title
            genre        = "Koncerty"
            municipality = "Trebic"
            venue        = "ROXY Club Trebic"
            startAt      = $range.startAt
            endAt        = $range.endAt
            startText    = $range.startAt.ToString("d. M. yyyy HH:mm")
            endText      = $range.endAt.ToString("d. M. yyyy HH:mm")
            dateLabel    = $dateText
            timeLabel    = ""
            summary      = ""
            keywords     = "club koncert party"
            link         = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
            detailLink   = Resolve-AbsoluteUrl -Url $match.Groups["href"].Value -BaseUrl $Url
            imageUrl     = Resolve-AbsoluteUrl -Url $match.Groups["image"].Value.Trim("'") -BaseUrl $Url
            dedupeKey    = "$($range.startAt.ToString("yyyy-MM-dd"))|$((Get-NormalizedPlaceKey -Value $title))"
        })
    }

    $items
}

function Merge-EventItems {
    param([object[]]$Items)

    $itemList = [object[]]$Items
    $map = @{}
    foreach ($item in $itemList) {
        if ($null -eq $item) { continue }
        if (-not $map.ContainsKey($item.dedupeKey)) {
            $map[$item.dedupeKey] = $item
            continue
        }

        $existing = $map[$item.dedupeKey]
        $regionItem = $null
        $otherItem = $null
        if ($existing.sourceKey -eq "region_calendar" -and $item.sourceKey -ne "region_calendar") {
            $regionItem = $existing
            $otherItem = $item
        } elseif ($item.sourceKey -eq "region_calendar" -and $existing.sourceKey -ne "region_calendar") {
            $regionItem = $item
            $otherItem = $existing
        }

        if ($null -ne $regionItem -and $null -ne $otherItem) {
            $mergedStartAt = if ($otherItem.startAt.TimeOfDay.TotalMinutes -gt 0) { $otherItem.startAt } else { $regionItem.startAt }
            $mergedEndAt = if ($otherItem.endAt.TimeOfDay.TotalMinutes -gt 0) { $otherItem.endAt } else { $regionItem.endAt }
            $mergedImageUrl = if (-not [string]::IsNullOrWhiteSpace($regionItem.imageUrl)) { $regionItem.imageUrl } else { $otherItem.imageUrl }
            $map[$item.dedupeKey] = [pscustomobject]@{
                sourceKey    = $otherItem.sourceKey
                sourceLabel  = $otherItem.sourceLabel
                title        = $otherItem.title
                genre        = $otherItem.genre
                municipality = $regionItem.municipality
                venue        = if (-not [string]::IsNullOrWhiteSpace($otherItem.venue)) { $otherItem.venue } else { $regionItem.venue }
                startAt      = $mergedStartAt
                endAt        = $mergedEndAt
                startText    = $mergedStartAt.ToString("d. M. yyyy HH:mm")
                endText      = $mergedEndAt.ToString("d. M. yyyy HH:mm")
                dateLabel    = $regionItem.dateLabel
                timeLabel    = $otherItem.timeLabel
                summary      = if (-not [string]::IsNullOrWhiteSpace($regionItem.summary)) { $regionItem.summary } else { $otherItem.summary }
                keywords     = $regionItem.keywords
                link         = $otherItem.link
                detailLink   = $regionItem.detailLink
                imageUrl     = $mergedImageUrl
                dedupeKey    = $item.dedupeKey
            }
        }
    }

    @($map.Values)
}

function Get-PreviewText {
    param(
        [string]$Text,
        [int]$MaxLength = 220
    )

    $clean = Normalize-Whitespace -Value $Text
    if ([string]::IsNullOrWhiteSpace($clean)) { return "" }
    if ($clean.Length -le $MaxLength) { return $clean }

    $cut = $clean.Substring(0, $MaxLength)
    $lastSpace = $cut.LastIndexOf(' ')
    if ($lastSpace -ge 120) {
        $cut = $cut.Substring(0, $lastSpace)
    }
    "$($cut.Trim())..."
}

function Get-DateLabelWithOptionalTime {
    param([datetime]$Value)

    if ($Value.TimeOfDay.TotalMinutes -gt 0) {
        return $Value.ToString("d. M. yyyy HH:mm")
    }

    $Value.ToString("d. M. yyyy")
}

function Convert-ItemsToHtml {
    param(
        [object[]]$Items,
        [string]$GeneratedAtText,
        [int]$HorizonDays,
        [double]$RadiusKm,
        [string[]]$SourceLabels,
        [datetime]$Now
    )

    $css = @"
:root{--bg:#f7f2e9;--card:#fffdfa;--ink:#1d2a34;--muted:#66727f;--accent:#0d7a70;--accent-2:#d55a1f;--line:#eadfce;--hero-a:#0d7a70;--hero-b:#cf5a22}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:'Segoe UI',Arial,sans-serif;color:var(--ink);background:linear-gradient(180deg,#f7f1e7 0%,#fbf8f2 100%)}
.wrap{max-width:1120px;margin:0 auto;padding:14px 14px 40px}
.hero{padding:20px 18px;border-radius:22px;background:linear-gradient(135deg,var(--hero-a),var(--hero-b));color:#fff;box-shadow:0 14px 34px rgba(29,42,52,.14)}
.hero h1{margin:0 0 8px;font-size:28px;line-height:1.08}
.hero p{margin:0;font-size:15px;line-height:1.45;max-width:980px}
.meta{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
.pill{padding:7px 11px;border-radius:999px;background:rgba(255,255,255,.18);font-size:12px;font-weight:700}
.next-box{margin-top:16px;padding:14px 15px;border-radius:18px;background:rgba(255,255,255,.14);backdrop-filter:blur(2px)}
.next-box h2{margin:0 0 10px;font-size:18px}
.next-list{margin:0;padding:0;list-style:none}
.next-list li{display:grid;grid-template-columns:140px 1fr;gap:10px;padding:4px 0;font-size:14px;line-height:1.35}
.next-date{font-weight:700}
.next-link{color:#fff;text-decoration:none}
.next-link:hover{text-decoration:underline}
.genre-section{margin-top:22px}
.genre-header{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:10px}
.genre-header h2{margin:0;font-size:22px}
.genre-header span{color:var(--muted);font-size:13px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:14px}
.card{background:var(--card);border:1px solid var(--line);border-radius:20px;padding:14px;box-shadow:0 8px 20px rgba(29,42,52,.05)}
.thumb{display:block;width:100%;aspect-ratio:16/9;object-fit:cover;border-radius:14px;margin:0 0 12px;background:#f0e7db}
.date{font-size:13px;font-weight:800;letter-spacing:.03em;color:var(--accent)}
.card h3{margin:8px 0 10px;font-size:18px;line-height:1.15}
.summary{color:var(--muted);font-size:14px;line-height:1.45;margin:10px 0 0;max-height:8.7em;overflow:hidden}
.summary.is-expanded{max-height:none;overflow:visible}
.summary-toggle{margin-top:8px;padding:0;border:0;background:none;color:var(--accent);font-size:13px;font-weight:800;cursor:pointer}
.summary-toggle:hover{text-decoration:underline}
.detail{margin-top:10px;color:var(--ink);font-size:14px;line-height:1.35}
.detail strong{color:var(--accent-2)}
.links{margin-top:14px;display:flex}
.btn{display:inline-block;width:100%;text-align:center;padding:12px 14px;border-radius:12px;background:var(--accent);color:#fff;text-decoration:none;font-weight:700}
.empty{margin-top:20px;padding:20px;border-radius:18px;background:#fff;border:1px solid var(--line);color:var(--muted)}
@media (min-width:721px){.wrap{padding:24px 18px 56px}.hero{padding:28px 26px}.hero h1{font-size:34px}.card h3{font-size:20px}}
@media (max-width:720px){.next-list li{grid-template-columns:1fr}.next-date{margin-bottom:1px}.hero p{max-width:none}}
"@

    $itemList = [object[]]$Items
    $nextItemsHtml = if ($itemList.Count -gt 0) {
        $nearestItems = $itemList | Where-Object { $_.startAt -ge $Now } | Sort-Object startAt, endAt, title | Select-Object -First 10
        $rows = ($nearestItems | ForEach-Object {
            "<li><span class='next-date'>$([System.Net.WebUtility]::HtmlEncode((Get-DateLabelWithOptionalTime -Value $_.startAt)))</span><a class='next-link' href='$([System.Net.WebUtility]::HtmlEncode($_.detailLink))' target='_blank' rel='noreferrer'>$([System.Net.WebUtility]::HtmlEncode($_.title))</a></li>"
        }) -join "`n"
        if ([string]::IsNullOrWhiteSpace($rows)) {
            "<div class='next-box'><h2>10 nejbli&#382;&#353;&#237;ch akc&#237;</h2><div>V nejblizsich dnech ted nejsou zadne budouci akce.</div></div>"
        } else {
            "<div class='next-box'><h2>10 nejbli&#382;&#353;&#237;ch akc&#237;</h2><ul class='next-list'>$rows</ul></div>"
        }
    } else {
        ""
    }
    $sourceText = [System.Net.WebUtility]::HtmlEncode(($SourceLabels | Sort-Object -Unique) -join " + ")
    if ($itemList.Count -eq 0) {
        $bodyHtml = "<div class='empty'>V zadanem okruhu a horizontu $HorizonDays dni ted nejsou zadne aktivni kulturni akce.</div>"
    } else {
        $bodyHtml = (($itemList | Group-Object genre | Sort-Object Name) | ForEach-Object {
            $cards = ($_.Group | Sort-Object sortAt, endAt, title | ForEach-Object {
                $fullSummary = Normalize-Whitespace -Value $_.summary
                $previewSummary = Get-PreviewText -Text $fullSummary -MaxLength 220
                $summaryId = "summary-$([Math]::Abs($_.dedupeKey.GetHashCode()))"
                $summary = if ([string]::IsNullOrWhiteSpace($previewSummary)) {
                    ""
                } elseif ($fullSummary.Length -gt $previewSummary.Length) {
                    "<div id='$summaryId' class='summary' data-full='$([System.Net.WebUtility]::HtmlEncode($fullSummary))' data-preview='$([System.Net.WebUtility]::HtmlEncode($previewSummary))'>$([System.Net.WebUtility]::HtmlEncode($previewSummary))</div><button class='summary-toggle' type='button' data-summary-id='$summaryId' aria-expanded='false'>Zobrazit vice</button>"
                } else {
                    "<div class='summary'>$([System.Net.WebUtility]::HtmlEncode($previewSummary))</div>"
                }
                $image = if ([string]::IsNullOrWhiteSpace($_.imageUrl)) { "" } else { "<img class='thumb' src='$([System.Net.WebUtility]::HtmlEncode($_.imageUrl))' alt='$([System.Net.WebUtility]::HtmlEncode($_.title))'>" }
                $distanceDetail = if ([double]$_.distanceKm -gt 0.04) { "<div class='detail'><strong>Vzd&#225;lenost:</strong> $([System.Net.WebUtility]::HtmlEncode(('{0:N1} km' -f $_.distanceKm)))</div>" } else { "" }
                "<article class='card'>$image<div class='date'>$([System.Net.WebUtility]::HtmlEncode($_.startText))</div><h3>$([System.Net.WebUtility]::HtmlEncode($_.title))</h3><div class='detail'><strong>M&#237;sto:</strong> $([System.Net.WebUtility]::HtmlEncode($_.venue))</div><div class='detail'><strong>Obec:</strong> $([System.Net.WebUtility]::HtmlEncode($_.municipality))</div>$distanceDetail<div class='detail'><strong>Term&#237;n:</strong> $([System.Net.WebUtility]::HtmlEncode($_.dateLabel))</div>$summary<div class='links'><a class='btn' href='$([System.Net.WebUtility]::HtmlEncode($_.detailLink))' target='_blank' rel='noreferrer'>Otev&#345;&#237;t detail</a></div></article>"
            }) -join "`n"
            "<section class='genre-section'><div class='genre-header'><h2>$(Get-GenreDisplayHtml -Genre $_.Name)</h2><span>$($_.Count) akc&#237;</span></div><div class='grid'>$cards</div></section>"
        }) -join "`n"
    }

    $script = @"
<script>
document.addEventListener('click', function (event) {
  var button = event.target.closest('.summary-toggle');
  if (!button) return;
  var summary = document.getElementById(button.getAttribute('data-summary-id'));
  if (!summary) return;
  var expanded = button.getAttribute('aria-expanded') === 'true';
  if (expanded) {
    summary.textContent = summary.getAttribute('data-preview');
    summary.classList.remove('is-expanded');
    button.setAttribute('aria-expanded', 'false');
    button.textContent = 'Zobrazit vice';
  } else {
    summary.textContent = summary.getAttribute('data-full');
    summary.classList.add('is-expanded');
    button.setAttribute('aria-expanded', 'true');
    button.textContent = 'Zobrazit mene';
  }
});
</script>
"@

    @"
<!DOCTYPE html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Kulturn&#237; akce v T&#345;eb&#237;&#269;i a okol&#237;</title>
  <style>$css</style>
</head>
<body>
  <div class="wrap">
    <section class="hero">
      <h1>Kulturn&#237; akce v T&#345;eb&#237;&#269;i a okol&#237; do $([int]$RadiusKm) km</h1>
      <p>Seznam je rozd&#283;len&#253; podle &#382;&#225;nru, &#345;azen&#253; od nejbli&#382;&#353;&#237; akce po nejpozd&#283;j&#353;&#237; a dr&#382;&#237; jen aktivn&#237; akce s horizontem $HorizonDays dn&#237; dop&#345;edu.</p>
      <div class="meta">
        <span class="pill">Aktualizov&#225;no: $([System.Net.WebUtility]::HtmlEncode($GeneratedAtText))</span>
        <span class="pill">Zdroje: $sourceText</span>
      </div>
      $nextItemsHtml
    </section>
    $bodyHtml
  </div>
  $script
</body>
</html>
"@
}

$configFullPath = Resolve-ProjectPath -Path $ConfigPath
$config = Read-JsonFile -Path $configFullPath
if ($null -eq $config) {
    throw "Nepodarilo se nacist konfiguraci '$configFullPath'."
}

$configDirectory = Split-Path -Parent $configFullPath
$reportPath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.reportPath
$itemsPath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.itemsPath
$runLogPath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.runLogPath
$locationCachePath = Resolve-PathFromBase -BasePath $configDirectory -RelativePath $config.output.locationCachePath

$knownPlaceMap = @{}
foreach ($property in $config.knownPlaces.PSObject.Properties) {
    $knownPlaceMap[(Get-NormalizedPlaceKey -Value $property.Name)] = [pscustomobject]@{
        latitude  = [double]$property.Value.latitude
        longitude = [double]$property.Value.longitude
        source    = "settings"
        display   = $property.Name
    }
}

$locationCache = Get-CacheMap -CacheObject (Read-JsonFile -Path $locationCachePath)
$timezoneInfo = Get-TimeZoneInfoSafe -PreferredId ([string]$config.timezone)
$now = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $timezoneInfo)
$windowEnd = $now.AddDays([int]$config.horizonDays)
$items = New-Object System.Collections.Generic.List[object]

$regionSource = $config.sources | Where-Object { $_.type -eq "region_calendar" } | Select-Object -First 1
if ($null -ne $regionSource) {
    foreach ($listItem in @(Get-RegionCalendarListItems -ListUrl $regionSource.listUrl)) {
        $roughRange = Parse-DateRange -Value $listItem.dateText -StartTimeText "" -EndTimeText ""
        if ($null -eq $roughRange) { continue }
        if ($roughRange.endAt -lt $now -or $roughRange.startAt -gt $windowEnd) { continue }
        $detailItem = Get-RegionCalendarDetailItem -ListItem $listItem
        if ($null -ne $detailItem) { $items.Add($detailItem) }
    }
}

foreach ($source in @($config.sources | Where-Object { $_.type -eq "mks_category" })) {
    foreach ($item in @(Get-MksCategoryItems -SourceKey $source.key -SourceLabel $source.label -Url $source.url -Genre $source.genre)) {
        $items.Add($item)
    }
}

foreach ($source in @($config.sources | Where-Object { $_.type -eq "unesco_articles" })) {
    foreach ($item in @(Get-UnescoItems -SourceKey $source.key -SourceLabel $source.label -Url $source.url)) {
        $items.Add($item)
    }
}

foreach ($source in @($config.sources | Where-Object { $_.type -eq "ddm_catalog" })) {
    foreach ($item in @(Get-DdmItems -SourceKey $source.key -SourceLabel $source.label -Url $source.url)) {
        $items.Add($item)
    }
}

foreach ($source in @($config.sources | Where-Object { $_.type -eq "muzeum_schedule" })) {
    foreach ($item in @(Get-MuzeumScheduleItems -SourceKey $source.key -SourceLabel $source.label -Url $source.url)) {
        $items.Add($item)
    }
}

foreach ($source in @($config.sources | Where-Object { $_.type -eq "roxy_program" })) {
    foreach ($item in @(Get-RoxyProgramItems -SourceKey $source.key -SourceLabel $source.label -Url $source.url)) {
        $items.Add($item)
    }
}

$mergedItems = Merge-EventItems -Items $items
$sortedMergedItems = [object[]]($mergedItems | Sort-Object startAt, endAt, title)
$finalItems = New-Object System.Collections.Generic.List[object]
foreach ($item in $sortedMergedItems) {
    if ($item.endAt -lt $now -or $item.startAt -gt $windowEnd) { continue }
    if ($item.genre -eq "Kino" -and $item.sourceKey -ne "mks_kino") { continue }
    $coordinates = Resolve-PlaceCoordinates -PlaceName $item.municipality -KnownPlaceMap $knownPlaceMap -LocationCache $locationCache
    if ($null -eq $coordinates) { continue }

    $distance = Get-HaversineDistanceKm -Latitude1 ([double]$config.center.latitude) -Longitude1 ([double]$config.center.longitude) -Latitude2 ([double]$coordinates.latitude) -Longitude2 ([double]$coordinates.longitude)
    if ($distance -gt [double]$config.radiusKm) { continue }
    $sortAt = if ($item.startAt -lt $now) { $now } else { $item.startAt }

    $finalItems.Add([pscustomobject]@{
        sourceKey    = $item.sourceKey
        sourceLabel  = $item.sourceLabel
        title        = $item.title
        genre        = $item.genre
        municipality = $item.municipality
        venue        = $item.venue
        startAt      = $item.startAt
        endAt        = $item.endAt
        startText    = $item.startText
        endText      = $item.endText
        dateLabel    = $item.dateLabel
        timeLabel    = $item.timeLabel
        summary      = $item.summary
        keywords     = $item.keywords
        link         = $item.link
        detailLink   = $item.detailLink
        imageUrl     = $item.imageUrl
        distanceKm   = [math]::Round($distance, 1)
        sortAt       = $sortAt
        dedupeKey    = $item.dedupeKey
    })
}

$finalItems = [System.Collections.Generic.List[object]]([object[]]($finalItems | Sort-Object sortAt, endAt, title))
$generatedAtText = $now.ToString("d. M. yyyy HH:mm")
$sourceLabels = [object[]]($config.sources | ForEach-Object { [string]$_.label })
$reportHtml = Convert-ItemsToHtml -Items $finalItems -GeneratedAtText $generatedAtText -HorizonDays ([int]$config.horizonDays) -RadiusKm ([double]$config.radiusKm) -SourceLabels $sourceLabels -Now $now

Ensure-Directory -Path (Split-Path -Parent $reportPath)
$reportHtml | Set-Content -LiteralPath $reportPath -Encoding UTF8

$itemsPayload = [pscustomobject]@{
    generatedAt = $now.ToString("o")
    center      = $config.center
    radiusKm    = [double]$config.radiusKm
    horizonDays = [int]$config.horizonDays
    items       = @(([object[]]$finalItems) | ForEach-Object {
        [pscustomobject]@{
            title        = $_.title
            genre        = $_.genre
            municipality = $_.municipality
            venue        = $_.venue
            startAt      = $_.startAt.ToString("o")
            endAt        = $_.endAt.ToString("o")
            dateLabel    = $_.dateLabel
            summary      = $_.summary
            distanceKm   = $_.distanceKm
            link         = $_.link
            detailLink   = $_.detailLink
            imageUrl     = $_.imageUrl
            sourceLabel  = $_.sourceLabel
        }
    })
}
Write-JsonFile -Path $itemsPath -Value $itemsPayload
Save-LocationCache -Path $locationCachePath -Cache $locationCache

$runLog = [pscustomobject]@{
    generatedAt = $now.ToString("o")
    totalItems  = $finalItems.Count
    reportPath  = $reportPath
    itemsPath   = $itemsPath
}
Write-JsonFile -Path $runLogPath -Value $runLog

Write-Host "Monitoring kulturnich akci hotov."
Write-Host "Akci v okruhu $($config.radiusKm) km a v horizontu $($config.horizonDays) dni: $($finalItems.Count)"
Write-Host "Report: $reportPath"
