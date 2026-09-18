param(
    [switch]$excludeListed,
    [switch]$download,
    [switch]$fix,

    [ValidateRange(0, 99)]
    [int]$age = 0,

    [ValidateSet("Id", "Size")]
    [string]$sort = "Id",

    [switch]$cleanup,

    # Optional arguments for -download and -cleanup.
    # Allowed values:
    #   download: Ask | All
    #   cleanup : Ignored | All
    #
    # ValueFromRemainingArguments enables the following syntax:
    #   -download
    #   -download All
    #   -cleanup
    #   -cleanup All
    [Parameter(ValueFromRemainingArguments = $true)]
    [ValidateSet("Ask", "All", "Ignored")]
    [string[]]$ModeArguments = @()
)

$DownloadEnabled = [bool]$download
$CleanupEnabled  = [bool]$cleanup

# Default modes
$DownloadMode = "Ask"
$CleanupMode  = "Ignored"

# PowerShell switch parameters cannot accept an optional string argument
# themselves. The following unbound mode words are collected in
# $ModeArguments. The original invocation line is used to determine which
# switch each word belongs to.
$invocationLine = [string]$MyInvocation.Line

if ($DownloadEnabled) {
    $downloadMatch =
        [regex]::Match(
            $invocationLine,
            '(?i)(?:^|\s)-download(?:\s+["'']?(?<mode>Ask|All)["'']?)?(?=\s|$)'
        )

    if (
        $downloadMatch.Success -and
        $downloadMatch.Groups["mode"].Success
    ) {
        $DownloadMode =
            $downloadMatch.Groups["mode"].Value
    }
}

if ($CleanupEnabled) {
    $cleanupMatch =
        [regex]::Match(
            $invocationLine,
            '(?i)(?:^|\s)-cleanup(?:\s+["'']?(?<mode>Ignored|All)["'']?)?(?=\s|$)'
        )

    if (
        $cleanupMatch.Success -and
        $cleanupMatch.Groups["mode"].Success
    ) {
        $CleanupMode =
            $cleanupMatch.Groups["mode"].Value
    }
}

# Fallback for invocation methods where $MyInvocation.Line does not contain the expanded
# parameter value (for example, certain programmatic invocations).
# If only one of the two modes is active, a single remaining argument can be
# assigned unambiguously.
if ($ModeArguments.Count -gt 0) {
    if ($DownloadEnabled -and -not $CleanupEnabled) {
        if ($ModeArguments.Count -gt 1) {
            throw "Zu viele Argumente fuer -download."
        }

        if ($ModeArguments[0] -notin @("Ask", "All")) {
            throw (
                "Ungueltiger Wert fuer -download: '{0}'. Erlaubt: Ask, All." -f
                $ModeArguments[0]
            )
        }

        $DownloadMode =
            [string]$ModeArguments[0]
    }
    elseif ($CleanupEnabled -and -not $DownloadEnabled) {
        if ($ModeArguments.Count -gt 1) {
            throw "Zu viele Argumente fuer -cleanup."
        }

        if ($ModeArguments[0] -notin @("Ignored", "All")) {
            throw (
                "Ungueltiger Wert fuer -cleanup: '{0}'. Erlaubt: Ignored, All." -f
                $ModeArguments[0]
            )
        }

        $CleanupMode =
            [string]$ModeArguments[0]
    }
    elseif ($DownloadEnabled -and $CleanupEnabled) {
        # For combined invocations, assignment primarily relies on
        # $MyInvocation.Line. Also make sure that no values invalid for both
        # modes are present.
        foreach ($modeArgument in $ModeArguments) {
            if ($modeArgument -notin @("Ask", "All", "Ignored")) {
                throw (
                    "Ungueltiges Modusargument: '{0}'." -f
                    $modeArgument
                )
            }
        }
    }
    else {
        throw (
            "Modusargument(e) '{0}' wurden angegeben, aber weder -download " +
            "noch -cleanup ist aktiv." -f
            ($ModeArguments -join ", ")
        )
    }
}

# Strictly validate the effective values once more.
if ($DownloadMode -notin @("Ask", "All")) {
    throw (
        "Ungueltiger Download-Modus '{0}'. Erlaubt: Ask, All." -f
        $DownloadMode
    )
}

if ($CleanupMode -notin @("Ignored", "All")) {
    throw (
        "Ungueltiger Cleanup-Modus '{0}'. Erlaubt: Ignored, All." -f
        $CleanupMode
    )
}

# Remember whether -age was actually specified.
# This is more robust than Nullable[int], because PowerShell can treat nullable values
# as regular Int32 values during parameter binding.
$AgeFilterEnabled = $PSBoundParameters.ContainsKey("age")

$ErrorActionPreference = "Stop"

# ============================================================================
# Configuration
# ============================================================================

$CatalogUrl = "https://ttapiv2.ravensburger.com/api/v2/catalog/de_DE"

# Official Ravensburger list of tiptoi audio files.
# It serves as a bridge between the book number derived from the ISBN and older
# Ravensburger article numbers in the tiptoi catalog.
$RavensburgerAudioIndex =
    "https://service.ravensburger.de/tiptoi%C2%AE/tiptoi%C2%AE_Audiodateien"

# RSS feed: tiptoi books held by Treffpunkt Buecherei
# Uhldingen-Muehlhofen. The current loan status is irrelevant.
$LibraryRssUrl =
    "https://opac.winbiap.net/uhldingen-muehlhofen/service/rss.aspx?data=Y21kPTUmYW1wO3NDPWNfMD0wJSVtXzA9MSUlZl8wPTIlJW9fMD04JSV2XzA9dGlwdG9pJSVnXzA9LTErK2NfMT0xJSVtXzE9MSUlZl8xPTQ1JSVvXzE9MiUldl8xPTYlJWdfMT0tMQ%3d%3d"

$ScriptDirectory = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($ScriptDirectory)) {
    $ScriptDirectory = (Get-Location).Path
}

# Optional lists of owned and ignored tiptoi books.
# Example structure:
#
# mine:
#  - isbn: 978-3-473-32909-0
#  - id: 923
# ignore:
#  - isbn: 978-3-473-32911-0
#  - id: 924
#
$MyBooksFile =
    Join-Path `
        $ScriptDirectory `
        "tiptoi_mybooks.yml"

# For Windows PowerShell 5.1 / legacy TLS configurations
try {
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor
        [Net.SecurityProtocolType]::Tls12
}
catch {
}

try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $OutputEncoding = [Console]::OutputEncoding
}
catch {
}


# ============================================================================
# General helper functions
# ============================================================================

function ConvertFrom-HtmlText {
    param(
        [string]$Html
    )

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $text = [regex]::Replace($Html, "<[^>]+>", " ")
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    $text = [regex]::Replace($text, "\s+", " ")

    return $text.Trim()
}


function ConvertTo-TitleKey {
    param(
        [string]$Title
    )

    if ([string]::IsNullOrWhiteSpace($Title)) {
        return ""
    }

    $text = [System.Net.WebUtility]::HtmlDecode($Title)

    $text = $text -replace '(?i)\btiptoi\b', ' '
    $text = $text -replace '(?i)\baudiodatei\b', ' '
    $text = $text -replace '(?i)wieso\s*\?\s*weshalb\s*\?\s*warum\s*\?', ' '

    $text = $text.Normalize(
        [System.Text.NormalizationForm]::FormD
    )

    $builder = New-Object System.Text.StringBuilder

    foreach ($character in $text.ToCharArray()) {
        $category =
            [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory(
                $character
            )

        if (
            $category -ne
            [System.Globalization.UnicodeCategory]::NonSpacingMark
        ) {
            [void]$builder.Append($character)
        }
    }

    $text = $builder.ToString().ToLowerInvariant()
    $text = $text -replace '[^a-z0-9]+', ' '
    $text = [regex]::Replace($text, '\s+', ' ')

    return $text.Trim()
}


function Get-TitleSimilarity {
    param(
        [string]$Title1,
        [string]$Title2
    )

    $a = ConvertTo-TitleKey $Title1
    $b = ConvertTo-TitleKey $Title2

    if (
        [string]::IsNullOrWhiteSpace($a) -or
        [string]::IsNullOrWhiteSpace($b)
    ) {
        return 0.0
    }

    if ($a -eq $b) {
        return 1.0
    }

    if ($a.Contains($b) -or $b.Contains($a)) {
        if (
            ($a.Split(" ").Count -ge 3) -and
            ($b.Split(" ").Count -ge 3)
        ) {
            return 0.95
        }
    }

    $stopWords = @(
        "der", "die", "das", "den", "dem", "des",
        "ein", "eine", "einer", "eines",
        "und", "oder", "mit", "von", "fur",
        "wieso", "weshalb", "warum", "tiptoi"
    )

    $tokensA = @(
        $a.Split(" ") |
        Where-Object {
            $_.Length -gt 1 -and
            $_ -notin $stopWords
        } |
        Select-Object -Unique
    )

    $tokensB = @(
        $b.Split(" ") |
        Where-Object {
            $_.Length -gt 1 -and
            $_ -notin $stopWords
        } |
        Select-Object -Unique
    )

    if (
        ($tokensA.Count -eq 0) -or
        ($tokensB.Count -eq 0)
    ) {
        return 0.0
    }

    $intersection = @(
        $tokensA |
        Where-Object {
            $_ -in $tokensB
        }
    ).Count

    return (
        2.0 * $intersection /
        ($tokensA.Count + $tokensB.Count)
    )
}


function Format-ByteSize {
    param(
        [long]$Bytes
    )

    if ($Bytes -ge 1GB) {
        return ("{0:N2} GB" -f ($Bytes / 1GB))
    }

    if ($Bytes -ge 1MB) {
        return ("{0:N2} MB" -f ($Bytes / 1MB))
    }

    if ($Bytes -ge 1KB) {
        return ("{0:N2} KB" -f ($Bytes / 1KB))
    }

    return "$Bytes Byte"
}


# ============================================================================
# ISBN
# ============================================================================

function Test-Isbn13 {
    param(
        [string]$Isbn
    )

    # WinBIAP labels the field as ISBN, but for older
    # Ravensburger products it sometimes contains the EAN-13 4005556...
    # The check digit is calculated identically for ISBN-13 and EAN-13.
    $digits = $Isbn -replace '\D', ''

    if ($digits.Length -ne 13) {
        return $false
    }

    $sum = 0

    for ($i = 0; $i -lt 12; $i++) {
        $digit = [int]::Parse($digits.Substring($i, 1))

        if (($i % 2) -eq 0) {
            $sum += $digit
        }
        else {
            $sum += 3 * $digit
        }
    }

    $expected = (10 - ($sum % 10)) % 10
    $actual = [int]::Parse($digits.Substring(12, 1))

    return ($expected -eq $actual)
}


function Get-Isbn13FromText {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    # Accepts, for example:
    # 978-3-473-32924-3
    # 9783473329243
    # 400-5-556-00590-1
    # 4005556005901
    #
    # For older Ravensburger products, WinBIAP stores an EAN-13
    # in the "ISBN" field. For collection matching, this identifier is
    # just as useful as a 978/979 ISBN.
    $pattern =
        '(?<!\d)((?:97[89]|4005556)(?:[\s\-\u2010-\u2015]*\d){10})(?!\d)'

    $result = @()

    foreach ($match in [regex]::Matches($Text, $pattern)) {
        $isbn = $match.Groups[1].Value -replace '\D', ''

        if (
            (Test-Isbn13 -Isbn $isbn) -and
            ($isbn -notin $result)
        ) {
            $result += $isbn
        }
    }

    return $result
}


function Get-RavensburgerBookNumberFromIsbn13 {
    param(
        [string]$Isbn
    )

    $digits = $Isbn -replace '\D', ''

    if ($digits.Length -ne 13) {
        return $null
    }

    # Ravensburger book ISBN:
    # 978-3-473-49286-2 -> 49286
    if ($digits.StartsWith("9783473")) {
        return $digits.Substring(7, 5)
    }

    # Older Ravensburger retail EAN, which WinBIAP sometimes
    # returns in the "ISBN" field:
    # 400-5-556-00590-1 -> 00590
    if ($digits.StartsWith("4005556")) {
        return $digits.Substring(7, 5)
    }

    return $null
}


# ============================================================================
# GME
# ============================================================================

function Get-GmeProductId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::ReadWrite
    )

    try {
        $header = New-Object byte[] 24
        $offset = 0

        while ($offset -lt $header.Length) {
            $read = $stream.Read(
                $header,
                $offset,
                $header.Length - $offset
            )

            if ($read -eq 0) {
                throw "Datei ist zu kurz, um eine gueltige GME-Datei zu sein."
            }

            $offset += $read
        }
    }
    finally {
        $stream.Dispose()
    }

    $magic = [BitConverter]::ToUInt32($header, 0x08)

    if ($magic -ne 0x238B) {
        throw (
            "Ungueltige GME Magic Number: 0x{0:X8}" -f $magic
        )
    }

    return [string](
        [BitConverter]::ToUInt32($header, 0x14)
    )
}


# ============================================================================
# Library RSS feed
# ============================================================================

function Get-LibraryRssItems {
    Write-Host ""
    Write-Host "RSS-Feed der Buecherei wird geladen ..."

    try {
        $response = Invoke-WebRequest `
            -Uri $LibraryRssUrl `
            -Method Get `
            -UseBasicParsing `
            -TimeoutSec 30

        # The WinBIAP RSS feed currently starts with a UTF-8 BOM
        # (U+FEFF). Windows PowerShell 5.1 sometimes passes this character
        # through as part of the string; a direct [xml] cast
        # then fails with "Cannot convert value \"<BOM><rss...\"".
        $rssText = [string]$response.Content

        # Remove the actual BOM/zero-width character as well as any BOM that may
        # already have been decoded incorrectly.
        $rssText = [regex]::Replace(
            $rssText,
            '^[\uFEFF\u200B\x00\s]+',
            ''
        )

        $rssText = [regex]::Replace(
            $rssText,
            '^ï»¿',
            ''
        )

        $xmlDocument = New-Object System.Xml.XmlDocument
        $xmlDocument.PreserveWhitespace = $false
        $xmlDocument.LoadXml($rssText)

        $rss = $xmlDocument
    }
    catch {
        throw (
            "RSS-Feed der Buecherei konnte nicht geladen/geparst werden: " +
            $_.Exception.Message
        )
    }

    $nodes = @(
        $rss.SelectNodes(
            "//*[local-name()='item']"
        )
    )

    $items = @()

    foreach ($node in $nodes) {
        $titleNode = $node.SelectSingleNode(
            "./*[local-name()='title']"
        )

        $linkNode = $node.SelectSingleNode(
            "./*[local-name()='link']"
        )

        $title = ""

        if ($null -ne $titleNode) {
            $title = [string]$titleNode.InnerText
        }

        $link = ""

        if ($null -ne $linkNode) {
            $link = [string]$linkNode.InnerText
        }

        # InnerText also contains the description / bibliographic information.
        $allText = [string]$node.InnerText

        $isbns = @(
            Get-Isbn13FromText -Text $allText
        )

        if ($isbns.Count -eq 0) {
            # Additional attempt on the XML itself in case the ISBN is embedded differently
            # in HTML fragments / CDATA.
            $isbns = @(
                Get-Isbn13FromText -Text $node.OuterXml
            )
        }

        $items += [PSCustomObject]@{
            Title = $title
            Link  = $link
            Isbns = $isbns
        }
    }

    $withIsbn = @(
        $items |
        Where-Object {
            $_.Isbns.Count -gt 0
        }
    )

    Write-Host (
        "{0} RSS-Eintraege gefunden, davon {1} mit ISBN/EAN." -f
        $items.Count,
        $withIsbn.Count
    )

    return $items
}


# ============================================================================
# Official Ravensburger book numbers / titles
# ============================================================================

function Get-RavensburgerBookIndex {
    Write-Host ""
    Write-Host "Ravensburger-Buchindex wird geladen ..."

    try {
        $response = Invoke-WebRequest `
            -Uri $RavensburgerAudioIndex `
            -Method Get `
            -UseBasicParsing `
            -TimeoutSec 30

        $html = $response.Content
    }
    catch {
        throw (
            "Ravensburger-Buchindex konnte nicht geladen werden: " +
            $_.Exception.Message
        )
    }

    $byNumber = @{}

    $anchors = [regex]::Matches(
        $html,
        '(?is)<a\b[^>]*href\s*=\s*["''](?<href>[^"'']+)["''][^>]*>(?<text>.*?)</a>'
    )

    foreach ($anchor in $anchors) {
        $href = [System.Net.WebUtility]::HtmlDecode(
            $anchor.Groups["href"].Value
        )

        try {
            $href = [System.Uri]::UnescapeDataString($href)
        }
        catch {
        }

        # Only the "Audiodateien tiptoi Buecher" section.
        $hrefKey = ConvertTo-TitleKey -Title $href

        if ($hrefKey -notmatch 'audiodateien.*bucher') {
            continue
        }

        $anchorText = ConvertFrom-HtmlText `
            -Html $anchor.Groups["text"].Value

        if ([string]::IsNullOrWhiteSpace($anchorText)) {
            continue
        }

        $numberMatches = [regex]::Matches(
            $anchorText,
            '(?<!\d)(\d{5})(?!\d)'
        )

        if ($numberMatches.Count -eq 0) {
            continue
        }

        $titleText = $anchorText -replace '(?<!\d)\d{5}(?!\d)', ' '
        $titleText = $titleText -replace '(?i)\s+\bund\b\s*$', ''
        $titleText = [regex]::Replace($titleText, '\s+', ' ').Trim()

        foreach ($numberMatch in $numberMatches) {
            $number = $numberMatch.Groups[1].Value

            if (-not $byNumber.ContainsKey($number)) {
                $byNumber[$number] = $titleText
            }
        }
    }

    Write-Host (
        "{0} Ravensburger-Buchnummern gefunden." -f
        $byNumber.Count
    )

    return $byNumber
}


# ============================================================================
# Ravensburger tiptoi catalog
# ============================================================================

function Get-RavensburgerCatalog {
    Write-Host ""
    Write-Host "Ravensburger-tiptoi-Katalog wird geladen ..."

    try {
        return Invoke-RestMethod `
            -Uri $CatalogUrl `
            -Method Get `
            -TimeoutSec 30
    }
    catch {
        throw (
            "Ravensburger-tiptoi-Katalog konnte nicht geladen werden: " +
            $_.Exception.Message
        )
    }
}


function Get-CatalogProductMaps {
    param(
        [Parameter(Mandatory = $true)]
        $Catalog
    )

    $products = @($Catalog.products)

    $byArticleNumber = @{}
    $byGameId = @{}

    foreach ($product in $products) {
        $articleNumber = (
            [string]$product.id -replace '\D', ''
        )

        if (
            $articleNumber.Length -gt 0 -and
            $articleNumber.Length -le 5
        ) {
            $articleNumber = $articleNumber.PadLeft(5, '0')

            if (-not $byArticleNumber.ContainsKey($articleNumber)) {
                $byArticleNumber[$articleNumber] = @()
            }

            $byArticleNumber[$articleNumber] += $product
        }

        foreach ($gameFile in @($product.gameFiles)) {
            if ($null -eq $gameFile) {
                continue
            }

            $gameId = [string]$gameFile.id

            if ([string]::IsNullOrWhiteSpace($gameId)) {
                continue
            }

            $entry = [PSCustomObject]@{
                Product       = $product
                GameFile      = $gameFile
                GameProductId = $gameId
                ProductName   = [string]$product.name
                ArticleNumber = [string]$product.id
                AgeFrom       = $product.ageFrom
                AgeTo         = $product.ageTo
                Url           = [string]$gameFile.url
                FileName      = [string]$gameFile.fileName
                Version       = [string]$gameFile.version
            }

            if (-not $byGameId.ContainsKey($gameId)) {
                $byGameId[$gameId] = @()
            }

            $byGameId[$gameId] += $entry
        }
    }

    return [PSCustomObject]@{
        Products        = $products
        ByArticleNumber = $byArticleNumber
        ByGameId        = $byGameId
    }
}


# ============================================================================
# Select the preferred Ravensburger download entry for a GME Product-ID
# ============================================================================

function Get-PreferredCatalogGameEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GameProductId,

        [Parameter(Mandatory = $true)]
        $CatalogMaps
    )

    if (
        -not $CatalogMaps.ByGameId.ContainsKey(
            [string]$GameProductId
        )
    ) {
        return $null
    }

    $entries = @(
        $CatalogMaps.ByGameId[
            [string]$GameProductId
        ] |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace(
                [string]$_.Url
            )
        }
    )

    if ($entries.Count -eq 0) {
        return $null
    }

    if ($entries.Count -eq 1) {
        return $entries[0]
    }

    # If multiple entries have the same GME Product-ID,
    # use the highest version value where possible.
    #
    # Compare numerically first where possible; otherwise
    # Sort-Object falls back to the string representation.
    $sorted = @(
        $entries |
        Sort-Object `
            -Property @{
                Expression = {
                    $numeric = 0.0

                    if (
                        [double]::TryParse(
                            [string]$_.Version,
                            [System.Globalization.NumberStyles]::Any,
                            [System.Globalization.CultureInfo]::InvariantCulture,
                            [ref]$numeric
                        )
                    ) {
                        return $numeric
                    }

                    return -1.0
                }
            }, @{
                Expression = {
                    [string]$_.Version
                }
            } `
            -Descending
    )

    return $sorted[0]
}


# ============================================================================
# ISBN/EAN from RSS -> Ravensburger product -> GME Product-ID
# ============================================================================

function Find-BestCatalogProductsByTitle {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [array]$Products
    )

    $bestScore = 0.0
    $best = @()

    foreach ($product in $Products) {
        $score = Get-TitleSimilarity `
            -Title1 $Title `
            -Title2 ([string]$product.name)

        if ($score -gt ($bestScore + 0.0001)) {
            $bestScore = $score
            $best = @($product)
        }
        elseif ([math]::Abs($score - $bestScore) -lt 0.0001) {
            $best += $product
        }
    }

    return [PSCustomObject]@{
        Score    = $bestScore
        Products = @($best)
    }
}


function Resolve-LibraryRssToCatalog {
    param(
        [Parameter(Mandatory = $true)]
        [array]$RssItems,

        [Parameter(Mandatory = $true)]
        [hashtable]$BookIndex,

        [Parameter(Mandatory = $true)]
        $CatalogMaps
    )

    $resolvedByGameId = @{}
    $unresolved = @()

    foreach ($rssItem in $RssItems) {
        $itemResolved = $false

        # The ISBN/EAN from the RSS feed is always used as the primary identifier.
        #
        # However, some WinBIAP entries do not contain any
        # ISBN/EAN at all (x13= is also empty in the feed). Only for these
        # cases is there a title fallback, so that such tiptoi media are not
        # categorically excluded from downloading.
        if ($rssItem.Isbns.Count -eq 0) {
            if (
                -not [string]::IsNullOrWhiteSpace($rssItem.Title)
            ) {
                $titleMatch =
                    Find-BestCatalogProductsByTitle `
                        -Title $rssItem.Title `
                        -Products $CatalogMaps.Products

                if ($titleMatch.Score -ge 0.90) {
                    foreach ($product in @($titleMatch.Products)) {
                        foreach ($gameFile in @($product.gameFiles)) {
                            if ($null -eq $gameFile) {
                                continue
                            }

                            $gameId = [string]$gameFile.id

                            if ([string]::IsNullOrWhiteSpace($gameId)) {
                                continue
                            }

                            $candidate = [PSCustomObject]@{
                                RssTitle       = [string]$rssItem.Title
                                Isbn           = ""
                                BookNumber     = ""
                                ProductName    = [string]$product.name
                                ArticleNumber  = [string]$product.id
                                AgeFrom        = $product.ageFrom
                                AgeTo          = $product.ageTo
                                GameProductId  = $gameId
                                Url            = [string]$gameFile.url
                                FileName       = [string]$gameFile.fileName
                                Version        = [string]$gameFile.version
                            }

                            if (-not $resolvedByGameId.ContainsKey($gameId)) {
                                $resolvedByGameId[$gameId] = $candidate
                            }

                            $itemResolved = $true
                        }
                    }
                }
            }

            if (-not $itemResolved) {
                $unresolved += [PSCustomObject]@{
                    Title  = $rssItem.Title
                    Isbn   = ""
                    Reason = "Keine ISBN/EAN; Titel-Fallback nicht eindeutig"
                }
            }

            continue
        }

        foreach ($isbn in $rssItem.Isbns) {
            $bookNumber =
                Get-RavensburgerBookNumberFromIsbn13 -Isbn $isbn

            if ([string]::IsNullOrWhiteSpace($bookNumber)) {
                continue
            }

            $matchedProducts = @()

            # 1. Normal case:
            # ISBN 978-3-473-49286-2 -> book number 49286
            # and product.id == 49286
            if (
                $CatalogMaps.ByArticleNumber.ContainsKey(
                    $bookNumber
                )
            ) {
                $matchedProducts = @(
                    $CatalogMaps.ByArticleNumber[$bookNumber]
                )
            }

            # 2. Legacy inventory:
            # Book number and retail/catalog number may differ.
            # Knight example:
            # ISBN/book number 32910, old article number 00590.
            #
            # Then use the book number determined from the ISBN,
            # retrieve the official Ravensburger title for it, and search
            # for that title in the tiptoi catalog.
            if (
                $matchedProducts.Count -eq 0 -and
                $BookIndex.ContainsKey($bookNumber)
            ) {
                $serviceTitle = [string]$BookIndex[$bookNumber]

                $titleMatch =
                    Find-BestCatalogProductsByTitle `
                        -Title $serviceTitle `
                        -Products $CatalogMaps.Products

                if ($titleMatch.Score -ge 0.82) {
                    $matchedProducts = @(
                        $titleMatch.Products
                    )
                }
            }

            # 3. Final fallback:
            # The ISBN/EAN remains the starting point; only for the technical
            # mapping to the historical tiptoi article number is the
            # RSS title used.
            if (
                $matchedProducts.Count -eq 0 -and
                -not [string]::IsNullOrWhiteSpace($rssItem.Title)
            ) {
                $titleMatch =
                    Find-BestCatalogProductsByTitle `
                        -Title $rssItem.Title `
                        -Products $CatalogMaps.Products

                if ($titleMatch.Score -ge 0.88) {
                    $matchedProducts = @(
                        $titleMatch.Products
                    )
                }
            }

            foreach ($product in $matchedProducts) {
                foreach ($gameFile in @($product.gameFiles)) {
                    if ($null -eq $gameFile) {
                        continue
                    }

                    $gameId = [string]$gameFile.id

                    if ([string]::IsNullOrWhiteSpace($gameId)) {
                        continue
                    }

                    $candidate = [PSCustomObject]@{
                        RssTitle       = [string]$rssItem.Title
                        Isbn           = [string]$isbn
                        BookNumber     = [string]$bookNumber
                        ProductName    = [string]$product.name
                        ArticleNumber  = [string]$product.id
                        AgeFrom        = $product.ageFrom
                        AgeTo          = $product.ageTo
                        GameProductId  = $gameId
                        Url            = [string]$gameFile.url
                        FileName       = [string]$gameFile.fileName
                        Version        = [string]$gameFile.version
                    }

                    # If the same GME Product-ID occurs multiple times:
                    # keep the higher version where possible.
                    if (-not $resolvedByGameId.ContainsKey($gameId)) {
                        $resolvedByGameId[$gameId] = $candidate
                    }
                    else {
                        $oldVersion = [string]$resolvedByGameId[$gameId].Version
                        $newVersion = [string]$candidate.Version

                        if ($newVersion -gt $oldVersion) {
                            $resolvedByGameId[$gameId] = $candidate
                        }
                    }

                    $itemResolved = $true
                }
            }
        }

        if (-not $itemResolved) {
            $unresolved += [PSCustomObject]@{
                Title  = $rssItem.Title
                Isbn   = ($rssItem.Isbns -join ", ")
                Reason = "ISBN/EAN konnte keiner GME Product-ID zugeordnet werden"
            }
        }
    }

    return [PSCustomObject]@{
        ByGameId    = $resolvedByGameId
        Entries     = @($resolvedByGameId.Values)
        Unresolved  = @($unresolved)
    }
}


# ============================================================================
# Read optional tiptoi_mybooks.yml file
#
# No external PowerShell YAML module is required for the simple YAML structure
# used here.
#
# Standard format:
#
# mine:
#  - isbn: 978-3-473-32909-0 # comment
#  - isbn: 9783473329100
#  - id: 923
# ignore:
#  - isbn: 978-3-473-32911-0
#  - id: 924
#
# ISBN separators are optional.
#
# Backward compatibility:
# - "my_books:" is treated like "mine:".
# - "ingore:" is accepted as a typo alias for "ignore:".
# ============================================================================

function Get-BookListsConfiguration {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $mine = [PSCustomObject]@{
        Isbns      = @()
        ProductIds = @()
    }

    $ignore = [PSCustomObject]@{
        Isbns      = @()
        ProductIds = @()
    }

    $result = [PSCustomObject]@{
        Exists = $false
        Mine   = $mine
        Ignore = $ignore
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        return $result
    }

    $result.Exists = $true

    Write-Host ""
    Write-Host (
        "Buchlisten werden aus '{0}' geladen ..." -f
        ([System.IO.Path]::GetFileName($Path))
    )

    try {
        $lines =
            Get-Content `
                -LiteralPath $Path `
                -ErrorAction Stop
    }
    catch {
        Write-Warning (
            "tiptoi_mybooks.yml konnte nicht gelesen werden: " +
            $_.Exception.Message
        )

        return $result
    }

    $currentSection = ""

    foreach ($rawLine in $lines) {
        $line = [string]$rawLine

        if (
            $line -match
            '^\s*(?<section>mine|ignore|ingore|my_books)\s*:\s*(?:#.*)?$'
        ) {
            $sectionName =
                $Matches["section"].ToLowerInvariant()

            switch ($sectionName) {
                "mine" {
                    $currentSection = "mine"
                }

                "my_books" {
                    $currentSection = "mine"

                    Write-Warning (
                        "Der YAML-Knoten 'my_books' ist veraltet. " +
                        "Bitte kuenftig 'mine' verwenden."
                    )
                }

                "ignore" {
                    $currentSection = "ignore"
                }

                "ingore" {
                    $currentSection = "ignore"

                    Write-Warning (
                        "Der YAML-Knoten 'ingore' wird als 'ignore' behandelt."
                    )
                }
            }

            continue
        }

        # Another non-indented YAML root node ends
        # the list currently being read.
        if (
            $line -match '^\S[^:]*\s*:'
        ) {
            $currentSection = ""
            continue
        }

        if ([string]::IsNullOrWhiteSpace($currentSection)) {
            continue
        }

        if (
            $line -match
            '^\s*-\s*(?<key>isbn|id)\s*:\s*(?<value>.*?)\s*(?:#.*)?$'
        ) {
            $key =
                $Matches["key"].ToLowerInvariant()

            $value =
                [string]$Matches["value"]

            $value = $value.Trim()
            $value = $value.Trim('"')
            $value = $value.Trim("'")
            $value = $value.Trim()

            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            if ($currentSection -eq "mine") {
                $target = $result.Mine
            }
            else {
                $target = $result.Ignore
            }

            if ($key -eq "id") {
                $id =
                    $value -replace '\D', ''

                if (
                    -not [string]::IsNullOrWhiteSpace($id)
                ) {
                    $target.ProductIds +=
                        [string]$id
                }

                continue
            }

            # Normalize ISBN. Hyphens/spaces are optional.
            $isbn =
                $value -replace '[^0-9Xx]', ''

            if (
                $isbn.Length -eq 13 -or
                $isbn.Length -eq 10
            ) {
                $target.Isbns +=
                    $isbn
            }
            else {
                Write-Warning (
                    "Ungueltige ISBN in tiptoi_mybooks.yml ignoriert: '{0}'" -f
                    $value
                )
            }
        }
    }

    foreach ($list in @($result.Mine, $result.Ignore)) {
        $list.Isbns =
            @(
                $list.Isbns |
                Select-Object -Unique
            )

        $list.ProductIds =
            @(
                $list.ProductIds |
                Select-Object -Unique
            )
    }

    Write-Host (
        "    mine   : {0} direkte IDs, {1} ISBN-Eintraege" -f
        $result.Mine.ProductIds.Count,
        $result.Mine.Isbns.Count
    )

    Write-Host (
        "    ignore : {0} direkte IDs, {1} ISBN-Eintraege" -f
        $result.Ignore.ProductIds.Count,
        $result.Ignore.Isbns.Count
    )

    return $result
}


# ============================================================================
# ISBN-10 -> ISBN-13
# ============================================================================

function Convert-Isbn10ToIsbn13 {
    param(
        [string]$Isbn
    )

    $isbn10 =
        $Isbn -replace '[^0-9Xx]', ''

    if ($isbn10.Length -ne 10) {
        return $null
    }

    # Do not require a valid ISBN-10 check digit for conversion.
    # The first 9 digits are transferred into the 978 ISBN-13 namespace.
    $body =
        "978" + $isbn10.Substring(0, 9)

    $sum = 0

    for ($i = 0; $i -lt 12; $i++) {
        $digit =
            [int]::Parse(
                $body.Substring($i, 1)
            )

        if (($i % 2) -eq 0) {
            $sum += $digit
        }
        else {
            $sum += 3 * $digit
        }
    }

    $checkDigit =
        (10 - ($sum % 10)) % 10

    return "$body$checkDigit"
}


# ============================================================================
# Map configured ISBNs / IDs to GME Product-IDs
#
# Direct id: entries are carried over unchanged.
# ISBNs are resolved using the same Ravensburger mapping logic as
# the library RSS entries.
# ============================================================================

function Resolve-ConfiguredBookListToGameIds {
    param(
        [Parameter(Mandatory = $true)]
        $List,

        [Parameter(Mandatory = $true)]
        [hashtable]$BookIndex,

        [Parameter(Mandatory = $true)]
        $CatalogMaps,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $gameIds = @{}

    foreach ($id in @($List.ProductIds)) {
        if (
            -not [string]::IsNullOrWhiteSpace(
                [string]$id
            )
        ) {
            $gameIds[[string]$id] = $true
        }
    }

    $pseudoRssItems = @()

    foreach ($rawIsbn in @($List.Isbns)) {
        $isbn =
            [string]$rawIsbn

        if ($isbn.Length -eq 10) {
            $isbn =
                Convert-Isbn10ToIsbn13 `
                    -Isbn $isbn
        }

        if ([string]::IsNullOrWhiteSpace($isbn)) {
            continue
        }

        $pseudoRssItems +=
            [PSCustomObject]@{
                Title = ""
                Link  = ""
                Isbns = @([string]$isbn)
            }
    }

    $unresolved = @()

    if ($pseudoRssItems.Count -gt 0) {
        $resolution =
            Resolve-LibraryRssToCatalog `
                -RssItems $pseudoRssItems `
                -BookIndex $BookIndex `
                -CatalogMaps $CatalogMaps

        foreach ($entry in @($resolution.Entries)) {
            if (
                -not [string]::IsNullOrWhiteSpace(
                    [string]$entry.GameProductId
                )
            ) {
                $gameIds[
                    [string]$entry.GameProductId
                ] = $true
            }
        }

        $unresolved =
            @($resolution.Unresolved)
    }

    Write-Host (
        "    {0}: {1} GME Product-ID(s) erkannt." -f
        $Label,
        $gameIds.Count
    )

    if ($unresolved.Count -gt 0) {
        Write-Warning (
            "{0} ISBN-Eintrag/Eintraege aus YAML-Liste '{1}' konnten " +
            "keiner GME Product-ID zugeordnet werden." -f
            $unresolved.Count,
            $Label
        )
    }

    return [PSCustomObject]@{
        GameIds    = $gameIds
        Unresolved = $unresolved
    }
}


# ============================================================================
# Determine availability status of a local GME
#
# Priority:
#   1. Besitz
#   2. Buecherei
#   3. Unbekannt
# ============================================================================
# Determine availability status of a local GME
#
# Priority:
#   1. Ignoriert
#   2. Besitz
#   3. Buecherei
#   4. Unbekannt
# ============================================================================

function Get-LocalAvailabilityStatus {
    param(
        [Parameter(Mandatory = $true)]
        $Record,

        [Parameter(Mandatory = $true)]
        [hashtable]$OwnedGameIds,

        [Parameter(Mandatory = $true)]
        [hashtable]$IgnoredGameIds,

        [Parameter(Mandatory = $true)]
        [hashtable]$LibraryGameIds
    )

    if (
        -not $Record.Valid -or
        [string]::IsNullOrWhiteSpace(
            [string]$Record.ProductId
        )
    ) {
        return "Unbekannt"
    }

    # Ignore has the highest priority. This keeps an explicitly ignored
    # book "Ignoriert" even when it is also listed in mine or in the RSS feed.
    if (
        $IgnoredGameIds.ContainsKey(
            [string]$Record.ProductId
        )
    ) {
        return "Ignoriert"
    }

    if (
        $OwnedGameIds.ContainsKey(
            [string]$Record.ProductId
        )
    ) {
        return "Besitz"
    }

    if (
        $LibraryGameIds.ContainsKey(
            [string]$Record.ProductId
        )
    ) {
        return "Buecherei"
    }

    return "Unbekannt"
}


# ============================================================================
# Format age recommendation for display
#
# Rules analogous to the download age filter:
# - both missing     -> "keine Angabe"
# - ageFrom missing  -> 0
# - ageTo missing    -> 99
# - otherwise        -> "<ageFrom>-<ageTo>"
# ============================================================================

function Get-AgeRecommendationText {
    param(
        $Entry
    )

    if ($null -eq $Entry) {
        return "keine Angabe"
    }

    $rawFrom = $Entry.AgeFrom
    $rawTo   = $Entry.AgeTo

    $hasFrom = $false
    $hasTo   = $false

    $ageFrom = 0
    $ageTo   = 99

    if (
        $null -ne $rawFrom -and
        -not [string]::IsNullOrWhiteSpace([string]$rawFrom)
    ) {
        $parsedFrom = 0

        if (
            [int]::TryParse(
                [string]$rawFrom,
                [ref]$parsedFrom
            )
        ) {
            $ageFrom = $parsedFrom
            $hasFrom = $true
        }
    }

    if (
        $null -ne $rawTo -and
        -not [string]::IsNullOrWhiteSpace([string]$rawTo)
    ) {
        $parsedTo = 0

        if (
            [int]::TryParse(
                [string]$rawTo,
                [ref]$parsedTo
            )
        ) {
            $ageTo = $parsedTo
            $hasTo = $true
        }
    }

    if (-not $hasFrom -and -not $hasTo) {
        return "keine Angabe"
    }

    if (-not $hasFrom) {
        $ageFrom = 0
    }

    if (-not $hasTo) {
        $ageTo = 99
    }

    return ("{0}-{1}" -f $ageFrom, $ageTo)
}


# ============================================================================
# Age filter for download
# ============================================================================

function Test-DownloadAgeMatch {
    param(
        [Parameter(Mandatory = $true)]
        $Entry,

        [Parameter(Mandatory = $true)]
        [int]$RequestedAge,

        [Parameter(Mandatory = $true)]
        [bool]$FilterEnabled
    )

    # No -age specified -> no age filtering.
    if (-not $FilterEnabled) {
        return $true
    }

    $rawFrom = $Entry.AgeFrom
    $rawTo   = $Entry.AgeTo

    $hasFrom = $false
    $hasTo   = $false

    $ageFrom = 0
    $ageTo   = 99

    if (
        $null -ne $rawFrom -and
        -not [string]::IsNullOrWhiteSpace([string]$rawFrom)
    ) {
        $parsedFrom = 0

        if (
            [int]::TryParse(
                [string]$rawFrom,
                [ref]$parsedFrom
            )
        ) {
            $ageFrom = $parsedFrom
            $hasFrom = $true
        }
    }

    if (
        $null -ne $rawTo -and
        -not [string]::IsNullOrWhiteSpace([string]$rawTo)
    ) {
        $parsedTo = 0

        if (
            [int]::TryParse(
                [string]$rawTo,
                [ref]$parsedTo
            )
        ) {
            $ageTo = $parsedTo
            $hasTo = $true
        }
    }

    # If both values are missing, the book is included
    # regardless of age.
    if (-not $hasFrom -and -not $hasTo) {
        return $true
    }

    # If only ageFrom is missing -> 0.
    if (-not $hasFrom) {
        $ageFrom = 0
    }

    # If only ageTo is missing -> 99.
    if (-not $hasTo) {
        $ageTo = 99
    }

    return (
        ($ageFrom -le $RequestedAge) -and
        ($RequestedAge -le $ageTo)
    )
}


# ============================================================================
# Local GME files
# ============================================================================

function Get-LocalGmeRecords {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        $CatalogMaps
    )

    $files = @(
        Get-ChildItem `
            -LiteralPath $Directory `
            -Filter "*.gme" `
            -File
    )

    $records = @()

    foreach ($file in $files) {
        try {
            $gameId = Get-GmeProductId -Path $file.FullName

            $title = "<Product-ID $gameId unbekannt>"

            if ($CatalogMaps.ByGameId.ContainsKey($gameId)) {
                $title = [string](
                    $CatalogMaps.ByGameId[$gameId][0].ProductName
                )
            }

            $records += [PSCustomObject]@{
                File          = $file
                FileName      = $file.Name
                FullName      = $file.FullName
                Length        = [long]$file.Length
                ProductId     = [string]$gameId
                Title         = $title
                Valid         = $true
                ErrorMessage  = ""
            }
        }
        catch {
            $records += [PSCustomObject]@{
                File          = $file
                FileName      = $file.Name
                FullName      = $file.FullName
                Length        = [long]$file.Length
                ProductId     = ""
                Title         = "<GME nicht lesbar>"
                Valid         = $false
                ErrorMessage  = $_.Exception.Message
            }
        }
    }

    return $records
}


# ============================================================================
# Remote file size
# ============================================================================

function Get-RemoteFileSizeBytes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    # Try HEAD first
    try {
        $request = [System.Net.HttpWebRequest]::Create($Url)
        $request.Method = "HEAD"
        $request.AllowAutoRedirect = $true
        $request.UserAgent = "PowerShell-TiptoiLibrarySync/16.0"
        $request.Timeout = 30000
        $request.ReadWriteTimeout = 30000

        $response = $request.GetResponse()

        try {
            $length = [long]$response.ContentLength

            if ($length -gt 0) {
                return $length
            }
        }
        finally {
            $response.Close()
        }
    }
    catch {
    }

    # Fallback: Range GET 0-0.
    # The body is not read; only Content-Range or
    # Content-Length is needed.
    try {
        $request = [System.Net.HttpWebRequest]::Create($Url)
        $request.Method = "GET"
        $request.AllowAutoRedirect = $true
        $request.UserAgent = "PowerShell-TiptoiLibrarySync/16.0"
        $request.Timeout = 30000
        $request.ReadWriteTimeout = 30000
        $request.AddRange(0, 0)

        $response = $request.GetResponse()

        try {
            $contentRange = [string]$response.Headers["Content-Range"]

            if (
                $contentRange -match '/(\d+)\s*$'
            ) {
                return [long]$Matches[1]
            }

            $length = [long]$response.ContentLength

            # If the server ignores Range and returns 200 with the complete
            # Content-Length, that value is also the file size.
            if ($length -gt 1) {
                return $length
            }
        }
        finally {
            $response.Close()
        }
    }
    catch {
    }

    return [long]-1
}


function Get-FreeSpaceBytes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $root = [System.IO.Path]::GetPathRoot($fullPath)

    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "Laufwerk fuer '$Path' konnte nicht bestimmt werden."
    }

    try {
        $drive = New-Object System.IO.DriveInfo($root)
        return [long]$drive.AvailableFreeSpace
    }
    catch {
        throw (
            "Freier Speicherplatz auf '$root' konnte nicht bestimmt werden: " +
            $_.Exception.Message
        )
    }
}


# ============================================================================
# Download
# ============================================================================

function Get-SafeDownloadPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$PreferredFileName,

        [Parameter(Mandatory = $true)]
        [string]$GameProductId
    )

    $name = $PreferredFileName

    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = "tiptoi_$GameProductId.gme"
    }

    foreach ($invalid in [System.IO.Path]::GetInvalidFileNameChars()) {
        $name = $name.Replace([string]$invalid, "_")
    }

    if (
        -not $name.EndsWith(
            ".gme",
            [System.StringComparison]::OrdinalIgnoreCase
        )
    ) {
        $name += ".gme"
    }

    $path = Join-Path $Directory $name

    if (-not (Test-Path -LiteralPath $path)) {
        return $path
    }

    # Name collision: by definition, the existing file does not have the same
    # Product-ID (otherwise it would not be in the download list).
    $base = [System.IO.Path]::GetFileNameWithoutExtension($name)

    return (
        Join-Path `
            $Directory `
            ("{0}_{1}.gme" -f $base, $GameProductId)
    )
}


function Download-GmeEntry {
    param(
        [Parameter(Mandatory = $true)]
        $Entry,

        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    if ([string]::IsNullOrWhiteSpace($Entry.Url)) {
        throw (
            "Keine Download-URL fuer GME Product-ID " +
            $Entry.GameProductId
        )
    }

    $destination =
        Get-SafeDownloadPath `
            -Directory $Directory `
            -PreferredFileName $Entry.FileName `
            -GameProductId $Entry.GameProductId

    $tempPath = "$destination.download"

    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force
    }

    try {
        Invoke-WebRequest `
            -Uri $Entry.Url `
            -OutFile $tempPath `
            -UseBasicParsing `
            -TimeoutSec 0

        $downloadedProductId =
            Get-GmeProductId -Path $tempPath

        if (
            [string]$downloadedProductId -ne
            [string]$Entry.GameProductId
        ) {
            throw (
                "Product-ID-Pruefung fehlgeschlagen. Erwartet: {0}, erhalten: {1}" -f
                $Entry.GameProductId,
                $downloadedProductId
            )
        }

        Move-Item `
            -LiteralPath $tempPath `
            -Destination $destination `
            -Force

        return $destination
    }
    catch {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }

        throw
    }
}


# ============================================================================
# Re-download and replace an existing GME file from the Ravensburger server
#
# The existing file is replaced only after:
#   1. the download has completed,
#   2. the GME Product-ID of the temporary file has been validated,
#   3. optionally, the expected server size has been validated.
# ============================================================================

function Reload-ExistingGmeFile {
    param(
        [Parameter(Mandatory = $true)]
        $CatalogEntry,

        [Parameter(Mandatory = $true)]
        [string]$ExistingPath,

        [long]$ExpectedSize = -1
    )

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$CatalogEntry.Url
        )
    ) {
        throw (
            "Keine Ravensburger-Download-URL fuer GME Product-ID " +
            $CatalogEntry.GameProductId
        )
    }

    $tempPath = "$ExistingPath.reload"

    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item `
            -LiteralPath $tempPath `
            -Force
    }

    try {
        Invoke-WebRequest `
            -Uri $CatalogEntry.Url `
            -OutFile $tempPath `
            -UseBasicParsing `
            -TimeoutSec 0

        $downloadedProductId =
            Get-GmeProductId `
                -Path $tempPath

        if (
            [string]$downloadedProductId -ne
            [string]$CatalogEntry.GameProductId
        ) {
            throw (
                "Product-ID-Pruefung fehlgeschlagen. Erwartet: {0}, erhalten: {1}" -f
                $CatalogEntry.GameProductId,
                $downloadedProductId
            )
        }

        $tempFile =
            Get-Item `
                -LiteralPath $tempPath

        if (
            $ExpectedSize -ge 0 -and
            [long]$tempFile.Length -ne
            [long]$ExpectedSize
        ) {
            throw (
                "Groessenpruefung nach Reload fehlgeschlagen. " +
                "Server: {0}, Download: {1}" -f
                (Format-ByteSize -Bytes $ExpectedSize),
                (Format-ByteSize -Bytes ([long]$tempFile.Length))
            )
        }

        Move-Item `
            -LiteralPath $tempPath `
            -Destination $ExistingPath `
            -Force
    }
    catch {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item `
                -LiteralPath $tempPath `
                -Force `
                -ErrorAction SilentlyContinue
        }

        throw
    }
}


# ============================================================================
# Determine expected Ravensburger filename
# ============================================================================

function Get-CatalogExpectedFileName {
    param(
        [Parameter(Mandatory = $true)]
        $CatalogEntry
    )

    $name = [string]$CatalogEntry.FileName

    if (-not [string]::IsNullOrWhiteSpace($name)) {
        return [System.IO.Path]::GetFileName($name)
    }

    # Fallback: derive filename from download URL.
    try {
        $uri = New-Object System.Uri([string]$CatalogEntry.Url)

        $name =
            [System.Uri]::UnescapeDataString(
                [System.IO.Path]::GetFileName(
                    $uri.AbsolutePath
                )
            )

        if (-not [string]::IsNullOrWhiteSpace($name)) {
            return $name
        }
    }
    catch {
    }

    return ""
}


# ============================================================================
# Handle duplicates with the same GME Product-ID
#
# Rules:
#
# 1. No duplicate matches the server size:
#       -> Offer to delete all duplicates + reload.
#
# 2. At least one duplicate matches the server size, while others do not:
#       -> Offer to delete the invalid duplicates.
#
# 3. Multiple remaining duplicates match the server size:
#       -> If exactly one has the Ravensburger filename,
#          offer to delete the other file(s).
#       -> If none matches the Ravensburger filename,
#          let the user choose which file to keep.
#
# The function returns the subsequently re-read local inventory.
# ============================================================================

function Resolve-DuplicateLocalGmeFiles {
    param(
        [Parameter(Mandatory = $true)]
        [array]$LocalRecords,

        [Parameter(Mandatory = $true)]
        $CatalogMaps,

        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $duplicateGroups = @(
        $LocalRecords |
        Where-Object {
            $_.Valid -and
            -not [string]::IsNullOrWhiteSpace(
                [string]$_.ProductId
            )
        } |
        Group-Object `
            -Property ProductId |
        Where-Object {
            $_.Count -gt 1
        }
    )

    if ($duplicateGroups.Count -eq 0) {
        return $LocalRecords
    }

    Write-Host ""
    Write-Warning (
        "{0} GME Product-ID(s) sind mehrfach im Ordner vorhanden." -f
        $duplicateGroups.Count
    )
    Write-Host ""

    # Check the same server URL only once.
    $remoteSizeCache = @{}

    foreach ($group in $duplicateGroups) {
        $productId = [string]$group.Name

        # Only consider files that still exist.
        $records = @(
            $group.Group |
            Where-Object {
                Test-Path -LiteralPath $_.FullName
            }
        )

        if ($records.Count -le 1) {
            continue
        }

        Write-Host "------------------------------------------------------------"
        Write-Host (
            "Duplikat: GME Product-ID {0}" -f
            $productId
        )

        foreach ($record in $records) {
            Write-Host (
                "    {0} ({1})" -f
                $record.FileName,
                (Format-ByteSize -Bytes ([long]$record.Length))
            )
        }

        $catalogEntry =
            Get-PreferredCatalogGameEntry `
                -GameProductId $productId `
                -CatalogMaps $CatalogMaps

        if ($null -eq $catalogEntry) {
            Write-Warning (
                "    Keine Ravensburger-Download-Information fuer diese " +
                "Product-ID gefunden. Duplikate bleiben unveraendert."
            )
            Write-Host ""
            continue
        }

        $url = [string]$catalogEntry.Url

        if ($remoteSizeCache.ContainsKey($url)) {
            $remoteSize = [long]$remoteSizeCache[$url]
        }
        else {
            $remoteSize =
                Get-RemoteFileSizeBytes `
                    -Url $url

            $remoteSizeCache[$url] =
                [long]$remoteSize
        }

        if ($remoteSize -lt 0) {
            Write-Warning (
                "    Server-Dateigroesse konnte nicht ermittelt werden. " +
                "Duplikate bleiben unveraendert."
            )
            Write-Host ""
            continue
        }

        $expectedFileName =
            Get-CatalogExpectedFileName `
                -CatalogEntry $catalogEntry

        Write-Host (
            "    Ravensburger-Groesse : {0}" -f
            (Format-ByteSize -Bytes $remoteSize)
        )

        if (
            -not [string]::IsNullOrWhiteSpace(
                $expectedFileName
            )
        ) {
            Write-Host (
                "    Ravensburger-Name   : {0}" -f
                $expectedFileName
            )
        }
        else {
            Write-Host (
                "    Ravensburger-Name   : nicht ermittelbar"
            )
        }

        $sizeValid = @(
            $records |
            Where-Object {
                [long]$_.Length -eq
                [long]$remoteSize
            }
        )

        $sizeInvalid = @(
            $records |
            Where-Object {
                [long]$_.Length -ne
                [long]$remoteSize
            }
        )

        # ----------------------------------------------------------------
        # Case 1:
        # All duplicates have an incorrect size.
        # -> Offer to delete all + reload.
        # ----------------------------------------------------------------

        if ($sizeValid.Count -eq 0) {
            Write-Warning (
                "    Alle {0} Duplikate weichen von der Ravensburger-" +
                "Dateigroesse ab." -f
                $records.Count
            )

            # Before deleting, make sure that after reclaiming the space from
            # the files to be deleted, enough space is available for the reload.
            $freeBytes =
                Get-FreeSpaceBytes `
                    -Path $Directory

            $reclaimableBytes =
                [long](
                    (
                        $records |
                        Measure-Object `
                            -Property Length `
                            -Sum
                    ).Sum
                )

            $projectedFree =
                $freeBytes + $reclaimableBytes

            if ($projectedFree -lt $remoteSize) {
                Write-Warning (
                    "    Selbst nach dem Loeschen aller Duplikate waere " +
                    "nicht genug Speicher fuer den Reload vorhanden. " +
                    "Es wird nichts geloescht."
                )
                Write-Host ""
                continue
            }

            $answer =
                Read-Host (
                    "    Alle Duplikate loeschen und eine korrekte Datei " +
                    "neu laden? [j/N]"
                )

            if ($answer -notmatch '^(?i:j|ja|y|yes)$') {
                Write-Host "    Keine Aenderung."
                Write-Host ""
                continue
            }

            $deleteFailed = $false

            foreach ($record in $records) {
                try {
                    Remove-Item `
                        -LiteralPath $record.FullName `
                        -Force

                    Write-Host (
                        "    Geloescht: {0}" -f
                        $record.FileName
                    )
                }
                catch {
                    Write-Warning (
                        "    Konnte '{0}' nicht loeschen: {1}" -f
                        $record.FileName,
                        $_.Exception.Message
                    )

                    $deleteFailed = $true
                }
            }

            if ($deleteFailed) {
                Write-Warning (
                    "    Reload wird nicht automatisch gestartet, weil nicht " +
                    "alle Duplikate geloescht werden konnten."
                )
                Write-Host ""
                continue
            }

            $destination =
                Get-SafeDownloadPath `
                    -Directory $Directory `
                    -PreferredFileName $expectedFileName `
                    -GameProductId $productId

            try {
                Write-Host "    Reload wird gestartet ..."

                Reload-ExistingGmeFile `
                    -CatalogEntry $catalogEntry `
                    -ExistingPath $destination `
                    -ExpectedSize $remoteSize

                Write-Host (
                    "    Reload erfolgreich: {0}" -f
                    ([System.IO.Path]::GetFileName($destination))
                )
            }
            catch {
                Write-Warning (
                    "    Reload fehlgeschlagen: " +
                    $_.Exception.Message
                )
            }

            Write-Host ""
            continue
        }

        # ----------------------------------------------------------------
        # Case 2:
        # At least one file has the correct size, while others do not.
        # -> Offer incorrect duplicates for deletion individually.
        # ----------------------------------------------------------------

        if ($sizeInvalid.Count -gt 0) {
            foreach ($record in $sizeInvalid) {
                Write-Warning (
                    "    Ungueltiges Duplikat: {0} ({1})" -f
                    $record.FileName,
                    (Format-ByteSize -Bytes ([long]$record.Length))
                )

                $answer =
                    Read-Host (
                        "    Ungueltiges Duplikat loeschen? [j/N]"
                    )

                if ($answer -match '^(?i:j|ja|y|yes)$') {
                    try {
                        Remove-Item `
                            -LiteralPath $record.FullName `
                            -Force

                        Write-Host (
                            "    Geloescht: {0}" -f
                            $record.FileName
                        )
                    }
                    catch {
                        Write-Warning (
                            "    Konnte '{0}' nicht loeschen: {1}" -f
                            $record.FileName,
                            $_.Exception.Message
                        )
                    }
                }
                else {
                    Write-Host (
                        "    Datei bleibt erhalten: {0}" -f
                        $record.FileName
                    )
                }
            }
        }

        # Use only correctly sized files for the filename decision.
        $remainingValid = @(
            $sizeValid |
            Where-Object {
                Test-Path -LiteralPath $_.FullName
            }
        )

        if ($remainingValid.Count -le 1) {
            Write-Host ""
            continue
        }

        # ----------------------------------------------------------------
        # Case 3:
        # Multiple correctly sized duplicates.
        # ----------------------------------------------------------------

        $nameMatches = @()

        if (
            -not [string]::IsNullOrWhiteSpace(
                $expectedFileName
            )
        ) {
            $nameMatches = @(
                $remainingValid |
                Where-Object {
                    [string]::Equals(
                        [string]$_.FileName,
                        [string]$expectedFileName,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }
            )
        }

        # Exactly one file has the Ravensburger filename:
        # keep it and offer the others for deletion.
        if ($nameMatches.Count -eq 1) {
            $keepRecord = $nameMatches[0]

            Write-Host (
                "    '{0}' entspricht dem Ravensburger-Dateinamen." -f
                $keepRecord.FileName
            )

            foreach (
                $record in
                $remainingValid |
                Where-Object {
                    $_.FullName -ne
                    $keepRecord.FullName
                }
            ) {
                $answer =
                    Read-Host (
                        "    Duplikat '{0}' loeschen? [j/N]" -f
                        $record.FileName
                    )

                if ($answer -match '^(?i:j|ja|y|yes)$') {
                    try {
                        Remove-Item `
                            -LiteralPath $record.FullName `
                            -Force

                        Write-Host (
                            "    Geloescht: {0}" -f
                            $record.FileName
                        )
                    }
                    catch {
                        Write-Warning (
                            "    Konnte '{0}' nicht loeschen: {1}" -f
                            $record.FileName,
                            $_.Exception.Message
                        )
                    }
                }
                else {
                    Write-Host (
                        "    Datei bleibt erhalten: {0}" -f
                        $record.FileName
                    )
                }
            }

            Write-Host ""
            continue
        }

        # No file matches the Ravensburger filename
        # (or the expected filename could not be determined):
        # the user chooses which valid file to keep.
        Write-Warning (
            "    Mehrere groessenrichtige Duplikate vorhanden; keines " +
            "entspricht eindeutig dem Ravensburger-Dateinamen."
        )

        Write-Host ""
        Write-Host "    Welche Datei soll behalten werden?"

        for (
            $i = 0;
            $i -lt $remainingValid.Count;
            $i++
        ) {
            Write-Host (
                "      [{0}] {1}" -f
                ($i + 1),
                $remainingValid[$i].FileName
            )
        }

        $selection =
            Read-Host (
                "    Nummer eingeben; Enter = keine Aenderung"
            )

        if ([string]::IsNullOrWhiteSpace($selection)) {
            Write-Host "    Keine Aenderung."
            Write-Host ""
            continue
        }

        $selectedIndex = 0

        if (
            -not [int]::TryParse(
                $selection,
                [ref]$selectedIndex
            )
        ) {
            Write-Warning "    Ungueltige Auswahl. Keine Aenderung."
            Write-Host ""
            continue
        }

        if (
            $selectedIndex -lt 1 -or
            $selectedIndex -gt
                $remainingValid.Count
        ) {
            Write-Warning "    Ungueltige Auswahl. Keine Aenderung."
            Write-Host ""
            continue
        }

        $keepRecord =
            $remainingValid[
                $selectedIndex - 1
            ]

        Write-Host (
            "    Wird behalten: {0}" -f
            $keepRecord.FileName
        )

        foreach (
            $record in
            $remainingValid |
            Where-Object {
                $_.FullName -ne
                $keepRecord.FullName
            }
        ) {
            try {
                Remove-Item `
                    -LiteralPath $record.FullName `
                    -Force

                Write-Host (
                    "    Geloescht: {0}" -f
                    $record.FileName
                )
            }
            catch {
                Write-Warning (
                    "    Konnte '{0}' nicht loeschen: {1}" -f
                    $record.FileName,
                    $_.Exception.Message
                )
            }
        }

        Write-Host ""
    }

    # Re-read the complete local inventory after all actions.
    return @(
        Get-LocalGmeRecords `
            -Directory $Directory `
            -CatalogMaps $CatalogMaps
    )
}


# ============================================================================
# Check existing local GME files against Ravensburger server sizes
# ============================================================================

function Fix-LocalGmeFiles {
    param(
        [Parameter(Mandatory = $true)]
        [array]$LocalRecords,

        [Parameter(Mandatory = $true)]
        $CatalogMaps,

        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    Write-Host ""
    Write-Host "============================================================"
    Write-Host " Pruefen und Reparieren vorhandener GME-Dateien"
    Write-Host "============================================================"
    Write-Host ""

    if ($LocalRecords.Count -eq 0) {
        Write-Host "Keine vorhandenen GME-Dateien zu pruefen."
        Write-Host ""

        return
    }

    # Handle duplicates with the same Product-ID first.
    $LocalRecords = @(
        Resolve-DuplicateLocalGmeFiles `
            -LocalRecords $LocalRecords `
            -CatalogMaps $CatalogMaps `
            -Directory $Directory
    )

    # If the user declined duplicate cleanup,
    # these Product-IDs are not bothered again with reload prompts
    # during the subsequent individual check.
    $remainingDuplicateIds = @{}

    $stillDuplicateGroups = @(
        $LocalRecords |
        Where-Object {
            $_.Valid -and
            -not [string]::IsNullOrWhiteSpace(
                [string]$_.ProductId
            )
        } |
        Group-Object `
            -Property ProductId |
        Where-Object {
            $_.Count -gt 1
        }
    )

    foreach ($group in $stillDuplicateGroups) {
        $remainingDuplicateIds[
            [string]$group.Name
        ] = $true
    }

    # The same URL only needs to be queried once via HEAD/Range.
    $remoteSizeCache = @{}

    $checked = 0
    $matching = 0
    $mismatches = 0
    $reloaded = 0
    $notVerifiable = 0

    $counter = 0

    foreach ($record in $LocalRecords) {
        $counter++

        Write-Host (
            "[{0}/{1}] {2}" -f
            $counter,
            $LocalRecords.Count,
            $record.FileName
        )

        if (-not $record.Valid) {
            Write-Warning (
                "    Datei ist keine lesbare GME-Datei und kann nicht " +
                "gegen Ravensburger verifiziert werden."
            )

            $notVerifiable++
            Write-Host ""
            continue
        }

        if (
            $remainingDuplicateIds.ContainsKey(
                [string]$record.ProductId
            )
        ) {
            Write-Warning (
                "    Product-ID {0} ist weiterhin mehrfach vorhanden. " +
                "Die Einzelpruefung wird fuer diese Datei uebersprungen." -f
                $record.ProductId
            )

            Write-Host ""
            continue
        }

        $catalogEntry =
            Get-PreferredCatalogGameEntry `
                -GameProductId ([string]$record.ProductId) `
                -CatalogMaps $CatalogMaps

        if ($null -eq $catalogEntry) {
            Write-Warning (
                "    GME Product-ID {0} ist im aktuellen Ravensburger-" +
                "Katalog nicht mit einer Download-URL vorhanden." -f
                $record.ProductId
            )

            $notVerifiable++
            Write-Host ""
            continue
        }

        $url = [string]$catalogEntry.Url

        if ($remoteSizeCache.ContainsKey($url)) {
            $remoteSize = [long]$remoteSizeCache[$url]
        }
        else {
            $remoteSize =
                Get-RemoteFileSizeBytes `
                    -Url $url

            $remoteSizeCache[$url] =
                [long]$remoteSize
        }

        if ($remoteSize -lt 0) {
            Write-Warning (
                "    Server-Dateigroesse konnte nicht ermittelt werden."
            )

            $notVerifiable++
            Write-Host ""
            continue
        }

        $checked++

        $localSize =
            [long]$record.Length

        Write-Host (
            "    GME Product-ID : {0}" -f
            $record.ProductId
        )

        Write-Host (
            "    Lokal          : {0}" -f
            (Format-ByteSize -Bytes $localSize)
        )

        Write-Host (
            "    Ravensburger   : {0}" -f
            (Format-ByteSize -Bytes $remoteSize)
        )

        if ($localSize -eq $remoteSize) {
            Write-Host "    Ergebnis       : OK"

            $matching++
            Write-Host ""
            continue
        }

        $mismatches++

        Write-Warning (
            "    GROESSENABWEICHUNG bei '{0}'." -f
            $record.FileName
        )

        $answer =
            Read-Host "    Datei vom Ravensburger-Server neu laden? [j/N]"

        if ($answer -notmatch '^(?i:j|ja|y|yes)$') {
            Write-Host "    Kein Reload."
            Write-Host ""
            continue
        }

        # A complete second file is temporarily required for the reload.
        # Therefore, check free space first.
        $freeBytes =
            Get-FreeSpaceBytes `
                -Path $Directory

        if ($freeBytes -lt $remoteSize) {
            Write-Warning (
                "    Reload nicht moeglich: Es werden temporaer {0} benoetigt, " +
                "aber nur {1} sind frei." -f
                (Format-ByteSize -Bytes $remoteSize),
                (Format-ByteSize -Bytes $freeBytes)
            )

            Write-Host ""
            continue
        }

        try {
            Write-Host "    Reload wird gestartet ..."

            Reload-ExistingGmeFile `
                -CatalogEntry $catalogEntry `
                -ExistingPath $record.FullName `
                -ExpectedSize $remoteSize

            # Update the local record so that a subsequent
            # -download run continues with current file sizes.
            $updatedFile =
                Get-Item `
                    -LiteralPath $record.FullName

            $record.File =
                $updatedFile

            $record.Length =
                [long]$updatedFile.Length

            Write-Host "    Reload erfolgreich."

            $reloaded++
        }
        catch {
            Write-Warning (
                "    Reload fehlgeschlagen: " +
                $_.Exception.Message
            )
        }

        Write-Host ""
    }

    Write-Host "Pruefung/Reparatur abgeschlossen:"
    Write-Host (
        "    Geprueft             : {0}" -f
        $checked
    )
    Write-Host (
        "    Groesse korrekt      : {0}" -f
        $matching
    )
    Write-Host (
        "    Abweichungen         : {0}" -f
        $mismatches
    )
    Write-Host (
        "    Neu geladen          : {0}" -f
        $reloaded
    )
    Write-Host (
        "    Nicht verifizierbar  : {0}" -f
        $notVerifiable
    )
    Write-Host ""

    # Return the final local inventory after reload/delete actions.
    return @(
        Get-LocalGmeRecords `
            -Directory $Directory `
            -CatalogMaps $CatalogMaps
    )
}


# ============================================================================
# Manual cleanup mode
#
# -cleanup All:
#     Files with availability status "Unbekannt" or "Ignoriert"
#
# -cleanup or -cleanup Ignored:
#     Only files with availability status "Ignoriert"
#
# Candidates are offered in descending file-size order.
# ============================================================================

function Invoke-ManualCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [array]$LocalRecords,

        [Parameter(Mandatory = $true)]
        [ValidateSet("All", "Ignored")]
        [string]$Mode
    )

    if ($Mode -eq "Ignored") {
        $candidates = @(
            $LocalRecords |
            Where-Object {
                $_.Valid -and
                $_.AvailabilityStatus -eq "Ignoriert"
            } |
            Sort-Object `
                -Property Length `
                -Descending
        )
    }
    else {
        $candidates = @(
            $LocalRecords |
            Where-Object {
                $_.Valid -and
                (
                    $_.AvailabilityStatus -eq "Unbekannt" -or
                    $_.AvailabilityStatus -eq "Ignoriert"
                )
            } |
            Sort-Object `
                -Property Length `
                -Descending
        )
    }

    Write-Host ""
    Write-Host "============================================================"
    Write-Host (
        " Cleanup-Modus: {0}" -f
        $Mode
    )
    Write-Host "============================================================"
    Write-Host ""

    if ($candidates.Count -eq 0) {
        Write-Host "Keine passenden Cleanup-Kandidaten gefunden."
        Write-Host ""

        return
    }

    Write-Host (
        "{0} Datei(en) werden nach Dateigroesse absteigend angeboten." -f
        $candidates.Count
    )

    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate.FullName)) {
            continue
        }

        Write-Host ""
        Write-Host (
            "Status:           {0}" -f
            $candidate.AvailabilityStatus
        )

        Write-Host (
            "Dateiname:        {0}" -f
            $candidate.FileName
        )

        Write-Host (
            "Titel:            {0}" -f
            $candidate.Title
        )

        Write-Host (
            "GME Product-ID:   {0}" -f
            $candidate.ProductId
        )

        Write-Host (
            "Dateigroesse:     {0}" -f
            (Format-ByteSize -Bytes ([long]$candidate.Length))
        )

        $answer =
            Read-Host "Datei loeschen? [j/N]"

        if ($answer -match '^(?i:j|ja|y|yes)$') {
            try {
                Remove-Item `
                    -LiteralPath $candidate.FullName `
                    -Force `
                    -ErrorAction Stop

                Write-Host "Datei geloescht."
            }
            catch {
                Write-Warning (
                    "Datei konnte nicht geloescht werden: " +
                    $_.Exception.Message
                )
            }
        }
        else {
            Write-Host "Datei bleibt erhalten."
        }
    }

    Write-Host ""
    Write-Host "Cleanup abgeschlossen."
    Write-Host ""
}


# ============================================================================
# Determine status for table listing
#
# This check is NOT destructive and runs automatically for every normal
# listing.
#
# Display status priority:
#
#   Duplikat       -> Product-ID exists more than once
#   Falsche Groesse -> not a duplicate, but local size != server size
#   OK             -> matching size and no duplicate
#
# If the file or server size cannot be checked,
# "Nicht pruefbar" is used.
# ============================================================================

function Add-LocalFileCheckStatus {
    param(
        [Parameter(Mandatory = $true)]
        [array]$LocalRecords,

        [Parameter(Mandatory = $true)]
        $CatalogMaps
    )

    if ($LocalRecords.Count -eq 0) {
        return $LocalRecords
    }

    Write-Host ""
    Write-Host "Dateistatus wird gegen Ravensburger geprueft ..."

    # Product-ID frequencies for duplicate detection.
    $productIdCounts = @{}

    foreach ($record in $LocalRecords) {
        if (
            $record.Valid -and
            -not [string]::IsNullOrWhiteSpace(
                [string]$record.ProductId
            )
        ) {
            $id = [string]$record.ProductId

            if (-not $productIdCounts.ContainsKey($id)) {
                $productIdCounts[$id] = 0
            }

            $productIdCounts[$id]++
        }
    }

    # Query the same server URL only once.
    $remoteSizeCache = @{}

    $counter = 0

    foreach ($record in $LocalRecords) {
        $counter++

        $fileStatus = "Nicht pruefbar"
        $remoteSize = [long]-1
        $isDuplicate = $false

        if (
            $record.Valid -and
            -not [string]::IsNullOrWhiteSpace(
                [string]$record.ProductId
            )
        ) {
            $id = [string]$record.ProductId

            if (
                $productIdCounts.ContainsKey($id) -and
                $productIdCounts[$id] -gt 1
            ) {
                $isDuplicate = $true
            }

            $catalogEntry =
                Get-PreferredCatalogGameEntry `
                    -GameProductId $id `
                    -CatalogMaps $CatalogMaps

            if ($null -ne $catalogEntry) {
                $url = [string]$catalogEntry.Url

                if (
                    -not [string]::IsNullOrWhiteSpace(
                        $url
                    )
                ) {
                    if ($remoteSizeCache.ContainsKey($url)) {
                        $remoteSize =
                            [long]$remoteSizeCache[$url]
                    }
                    else {
                        $remoteSize =
                            Get-RemoteFileSizeBytes `
                                -Url $url

                        $remoteSizeCache[$url] =
                            [long]$remoteSize
                    }
                }
            }

            # Duplikat takes precedence in the table display.
            if ($isDuplicate) {
                $fileStatus = "Duplikat"
            }
            elseif ($remoteSize -ge 0) {
                if (
                    [long]$record.Length -eq
                    [long]$remoteSize
                ) {
                    $fileStatus = "OK"
                }
                else {
                    $fileStatus = "Falsche Groesse"
                }
            }
        }

        $record |
            Add-Member `
                -NotePropertyName "FileStatus" `
                -NotePropertyValue $fileStatus `
                -Force

        $record |
            Add-Member `
                -NotePropertyName "RemoteSizeBytes" `
                -NotePropertyValue ([long]$remoteSize) `
                -Force

        $record |
            Add-Member `
                -NotePropertyName "IsDuplicate" `
                -NotePropertyValue $isDuplicate `
                -Force
    }

    return $LocalRecords
}


# ============================================================================
# Handle orphaned temporary download files from a previous run
# ============================================================================

function Remove-StaleDownloadFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    $staleFiles = @(
        Get-ChildItem `
            -LiteralPath $Directory `
            -File `
            -Filter "*.download" `
            -ErrorAction SilentlyContinue |
        Sort-Object `
            -Property Length `
            -Descending
    )

    if ($staleFiles.Count -eq 0) {
        return
    }

    Write-Host ""
    Write-Warning (
        "{0} temporaere Download-Datei(en) eines vorherigen Durchlaufs " +
        "gefunden." -f
        $staleFiles.Count
    )

    foreach ($file in $staleFiles) {
        Write-Host ""
        Write-Host (
            "Temporäre Datei: {0}" -f
            $file.Name
        )

        Write-Host (
            "Dateigröße:      {0}" -f
            (Format-ByteSize -Bytes ([long]$file.Length))
        )

        $answer =
            Read-Host "Temporäre Datei löschen? [j/N]"

        if ($answer -match '^(?i:j|ja|y|yes)$') {
            try {
                Remove-Item `
                    -LiteralPath $file.FullName `
                    -Force `
                    -ErrorAction Stop

                Write-Host "Temporäre Datei gelöscht."
            }
            catch {
                Write-Warning (
                    "Temporäre Datei konnte nicht gelöscht werden: " +
                    $_.Exception.Message
                )
            }
        }
        else {
            Write-Host "Temporäre Datei bleibt erhalten."
        }
    }

    Write-Host ""
}


# ============================================================================
# Load data and resolve RSS -> GME
# ============================================================================

$catalog = Get-RavensburgerCatalog
$catalogMaps = Get-CatalogProductMaps -Catalog $catalog
$bookIndex = Get-RavensburgerBookIndex
$rssItems = @(Get-LibraryRssItems)

$libraryResolution =
    Resolve-LibraryRssToCatalog `
        -RssItems $rssItems `
        -BookIndex $bookIndex `
        -CatalogMaps $catalogMaps

$libraryEntries = @($libraryResolution.Entries)
$unresolvedRssItems = @($libraryResolution.Unresolved)

$libraryGameIds = @{}

foreach ($entry in $libraryEntries) {
    $libraryGameIds[[string]$entry.GameProductId] = $true
}

Write-Host ""
Write-Host (
    "{0} eindeutige GME Product-IDs aus dem RSS-Feed zugeordnet." -f
    $libraryGameIds.Count
)


# ============================================================================
# Load optional "mine" and "ignore" lists
# ============================================================================

$bookListsConfiguration =
    Get-BookListsConfiguration `
        -Path $MyBooksFile

$mineResolution =
    Resolve-ConfiguredBookListToGameIds `
        -List $bookListsConfiguration.Mine `
        -BookIndex $bookIndex `
        -CatalogMaps $catalogMaps `
        -Label "mine"

$ignoreResolution =
    Resolve-ConfiguredBookListToGameIds `
        -List $bookListsConfiguration.Ignore `
        -BookIndex $bookIndex `
        -CatalogMaps $catalogMaps `
        -Label "ignore"

$ownedGameIds =
    $mineResolution.GameIds

$ignoredGameIds =
    $ignoreResolution.GameIds

if ($unresolvedRssItems.Count -gt 0) {
    Write-Warning (
        "{0} RSS-Eintrag/Eintraege konnten nicht sicher einer GME Product-ID zugeordnet werden." -f
        $unresolvedRssItems.Count
    )

    $unresolvedRssItems |
        Format-Table `
            -Property Title, Isbn, Reason `
            -AutoSize
}


# ============================================================================
# Scan local files
# ============================================================================

$localRecords =
    @(
        Get-LocalGmeRecords `
            -Directory $ScriptDirectory `
            -CatalogMaps $catalogMaps
    )


# ============================================================================
# Optional: repair existing files
# ============================================================================

if ($fix) {
    $localRecords = @(
        Fix-LocalGmeFiles `
            -LocalRecords $localRecords `
            -CatalogMaps $catalogMaps `
            -Directory $ScriptDirectory
    )
}


# After an optional repair, determine the availability status of each local file.
# This status is used both for the table and for the
# deletion candidates in download mode.
foreach ($record in $localRecords) {
    $status =
        Get-LocalAvailabilityStatus `
            -Record $record `
            -OwnedGameIds $ownedGameIds `
            -IgnoredGameIds $ignoredGameIds `
            -LibraryGameIds $libraryGameIds

    $record |
        Add-Member `
            -NotePropertyName "AvailabilityStatus" `
            -NotePropertyValue $status `
            -Force
}


# ============================================================================
# Optional manual cleanup
# ============================================================================

if ($CleanupEnabled) {
    Invoke-ManualCleanup `
        -LocalRecords $localRecords `
        -Mode $CleanupMode

    # Re-read local inventory after possible deletions.
    $localRecords = @(
        Get-LocalGmeRecords `
            -Directory $ScriptDirectory `
            -CatalogMaps $catalogMaps
    )

    foreach ($record in $localRecords) {
        $status =
            Get-LocalAvailabilityStatus `
                -Record $record `
                -OwnedGameIds $ownedGameIds `
                -IgnoredGameIds $ignoredGameIds `
                -LibraryGameIds $libraryGameIds

        $record |
            Add-Member `
                -NotePropertyName "AvailabilityStatus" `
                -NotePropertyValue $status `
                -Force
    }

    # If cleanup was invoked as the only mode, exit after
    # cleanup. When combined with -download, the
    # download mode continues with the cleaned-up inventory.
    if (-not $download) {
        exit 0
    }
}


# Rebuild after an optional repair because -fix can delete or
# reload files. This ensures download mode always works with the
# actual current inventory.
$localGameIds = @{}

foreach ($record in $localRecords) {
    if (
        $record.Valid -and
        -not [string]::IsNullOrWhiteSpace($record.ProductId)
    ) {
        $localGameIds[[string]$record.ProductId] = $true
    }
}


# ============================================================================
# DOWNLOAD MODE
# ============================================================================

if ($download) {
    # Handle orphaned temporary files from previous download runs first.
    Remove-StaleDownloadFiles `
        -Directory $ScriptDirectory

    Write-Host ""
    Write-Host "============================================================"
    Write-Host (
        " Download-Modus ({0})" -f
        $DownloadMode
    )
    Write-Host "============================================================"
    Write-Host ""

    if ($unresolvedRssItems.Count -gt 0) {
        Write-Warning (
            "{0} RSS-Eintrag/Eintraege konnten nicht sicher einer GME " +
            "Product-ID zugeordnet werden. Der Download wird mit den erfolgreich " +
            "zugeordneten Eintraegen fortgesetzt." -f
            $unresolvedRssItems.Count
        )

        Write-Host ""
    }

    # In download mode, files can optionally be filtered by the Ravensburger
    # age recommendation.
    #
    # Rules:
    # - both ageFrom and ageTo are missing -> always include
    # - only ageFrom is missing          -> ageFrom = 0
    # - only ageTo is missing            -> ageTo   = 99
    # - otherwise                        -> ageFrom <= age <= ageTo
    #
    # Without -age, NO age filtering is applied.
    $ignoredLibraryEntries = @(
        $libraryEntries |
        Where-Object {
            $ignoredGameIds.ContainsKey(
                [string]$_.GameProductId
            )
        }
    )

    $eligibleLibraryEntries = @(
        $libraryEntries |
        Where-Object {
            $gameId =
                [string]$_.GameProductId

            if ($ignoredGameIds.ContainsKey($gameId)) {
                return $false
            }

            Test-DownloadAgeMatch `
                -Entry $_ `
                -RequestedAge $age `
                -FilterEnabled $AgeFilterEnabled
        }
    )

    if ($ignoredLibraryEntries.Count -gt 0) {
        Write-Host (
            "{0} RSS-Eintrag/Eintraege sind in 'ignore' gelistet und " +
            "werden nicht heruntergeladen." -f
            $ignoredLibraryEntries.Count
        )

        Write-Host ""
    }

    if ($AgeFilterEnabled) {
        Write-Host (
            "Altersfilter aktiv: {0} Jahr(e)." -f
            $age
        )

        Write-Host (
            "{0} von {1} erfolgreich zugeordneten RSS-Eintraegen " +
            "erfuellen den Altersfilter." -f
            $eligibleLibraryEntries.Count,
            $libraryEntries.Count
        )

        Write-Host ""
    }

    # Missing GME files are determined exclusively by the GME Product-ID,
    # NOT by the filename.
    $toDownload = @(
        $eligibleLibraryEntries |
        Where-Object {
            -not $localGameIds.ContainsKey(
                [string]$_.GameProductId
            )
        } |
        Sort-Object `
            -Property ProductName, GameProductId
    )

    if ($toDownload.Count -eq 0) {
        if ($AgeFilterEnabled) {
            Write-Host (
                "Von den erfolgreich zugeordneten RSS-Eintraegen, die den " +
                "Altersfilter erfuellen, fehlt keine tiptoi-GME-Datei im " +
                "aktuellen Ordner."
            )
        }
        else {
            Write-Host (
                "Von den erfolgreich zugeordneten RSS-Eintraegen fehlt keine " +
                "tiptoi-GME-Datei im aktuellen Ordner."
            )
        }

        if ($unresolvedRssItems.Count -gt 0) {
            Write-Warning (
                "{0} RSS-Eintrag/Eintraege konnten nicht sicher einer GME " +
                "Product-ID zugeordnet werden und wurden daher beim Download " +
                "nicht beruecksichtigt." -f
                $unresolvedRssItems.Count
            )
        }

        exit 0
    }

    Write-Host (
        "{0} GME-Datei(en) fehlen und sollen heruntergeladen werden." -f
        $toDownload.Count
    )

    Write-Host ""
    Write-Host "Remote-Dateigroessen werden ermittelt ..."

    $sizedDownloads = @()
    $sizeUnknown = @()

    $sizeCounter = 0

    foreach ($entry in $toDownload) {
        $sizeCounter++

        Write-Host (
            "[{0}/{1}] {2}" -f
            $sizeCounter,
            $toDownload.Count,
            $entry.ProductName
        )

        $remoteSize =
            Get-RemoteFileSizeBytes `
                -Url $entry.Url

        if ($remoteSize -lt 0) {
            $sizeUnknown += $entry
            continue
        }

        $sizedDownloads += [PSCustomObject]@{
            Entry     = $entry
            SizeBytes = [long]$remoteSize
        }
    }

    if ($sizeUnknown.Count -gt 0) {
        Write-Warning (
            "Die Groesse von {0} Download(s) konnte nicht vorab bestimmt werden. " +
            "Da die Gesamtgroesse vor dem Download sicher bekannt sein soll, " +
            "wird nicht heruntergeladen." -f
            $sizeUnknown.Count
        )

        $sizeUnknown |
            Select-Object `
                ProductName,
                GameProductId,
                Url |
            Format-Table -AutoSize

        exit 3
    }

    # --------------------------------------------------------------------
    # Confirm download candidates.
    #
    # Default:
    #   -download or -download Ask
    #       -> ask for each candidate individually
    #
    # Optional:
    #   -download All
    #       -> accept all candidates without prompting
    #
    # IMPORTANT:
    # This selection takes place BEFORE calculating the required
    # total storage space.
    # --------------------------------------------------------------------

    if ($DownloadMode -eq "Ask") {
        Write-Host ""
        Write-Host "Download-Kandidaten werden einzeln abgefragt ..."
        Write-Host ""

        $selectedDownloads = @()

        foreach ($item in $sizedDownloads) {
            $entry = $item.Entry

            $isbnText =
                [string]$entry.Isbn

            if ([string]::IsNullOrWhiteSpace($isbnText)) {
                $isbnText = "keine Angabe"
            }

            $ageRecommendation =
                Get-AgeRecommendationText `
                    -Entry $entry

            Write-Host (
                "Titel:            {0}" -f
                $entry.ProductName
            )

            Write-Host (
                "ISBN:             {0}" -f
                $isbnText
            )

            Write-Host (
                "GME Product-ID:   {0}" -f
                $entry.GameProductId
            )

            Write-Host (
                "Altersempfehlung: {0}" -f
                $ageRecommendation
            )

            Write-Host (
                "Dateigroesse:     {0}" -f
                (Format-ByteSize -Bytes $item.SizeBytes)
            )

            $answer =
                Read-Host "Herunterladen? [j/N]"

            if ($answer -match '^(?i:j|ja|y|yes)$') {
                $selectedDownloads += $item

                Write-Host "Zum Download ausgewaehlt."
            }
            else {
                Write-Host "Wird nicht heruntergeladen."
            }

            Write-Host ""
        }

        $sizedDownloads =
            @($selectedDownloads)

        if ($sizedDownloads.Count -eq 0) {
            Write-Host (
                "Es wurde kein Download-Kandidat ausgewaehlt. " +
                "Es wird nichts heruntergeladen."
            )

            exit 0
        }

        Write-Host (
            "{0} Datei(en) wurden fuer den Download ausgewaehlt." -f
            $sizedDownloads.Count
        )
    }
    else {
        Write-Host ""
        Write-Host (
            "Download-Modus All: Alle {0} Kandidaten werden ohne " +
            "Einzelabfrage heruntergeladen." -f
            $sizedDownloads.Count
        )
    }

    $totalDownloadBytes = [long](
        (
            $sizedDownloads |
            Measure-Object `
                -Property SizeBytes `
                -Sum
        ).Sum
    )

    Write-Host ""
    Write-Host (
        "Gesamtgroesse der Downloads: {0}" -f
        (Format-ByteSize -Bytes $totalDownloadBytes)
    )

    $freeBytes =
        Get-FreeSpaceBytes `
            -Path $ScriptDirectory

    Write-Host (
        "Freier Speicherplatz:          {0}" -f
        (Format-ByteSize -Bytes $freeBytes)
    )

    # --------------------------------------------------------------------
    # If insufficient space is available:
    # offer local GMEs that are NOT listed in the RSS feed in descending
    # size order.
    # --------------------------------------------------------------------

    if ($freeBytes -lt $totalDownloadBytes) {
        Write-Host ""
        Write-Warning "Der freie Speicherplatz reicht nicht aus."

        if ($unresolvedRssItems.Count -gt 0) {
            Write-Warning (
                "ACHTUNG: {0} RSS-Eintrag/Eintraege konnten nicht sicher einer " +
                "GME Product-ID zugeordnet werden. Deshalb koennen im folgenden " +
                "Loeschdialog auch lokale GME-Dateien angeboten werden, die " +
                "moeglicherweise zu einem dieser nicht zugeordneten RSS-Titel " +
                "gehoeren. Bitte gleichen Sie jeden Loeschkandidaten manuell mit " +
                "den zuvor ausgegebenen, nicht zugeordneten Titeln/ISBN-Angaben ab, " +
                "bevor Sie das Loeschen bestaetigen." -f
                $unresolvedRssItems.Count
            )

            Write-Host ""
            Write-Host "Nicht sicher zugeordnete RSS-Eintraege:"
            $unresolvedRssItems |
                Format-Table `
                    -Property Title, Isbn, Reason `
                    -AutoSize

            Write-Host ""
        }

        # To free storage space, files with status
        # "Unbekannt" or "Ignoriert" may be offered.
        #
        # An explicitly ignored book may therefore be deleted even
        # if it is also contained in the library RSS feed.
        # "Besitz" and "Buecherei" are not offered.
        $deleteCandidates = @(
            $localRecords |
            Where-Object {
                $_.Valid -and
                (
                    $_.AvailabilityStatus -eq "Unbekannt" -or
                    $_.AvailabilityStatus -eq "Ignoriert"
                )
            } |
            Sort-Object `
                -Property Length `
                -Descending
        )

        # The list is repeated until sufficient storage space
        # is available. Files already deleted are skipped in later
        # passes. There is no automatic abort
        # at this point; the user can abort with Ctrl+C.
        $deletePass = 0

        :StorageCleanup while ($true) {
            $freeBytes =
                Get-FreeSpaceBytes `
                    -Path $ScriptDirectory

            if ($freeBytes -ge $totalDownloadBytes) {
                Write-Host ""
                Write-Host (
                    "Genuegend freier Speicherplatz vorhanden. " +
                    "Die Loeschabfrage wird beendet und der Download beginnt."
                )

                break StorageCleanup
            }

            $remainingCandidates = @(
                $deleteCandidates |
                Where-Object {
                    Test-Path -LiteralPath $_.FullName
                }
            )

            if ($remainingCandidates.Count -eq 0) {
                Write-Host ""
                Write-Error (
                    "Der freie Speicherplatz reicht weiterhin nicht aus und " +
                    "es sind keine verbleibenden Loeschkandidaten mit Status " +
                    "'Unbekannt' oder 'Ignoriert' vorhanden. Der Download wird abgebrochen."
                )

                exit 5
            }

            $deletePass++

            if ($deletePass -gt 1) {
                Write-Host ""
                Write-Warning (
                    "Der freie Speicherplatz reicht weiterhin nicht aus. " +
                    "Die verbleibenden Loeschkandidaten werden erneut " +
                    "durchgegangen. Abbruch mit Ctrl+C."
                )
            }

            foreach ($candidate in $remainingCandidates) {
                $freeBytes =
                    Get-FreeSpaceBytes `
                        -Path $ScriptDirectory

                if ($freeBytes -ge $totalDownloadBytes) {
                    break
                }

                # If the file was already deleted in a previous pass,
                # do not offer it again.
                if (-not (Test-Path -LiteralPath $candidate.FullName)) {
                    continue
                }

                $missingBytes =
                    $totalDownloadBytes - $freeBytes

                Write-Host ""
                Write-Host (
                    "Noch benoetigter Speicherplatz: {0}" -f
                    (Format-ByteSize -Bytes $missingBytes)
                )

                Write-Host (
                    "Status {0} : {1}" -f
                    $candidate.AvailabilityStatus,
                    $candidate.FileName
                )

                Write-Host (
                    "Titel:             {0}" -f
                    $candidate.Title
                )

                Write-Host (
                    "GME Product-ID:    {0}" -f
                    $candidate.ProductId
                )

                $candidateCatalogEntry =
                    Get-PreferredCatalogGameEntry `
                        -GameProductId ([string]$candidate.ProductId) `
                        -CatalogMaps $catalogMaps

                $ageRecommendation =
                    Get-AgeRecommendationText `
                        -Entry $candidateCatalogEntry

                Write-Host (
                    "Altersempfehlung:  {0}" -f
                    $ageRecommendation
                )

                Write-Host (
                    "Dateigroesse:      {0}" -f
                    (Format-ByteSize -Bytes $candidate.Length)
                )

                $answer =
                    Read-Host "Datei loeschen? [j/N]"

                if ($answer -match '^(?i:j|ja|y|yes)$') {
                    try {
                        Remove-Item `
                            -LiteralPath $candidate.FullName `
                            -Force

                        Write-Host "Datei geloescht."

                        # Re-check free space after EVERY deletion.
                        $freeBytes =
                            Get-FreeSpaceBytes `
                                -Path $ScriptDirectory

                        Write-Host (
                            "Freier Speicherplatz jetzt: {0}" -f
                            (Format-ByteSize -Bytes $freeBytes)
                        )

                        if ($freeBytes -ge $totalDownloadBytes) {
                            Write-Host ""
                            Write-Host (
                                "Genuegend freier Speicherplatz erreicht. " +
                                "Die Loeschabfrage wird beendet und der " +
                                "Download beginnt."
                            )

                            break StorageCleanup
                        }
                    }
                    catch {
                        Write-Warning (
                            "Datei konnte nicht geloescht werden: " +
                            $_.Exception.Message
                        )
                    }
                }
                else {
                    Write-Host "Datei bleibt erhalten."
                }
            }

            $freeBytes =
                Get-FreeSpaceBytes `
                    -Path $ScriptDirectory

            if ($freeBytes -lt $totalDownloadBytes) {
                Write-Host ""
                Write-Warning (
                    "Nach diesem Durchlauf fehlt weiterhin Speicherplatz. " +
                    "Die Liste der verbleibenden Loeschkandidaten wird " +
                    "erneut gestartet. Abbruch mit Ctrl+C."
                )
            }
        }
    }

    # --------------------------------------------------------------------
    # Downloads
    # --------------------------------------------------------------------

    Write-Host ""
    Write-Host "Downloads werden gestartet ..."
    Write-Host ""

    $downloadCounter = 0
    $downloadErrors = @()

    foreach ($item in $sizedDownloads) {
        $downloadCounter++
        $entry = $item.Entry

        Write-Host (
            "[{0}/{1}] {2}" -f
            $downloadCounter,
            $sizedDownloads.Count,
            $entry.ProductName
        )

        Write-Host (
            "    ISBN           : {0}" -f
            $entry.Isbn
        )

        Write-Host (
            "    GME Product-ID : {0}" -f
            $entry.GameProductId
        )

        Write-Host (
            "    Groesse        : {0}" -f
            (Format-ByteSize -Bytes $item.SizeBytes)
        )

        try {
            $savedPath =
                Download-GmeEntry `
                    -Entry $entry `
                    -Directory $ScriptDirectory

            Write-Host (
                "    Gespeichert    : {0}" -f
                ([System.IO.Path]::GetFileName($savedPath))
            )
        }
        catch {
            $downloadErrors += [PSCustomObject]@{
                Title     = $entry.ProductName
                ProductId = $entry.GameProductId
                Error     = $_.Exception.Message
            }

            Write-Warning (
                "Download fehlgeschlagen: " +
                $_.Exception.Message
            )
        }

        Write-Host ""
    }

    if ($downloadErrors.Count -gt 0) {
        Write-Warning (
            "{0} Download(s) sind fehlgeschlagen." -f
            $downloadErrors.Count
        )

        $downloadErrors |
            Format-Table `
                -Property Title, ProductId, Error `
                -AutoSize

        exit 6
    }

    Write-Host (
        "Fertig. {0} GME-Datei(en) wurden heruntergeladen." -f
        $sizedDownloads.Count
    )

    exit 0
}


# ============================================================================
# Automatically check file size + duplicates before every table listing
# ============================================================================

$localRecords = @(
    Add-LocalFileCheckStatus `
        -LocalRecords $localRecords `
        -CatalogMaps $catalogMaps
)


# ============================================================================
# NORMAL MODE: table of existing files
# ============================================================================

$results = @()

foreach ($record in $localRecords) {
    $availability =
        [string]$record.AvailabilityStatus

    if ([string]::IsNullOrWhiteSpace($availability)) {
        $availability = "Unbekannt"
    }

    # Generate umlauts only here so that the script file remains robust under
    # Windows PowerShell 5.1.
    if ($availability -eq "Buecherei") {
        $availability = "Bücherei"
    }

    $fileStatus =
        [string]$record.FileStatus

    if ([string]::IsNullOrWhiteSpace($fileStatus)) {
        $fileStatus = "Nicht pruefbar"
    }

    # Desired German display.
    if ($fileStatus -eq "Falsche Groesse") {
        $fileStatus = "Falsche Größe"
    }
    elseif ($fileStatus -eq "Nicht pruefbar") {
        $fileStatus = "Nicht prüfbar"
    }

    $results += [PSCustomObject]@{
        "ID"             = $record.ProductId
        "Dateiname"      = $record.FileName
        "Titel"          = $record.Title
        "Verfügbarkeit" = $availability
        "Status"         = $fileStatus
        "Größe in MB"   = [math]::Round(
            $record.Length / 1MB,
            2
        )
    }
}

if ($excludeListed) {
    # "excludeListed" continues to show only status "Unbekannt".
    # Explicitly ignored files therefore remain separate from this view.
    $results = @(
        $results |
        Where-Object {
            $_."Verfügbarkeit" -eq "Unbekannt"
        }
    )
}

# ============================================================================
# Sorting of table listing
#
# Default:
#     -sort Id
#         GME Product-ID numerically ascending
#
# Optional:
#     -sort Size
#         File size descending
# ============================================================================

if ($sort -eq "Size") {
    $results = @(
        $results |
        Sort-Object `
            -Property "Größe in MB" `
            -Descending
    )
}
else {
    # Sort ID numerically, not lexicographically.
    # Unreadable/empty IDs are placed at the end.
    $results = @(
        $results |
        Sort-Object `
            -Property @{
                Expression = {
                    $numericId = [long]0

                    if (
                        [long]::TryParse(
                            [string]$_.ID,
                            [ref]$numericId
                        )
                    ) {
                        return $numericId
                    }

                    return [long]::MaxValue
                }
            }, @{
                Expression = {
                    [string]$_.ID
                }
            }
    )
}

Write-Host ""
Write-Host "=========================================================================="
Write-Host " Ergebnis"
Write-Host "=========================================================================="
Write-Host ""

if ($results.Count -eq 0) {
    if ($excludeListed) {
        Write-Host (
            "Keine Dateien mit Verfuegbarkeit Unbekannt gefunden."
        )
    }
    else {
        Write-Host "Keine GME-Dateien im Script-Verzeichnis gefunden."
    }
}
else {
    $results |
        Format-Table `
            -Property `
                "ID",
                "Dateiname",
                "Titel",
                "Verfügbarkeit",
                "Status",
                "Größe in MB" `
            -AutoSize
}
