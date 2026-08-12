function Build-EntraWordReport {
    <#
        Builds a Word system-documentation report (.docx) from an export, in pure
        PowerShell - it writes the Office Open XML parts and zips them, so it needs
        neither Word installed nor any extra module.

        This first milestone lays down the whole document skeleton (cover page,
        table of contents, and every section heading so the TOC is complete) and
        fills section 1 (Inledning / Versionshantering). The remaining sections are
        placeholders that we fill in area by area.

        Metadata that isn't in Microsoft Graph (author, customer, contacts, logos,
        version) comes from Entra\Config\WordReportProfile.psd1; anything blank is
        shown as a [placeholder] so it's easy to complete in Word.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ExportsRoot,
        [string]$SourceFolder,
        [string]$Language = 'sv',
        [string]$ConsultantLogo,
        [string]$CustomerLogo
    )

    $src = if ($SourceFolder) { $SourceFolder } else { Resolve-AutodocSource -ExportsRoot $ExportsRoot }
    if (-not $src) { Write-Host "No export to build a Word report from." -ForegroundColor Yellow; return }

    # ---- profile (metadata that isn't in Graph) ----
    $entraRoot = Split-Path (Split-Path (Split-Path $PSScriptRoot))   # Functions\Reports\Word -> Entra
    $profilePath = Join-Path $entraRoot 'Config\WordReportProfile.psd1'
    $prof = @{}
    if (Test-Path $profilePath) { try { $prof = Import-PowerShellDataFile -Path $profilePath } catch { $prof = @{} } }
    function PV($k, $ph) { $v = "$($prof[$k])"; if ($v) { $v } else { "[$ph]" } }

    $consultant = PV 'ConsultantCompany' 'Konsultbolag'
    $author     = PV 'Author'            'Utarbetad av'
    $customer   = PV 'CustomerName'      'Kundens namn'
    $docId      = if ("$($prof.DocumentId)") { "$($prof.DocumentId)" } else { 'Systemdokumentation Entra ID' }
    $version    = if ("$($prof.Version)")    { "$($prof.Version)" }    else { '1.0' }
    $firstChg   = if ("$($prof.FirstChange)"){ "$($prof.FirstChange)" }else { 'Första utkast' }
    $today      = Get-Date -Format 'yyyy-MM-dd'

    # logos: GUI parameter -> cached logo (files\cache\wordTemplates\template\img) -> profile -> text placeholder
    $logoImgDir = Join-Path $entraRoot 'files\cache\wordTemplates\template\img'
    function CachedLogo($base) { if (Test-Path $logoImgDir) { $f = Get-ChildItem -Path $logoImgDir -Filter "$base.*" -ErrorAction SilentlyContinue | Select-Object -First 1; if ($f) { return $f.FullName } } ; return '' }
    $consultantLogo = if ($ConsultantLogo -and (Test-Path $ConsultantLogo)) { $ConsultantLogo } elseif (CachedLogo 'companylogo') { CachedLogo 'companylogo' } elseif ("$($prof.ConsultantLogo)" -and (Test-Path "$($prof.ConsultantLogo)")) { "$($prof.ConsultantLogo)" } else { '' }
    $customerLogo   = if ($CustomerLogo   -and (Test-Path $CustomerLogo))   { $CustomerLogo }   elseif (CachedLogo 'clientlogo')  { CachedLogo 'clientlogo' }  elseif ("$($prof.CustomerLogo)"   -and (Test-Path "$($prof.CustomerLogo)"))   { "$($prof.CustomerLogo)" }   else { '' }

    $tenant = @(Get-AutodocExportData -ExportFolder $src -BaseName 'TenantInformation') | Select-Object -First 1

    # ---- OOXML helpers ----
    function XmlEsc([string]$s) { if ($null -eq $s) { return '' }; $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' }
    function Par {
        param([string]$Text = '', [string]$Style = '', [string]$Align = '', [switch]$Bold, [int]$Size = 0)
        $ppr = ''; if ($Style) { $ppr += "<w:pStyle w:val='$Style'/>" }; if ($Align) { $ppr += "<w:jc w:val='$Align'/>" }
        $rpr = ''; if ($Bold) { $rpr += '<w:b/>' }; if ($Size) { $rpr += "<w:sz w:val='$Size'/><w:szCs w:val='$Size'/>" }
        $ppr = if ($ppr) { "<w:pPr>$ppr</w:pPr>" } else { '' }; $rpr = if ($rpr) { "<w:rPr>$rpr</w:rPr>" } else { '' }
        "<w:p>$ppr<w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc $Text)</w:t></w:r></w:p>"
    }
    function H1($t) { Par -Text $t -Style 'Heading1' }
    function H2($t) { Par -Text $t -Style 'Heading2' }
    function PageBreak { "<w:p><w:r><w:br w:type='page'/></w:r></w:p>" }
    function Tbl {
        param([object[]]$Rows, [int[]]$ColW, [switch]$Header, [switch]$BoldFirstCol)
        $total = ($ColW | Measure-Object -Sum).Sum
        $grid = ($ColW | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
        $trs = ''
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $tcs = ''
            for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
                $hdr = ($Header -and $r -eq 0)
                $shd = if ($hdr) { "<w:shd w:val='clear' w:color='auto' w:fill='4472C4'/>" } else { '' }
                $b = ($hdr -or ($BoldFirstCol -and $c -eq 0))
                $col = if ($hdr) { "<w:color w:val='FFFFFF'/>" } else { '' }
                $rpr = if ($b -or $col) { "<w:rPr>$(if($b){'<w:b/>'})$col</w:rPr>" } else { '' }
                $tcs += "<w:tc><w:tcPr><w:tcW w:w='$($ColW[$c])' w:type='dxa'/>$shd</w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc "$($Rows[$r][$c])")</w:t></w:r></w:p></w:tc>"
            }
            $trs += "<w:tr>$tcs</w:tr>"
        }
        "<w:tbl><w:tblPr><w:tblW w:w='$total' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:left w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:right w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$grid</w:tblGrid>$trs</w:tbl>"
    }

    # ---- image helpers (logos) ----
    $images = [System.Collections.Generic.List[object]]::new()
    $imgSeq = 0
    function Get-ImageSize($path) {
        try {
            $b = [System.IO.File]::ReadAllBytes($path)
            if ($b.Length -gt 24 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50) {                     # PNG
                return @{ W = (([int]$b[16] -shl 24) -bor ([int]$b[17] -shl 16) -bor ([int]$b[18] -shl 8) -bor [int]$b[19]); H = (([int]$b[20] -shl 24) -bor ([int]$b[21] -shl 16) -bor ([int]$b[22] -shl 8) -bor [int]$b[23]) }
            }
            if ($b.Length -gt 4 -and $b[0] -eq 0xFF -and $b[1] -eq 0xD8) {                        # JPEG
                $i = 2
                while ($i -lt $b.Length - 8) {
                    if ($b[$i] -ne 0xFF) { $i++; continue }
                    $m = $b[$i + 1]
                    if ($m -ge 0xC0 -and $m -le 0xC3) { return @{ W = (([int]$b[$i + 7] -shl 8) -bor [int]$b[$i + 8]); H = (([int]$b[$i + 5] -shl 8) -bor [int]$b[$i + 6]) } }
                    $i += 2 + (([int]$b[$i + 2] -shl 8) -bor [int]$b[$i + 3])
                }
            }
        } catch {}
        return @{ W = 0; H = 0 }
    }
    function LogoDrawing($path, $maxWidthInch) {
        if (-not $path -or -not (Test-Path $path)) { return $null }
        $script:imgSeq++
        $ext = ([System.IO.Path]::GetExtension($path)).TrimStart('.').ToLower(); if ($ext -eq 'jpeg') { $ext = 'jpg' }
        $mediaName = "image$($script:imgSeq).$ext"; $relId = "rIdImg$($script:imgSeq)"
        $d = Get-ImageSize $path
        $wIn = $maxWidthInch; $hIn = if ($d.W -gt 0) { [math]::Round($maxWidthInch * $d.H / $d.W, 3) } else { [math]::Round($maxWidthInch * 0.4, 3) }
        $wEmu = [int]($wIn * 914400); $hEmu = [int]($hIn * 914400)
        $images.Add(@{ Path = $path; MediaName = $mediaName; RelId = $relId })
        $id = $script:imgSeq
        "<w:drawing><wp:inline distT='0' distB='0' distL='0' distR='0' xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'><wp:extent cx='$wEmu' cy='$hEmu'/><wp:docPr id='$id' name='Logo$id'/><a:graphic xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/picture'><pic:pic xmlns:pic='http://schemas.openxmlformats.org/drawingml/2006/picture'><pic:nvPicPr><pic:cNvPr id='$id' name='Logo$id'/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed='$relId'/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='$wEmu' cy='$hEmu'/></a:xfrm><a:prstGeom prst='rect'><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"
    }

    function Get-MapPng($namePattern, $builder, $maxWidthInch) {
        $svgDir = Join-Path $src 'Report\svg'
        $find = { Get-ChildItem -Path $svgDir -Filter "*$namePattern*.svg" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime | Select-Object -Last 1 }
        $svg = & $find
        if (-not $svg -and $builder) { try { & $builder | Out-Null } catch { } ; $svg = & $find }
        if (-not $svg) { return $null }
        $png = Join-Path ([System.IO.Path]::GetTempPath()) ("wrmap_" + [guid]::NewGuid().ToString('N') + '.png')
        if (Convert-SvgToPng -SvgPath $svg.FullName -PngPath $png) { return (LogoDrawing $png $maxWidthInch) }
        return $null
    }

    # ---- body ----
    $sb = [System.Text.StringBuilder]::new()

    # cover
    [void]$sb.Append((Tbl -BoldFirstCol -Rows @(@('Dokumenttyp', 'Systemdokumentation'), @('Område', 'Entra ID')) -ColW @(1900, 2600)))
    [void]$sb.Append((Par ''))
    $cLogo = LogoDrawing $consultantLogo 2.2
    if ($cLogo) { [void]$sb.Append("<w:p><w:r>$cLogo</w:r></w:p>") } else { [void]$sb.Append((Par -Text "[$consultant logo]")) }
    1..4 | ForEach-Object { [void]$sb.Append((Par '')) }
    [void]$sb.Append((Par -Text $docId -Style 'Title' -Align 'center'))
    [void]$sb.Append((Par -Text $customer -Align 'center' -Size 36))
    1..7 | ForEach-Object { [void]$sb.Append((Par '')) }
    $kLogo = LogoDrawing $customerLogo 2.6
    if ($kLogo) { [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$kLogo</w:r></w:p>") } else { [void]$sb.Append((Par -Text '[Kundens logo]' -Align 'center')) }
    [void]$sb.Append((PageBreak))

    # table of contents (populated when the field is updated in Word)
    [void]$sb.Append((Par -Text 'Innehållsförteckning' -Style 'TOCHeading'))
    [void]$sb.Append("<w:p><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> TOC \o &quot;1-3&quot; \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType='separate'/></w:r><w:r><w:t xml:space='preserve'>Uppdatera fältet (markera allt, F9) i Word f&#246;r att fylla inneh&#229;llsf&#246;rteckningen.</w:t></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r></w:p>")
    [void]$sb.Append((PageBreak))

    # 1 Inledning
    [void]$sb.Append((H1 'Inledning'))
    [void]$sb.Append((H2 'Uppdatering av dokument'))
    [void]$sb.Append((Par 'Systemansvarig ansvarar för att detta dokument är aktuellt och relevant. En årlig översyn ska utföras.'))
    [void]$sb.Append((H2 'Versionshantering'))
    [void]$sb.Append((Tbl -Header -Rows @(@('Version', 'Datum', 'Förändrad av', 'Förändring'), @($version, $today, $author, $firstChg)) -ColW @(1200, 1800, 3000, 3000)))

    # 2 Förvaltning och support
    [void]$sb.Append((H1 'Förvaltning och support'))
    [void]$sb.Append((H2 "Kontakt $consultant"))
    [void]$sb.Append((Par (PV 'ContactConsultant' 'Kontaktuppgifter till konsultbolaget')))
    [void]$sb.Append((H2 'Kontaktpersoner kund'))
    [void]$sb.Append((Par (PV 'ContactCustomer' 'Kontaktpersoner hos kunden')))

    # 3 Allmän systembeskrivning
    [void]$sb.Append((H1 'Allmän systembeskrivning'))
    [void]$sb.Append((Par '[Innehåll kommer]'))

    # 4 Systemkarta
    [void]$sb.Append((H1 'Systemkarta'))
    [void]$sb.Append((Par 'Översiktskarta över tenanten (hub). Detaljerade områdeskartor finns i avsnitt 5 samt som separata SVG- och PDF-filer.'))
    $hubDraw = Get-MapPng 'Hub' { Build-OverviewMapSvg -SourceFolder $src -Style hub } 6.3
    if ($hubDraw) { [void]$sb.Append("<w:p><w:pPr><w:jc w:val='center'/></w:pPr><w:r>$hubDraw</w:r></w:p>") }
    else { [void]$sb.Append((Par '[Hub-karta kunde inte skapas - kontrollera att SVG-kartorna byggs]')) }

    # 5 Systemkonfiguration / inställningar
    [void]$sb.Append((H1 'Systemkonfiguration / inställningar'))
    foreach ($a in 'Users', 'Groups', 'Licenses', 'Departments & offices', 'Conditional Access', 'App registrations', 'Enterprise apps', 'Domains', 'Self-service password reset (SSPR)', 'Service limits & recommendations') {
        [void]$sb.Append((H2 $a)); [void]$sb.Append((Par '[Innehåll kommer]'))
    }

    # 6-8
    [void]$sb.Append((H1 'Behörighetsgrupper / RBAC (roller)')); [void]$sb.Append((Par '[Innehåll kommer]'))
    [void]$sb.Append((H1 'Logghantering')); [void]$sb.Append((Par '[Innehåll kommer]'))
    [void]$sb.Append((H1 'Överlämning')); [void]$sb.Append((Par '[Innehåll kommer]'))

    $sectPr = "<w:sectPr><w:footerReference w:type='default' r:id='rId3'/><w:pgSz w:w='11906' w:h='16838'/><w:pgMar w:top='1417' w:right='1417' w:bottom='1417' w:left='1417' w:header='708' w:footer='708' w:gutter='0'/></w:sectPr>"
    $bodyXml = $sb.ToString() + $sectPr

    # ---- fixed parts ----
    $documentXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n" +
    "<w:document xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><w:body>$bodyXml</w:body></w:document>"

    # footer with the metadata table + page number
    $pageField = "<w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r><w:r><w:t xml:space='preserve'> (</w:t></w:r><w:r><w:fldChar w:fldCharType='begin'/></w:r><w:r><w:instrText xml:space='preserve'> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType='end'/></w:r><w:r><w:t>)</w:t></w:r></w:p>"
    function FtrCell($w, $text, $bold) {
        $rpr = if ($bold) { '<w:rPr><w:b/><w:sz w:val=''16''/></w:rPr>' } else { '<w:rPr><w:sz w:val=''16''/></w:rPr>' }
        "<w:tc><w:tcPr><w:tcW w:w='$w' w:type='dxa'/></w:tcPr><w:p><w:pPr><w:spacing w:after='0'/></w:pPr><w:r>$rpr<w:t xml:space='preserve'>$(XmlEsc "$text")</w:t></w:r></w:p></w:tc>"
    }
    $fw = @(2400, 2400, 1500, 1000, 1700)
    $ftrHead = '<w:tr>' + (FtrCell $fw[0] 'Utarbetad av' $true) + (FtrCell $fw[1] 'Dokument-ID' $true) + (FtrCell $fw[2] 'Datum' $true) + (FtrCell $fw[3] 'Version' $true) + (FtrCell $fw[4] 'Sida' $true) + '</w:tr>'
    $sidaCell = "<w:tc><w:tcPr><w:tcW w:w='$($fw[4])' w:type='dxa'/></w:tcPr>$pageField</w:tc>"
    $ftrVal = '<w:tr>' + (FtrCell $fw[0] $author $false) + (FtrCell $fw[1] $docId $false) + (FtrCell $fw[2] $today $false) + (FtrCell $fw[3] $version $false) + $sidaCell + '</w:tr>'
    $ftrGrid = ($fw | ForEach-Object { "<w:gridCol w:w='$_'/>" }) -join ''
    $footerTbl = "<w:tbl><w:tblPr><w:tblW w:w='$(($fw|Measure-Object -Sum).Sum)' w:type='dxa'/><w:tblBorders><w:top w:val='single' w:sz='4' w:color='BFBFBF'/><w:bottom w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideH w:val='single' w:sz='4' w:color='BFBFBF'/><w:insideV w:val='single' w:sz='4' w:color='BFBFBF'/></w:tblBorders></w:tblPr><w:tblGrid>$ftrGrid</w:tblGrid>$ftrHead$ftrVal</w:tbl>"
    $footerXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<w:ftr xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`">$footerTbl</w:ftr>"

    $contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>
'@
    $relsRoot = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
'@
    $imgRels = ''
    foreach ($im in $images) { $imgRels += "<Relationship Id=`"$($im.RelId)`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`" Target=`"media/$($im.MediaName)`"/>" }
    $relsDoc = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>`n<Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/><Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings`" Target=`"settings.xml`"/><Relationship Id=`"rId3`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer`" Target=`"footer1.xml`"/><Relationship Id=`"rId4`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering`" Target=`"numbering.xml`"/>$imgRels</Relationships>"
    $settingsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:updateFields w:val="true"/></w:settings>
'@
    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="sv-SE"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="6AA84F"/></w:pBdr></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="2"/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="404040"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:pPr><w:outlineLvl w:val="9"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:style></w:styles>
'@

    $numberingXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:suff w:val="space"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:suff w:val="space"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:suff w:val="space"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>
'@

    # ---- assemble + zip ----
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("wr_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp '_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp 'word') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp 'word\_rels') -Force | Out-Null

    Set-Content -LiteralPath (Join-Path $tmp '[Content_Types].xml') -Value $contentTypes -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp '_rels\.rels') -Value $relsRoot -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\document.xml') -Value $documentXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\styles.xml') -Value $stylesXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\settings.xml') -Value $settingsXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\numbering.xml') -Value $numberingXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\footer1.xml') -Value $footerXml -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'word\_rels\document.xml.rels') -Value $relsDoc -Encoding UTF8

    if ($images.Count) {
        New-Item -ItemType Directory -Path (Join-Path $tmp 'word\media') -Force | Out-Null
        foreach ($im in $images) { Copy-Item -LiteralPath $im.Path -Destination (Join-Path $tmp "word\media\$($im.MediaName)") -Force }
    }

    $wordDir = Join-Path $src 'Report\word'
    if (-not (Test-Path $wordDir)) { New-Item -ItemType Directory -Path $wordDir -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyy-MM-dd_HH.mm.ss'
    $out = Join-Path $wordDir "Systemdokumentation Entra ID $stamp.docx"
    if (Test-Path $out) { Remove-Item -LiteralPath $out -Force }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $out)
    Remove-Item -LiteralPath $tmp -Recurse -Force

    Write-Host "  Word report: $out" -ForegroundColor DarkGreen
    return $out
}
