# Maakt Cursussen.xlsx aan (via Excel COM) en vult het met de gegevens uit een bestaande JSON.
# Enkel nodig tijdens ontwikkeling; de tool zelf (json-builder.html) heeft dit niet nodig.
#
# Gebruik:  pwsh tools/maak-sjabloon.ps1 [-Json example/Cursussen.json] [-Uit Cursussen.xlsx] [-Leeg]

param(
    [string]$Json = (Join-Path $PSScriptRoot '..\example\Cursussen.json'),
    [string]$Uit  = (Join-Path $PSScriptRoot '..\Cursussen.xlsx'),
    [switch]$Leeg
)

$ErrorActionPreference = 'Stop'

# ---- Lijsten (toegelaten waarden) ----
$lijsten = [ordered]@{
    'Taal'          = @('NEDERLANDS', 'FRANS')
    'Statuut'       = @('WERKNEMER', 'GEPENSIONEERDE')
    'JaNee'         = @('JA', 'NEE')
    'Verlofstelsel' = @('VRIJETIJD', 'VLAAMS_OPLEIDINGS_VERLOF', 'BETAALD_EDUCATIEF_VERLOF', 'SYNDICAAL_VERLOF')
    'Mandaat'       = @('GEEN_VEREIST', 'EEN_MANDAAT_VEREIST', 'MANDAAT_ONDERNEMINGSRAAD', 'MANDAAT_COMITE_VOOR_PREVENTIE_EN_BESCHERMING', 'MANDAAT_SYNDICALE_DELEGATIE')
    'Regio'         = @('VLAANDEREN', 'WALLONIE', 'BRUSSEL')
}
$lijstRijen = 100   # bereik voor keuzelijsten (ruimte om codes toe te voegen)
$maxRijen   = 1000  # rijen met gegevensvalidatie
$aantalDatums = 8

$cursusKoppen = @('Titel', 'Categorie', 'Categorie URL', 'Vormingsorganisatie', 'Taal',
    'Begin inschrijving', 'Einde inschrijving', 'Vormingsjaar', 'Credits', 'Statuut') +
    $lijsten['Verlofstelsel'] + $lijsten['Mandaat'] + $lijsten['Regio']
$actKoppen = @('Cursus', 'Gemeente', 'Residentieel', 'Actief') + (1..$aantalDatums | ForEach-Object { "Datum $_" })

function Get-OADate([string]$iso) { ([datetime]::ParseExact($iso, 'yyyy-MM-dd', $null)).ToOADate() }
function Get-ColLetter([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; $s }

# ---- Gegevens uit JSON ----
$cursusRijen = @()
$actRijen = @()
if (-not $Leeg) {
    $cursussen = Get-Content -Raw -Encoding UTF8 $Json | ConvertFrom-Json
    foreach ($c in $cursussen) {
        $rij = @($c.titel, $c.categorie.naam, $c.categorie.url, $c.vormingsorganisatie.naam, $c.vormingsorganisatie.taal,
            (Get-OADate $c.vormingsorganisatie.begindatumInschrijvingsperiode),
            (Get-OADate $c.vormingsorganisatie.einddatumInschrijvingsperiode),
            $c.vormingsjaar.vormingsjaar, $c.vereistAantalCreditsCursus, $c.vereisteEigenschappenCursisten.statuut)
        foreach ($v in $lijsten['Verlofstelsel']) { $rij += $(if ($c.mogelijkeVerlofstelsels -contains $v) { 'X' } else { $null }) }
        foreach ($v in $lijsten['Mandaat']) { $rij += $(if ($c.vereisteEigenschappenCursisten.mandaatVereiste -contains $v) { 'X' } else { $null }) }
        foreach ($v in $lijsten['Regio']) { $rij += $(if ($c.cursusRegios -contains $v) { 'X' } else { $null }) }
        $cursusRijen += , $rij

        foreach ($a in $c.activiteiten) {
            $r = @($c.titel, $a.locatie.gemeente, $(if ($a.residentieel) { 'JA' } else { 'NEE' }), $(if ($a.actief) { 'JA' } else { 'NEE' }))
            for ($i = 0; $i -lt $aantalDatums; $i++) {
                $r += $(if ($a.data -and $i -lt @($a.data).Count) { Get-OADate @($a.data)[$i].localDate } else { $null })
            }
            $actRijen += , $r
        }
    }
}

function Write-Rows($sheet, $rows, [int]$cols) {
    # cel per cel: PowerShell 7 kan geen object[,] via COM aan een Range toekennen
    for ($i = 0; $i -lt $rows.Count; $i++) {
        for ($j = 0; $j -lt $cols; $j++) {
            $v = $rows[$i][$j]
            if ($null -eq $v) { continue }
            $cel = $sheet.Cells($i + 2, $j + 1)
            # aparte toewijzingen per type: de COM-binder van PS7 onthoudt het type per regel code
            if ($v -is [string]) { $cel.Value2 = [string]$v } else { $cel.Value2 = [double]$v }
        }
    }
}

function Set-Header($sheet, $koppen) {
    for ($i = 0; $i -lt $koppen.Count; $i++) { $sheet.Cells(1, $i + 1).Value2 = $koppen[$i] }
    $h = $sheet.Range($sheet.Cells(1, 1), $sheet.Cells(1, $koppen.Count))
    $h.Font.Bold = $true
    $h.Font.Color = 0xFFFFFF
    $h.Interior.Color = 0x6B3A1F   # BGR: donkerblauw
    $h.WrapText = $true
    $h.VerticalAlignment = -4108
}

function Add-List($range, [string]$formula) {
    $range.Validation.Delete()
    $range.Validation.Add(3, 1, 1, $formula) | Out-Null
    $range.Validation.IgnoreBlank = $true
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 3) { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) | Out-Null }
    $wsC = $wb.Worksheets.Item(1); $wsC.Name = 'Cursussen'
    $wsA = $wb.Worksheets.Item(2); $wsA.Name = 'Activiteiten'
    $wsL = $wb.Worksheets.Item(3); $wsL.Name = 'Lijsten'

    # ---- Lijsten ----
    $kol = 1
    $lijstKolom = @{}
    foreach ($naam in $lijsten.Keys) {
        $wsL.Cells(1, $kol).Value2 = $naam
        $vals = $lijsten[$naam]
        for ($i = 0; $i -lt $vals.Count; $i++) { $wsL.Cells($i + 2, $kol).Value2 = $vals[$i] }
        $lijstKolom[$naam] = Get-ColLetter $kol
        $kol++
    }
    Set-Header $wsL @($lijsten.Keys)
    $wsL.Columns.AutoFit() | Out-Null
    function Lijst($naam) { $l = $lijstKolom[$naam]; "=Lijsten!`$$l`$2:`$$l`$$lijstRijen" }

    # ---- Cursussen ----
    Set-Header $wsC $cursusKoppen
    Write-Rows $wsC $cursusRijen $cursusKoppen.Count
    $laatste = $maxRijen
    $col = @{}; for ($i = 0; $i -lt $cursusKoppen.Count; $i++) { $col[$cursusKoppen[$i]] = $i + 1 }
    # komma-operator: anders splitst PowerShell de Range op in losse cellen
    function KolomBereik($ws, $c) { , $ws.Range($ws.Cells(2, $c), $ws.Cells($laatste, $c)) }

    Add-List (KolomBereik $wsC $col['Taal']) (Lijst 'Taal')
    Add-List (KolomBereik $wsC $col['Statuut']) (Lijst 'Statuut')
    (KolomBereik $wsC $col['Begin inschrijving']).NumberFormat = 'dd/mm/yyyy'
    (KolomBereik $wsC $col['Einde inschrijving']).NumberFormat = 'dd/mm/yyyy'
    (KolomBereik $wsC $col['Vormingsjaar']).NumberFormat = '@'
    $eersteCode = $col[$lijsten['Verlofstelsel'][0]]
    $laatsteCode = $cursusKoppen.Count
    $codes = $wsC.Range($wsC.Cells(2, $eersteCode), $wsC.Cells($laatste, $laatsteCode))
    Add-List $codes 'X'
    $codes.HorizontalAlignment = -4108

    $wsC.Columns.AutoFit() | Out-Null
    $wsC.Columns($col['Titel']).ColumnWidth = 55
    $wsC.Columns($col['Categorie']).ColumnWidth = 35
    $wsC.Columns($col['Categorie URL']).ColumnWidth = 40
    for ($c = $eersteCode; $c -le $laatsteCode; $c++) {
        $wsC.Columns($c).ColumnWidth = 14
        $wsC.Cells(1, $c).Orientation = 90
    }
    $wsC.Rows(1).RowHeight = 190
    # groepskleuren voor de X-kolommen
    $kleuren = @{ 'Verlofstelsel' = 0x4F7A2E; 'Mandaat' = 0x1F5FA8; 'Regio' = 0x7A3E6B }
    foreach ($g in 'Verlofstelsel', 'Mandaat', 'Regio') {
        foreach ($v in $lijsten[$g]) { $wsC.Cells(1, $col[$v]).Interior.Color = $kleuren[$g] }
    }

    # ---- Activiteiten ----
    Set-Header $wsA $actKoppen
    Write-Rows $wsA $actRijen $actKoppen.Count
    Add-List (KolomBereik $wsA 1) "=Cursussen!`$A`$2:`$A`$$maxRijen"
    Add-List (KolomBereik $wsA 3) (Lijst 'JaNee')
    Add-List (KolomBereik $wsA 4) (Lijst 'JaNee')
    $wsA.Range($wsA.Cells(2, 5), $wsA.Cells($laatste, 4 + $aantalDatums)).NumberFormat = 'dd/mm/yyyy'
    $wsA.Columns.AutoFit() | Out-Null
    $wsA.Columns(1).ColumnWidth = 55
    $wsA.Rows(1).RowHeight = 20
    $wsC.Rows(1).VerticalAlignment = -4107

    foreach ($ws in $wsC, $wsA) {
        $ws.Activate()
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.SplitColumn = 1
        $excel.ActiveWindow.FreezePanes = $true
    }
    $wsC.Activate()

    $pad = [System.IO.Path]::GetFullPath($Uit)
    if (Test-Path $pad) { Remove-Item $pad }
    $wb.SaveAs($pad, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    "Aangemaakt: $pad ($($cursusRijen.Count) cursussen, $($actRijen.Count) activiteiten)"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
