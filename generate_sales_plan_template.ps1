param(
  [string]$OutputPath = (Join-Path (Get-Location) 'mau_xuat_sales_plan_start_2026-06.xlsx'),
  [datetime]$CurrentDate = [datetime]'2026-05-28',
  [int]$DataStartRow = 18,
  [int]$DataEndRow = 217
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Security

function Esc([object]$Value) {
  if ($null -eq $Value) { return '' }
  return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ColName([int]$Num) {
  $s = ''
  while ($Num -gt 0) {
    $Num--
    $s = [char](65 + ($Num % 26)) + $s
    $Num = [math]::Floor($Num / 26)
  }
  return $s
}

function CellRef([int]$Col, [int]$Row) {
  return (ColName $Col) + $Row
}

function TextCell([int]$Col, [int]$Row, [string]$Text, [int]$Style = 0) {
  $r = CellRef $Col $Row
  return '<c r="' + $r + '" s="' + $Style + '" t="inlineStr"><is><t xml:space="preserve">' + (Esc $Text) + '</t></is></c>'
}

function NumCell([int]$Col, [int]$Row, [object]$Num, [int]$Style = 0) {
  $r = CellRef $Col $Row
  if ($null -eq $Num -or [string]$Num -eq '') {
    return '<c r="' + $r + '" s="' + $Style + '"/>'
  }
  $text = ([string]$Num).Replace(',', '.')
  return '<c r="' + $r + '" s="' + $Style + '"><v>' + $text + '</v></c>'
}

function FormulaCell([int]$Col, [int]$Row, [string]$Formula, [int]$Style = 0) {
  $r = CellRef $Col $Row
  return '<c r="' + $r + '" s="' + $Style + '"><f>' + (Esc $Formula) + '</f></c>'
}

function SumRefs([string[]]$Refs) {
  if ($Refs.Count -eq 0) { return '0' }
  if ($Refs.Count -eq 1) { return $Refs[0] }
  return ($Refs -join '+')
}

function Add-ZipEntry([System.IO.Compression.ZipArchive]$Zip, [string]$Name, [string]$Content) {
  $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
  $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
  try { $writer.Write($Content) } finally { $writer.Dispose() }
}

$products = @(
  @{ Name = 'SPR EPO'; Price = 590000 },
  @{ Name = 'SPR Collagen'; Price = 590000 },
  @{ Name = 'SPR Vitamin C 500'; Price = 750000 },
  @{ Name = 'SPR Vitamin C 1000'; Price = 890000 },
  @{ Name = 'SPR Milky Calcium'; Price = 745000 },
  @{ Name = 'SPR Kids Fish Oil'; Price = 745000 },
  @{ Name = 'SPR Garlic Oil'; Price = 525000 },
  @{ Name = 'SPR Colostrum Power'; Price = 775000 },
  @{ Name = 'SPR Colostrum Kids'; Price = 745000 },
  @{ Name = 'SPR Fatclear'; Price = 949000 },
  @{ Name = 'SPR Joint Swelling'; Price = 495000 }
)

$channels = @(
  @{ Code = 1001; Channel = 'MTC'; Sub = 'LC, PMC, AK' },
  @{ Code = 1002; Channel = 'MTC'; Sub = 'Health & Beauty' },
  @{ Code = 1003; Channel = 'Ecom'; Sub = 'Ecommerce' },
  @{ Code = 1004; Channel = 'Baby & Mom'; Sub = 'Baby & Mom' },
  @{ Code = 1005; Channel = 'OTC'; Sub = 'Key/Chain' },
  @{ Code = 1006; Channel = 'ETC'; Sub = 'ETC' },
  @{ Code = 1007; Channel = 'OTC'; Sub = 'IDP' }
)

$fixedHeaders = @('Code channel', 'Channel', 'Sub-channel', 'Customer', 'Stores', 'Province', 'Regional', 'Sub-Channel 1', 'STAFF')
$productCount = $products.Count
$monthNames = @('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC')

$startBase = $CurrentDate.AddMonths(1)
$startMonth = [datetime]::new($startBase.Year, $startBase.Month, 1)
$endMonth = [datetime]::new($startMonth.Year, 12, 1)

$blocks = [System.Collections.Generic.List[object]]::new()
$monthBlocksByQuarter = @{}
$currentCol = 10

for ($dt = $startMonth; $dt -le $endMonth; $dt = $dt.AddMonths(1)) {
  $quarter = [math]::Ceiling($dt.Month / 3)
  $productCols = @()
  for ($i = 0; $i -lt $productCount; $i++) { $productCols += ($currentCol + $i) }

  $monthBlock = [pscustomobject]@{
    Type = 'month'
    Label = ($monthNames[$dt.Month - 1] + ' ' + $dt.Year)
    Quarter = $quarter
    Year = $dt.Year
    Start = $currentCol
    End = ($currentCol + $productCount + 1)
    ProductCols = $productCols
    TotalCol = ($currentCol + $productCount)
    GrossCol = ($currentCol + $productCount + 1)
  }
  $blocks.Add($monthBlock)

  $quarterKey = [string]$quarter
  if (-not $monthBlocksByQuarter.ContainsKey($quarterKey)) {
    $monthBlocksByQuarter[$quarterKey] = [System.Collections.Generic.List[object]]::new()
  }
  $monthBlocksByQuarter[$quarterKey].Add($monthBlock)
  $currentCol += $productCount + 2

  $next = $dt.AddMonths(1)
  $isLastInQuarter = ($next.Month -eq 1) -or ([math]::Ceiling($next.Month / 3) -ne $quarter) -or ($next -gt $endMonth)
  if ($isLastInQuarter) {
    $qProductCols = @()
    for ($i = 0; $i -lt $productCount; $i++) { $qProductCols += ($currentCol + $i) }
    $quarterBlock = [pscustomobject]@{
      Type = 'quarter'
      Label = ('TOTAL Q' + $quarter + ' ' + $dt.Year)
      Quarter = $quarter
      Year = $dt.Year
      Start = $currentCol
      End = ($currentCol + $productCount)
      ProductCols = $qProductCols
      TotalCol = ($currentCol + $productCount)
      GrossCol = $null
      MonthBlocks = $monthBlocksByQuarter[$quarterKey]
    }
    $blocks.Add($quarterBlock)
    $currentCol += $productCount + 1
  }
}

$quarterBlocks = @($blocks | Where-Object { $_.Type -eq 'quarter' })
$yearProductCols = @()
for ($i = 0; $i -lt $productCount; $i++) { $yearProductCols += ($currentCol + $i) }
$yearBlock = [pscustomobject]@{
  Type = 'year'
  Label = ('TOTAL YEAR ' + $startMonth.Year)
  Quarter = $null
  Year = $startMonth.Year
  Start = $currentCol
  End = ($currentCol + $productCount + 1)
  ProductCols = $yearProductCols
  TotalCol = ($currentCol + $productCount)
  GrossCol = ($currentCol + $productCount + 1)
  QuarterBlocks = $quarterBlocks
}
$blocks.Add($yearBlock)

$lastCol = $yearBlock.End
$lastColName = ColName $lastCol

$rows = @{}
function AddCell([int]$Row, [string]$Xml) {
  if (-not $script:rows.ContainsKey($Row)) {
    $script:rows[$Row] = [System.Collections.Generic.List[string]]::new()
  }
  $script:rows[$Row].Add($Xml)
}

$mergeRefs = [System.Collections.Generic.List[string]]::new()

AddCell 1 (TextCell 1 1 ('SALES PLAN EXPORT TEMPLATE - START ' + $monthNames[$startMonth.Month - 1] + ' ' + $startMonth.Year) 1)
$mergeRefs.Add('A1:' + $lastColName + '1')
AddCell 2 (TextCell 1 2 'Rule: first month = current month + 1. Quarter totals are created even when the first quarter is incomplete.' 2)
$mergeRefs.Add('A2:' + $lastColName + '2')

for ($i = 0; $i -lt $fixedHeaders.Count; $i++) {
  AddCell 3 (TextCell ($i + 1) 3 $fixedHeaders[$i] 4)
  AddCell 16 (TextCell ($i + 1) 16 $fixedHeaders[$i] 4)
  AddCell 17 (TextCell ($i + 1) 17 '' 3)
}

for ($idx = 0; $idx -lt $channels.Count; $idx++) {
  $r = 5 + $idx
  $ch = $channels[$idx]
  AddCell $r (NumCell 1 $r $ch.Code 6)
  AddCell $r (TextCell 2 $r $ch.Channel 6)
  AddCell $r (TextCell 3 $r $ch.Sub 6)
  for ($c = 4; $c -le 9; $c++) { AddCell $r (TextCell $c $r '' 6) }
}
AddCell 12 (NumCell 1 12 1000 7)
AddCell 12 (TextCell 2 12 'Nationwide' 7)
for ($c = 3; $c -le 9; $c++) { AddCell 12 (TextCell $c 12 '' 7) }

foreach ($block in $blocks) {
  $labelStyle = if ($block.Type -eq 'month') { 8 } elseif ($block.Type -eq 'quarter') { 9 } else { 10 }
  AddCell 15 (TextCell $block.Start 15 $block.Label $labelStyle)
  if ($block.Start -lt $block.End) {
    $mergeRefs.Add((ColName $block.Start) + '15:' + (ColName $block.End) + '15')
  }

  for ($i = 0; $i -lt $productCount; $i++) {
    $col = $block.ProductCols[$i]
    AddCell 16 (TextCell $col 16 $products[$i].Name 4)
    AddCell 17 (NumCell $col 17 $products[$i].Price 3)
  }

  if ($block.Type -eq 'quarter') {
    AddCell 16 (TextCell $block.TotalCol 16 'Quarter Total Volume' 5)
    AddCell 17 (TextCell $block.TotalCol 17 '' 3)
  } else {
    AddCell 16 (TextCell $block.TotalCol 16 'Total Volume' 5)
    AddCell 16 (TextCell $block.GrossCol 16 "Gross Sales`n+VAT" 5)
    AddCell 17 (TextCell $block.TotalCol 17 '' 3)
    AddCell 17 (TextCell $block.GrossCol 17 '' 3)
  }
}

foreach ($r in 5..12) {
  foreach ($block in $blocks) {
    if ($block.Type -eq 'month') {
      foreach ($col in $block.ProductCols) {
        if ($r -eq 12) {
          $formula = 'SUM(' + (CellRef $col 5) + ':' + (CellRef $col 11) + ')'
        } else {
          $formula = 'SUMIF($A$' + $DataStartRow + ':$A$' + $DataEndRow + ',$A' + $r + ',' + (ColName $col) + '$' + $DataStartRow + ':' + (ColName $col) + '$' + $DataEndRow + ')'
        }
        AddCell $r (FormulaCell $col $r $formula 11)
      }
      if ($r -eq 12) {
        AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.TotalCol 5) + ':' + (CellRef $block.TotalCol 11) + ')') 12)
        AddCell $r (FormulaCell $block.GrossCol $r ('SUM(' + (CellRef $block.GrossCol 5) + ':' + (CellRef $block.GrossCol 11) + ')') 12)
      } else {
        AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 12)
        AddCell $r (FormulaCell $block.GrossCol $r ('SUMIF($A$' + $DataStartRow + ':$A$' + $DataEndRow + ',$A' + $r + ',' + (ColName $block.GrossCol) + '$' + $DataStartRow + ':' + (ColName $block.GrossCol) + '$' + $DataEndRow + ')') 12)
      }
    } elseif ($block.Type -eq 'quarter') {
      for ($i = 0; $i -lt $productCount; $i++) {
        $refs = @()
        foreach ($monthBlock in $block.MonthBlocks) { $refs += (CellRef $monthBlock.ProductCols[$i] $r) }
        AddCell $r (FormulaCell $block.ProductCols[$i] $r (SumRefs $refs) 13)
      }
      AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 13)
    } else {
      for ($i = 0; $i -lt $productCount; $i++) {
        $refs = @()
        foreach ($quarterBlock in $block.QuarterBlocks) { $refs += (CellRef $quarterBlock.ProductCols[$i] $r) }
        AddCell $r (FormulaCell $block.ProductCols[$i] $r (SumRefs $refs) 14)
      }
      AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 14)
      AddCell $r (FormulaCell $block.GrossCol $r ('SUMPRODUCT(' + (CellRef $block.ProductCols[0] 17) + ':' + (CellRef $block.ProductCols[-1] 17) + ',' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 14)
    }
  }
}

for ($r = $DataStartRow; $r -le $DataEndRow; $r++) {
  for ($c = 1; $c -le 9; $c++) { AddCell $r (TextCell $c $r '' 15) }
  foreach ($block in $blocks) {
    if ($block.Type -eq 'month') {
      foreach ($col in $block.ProductCols) { AddCell $r (NumCell $col $r '' 16) }
      AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 17)
      AddCell $r (FormulaCell $block.GrossCol $r ('SUMPRODUCT(' + (CellRef $block.ProductCols[0] 17) + ':' + (CellRef $block.ProductCols[-1] 17) + ',' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 17)
    } elseif ($block.Type -eq 'quarter') {
      for ($i = 0; $i -lt $productCount; $i++) {
        $refs = @()
        foreach ($monthBlock in $block.MonthBlocks) { $refs += (CellRef $monthBlock.ProductCols[$i] $r) }
        AddCell $r (FormulaCell $block.ProductCols[$i] $r (SumRefs $refs) 18)
      }
      AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 18)
    } else {
      for ($i = 0; $i -lt $productCount; $i++) {
        $refs = @()
        foreach ($quarterBlock in $block.QuarterBlocks) { $refs += (CellRef $quarterBlock.ProductCols[$i] $r) }
        AddCell $r (FormulaCell $block.ProductCols[$i] $r (SumRefs $refs) 19)
      }
      AddCell $r (FormulaCell $block.TotalCol $r ('SUM(' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 19)
      AddCell $r (FormulaCell $block.GrossCol $r ('SUMPRODUCT(' + (CellRef $block.ProductCols[0] 17) + ':' + (CellRef $block.ProductCols[-1] 17) + ',' + (CellRef $block.ProductCols[0] $r) + ':' + (CellRef $block.ProductCols[-1] $r) + ')') 19)
    }
  }
}

$sheet = [System.Text.StringBuilder]::new()
[void]$sheet.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$sheet.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
[void]$sheet.Append('<dimension ref="A1:' + $lastColName + $DataEndRow + '"/>')
[void]$sheet.Append('<sheetViews><sheetView workbookViewId="0"><pane xSplit="9" ySplit="17" topLeftCell="J18" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="J18" sqref="J18"/></sheetView></sheetViews>')
[void]$sheet.Append('<sheetFormatPr defaultRowHeight="18"/>')
[void]$sheet.Append('<cols>')
$widths = @{ 1 = 12; 2 = 20; 3 = 24; 4 = 30; 5 = 10; 6 = 18; 7 = 14; 8 = 18; 9 = 12 }
for ($c = 1; $c -le 9; $c++) {
  [void]$sheet.Append('<col min="' + $c + '" max="' + $c + '" width="' + $widths[$c] + '" customWidth="1"/>')
}
[void]$sheet.Append('<col min="10" max="' + $lastCol + '" width="18" customWidth="1"/>')
[void]$sheet.Append('</cols><sheetData>')
foreach ($rowKey in ($rows.Keys | Sort-Object { [int]$_ })) {
  $height = if ([int]$rowKey -eq 16) { 32 } elseif ([int]$rowKey -eq 15) { 24 } else { 18 }
  [void]$sheet.Append('<row r="' + $rowKey + '" ht="' + $height + '" customHeight="1">')
  foreach ($cellXml in $rows[$rowKey]) { [void]$sheet.Append($cellXml) }
  [void]$sheet.Append('</row>')
}
[void]$sheet.Append('</sheetData>')
[void]$sheet.Append('<autoFilter ref="A16:' + $lastColName + $DataEndRow + '"/>')
[void]$sheet.Append('<mergeCells count="' + $mergeRefs.Count + '">')
foreach ($mergeRef in $mergeRefs) { [void]$sheet.Append('<mergeCell ref="' + $mergeRef + '"/>') }
[void]$sheet.Append('</mergeCells>')
[void]$sheet.Append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
[void]$sheet.Append('</worksheet>')

$productSheetRows = @{}
function AddProductCell([int]$Row, [string]$Xml) {
  if (-not $script:productSheetRows.ContainsKey($Row)) {
    $script:productSheetRows[$Row] = [System.Collections.Generic.List[string]]::new()
  }
  $script:productSheetRows[$Row].Add($Xml)
}
AddProductCell 1 (TextCell 1 1 'Product Name' 4)
AddProductCell 1 (TextCell 2 1 'Price' 4)
for ($i = 0; $i -lt $productCount; $i++) {
  $r = $i + 2
  AddProductCell $r (TextCell 1 $r $products[$i].Name 15)
  AddProductCell $r (NumCell 2 $r $products[$i].Price 16)
}
$productSheet = [System.Text.StringBuilder]::new()
[void]$productSheet.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B' + ($productCount + 1) + '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/></cols><sheetData>')
foreach ($rowKey in ($productSheetRows.Keys | Sort-Object { [int]$_ })) {
  [void]$productSheet.Append('<row r="' + $rowKey + '">')
  foreach ($cellXml in $productSheetRows[$rowKey]) { [void]$productSheet.Append($cellXml) }
  [void]$productSheet.Append('</row>')
}
[void]$productSheet.Append('</sheetData></worksheet>')

$mapLines = @(
  @('Field', 'Value'),
  @('Main sheet', 'DetailPlan_Template'),
  @('Start month rule', 'Current month + 1'),
  @('Template start month', ($monthNames[$startMonth.Month - 1] + ' ' + $startMonth.Year)),
  @('Customer data starts at row', [string]$DataStartRow),
  @('Customer data prepared rows', ($DataStartRow.ToString() + ':' + $DataEndRow.ToString())),
  @('Fixed customer columns', 'A:I = Code channel, Channel, Sub-channel, Customer, Stores, Province, Regional, Sub-Channel 1, STAFF'),
  @('Product input source', 'UI should provide Product Name + Price; this template uses Products sheet as sample values'),
  @('Month block rule', 'For each month: product quantity columns, Total Volume, Gross Sales +VAT'),
  @('Quarter block rule', 'After each quarter appearing in the timeline: product quantity totals, Quarter Total Volume. Create it even if the quarter is incomplete.'),
  @('Year block rule', 'At year end: product quantity totals, Total Volume, Gross Sales +VAT'),
  @('Channel summary rows', 'Rows 5:11 by channel code; row 12 Nationwide total'),
  @('Do not overwrite', 'Rows 15:17 contain labels, headers, prices; write customer rows from row 18')
)
$mappingSheet = [System.Text.StringBuilder]::new()
[void]$mappingSheet.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B' + $mapLines.Count + '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="110" customWidth="1"/></cols><sheetData>')
for ($i = 0; $i -lt $mapLines.Count; $i++) {
  $r = $i + 1
  $style = if ($r -eq 1) { 4 } else { 15 }
  [void]$mappingSheet.Append('<row r="' + $r + '">')
  [void]$mappingSheet.Append((TextCell 1 $r $mapLines[$i][0] $style))
  [void]$mappingSheet.Append((TextCell 2 $r $mapLines[$i][1] $style))
  [void]$mappingSheet.Append('</row>')
}
[void]$mappingSheet.Append('</sheetData></worksheet>')

$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="8">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF6A329F"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0F6B5F"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="20">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="2" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
'@

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@

$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
'@

$workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="DetailPlan_Template" sheetId="1" r:id="rId1"/><sheet name="Products" sheetId="2" r:id="rId2"/><sheet name="Mapping" sheetId="3" r:id="rId3"/></sheets><calcPr calcId="171027" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>
'@

$workbookRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
'@

if (Test-Path -LiteralPath $OutputPath) {
  Remove-Item -LiteralPath $OutputPath -Force
}

$fs = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::CreateNew)
$zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
  Add-ZipEntry $zip '[Content_Types].xml' $contentTypes
  Add-ZipEntry $zip '_rels/.rels' $rootRels
  Add-ZipEntry $zip 'xl/workbook.xml' $workbook
  Add-ZipEntry $zip 'xl/_rels/workbook.xml.rels' $workbookRels
  Add-ZipEntry $zip 'xl/styles.xml' $stylesXml
  Add-ZipEntry $zip 'xl/worksheets/sheet1.xml' $sheet.ToString()
  Add-ZipEntry $zip 'xl/worksheets/sheet2.xml' $productSheet.ToString()
  Add-ZipEntry $zip 'xl/worksheets/sheet3.xml' $mappingSheet.ToString()
} finally {
  $zip.Dispose()
  $fs.Dispose()
}

[pscustomobject]@{
  OutputPath = $OutputPath
  StartMonth = ($monthNames[$startMonth.Month - 1] + ' ' + $startMonth.Year)
  EndColumn = $lastColName
  DataRows = ($DataStartRow.ToString() + ':' + $DataEndRow.ToString())
}
