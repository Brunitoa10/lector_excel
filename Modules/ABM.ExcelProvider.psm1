Set-StrictMode -Version Latest

<#
    Estrategia para libros de Excel (.xlsx / .xlsm) usando OpenXML nativo (ZIP + XML de .NET).
    No necesita Excel instalado ni COM.

    Lectura : primera fila con datos = encabezados; detecta columna inicial, textos compartidos,
              textos en línea, booleanos y fechas (según el formato de celda).
    Escritura: reemplaza sólo las filas de la tabla en la hoja elegida y conserva:
              filas por encima del encabezado (títulos), estilos de cada columna, anchos,
              paneles inmovilizados, autofiltro, tablas de Excel y el resto de las hojas.
    Limitación conocida: las fórmulas dentro del rango editado se guardan como valores.
#>

$script:NsMain    = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:NsDocRel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:NsXml     = 'http://www.w3.org/XML/1998/namespace'
$script:Invariant = [System.Globalization.CultureInfo]::InvariantCulture
$script:BuiltInDateFormats = @(14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47)
$script:DateInputFormats   = [string[]]@(
    'dd/MM/yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm', 'dd/MM/yyyy',
    'd/M/yyyy H:mm:ss', 'd/M/yyyy H:mm', 'd/M/yyyy',
    'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd'
)
$script:InvalidXmlChars = New-Object System.Text.RegularExpressions.Regex('[\x00-\x08\x0B\x0C\x0E-\x1F]')
$script:NumberPattern   = New-Object System.Text.RegularExpressions.Regex('^-?(0|[1-9]\d{0,14})(\.\d+)?([eE][-+]?\d+)?$')
$script:TrueText  = 'VERDADERO'
$script:FalseText = 'FALSO'

#region API pública
function Read-ExcelSheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SheetName
    )

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $info   = Get-WorkbookInfo -Zip $zip
        $sheet  = Select-Sheet -Info $info -SheetName $SheetName
        $shared = Get-SharedStrings -Zip $zip -Part $info.SharedStringsPart
        $dates  = Get-DateStyleSet -Zip $zip -StylesPart $info.StylesPart
        $doc    = Open-ZipXml -Zip $zip -EntryName $sheet.Part
        if ($null -eq $doc) { throw "No se encontró la hoja '$($sheet.Name)' dentro del archivo." }

        $rows      = Read-SheetRows -Document $doc -Ns (New-NsManager $doc) -Shared $shared -DateStyles $dates
        $sheetList = @($info.Sheets | ForEach-Object { $_.Name })
        $table     = ConvertTo-SheetTable -Rows $rows
        $layout    = Get-SheetLayout -Rows $rows

        return @{
            Table = $table
            Meta  = @{ Sheets = $sheetList; Sheet = $sheet.Name; HeaderRow = $layout.HeaderRow; ColOffset = $layout.ColOffset }
        }
    }
    finally { $zip.Dispose() }
}

function Save-ExcelSheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SheetName,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [int]$HeaderRow = 1,
        [int]$ColOffset = 0,
        [string]$DefaultHeaderStyle = '',
        [string]$DefaultDataStyle = ''
    )

    if ($HeaderRow -lt 1) { $HeaderRow = 1 }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $info  = Get-WorkbookInfo -Zip $zip
        $sheet = Select-Sheet -Info $info -SheetName $SheetName
        $dates = Get-DateStyleSet -Zip $zip -StylesPart $info.StylesPart
        $doc   = Open-ZipXml -Zip $zip -EntryName $sheet.Part
        $ns    = New-NsManager $doc
        $sheetData = $doc.SelectSingleNode('/x:worksheet/x:sheetData', $ns)
        if ($null -eq $sheetData) { throw "La hoja '$($sheet.Name)' no tiene una estructura válida." }

        $styles  = Remove-TableRows -SheetData $sheetData -Ns $ns -HeaderRow $HeaderRow
        $lastRow = Write-TableRows -Document $doc -SheetData $sheetData -Table $Table -HeaderRow $HeaderRow -ColOffset $ColOffset `
                                   -Styles $styles -DateStyles $dates -DefaultHeaderStyle $DefaultHeaderStyle -DefaultDataStyle $DefaultDataStyle

        $firstColumn = ConvertTo-ColumnLetters $ColOffset
        $lastColumn  = ConvertTo-ColumnLetters ($ColOffset + [math]::Max(1, $Table.Columns.Count) - 1)
        Update-SheetRanges -Document $doc -Ns $ns -FirstColumn $firstColumn -LastColumn $lastColumn -HeaderRow $HeaderRow -LastRow $lastRow

        Save-ZipXml -Zip $zip -EntryName $sheet.Part -Document $doc
        $tableRange = '{0}{1}:{2}{3}' -f $firstColumn, $HeaderRow, $lastColumn, [math]::Max($lastRow, $HeaderRow + 1)
        Update-SheetTables -Zip $zip -SheetPart $sheet.Part -Range $tableRange
        Remove-CalcChain -Zip $zip -Info $info
    }
    finally { $zip.Dispose() }
}

function New-ExcelWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SheetName = 'Datos',
        [Parameter(Mandatory)][System.Data.DataTable]$Table
    )

    $SheetName = ConvertTo-SafeSheetName $SheetName
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $parts = Get-NewWorkbookParts -SheetName $SheetName -ColumnCount ([math]::Max(1, $Table.Columns.Count))
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) { Save-ZipText -Zip $zip -EntryName $name -Content $parts[$name] }
    }
    finally { $zip.Dispose() }

    # Estilo 1 = encabezado corporativo, estilo 2 = celda con borde (definidos en styles.xml).
    Save-ExcelSheet -Path $Path -SheetName $SheetName -Table $Table -DefaultHeaderStyle '1' -DefaultDataStyle '2'
}
#endregion

#region ZIP / XML
function Open-ZipXml {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName)

    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8, $true)
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $true
        $doc.XmlResolver = $null
        $doc.LoadXml($reader.ReadToEnd())
        return ,$doc
    }
    finally { $reader.Dispose() }
}

function Save-ZipXml {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName, [System.Xml.XmlDocument]$Document)

    $existing = $Zip.GetEntry($EntryName)
    if ($null -ne $existing) { $existing.Delete() }
    $entry  = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try { $Document.Save($writer) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Save-ZipText {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName, [string]$Content)

    $entry  = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Content) } finally { $writer.Dispose() }
}

function New-NsManager {
    param([System.Xml.XmlDocument]$Document)
    $ns = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
    $ns.AddNamespace('x', $script:NsMain)
    return ,$ns
}

function Get-PartDirectory {
    param([string]$PartName)
    $index = $PartName.LastIndexOf('/')
    if ($index -lt 0) { return '' }
    return $PartName.Substring(0, $index)
}

function Get-RelsPath {
    param([string]$PartName)
    $directory = Get-PartDirectory $PartName
    $file = $PartName.Substring($PartName.LastIndexOf('/') + 1)
    if ($directory) { return "$directory/_rels/$file.rels" }
    return "_rels/$file.rels"
}

function Resolve-PartName {
    param([string]$BaseDirectory, [string]$Target)

    if ($Target.StartsWith('/')) { return $Target.TrimStart('/') }
    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($segment in ($BaseDirectory -split '/')) { if ($segment) { $segments.Add($segment) } }
    foreach ($segment in ($Target -split '/')) {
        if ($segment -eq '..') { if ($segments.Count -gt 0) { $segments.RemoveAt($segments.Count - 1) } }
        elseif ($segment -and $segment -ne '.') { $segments.Add($segment) }
    }
    return ($segments -join '/')
}

function Get-Relationships {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$PartName)

    $list = New-Object System.Collections.Generic.List[object]
    $doc  = Open-ZipXml -Zip $Zip -EntryName (Get-RelsPath $PartName)
    if ($null -eq $doc) { return ,$list.ToArray() }

    $baseDirectory = Get-PartDirectory $PartName
    foreach ($node in $doc.DocumentElement.ChildNodes) {
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $target   = $node.GetAttribute('Target')
        $external = $node.GetAttribute('TargetMode') -eq 'External'
        $list.Add([pscustomobject]@{
            Id     = $node.GetAttribute('Id')
            Type   = $node.GetAttribute('Type')
            Target = $target
            Part   = if ($external) { '' } else { Resolve-PartName -BaseDirectory $baseDirectory -Target $target }
        })
    }
    return ,$list.ToArray()
}
#endregion

#region Estructura del libro
function Get-WorkbookInfo {
    param([System.IO.Compression.ZipArchive]$Zip)

    $workbookPart = 'xl/workbook.xml'
    foreach ($rel in (Get-Relationships -Zip $Zip -PartName '')) {
        if ($rel.Type -like '*/officeDocument') { $workbookPart = $rel.Part }
    }

    $doc = Open-ZipXml -Zip $Zip -EntryName $workbookPart
    if ($null -eq $doc) { throw 'El archivo no es un libro de Excel válido (no contiene workbook.xml).' }
    $ns   = New-NsManager $doc
    $rels = Get-Relationships -Zip $Zip -PartName $workbookPart
    $byId = @{}
    foreach ($rel in $rels) { $byId[$rel.Id] = $rel }

    $sheets = New-Object System.Collections.Generic.List[object]
    foreach ($node in $doc.SelectNodes('/x:workbook/x:sheets/x:sheet', $ns)) {
        $relationId = $node.GetAttribute('id', $script:NsDocRel)
        if ($byId.ContainsKey($relationId) -and $byId[$relationId].Type -like '*/worksheet') {
            $sheets.Add([pscustomobject]@{ Name = $node.GetAttribute('name'); Part = $byId[$relationId].Part })
        }
    }

    $shared = @($rels | Where-Object { $_.Type -like '*/sharedStrings' })
    $styles = @($rels | Where-Object { $_.Type -like '*/styles' })

    [pscustomobject]@{
        Part              = $workbookPart
        Rels              = $rels
        Sheets            = $sheets.ToArray()
        SharedStringsPart = if ($shared.Count -gt 0) { $shared[0].Part } else { '' }
        StylesPart        = if ($styles.Count -gt 0) { $styles[0].Part } else { '' }
    }
}

function Select-Sheet {
    param($Info, [string]$SheetName)

    $all = @($Info.Sheets)
    if ($all.Count -eq 0) { throw 'El libro no contiene hojas de cálculo editables.' }
    if ($SheetName) {
        foreach ($sheet in $all) { if ($sheet.Name -eq $SheetName) { return $sheet } }
    }
    return $all[0]
}

function Get-SharedStrings {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$Part)

    $list = New-Object System.Collections.Generic.List[string]
    if ($Part) {
        $doc = Open-ZipXml -Zip $Zip -EntryName $Part
        if ($null -ne $doc) {
            $ns = New-NsManager $doc
            foreach ($item in $doc.SelectNodes('/x:sst/x:si', $ns)) { $list.Add((Get-RichText -Node $item -Ns $ns)) }
        }
    }
    return ,$list.ToArray()
}

function Get-RichText {
    param($Node, $Ns)

    $parts = $Node.SelectNodes('.//x:t[not(ancestor::x:rPh)]', $Ns)
    if ($parts.Count -eq 1) { return $parts[0].InnerText }
    $builder = New-Object System.Text.StringBuilder
    foreach ($part in $parts) { [void]$builder.Append($part.InnerText) }
    return $builder.ToString()
}

function Get-DateStyleSet {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$StylesPart)

    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    if (-not $StylesPart) { return ,$set }
    $doc = Open-ZipXml -Zip $Zip -EntryName $StylesPart
    if ($null -eq $doc) { return ,$set }
    $ns = New-NsManager $doc

    $customDates = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($format in $doc.SelectNodes('/x:styleSheet/x:numFmts/x:numFmt', $ns)) {
        $clean = [regex]::Replace($format.GetAttribute('formatCode'), '"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', '')
        $clean = $clean -ireplace 'General', ''
        if ($clean -imatch '[dymhs]') { [void]$customDates.Add([int]$format.GetAttribute('numFmtId')) }
    }

    $index = 0
    foreach ($xf in $doc.SelectNodes('/x:styleSheet/x:cellXfs/x:xf', $ns)) {
        $formatId = 0
        [void][int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$formatId)
        if ($script:BuiltInDateFormats -contains $formatId -or $customDates.Contains($formatId)) { [void]$set.Add($index) }
        $index++
    }
    return ,$set
}
#endregion

#region Lectura de celdas
function Read-SheetRows {
    param([System.Xml.XmlDocument]$Document, $Ns, [string[]]$Shared, $DateStyles)

    $result = New-Object System.Collections.Generic.List[object]
    $rowNumber = 0
    foreach ($row in $Document.SelectNodes('/x:worksheet/x:sheetData/x:row', $Ns)) {
        $rowNumber = Get-RowNumber -Row $row -Previous $rowNumber
        $cells  = @{}
        $column = -1
        foreach ($cell in $row.SelectNodes('x:c', $Ns)) {
            $column = Get-ColumnIndex -Cell $cell -Previous $column
            $value  = Get-CellText -Cell $cell -Ns $Ns -Shared $Shared -DateStyles $DateStyles
            if ($value -ne '') { $cells[$column] = $value }
        }
        $result.Add([pscustomobject]@{ Number = $rowNumber; Cells = $cells })
    }
    return ,$result.ToArray()
}

function Get-SheetLayout {
    <# Encabezado = la fila más completa entre las primeras con datos (salta títulos y subtítulos). #>
    param([object[]]$Rows, [int]$ScanLimit = 20)

    $candidates = @($Rows | Where-Object { $_.Cells.Count -gt 0 } | Select-Object -First $ScanLimit)
    if ($candidates.Count -eq 0) { return @{ HeaderRow = 1; ColOffset = 0 } }

    # Primera fila con al menos el 60 % de las celdas de la fila más completa.
    $widest = ($candidates | ForEach-Object { $_.Cells.Count } | Measure-Object -Maximum).Maximum
    $best = $candidates | Where-Object { $_.Cells.Count -ge ($widest * 0.6) } | Select-Object -First 1

    $offset = [int](($best.Cells.Keys | Measure-Object -Minimum).Minimum)
    return @{ HeaderRow = $best.Number; ColOffset = $offset }
}

function ConvertTo-SheetTable {
    param([object[]]$Rows)

    $layout = Get-SheetLayout -Rows $Rows
    $header = $null
    foreach ($row in $Rows) { if ($row.Number -eq $layout.HeaderRow -and $row.Cells.Count -gt 0) { $header = $row; break } }
    if ($null -eq $header) { return ,(New-AbmDataTable -Columns @('Columna 1')) }

    $offset = $layout.ColOffset
    $maxColumn = $offset
    foreach ($row in $Rows) {
        if ($row.Number -lt $header.Number) { continue }
        foreach ($key in $row.Cells.Keys) { if ($key -gt $maxColumn) { $maxColumn = $key } }
    }

    $width   = $maxColumn - $offset + 1
    $headers = [string[]]::new($width)
    for ($i = 0; $i -lt $width; $i++) {
        $headers[$i] = if ($header.Cells.ContainsKey($offset + $i)) { $header.Cells[$offset + $i] } else { '' }
    }
    $table = New-AbmDataTable -Columns $headers

    foreach ($row in $Rows) {
        if ($row.Number -le $header.Number -or $row.Cells.Count -eq 0) { continue }
        $values = [object[]]::new($width)
        for ($i = 0; $i -lt $width; $i++) {
            $values[$i] = if ($row.Cells.ContainsKey($offset + $i)) { $row.Cells[$offset + $i] } else { '' }
        }
        [void]$table.Rows.Add($values)
    }
    $table.AcceptChanges()
    return ,$table
}

function Get-CellText {
    param($Cell, $Ns, [string[]]$Shared, $DateStyles)

    $type = $Cell.GetAttribute('t')
    if ($type -eq 'inlineStr') {
        $inline = $Cell.SelectSingleNode('x:is', $Ns)
        if ($null -eq $inline) { return '' }
        return (Get-RichText -Node $inline -Ns $Ns)
    }

    $valueNode = $Cell.SelectSingleNode('x:v', $Ns)
    if ($null -eq $valueNode) { return '' }
    $raw = $valueNode.InnerText

    switch ($type) {
        's' {
            $index = 0
            if ([int]::TryParse($raw, [ref]$index) -and $index -lt $Shared.Length) { return $Shared[$index] }
            return ''
        }
        'b'   { if ($raw -eq '1') { return $script:TrueText }; return $script:FalseText }
        'str' { return $raw }
        'e'   { return $raw }
    }

    $style = $Cell.GetAttribute('s')
    if ($style -and $DateStyles.Contains([int]$style)) {
        $number = 0.0
        if ([double]::TryParse($raw, [System.Globalization.NumberStyles]::Float, $script:Invariant, [ref]$number)) {
            return (Format-OADate -Value $number)
        }
    }
    return (Format-CellNumber -Raw $raw)
}

function Format-CellNumber {
    <# Igual que Excel: 15 dígitos significativos (evita ruido como 656.6900000000001). #>
    param([string]$Raw)
    if ($Raw.Length -le 15) { return $Raw }
    $number = 0.0
    if ([double]::TryParse($Raw, [System.Globalization.NumberStyles]::Float, $script:Invariant, [ref]$number)) {
        return $number.ToString('G15', $script:Invariant)
    }
    return $Raw
}

function Format-OADate {
    param([double]$Value)

    try { $date = [datetime]::FromOADate($Value) }
    catch { return $Value.ToString($script:Invariant) }

    if ($Value -lt 1) { return $date.ToString('HH:mm:ss') }
    if ([math]::Abs($Value - [math]::Floor($Value)) -lt 1e-9) { return $date.ToString('dd/MM/yyyy') }
    return $date.ToString('dd/MM/yyyy HH:mm:ss')
}

function Get-RowNumber {
    param($Row, [int]$Previous)
    $number = 0
    if ([int]::TryParse($Row.GetAttribute('r'), [ref]$number)) { return $number }
    return $Previous + 1
}

function Get-ColumnIndex {
    param($Cell, [int]$Previous)
    $letters = $Cell.GetAttribute('r') -replace '[0-9$]', ''
    if ($letters) { return (ConvertFrom-ColumnLetters $letters) }
    return $Previous + 1
}

function ConvertFrom-ColumnLetters {
    param([string]$Letters)
    $number = 0
    foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) { $number = ($number * 26) + ([int]$ch - 64) }
    return $number - 1
}

function ConvertTo-ColumnLetters {
    param([int]$Index)
    $number  = $Index + 1
    $letters = ''
    while ($number -gt 0) {
        $remainder = ($number - 1) % 26
        $letters   = [string][char](65 + $remainder) + $letters
        $number    = [int][math]::Floor(($number - 1) / 26)
    }
    return $letters
}
#endregion

#region Escritura de celdas
function Remove-TableRows {
    <# Quita las filas desde el encabezado hacia abajo y devuelve los estilos por columna para reutilizarlos. #>
    param($SheetData, $Ns, [int]$HeaderRow)

    $styles   = @{ Header = @{}; Data = @{} }
    $obsolete = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    $rowNumber = 0
    foreach ($row in @($SheetData.SelectNodes('x:row', $Ns))) {
        $rowNumber = Get-RowNumber -Row $row -Previous $rowNumber
        if ($rowNumber -lt $HeaderRow) { continue }
        $obsolete.Add($row)
        if ($rowNumber -gt $HeaderRow + 1) { continue }

        $target = if ($rowNumber -eq $HeaderRow) { $styles.Header } else { $styles.Data }
        $column = -1
        foreach ($cell in $row.SelectNodes('x:c', $Ns)) {
            $column = Get-ColumnIndex -Cell $cell -Previous $column
            $style  = $cell.GetAttribute('s')
            if ($style) { $target[$column] = $style }
        }
    }
    foreach ($node in $obsolete) { [void]$SheetData.RemoveChild($node) }
    return $styles
}

function Write-TableRows {
    <# Escribe encabezado + registros. Devuelve el número de la última fila escrita. #>
    param(
        [System.Xml.XmlDocument]$Document, $SheetData, [System.Data.DataTable]$Table,
        [int]$HeaderRow, [int]$ColOffset, [hashtable]$Styles, $DateStyles,
        [string]$DefaultHeaderStyle, [string]$DefaultDataStyle
    )

    $prefix      = $SheetData.Prefix
    $columnCount = $Table.Columns.Count
    $letters     = [string[]]::new($columnCount)
    $dataStyles  = [string[]]::new($columnCount)
    $dateColumns = [bool[]]::new($columnCount)

    for ($i = 0; $i -lt $columnCount; $i++) {
        $column = $ColOffset + $i
        $letters[$i]     = ConvertTo-ColumnLetters $column
        $dataStyles[$i]  = if ($Styles.Data.ContainsKey($column)) { $Styles.Data[$column] } else { $DefaultDataStyle }
        $dateColumns[$i] = ($dataStyles[$i] -ne '') -and $DateStyles.Contains([int]$dataStyles[$i])
    }

    $rowNumber = $HeaderRow
    $headerElement = New-RowElement -Document $Document -Prefix $prefix -Number $rowNumber
    for ($i = 0; $i -lt $columnCount; $i++) {
        $column = $ColOffset + $i
        $style  = if ($Styles.Header.ContainsKey($column)) { $Styles.Header[$column] } else { $DefaultHeaderStyle }
        Add-CellElement -Document $Document -Prefix $prefix -Row $headerElement -Reference ($letters[$i] + $rowNumber) `
                        -Value $Table.Columns[$i].ColumnName -Style $style -AsText
    }
    [void]$SheetData.AppendChild($headerElement)

    foreach ($dataRow in $Table.Rows) {
        if ($dataRow.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
        $rowNumber++
        $element = New-RowElement -Document $Document -Prefix $prefix -Number $rowNumber
        for ($i = 0; $i -lt $columnCount; $i++) {
            Add-CellElement -Document $Document -Prefix $prefix -Row $element -Reference ($letters[$i] + $rowNumber) `
                            -Value ([string]$dataRow[$i]) -Style $dataStyles[$i] -IsDateColumn:($dateColumns[$i])
        }
        [void]$SheetData.AppendChild($element)
    }
    return $rowNumber
}

function New-RowElement {
    param([System.Xml.XmlDocument]$Document, [string]$Prefix, [int]$Number)
    $row = $Document.CreateElement($Prefix, 'row', $script:NsMain)
    $row.SetAttribute('r', [string]$Number)
    return ,$row
}

function Add-CellElement {
    param(
        [System.Xml.XmlDocument]$Document, [string]$Prefix, [System.Xml.XmlElement]$Row,
        [string]$Reference, [string]$Value, [string]$Style, [switch]$AsText, [switch]$IsDateColumn
    )

    $Value = $script:InvalidXmlChars.Replace($Value, '')
    if ($Value -eq '' -and -not $Style) { return }

    $cell = $Document.CreateElement($Prefix, 'c', $script:NsMain)
    $cell.SetAttribute('r', $Reference)
    if ($Style) { $cell.SetAttribute('s', $Style) }

    if ($Value -ne '') {
        $number = if ($AsText) { $null } else { ConvertTo-CellNumber -Value $Value -IsDateColumn:$IsDateColumn }
        if ($null -ne $number) {
            $v = $Document.CreateElement($Prefix, 'v', $script:NsMain)
            $v.InnerText = $number
            [void]$cell.AppendChild($v)
        }
        elseif (-not $AsText -and ($Value -ceq $script:TrueText -or $Value -ceq $script:FalseText)) {
            $cell.SetAttribute('t', 'b')
            $v = $Document.CreateElement($Prefix, 'v', $script:NsMain)
            $v.InnerText = if ($Value -ceq $script:TrueText) { '1' } else { '0' }
            [void]$cell.AppendChild($v)
        }
        else {
            $cell.SetAttribute('t', 'inlineStr')
            $is = $Document.CreateElement($Prefix, 'is', $script:NsMain)
            $t  = $Document.CreateElement($Prefix, 't', $script:NsMain)
            if ($Value -ne $Value.Trim()) { [void]$t.SetAttribute('space', $script:NsXml, 'preserve') }
            $t.InnerText = $Value
            [void]$is.AppendChild($t)
            [void]$cell.AppendChild($is)
        }
    }
    [void]$Row.AppendChild($cell)
}

function ConvertTo-CellNumber {
    <# Devuelve el valor numérico en formato invariante, o $null si debe guardarse como texto. #>
    param([string]$Value, [switch]$IsDateColumn)

    if ($script:NumberPattern.IsMatch($Value)) {
        $number = 0.0
        if ([double]::TryParse($Value, [System.Globalization.NumberStyles]::Float, $script:Invariant, [ref]$number)) {
            return $number.ToString('R', $script:Invariant)
        }
    }
    if ($IsDateColumn) {
        $date = [datetime]::MinValue
        if ([datetime]::TryParseExact($Value, $script:DateInputFormats, $script:Invariant, [System.Globalization.DateTimeStyles]::None, [ref]$date)) {
            return $date.ToOADate().ToString('R', $script:Invariant)
        }
        $time = [timespan]::Zero
        if ($Value -match '^\d{1,2}:\d{2}(:\d{2})?$' -and [timespan]::TryParse($Value, $script:Invariant, [ref]$time)) {
            return $time.TotalDays.ToString('R', $script:Invariant)
        }
    }
    return $null
}

function Update-SheetRanges {
    param($Document, $Ns, [string]$FirstColumn, [string]$LastColumn, [int]$HeaderRow, [int]$LastRow)

    $dimension = $Document.SelectSingleNode('/x:worksheet/x:dimension', $Ns)
    if ($null -ne $dimension) { $dimension.SetAttribute('ref', ('A1:{0}{1}' -f $LastColumn, $LastRow)) }

    $autoFilter = $Document.SelectSingleNode('/x:worksheet/x:autoFilter', $Ns)
    if ($null -ne $autoFilter) { $autoFilter.SetAttribute('ref', ('{0}{1}:{2}{3}' -f $FirstColumn, $HeaderRow, $LastColumn, $LastRow)) }
}

function Update-SheetTables {
    <# Ajusta el rango de la tabla de Excel (ListObject) si la hoja tiene exactamente una. #>
    param([System.IO.Compression.ZipArchive]$Zip, [string]$SheetPart, [string]$Range)

    $relationships = Get-Relationships -Zip $Zip -PartName $SheetPart
    $tables = @($relationships | Where-Object { $_.Type -like '*/table' -and $_.Part })
    if ($tables.Count -ne 1) { return }

    $doc = Open-ZipXml -Zip $Zip -EntryName $tables[0].Part
    if ($null -eq $doc) { return }
    $doc.DocumentElement.SetAttribute('ref', $Range)
    $autoFilter = $doc.SelectSingleNode('/x:table/x:autoFilter', (New-NsManager $doc))
    if ($null -ne $autoFilter) { $autoFilter.SetAttribute('ref', $Range) }
    Save-ZipXml -Zip $Zip -EntryName $tables[0].Part -Document $doc
}

function Remove-CalcChain {
    <# La cadena de cálculo queda desactualizada al reescribir celdas; Excel la regenera sola. #>
    param([System.IO.Compression.ZipArchive]$Zip, $Info)

    $calc = @($Info.Rels | Where-Object { $_.Type -like '*/calcChain' })
    if ($calc.Count -eq 0) { return }

    $entry = $Zip.GetEntry($calc[0].Part)
    if ($null -ne $entry) { $entry.Delete() }

    $relsPath = Get-RelsPath $Info.Part
    $relsDoc  = Open-ZipXml -Zip $Zip -EntryName $relsPath
    Remove-ChildByAttribute -Document $relsDoc -Attribute 'Id' -Value $calc[0].Id
    Save-ZipXml -Zip $Zip -EntryName $relsPath -Document $relsDoc

    $types = Open-ZipXml -Zip $Zip -EntryName '[Content_Types].xml'
    if ($null -ne $types) {
        Remove-ChildByAttribute -Document $types -Attribute 'PartName' -Value ('/' + $calc[0].Part)
        Save-ZipXml -Zip $Zip -EntryName '[Content_Types].xml' -Document $types
    }
}

function Remove-ChildByAttribute {
    param([System.Xml.XmlDocument]$Document, [string]$Attribute, [string]$Value)
    foreach ($node in @($Document.DocumentElement.ChildNodes)) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element -and $node.GetAttribute($Attribute) -eq $Value) {
            [void]$Document.DocumentElement.RemoveChild($node)
        }
    }
}
#endregion

#region Libro nuevo
function ConvertTo-SafeSheetName {
    param([string]$Name)
    $clean = ($Name -replace '[\[\]\:\*\?\/\\]', '').Trim().Trim("'")
    if (-not $clean) { $clean = 'Datos' }
    if ($clean.Length -gt 31) { $clean = $clean.Substring(0, 31) }
    return $clean
}

function Get-NewWorkbookParts {
    param([string]$SheetName, [int]$ColumnCount)

    $xmlHeader  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $main       = $script:NsMain
    $rel        = $script:NsDocRel
    $pkgRel     = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $lastColumn = ConvertTo-ColumnLetters ($ColumnCount - 1)
    $columns    = (1..$ColumnCount | ForEach-Object { '<col min="{0}" max="{0}" width="24" customWidth="1"/>' -f $_ }) -join ''
    $sheetTitle = [System.Security.SecurityElement]::Escape($SheetName)

    $parts = [ordered]@{}
    $parts['[Content_Types].xml'] = $xmlHeader +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>'

    $parts['_rels/.rels'] = $xmlHeader +
        "<Relationships xmlns=`"$pkgRel`">" +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'

    $parts['xl/workbook.xml'] = $xmlHeader +
        "<workbook xmlns=`"$main`" xmlns:r=`"$rel`"><sheets><sheet name=`"$sheetTitle`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"

    $parts['xl/_rels/workbook.xml.rels'] = $xmlHeader +
        "<Relationships xmlns=`"$pkgRel`">" +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>'

    $parts['xl/styles.xml'] = $xmlHeader +
        "<styleSheet xmlns=`"$main`">" +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF112240"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
        '<border><left style="thin"><color rgb="FFDDE3EC"/></left><right style="thin"><color rgb="FFDDE3EC"/></right>' +
        '<top style="thin"><color rgb="FFDDE3EC"/></top><bottom style="thin"><color rgb="FFDDE3EC"/></bottom><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>'

    $parts['xl/worksheets/sheet1.xml'] = $xmlHeader +
        "<worksheet xmlns=`"$main`" xmlns:r=`"$rel`">" +
        "<dimension ref=`"A1:$($lastColumn)1`"/>" +
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' +
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="15"/>' +
        "<cols>$columns</cols>" +
        '<sheetData/>' +
        "<autoFilter ref=`"A1:$($lastColumn)1`"/>" +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
        '</worksheet>'

    return $parts
}
#endregion

Export-ModuleMember -Function Read-ExcelSheet, Save-ExcelSheet, New-ExcelWorkbook
