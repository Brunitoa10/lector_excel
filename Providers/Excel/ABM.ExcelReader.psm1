Set-StrictMode -Version Latest

<#
    Lectura de una hoja: celdas -> filas dispersas -> DataTable.
    El encabezado es la primera fila (entre las primeras con datos) que tiene al menos el 60 % de las celdas
    de la fila más completa; así se saltean títulos y subtítulos.
#>

$script:HeaderScanLimit = 20
$script:HeaderRatio     = 0.6

function Get-AbmSharedStrings {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [AllowEmptyString()][string]$Part)

    $list = New-Object System.Collections.Generic.List[string]
    $doc  = if ($Part) { Open-AbmZipXml -Zip $Zip -EntryName $Part } else { $null }
    if ($null -ne $doc) {
        $ns = New-AbmXmlNamespace $doc
        foreach ($item in $doc.SelectNodes('/x:sst/x:si', $ns)) { $list.Add((Get-RichText -Node $item -Ns $ns)) }
    }
    return ,$list.ToArray()
}

function Read-AbmSheetRows {
    <# Devuelve filas { Number; Cells = @{ índiceColumna = texto } } sólo con celdas no vacías. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Xml.XmlDocument]$Document,
        [AllowEmptyCollection()][string[]]$Shared,
        [Parameter(Mandatory)]$DateStyles
    )

    $ns = New-AbmXmlNamespace $Document
    $result = New-Object System.Collections.Generic.List[object]
    $rowNumber = 0
    foreach ($row in $Document.SelectNodes('/x:worksheet/x:sheetData/x:row', $ns)) {
        $rowNumber = Get-AbmRowNumber -Row $row -Previous $rowNumber
        $cells  = @{}
        $column = -1
        foreach ($cell in $row.SelectNodes('x:c', $ns)) {
            $column = Get-AbmColumnIndex -Cell $cell -Previous $column
            $value  = Get-CellText -Cell $cell -Ns $ns -Shared $Shared -DateStyles $DateStyles
            if ($value -ne '') { $cells[$column] = $value }
        }
        $result.Add([pscustomobject]@{ Number = $rowNumber; Cells = $cells })
    }
    return ,$result.ToArray()
}

function Get-AbmSheetLayout {
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Rows)

    $candidates = @($Rows | Where-Object { $_.Cells.Count -gt 0 } | Select-Object -First $script:HeaderScanLimit)
    if ($candidates.Count -eq 0) { return @{ HeaderRow = 1; ColOffset = 0 } }

    $widest = ($candidates | ForEach-Object { $_.Cells.Count } | Measure-Object -Maximum).Maximum
    $header = $candidates | Where-Object { $_.Cells.Count -ge ($widest * $script:HeaderRatio) } | Select-Object -First 1
    $offset = [int](($header.Cells.Keys | Measure-Object -Minimum).Minimum)
    return @{ HeaderRow = $header.Number; ColOffset = $offset }
}

function ConvertTo-AbmSheetTable {
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][hashtable]$Layout)

    $header = @($Rows | Where-Object { $_.Number -eq $Layout.HeaderRow -and $_.Cells.Count -gt 0 }) | Select-Object -First 1
    if ($null -eq $header) { return ,(New-AbmDataTable -Columns @('Columna 1')) }

    $offset = $Layout.ColOffset
    $width  = (Get-LastColumn -Rows $Rows -FromRow $header.Number -Minimum $offset) - $offset + 1
    $table  = New-AbmDataTable -Columns (Get-RowValues -Row $header -Offset $offset -Width $width)
    foreach ($row in $Rows) {
        if ($row.Number -le $header.Number -or $row.Cells.Count -eq 0) { continue }
        [void]$table.Rows.Add([object[]](Get-RowValues -Row $row -Offset $offset -Width $width))
    }
    $table.AcceptChanges()
    return ,$table
}

#region Internos
function Get-LastColumn {
    param([object[]]$Rows, [int]$FromRow, [int]$Minimum)
    $max = $Minimum
    foreach ($row in $Rows) {
        if ($row.Number -lt $FromRow) { continue }
        foreach ($key in $row.Cells.Keys) { if ($key -gt $max) { $max = $key } }
    }
    return $max
}

function Get-RowValues {
    param($Row, [int]$Offset, [int]$Width)
    $values = [string[]]::new($Width)
    for ($i = 0; $i -lt $Width; $i++) {
        $values[$i] = if ($Row.Cells.ContainsKey($Offset + $i)) { $Row.Cells[$Offset + $i] } else { '' }
    }
    return ,$values
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
        's'   { return (Get-SharedText -Shared $Shared -Raw $raw) }
        'b'   { return (ConvertTo-AbmBooleanText -Raw $raw) }
        'str' { return $raw }
        'e'   { return $raw }
    }

    $style = $Cell.GetAttribute('s')
    if ($style -and $DateStyles.Contains([int]$style)) {
        $date = ConvertFrom-AbmStoredDate -Raw $raw
        if ($null -ne $date) { return $date }
    }
    return (Format-AbmCellNumber -Raw $raw)
}

function Get-SharedText {
    param([string[]]$Shared, [string]$Raw)
    $index = 0
    if ([int]::TryParse($Raw, [ref]$index) -and $index -lt $Shared.Length) { return $Shared[$index] }
    return ''
}

function Get-RichText {
    <# Concatena los fragmentos <t> (texto enriquecido), ignorando la fonética (rPh). #>
    param($Node, $Ns)
    $parts = $Node.SelectNodes('.//x:t[not(ancestor::x:rPh)]', $Ns)
    if ($parts.Count -eq 1) { return $parts[0].InnerText }
    $builder = New-Object System.Text.StringBuilder
    foreach ($part in $parts) { [void]$builder.Append($part.InnerText) }
    return $builder.ToString()
}
#endregion

Export-ModuleMember -Function Get-AbmSharedStrings, Read-AbmSheetRows, Get-AbmSheetLayout, ConvertTo-AbmSheetTable
