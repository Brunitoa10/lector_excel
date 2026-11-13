Set-StrictMode -Version Latest

<#
    Escritura de la tabla en una hoja, optimizada para archivos grandes:
      1. Reset-AbmSheetData: conserva las filas por encima del encabezado, recuerda el estilo de cada
         columna y deja una marca donde irán las filas nuevas (sin recorrer toda la hoja).
      2. ConvertTo-AbmSheetRowsXml: genera el XML de encabezado + registros en un StringBuilder,
         en un bucle sin llamadas a funciones por celda.
      3. Complete-AbmSheetXml: reemplaza la marca por las filas y devuelve el XML final de la hoja.
#>

$script:Ns     = Get-AbmSpreadsheetNamespace
$script:Codec  = Get-AbmCellCodec
$script:Marker = 'ABM:ROWS'
$script:InvalidXmlChars = New-Object System.Text.RegularExpressions.Regex('[\x00-\x08\x0B\x0C\x0E-\x1F]')
$script:InvalidCharSet  = [char[]](@(0..8) + @(11, 12) + @(14..31))

function Reset-AbmSheetData {
    <# Devuelve @{ Header = @{col=estilo}; Data = @{col=estilo} } y deja la marca en sheetData. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Xml.XmlElement]$SheetData, [Parameter(Mandatory)]$Ns, [int]$HeaderRow)

    $styles = @{ Header = @{}; Data = @{} }
    $kept   = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    $rowNumber = 0
    foreach ($row in $SheetData.SelectNodes('x:row', $Ns)) {
        $rowNumber = Get-AbmRowNumber -Row $row -Previous $rowNumber
        if ($rowNumber -lt $HeaderRow) { $kept.Add($row); continue }
        if ($rowNumber -eq $HeaderRow)     { Read-ColumnStyles -Row $row -Ns $Ns -Target $styles.Header; continue }
        if ($rowNumber -eq $HeaderRow + 1) { Read-ColumnStyles -Row $row -Ns $Ns -Target $styles.Data }
        break
    }

    $SheetData.RemoveAll()
    foreach ($row in $kept) { [void]$SheetData.AppendChild($row) }
    [void]$SheetData.AppendChild($SheetData.OwnerDocument.CreateComment($script:Marker))
    return $styles
}

function ConvertTo-AbmSheetRowsXml {
    <# Devuelve @{ Xml = <filas>; LastRow = <número de la última fila escrita> }. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [int]$HeaderRow, [int]$ColOffset, [AllowEmptyString()][string]$Prefix = '',
        [Parameter(Mandatory)][hashtable]$Styles, [Parameter(Mandatory)]$DateStyles,
        [string]$DefaultHeaderStyle = '', [string]$DefaultDataStyle = ''
    )

    $layout  = Get-ColumnLayout -Table $Table -ColOffset $ColOffset -Styles $Styles -DateStyles $DateStyles `
                                -DefaultHeaderStyle $DefaultHeaderStyle -DefaultDataStyle $DefaultDataStyle
    $tags    = Get-Tags -Prefix $Prefix
    $builder = New-Object System.Text.StringBuilder (1MB)
    $codec   = $script:Codec
    $invalid = $script:InvalidCharSet
    $count   = $layout.Count
    $number  = 0.0
    $date    = [datetime]::MinValue

    $rowNumber = $HeaderRow
    [void]$builder.Append($tags.RowOpen + $rowNumber + '">')
    for ($i = 0; $i -lt $count; $i++) {
        $column = $layout[$i]
        $text = [System.Security.SecurityElement]::Escape($script:InvalidXmlChars.Replace($Table.Columns[$i].ColumnName, ''))
        [void]$builder.Append($tags.CellOpen + $column.Letters + $rowNumber + $column.HeaderAttr + ' t="inlineStr">' + $tags.InlineOpen + $text + $tags.InlineClose)
    }
    [void]$builder.Append($tags.RowClose)

    foreach ($dataRow in $Table.Rows) {
        if ($dataRow.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
        $rowNumber++
        $items = $dataRow.ItemArray
        [void]$builder.Append($tags.RowOpen + $rowNumber + '">')
        for ($i = 0; $i -lt $count; $i++) {
            $column = $layout[$i]
            $value  = [string]$items[$i]
            if ($value.IndexOfAny($invalid) -ge 0) { $value = $script:InvalidXmlChars.Replace($value, '') }
            $open = $tags.CellOpen + $column.Letters + $rowNumber + $column.DataAttr

            if ($value.Length -eq 0) { if ($column.DataAttr.Length -gt 1) { [void]$builder.Append($open + '/>') }; continue }

            # Misma regla que ConvertTo-AbmCellNumber / ConvertTo-AbmStoredBoolean, sin llamar funciones por celda.
            if ($codec.NumberPattern.IsMatch($value) -and [double]::TryParse($value, [System.Globalization.NumberStyles]::Float, $codec.Culture, [ref]$number)) {
                [void]$builder.Append($open + '>' + $tags.ValueOpen + $number.ToString('R', $codec.Culture) + $tags.ValueClose)
            }
            elseif ($column.IsDate -and [datetime]::TryParseExact($value, $codec.DateFormats, $codec.Culture, [System.Globalization.DateTimeStyles]::None, [ref]$date)) {
                [void]$builder.Append($open + '>' + $tags.ValueOpen + $date.ToOADate().ToString('R', $codec.Culture) + $tags.ValueClose)
            }
            elseif ($value -ceq $codec.TrueText -or $value -ceq $codec.FalseText) {
                $flag = if ($value -ceq $codec.TrueText) { '1' } else { '0' }
                [void]$builder.Append($open + ' t="b">' + $tags.ValueOpen + $flag + $tags.ValueClose)
            }
            else {
                $space = if ($value.Length -ne $value.Trim().Length) { ' xml:space="preserve"' } else { '' }
                [void]$builder.Append($open + ' t="inlineStr">' + $tags.IsOpen + $tags.TOpen + $space + '>' + [System.Security.SecurityElement]::Escape($value) + $tags.InlineClose)
            }
        }
        [void]$builder.Append($tags.RowClose)
    }
    return @{ Xml = $builder.ToString(); LastRow = $rowNumber }
}

function Complete-AbmSheetXml {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][System.Xml.XmlDocument]$Document, [Parameter(Mandatory)][AllowEmptyString()][string]$RowsXml)
    return $Document.OuterXml.Replace("<!--$($script:Marker)-->", $RowsXml)
}

#region Internos
function Read-ColumnStyles {
    param($Row, $Ns, [hashtable]$Target)
    $column = -1
    foreach ($cell in $Row.SelectNodes('x:c', $Ns)) {
        $column = Get-AbmColumnIndex -Cell $cell -Previous $column
        $style  = $cell.GetAttribute('s')
        if ($style) { $Target[$column] = $style }
    }
}

function Get-ColumnLayout {
    <# Precalcula por columna: letras, atributo de estilo del encabezado y de los datos, y si es fecha.
       El estilo se busca por la posición ORIGINAL de la columna (AbmSourceIndex), no por la actual. #>
    param($Table, [int]$ColOffset, [hashtable]$Styles, $DateStyles, [string]$DefaultHeaderStyle, [string]$DefaultDataStyle)

    # Columna nueva: encabezado con el estilo del primer encabezado existente y datos sin estilo propio.
    $newHeader = if ($Styles.Header.ContainsKey($ColOffset)) { $Styles.Header[$ColOffset] } else { $DefaultHeaderStyle }
    $layout = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
        $source = $Table.Columns[$i].ExtendedProperties['AbmSourceIndex']
        if ($null -ne $source) {
            $key    = $ColOffset + [int]$source
            $header = if ($Styles.Header.ContainsKey($key)) { $Styles.Header[$key] } else { $DefaultHeaderStyle }
            $data   = if ($Styles.Data.ContainsKey($key)) { $Styles.Data[$key] } else { $DefaultDataStyle }
        }
        else { $header = $newHeader; $data = $DefaultDataStyle }
        $layout.Add([pscustomobject]@{
            Letters    = ConvertTo-AbmColumnLetters ($ColOffset + $i)
            HeaderAttr = Get-StyleAttribute $header
            DataAttr   = Get-StyleAttribute $data
            IsDate     = ($data -ne '') -and $DateStyles.Contains([int]$data)
        })
    }
    return ,$layout.ToArray()
}

function Get-StyleAttribute {
    <# Cierra el atributo r="…" y agrega s="…" si hay estilo. #>
    param([string]$Style)
    if ($Style) { return '" s="' + $Style + '"' }
    return '"'
}

function Get-Tags {
    param([string]$Prefix)
    $p = if ($Prefix) { "$Prefix`:" } else { '' }
    return [pscustomobject]@{
        RowOpen     = "<${p}row r=`""
        RowClose    = "</${p}row>"
        CellOpen    = "<${p}c r=`""
        ValueOpen   = "<${p}v>"
        ValueClose  = "</${p}v></${p}c>"
        IsOpen      = "<${p}is>"
        TOpen       = "<${p}t"
        InlineOpen  = "<${p}is><${p}t>"
        InlineClose = "</${p}t></${p}is></${p}c>"
    }
}
#endregion

Export-ModuleMember -Function Reset-AbmSheetData, ConvertTo-AbmSheetRowsXml, Complete-AbmSheetXml
