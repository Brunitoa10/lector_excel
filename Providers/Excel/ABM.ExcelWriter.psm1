Set-StrictMode -Version Latest

<#
    Escritura de la tabla en el XML de una hoja: quita las filas viejas (desde el encabezado),
    recuerda el estilo de cada columna y escribe encabezado + registros con esos estilos.
#>

$script:Ns = Get-AbmSpreadsheetNamespace
$script:InvalidXmlChars = New-Object System.Text.RegularExpressions.Regex('[\x00-\x08\x0B\x0C\x0E-\x1F]')

function Remove-AbmTableRows {
    <# Quita las filas desde $HeaderRow hacia abajo y devuelve @{ Header = @{col=estilo}; Data = @{col=estilo} }. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Xml.XmlElement]$SheetData, [Parameter(Mandatory)]$Ns, [int]$HeaderRow)

    $styles   = @{ Header = @{}; Data = @{} }
    $obsolete = New-Object System.Collections.Generic.List[System.Xml.XmlNode]
    $rowNumber = 0
    foreach ($row in @($SheetData.SelectNodes('x:row', $Ns))) {
        $rowNumber = Get-AbmRowNumber -Row $row -Previous $rowNumber
        if ($rowNumber -lt $HeaderRow) { continue }
        $obsolete.Add($row)
        if ($rowNumber -eq $HeaderRow)         { Read-ColumnStyles -Row $row -Ns $Ns -Target $styles.Header }
        elseif ($rowNumber -eq $HeaderRow + 1) { Read-ColumnStyles -Row $row -Ns $Ns -Target $styles.Data }
    }
    foreach ($node in $obsolete) { [void]$SheetData.RemoveChild($node) }
    return $styles
}

function Write-AbmTableRows {
    <# Escribe encabezado + registros. Devuelve el número de la última fila escrita. #>
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory)][System.Xml.XmlElement]$SheetData,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [int]$HeaderRow, [int]$ColOffset,
        [Parameter(Mandatory)][hashtable]$Styles, [Parameter(Mandatory)]$DateStyles,
        [string]$DefaultHeaderStyle = '', [string]$DefaultDataStyle = ''
    )

    $layout  = Get-ColumnLayout -Table $Table -ColOffset $ColOffset -Styles $Styles -DateStyles $DateStyles `
                                -DefaultHeaderStyle $DefaultHeaderStyle -DefaultDataStyle $DefaultDataStyle
    $context = @{ Document = $SheetData.OwnerDocument; Prefix = $SheetData.Prefix }

    $rowNumber = $HeaderRow
    $names = @(foreach ($column in $Table.Columns) { $column.ColumnName })
    [void]$SheetData.AppendChild((New-RowElement -Context $context -Number $rowNumber -Values $names -Layout $layout -IsHeader))

    foreach ($dataRow in $Table.Rows) {
        if ($dataRow.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
        $rowNumber++
        $values = @(for ($i = 0; $i -lt $Table.Columns.Count; $i++) { [string]$dataRow[$i] })
        [void]$SheetData.AppendChild((New-RowElement -Context $context -Number $rowNumber -Values $values -Layout $layout))
    }
    return $rowNumber
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
    <# Precalcula por columna: letra, estilo de encabezado, estilo de datos y si es columna de fecha. #>
    param($Table, [int]$ColOffset, [hashtable]$Styles, $DateStyles, [string]$DefaultHeaderStyle, [string]$DefaultDataStyle)

    $layout = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
        $column = $ColOffset + $i
        $data   = if ($Styles.Data.ContainsKey($column)) { $Styles.Data[$column] } else { $DefaultDataStyle }
        $layout.Add([pscustomobject]@{
            Letters     = ConvertTo-AbmColumnLetters $column
            HeaderStyle = if ($Styles.Header.ContainsKey($column)) { $Styles.Header[$column] } else { $DefaultHeaderStyle }
            DataStyle   = $data
            IsDate      = ($data -ne '') -and $DateStyles.Contains([int]$data)
        })
    }
    return ,$layout.ToArray()
}

function New-RowElement {
    param([hashtable]$Context, [int]$Number, [string[]]$Values, [object[]]$Layout, [switch]$IsHeader)

    $row = New-XmlElement -Context $Context -Name 'row'
    $row.SetAttribute('r', [string]$Number)
    for ($i = 0; $i -lt $Layout.Count; $i++) {
        $column = $Layout[$i]
        $style  = if ($IsHeader) { $column.HeaderStyle } else { $column.DataStyle }
        $cell   = New-CellElement -Context $Context -Reference ($column.Letters + $Number) -Value $Values[$i] `
                                  -Style $style -AsText:$IsHeader -IsDateColumn:($column.IsDate)
        if ($null -ne $cell) { [void]$row.AppendChild($cell) }
    }
    return ,$row
}

function New-CellElement {
    param([hashtable]$Context, [string]$Reference, [string]$Value, [string]$Style, [switch]$AsText, [switch]$IsDateColumn)

    $Value = $script:InvalidXmlChars.Replace($Value, '')
    if ($Value -eq '' -and -not $Style) { return $null }

    $cell = New-XmlElement -Context $Context -Name 'c'
    $cell.SetAttribute('r', $Reference)
    if ($Style) { $cell.SetAttribute('s', $Style) }
    if ($Value -eq '') { return ,$cell }

    $number  = if ($AsText) { $null } else { ConvertTo-AbmCellNumber -Value $Value -IsDateColumn:$IsDateColumn }
    $boolean = if ($AsText) { $null } else { ConvertTo-AbmStoredBoolean -Value $Value }
    if ($null -ne $number)      { Add-ValueNode -Context $Context -Cell $cell -Value $number }
    elseif ($null -ne $boolean) { $cell.SetAttribute('t', 'b'); Add-ValueNode -Context $Context -Cell $cell -Value $boolean }
    else                        { Add-InlineString -Context $Context -Cell $cell -Value $Value }
    return ,$cell
}

function Add-ValueNode {
    param([hashtable]$Context, $Cell, [string]$Value)
    $v = New-XmlElement -Context $Context -Name 'v'
    $v.InnerText = $Value
    [void]$Cell.AppendChild($v)
}

function Add-InlineString {
    param([hashtable]$Context, $Cell, [string]$Value)
    $Cell.SetAttribute('t', 'inlineStr')
    $is = New-XmlElement -Context $Context -Name 'is'
    $t  = New-XmlElement -Context $Context -Name 't'
    if ($Value -ne $Value.Trim()) { [void]$t.SetAttribute('space', $script:Ns.Xml, 'preserve') }
    $t.InnerText = $Value
    [void]$is.AppendChild($t)
    [void]$Cell.AppendChild($is)
}

function New-XmlElement {
    param([hashtable]$Context, [string]$Name)
    return ,$Context.Document.CreateElement($Context.Prefix, $Name, $script:Ns.Main)
}
#endregion

Export-ModuleMember -Function Remove-AbmTableRows, Write-AbmTableRows
