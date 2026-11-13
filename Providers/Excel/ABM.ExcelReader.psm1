Set-StrictMode -Version Latest

<#
    Lectura de una hoja en streaming (System.Xml.XmlReader): recorre el XML una sola vez, sin cargarlo
    completo en memoria ni usar XPath por celda. El bucle principal no llama funciones de PowerShell
    (son costosas): resuelve columnas y fechas con cachés y copia filas por bloque al DataTable.

    El encabezado es la primera fila (entre las primeras con datos) que tiene al menos el 60 % de las
    celdas de la fila más completa; así se saltean títulos y subtítulos.
#>

$script:HeaderScanLimit = 20
$script:HeaderRatio     = 0.6
$script:ProgressEvery   = 5000
$script:RefDigits       = [char[]]'0123456789$'
$script:ElementNode     = [System.Xml.XmlNodeType]::Element
$script:EndElementNode  = [System.Xml.XmlNodeType]::EndElement

function Get-AbmSharedStrings {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [AllowEmptyString()][string]$Part)

    $list = New-Object System.Collections.Generic.List[string]
    if (-not $Part -or $null -eq $Zip.GetEntry($Part)) { return ,$list.ToArray() }

    $reader  = Open-XmlStream -Zip $Zip -Part $Part
    $builder = New-Object System.Text.StringBuilder
    $inPhonetic = $false
    try {
        while ($reader.Read()) {
            if ($reader.NodeType -eq $script:ElementNode) {
                $name = $reader.LocalName
                if ($name -eq 't') { if (-not $inPhonetic -and -not $reader.IsEmptyElement) { [void]$builder.Append($reader.ReadString()) } }
                elseif ($name -eq 'si') { [void]$builder.Clear(); if ($reader.IsEmptyElement) { $list.Add('') } }
                elseif ($name -eq 'rPh' -and -not $reader.IsEmptyElement) { $inPhonetic = $true }
            }
            elseif ($reader.NodeType -eq $script:EndElementNode) {
                if ($reader.LocalName -eq 'si') { $list.Add($builder.ToString()) }
                elseif ($reader.LocalName -eq 'rPh') { $inPhonetic = $false }
            }
        }
    }
    finally { $reader.Dispose() }
    return ,$list.ToArray()
}

function Read-AbmSheetGrid {
    <# Devuelve la hoja como listas paralelas: Numbers (nro. de fila), Values (string[] densos) y Filled (celdas con dato). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory)][string]$Part,
        [AllowEmptyCollection()][string[]]$Shared,
        [Parameter(Mandatory)]$DateStyles,
        [scriptblock]$OnProgress
    )

    $dateSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($index in $DateStyles) { [void]$dateSet.Add([string]$index) }
    $columnCache = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $dateCache   = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    $trueText    = ConvertTo-AbmBooleanText -Raw '1'
    $falseText   = ConvertTo-AbmBooleanText -Raw '0'

    $numbers = New-Object System.Collections.Generic.List[int]
    $values  = New-Object 'System.Collections.Generic.List[string[]]'
    $filled  = New-Object System.Collections.Generic.List[int]
    $row     = New-Object System.Collections.Generic.List[string]
    $inline  = New-Object System.Text.StringBuilder
    $rowNumber = 0; $rowFilled = 0; $column = -1
    $cellType = $null; $cellStyle = $null; $raw = $null; $text = ''; $position = 0
    $inCell = $false; $inPhonetic = $false

    $reader = Open-XmlStream -Zip $Zip -Part $Part
    try {
        while ($reader.Read()) {
            if ($reader.NodeType -eq $script:ElementNode) {
                $name = $reader.LocalName
                if ($name -eq 'c') {
                    $ref = $reader.GetAttribute('r')
                    if ($ref) {
                        $letters = $ref.TrimEnd($script:RefDigits)
                        if (-not $columnCache.TryGetValue($letters, [ref]$position)) {
                            $position = ConvertFrom-AbmColumnLetters $letters
                            $columnCache[$letters] = $position
                        }
                        $column = $position
                    }
                    else { $column++ }
                    $cellType = $reader.GetAttribute('t'); $cellStyle = $reader.GetAttribute('s'); $raw = $null
                    [void]$inline.Clear()
                    $inCell = -not $reader.IsEmptyElement
                }
                elseif ($name -eq 'v') { if ($inCell -and -not $reader.IsEmptyElement) { $raw = $reader.ReadString() } }
                elseif ($name -eq 't') { if ($inCell -and -not $inPhonetic -and -not $reader.IsEmptyElement) { [void]$inline.Append($reader.ReadString()) } }
                elseif ($name -eq 'rPh') { if (-not $reader.IsEmptyElement) { $inPhonetic = $true } }
                elseif ($name -eq 'row') {
                    $attribute = $reader.GetAttribute('r')
                    if ($attribute -and [int]::TryParse($attribute, [ref]$position)) { $rowNumber = $position } else { $rowNumber++ }
                    $row.Clear(); $rowFilled = 0; $column = -1
                }
                continue
            }
            if ($reader.NodeType -ne $script:EndElementNode) { continue }

            $name = $reader.LocalName
            if ($name -eq 'c') {
                $inCell = $false
                # --- valor visible de la celda (inline para evitar llamadas por celda) ---
                if ($cellType -eq 'inlineStr')                 { $text = $inline.ToString() }
                elseif ($null -eq $raw)                        { $text = '' }
                elseif ($cellType -eq 's')                     { $text = ''; if ([int]::TryParse($raw, [ref]$position) -and $position -lt $Shared.Length) { $text = $Shared[$position] } }
                elseif ($cellType -eq 'b')                     { if ($raw -eq '1') { $text = $trueText } else { $text = $falseText } }
                elseif ($cellType -eq 'str' -or $cellType -eq 'e') { $text = $raw }
                elseif ($cellStyle -and $dateSet.Contains($cellStyle)) {
                    if (-not $dateCache.TryGetValue($raw, [ref]$text)) {
                        $text = ConvertFrom-AbmStoredDate -Raw $raw
                        if ($null -eq $text) { $text = $raw }
                        $dateCache[$raw] = $text
                    }
                }
                elseif ($raw.Length -gt 15)                    { $text = Format-AbmCellNumber -Raw $raw }
                else                                           { $text = $raw }

                if ($text.Length -gt 0) {
                    while ($row.Count -lt $column) { $row.Add('') }
                    if ($row.Count -eq $column) { $row.Add($text) } else { $row[$column] = $text }
                    $rowFilled++
                }
            }
            elseif ($name -eq 'row') {
                $numbers.Add($rowNumber); $values.Add($row.ToArray()); $filled.Add($rowFilled)
                if ($null -ne $OnProgress -and ($numbers.Count % $script:ProgressEvery) -eq 0) { & $OnProgress $numbers.Count }
            }
            elseif ($name -eq 'rPh') { $inPhonetic = $false }
            elseif ($name -eq 'sheetData') { break }
        }
    }
    finally { $reader.Dispose() }

    return [pscustomobject]@{ Numbers = $numbers; Values = $values; Filled = $filled }
}

function Get-AbmSheetLayout {
    <# Devuelve @{ HeaderRow; ColOffset; HeaderIndex } (HeaderIndex = -1 si la hoja está vacía). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Grid)

    $candidates = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $Grid.Filled.Count -and $candidates.Count -lt $script:HeaderScanLimit; $i++) {
        if ($Grid.Filled[$i] -gt 0) { $candidates.Add($i) }
    }
    if ($candidates.Count -eq 0) { return @{ HeaderRow = 1; ColOffset = 0; HeaderIndex = -1 } }

    $widest = 0
    foreach ($i in $candidates) { if ($Grid.Filled[$i] -gt $widest) { $widest = $Grid.Filled[$i] } }
    $header = $candidates[0]
    foreach ($i in $candidates) { if ($Grid.Filled[$i] -ge ($widest * $script:HeaderRatio)) { $header = $i; break } }

    $cells  = $Grid.Values[$header]
    $offset = 0
    while ($offset -lt $cells.Length -and $cells[$offset] -eq '') { $offset++ }
    return @{ HeaderRow = $Grid.Numbers[$header]; ColOffset = $offset; HeaderIndex = $header }
}

function ConvertTo-AbmSheetTable {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Grid, [Parameter(Mandatory)][hashtable]$Layout)

    $header = $Layout.HeaderIndex
    if ($header -lt 0) { return ,(New-AbmDataTable -Columns @('Columna 1')) }

    $offset = $Layout.ColOffset
    $width  = 1
    for ($i = $header; $i -lt $Grid.Values.Count; $i++) {
        $length = $Grid.Values[$i].Length - $offset
        if ($length -gt $width) { $width = $length }
    }

    $names = [string[]]::new($width)
    Copy-Cells -Source $Grid.Values[$header] -Offset $offset -Target $names
    for ($i = 0; $i -lt $width; $i++) { if ($null -eq $names[$i]) { $names[$i] = '' } }
    $table = New-AbmDataTable -Columns $names

    $table.BeginLoadData()
    for ($i = $header + 1; $i -lt $Grid.Values.Count; $i++) {
        if ($Grid.Filled[$i] -eq 0) { continue }
        $source = $Grid.Values[$i]
        $cells  = [object[]]::new($width)
        $count  = [math]::Min($source.Length - $offset, $width)
        if ($count -gt 0) { [System.Array]::Copy($source, $offset, $cells, 0, $count) }
        [void]$table.Rows.Add($cells)
    }
    $table.EndLoadData()
    $table.AcceptChanges()
    # Posición original de cada columna: permite conservar su estilo aunque se agreguen, muevan o quiten columnas.
    for ($i = 0; $i -lt $table.Columns.Count; $i++) { $table.Columns[$i].ExtendedProperties['AbmSourceIndex'] = $i }
    return ,$table
}

#region Internos
function Open-XmlStream {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$Part)

    $entry = $Zip.GetEntry($Part)
    if ($null -eq $entry) { throw "No se encontró la parte '$Part' dentro del archivo." }
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.IgnoreComments               = $true
    $settings.IgnoreProcessingInstructions = $true
    $settings.DtdProcessing                = [System.Xml.DtdProcessing]::Prohibit
    $settings.CloseInput                   = $true
    return [System.Xml.XmlReader]::Create($entry.Open(), $settings)
}

function Copy-Cells {
    param([string[]]$Source, [int]$Offset, [string[]]$Target)
    $count = [math]::Min($Source.Length - $Offset, $Target.Length)
    if ($count -gt 0) { [System.Array]::Copy($Source, $Offset, $Target, 0, $count) }
}
#endregion

Export-ModuleMember -Function Get-AbmSharedStrings, Read-AbmSheetGrid, Get-AbmSheetLayout, ConvertTo-AbmSheetTable
