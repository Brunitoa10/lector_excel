Set-StrictMode -Version Latest

<#
    Estrategia para archivos .txt: orquesta codificación (ABM.TextEncoding) y reglas de
    delimitado (ABM.DelimitedText). Sin separador consistente -> modo texto plano (una línea = un registro).
#>

$script:SampleSize  = 10
$script:PlainColumn = 'Línea'
$script:NewLine     = "`r`n"
$script:ProgressEvery = 5000

function Read-TextSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [scriptblock]$OnProgress)

    $bytes    = [System.IO.File]::ReadAllBytes($Path)
    $encoding = Get-AbmTextEncoding -Bytes $bytes
    $text     = $encoding.Encoding.GetString($bytes, $encoding.Offset, $bytes.Length - $encoding.Offset)
    $lines    = [regex]::Split($text, '\r\n|\n|\r')
    $count    = $lines.Length
    if ($count -gt 0 -and $lines[$count - 1] -eq '') { $count-- }

    $nonEmpty = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $count; $i++) { if (-not [string]::IsNullOrWhiteSpace($lines[$i])) { $nonEmpty.Add($i) } }
    $sampleSize = [math]::Min($nonEmpty.Count, $script:SampleSize)
    $sample     = [string[]]::new($sampleSize)
    for ($i = 0; $i -lt $sampleSize; $i++) { $sample[$i] = $lines[$nonEmpty[$i]] }
    $delimiter = if ($sampleSize -gt 0) { Find-AbmDelimiter -Sample $sample } else { '' }

    if ($delimiter -eq '') { $table = ConvertFrom-PlainLines -Lines $lines -Count $count -OnProgress $OnProgress }
    else { $table = ConvertFrom-DelimitedLines -Lines $lines -Indexes $nonEmpty.ToArray() -Delimiter $delimiter -OnProgress $OnProgress }
    $table.AcceptChanges()

    return @{
        Table = $table
        Meta  = @{ Delimiter = $delimiter; Encoding = $encoding.Encoding; EncodingName = $encoding.Name }
    }
}

function Save-TextSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [AllowEmptyString()][string]$Delimiter = '',
        [System.Text.Encoding]$Encoding
    )

    if ($null -eq $Encoding) { $Encoding = New-Object System.Text.UTF8Encoding($true) }
    $content = if ([string]::IsNullOrEmpty($Delimiter)) { ConvertTo-PlainText -Table $Table }
               else { ConvertTo-DelimitedText -Table $Table -Delimiter $Delimiter }
    [System.IO.File]::WriteAllText($Path, $content, $Encoding)
}

#region Conversión texto <-> tabla
function ConvertFrom-PlainLines {
    param([string[]]$Lines, [int]$Count, [scriptblock]$OnProgress)

    $table = New-AbmDataTable -Columns @($script:PlainColumn)
    $table.BeginLoadData()
    for ($i = 0; $i -lt $Count; $i++) {
        [void]$table.Rows.Add([object[]]@($Lines[$i]))
        if ($null -ne $OnProgress -and ($i % $script:ProgressEvery) -eq 0 -and $i -gt 0) { & $OnProgress $i }
    }
    $table.EndLoadData()
    return ,$table
}

function ConvertFrom-DelimitedLines {
    <# Bucle sin llamadas a funciones: Split nativo si la línea no tiene comillas; copia por bloque a la fila. #>
    param([string[]]$Lines, [int[]]$Indexes, [string]$Delimiter, [scriptblock]$OnProgress)

    $separator = [char[]]$Delimiter
    $parsed = New-Object 'System.Collections.Generic.List[string[]]' ($Indexes.Length)
    $width  = 0
    foreach ($index in $Indexes) {
        $line = $Lines[$index]
        if ($line.IndexOf('"') -lt 0) { $fields = $line.Split($separator) }
        else { $fields = Split-AbmDelimitedLine -Line $line -Delimiter $Delimiter }
        if ($fields.Length -gt $width) { $width = $fields.Length }
        $parsed.Add($fields)
    }

    $names = [string[]]::new($width)
    [System.Array]::Copy($parsed[0], $names, $parsed[0].Length)
    for ($i = 0; $i -lt $width; $i++) { if ($null -eq $names[$i]) { $names[$i] = '' } }
    $table = New-AbmDataTable -Columns $names

    $table.BeginLoadData()
    for ($r = 1; $r -lt $parsed.Count; $r++) {
        $source = $parsed[$r]
        $cells  = [object[]]::new($width)
        [System.Array]::Copy($source, $cells, $source.Length)
        for ($c = $source.Length; $c -lt $width; $c++) { $cells[$c] = '' }
        [void]$table.Rows.Add($cells)
        if ($null -ne $OnProgress -and ($r % $script:ProgressEvery) -eq 0) { & $OnProgress $r }
    }
    $table.EndLoadData()
    return ,$table
}

function ConvertTo-PlainText {
    param([System.Data.DataTable]$Table)
    $builder = New-Object System.Text.StringBuilder (1MB)
    foreach ($row in $Table.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
        [void]$builder.Append([string]$row[0]).Append($script:NewLine)
    }
    return $builder.ToString()
}

function ConvertTo-DelimitedText {
    <# Camino rápido: une la fila completa y sólo escapa campo por campo si hace falta (comillas, saltos o delimitador). #>
    param([System.Data.DataTable]$Table, [string]$Delimiter)

    $builder  = New-Object System.Text.StringBuilder (1MB)
    $expected = $Table.Columns.Count - 1
    $special  = [char[]]("`"`r`n")
    $headers  = @(foreach ($column in $Table.Columns) { Format-AbmDelimitedField -Value $column.ColumnName -Delimiter $Delimiter })
    [void]$builder.Append($headers -join $Delimiter).Append($script:NewLine)

    foreach ($row in $Table.Rows) {
        if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
        $line = [string]::Join($Delimiter, $row.ItemArray)
        $needsQuoting = $line.IndexOfAny($special) -ge 0 -or ($line.Split([char[]]$Delimiter).Length - 1) -ne $expected
        if ($needsQuoting) {
            $fields = @(foreach ($item in $row.ItemArray) { Format-AbmDelimitedField -Value ([string]$item) -Delimiter $Delimiter })
            $line = $fields -join $Delimiter
        }
        [void]$builder.Append($line).Append($script:NewLine)
    }
    return $builder.ToString()
}

#endregion

Export-ModuleMember -Function Read-TextSource, Save-TextSource
