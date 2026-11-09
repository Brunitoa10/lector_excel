Set-StrictMode -Version Latest

<#
    Estrategia para archivos .txt: orquesta codificación (ABM.TextEncoding) y reglas de
    delimitado (ABM.DelimitedText). Sin separador consistente -> modo texto plano (una línea = un registro).
#>

$script:SampleSize  = 10
$script:PlainColumn = 'Línea'
$script:NewLine     = "`r`n"

function Read-TextSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $bytes    = [System.IO.File]::ReadAllBytes($Path)
    $encoding = Get-AbmTextEncoding -Bytes $bytes
    $text     = $encoding.Encoding.GetString($bytes, $encoding.Offset, $bytes.Length - $encoding.Offset)
    $lines    = [regex]::Split($text, '\r\n|\n|\r')
    $count    = $lines.Length
    if ($count -gt 0 -and $lines[$count - 1] -eq '') { $count-- }

    $nonEmpty = @(for ($i = 0; $i -lt $count; $i++) { if ($lines[$i].Trim() -ne '') { $i } })
    $sample   = @($nonEmpty | Select-Object -First $script:SampleSize | ForEach-Object { $lines[$_] })
    $delimiter = if ($sample.Count -gt 0) { Find-AbmDelimiter -Sample $sample } else { '' }

    $table = if ($delimiter -eq '') { ConvertFrom-PlainLines -Lines $lines -Count $count }
             else { ConvertFrom-DelimitedLines -Lines $lines -Indexes $nonEmpty -Delimiter $delimiter }
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
    param([string[]]$Lines, [int]$Count)

    $table = New-AbmDataTable -Columns @($script:PlainColumn)
    for ($i = 0; $i -lt $Count; $i++) { [void]$table.Rows.Add([object[]]@($Lines[$i])) }
    return ,$table
}

function ConvertFrom-DelimitedLines {
    param([string[]]$Lines, [int[]]$Indexes, [string]$Delimiter)

    $parsed = New-Object 'System.Collections.Generic.List[string[]]'
    $width  = 0
    foreach ($index in $Indexes) {
        $fields = Split-AbmDelimitedLine -Line $Lines[$index] -Delimiter $Delimiter
        if ($fields.Length -gt $width) { $width = $fields.Length }
        $parsed.Add($fields)
    }

    $table = New-AbmDataTable -Columns (Get-PaddedFields -Fields $parsed[0] -Width $width)
    for ($r = 1; $r -lt $parsed.Count; $r++) {
        [void]$table.Rows.Add([object[]](Get-PaddedFields -Fields $parsed[$r] -Width $width))
    }
    return ,$table
}

function Get-PaddedFields {
    param([string[]]$Fields, [int]$Width)
    $result = [string[]]::new($Width)
    for ($i = 0; $i -lt $Width; $i++) { $result[$i] = if ($i -lt $Fields.Length) { $Fields[$i] } else { '' } }
    return ,$result
}

function ConvertTo-PlainText {
    param([System.Data.DataTable]$Table)
    $builder = New-Object System.Text.StringBuilder
    foreach ($row in (Get-ActiveRows -Table $Table)) { [void]$builder.Append([string]$row[0]).Append($script:NewLine) }
    return $builder.ToString()
}

function ConvertTo-DelimitedText {
    param([System.Data.DataTable]$Table, [string]$Delimiter)

    $builder = New-Object System.Text.StringBuilder
    $headers = @(foreach ($column in $Table.Columns) { Format-AbmDelimitedField -Value $column.ColumnName -Delimiter $Delimiter })
    [void]$builder.Append($headers -join $Delimiter).Append($script:NewLine)

    foreach ($row in (Get-ActiveRows -Table $Table)) {
        $fields = @(for ($i = 0; $i -lt $Table.Columns.Count; $i++) { Format-AbmDelimitedField -Value ([string]$row[$i]) -Delimiter $Delimiter })
        [void]$builder.Append($fields -join $Delimiter).Append($script:NewLine)
    }
    return $builder.ToString()
}

function Get-ActiveRows {
    param([System.Data.DataTable]$Table)
    foreach ($row in $Table.Rows) { if ($row.RowState -ne [System.Data.DataRowState]::Deleted) { $row } }
}
#endregion

Export-ModuleMember -Function Read-TextSource, Save-TextSource
