Set-StrictMode -Version Latest

<#
    Estrategia para archivos .txt.
      - Detecta codificación (BOM UTF-8/UTF-16, UTF-8 sin BOM o ANSI) y la conserva al guardar.
      - Detecta delimitador (Tab, ; | ,). Si no hay delimitador consistente -> modo "texto plano" (una línea = un registro).
      - Soporta campos entre comillas ("valor; con delimitador").
#>

$script:Candidates     = @("`t", ';', '|', ',')
$script:DelimiterNames = @{ "`t" = 'Tabulación'; ';' = 'Punto y coma'; '|' = 'Barra vertical'; ',' = 'Coma'; '' = 'Texto plano' }
$script:SampleSize     = 10
$script:PlainColumn    = 'Línea'

#region API pública
function Read-TextSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $bytes    = [System.IO.File]::ReadAllBytes($Path)
    $encoding = Get-TextEncodingInfo -Bytes $bytes
    $text     = $encoding.Encoding.GetString($bytes, $encoding.Offset, $bytes.Length - $encoding.Offset)
    $lines    = [regex]::Split($text, '\r\n|\n|\r')
    $count    = $lines.Length
    if ($count -gt 0 -and $lines[$count - 1] -eq '') { $count-- }

    $nonEmpty = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $count; $i++) { if ($lines[$i].Trim() -ne '') { $nonEmpty.Add($i) } }

    $sample    = @($nonEmpty | Select-Object -First $script:SampleSize | ForEach-Object { $lines[$_] })
    $delimiter = if ($sample.Count -gt 0) { Find-Delimiter -Sample $sample } else { '' }
    $meta      = @{ Delimiter = $delimiter; Encoding = $encoding.Encoding; EncodingName = $encoding.Name }

    $table = if ($delimiter -eq '') { ConvertFrom-PlainLines -Lines $lines -Count $count }
             else { ConvertFrom-DelimitedLines -Lines $lines -Indexes $nonEmpty.ToArray() -Delimiter $delimiter }
    $table.AcceptChanges()

    return @{ Table = $table; Meta = $meta }
}

function Save-TextSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [string]$Delimiter = '',
        [System.Text.Encoding]$Encoding
    )

    if ($null -eq $Encoding) { $Encoding = New-Object System.Text.UTF8Encoding($true) }
    $newLine = "`r`n"
    $builder = New-Object System.Text.StringBuilder

    if ([string]::IsNullOrEmpty($Delimiter)) {
        foreach ($row in $Table.Rows) {
            if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
            [void]$builder.Append([string]$row[0]).Append($newLine)
        }
    }
    else {
        $headers = @(foreach ($column in $Table.Columns) { Format-DelimitedField -Value $column.ColumnName -Delimiter $Delimiter })
        [void]$builder.Append($headers -join $Delimiter).Append($newLine)
        foreach ($row in $Table.Rows) {
            if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
            $fields = @(for ($i = 0; $i -lt $Table.Columns.Count; $i++) { Format-DelimitedField -Value ([string]$row[$i]) -Delimiter $Delimiter })
            [void]$builder.Append($fields -join $Delimiter).Append($newLine)
        }
    }

    [System.IO.File]::WriteAllText($Path, $builder.ToString(), $Encoding)
}

function Get-AbmDelimiterName {
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Delimiter)

    if ($script:DelimiterNames.ContainsKey($Delimiter)) { return $script:DelimiterNames[$Delimiter] }
    return "'$Delimiter'"
}
#endregion

#region Codificación
function Get-TextEncodingInfo {
    param([byte[]]$Bytes)

    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return @{ Encoding = (New-Object System.Text.UTF8Encoding($true)); Offset = 3; Name = 'UTF-8 (BOM)' }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return @{ Encoding = [System.Text.Encoding]::Unicode; Offset = 2; Name = 'UTF-16 LE' }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) {
        return @{ Encoding = [System.Text.Encoding]::BigEndianUnicode; Offset = 2; Name = 'UTF-16 BE' }
    }
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$strict.GetString($Bytes)
        return @{ Encoding = (New-Object System.Text.UTF8Encoding($false)); Offset = 0; Name = 'UTF-8' }
    }
    catch {
        $ansi = [System.Text.Encoding]::Default
        return @{ Encoding = $ansi; Offset = 0; Name = "ANSI ($($ansi.WebName))" }
    }
}
#endregion

#region Parseo
function Find-Delimiter {
    param([string[]]$Sample)

    $best = ''
    $bestCount = 0
    foreach ($candidate in $script:Candidates) {
        $separator = [char[]]$candidate
        $counts = @($Sample | ForEach-Object { $_.Split($separator).Length - 1 })
        $first  = $counts[0]
        if ($first -le 0) { continue }

        $consistent = @($counts | Where-Object { $_ -ne $first }).Count -eq 0
        if (-not $consistent -and $candidate -eq ',') { continue }   # la coma aparece en prosa: exigir consistencia
        if ($first -gt $bestCount) { $best = $candidate; $bestCount = $first }
    }
    return $best
}

function Split-DelimitedLine {
    param([string]$Line, [string]$Delimiter)

    if ($Line.IndexOf('"') -lt 0) { return ,$Line.Split([char[]]$Delimiter) }

    $fields   = New-Object System.Collections.Generic.List[string]
    $current  = New-Object System.Text.StringBuilder
    $inQuotes = $false
    $sep      = $Delimiter[0]

    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($inQuotes) {
            if ($ch -eq [char]'"') {
                if ($i + 1 -lt $Line.Length -and $Line[$i + 1] -eq [char]'"') { [void]$current.Append('"'); $i++ }
                else { $inQuotes = $false }
            }
            else { [void]$current.Append($ch) }
        }
        elseif ($ch -eq [char]'"' -and $current.Length -eq 0) { $inQuotes = $true }
        elseif ($ch -eq $sep) { $fields.Add($current.ToString()); [void]$current.Clear() }
        else { [void]$current.Append($ch) }
    }
    $fields.Add($current.ToString())
    return ,$fields.ToArray()
}

function Format-DelimitedField {
    param([string]$Value, [string]$Delimiter)

    if ($Value.Contains($Delimiter) -or $Value.Contains('"') -or $Value.IndexOfAny([char[]]"`r`n") -ge 0) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

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
        $fields = Split-DelimitedLine -Line $Lines[$index] -Delimiter $Delimiter
        if ($fields.Length -gt $width) { $width = $fields.Length }
        $parsed.Add($fields)
    }

    $headers = [string[]]::new($width)
    for ($c = 0; $c -lt $width; $c++) { $headers[$c] = if ($c -lt $parsed[0].Length) { $parsed[0][$c] } else { '' } }
    $table = New-AbmDataTable -Columns $headers

    for ($r = 1; $r -lt $parsed.Count; $r++) {
        $source = $parsed[$r]
        $values = [object[]]::new($width)
        for ($c = 0; $c -lt $width; $c++) { $values[$c] = if ($c -lt $source.Length) { $source[$c] } else { '' } }
        [void]$table.Rows.Add($values)
    }
    return ,$table
}
#endregion

Export-ModuleMember -Function Read-TextSource, Save-TextSource, Get-AbmDelimiterName
