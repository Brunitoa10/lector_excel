Set-StrictMode -Version Latest

<#
    Reglas del texto delimitado: detección de separador, parseo de líneas con comillas y escape de campos.
    Funciones puras (sin E/S), fáciles de probar.
#>

$script:Candidates     = @("`t", ';', '|', ',')
$script:DelimiterNames = @{ "`t" = 'Tabulación'; ';' = 'Punto y coma'; '|' = 'Barra vertical'; ',' = 'Coma'; '' = 'Texto plano' }

function Find-AbmDelimiter {
    <# Devuelve el separador con más apariciones en la muestra; '' si no hay ninguno confiable. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string[]]$Sample)

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

function Split-AbmDelimitedLine {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Line, [Parameter(Mandatory)][string]$Delimiter)

    if ($Line.IndexOf('"') -lt 0) { return ,$Line.Split([char[]]$Delimiter) }

    $fields   = New-Object System.Collections.Generic.List[string]
    $current  = New-Object System.Text.StringBuilder
    $inQuotes = $false
    $quote    = [char]'"'
    $sep      = $Delimiter[0]

    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($inQuotes) {
            if ($ch -ne $quote) { [void]$current.Append($ch); continue }
            if ($i + 1 -lt $Line.Length -and $Line[$i + 1] -eq $quote) { [void]$current.Append($quote); $i++ }
            else { $inQuotes = $false }
        }
        elseif ($ch -eq $quote -and $current.Length -eq 0) { $inQuotes = $true }
        elseif ($ch -eq $sep) { $fields.Add($current.ToString()); [void]$current.Clear() }
        else { [void]$current.Append($ch) }
    }
    $fields.Add($current.ToString())
    return ,$fields.ToArray()
}

function Format-AbmDelimitedField {
    <# Pone entre comillas el valor si contiene el separador, comillas o saltos de línea. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Value, [Parameter(Mandatory)][string]$Delimiter)

    if ($Value.Contains($Delimiter) -or $Value.Contains('"') -or $Value.IndexOfAny([char[]]"`r`n") -ge 0) {
        return '"' + $Value.Replace('"', '""') + '"'
    }
    return $Value
}

function Get-AbmDelimiterName {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Delimiter)

    if ($script:DelimiterNames.ContainsKey($Delimiter)) { return $script:DelimiterNames[$Delimiter] }
    return "'$Delimiter'"
}

Export-ModuleMember -Function Find-AbmDelimiter, Split-AbmDelimitedLine, Format-AbmDelimitedField, Get-AbmDelimiterName
