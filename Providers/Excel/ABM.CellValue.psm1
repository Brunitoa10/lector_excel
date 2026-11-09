Set-StrictMode -Version Latest

<#
    Conversión de valores de celda en ambos sentidos (texto que ve el usuario <-> valor almacenado):
    números en formato invariante, fechas OLE Automation y booleanos.
#>

$script:Invariant     = [System.Globalization.CultureInfo]::InvariantCulture
$script:NumberPattern = New-Object System.Text.RegularExpressions.Regex('^-?(0|[1-9]\d{0,14})(\.\d+)?([eE][-+]?\d+)?$')
$script:TrueText      = 'VERDADERO'
$script:FalseText     = 'FALSO'
$script:DateFormats   = [string[]]@(
    'dd/MM/yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm', 'dd/MM/yyyy',
    'd/M/yyyy H:mm:ss', 'd/M/yyyy H:mm', 'd/M/yyyy',
    'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd'
)

#region Lectura (almacenado -> texto)
function Format-AbmOADate {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][double]$Value)

    try { $date = [datetime]::FromOADate($Value) }
    catch { return $Value.ToString($script:Invariant) }

    if ($Value -lt 1) { return $date.ToString('HH:mm:ss') }
    if ([math]::Abs($Value - [math]::Floor($Value)) -lt 1e-9) { return $date.ToString('dd/MM/yyyy') }
    return $date.ToString('dd/MM/yyyy HH:mm:ss')
}

function Format-AbmCellNumber {
    <# Igual que Excel: 15 dígitos significativos (evita ruido como 656.6900000000001). #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Raw)

    if ($Raw.Length -le 15) { return $Raw }
    $number = 0.0
    if (ConvertTo-Double -Text $Raw -Result ([ref]$number)) { return $number.ToString('G15', $script:Invariant) }
    return $Raw
}

function ConvertFrom-AbmStoredDate {
    <# Valor numérico de una celda con formato de fecha -> texto de fecha; $null si no es número. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Raw)
    $number = 0.0
    if (ConvertTo-Double -Text $Raw -Result ([ref]$number)) { return (Format-AbmOADate -Value $number) }
    return $null
}

function ConvertTo-AbmBooleanText {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Raw)
    if ($Raw -eq '1') { return $script:TrueText }
    return $script:FalseText
}
#endregion

#region Escritura (texto -> almacenado)
function ConvertTo-AbmCellNumber {
    <# Devuelve el valor numérico en formato invariante, o $null si debe guardarse como texto. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value, [switch]$IsDateColumn)

    $number = 0.0
    if ($script:NumberPattern.IsMatch($Value) -and (ConvertTo-Double -Text $Value -Result ([ref]$number))) {
        return $number.ToString('R', $script:Invariant)
    }
    if (-not $IsDateColumn) { return $null }

    $date = [datetime]::MinValue
    if ([datetime]::TryParseExact($Value, $script:DateFormats, $script:Invariant, [System.Globalization.DateTimeStyles]::None, [ref]$date)) {
        return $date.ToOADate().ToString('R', $script:Invariant)
    }
    $time = [timespan]::Zero
    if ($Value -match '^\d{1,2}:\d{2}(:\d{2})?$' -and [timespan]::TryParse($Value, $script:Invariant, [ref]$time)) {
        return $time.TotalDays.ToString('R', $script:Invariant)
    }
    return $null
}

#endregion

function Get-AbmCellCodec {
    <# Reglas de conversión precompiladas, para bucles de alto volumen que no pueden llamar funciones por celda. #>
    [CmdletBinding()]
    param()
    return [pscustomobject]@{
        NumberPattern = $script:NumberPattern
        DateFormats   = $script:DateFormats
        Culture       = $script:Invariant
        TrueText      = $script:TrueText
        FalseText     = $script:FalseText
    }
}

function ConvertTo-Double {
    param([string]$Text, [ref]$Result)
    $value = 0.0
    $ok = [double]::TryParse($Text, [System.Globalization.NumberStyles]::Float, $script:Invariant, [ref]$value)
    $Result.Value = $value
    return $ok
}

Export-ModuleMember -Function Format-AbmOADate, Format-AbmCellNumber, ConvertFrom-AbmStoredDate, ConvertTo-AbmBooleanText,
                              ConvertTo-AbmCellNumber, Get-AbmCellCodec
