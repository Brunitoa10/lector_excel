Set-StrictMode -Version Latest

<#
    Direcciones de celda de Excel: letras de columna <-> índice (base 0) y número de fila/columna
    a partir de los atributos "r" de SpreadsheetML (con respaldo secuencial si faltan).
#>

function ConvertFrom-AbmColumnLetters {
    <# "A" -> 0, "Z" -> 25, "AA" -> 26 #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][string]$Letters)

    $number = 0
    foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) { $number = ($number * 26) + ([int]$ch - 64) }
    return $number - 1
}

function ConvertTo-AbmColumnLetters {
    <# 0 -> "A", 25 -> "Z", 26 -> "AA" #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][int]$Index)

    $number  = $Index + 1
    $letters = ''
    while ($number -gt 0) {
        $remainder = ($number - 1) % 26
        $letters   = [string][char](65 + $remainder) + $letters
        $number    = [int][math]::Floor(($number - 1) / 26)
    }
    return $letters
}

function Get-AbmRowNumber {
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][System.Xml.XmlElement]$Row, [int]$Previous)

    $number = 0
    if ([int]::TryParse($Row.GetAttribute('r'), [ref]$number)) { return $number }
    return $Previous + 1
}

function Get-AbmColumnIndex {
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][System.Xml.XmlElement]$Cell, [int]$Previous)

    $letters = $Cell.GetAttribute('r') -replace '[0-9$]', ''
    if ($letters) { return (ConvertFrom-AbmColumnLetters $letters) }
    return $Previous + 1
}

Export-ModuleMember -Function ConvertFrom-AbmColumnLetters, ConvertTo-AbmColumnLetters, Get-AbmRowNumber, Get-AbmColumnIndex
