Set-StrictMode -Version Latest

<#
    Detección de codificación de archivos de texto: BOM UTF-8 / UTF-16, UTF-8 sin BOM o ANSI del sistema.
    Devuelve la codificación a usar para leer y para volver a guardar el archivo igual que estaba.
#>

function Get-AbmTextEncoding {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes)

    if (Test-Prefix -Bytes $Bytes -Prefix 0xEF, 0xBB, 0xBF) {
        return @{ Encoding = (New-Object System.Text.UTF8Encoding($true)); Offset = 3; Name = 'UTF-8 (BOM)' }
    }
    if (Test-Prefix -Bytes $Bytes -Prefix 0xFF, 0xFE) {
        return @{ Encoding = [System.Text.Encoding]::Unicode; Offset = 2; Name = 'UTF-16 LE' }
    }
    if (Test-Prefix -Bytes $Bytes -Prefix 0xFE, 0xFF) {
        return @{ Encoding = [System.Text.Encoding]::BigEndianUnicode; Offset = 2; Name = 'UTF-16 BE' }
    }
    if (Test-ValidUtf8 -Bytes $Bytes) {
        return @{ Encoding = (New-Object System.Text.UTF8Encoding($false)); Offset = 0; Name = 'UTF-8' }
    }
    $ansi = [System.Text.Encoding]::Default
    return @{ Encoding = $ansi; Offset = 0; Name = "ANSI ($($ansi.WebName))" }
}

function Test-Prefix {
    param([byte[]]$Bytes, [byte[]]$Prefix)
    if ($Bytes.Length -lt $Prefix.Length) { return $false }
    for ($i = 0; $i -lt $Prefix.Length; $i++) { if ($Bytes[$i] -ne $Prefix[$i]) { return $false } }
    return $true
}

function Test-ValidUtf8 {
    param([byte[]]$Bytes)
    try {
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        [void]$strict.GetString($Bytes)
        return $true
    }
    catch { return $false }
}

Export-ModuleMember -Function Get-AbmTextEncoding
