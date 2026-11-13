Set-StrictMode -Version Latest

<#
    Preferencias del usuario que la aplicación recuerda entre sesiones (p. ej. el tema).
    Se guardan en preferencias.json, separado de config.json (que es configuración editable a mano).
#>

$script:Path   = $null
$script:Values = @{}

function Initialize-AbmPreferences {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)

    $script:Path   = Join-Path $Root 'preferencias.json'
    $script:Values = @{}
    if (-not (Test-Path -LiteralPath $script:Path)) { return }
    try {
        $json = Get-Content -LiteralPath $script:Path -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $json.PSObject.Properties) { $script:Values[$property.Name] = $property.Value }
    }
    catch { Write-Verbose "preferencias.json inválido; se ignora. $($_.Exception.Message)" }
}

function Get-AbmPreference {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name, $Default = $null)
    if ($script:Values.ContainsKey($Name)) { return $script:Values[$Name] }
    return $Default
}

function Set-AbmPreference {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name, $Value)

    $script:Values[$Name] = $Value
    if (-not $script:Path) { return }
    try {
        $json = [pscustomobject]$script:Values | ConvertTo-Json
        [System.IO.File]::WriteAllText($script:Path, $json, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch { Write-Verbose "No se pudo guardar preferencias.json: $($_.Exception.Message)" }
}

Export-ModuleMember -Function Initialize-AbmPreferences, Get-AbmPreference, Set-AbmPreference
