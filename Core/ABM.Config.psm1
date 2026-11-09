Set-StrictMode -Version Latest

<#
    Responsabilidad única: proveer la configuración de la aplicación.
    Toma valores por defecto y los sobrescribe con los de config.json (si existe y es válido).
#>

function Get-AbmConfig {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $config = [ordered]@{
        Auditoria                = $true
        Backup                   = $true
        MaxBackupsPorArchivo     = 10
        CarpetaLogs              = 'Logs'
        CarpetaBackups           = 'Backups'
        DelimitadorTxtPorDefecto = ';'
    }

    if (Test-Path -LiteralPath $Path) {
        try {
            $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($property in $json.PSObject.Properties) {
                if ($config.Contains($property.Name)) { $config[$property.Name] = $property.Value }
            }
        }
        catch {
            Write-Warning "config.json inválido; se usan valores por defecto. Detalle: $($_.Exception.Message)"
        }
    }

    $config.DelimitadorTxtPorDefecto = ConvertTo-AbmDelimiter ([string]$config.DelimitadorTxtPorDefecto)
    return [pscustomobject]$config
}

function ConvertTo-AbmDelimiter {
    param([string]$Value)

    switch ($Value.Trim().ToLowerInvariant()) {
        { $_ -in 'tab', '\t', 'tabulacion', 'tabulación' } { return "`t" }
        { $_ -in ';', '|', ',' }                           { return $_ }
        default {
            if ($Value -eq "`t") { return "`t" }
            return ';'
        }
    }
}

Export-ModuleMember -Function Get-AbmConfig
