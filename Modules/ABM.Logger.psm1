Set-StrictMode -Version Latest

<#
    Responsabilidad única: trazabilidad.
      - Auditoría de operaciones (ABRIR, ALTA, BAJA, MODIFICACION, GUARDAR...) en Logs\auditoria_yyyyMM.txt
      - Respaldo del archivo antes de sobrescribirlo, con rotación (MaxBackupsPorArchivo).
#>

$script:Settings = $null
$script:Utf8Bom  = New-Object System.Text.UTF8Encoding($true)

function Initialize-AbmLogger {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)]$Config
    )

    $script:Settings = [pscustomobject]@{
        LogDir     = Join-Path $Root ([string]$Config.CarpetaLogs)
        BackupDir  = Join-Path $Root ([string]$Config.CarpetaBackups)
        Audit      = [bool]$Config.Auditoria
        Backup     = [bool]$Config.Backup
        MaxBackups = [math]::Max(1, [int]$Config.MaxBackupsPorArchivo)
    }
}

function Confirm-AbmDirectory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { [void](New-Item -ItemType Directory -Path $Path -Force) }
}

function Write-AbmAudit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Action,
        [string]$Target = '',
        [string]$Detail = ''
    )

    if ($null -eq $script:Settings -or -not $script:Settings.Audit) { return }

    try {
        Confirm-AbmDirectory $script:Settings.LogDir
        $now  = Get-Date
        $file = Join-Path $script:Settings.LogDir ('auditoria_{0:yyyyMM}.txt' -f $now)
        $line = '{0:yyyy-MM-dd HH:mm:ss} | {1,-12} | {2,-15} | {3} | {4}' -f $now, $Action, $env:USERNAME, $Target, ($Detail -replace '\r?\n', ' ')
        [System.IO.File]::AppendAllText($file, $line + [Environment]::NewLine, $script:Utf8Bom)
    }
    catch {
        Write-Verbose "No se pudo escribir la auditoría: $($_.Exception.Message)"
    }
}

function Backup-AbmFile {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    if ($null -eq $script:Settings -or -not $script:Settings.Backup) { return '' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }

    Confirm-AbmDirectory $script:Settings.BackupDir
    $name   = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext    = [System.IO.Path]::GetExtension($Path)
    $target = Join-Path $script:Settings.BackupDir ('{0}_{1:yyyyMMdd_HHmmss}{2}' -f $name, (Get-Date), $ext)
    Copy-Item -LiteralPath $Path -Destination $target -Force

    Remove-AbmOldBackups -BaseName $name -Extension $ext
    return $target
}

function Remove-AbmOldBackups {
    param([string]$BaseName, [string]$Extension)

    $pattern = '{0}_*{1}' -f [System.Management.Automation.WildcardPattern]::Escape($BaseName), $Extension
    $old = @(
        Get-ChildItem -LiteralPath $script:Settings.BackupDir -File |
            Where-Object { $_.Name -like $pattern } |
            Sort-Object Name -Descending |
            Select-Object -Skip $script:Settings.MaxBackups
    )
    foreach ($item in $old) { Remove-Item -LiteralPath $item.FullName -Force -ErrorAction SilentlyContinue }
}

Export-ModuleMember -Function Initialize-AbmLogger, Write-AbmAudit, Backup-AbmFile
