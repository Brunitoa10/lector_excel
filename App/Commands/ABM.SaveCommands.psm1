Set-StrictMode -Version Latest

<#
    Comandos de persistencia: guardar, guardar como (con conversión Excel <-> TXT) y crear archivo nuevo.
    Todos respaldan antes de sobrescribir y dejan registro en la auditoría.
#>

$script:SaveAsFilter = 'Libro de Excel (*.xlsx)|*.xlsx|Archivo de texto (*.txt)|*.txt'

function Invoke-AbmSave {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    Complete-AbmPendingEdit
    $state  = Get-AbmState
    $source = $state.Source
    if ($null -eq $source) { return $false }
    if (-not $state.Dirty) { Set-AbmStatus 'No hay cambios pendientes de guardar.'; return $true }

    try {
        Set-AbmBusy -Busy $true -Text 'Guardando cambios…'
        $backup = Backup-AbmFile -Path $source.Path
        Save-AbmSource -Source $source -Table $state.Table
        $state.Table.AcceptChanges()
        Set-AbmDirty -Value $false
        Write-AbmAudit -Action 'GUARDAR' -Target (Get-AbmSourceLabel) -Detail (Get-SaveDetail -Count $state.Table.Rows.Count -Backup $backup)
        Set-AbmStatus ('Cambios guardados en {0} a las {1:HH:mm:ss}' -f (Get-AbmSourceLabel), (Get-Date))
        return $true
    }
    catch {
        Show-AbmError -Owner (Get-AbmUi).Form -ErrorRecord $_ `
            -Message "No se pudo guardar el archivo.`nVerificá que no esté abierto en Excel u otro programa."
        return $false
    }
    finally { Set-AbmBusy -Busy $false }
}

function Invoke-AbmSaveAs {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    Complete-AbmPendingEdit
    $state  = Get-AbmState
    $source = $state.Source
    if ($null -eq $source) { return $false }

    $target = Select-AbmSavePath -Title 'Guardar como' -Filter $script:SaveAsFilter `
        -FilterIndex $(if ($source.Kind -eq 'Excel') { 1 } else { 2 }) `
        -FileName ([System.IO.Path]::GetFileNameWithoutExtension($source.Path) + '_copia') `
        -Directory ([System.IO.Path]::GetDirectoryName($source.Path))
    if (-not $target) { return $false }

    if ([string]::Equals([System.IO.Path]::GetFullPath($target), $source.Path, [System.StringComparison]::OrdinalIgnoreCase)) {
        $state.Dirty = $true
        return (Invoke-AbmSave)
    }

    $created = New-AbmFile -Path $target -Table $state.Table -Options (Get-SaveAsOptions -Source $source) `
                           -Action 'GUARDAR_COMO' -Detail ('Origen: {0}' -f (Get-AbmSourceLabel))
    if ($created) { $state.Dirty = $false; Open-AbmDocument -Path $target }
    return $created
}

function Invoke-AbmNewFile {
    [CmdletBinding()]
    param()

    if (-not (Confirm-AbmPendingChanges)) { return }
    $config = Get-AbmSetting -Name Config
    $spec = Show-AbmNewFileDialog -Owner (Get-AbmUi).Form -DefaultDelimiter ([string]$config.DelimitadorTxtPorDefecto)
    if ($null -eq $spec) { return }

    $filter = if ($spec.Format -eq 'Excel') { 'Libro de Excel (*.xlsx)|*.xlsx' } else { 'Archivo de texto (*.txt)|*.txt' }
    $target = Select-AbmSavePath -Title 'Crear archivo' -Filter $filter -FileName 'NuevoArchivo' `
                                 -Directory ([Environment]::GetFolderPath('MyDocuments'))
    if (-not $target) { return }

    $table   = New-AbmDataTable -Columns $spec.Columns
    $options = @{ SheetName = $spec.SheetName; Delimiter = $spec.Delimiter }
    if (New-AbmFile -Path $target -Table $table -Options $options -Action 'CREAR' -Detail ('Columnas: {0}' -f ($spec.Columns -join ', '))) {
        Open-AbmDocument -Path $target
    }
}

#region Internos
function New-AbmFile {
    <# Crea el archivo (respaldando el existente) y audita. Devuelve $true si salió bien. #>
    param([string]$Path, [System.Data.DataTable]$Table, [hashtable]$Options, [string]$Action, [string]$Detail)
    try {
        Set-AbmBusy -Busy $true -Text 'Generando archivo…'
        $null = Backup-AbmFile -Path $Path
        New-AbmSource -Path $Path -Table $Table -Options $Options
        Write-AbmAudit -Action $Action -Target ([System.IO.Path]::GetFileName($Path)) -Detail $Detail
        return $true
    }
    catch {
        Show-AbmError -Owner (Get-AbmUi).Form -Message 'No se pudo generar el archivo.' -ErrorRecord $_
        return $false
    }
    finally { Set-AbmBusy -Busy $false }
}

function Get-SaveAsOptions {
    param($Source)
    $config = Get-AbmSetting -Name Config
    return @{
        SheetName = if ($Source.Kind -eq 'Excel') { $Source.Meta.Sheet } else { 'Datos' }
        Delimiter = if ($Source.Kind -eq 'Txt') { $Source.Meta.Delimiter } else { [string]$config.DelimitadorTxtPorDefecto }
    }
}

function Get-SaveDetail {
    param([int]$Count, [string]$Backup)
    $detail = '{0} registros' -f $Count
    if ($Backup) { $detail += ' · respaldo: ' + [System.IO.Path]::GetFileName($Backup) }
    return $detail
}
#endregion

Export-ModuleMember -Function Invoke-AbmSave, Invoke-AbmSaveAs, Invoke-AbmNewFile
