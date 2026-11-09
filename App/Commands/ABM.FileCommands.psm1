Set-StrictMode -Version Latest

<#
    Comandos de archivo: abrir (diálogo, arrastrar y soltar, recargar, cambiar de hoja)
    y la protección de cambios pendientes.
#>

function Open-AbmDocument {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$Sheet)

    $form = (Get-AbmUi).Form
    try {
        Set-AbmBusy -Busy $true -Text ('Abriendo {0}…' -f [System.IO.Path]::GetFileName($Path))
        $source = Open-AbmSource -Path $Path -Sheet $Sheet -OnProgress {
            param($rows) Set-AbmBusy -Busy $true -Text ('Leyendo… {0:N0} filas' -f $rows)
        }
        Set-AbmBusy -Busy $true -Text 'Preparando la vista…'
        Set-AbmSource -Source $source
        $count = $source.Table.Rows.Count
        Write-AbmAudit -Action 'ABRIR' -Target (Get-AbmSourceLabel) -Detail ('{0} registros' -f $count)
        Set-AbmStatus ('Archivo abierto: {0} · {1:N0} registros' -f (Get-AbmSourceLabel), $count)
    }
    catch { Show-AbmError -Owner $form -Message "No se pudo abrir el archivo.`n$Path" -ErrorRecord $_ }
    finally { Set-AbmBusy -Busy $false }
}

function Invoke-AbmOpen {
    [CmdletBinding()]
    param()

    if (-not (Confirm-AbmPendingChanges)) { return }
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dialog.Title  = 'Abrir archivo'
        $dialog.Filter = Get-AbmFileFilter
        if ($dialog.ShowDialog((Get-AbmUi).Form) -eq [System.Windows.Forms.DialogResult]::OK) { Open-AbmDocument -Path $dialog.FileName }
    }
    finally { $dialog.Dispose() }
}

function Invoke-AbmDrop {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Files)

    if ($null -eq $Files -or $Files.Count -eq 0) { return }
    $form = (Get-AbmUi).Form
    try { $null = Get-AbmProvider -Path $Files[0] }
    catch { Show-AbmMessage -Owner $form -Message $_.Exception.Message -Icon Warning; return }

    if (-not (Confirm-AbmPendingChanges)) { return }
    $form.Activate()
    Open-AbmDocument -Path $Files[0]
}

function Invoke-AbmReload {
    [CmdletBinding()]
    param()

    $source = (Get-AbmState).Source
    if ($null -eq $source -or -not (Confirm-AbmPendingChanges)) { return }
    Open-AbmDocument -Path $source.Path -Sheet $source.Meta.Sheet
}

function Invoke-AbmSheetChange {
    [CmdletBinding()]
    param()

    $source = (Get-AbmState).Source
    if ($null -eq $source -or $source.Kind -ne 'Excel') { return }
    $combo = (Get-AbmUi).SheetCombo
    $name  = [string]$combo.SelectedItem
    if ($name -eq $source.Meta.Sheet) { return }
    if (-not (Confirm-AbmPendingChanges)) { $combo.SelectedItem = $source.Meta.Sheet; return }
    Open-AbmDocument -Path $source.Path -Sheet $name
}

function Confirm-AbmPendingChanges {
    <# $true si se puede continuar (no había cambios, se guardaron o se descartaron). #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $state = Get-AbmState
    if (-not $state.Dirty) { return $true }

    $message  = "Hay cambios sin guardar en {0}.`n`n¿Querés guardarlos antes de continuar?" -f (Get-AbmSourceLabel)
    $decision = Request-AbmSaveDecision -Owner (Get-AbmUi).Form -Message $message
    switch ($decision) {
        'Save'    { return [bool](Invoke-AbmSave) }
        'Discard' { $state.Dirty = $false; return $true }
        default   { return $false }
    }
}

function Select-AbmSavePath {
    <# Diálogo "Guardar como"; devuelve la ruta elegida o '' si se canceló. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$Title, [string]$Filter, [int]$FilterIndex = 1, [string]$FileName, [string]$Directory)

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    try {
        $dialog.Title            = $Title
        $dialog.Filter           = $Filter
        $dialog.FilterIndex      = $FilterIndex
        $dialog.FileName         = $FileName
        $dialog.InitialDirectory = $Directory
        $dialog.OverwritePrompt  = $true
        $dialog.AddExtension     = $true
        if ($dialog.ShowDialog((Get-AbmUi).Form) -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
        return ''
    }
    finally { $dialog.Dispose() }
}

Export-ModuleMember -Function Open-AbmDocument, Invoke-AbmOpen, Invoke-AbmDrop, Invoke-AbmReload, Invoke-AbmSheetChange,
                              Confirm-AbmPendingChanges, Select-AbmSavePath
