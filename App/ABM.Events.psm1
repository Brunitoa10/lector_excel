Set-StrictMode -Version Latest

<#
    Conexión de eventos de la UI con los comandos. Todo pasa por Invoke-AbmSafely:
    ningún error en un evento cierra la aplicación.
#>

function Register-AbmEvents {
    [CmdletBinding()]
    param()

    Register-ButtonEvents
    Register-GridEvents
    Register-SearchEvents
    Register-DragDropEvents
    Register-FormEvents
}

function Invoke-AbmSafely {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action, $Argument)
    try {
        if ($PSBoundParameters.ContainsKey('Argument')) { & $Action $Argument } else { & $Action }
    }
    catch {
        Set-AbmBusy -Busy $false
        Show-AbmError -Owner (Get-AbmUi).Form -Message 'Ocurrió un error inesperado.' -ErrorRecord $_
    }
}

#region Grupos de eventos
function Register-ButtonEvents {
    $ui = Get-AbmUi
    $ui.BtnOpen.Add_Click({   Invoke-AbmSafely { Invoke-AbmOpen } })
    $ui.BtnNew.Add_Click({    Invoke-AbmSafely { Invoke-AbmNewFile } })
    $ui.BtnAdd.Add_Click({    Invoke-AbmSafely { Invoke-AbmAdd } })
    $ui.BtnEdit.Add_Click({   Invoke-AbmSafely { Invoke-AbmEdit } })
    $ui.BtnDelete.Add_Click({ Invoke-AbmSafely { Invoke-AbmDelete } })
    $ui.BtnSave.Add_Click({   Invoke-AbmSafely { [void](Invoke-AbmSave) } })
    $ui.BtnSaveAs.Add_Click({ Invoke-AbmSafely { [void](Invoke-AbmSaveAs) } })
}

function Register-GridEvents {
    $ui = Get-AbmUi
    $ui.SheetCombo.Add_SelectionChangeCommitted({ Invoke-AbmSafely { Invoke-AbmSheetChange } })
    $ui.Grid.Add_SelectionChanged({ Update-AbmActions })
    $ui.Grid.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -ge 0) { Invoke-AbmSafely { Invoke-AbmEdit } }
    })
}

function Register-SearchEvents {
    <# Búsqueda con espera de 300 ms para no filtrar en cada tecla. #>
    $ui = Get-AbmUi
    $ui.Search.Add_TextChanged({ $timer = (Get-AbmUi).SearchTimer; $timer.Stop(); $timer.Start() })
    $ui.SearchTimer.Add_Tick({ (Get-AbmUi).SearchTimer.Stop(); Invoke-AbmSafely { Set-AbmFilter } })
}

function Register-DragDropEvents {
    $ui = Get-AbmUi
    $dragEnter = {
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    }
    $dragDrop = {
        param($s, $e)
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        Invoke-AbmSafely -Action { param($dropped) Invoke-AbmDrop -Files $dropped } -Argument $files
    }
    foreach ($target in @($ui.Form, $ui.Grid, $ui.EmptyLabel)) {
        $target.Add_DragEnter($dragEnter)
        $target.Add_DragDrop($dragDrop)
    }
}

function Register-FormEvents {
    $form = (Get-AbmUi).Form
    $form.Add_KeyDown({
        param($s, $e)
        Invoke-AbmSafely -Action { param($keyEvent) Invoke-AbmShortcut -KeyEvent $keyEvent } -Argument $e
    })
    $form.Add_Shown({
        $path = Get-AbmSetting -Name StartupPath
        if ($path) { Invoke-AbmSafely -Action { param($p) Open-AbmDocument -Path $p } -Argument $path }
    })
    $form.Add_FormClosing({
        param($s, $e)
        if ((Get-AbmState).Busy) { $e.Cancel = $true; return }
        try { if (-not (Confirm-AbmPendingChanges)) { $e.Cancel = $true } }
        catch { $e.Cancel = $false }
    })
}
#endregion

Export-ModuleMember -Function Register-AbmEvents, Invoke-AbmSafely
