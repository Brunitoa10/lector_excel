Set-StrictMode -Version Latest

<#
    Edición directa en la grilla (alternativa rápida al formulario):
      - Start-AbmCellEdit    : recuerda el valor anterior al empezar a editar.
      - Complete-AbmCellEdit : si el valor cambió, lo audita y marca el archivo como pendiente.
      - Complete-AbmPendingEdit: confirma una edición en curso antes de guardar, abrir, cerrar, etc.
#>

$script:Pending = $null

function Start-AbmCellEdit {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Windows.Forms.DataGridViewCellCancelEventArgs]$EventArgs)

    $grid = (Get-AbmUi).Grid
    $item = $grid.Rows[$EventArgs.RowIndex].DataBoundItem
    $name = $grid.Columns[$EventArgs.ColumnIndex].DataPropertyName
    if (-not ($item -is [System.Data.DataRowView])) { $script:Pending = $null; return }
    $script:Pending = @{ Item = $item; Column = $name; Value = [string]$item[$name] }
}

function Complete-AbmCellEdit {
    [CmdletBinding()]
    param()

    $pending = $script:Pending
    $script:Pending = $null
    if ($null -eq $pending) { return }

    # DataRowView devuelve el valor propuesto mientras la fila sigue en edición.
    $newValue = [string]$pending.Item[$pending.Column]
    if ($newValue -ceq $pending.Value) { return }

    $detail = "{0}: '{1}' → '{2}' (edición en grilla)" -f $pending.Column, $pending.Value, $newValue
    Write-AbmAudit -Action 'MODIFICACION' -Target (Get-AbmSourceLabel) -Detail $detail
    Set-AbmDirty
    Set-AbmStatus ("Celda «{0}» modificada. Recordá guardar los cambios (Ctrl+S)." -f $pending.Column)
}

function Complete-AbmPendingEdit {
    <# Cierra la edición de la celda y confirma la fila actual en la tabla. #>
    [CmdletBinding()]
    param()

    $grid = (Get-AbmUi).Grid
    $view = (Get-AbmState).View
    if ($null -eq $grid -or $null -eq $view) { return }
    if ($grid.IsCurrentCellInEditMode) { [void]$grid.EndEdit() }
    $grid.BindingContext[$view].EndCurrentEdit()
}

Export-ModuleMember -Function Start-AbmCellEdit, Complete-AbmCellEdit, Complete-AbmPendingEdit
