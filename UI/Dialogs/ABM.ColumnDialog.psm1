Set-StrictMode -Version Latest

<#
    Diálogo de columnas:
      -Mode Add    -> nombre + ubicación.    Devuelve @{ Name; Position } (Position = índice donde insertar).
      -Mode Rename -> columna + nuevo nombre. Devuelve @{ Index; Name }.
    La validación del nombre la inyecta quien llama (-Validator), así el diálogo no conoce la tabla.
#>

$script:Dialog = @{}

function Show-AbmColumnDialog {
    [CmdletBinding()]
    param(
        $Owner,
        [Parameter(Mandatory)][ValidateSet('Add', 'Rename')][string]$Mode,
        [Parameter(Mandatory)][string[]]$Columns,
        [int]$Selected = -1,
        [Parameter(Mandatory)][scriptblock]$Validator
    )

    if ($Mode -eq 'Add') { $shell = New-AddLayout -Columns $Columns -Selected $Selected }
    else { $shell = New-RenameLayout -Columns $Columns -Selected $Selected }
    $script:Dialog.Mode      = $Mode
    $script:Dialog.Validator = $Validator
    $shell.Ok.Add_Click({ Confirm-ColumnDialog })
    $shell.Form.Add_Shown({ [void]$script:Dialog.Name.Focus(); $script:Dialog.Name.SelectAll() })

    try {
        if ((Invoke-AbmDialog -Shell $shell -Owner $Owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        $name = $script:Dialog.Name.Text.Trim()
        if ($Mode -eq 'Add') { return @{ Name = $name; Position = $script:Dialog.Choice.SelectedIndex } }
        return @{ Index = $script:Dialog.Choice.SelectedIndex; Name = $name }
    }
    finally {
        $shell.Form.Dispose()
        $script:Dialog = @{}
    }
}

#region Diseños
function New-AddLayout {
    param([string[]]$Columns, [int]$Selected)

    $shell = New-AbmDialogShell -Title 'Agregar columna' -Subtitle 'La columna nueva queda vacía en todos los registros' -BodyHeight 170
    $places = @('Al principio') + @($Columns | ForEach-Object { 'Después de «{0}»' -f $_ })
    $default = if ($Selected -ge 0 -and $Selected -lt $Columns.Count) { $Selected + 1 } else { $Columns.Count }
    $name   = New-AbmTextBox -MaxLength 255
    $choice = New-AbmComboBox -Items $places -Selected $default

    Add-AbmDialogField -Shell $shell -Caption 'Nombre de la columna' -Control $name
    Add-AbmDialogField -Shell $shell -Caption 'Ubicación' -Control $choice
    $script:Dialog = @{ Form = $shell.Form; Name = $name; Choice = $choice; Columns = $Columns }
    return $shell
}

function New-RenameLayout {
    param([string[]]$Columns, [int]$Selected)

    $shell = New-AbmDialogShell -Title 'Renombrar columna' -Subtitle 'Los datos de la columna no cambian' -BodyHeight 170
    $index  = if ($Selected -ge 0 -and $Selected -lt $Columns.Count) { $Selected } else { 0 }
    $choice = New-AbmComboBox -Items $Columns -Selected $index
    $name   = New-AbmTextBox -Text $Columns[$index] -MaxLength 255

    Add-AbmDialogField -Shell $shell -Caption 'Columna' -Control $choice
    Add-AbmDialogField -Shell $shell -Caption 'Nuevo nombre' -Control $name
    $script:Dialog = @{ Form = $shell.Form; Name = $name; Choice = $choice; Columns = $Columns }
    $choice.Add_SelectedIndexChanged({ $script:Dialog.Name.Text = $script:Dialog.Columns[$script:Dialog.Choice.SelectedIndex] })
    return $shell
}
#endregion

function Confirm-ColumnDialog {
    $except  = if ($script:Dialog.Mode -eq 'Rename') { $script:Dialog.Choice.SelectedIndex } else { -1 }
    $problem = & $script:Dialog.Validator $script:Dialog.Name.Text $except
    if ($problem) {
        Show-AbmMessage -Owner $script:Dialog.Form -Message $problem -Icon Warning
        return
    }
    $script:Dialog.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
}

Export-ModuleMember -Function Show-AbmColumnDialog
