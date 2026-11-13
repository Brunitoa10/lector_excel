Set-StrictMode -Version Latest

<#
    Administración de columnas: agregar, renombrar y quitar.
    Cada cambio de estructura desconecta la grilla, modifica la tabla y la vuelve a conectar,
    reaplicando la búsqueda activa.
#>

function Invoke-AbmAddColumn {
    [CmdletBinding()]
    param([int]$ColumnIndex = -1)

    if (-not (Test-AbmHasData)) { return }
    Complete-AbmPendingEdit
    $result = Show-AbmColumnDialog -Owner (Get-AbmUi).Form -Mode Add -Columns @(Get-AbmColumnNames) `
                                   -Selected $ColumnIndex -Validator (Get-NameValidator)
    if ($null -eq $result) { return }

    $table = (Get-AbmState).Table
    Invoke-SchemaChange { Add-AbmTableColumn -Table $table -Name $result.Name -Position $result.Position }
    Complete-ColumnChange -Action 'COLUMNA_ALTA' -Detail ('«{0}» en la posición {1}' -f $result.Name, ($result.Position + 1)) `
                          -Status ('Columna «{0}» agregada.' -f $result.Name)
}

function Invoke-AbmRenameColumn {
    [CmdletBinding()]
    param([int]$ColumnIndex = -1)

    if (-not (Test-AbmHasData)) { return }
    Complete-AbmPendingEdit
    $names  = @(Get-AbmColumnNames)
    $result = Show-AbmColumnDialog -Owner (Get-AbmUi).Form -Mode Rename -Columns $names `
                                   -Selected $ColumnIndex -Validator (Get-NameValidator)
    if ($null -eq $result -or $result.Name -ceq $names[$result.Index]) { return }

    $table = (Get-AbmState).Table
    Invoke-SchemaChange { Rename-AbmTableColumn -Table $table -Index $result.Index -Name $result.Name }
    Complete-ColumnChange -Action 'COLUMNA_RENOMBRAR' -Detail ("'{0}' → '{1}'" -f $names[$result.Index], $result.Name) `
                          -Status ('Columna renombrada a «{0}».' -f $result.Name)
}

function Invoke-AbmRemoveColumn {
    [CmdletBinding()]
    param([int]$ColumnIndex = -1)

    if (-not (Test-AbmHasData) -or $ColumnIndex -lt 0) { return }
    Complete-AbmPendingEdit
    $table = (Get-AbmState).Table
    $name  = $table.Columns[$ColumnIndex].ColumnName
    $message = "¿Quitar la columna «{0}»?`n`nSe borrarán sus datos en los {1:N0} registros. El cambio se aplicará al archivo cuando guardes." -f $name, $table.Rows.Count
    if (-not (Confirm-AbmAction -Owner (Get-AbmUi).Form -Title 'Quitar columna' -Message $message)) { return }

    Invoke-SchemaChange { Remove-AbmTableColumn -Table $table -Index $ColumnIndex }
    Complete-ColumnChange -Action 'COLUMNA_BAJA' -Detail ('«{0}»' -f $name) -Status ('Columna «{0}» quitada.' -f $name)
}

#region Internos
function Get-NameValidator {
    return { param($name, $except) Test-AbmColumnName -Table (Get-AbmState).Table -Name $name -ExceptIndex $except }
}

function Invoke-SchemaChange {
    param([scriptblock]$Change)
    Suspend-AbmGridBinding
    try { & $Change }
    finally { Resume-AbmGridBinding }
}

function Complete-ColumnChange {
    param([string]$Action, [string]$Detail, [string]$Status)
    Write-AbmAudit -Action $Action -Target (Get-AbmSourceLabel) -Detail $Detail
    Set-AbmDirty
    Set-AbmStatus "$Status Recordá guardar los cambios (Ctrl+S)."
}
#endregion

Export-ModuleMember -Function Invoke-AbmAddColumn, Invoke-AbmRenameColumn, Invoke-AbmRemoveColumn
