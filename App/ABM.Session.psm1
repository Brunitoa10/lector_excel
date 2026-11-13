Set-StrictMode -Version Latest

<#
    Sesión de trabajo: transiciones del estado (abrir origen, marcar cambios) y consultas sobre
    el archivo abierto y la selección actual.
#>

function Set-AbmSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Source)

    $state = Get-AbmState
    $state.Source = $Source
    $state.Table  = $Source.Table
    $state.View   = [System.Data.DataView]::new($Source.Table)
    $state.Dirty  = $false

    $ui = Get-AbmUi
    $ui.SearchTimer.Stop()
    $ui.Search.Text = ''
    $ui.SearchTimer.Stop()
    $ui.Grid.DataSource = $state.View

    Set-AbmGridColumns
    Update-AbmSheetSelector
    Update-AbmView
}

function Set-AbmDirty {
    [CmdletBinding()]
    param([bool]$Value = $true)
    (Get-AbmState).Dirty = $Value
    Update-AbmView
}

function Test-AbmHasData {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    return ($null -ne (Get-AbmState).Table)
}

function Get-AbmSourceLabel {
    <# "archivo.xlsx [Hoja]" o "archivo.txt". #>
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $source = (Get-AbmState).Source
    if ($null -eq $source) { return '' }
    $name = [System.IO.Path]::GetFileName($source.Path)
    if ($source.Kind -eq 'Excel') { return '{0} [{1}]' -f $name, $source.Meta.Sheet }
    return $name
}

function Get-AbmColumnNames {
    [CmdletBinding()]
    param()
    foreach ($column in (Get-AbmState).Table.Columns) { $column.ColumnName }
}

function Get-AbmSelectedRows {
    <# DataRows que tienen al menos una celda seleccionada (respeta filtro y orden de la vista). #>
    [CmdletBinding()]
    param()
    if (-not (Test-AbmHasData)) { return }
    $grid = (Get-AbmUi).Grid
    $seen = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($cell in $grid.SelectedCells) {
        if (-not $seen.Add($cell.RowIndex)) { continue }
        $item = $grid.Rows[$cell.RowIndex].DataBoundItem
        if ($item -is [System.Data.DataRowView]) { $item.Row }
    }
}

function Get-AbmSelectedRowCount {
    <# Cantidad de filas seleccionadas; con selecciones enormes devuelve $Limit (sólo importa si es 0, 1 o más). #>
    [CmdletBinding()]
    [OutputType([int])]
    param([int]$Limit = 2000)

    if (-not (Test-AbmHasData)) { return 0 }
    $cells = (Get-AbmUi).Grid.SelectedCells
    if ($cells.Count -gt $Limit) { return $Limit }
    $rows = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($cell in $cells) { [void]$rows.Add($cell.RowIndex) }
    return $rows.Count
}

function Get-AbmCurrentColumnIndex {
    <# Índice (en la tabla) de la columna de la celda actual; -1 si no hay. #>
    [CmdletBinding()]
    [OutputType([int])]
    param()
    $grid = (Get-AbmUi).Grid
    if (-not (Test-AbmHasData) -or $null -eq $grid.CurrentCell) { return -1 }
    return (ConvertTo-AbmTableColumnIndex -GridColumnIndex $grid.CurrentCell.ColumnIndex)
}

function ConvertTo-AbmTableColumnIndex {
    <# La grilla puede ordenar columnas distinto que la tabla: se traduce por nombre. #>
    [CmdletBinding()]
    [OutputType([int])]
    param([int]$GridColumnIndex)
    if ($GridColumnIndex -lt 0) { return -1 }
    $name = (Get-AbmUi).Grid.Columns[$GridColumnIndex].DataPropertyName
    return (Get-AbmState).Table.Columns.IndexOf($name)
}

function Suspend-AbmGridBinding {
    <# Desconecta la grilla antes de cambiar la estructura de la tabla (columnas). #>
    [CmdletBinding()]
    param()
    $view = (Get-AbmState).View
    $view.RowFilter = ''
    $view.Sort = ''
    (Get-AbmUi).Grid.DataSource = $null
}

function Resume-AbmGridBinding {
    [CmdletBinding()]
    param()
    (Get-AbmUi).Grid.DataSource = (Get-AbmState).View
    Set-AbmGridColumns
    Set-AbmFilter
}

Export-ModuleMember -Function Set-AbmSource, Set-AbmDirty, Test-AbmHasData, Get-AbmSourceLabel, Get-AbmColumnNames, Get-AbmSelectedRows,
                              Get-AbmSelectedRowCount, Get-AbmCurrentColumnIndex, ConvertTo-AbmTableColumnIndex,
                              Suspend-AbmGridBinding, Resume-AbmGridBinding
