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
    <# DataRows seleccionadas en la grilla (respeta filtro y orden de la vista). #>
    [CmdletBinding()]
    param()
    if (-not (Test-AbmHasData)) { return }
    foreach ($gridRow in (Get-AbmUi).Grid.SelectedRows) {
        $item = $gridRow.DataBoundItem
        if ($item -is [System.Data.DataRowView]) { $item.Row }
    }
}

Export-ModuleMember -Function Set-AbmSource, Set-AbmDirty, Test-AbmHasData, Get-AbmSourceLabel, Get-AbmColumnNames, Get-AbmSelectedRows
