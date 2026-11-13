Set-StrictMode -Version Latest

<#
    Refresco de la vista a partir del estado (la vista nunca guarda estado propio):
    indicadores, título, encabezado, botones habilitados, selector de hoja y columnas de la grilla.
#>

function Update-AbmView {
    [CmdletBinding()]
    param()

    $ui = Get-AbmUi
    $hasData = Test-AbmHasData
    $ui.Grid.Visible       = $hasData
    $ui.EmptyLabel.Visible = -not $hasData

    if ($hasData) { Update-DataSummary } else { Clear-DataSummary }
    Update-AbmActions
}

function Update-AbmActions {
    [CmdletBinding()]
    param()

    $ui = Get-AbmUi
    $hasData  = Test-AbmHasData
    $selected = Get-AbmSelectedRowCount

    $ui.BtnAdd.Enabled    = $hasData
    $ui.BtnEdit.Enabled   = $selected -eq 1
    $ui.BtnDelete.Enabled = $selected -ge 1
    $ui.BtnSave.Enabled   = $hasData -and (Get-AbmState).Dirty
    $ui.BtnSaveAs.Enabled = $hasData
    $ui.BtnColumns.Enabled = $hasData
    $ui.Search.Enabled    = $hasData
}

function Update-AbmSheetSelector {
    [CmdletBinding()]
    param()

    $combo  = (Get-AbmUi).SheetCombo
    $source = (Get-AbmState).Source
    $combo.BeginUpdate()
    try {
        $combo.Items.Clear()
        if ($null -ne $source -and $source.Kind -eq 'Excel') {
            foreach ($name in @($source.Meta.Sheets)) { [void]$combo.Items.Add($name) }
            $combo.SelectedItem = $source.Meta.Sheet
            $combo.Enabled = $combo.Items.Count -gt 1
        }
        else {
            [void]$combo.Items.Add('— (archivo de texto)')
            $combo.SelectedIndex = 0
            $combo.Enabled = $false
        }
    }
    finally { $combo.EndUpdate() }
}

function Set-AbmGridColumns {
    <# Hasta 6 columnas: ocupan todo el ancho. Más: ancho según contenido, entre 110 y 320 px. #>
    [CmdletBinding()]
    param()

    $grid = (Get-AbmUi).Grid
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    foreach ($column in $grid.Columns) {
        $column.SortMode     = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $column.MinimumWidth = Get-AbmScaled 70
    }
    if ($grid.Columns.Count -le 6) {
        $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        return
    }
    $grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
    $min = Get-AbmScaled 110
    $max = Get-AbmScaled 320
    foreach ($column in $grid.Columns) { $column.Width = [math]::Min([math]::Max($column.Width, $min), $max) }
}

function Select-AbmRow {
    <# Selecciona y muestra en la grilla la fila indicada (si está visible con el filtro actual). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Data.DataRow]$Row)

    $grid = (Get-AbmUi).Grid
    for ($i = $grid.Rows.Count - 1; $i -ge 0; $i--) {
        $item = $grid.Rows[$i].DataBoundItem
        if ($item -is [System.Data.DataRowView] -and [object]::ReferenceEquals($item.Row, $Row)) {
            $grid.ClearSelection()
            $grid.CurrentCell = $grid.Rows[$i].Cells[0]
            try { $grid.FirstDisplayedScrollingRowIndex = $i } catch { Write-Verbose 'Fila fuera del área visible.' }
            return
        }
    }
}

#region Internos
function Update-DataSummary {
    $ui     = Get-AbmUi
    $state  = Get-AbmState
    $source = $state.Source
    $total   = $state.Table.Rows.Count
    $visible = $state.View.Count

    $rowsCaption = if ($visible -ne $total) { 'Filtro activo: {0:N0} visibles' -f $visible } else { 'Registros en el archivo' }
    Set-AbmKpiCard -Card $ui.KpiRows -Value ('{0:N0}' -f $total) -Caption $rowsCaption
    Set-AbmKpiCard -Card $ui.KpiCols -Value ('{0:N0}' -f $state.Table.Columns.Count) -Caption 'Campos por registro'
    Set-AbmKpiCard -Card $ui.KpiSource -Value (Get-SourceKindText $source) -Caption (Get-SourceDetail $source)
    if ($state.Dirty) { Set-AbmKpiCard -Card $ui.KpiState -Value 'Pendiente' -Caption 'Hay cambios sin guardar' -Color 'Warning' }
    else              { Set-AbmKpiCard -Card $ui.KpiState -Value 'Al día' -Caption 'Todos los cambios guardados' -Color 'Success' }

    $ui.HeaderFile.Text = '{0}{1}{2}' -f [System.IO.Path]::GetFileName($source.Path), [Environment]::NewLine, [System.IO.Path]::GetDirectoryName($source.Path)
    $marker = if ($state.Dirty) { '● ' } else { '' }
    $ui.Form.Text = '{0}{1} — {2}' -f $marker, (Get-AbmSourceLabel), (Get-AbmSetting -Name AppTitle)
}

function Clear-DataSummary {
    $ui = Get-AbmUi
    foreach ($card in @($ui.KpiRows, $ui.KpiCols, $ui.KpiSource, $ui.KpiState)) { Set-AbmKpiCard -Card $card }
    $ui.HeaderFile.Text = 'Sin archivo abierto'
    $ui.Form.Text = Get-AbmSetting -Name AppTitle
}

function Get-SourceKindText {
    param($Source)
    if ($Source.Kind -eq 'Excel') { return 'Excel' }
    return 'TXT'
}

function Get-SourceDetail {
    param($Source)
    if ($Source.Kind -eq 'Excel') {
        return 'Hoja «{0}» · {1} hoja(s) en el libro' -f $Source.Meta.Sheet, @($Source.Meta.Sheets).Count
    }
    return '{0} · {1}' -f (Get-AbmDelimiterName -Delimiter $Source.Meta.Delimiter), $Source.Meta.EncodingName
}
#endregion

Export-ModuleMember -Function Update-AbmView, Update-AbmActions, Update-AbmSheetSelector, Set-AbmGridColumns, Select-AbmRow
