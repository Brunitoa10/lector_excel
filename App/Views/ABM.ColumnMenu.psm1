Set-StrictMode -Version Latest

<#
    Menú de columnas (botón "Columnas…" y clic derecho sobre un encabezado de la grilla).
    Las opciones se rotulan con el nombre de la columna elegida.
#>

function New-AbmColumnMenu {
    [CmdletBinding()]
    param()

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $menu.ShowImageMargin = $false
    $menu.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
    $menu.Font = Get-AbmFont 'Body'
    Set-AbmThemeColor -Target $menu -Property BackColor -Role 'Card'

    $ui = Get-AbmUi
    $ui.MenuAddColumn    = New-MenuItem -Text 'Agregar columna…'
    $ui.MenuRenameColumn = New-MenuItem -Text 'Renombrar columna…'
    $ui.MenuRemoveColumn = New-MenuItem -Text 'Quitar columna…'
    [void]$menu.Items.AddRange([System.Windows.Forms.ToolStripItem[]]@($ui.MenuAddColumn, $ui.MenuRenameColumn, $ui.MenuRemoveColumn))
    $ui.ColumnMenu = $menu
    $ui.MenuColumn = -1
}

function Show-AbmColumnMenu {
    <# Muestra el menú para la columna indicada (índice de la tabla; -1 = ninguna) en un punto de pantalla. #>
    [CmdletBinding()]
    param([int]$ColumnIndex = -1, [Parameter(Mandatory)][System.Drawing.Point]$ScreenPoint)

    if (-not (Test-AbmHasData)) { return }
    $ui = Get-AbmUi
    $names = @(Get-AbmColumnNames)
    $hasColumn = $ColumnIndex -ge 0 -and $ColumnIndex -lt $names.Count
    $ui.MenuColumn = if ($hasColumn) { $ColumnIndex } else { -1 }

    $label = if ($hasColumn) { ' «{0}»' -f $names[$ColumnIndex] } else { '' }
    $ui.MenuAddColumn.Text    = if ($hasColumn) { 'Agregar columna después de{0}…' -f $label } else { 'Agregar columna…' }
    $ui.MenuRenameColumn.Text = 'Renombrar columna{0}…' -f $label
    $ui.MenuRemoveColumn.Text = 'Quitar columna{0}…' -f $label
    $ui.MenuRemoveColumn.Enabled = $hasColumn -and $names.Count -gt 1
    $ui.ColumnMenu.Show($ScreenPoint)
}

function New-MenuItem {
    param([string]$Text)
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = $Text
    Set-AbmThemeColor -Target $item -Property ForeColor -Role 'Text'
    return $item
}

Export-ModuleMember -Function New-AbmColumnMenu, Show-AbmColumnMenu
