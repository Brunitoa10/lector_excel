Set-StrictMode -Version Latest

<#
    Área de datos: tarjeta con barra superior (selector de hoja + búsqueda), grilla y
    mensaje de estado vacío cuando no hay archivo abierto.
#>

$script:DesignWidth = 1000   # ancho de referencia para anclar los controles a la derecha

function New-AbmGridView {
    [CmdletBinding()]
    param()

    $area = New-AbmPanel -Dock Fill -Color 'Surface'
    $area.Padding = New-AbmPadding -Left 20 -Top 4 -Right 20 -Bottom 14
    $frame = New-AbmPanel -Dock Fill -Color 'Border'
    $frame.Padding = New-Object System.Windows.Forms.Padding(1)
    $inner = New-AbmPanel -Dock Fill -Color 'Card'
    $separator = New-AbmPanel -Dock Top -Color 'Border'
    $separator.Height = 1

    $inner.Controls.AddRange([System.Windows.Forms.Control[]]@((New-DataGrid), (New-EmptyState), $separator, (New-GridBar)))
    $frame.Controls.Add($inner)
    $area.Controls.Add($frame)
    return $area
}

function New-GridBar {
    $bar = New-AbmPanel -Dock Top -Height 54 -Color 'Card'
    $bar.Width = Get-AbmScaled $script:DesignWidth
    $right = [System.Windows.Forms.AnchorStyles]'Top, Right'

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $combo.FlatStyle     = [System.Windows.Forms.FlatStyle]::Flat
    $combo.Width         = Get-AbmScaled 220
    $combo.Location      = New-Object System.Drawing.Point((Get-AbmScaled 62), (Get-AbmScaled 15))

    $search = New-Object System.Windows.Forms.TextBox
    $search.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $search.Width       = Get-AbmScaled 280
    $search.Location    = New-Object System.Drawing.Point(($bar.Width - (Get-AbmScaled 296)), (Get-AbmScaled 15))
    $search.Anchor      = $right

    $searchLabel = New-AbmLabel -Text 'BUSCAR' -Font 'KpiTitle' -Color 'Muted' -Y 20
    $searchLabel.Left   = $bar.Width - (Get-AbmScaled 356)
    $searchLabel.Anchor = $right

    $sheetLabel = New-AbmLabel -Text 'HOJA' -Font 'KpiTitle' -Color 'Muted' -X 16 -Y 20
    $bar.Controls.AddRange([System.Windows.Forms.Control[]]@($sheetLabel, $combo, $searchLabel, $search))

    $ui = Get-AbmUi
    $ui.SheetCombo  = $combo
    $ui.Search      = $search
    $ui.SearchTimer = New-Object System.Windows.Forms.Timer
    $ui.SearchTimer.Interval = 300
    return $bar
}

function New-DataGrid {
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $grid.AllowDrop = $true
    Set-AbmGridStyle -Grid $grid
    (Get-AbmUi).Grid = $grid
    return $grid
}

function New-EmptyState {
    $label = New-Object System.Windows.Forms.Label
    $label.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $label.Font      = Get-AbmFont 'EmptyState'
    $label.ForeColor = Get-AbmColor 'Muted'
    $label.AllowDrop = $true
    $label.Text      = "Abrí un archivo Excel o TXT para comenzar`r`n`r`nUsá «Abrir…», «Nuevo archivo» o arrastrá el archivo a esta ventana"
    (Get-AbmUi).EmptyLabel = $label
    return $label
}

Export-ModuleMember -Function New-AbmGridView
