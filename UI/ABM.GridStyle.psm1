Set-StrictMode -Version Latest

<#
    Estilo y comportamiento de la grilla de datos: encabezado corporativo, filas alternadas,
    sólo lectura (el ABM se hace por diálogos) y selección de fila completa.
#>

function Set-AbmGridStyle {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid)

    Set-GridAppearance -Grid $Grid
    Set-GridBehavior -Grid $Grid
}

function Set-GridAppearance {
    param([System.Windows.Forms.DataGridView]$Grid)

    $cellPadding = New-AbmPadding -Left 8 -Right 4
    $Grid.BackgroundColor = Get-AbmColor 'Card'
    $Grid.BorderStyle     = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor       = Get-AbmColor 'Border'
    $Grid.EnableHeadersVisualStyles   = $false
    $Grid.ColumnHeadersBorderStyle    = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight         = Get-AbmScaled 38
    $Grid.RowTemplate.Height          = Get-AbmScaled 30
    $Grid.RowHeadersVisible           = $false

    $header = $Grid.ColumnHeadersDefaultCellStyle
    $header.BackColor          = Get-AbmColor 'Primary'
    $header.ForeColor          = Get-AbmColor 'HeaderText'
    $header.SelectionBackColor = Get-AbmColor 'Primary'
    $header.SelectionForeColor = Get-AbmColor 'HeaderText'
    $header.Font               = Get-AbmFont 'GridHeader'
    $header.Padding            = $cellPadding
    $header.Alignment          = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft

    $cells = $Grid.DefaultCellStyle
    $cells.Font               = Get-AbmFont 'Body'
    $cells.BackColor          = Get-AbmColor 'Card'
    $cells.ForeColor          = Get-AbmColor 'Text'
    $cells.SelectionBackColor = Get-AbmColor 'GridSelect'
    $cells.SelectionForeColor = Get-AbmColor 'Text'
    $cells.Padding            = $cellPadding
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = Get-AbmColor 'GridAlt'
}

function Set-GridBehavior {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.AllowUserToAddRows    = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.ReadOnly              = $true
    $Grid.MultiSelect           = $true
    $Grid.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.StandardTab           = $true
}

Export-ModuleMember -Function Set-AbmGridStyle
