Set-StrictMode -Version Latest

<#
    Estilo y comportamiento de la grilla de datos: encabezado corporativo, filas alternadas,
    selección de celdas (como en Excel) y edición directa al escribir, con F2 o con doble clic.
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
    $Grid.BorderStyle     = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.EnableHeadersVisualStyles   = $false
    $Grid.ColumnHeadersBorderStyle    = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight         = Get-AbmScaled 38
    $Grid.RowTemplate.Height          = Get-AbmScaled 30
    $Grid.RowHeadersVisible           = $false
    Set-AbmThemeColor -Target $Grid -Property BackgroundColor -Role 'Card'
    Set-AbmThemeColor -Target $Grid -Property GridColor -Role 'Border'

    $header = $Grid.ColumnHeadersDefaultCellStyle
    $header.Font      = Get-AbmFont 'GridHeader'
    $header.Padding   = $cellPadding
    $header.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
    Set-StyleColors -Style $header -Owner $Grid -Back 'Primary' -Fore 'HeaderText' -SelectionBack 'Primary' -SelectionFore 'HeaderText'

    $cells = $Grid.DefaultCellStyle
    $cells.Font    = Get-AbmFont 'Body'
    $cells.Padding = $cellPadding
    Set-StyleColors -Style $cells -Owner $Grid -Back 'Card' -Fore 'Text' -SelectionBack 'GridSelect' -SelectionFore 'Text'
    Set-AbmThemeColor -Target $Grid.AlternatingRowsDefaultCellStyle -Property BackColor -Role 'GridAlt' -Owner $Grid
}

function Set-GridBehavior {
    param([System.Windows.Forms.DataGridView]$Grid)

    $Grid.AllowUserToAddRows    = $false
    $Grid.AllowUserToDeleteRows = $false
    $Grid.AllowUserToResizeRows = $false
    $Grid.ReadOnly              = $false
    $Grid.EditMode              = [System.Windows.Forms.DataGridViewEditMode]::EditOnKeystrokeOrF2
    $Grid.MultiSelect           = $true
    $Grid.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $Grid.StandardTab           = $true
}

function Set-StyleColors {
    param($Style, $Owner, [string]$Back, [string]$Fore, [string]$SelectionBack, [string]$SelectionFore)
    Set-AbmThemeColor -Target $Style -Property BackColor -Role $Back -Owner $Owner
    Set-AbmThemeColor -Target $Style -Property ForeColor -Role $Fore -Owner $Owner
    Set-AbmThemeColor -Target $Style -Property SelectionBackColor -Role $SelectionBack -Owner $Owner
    Set-AbmThemeColor -Target $Style -Property SelectionForeColor -Role $SelectionFore -Owner $Owner
}

Export-ModuleMember -Function Set-AbmGridStyle
