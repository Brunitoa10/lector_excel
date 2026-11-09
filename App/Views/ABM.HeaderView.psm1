Set-StrictMode -Version Latest

<#
    Encabezado corporativo: título, subtítulo y, a la derecha, el archivo abierto.
#>

function New-AbmHeaderView {
    [CmdletBinding()]
    param()

    $panel  = New-AbmPanel -Dock Top -Height 78 -Color 'Primary'
    $accent = New-AbmPanel -Dock Bottom -Height 3 -Color 'Gold'
    $file   = New-FileLabel

    $title    = New-AbmLabel -Text 'ABM de Datos' -Font 'Title' -Color 'HeaderText' -X 22 -Y 10
    $subtitle = New-AbmLabel -Text 'Alta, baja y modificación de registros · Excel (.xlsx) y texto (.txt)' `
                             -Font 'Subtitle' -Color 'HeaderMuted' -X 24 -Y 47

    $panel.Controls.AddRange([System.Windows.Forms.Control[]]@($accent, $file, $title, $subtitle))
    (Get-AbmUi).HeaderFile = $file
    return $panel
}

function New-FileLabel {
    $label = New-Object System.Windows.Forms.Label
    $label.Dock         = [System.Windows.Forms.DockStyle]::Right
    $label.Width        = Get-AbmScaled 560
    $label.AutoSize     = $false
    $label.AutoEllipsis = $true
    $label.TextAlign    = [System.Drawing.ContentAlignment]::MiddleRight
    $label.Padding      = New-AbmPadding -Right 22 -Bottom 3
    $label.Font         = Get-AbmFont 'Label'
    $label.ForeColor    = Get-AbmColor 'HeaderMuted'
    return $label
}

Export-ModuleMember -Function New-AbmHeaderView
