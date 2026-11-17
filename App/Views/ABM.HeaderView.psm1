Set-StrictMode -Version Latest

<#
    Encabezado corporativo: título, subtítulo, el archivo abierto y el botón de tema (Claro / Oscuro).
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

    # Se acoplan en orden inverso al agregado: primero la línea dorada (ancho completo),
    # después el botón de tema (borde derecho) y a su izquierda el archivo.
    $panel.Controls.AddRange([System.Windows.Forms.Control[]]@($file, (New-ThemeToggle), $accent, $title, $subtitle))
    (Get-AbmUi).HeaderFile = $file
    return $panel
}

function New-ThemeToggle {
    $holder = New-AbmPanel -Dock Right -Width 150
    $button = New-AbmButton -Text (Get-AbmThemeToggleText) -Variant Header -Width 124
    $button.Location = New-Object System.Drawing.Point(0, (Get-AbmScaled 22))
    $holder.Controls.Add($button)

    $ui = Get-AbmUi
    $ui.BtnTheme = $button
    $ui.Tips.SetToolTip($button, 'Cambiar entre modo claro y oscuro (Ctrl+T)')
    return $holder
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
    Set-AbmThemeColor -Target $label -Property ForeColor -Role 'HeaderMuted'
    return $label
}

Export-ModuleMember -Function New-AbmHeaderView
