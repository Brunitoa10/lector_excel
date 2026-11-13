Set-StrictMode -Version Latest

<#
    Tarjeta de indicador (KPI): franja de color + título + valor grande + leyenda.
    Devuelve los controles internos para que la vista pueda actualizarlos.
#>

function New-AbmKpiCard {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$AccentColor)

    $card = New-AbmPanel -Dock Fill -Color 'Card'
    $card.Margin = New-AbmPadding -Right 14
    $card.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen((Get-AbmColor 'Border'))
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width - 1, $s.Height - 1)
        $pen.Dispose()
    })

    $strip = New-AbmPanel -Dock Left -Width 5 -Color $AccentColor
    $body  = New-AbmPanel -Dock Fill
    $body.Padding = New-AbmPadding -Left 14 -Top 8 -Right 8 -Bottom 6

    $titleLabel = New-CardLabel -Text $Title.ToUpperInvariant() -Font 'KpiTitle' -Color 'Muted' -Height 18
    $valueLabel = New-CardLabel -Text '—' -Font 'Kpi' -Color 'Text' -Height 40
    $caption    = New-CardLabel -Text '' -Font 'Small' -Color 'Muted' -Height 18

    # Dock Top apila en orden inverso de agregado: el último queda arriba.
    $body.Controls.AddRange([System.Windows.Forms.Control[]]@($caption, $valueLabel, $titleLabel))
    $card.Controls.AddRange([System.Windows.Forms.Control[]]@($body, $strip))

    return [pscustomobject]@{ Panel = $card; Strip = $strip; Value = $valueLabel; Caption = $caption }
}

function Set-AbmKpiCard {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Card, [string]$Value = '—', [string]$Caption = '', [string]$Color = '')

    $Card.Value.Text   = $Value
    $Card.Caption.Text = $Caption
    if ($Color) {
        Set-AbmThemeColor -Target $Card.Strip -Property BackColor -Role $Color
        Set-AbmThemeColor -Target $Card.Value -Property ForeColor -Role $Color
    }
}

function New-CardLabel {
    param([string]$Text, [string]$Font, [string]$Color, [int]$Height)

    $label = New-Object System.Windows.Forms.Label
    $label.Text         = $Text
    $label.Font         = Get-AbmFont $Font
    $label.AutoSize     = $false
    $label.AutoEllipsis = $true
    $label.Dock         = [System.Windows.Forms.DockStyle]::Top
    $label.Height       = Get-AbmScaled $Height
    $label.TextAlign    = [System.Drawing.ContentAlignment]::MiddleLeft
    Set-AbmThemeColor -Target $label -Property ForeColor -Role $Color
    return $label
}

Export-ModuleMember -Function New-AbmKpiCard, Set-AbmKpiCard
