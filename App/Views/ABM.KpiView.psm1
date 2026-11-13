Set-StrictMode -Version Latest

<#
    Fila de indicadores: Registros, Columnas, Origen y Estado.
#>

$script:Cards = @(
    @{ Name = 'KpiRows';   Title = 'Registros'; Color = 'Accent' }
    @{ Name = 'KpiCols';   Title = 'Columnas';  Color = 'PrimaryLight' }
    @{ Name = 'KpiSource'; Title = 'Origen';    Color = 'Gold' }
    @{ Name = 'KpiState';  Title = 'Estado';    Color = 'Success' }
)

function New-AbmKpiView {
    [CmdletBinding()]
    param()

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock        = [System.Windows.Forms.DockStyle]::Top
    $layout.Height      = Get-AbmScaled 108
    $layout.Padding     = New-AbmPadding -Left 20 -Top 14 -Right 6 -Bottom 8
    Set-AbmThemeColor -Target $layout -Property BackColor -Role 'Surface'
    $layout.ColumnCount = $script:Cards.Count
    $layout.RowCount    = 1
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $ui = Get-AbmUi
    for ($i = 0; $i -lt $script:Cards.Count; $i++) {
        $spec = $script:Cards[$i]
        [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100 / $script:Cards.Count))))
        $card = New-AbmKpiCard -Title $spec.Title -AccentColor $spec.Color
        $layout.Controls.Add($card.Panel, $i, 0)
        $ui[$spec.Name] = $card
    }
    return $layout
}

Export-ModuleMember -Function New-AbmKpiView
