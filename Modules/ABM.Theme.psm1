Set-StrictMode -Version Latest

<#
    Sistema visual (paleta, tipografías, escala DPI) y fábrica de componentes reutilizables.
    Centralizar el estilo acá evita "colores mágicos" repartidos por la UI (DRY).
#>

$script:Palette = @{
    Primary        = [System.Drawing.Color]::FromArgb(17, 34, 64)
    PrimaryLight   = [System.Drawing.Color]::FromArgb(52, 78, 125)
    Accent         = [System.Drawing.Color]::FromArgb(47, 111, 237)
    AccentHover    = [System.Drawing.Color]::FromArgb(33, 90, 204)
    Gold           = [System.Drawing.Color]::FromArgb(201, 162, 39)
    Success        = [System.Drawing.Color]::FromArgb(36, 145, 98)
    SuccessHover   = [System.Drawing.Color]::FromArgb(28, 122, 82)
    Warning        = [System.Drawing.Color]::FromArgb(214, 150, 0)
    Danger         = [System.Drawing.Color]::FromArgb(200, 55, 55)
    DangerSoft     = [System.Drawing.Color]::FromArgb(253, 236, 236)
    Surface        = [System.Drawing.Color]::FromArgb(243, 245, 249)
    Card           = [System.Drawing.Color]::White
    Border         = [System.Drawing.Color]::FromArgb(221, 227, 236)
    Text           = [System.Drawing.Color]::FromArgb(30, 38, 51)
    Muted          = [System.Drawing.Color]::FromArgb(107, 118, 137)
    HeaderText     = [System.Drawing.Color]::White
    HeaderMuted    = [System.Drawing.Color]::FromArgb(168, 182, 206)
    GridAlt        = [System.Drawing.Color]::FromArgb(248, 250, 253)
    GridSelect     = [System.Drawing.Color]::FromArgb(220, 232, 255)
    SecondaryHover = [System.Drawing.Color]::FromArgb(234, 239, 247)
}

$script:Fonts = @{
    Body        = New-Object System.Drawing.Font('Segoe UI', 9.75)
    Small       = New-Object System.Drawing.Font('Segoe UI', 8.5)
    Label       = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    Button      = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
    Title       = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
    Subtitle    = New-Object System.Drawing.Font('Segoe UI', 9.5)
    DialogTitle = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    Kpi         = New-Object System.Drawing.Font('Segoe UI Semibold', 20)
    KpiTitle    = New-Object System.Drawing.Font('Segoe UI Semibold', 8.25)
    GridHeader  = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
    EmptyState  = New-Object System.Drawing.Font('Segoe UI', 12)
}

$script:Scale = 1.0
try {
    $graphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    $script:Scale = [math]::Max(1.0, $graphics.DpiX / 96.0)
    $graphics.Dispose()
}
catch { $script:Scale = 1.0 }

function Get-AbmColor {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    if (-not $script:Palette.ContainsKey($Name)) { throw "Color no definido en la paleta: $Name" }
    return $script:Palette[$Name]
}

function Get-AbmFont {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    if (-not $script:Fonts.ContainsKey($Name)) { throw "Fuente no definida: $Name" }
    return $script:Fonts[$Name]
}

function Get-AbmScaled {
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][int]$Value)
    return [int][math]::Round($Value * $script:Scale)
}

function New-AbmButton {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Primary', 'Secondary', 'Success', 'Danger')][string]$Variant = 'Secondary',
        [int]$Width = 112
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text      = $Text
    $button.Width     = Get-AbmScaled $Width
    $button.Height    = Get-AbmScaled 34
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.Font      = $script:Fonts.Button
    $button.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $button.Margin    = New-Object System.Windows.Forms.Padding(0, 0, (Get-AbmScaled 8), 0)
    $button.UseVisualStyleBackColor = $false

    $p = $script:Palette
    $scheme = switch ($Variant) {
        'Primary' { @{ Back = $p.Accent;  Hover = $p.AccentHover;    Fore = $p.HeaderText; Border = $p.Accent } }
        'Success' { @{ Back = $p.Success; Hover = $p.SuccessHover;   Fore = $p.HeaderText; Border = $p.Success } }
        'Danger'  { @{ Back = $p.Card;    Hover = $p.DangerSoft;     Fore = $p.Danger;     Border = $p.Danger } }
        default   { @{ Back = $p.Card;    Hover = $p.SecondaryHover; Fore = $p.Text;       Border = $p.Border } }
    }

    $button.BackColor = $scheme.Back
    $button.ForeColor = $scheme.Fore
    $button.FlatAppearance.BorderSize         = 1
    $button.FlatAppearance.BorderColor        = $scheme.Border
    $button.FlatAppearance.MouseOverBackColor = $scheme.Hover
    $button.FlatAppearance.MouseDownBackColor = $scheme.Hover
    return $button
}

function New-AbmKpiCard {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][System.Drawing.Color]$AccentColor
    )

    $card = New-Object System.Windows.Forms.Panel
    $card.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $card.BackColor = $script:Palette.Card
    $card.Margin    = New-Object System.Windows.Forms.Padding(0, 0, (Get-AbmScaled 14), 0)
    $card.Add_Paint({
        param($sender, $e)
        $pen = New-Object System.Drawing.Pen((Get-AbmColor 'Border'))
        $e.Graphics.DrawRectangle($pen, 0, 0, $sender.Width - 1, $sender.Height - 1)
        $pen.Dispose()
    })

    $strip = New-Object System.Windows.Forms.Panel
    $strip.Dock      = [System.Windows.Forms.DockStyle]::Left
    $strip.Width     = Get-AbmScaled 5
    $strip.BackColor = $AccentColor

    $body = New-Object System.Windows.Forms.Panel
    $body.Dock    = [System.Windows.Forms.DockStyle]::Fill
    $body.Padding = New-Object System.Windows.Forms.Padding((Get-AbmScaled 14), (Get-AbmScaled 8), (Get-AbmScaled 8), (Get-AbmScaled 6))

    $titleLabel = New-AbmCardLabel -Text $Title.ToUpperInvariant() -Font $script:Fonts.KpiTitle -Color $script:Palette.Muted -Height 18
    $valueLabel = New-AbmCardLabel -Text '—' -Font $script:Fonts.Kpi -Color $script:Palette.Text -Height 40
    $caption    = New-AbmCardLabel -Text '' -Font $script:Fonts.Small -Color $script:Palette.Muted -Height 18

    # Dock Top se apila en orden inverso de agregado: el último queda arriba.
    $body.Controls.AddRange([System.Windows.Forms.Control[]]@($caption, $valueLabel, $titleLabel))
    $card.Controls.AddRange([System.Windows.Forms.Control[]]@($body, $strip))

    return [pscustomobject]@{ Panel = $card; Strip = $strip; Value = $valueLabel; Caption = $caption }
}

function New-AbmCardLabel {
    param([string]$Text, [System.Drawing.Font]$Font, [System.Drawing.Color]$Color, [int]$Height)

    $label = New-Object System.Windows.Forms.Label
    $label.Text         = $Text
    $label.Font         = $Font
    $label.ForeColor    = $Color
    $label.AutoSize     = $false
    $label.AutoEllipsis = $true
    $label.Dock         = [System.Windows.Forms.DockStyle]::Top
    $label.Height       = Get-AbmScaled $Height
    $label.TextAlign    = [System.Drawing.ContentAlignment]::MiddleLeft
    return $label
}

function Set-AbmGridStyle {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid)

    $p = $script:Palette
    $cellPadding = New-Object System.Windows.Forms.Padding((Get-AbmScaled 8), 0, (Get-AbmScaled 4), 0)

    $Grid.BackgroundColor = $p.Card
    $Grid.BorderStyle     = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor       = $p.Border
    $Grid.EnableHeadersVisualStyles   = $false
    $Grid.ColumnHeadersBorderStyle    = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $Grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $Grid.ColumnHeadersHeight         = Get-AbmScaled 38

    $header = $Grid.ColumnHeadersDefaultCellStyle
    $header.BackColor          = $p.Primary
    $header.ForeColor          = $p.HeaderText
    $header.SelectionBackColor = $p.Primary
    $header.SelectionForeColor = $p.HeaderText
    $header.Font               = $script:Fonts.GridHeader
    $header.Padding            = $cellPadding
    $header.Alignment          = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft

    $cells = $Grid.DefaultCellStyle
    $cells.Font               = $script:Fonts.Body
    $cells.BackColor          = $p.Card
    $cells.ForeColor          = $p.Text
    $cells.SelectionBackColor = $p.GridSelect
    $cells.SelectionForeColor = $p.Text
    $cells.Padding            = $cellPadding

    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $p.GridAlt
    $Grid.RowTemplate.Height      = Get-AbmScaled 30
    $Grid.RowHeadersVisible       = $false
    $Grid.AllowUserToAddRows      = $false
    $Grid.AllowUserToDeleteRows   = $false
    $Grid.AllowUserToResizeRows   = $false
    $Grid.ReadOnly                = $true
    $Grid.MultiSelect             = $true
    $Grid.SelectionMode           = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $Grid.StandardTab             = $true
}

Export-ModuleMember -Function Get-AbmColor, Get-AbmFont, Get-AbmScaled, New-AbmButton, New-AbmKpiCard, Set-AbmGridStyle
