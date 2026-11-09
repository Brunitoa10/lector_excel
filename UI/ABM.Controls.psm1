Set-StrictMode -Version Latest

<#
    Fábrica de controles básicos con el estilo del tema (Factory). Todas las medidas se expresan
    en píxeles de diseño y se escalan según el DPI de la pantalla.
#>

$script:ButtonSchemes = @{
    Primary   = @{ Back = 'Accent';  Hover = 'AccentHover';    Fore = 'HeaderText'; Border = 'Accent' }
    Success   = @{ Back = 'Success'; Hover = 'SuccessHover';   Fore = 'HeaderText'; Border = 'Success' }
    Danger    = @{ Back = 'Card';    Hover = 'DangerSoft';     Fore = 'Danger';     Border = 'Danger' }
    Secondary = @{ Back = 'Card';    Hover = 'SecondaryHover'; Fore = 'Text';       Border = 'Border' }
}

function New-AbmPadding {
    [CmdletBinding()]
    param([int]$Left = 0, [int]$Top = 0, [int]$Right = 0, [int]$Bottom = 0)
    return New-Object System.Windows.Forms.Padding((Get-AbmScaled $Left), (Get-AbmScaled $Top), (Get-AbmScaled $Right), (Get-AbmScaled $Bottom))
}

function New-AbmPanel {
    [CmdletBinding()]
    param([System.Windows.Forms.DockStyle]$Dock = 'None', [int]$Height = 0, [int]$Width = 0, [string]$Color = '')

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = $Dock
    if ($Height -gt 0) { $panel.Height = Get-AbmScaled $Height }
    if ($Width -gt 0)  { $panel.Width  = Get-AbmScaled $Width }
    if ($Color)        { $panel.BackColor = Get-AbmColor $Color }
    return $panel
}

function New-AbmLabel {
    [CmdletBinding()]
    param([string]$Text = '', [string]$Font = 'Body', [string]$Color = 'Text', [int]$X = 0, [int]$Y = 0)

    $label = New-Object System.Windows.Forms.Label
    $label.Text      = $Text
    $label.Font      = Get-AbmFont $Font
    $label.ForeColor = Get-AbmColor $Color
    $label.AutoSize  = $true
    $label.Location  = New-Object System.Drawing.Point((Get-AbmScaled $X), (Get-AbmScaled $Y))
    return $label
}

function New-AbmButton {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Primary', 'Secondary', 'Success', 'Danger')][string]$Variant = 'Secondary',
        [int]$Width = 112
    )

    $scheme = $script:ButtonSchemes[$Variant]
    $button = New-Object System.Windows.Forms.Button
    $button.Text      = $Text
    $button.Width     = Get-AbmScaled $Width
    $button.Height    = Get-AbmScaled 34
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.Font      = Get-AbmFont 'Button'
    $button.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $button.Margin    = New-AbmPadding -Right 8
    $button.UseVisualStyleBackColor = $false
    $button.BackColor = Get-AbmColor $scheme.Back
    $button.ForeColor = Get-AbmColor $scheme.Fore
    $button.FlatAppearance.BorderSize         = 1
    $button.FlatAppearance.BorderColor        = Get-AbmColor $scheme.Border
    $button.FlatAppearance.MouseOverBackColor = Get-AbmColor $scheme.Hover
    $button.FlatAppearance.MouseDownBackColor = Get-AbmColor $scheme.Hover
    return $button
}

function New-AbmSeparator {
    <# Separador vertical para barras de herramientas. #>
    [CmdletBinding()]
    param([int]$Height = 34)
    $separator = New-Object System.Windows.Forms.Panel
    $separator.Width     = 1
    $separator.Height    = Get-AbmScaled $Height
    $separator.BackColor = Get-AbmColor 'Border'
    $separator.Margin    = New-AbmPadding -Left 6 -Right 14
    return $separator
}

Export-ModuleMember -Function New-AbmPadding, New-AbmPanel, New-AbmLabel, New-AbmButton, New-AbmSeparator
