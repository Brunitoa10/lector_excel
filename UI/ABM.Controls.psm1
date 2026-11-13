Set-StrictMode -Version Latest

<#
    Fábrica de controles básicos con el estilo del tema (Factory). Todos los colores se enlazan a roles
    de la paleta, así que responden al cambio Claro / Oscuro. Las medidas están en píxeles de diseño.
#>

$script:ButtonSchemes = @{
    Primary   = @{ Back = 'Accent';  Hover = 'AccentHover';    Fore = 'HeaderText'; Border = 'Accent' }
    Success   = @{ Back = 'Success'; Hover = 'SuccessHover';   Fore = 'HeaderText'; Border = 'Success' }
    Danger    = @{ Back = 'Card';    Hover = 'DangerSoft';     Fore = 'Danger';     Border = 'Danger' }
    Secondary = @{ Back = 'Card';    Hover = 'SecondaryHover'; Fore = 'Text';       Border = 'Border' }
    Header    = @{ Back = 'Primary'; Hover = 'PrimaryLight';   Fore = 'HeaderText'; Border = 'HeaderMuted' }
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
    if ($Color)        { Set-AbmThemeColor -Target $panel -Property BackColor -Role $Color }
    return $panel
}

function New-AbmLabel {
    [CmdletBinding()]
    param([string]$Text = '', [string]$Font = 'Body', [string]$Color = 'Text', [int]$X = 0, [int]$Y = 0)

    $label = New-Object System.Windows.Forms.Label
    $label.Text      = $Text
    $label.Font      = Get-AbmFont $Font
    $label.AutoSize  = $true
    $label.BackColor = [System.Drawing.Color]::Transparent
    $label.Location  = New-Object System.Drawing.Point((Get-AbmScaled $X), (Get-AbmScaled $Y))
    Set-AbmThemeColor -Target $label -Property ForeColor -Role $Color
    return $label
}

function New-AbmButton {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Primary', 'Secondary', 'Success', 'Danger', 'Header')][string]$Variant = 'Secondary',
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
    $button.UseVisualStyleBackColor   = $false
    $button.FlatAppearance.BorderSize = 1

    Set-AbmThemeColor -Target $button -Property BackColor -Role $scheme.Back
    Set-AbmThemeColor -Target $button -Property ForeColor -Role $scheme.Fore
    $appearance = $button.FlatAppearance
    Set-AbmThemeColor -Target $appearance -Property BorderColor -Role $scheme.Border -Owner $button
    Set-AbmThemeColor -Target $appearance -Property MouseOverBackColor -Role $scheme.Hover -Owner $button
    Set-AbmThemeColor -Target $appearance -Property MouseDownBackColor -Role $scheme.Hover -Owner $button
    return $button
}

function New-AbmTextBox {
    [CmdletBinding()]
    param([string]$Text = '', [int]$MaxLength = 32767, [switch]$Multiline)

    $box = New-Object System.Windows.Forms.TextBox
    $box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $box.MaxLength   = $MaxLength
    $box.Text        = $Text
    if ($Multiline) {
        $box.Multiline     = $true
        $box.AcceptsReturn = $true
        $box.ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
    }
    Set-InputColors -Control $box
    return $box
}

function New-AbmComboBox {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Items = @(), [int]$Selected = -1)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $combo.FlatStyle     = [System.Windows.Forms.FlatStyle]::Flat
    if ($Items.Count -gt 0) { [void]$combo.Items.AddRange([object[]]$Items) }
    if ($Selected -ge 0 -and $Selected -lt $Items.Count) { $combo.SelectedIndex = $Selected }
    Set-InputColors -Control $combo
    return $combo
}

function New-AbmSeparator {
    <# Separador vertical para barras de herramientas. #>
    [CmdletBinding()]
    param([int]$Height = 34)
    $separator = New-Object System.Windows.Forms.Panel
    $separator.Width  = 1
    $separator.Height = Get-AbmScaled $Height
    $separator.Margin = New-AbmPadding -Left 6 -Right 14
    Set-AbmThemeColor -Target $separator -Property BackColor -Role 'Border'
    return $separator
}

function Set-InputColors {
    param([System.Windows.Forms.Control]$Control)
    Set-AbmThemeColor -Target $Control -Property BackColor -Role 'Input'
    Set-AbmThemeColor -Target $Control -Property ForeColor -Role 'Text'
}

Export-ModuleMember -Function New-AbmPadding, New-AbmPanel, New-AbmLabel, New-AbmButton, New-AbmTextBox, New-AbmComboBox, New-AbmSeparator
