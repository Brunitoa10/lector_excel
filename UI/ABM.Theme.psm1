Set-StrictMode -Version Latest

<#
    Sistema visual: paleta, tipografías y escala DPI. Única fuente de verdad del estilo (sin "colores mágicos").
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
    <# Convierte píxeles de diseño (96 DPI) a píxeles reales de la pantalla. #>
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][int]$Value)
    return [int][math]::Round($Value * $script:Scale)
}

Export-ModuleMember -Function Get-AbmColor, Get-AbmFont, Get-AbmScaled
