Set-StrictMode -Version Latest

<#
    Sistema visual: paletas (Claro / Oscuro), tipografías y escala DPI.

    Los controles no guardan colores fijos: se "enlazan" a un rol de la paleta con Set-AbmThemeColor
    (Observer). Al cambiar de tema, Set-AbmTheme vuelve a aplicar cada rol y ejecuta los ganchos
    registrados, sin reconstruir la ventana.
#>

$script:Palettes = @{
    Claro = @{
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
        Input          = [System.Drawing.Color]::White
        Border         = [System.Drawing.Color]::FromArgb(221, 227, 236)
        Text           = [System.Drawing.Color]::FromArgb(30, 38, 51)
        Muted          = [System.Drawing.Color]::FromArgb(107, 118, 137)
        HeaderText     = [System.Drawing.Color]::White
        HeaderMuted    = [System.Drawing.Color]::FromArgb(168, 182, 206)
        GridAlt        = [System.Drawing.Color]::FromArgb(248, 250, 253)
        GridSelect     = [System.Drawing.Color]::FromArgb(220, 232, 255)
        SecondaryHover = [System.Drawing.Color]::FromArgb(234, 239, 247)
    }
    Oscuro = @{
        Primary        = [System.Drawing.Color]::FromArgb(14, 22, 38)
        PrimaryLight   = [System.Drawing.Color]::FromArgb(96, 128, 186)
        Accent         = [System.Drawing.Color]::FromArgb(77, 139, 255)
        AccentHover    = [System.Drawing.Color]::FromArgb(58, 118, 236)
        Gold           = [System.Drawing.Color]::FromArgb(214, 178, 62)
        Success        = [System.Drawing.Color]::FromArgb(52, 178, 123)
        SuccessHover   = [System.Drawing.Color]::FromArgb(40, 150, 103)
        Warning        = [System.Drawing.Color]::FromArgb(236, 176, 46)
        Danger         = [System.Drawing.Color]::FromArgb(240, 104, 104)
        DangerSoft     = [System.Drawing.Color]::FromArgb(74, 38, 42)
        Surface        = [System.Drawing.Color]::FromArgb(22, 25, 31)
        Card           = [System.Drawing.Color]::FromArgb(31, 35, 43)
        Input          = [System.Drawing.Color]::FromArgb(40, 45, 55)
        Border         = [System.Drawing.Color]::FromArgb(54, 60, 72)
        Text           = [System.Drawing.Color]::FromArgb(229, 233, 240)
        Muted          = [System.Drawing.Color]::FromArgb(148, 158, 175)
        HeaderText     = [System.Drawing.Color]::White
        HeaderMuted    = [System.Drawing.Color]::FromArgb(150, 166, 194)
        GridAlt        = [System.Drawing.Color]::FromArgb(36, 41, 50)
        GridSelect     = [System.Drawing.Color]::FromArgb(44, 72, 122)
        SecondaryHover = [System.Drawing.Color]::FromArgb(45, 51, 62)
    }
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

$script:ThemeName = 'Claro'
$script:Bindings  = New-Object 'System.Collections.Generic.Dictionary[int,System.Collections.Generic.List[object]]'
$script:Hooks     = New-Object System.Collections.Generic.List[scriptblock]
$script:SinceCleanup = 0

$script:Scale = 1.0
try {
    $graphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    $script:Scale = [math]::Max(1.0, $graphics.DpiX / 96.0)
    $graphics.Dispose()
}
catch { $script:Scale = 1.0 }

#region Consulta
function Get-AbmColor {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    $palette = $script:Palettes[$script:ThemeName]
    if (-not $palette.ContainsKey($Name)) { throw "Color no definido en la paleta: $Name" }
    return $palette[$Name]
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

function Get-AbmThemeName {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return $script:ThemeName
}

function Get-AbmThemeNames {
    [CmdletBinding()]
    param()
    return @($script:Palettes.Keys | Sort-Object)
}
#endregion

#region Enlaces de color y cambio de tema
function Set-AbmThemeColor {
    <#
        Asigna el color del rol a la propiedad y recuerda el enlace para cuando cambie el tema.
        -Owner: control dueño del objeto (p. ej. el botón de un FlatAppearance), para descartar el
        enlace cuando ese control se destruye.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Target,
        [Parameter(Mandatory)][string]$Property,
        [Parameter(Mandatory)][string]$Role,
        $Owner
    )

    $Target.$Property = Get-AbmColor $Role
    $binding = Find-Binding -Target $Target
    if ($null -eq $binding) {
        $key = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Target)
        if (-not $script:Bindings.ContainsKey($key)) { $script:Bindings[$key] = New-Object System.Collections.Generic.List[object] }
        $owner = if ($null -ne $Owner) { $Owner } else { $Target }
        $binding = [pscustomobject]@{ Target = $Target; Owner = $owner; Roles = @{} }
        $script:Bindings[$key].Add($binding)
    }
    $binding.Roles[$Property] = $Role

    $script:SinceCleanup++
    if ($script:SinceCleanup -ge 500) { Remove-DisposedBindings }
}

function Set-AbmTheme {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    if (-not $script:Palettes.ContainsKey($Name)) { throw "Tema no definido: $Name" }
    $script:ThemeName = $Name
    Remove-DisposedBindings
    # Se recorre una copia: aplicar un color puede disparar eventos que registren nuevos enlaces.
    foreach ($bucket in @($script:Bindings.Values)) {
        foreach ($binding in @($bucket)) {
            foreach ($property in @($binding.Roles.Keys)) { $binding.Target.$property = Get-AbmColor $binding.Roles[$property] }
        }
    }
    foreach ($hook in $script:Hooks) { & $hook }
}

function Register-AbmThemeHook {
    <# Acción a ejecutar después de cada cambio de tema (p. ej. redibujar o cambiar un texto). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Action)
    $script:Hooks.Add($Action)
}

function Find-Binding {
    param($Target)
    $key = [System.Runtime.CompilerServices.RuntimeHelpers]::GetHashCode($Target)
    if (-not $script:Bindings.ContainsKey($key)) { return $null }
    foreach ($binding in $script:Bindings[$key]) { if ([object]::ReferenceEquals($binding.Target, $Target)) { return $binding } }
    return $null
}

function Remove-DisposedBindings {
    foreach ($bucket in $script:Bindings.Values) {
        for ($i = $bucket.Count - 1; $i -ge 0; $i--) {
            $owner = $bucket[$i].Owner
            if ($null -ne $owner.PSObject.Properties['IsDisposed'] -and $owner.IsDisposed) { $bucket.RemoveAt($i) }
        }
    }
    $script:SinceCleanup = 0
}
#endregion

Export-ModuleMember -Function Get-AbmColor, Get-AbmFont, Get-AbmScaled, Get-AbmThemeName, Get-AbmThemeNames,
                              Set-AbmThemeColor, Set-AbmTheme, Register-AbmThemeHook
