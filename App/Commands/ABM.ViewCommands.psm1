Set-StrictMode -Version Latest

<#
    Comandos de presentación: tema Claro / Oscuro (recordado entre sesiones).
#>

$script:PreferenceKey = 'Tema'

function Initialize-AbmTheme {
    <# Aplica el tema guardado antes de construir la ventana y registra el refresco posterior. #>
    [CmdletBinding()]
    param()

    $name = [string](Get-AbmPreference -Name $script:PreferenceKey -Default 'Claro')
    if (@(Get-AbmThemeNames) -notcontains $name) { $name = 'Claro' }
    Set-AbmTheme -Name $name
    Register-AbmThemeHook { Update-ThemeChrome }
}

function Invoke-AbmToggleTheme {
    [CmdletBinding()]
    param()

    $next = if ((Get-AbmThemeName) -eq 'Oscuro') { 'Claro' } else { 'Oscuro' }
    Set-AbmTheme -Name $next
    Set-AbmPreference -Name $script:PreferenceKey -Value $next
    Set-AbmStatus ('Tema {0} aplicado.' -f $next.ToLowerInvariant())
}

function Get-AbmThemeToggleText {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if ((Get-AbmThemeName) -eq 'Oscuro') { return 'Modo claro' }
    return 'Modo oscuro'
}

function Update-ThemeChrome {
    $ui = Get-AbmUi
    if ($ui.ContainsKey('BtnTheme')) { $ui.BtnTheme.Text = Get-AbmThemeToggleText }
    if ($ui.ContainsKey('Form'))     { $ui.Form.Refresh() }
}

Export-ModuleMember -Function Initialize-AbmTheme, Invoke-AbmToggleTheme, Get-AbmThemeToggleText
