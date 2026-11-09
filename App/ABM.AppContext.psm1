Set-StrictMode -Version Latest

<#
    Contexto compartido de la aplicación: registro de controles de la UI, estado de la sesión
    y ajustes de arranque. Es el único lugar con estado mutable de la capa App.
#>

$script:Ui    = @{}
$script:State = @{ Source = $null; Table = $null; View = $null; Dirty = $false; Busy = $false }
$script:Settings = @{ Config = $null; StartupPath = ''; AppTitle = 'ABM de Datos · Excel & TXT' }

function Initialize-AbmContext {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Config, [string]$StartupPath = '')

    $script:Ui.Clear()
    $script:State.Source = $null
    $script:State.Table  = $null
    $script:State.View   = $null
    $script:State.Dirty  = $false
    $script:State.Busy   = $false
    $script:Settings.Config      = $Config
    $script:Settings.StartupPath = $StartupPath
}

function Get-AbmUi {
    <# Controles registrados por las vistas (Form, Grid, Search, botones, tarjetas…). #>
    [CmdletBinding()]
    param()
    return $script:Ui
}

function Get-AbmState {
    <# Estado de la sesión: Source (archivo abierto), Table, View (filtro/orden), Dirty y Busy. #>
    [CmdletBinding()]
    param()
    return $script:State
}

function Get-AbmSetting {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Config', 'StartupPath', 'AppTitle')][string]$Name)
    return $script:Settings[$Name]
}

Export-ModuleMember -Function Initialize-AbmContext, Get-AbmUi, Get-AbmState, Get-AbmSetting
