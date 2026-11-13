Set-StrictMode -Version Latest

<#
    Ventana principal: compone las vistas, conecta los eventos y muestra la aplicación.
    No contiene lógica de negocio (está en App\Commands) ni de datos (está en Data y Providers).
#>

function Show-AbmMainForm {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Config, [string]$InitialPath = '')

    Initialize-AbmContext -Config $Config -StartupPath $InitialPath
    Initialize-AbmTheme
    $form = New-MainWindow
    $ui = Get-AbmUi
    $ui.Form = $form
    $ui.Tips = New-Object System.Windows.Forms.ToolTip
    New-AbmColumnMenu

    # Orden de agregado = orden inverso de acoplamiento (el último agregado queda en el borde).
    $form.Controls.AddRange([System.Windows.Forms.Control[]]@(
        (New-AbmGridView), (New-AbmKpiView), (New-AbmToolbarView), (New-AbmHeaderView), (New-AbmStatusView)
    ))

    Register-AbmEvents
    Update-AbmView
    try { [void]$form.ShowDialog() }
    finally { $form.Dispose() }
}

function New-MainWindow {
    $form = New-Object System.Windows.Forms.Form
    $form.Text          = Get-AbmSetting -Name AppTitle
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Size          = New-Object System.Drawing.Size((Get-AbmScaled 1280), (Get-AbmScaled 800))
    $form.MinimumSize   = New-Object System.Drawing.Size((Get-AbmScaled 1040), (Get-AbmScaled 640))
    Set-AbmThemeColor -Target $form -Property BackColor -Role 'Surface'
    $form.Font          = Get-AbmFont 'Body'
    $form.KeyPreview    = $true
    $form.AllowDrop     = $true
    return $form
}

Export-ModuleMember -Function Show-AbmMainForm
