Set-StrictMode -Version Latest

<#
    Mensajes y confirmaciones estándar (MessageBox) con un único título de aplicación.
#>

$script:Caption = 'ABM de Datos'

function Show-AbmMessage {
    [CmdletBinding()]
    param(
        $Owner,
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Information', 'Warning', 'Error')][string]$Icon = 'Information'
    )
    $buttons  = [System.Windows.Forms.MessageBoxButtons]::OK
    $iconType = [System.Windows.Forms.MessageBoxIcon]$Icon
    if ($Owner) { [void][System.Windows.Forms.MessageBox]::Show($Owner, $Message, $script:Caption, $buttons, $iconType) }
    else        { [void][System.Windows.Forms.MessageBox]::Show($Message, $script:Caption, $buttons, $iconType) }
}

function Show-AbmError {
    [CmdletBinding()]
    param($Owner, [Parameter(Mandatory)][string]$Message, $ErrorRecord)
    $detail = if ($ErrorRecord) { "`n`nDetalle: $($ErrorRecord.Exception.Message)" } else { '' }
    Show-AbmMessage -Owner $Owner -Message ($Message + $detail) -Icon Error
}

function Confirm-AbmAction {
    <# Sí / No con "No" como opción por defecto (para acciones destructivas). #>
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][string]$Message, [string]$Title = 'Confirmar')
    $result = [System.Windows.Forms.MessageBox]::Show($Owner, $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Request-AbmSaveDecision {
    <# Devuelve 'Save', 'Discard' o 'Cancel'. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)]$Owner, [Parameter(Mandatory)][string]$Message)
    $result = [System.Windows.Forms.MessageBox]::Show($Owner, $Message, 'Cambios pendientes',
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return 'Save' }
    if ($result -eq [System.Windows.Forms.DialogResult]::No)  { return 'Discard' }
    return 'Cancel'
}

Export-ModuleMember -Function Show-AbmMessage, Show-AbmError, Confirm-AbmAction, Request-AbmSaveDecision
