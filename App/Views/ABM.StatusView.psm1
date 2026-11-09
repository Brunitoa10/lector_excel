Set-StrictMode -Version Latest

<#
    Barra de estado: mensaje de la última acción, ayuda de atajos y usuario/equipo.
    También expone el indicador de "trabajando" (cursor de espera).
#>

function New-AbmStatusView {
    [CmdletBinding()]
    param()

    $strip = New-Object System.Windows.Forms.StatusStrip
    $strip.BackColor  = Get-AbmColor 'Card'
    $strip.SizingGrip = $false

    $status = New-StatusLabel -Text 'Listo' -Color 'Text'
    $status.Spring    = $true
    $status.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft

    [void]$strip.Items.Add($status)
    [void]$strip.Items.Add((New-StatusLabel -Text (Get-AbmShortcutHelp) -Color 'Muted'))
    [void]$strip.Items.Add((New-StatusLabel -Text "   $env:USERNAME @ $env:COMPUTERNAME" -Color 'Muted'))
    (Get-AbmUi).StatusText = $status
    return $strip
}

function Set-AbmStatus {
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Text)
    (Get-AbmUi).StatusText.Text = $Text
}

function Set-AbmBusy {
    [CmdletBinding()]
    param([Parameter(Mandatory)][bool]$Busy, [string]$Text = '')

    if (-not $Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        return
    }
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    if ($Text) { Set-AbmStatus $Text }
    [System.Windows.Forms.Application]::DoEvents()
}

function New-StatusLabel {
    param([string]$Text, [string]$Color)
    $label = New-Object System.Windows.Forms.ToolStripStatusLabel
    $label.Text      = $Text
    $label.ForeColor = Get-AbmColor $Color
    return $label
}

Export-ModuleMember -Function New-AbmStatusView, Set-AbmStatus, Set-AbmBusy
