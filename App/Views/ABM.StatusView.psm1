Set-StrictMode -Version Latest

<#
    Barra de estado: mensaje de la última acción, barra de progreso, ayuda de atajos y usuario/equipo.
    También expone el modo "trabajando": cursor de espera, barra en movimiento y controles bloqueados.
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

    $progress = New-Object System.Windows.Forms.ToolStripProgressBar
    $progress.Style   = [System.Windows.Forms.ProgressBarStyle]::Marquee
    $progress.MarqueeAnimationSpeed = 30
    $progress.Width   = Get-AbmScaled 160
    $progress.Visible = $false

    [void]$strip.Items.Add($status)
    [void]$strip.Items.Add($progress)
    [void]$strip.Items.Add((New-StatusLabel -Text (Get-AbmShortcutHelp) -Color 'Muted'))
    [void]$strip.Items.Add((New-StatusLabel -Text "   $env:USERNAME @ $env:COMPUTERNAME" -Color 'Muted'))
    $ui = Get-AbmUi
    $ui.StatusText = $status
    $ui.Progress   = $progress
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

    $ui = Get-AbmUi
    (Get-AbmState).Busy  = $Busy
    $ui.Progress.Visible = $Busy
    $ui.Toolbar.Enabled  = -not $Busy
    $ui.DataArea.Enabled = -not $Busy
    $cursor = if ($Busy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    [System.Windows.Forms.Cursor]::Current = $cursor
    if ($Text) { Set-AbmStatus $Text }
    # Permite redibujar (barra en movimiento y contador) durante procesos largos.
    if ($Busy) { [System.Windows.Forms.Application]::DoEvents() }
}

function New-StatusLabel {
    param([string]$Text, [string]$Color)
    $label = New-Object System.Windows.Forms.ToolStripStatusLabel
    $label.Text      = $Text
    $label.ForeColor = Get-AbmColor $Color
    return $label
}

Export-ModuleMember -Function New-AbmStatusView, Set-AbmStatus, Set-AbmBusy
