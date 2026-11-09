Set-StrictMode -Version Latest

<#
    Armazón común de todos los diálogos (Template Method): encabezado, cuerpo con scroll y pie con
    Aceptar / Cancelar. Cada diálogo concreto sólo agrega sus campos y su validación.
#>

function New-AbmDialogShell {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Title, [string]$Subtitle = '', [int]$Width = 560, [int]$BodyHeight = 300)

    $form = New-DialogForm -Title $Title -Width $Width -Height (74 + 64 + $BodyHeight)
    $clientWidth = $form.ClientSize.Width
    $body   = New-DialogBody
    $footer = New-DialogFooter -Width $clientWidth
    $header = New-DialogHeader -Title $Title -Subtitle $Subtitle -Width $clientWidth
    $form.Controls.AddRange([System.Windows.Forms.Control[]]@($body, $footer.Panel, $header))
    $form.AcceptButton = $footer.Ok
    $form.CancelButton = $footer.Cancel

    return [pscustomobject]@{ Form = $form; Body = $body; Ok = $footer.Ok; FieldWidth = $clientWidth - (Get-AbmScaled 60) }
}

function Add-AbmDialogField {
    <# Agrega una etiqueta y su control al cuerpo del diálogo. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Shell, [Parameter(Mandatory)][string]$Caption, [Parameter(Mandatory)][System.Windows.Forms.Control]$Control)

    $label = New-AbmLabel -Text $Caption -Font 'Label' -Color 'Muted'
    $label.Margin = New-AbmPadding -Top 6 -Bottom 3

    $Control.Width  = $Shell.FieldWidth
    $Control.Font   = Get-AbmFont 'Body'
    $Control.Margin = New-AbmPadding -Bottom 6

    $Shell.Body.Controls.Add($label)
    $Shell.Body.Controls.Add($Control)
}

function Invoke-AbmDialog {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Shell, $Owner)
    if ($Owner) { return $Shell.Form.ShowDialog($Owner) }
    return $Shell.Form.ShowDialog()
}

#region Partes del armazón
function New-DialogForm {
    param([string]$Title, [int]$Width, [int]$Height)
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = $Title
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.ShowInTaskbar   = $false
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.BackColor       = Get-AbmColor 'Card'
    $form.Font            = Get-AbmFont 'Body'
    $form.ClientSize      = New-Object System.Drawing.Size((Get-AbmScaled $Width), (Get-AbmScaled $Height))
    return $form
}

function New-DialogHeader {
    param([string]$Title, [string]$Subtitle, [int]$Width)
    $header = New-AbmPanel -Dock Top -Height 74 -Color 'Primary'
    $accent = New-AbmPanel -Dock Bottom -Height 3 -Color 'Gold'
    $titleLabel = New-AbmLabel -Text $Title -Font 'DialogTitle' -Color 'HeaderText' -X 22 -Y 12
    $subLabel   = New-AbmLabel -Text $Subtitle -Font 'Subtitle' -Color 'HeaderMuted' -X 24 -Y 44
    $subLabel.AutoSize     = $false
    $subLabel.AutoEllipsis = $true
    $subLabel.Size         = New-Object System.Drawing.Size(($Width - (Get-AbmScaled 48)), (Get-AbmScaled 20))
    $header.Controls.AddRange([System.Windows.Forms.Control[]]@($accent, $titleLabel, $subLabel))
    return $header
}

function New-DialogFooter {
    param([int]$Width)
    $footer = New-AbmPanel -Dock Bottom -Height 64 -Color 'Surface'
    $line   = New-AbmPanel -Dock Top -Color 'Border'
    $line.Height = 1

    $ok     = New-AbmButton -Text 'Aceptar' -Variant Primary -Width 120
    $cancel = New-AbmButton -Text 'Cancelar' -Width 110
    $top    = Get-AbmScaled 15
    $ok.Location     = New-Object System.Drawing.Point(($Width - (Get-AbmScaled 24) - $ok.Width), $top)
    $cancel.Location = New-Object System.Drawing.Point(($ok.Left - (Get-AbmScaled 10) - $cancel.Width), $top)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $footer.Controls.AddRange([System.Windows.Forms.Control[]]@($line, $ok, $cancel))
    return [pscustomobject]@{ Panel = $footer; Ok = $ok; Cancel = $cancel }
}

function New-DialogBody {
    $body = New-Object System.Windows.Forms.FlowLayoutPanel
    $body.Dock          = [System.Windows.Forms.DockStyle]::Fill
    $body.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $body.WrapContents  = $false
    $body.AutoScroll    = $true
    $body.BackColor     = Get-AbmColor 'Card'
    $body.Padding       = New-AbmPadding -Left 24 -Top 12 -Right 12 -Bottom 8
    return $body
}
#endregion

Export-ModuleMember -Function New-AbmDialogShell, Add-AbmDialogField, Invoke-AbmDialog
