Set-StrictMode -Version Latest

<#
    Diálogos modales reutilizables. Todos comparten el mismo "shell" (encabezado, cuerpo, pie)
    para mantener una estética uniforme (Template Method sobre New-AbmDialogShell).
#>

$script:Dialog  = @{}
$script:Caption = 'ABM de Datos'

#region Mensajes
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
    [CmdletBinding()]
    [OutputType([bool])]
    param($Owner, [Parameter(Mandatory)][string]$Message, [string]$Title = 'Confirmar')
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
    param($Owner, [Parameter(Mandatory)][string]$Message)
    $result = [System.Windows.Forms.MessageBox]::Show($Owner, $Message, 'Cambios pendientes',
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { return 'Save' }
    if ($result -eq [System.Windows.Forms.DialogResult]::No)  { return 'Discard' }
    return 'Cancel'
}
#endregion

#region Shell común
function New-AbmDialogShell {
    param([string]$Title, [string]$Subtitle, [int]$Width = 560, [int]$BodyHeight = 300)

    $width  = Get-AbmScaled $Width
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = $Title
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox     = $false
    $form.MinimizeBox     = $false
    $form.ShowInTaskbar   = $false
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.BackColor       = Get-AbmColor 'Card'
    $form.Font            = Get-AbmFont 'Body'
    $form.ClientSize      = New-Object System.Drawing.Size($width, (Get-AbmScaled (74 + 64 + $BodyHeight)))

    # Encabezado
    $header = New-Object System.Windows.Forms.Panel
    $header.Dock      = [System.Windows.Forms.DockStyle]::Top
    $header.Height    = Get-AbmScaled 74
    $header.BackColor = Get-AbmColor 'Primary'
    $accent = New-Object System.Windows.Forms.Panel
    $accent.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $accent.Height    = Get-AbmScaled 3
    $accent.BackColor = Get-AbmColor 'Gold'
    $titleLabel = New-AbmDialogLabel -Text $Title -Font (Get-AbmFont 'DialogTitle') -Color (Get-AbmColor 'HeaderText') -X 22 -Y 12
    $subLabel   = New-AbmDialogLabel -Text $Subtitle -Font (Get-AbmFont 'Subtitle') -Color (Get-AbmColor 'HeaderMuted') -X 24 -Y 44
    $subLabel.AutoSize     = $false
    $subLabel.AutoEllipsis = $true
    $subLabel.Size         = New-Object System.Drawing.Size(($width - (Get-AbmScaled 48)), (Get-AbmScaled 20))
    $header.Controls.AddRange([System.Windows.Forms.Control[]]@($accent, $titleLabel, $subLabel))

    # Pie
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $footer.Height    = Get-AbmScaled 64
    $footer.BackColor = Get-AbmColor 'Surface'
    $line = New-Object System.Windows.Forms.Panel
    $line.Dock      = [System.Windows.Forms.DockStyle]::Top
    $line.Height    = 1
    $line.BackColor = Get-AbmColor 'Border'

    $ok     = New-AbmButton -Text 'Aceptar' -Variant Primary -Width 120
    $cancel = New-AbmButton -Text 'Cancelar' -Width 110
    $top    = Get-AbmScaled 15
    $ok.Location     = New-Object System.Drawing.Point(($width - (Get-AbmScaled 24) - $ok.Width), $top)
    $cancel.Location = New-Object System.Drawing.Point(($ok.Left - (Get-AbmScaled 10) - $cancel.Width), $top)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $footer.Controls.AddRange([System.Windows.Forms.Control[]]@($line, $ok, $cancel))

    # Cuerpo (apila etiqueta + control, con scroll si hay muchos campos)
    $body = New-Object System.Windows.Forms.FlowLayoutPanel
    $body.Dock          = [System.Windows.Forms.DockStyle]::Fill
    $body.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $body.WrapContents  = $false
    $body.AutoScroll    = $true
    $body.BackColor     = Get-AbmColor 'Card'
    $body.Padding       = New-Object System.Windows.Forms.Padding((Get-AbmScaled 24), (Get-AbmScaled 12), (Get-AbmScaled 12), (Get-AbmScaled 8))

    $form.Controls.AddRange([System.Windows.Forms.Control[]]@($body, $footer, $header))
    $form.AcceptButton = $ok
    $form.CancelButton = $cancel

    return [pscustomobject]@{ Form = $form; Body = $body; Ok = $ok; FieldWidth = $width - (Get-AbmScaled 60) }
}

function New-AbmDialogLabel {
    param([string]$Text, [System.Drawing.Font]$Font, [System.Drawing.Color]$Color, [int]$X, [int]$Y)
    $label = New-Object System.Windows.Forms.Label
    $label.Text      = $Text
    $label.Font      = $Font
    $label.ForeColor = $Color
    $label.AutoSize  = $true
    $label.Location  = New-Object System.Drawing.Point((Get-AbmScaled $X), (Get-AbmScaled $Y))
    return $label
}

function Add-AbmDialogField {
    param($Shell, [string]$Caption, [System.Windows.Forms.Control]$Control)

    $label = New-Object System.Windows.Forms.Label
    $label.Text      = $Caption
    $label.AutoSize  = $true
    $label.Font      = Get-AbmFont 'Label'
    $label.ForeColor = Get-AbmColor 'Muted'
    $label.Margin    = New-Object System.Windows.Forms.Padding(0, (Get-AbmScaled 6), 0, (Get-AbmScaled 3))

    $Control.Width  = $Shell.FieldWidth
    $Control.Font   = Get-AbmFont 'Body'
    $Control.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, (Get-AbmScaled 6))

    $Shell.Body.Controls.Add($label)
    $Shell.Body.Controls.Add($Control)
}

function Invoke-AbmDialog {
    param($Shell, $Owner)
    if ($Owner) { return $Shell.Form.ShowDialog($Owner) }
    return $Shell.Form.ShowDialog()
}
#endregion

#region Alta / Modificación
function Show-AbmRecordDialog {
    <# Devuelve un diccionario ordenado columna -> valor, o $null si se canceló. #>
    [CmdletBinding()]
    param(
        $Owner,
        [Parameter(Mandatory)][string]$Title,
        [string]$Subtitle = '',
        [Parameter(Mandatory)][string[]]$Columns,
        [System.Collections.IDictionary]$Values
    )

    $bodyHeight = [math]::Min([math]::Max($Columns.Count * 58 + 24, 140), 480)
    $shell = New-AbmDialogShell -Title $Title -Subtitle $Subtitle -BodyHeight $bodyHeight
    $boxes = [ordered]@{}
    foreach ($column in $Columns) {
        $box = New-Object System.Windows.Forms.TextBox
        $box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        if ($null -ne $Values -and $Values.Contains($column)) { $box.Text = [string]$Values[$column] }
        Add-AbmDialogField -Shell $shell -Caption $column -Control $box
        $boxes[$column] = $box
    }

    $script:Dialog = @{ Form = $shell.Form; Boxes = $boxes }
    $shell.Ok.Add_Click({
        $hasValue = $false
        foreach ($box in $script:Dialog.Boxes.Values) { if ($box.Text.Trim() -ne '') { $hasValue = $true; break } }
        if (-not $hasValue) {
            Show-AbmMessage -Owner $script:Dialog.Form -Message 'Completá al menos un campo para registrar.' -Icon Warning
            return
        }
        $script:Dialog.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $shell.Form.Add_Shown({
        $first = @($script:Dialog.Boxes.Values) | Select-Object -First 1
        if ($null -ne $first) { [void]$first.Focus(); $first.SelectAll() }
    })

    try {
        if ((Invoke-AbmDialog -Shell $shell -Owner $Owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        $result = [ordered]@{}
        foreach ($key in $boxes.Keys) { $result[$key] = $boxes[$key].Text }
        return $result
    }
    finally {
        $shell.Form.Dispose()
        $script:Dialog = @{}
    }
}
#endregion

#region Nuevo archivo
function Show-AbmNewFileDialog {
    <# Devuelve @{ Format; SheetName; Delimiter; Columns } o $null si se canceló. #>
    [CmdletBinding()]
    param($Owner, [string]$DefaultDelimiter = ';')

    $delimiters = @(
        @{ Text = 'Punto y coma  ( ; )';  Value = ';' }
        @{ Text = 'Tabulación';           Value = "`t" }
        @{ Text = 'Barra vertical  ( | )'; Value = '|' }
        @{ Text = 'Coma  ( , )';          Value = ',' }
    )

    $shell = New-AbmDialogShell -Title 'Nuevo archivo' -Subtitle 'Definí el formato y las columnas del archivo a crear' -BodyHeight 380

    $format = New-Object System.Windows.Forms.ComboBox
    $format.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$format.Items.AddRange([object[]]@('Libro de Excel (.xlsx)', 'Archivo de texto (.txt)'))
    $format.SelectedIndex = 0

    $sheet = New-Object System.Windows.Forms.TextBox
    $sheet.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $sheet.MaxLength   = 31
    $sheet.Text        = 'Datos'

    $delimiter = New-Object System.Windows.Forms.ComboBox
    $delimiter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($item in $delimiters) { [void]$delimiter.Items.Add($item.Text) }
    $delimiter.SelectedIndex = 0
    for ($i = 0; $i -lt $delimiters.Count; $i++) { if ($delimiters[$i].Value -eq $DefaultDelimiter) { $delimiter.SelectedIndex = $i } }
    $delimiter.Enabled = $false

    $columns = New-Object System.Windows.Forms.TextBox
    $columns.BorderStyle   = [System.Windows.Forms.BorderStyle]::FixedSingle
    $columns.Multiline     = $true
    $columns.AcceptsReturn = $true
    $columns.ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
    $columns.Height        = Get-AbmScaled 150

    Add-AbmDialogField -Shell $shell -Caption 'Formato' -Control $format
    Add-AbmDialogField -Shell $shell -Caption 'Nombre de la hoja (Excel)' -Control $sheet
    Add-AbmDialogField -Shell $shell -Caption 'Delimitador (TXT)' -Control $delimiter
    Add-AbmDialogField -Shell $shell -Caption 'Columnas — una por línea' -Control $columns

    $script:Dialog = @{ Form = $shell.Form; Format = $format; Sheet = $sheet; Delimiter = $delimiter; Columns = $columns }
    $format.Add_SelectedIndexChanged({
        $isExcel = $script:Dialog.Format.SelectedIndex -eq 0
        $script:Dialog.Sheet.Enabled     = $isExcel
        $script:Dialog.Delimiter.Enabled = -not $isExcel
    })
    $shell.Ok.Add_Click({
        $names = @(Get-AbmColumnList -Lines $script:Dialog.Columns.Lines)
        if ($names.Count -eq 0) {
            Show-AbmMessage -Owner $script:Dialog.Form -Message 'Ingresá al menos una columna (una por línea).' -Icon Warning
            return
        }
        $duplicates = @($names | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
        if ($duplicates.Count -gt 0) {
            Show-AbmMessage -Owner $script:Dialog.Form -Message ("Hay columnas repetidas: {0}" -f ($duplicates -join ', ')) -Icon Warning
            return
        }
        $script:Dialog.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    $shell.Form.Add_Shown({ [void]$script:Dialog.Columns.Focus() })

    try {
        if ((Invoke-AbmDialog -Shell $shell -Owner $Owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        return @{
            Format    = if ($format.SelectedIndex -eq 0) { 'Excel' } else { 'Txt' }
            SheetName = $sheet.Text.Trim()
            Delimiter = $delimiters[$delimiter.SelectedIndex].Value
            Columns   = [string[]]@(Get-AbmColumnList -Lines $columns.Lines)
        }
    }
    finally {
        $shell.Form.Dispose()
        $script:Dialog = @{}
    }
}

function Get-AbmColumnList {
    param([string[]]$Lines)
    foreach ($line in $Lines) {
        $name = $line.Trim()
        if ($name -ne '') { $name }
    }
}
#endregion

Export-ModuleMember -Function Show-AbmMessage, Show-AbmError, Confirm-AbmAction, Request-AbmSaveDecision,
                              Show-AbmRecordDialog, Show-AbmNewFileDialog
