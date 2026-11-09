Set-StrictMode -Version Latest

<#
    Pantalla principal.
      - Vista     : construcción de controles (New-*Area).
      - Controlador: acciones de usuario (Invoke-Abm*), que delegan en la fachada de datos (ABM.DataSource),
                     la auditoría (ABM.Logger) y los diálogos (ABM.Dialogs).
      - Estado    : un único objeto $script:State; la vista se refresca con Update-AbmView.
#>

$script:UI          = @{}
$script:State       = @{ Source = $null; Table = $null; View = $null; Dirty = $false }
$script:Config      = $null
$script:StartupPath = ''
$script:AppTitle    = 'ABM de Datos · Excel & TXT'

#region Punto de entrada
function Show-AbmMainForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [string]$InitialPath = ''
    )

    $script:Config      = $Config
    $script:StartupPath = $InitialPath

    $form = New-Object System.Windows.Forms.Form
    $form.Text          = $script:AppTitle
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Size          = New-Object System.Drawing.Size((Get-AbmScaled 1280), (Get-AbmScaled 800))
    $form.MinimumSize   = New-Object System.Drawing.Size((Get-AbmScaled 1040), (Get-AbmScaled 640))
    $form.BackColor     = Get-AbmColor 'Surface'
    $form.Font          = Get-AbmFont 'Body'
    $form.KeyPreview    = $true
    $form.AllowDrop     = $true
    $script:UI.Form     = $form

    # Orden de agregado = orden inverso de acoplamiento (el último queda en el borde).
    $form.Controls.AddRange([System.Windows.Forms.Control[]]@(
        (New-GridArea), (New-KpiArea), (New-ToolbarArea), (New-HeaderArea), (New-StatusArea)
    ))

    Register-AbmEvents
    Update-AbmView
    [void]$form.ShowDialog()
    $form.Dispose()
}
#endregion

#region Vista
function New-UiPanel {
    param([System.Windows.Forms.DockStyle]$Dock, [int]$Height = 0, [string]$Color = '')
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = $Dock
    if ($Height -gt 0) { $panel.Height = Get-AbmScaled $Height }
    if ($Color) { $panel.BackColor = Get-AbmColor $Color }
    return $panel
}

function New-UiLabel {
    param([string]$Text, [string]$Font, [string]$Color, [int]$X = 0, [int]$Y = 0)
    $label = New-Object System.Windows.Forms.Label
    $label.Text      = $Text
    $label.Font      = Get-AbmFont $Font
    $label.ForeColor = Get-AbmColor $Color
    $label.AutoSize  = $true
    $label.Location  = New-Object System.Drawing.Point((Get-AbmScaled $X), (Get-AbmScaled $Y))
    return $label
}

function New-UiPadding {
    param([int]$Left, [int]$Top, [int]$Right, [int]$Bottom)
    return New-Object System.Windows.Forms.Padding((Get-AbmScaled $Left), (Get-AbmScaled $Top), (Get-AbmScaled $Right), (Get-AbmScaled $Bottom))
}

function New-HeaderArea {
    $panel  = New-UiPanel -Dock Top -Height 78 -Color 'Primary'
    $accent = New-UiPanel -Dock Bottom -Height 3 -Color 'Gold'

    $file = New-Object System.Windows.Forms.Label
    $file.Dock         = [System.Windows.Forms.DockStyle]::Right
    $file.Width        = Get-AbmScaled 560
    $file.AutoSize     = $false
    $file.AutoEllipsis = $true
    $file.TextAlign    = [System.Drawing.ContentAlignment]::MiddleRight
    $file.Padding      = New-UiPadding 0 0 22 3
    $file.Font         = Get-AbmFont 'Label'
    $file.ForeColor    = Get-AbmColor 'HeaderMuted'
    $script:UI.HeaderFile = $file

    $title    = New-UiLabel -Text 'ABM de Datos' -Font 'Title' -Color 'HeaderText' -X 22 -Y 10
    $subtitle = New-UiLabel -Text 'Alta, baja y modificación de registros · Excel (.xlsx) y texto (.txt)' -Font 'Subtitle' -Color 'HeaderMuted' -X 24 -Y 47

    $panel.Controls.AddRange([System.Windows.Forms.Control[]]@($accent, $file, $title, $subtitle))
    return $panel
}

function New-ToolbarArea {
    $panel = New-UiPanel -Dock Top -Height 62 -Color 'Card'
    $line  = New-UiPanel -Dock Bottom -Height 1 -Color 'Border'

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock         = [System.Windows.Forms.DockStyle]::Fill
    $flow.WrapContents = $false
    $flow.Padding      = New-UiPadding 20 14 0 0
    $flow.BackColor    = Get-AbmColor 'Card'

    $ui = $script:UI
    $ui.BtnOpen   = New-AbmButton -Text 'Abrir…' -Width 96
    $ui.BtnNew    = New-AbmButton -Text 'Nuevo archivo' -Width 128
    $ui.BtnAdd    = New-AbmButton -Text '+  Alta' -Variant Primary -Width 96
    $ui.BtnEdit   = New-AbmButton -Text 'Modificar' -Width 106
    $ui.BtnDelete = New-AbmButton -Text 'Baja' -Variant Danger -Width 90
    $ui.BtnSave   = New-AbmButton -Text 'Guardar' -Variant Success -Width 104
    $ui.BtnSaveAs = New-AbmButton -Text 'Guardar como…' -Width 136

    $ui.Tips = New-Object System.Windows.Forms.ToolTip
    $ui.Tips.SetToolTip($ui.BtnOpen,   'Abrir un archivo Excel o TXT (Ctrl+O)')
    $ui.Tips.SetToolTip($ui.BtnNew,    'Crear un archivo nuevo definiendo sus columnas')
    $ui.Tips.SetToolTip($ui.BtnAdd,    'Agregar un registro (Ctrl+N)')
    $ui.Tips.SetToolTip($ui.BtnEdit,   'Modificar el registro seleccionado (F2 o doble clic)')
    $ui.Tips.SetToolTip($ui.BtnDelete, 'Dar de baja los registros seleccionados (Supr)')
    $ui.Tips.SetToolTip($ui.BtnSave,   'Guardar los cambios en el archivo (Ctrl+S)')
    $ui.Tips.SetToolTip($ui.BtnSaveAs, 'Guardar una copia en Excel o TXT (Ctrl+Shift+S)')

    $flow.Controls.AddRange([System.Windows.Forms.Control[]]@(
        $ui.BtnOpen, $ui.BtnNew, (New-ToolbarSeparator),
        $ui.BtnAdd, $ui.BtnEdit, $ui.BtnDelete, (New-ToolbarSeparator),
        $ui.BtnSave, $ui.BtnSaveAs
    ))
    $panel.Controls.AddRange([System.Windows.Forms.Control[]]@($flow, $line))
    return $panel
}

function New-ToolbarSeparator {
    $separator = New-Object System.Windows.Forms.Panel
    $separator.Width     = 1
    $separator.Height    = Get-AbmScaled 34
    $separator.BackColor = Get-AbmColor 'Border'
    $separator.Margin    = New-UiPadding 6 0 14 0
    return $separator
}

function New-KpiArea {
    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock        = [System.Windows.Forms.DockStyle]::Top
    $layout.Height      = Get-AbmScaled 108
    $layout.Padding     = New-UiPadding 20 14 6 8
    $layout.BackColor   = Get-AbmColor 'Surface'
    $layout.ColumnCount = 4
    $layout.RowCount    = 1
    for ($i = 0; $i -lt 4; $i++) {
        [void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    }
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $ui = $script:UI
    $ui.KpiRows   = New-AbmKpiCard -Title 'Registros' -AccentColor (Get-AbmColor 'Accent')
    $ui.KpiCols   = New-AbmKpiCard -Title 'Columnas'  -AccentColor (Get-AbmColor 'PrimaryLight')
    $ui.KpiSource = New-AbmKpiCard -Title 'Origen'    -AccentColor (Get-AbmColor 'Gold')
    $ui.KpiState  = New-AbmKpiCard -Title 'Estado'    -AccentColor (Get-AbmColor 'Success')

    $cards = @($ui.KpiRows, $ui.KpiCols, $ui.KpiSource, $ui.KpiState)
    for ($i = 0; $i -lt $cards.Count; $i++) { $layout.Controls.Add($cards[$i].Panel, $i, 0) }
    return $layout
}

function New-GridArea {
    $area = New-UiPanel -Dock Fill -Color 'Surface'
    $area.Padding = New-UiPadding 20 4 20 14

    $frame = New-UiPanel -Dock Fill -Color 'Border'
    $frame.Padding = New-Object System.Windows.Forms.Padding(1)
    $inner = New-UiPanel -Dock Fill -Color 'Card'

    # Barra superior de la tabla: selector de hoja + búsqueda
    $bar = New-UiPanel -Dock Top -Height 54 -Color 'Card'
    $bar.Width = Get-AbmScaled 1000
    $right = [System.Windows.Forms.AnchorStyles]'Top, Right'

    $sheetLabel = New-UiLabel -Text 'HOJA' -Font 'KpiTitle' -Color 'Muted' -X 16 -Y 20
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $combo.FlatStyle     = [System.Windows.Forms.FlatStyle]::Flat
    $combo.Width         = Get-AbmScaled 220
    $combo.Location      = New-Object System.Drawing.Point((Get-AbmScaled 62), (Get-AbmScaled 15))

    $search = New-Object System.Windows.Forms.TextBox
    $search.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $search.Width       = Get-AbmScaled 280
    $search.Location    = New-Object System.Drawing.Point(($bar.Width - (Get-AbmScaled 296)), (Get-AbmScaled 15))
    $search.Anchor      = $right
    $searchLabel = New-UiLabel -Text 'BUSCAR' -Font 'KpiTitle' -Color 'Muted' -X 0 -Y 20
    $searchLabel.Left   = $bar.Width - (Get-AbmScaled 356)
    $searchLabel.Anchor = $right
    $bar.Controls.AddRange([System.Windows.Forms.Control[]]@($sheetLabel, $combo, $searchLabel, $search))

    $separator = New-UiPanel -Dock Top -Height 1 -Color 'Border'

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $grid.AllowDrop = $true
    Set-AbmGridStyle -Grid $grid

    $empty = New-Object System.Windows.Forms.Label
    $empty.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $empty.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $empty.Font      = Get-AbmFont 'EmptyState'
    $empty.ForeColor = Get-AbmColor 'Muted'
    $empty.AllowDrop = $true
    $empty.Text      = "Abrí un archivo Excel o TXT para comenzar`r`n`r`nUsá «Abrir…», «Nuevo archivo» o arrastrá el archivo a esta ventana"

    $inner.Controls.AddRange([System.Windows.Forms.Control[]]@($grid, $empty, $separator, $bar))
    $frame.Controls.Add($inner)
    $area.Controls.Add($frame)

    $script:UI.Grid        = $grid
    $script:UI.EmptyLabel  = $empty
    $script:UI.SheetCombo  = $combo
    $script:UI.Search      = $search
    $script:UI.SearchTimer = New-Object System.Windows.Forms.Timer
    $script:UI.SearchTimer.Interval = 300
    return $area
}

function New-StatusArea {
    $strip = New-Object System.Windows.Forms.StatusStrip
    $strip.BackColor  = Get-AbmColor 'Card'
    $strip.SizingGrip = $false

    $status = New-Object System.Windows.Forms.ToolStripStatusLabel
    $status.Spring    = $true
    $status.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $status.Text      = 'Listo'

    $keys = New-Object System.Windows.Forms.ToolStripStatusLabel
    $keys.ForeColor = Get-AbmColor 'Muted'
    $keys.Text      = 'Ctrl+O Abrir · Ctrl+N Alta · F2 Modificar · Supr Baja · Ctrl+S Guardar · Ctrl+F Buscar · F5 Recargar'

    $user = New-Object System.Windows.Forms.ToolStripStatusLabel
    $user.ForeColor = Get-AbmColor 'Muted'
    $user.Text      = "   $env:USERNAME @ $env:COMPUTERNAME"

    [void]$strip.Items.Add($status)
    [void]$strip.Items.Add($keys)
    [void]$strip.Items.Add($user)
    $script:UI.StatusText = $status
    return $strip
}
#endregion

#region Eventos
function Register-AbmEvents {
    $ui = $script:UI

    $ui.BtnOpen.Add_Click({   Invoke-AbmSafely { Invoke-AbmOpen } })
    $ui.BtnNew.Add_Click({    Invoke-AbmSafely { Invoke-AbmNewFile } })
    $ui.BtnAdd.Add_Click({    Invoke-AbmSafely { Invoke-AbmAdd } })
    $ui.BtnEdit.Add_Click({   Invoke-AbmSafely { Invoke-AbmEdit } })
    $ui.BtnDelete.Add_Click({ Invoke-AbmSafely { Invoke-AbmDelete } })
    $ui.BtnSave.Add_Click({   Invoke-AbmSafely { [void](Invoke-AbmSave) } })
    $ui.BtnSaveAs.Add_Click({ Invoke-AbmSafely { [void](Invoke-AbmSaveAs) } })

    $ui.SheetCombo.Add_SelectionChangeCommitted({ Invoke-AbmSafely { Invoke-AbmSheetChange } })
    $ui.Grid.Add_SelectionChanged({ Update-AbmActions })
    $ui.Grid.Add_CellDoubleClick({
        param($s, $e)
        if ($e.RowIndex -ge 0) { Invoke-AbmSafely { Invoke-AbmEdit } }
    })

    $ui.Search.Add_TextChanged({ $script:UI.SearchTimer.Stop(); $script:UI.SearchTimer.Start() })
    $ui.SearchTimer.Add_Tick({ $script:UI.SearchTimer.Stop(); Invoke-AbmSafely { Set-AbmFilter } })

    $ui.Form.Add_KeyDown({
        param($s, $e)
        Invoke-AbmSafely -Action { param($keyEvent) Invoke-AbmShortcut -KeyEvent $keyEvent } -Argument $e
    })

    $dragEnter = {
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    }
    $dragDrop = {
        param($s, $e)
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        Invoke-AbmSafely -Action { param($dropped) Invoke-AbmDrop -Files $dropped } -Argument $files
    }
    foreach ($target in @($ui.Form, $ui.Grid, $ui.EmptyLabel)) {
        $target.Add_DragEnter($dragEnter)
        $target.Add_DragDrop($dragDrop)
    }

    $ui.Form.Add_Shown({
        if ($script:StartupPath) { Invoke-AbmSafely { Open-AbmDocument -Path $script:StartupPath } }
    })
    $ui.Form.Add_FormClosing({
        param($s, $e)
        try { if (-not (Confirm-AbmPendingChanges)) { $e.Cancel = $true } }
        catch { $e.Cancel = $false }
    })
}

function Invoke-AbmSafely {
    <# Barrera de errores para los eventos: ningún fallo cierra la aplicación. #>
    param([Parameter(Mandatory)][scriptblock]$Action, $Argument)
    try {
        if ($PSBoundParameters.ContainsKey('Argument')) { & $Action $Argument } else { & $Action }
    }
    catch {
        Set-AbmBusy -Busy $false
        Show-AbmError -Owner $script:UI.Form -Message 'Ocurrió un error inesperado.' -ErrorRecord $_
    }
}

function Invoke-AbmShortcut {
    param([System.Windows.Forms.KeyEventArgs]$KeyEvent)

    $key  = $KeyEvent.KeyCode
    $ctrl = $KeyEvent.Control
    $grid = $script:UI.Grid
    $handled = $true

    if     ($ctrl -and $KeyEvent.Shift -and $key -eq 'S') { [void](Invoke-AbmSaveAs) }
    elseif ($ctrl -and $key -eq 'S')                      { [void](Invoke-AbmSave) }
    elseif ($ctrl -and $key -eq 'O')                      { Invoke-AbmOpen }
    elseif ($ctrl -and $key -eq 'N')                      { Invoke-AbmAdd }
    elseif ($ctrl -and $key -eq 'F')                      { if ($script:UI.Search.Enabled) { [void]$script:UI.Search.Focus(); $script:UI.Search.SelectAll() } }
    elseif ($key -eq 'F2')                                { Invoke-AbmEdit }
    elseif ($key -eq 'F5')                                { Invoke-AbmReload }
    elseif ($key -eq 'Delete' -and $grid.Focused)         { Invoke-AbmDelete }
    elseif ($key -eq 'Escape' -and $script:UI.Search.Focused) { $script:UI.Search.Text = '' }
    else { $handled = $false }

    if ($handled) { $KeyEvent.Handled = $true; $KeyEvent.SuppressKeyPress = $true }
}
#endregion

#region Acciones (controlador)
function Invoke-AbmOpen {
    if (-not (Confirm-AbmPendingChanges)) { return }
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dialog.Title  = 'Abrir archivo'
        $dialog.Filter = Get-AbmFileFilter
        if ($dialog.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            Open-AbmDocument -Path $dialog.FileName
        }
    }
    finally { $dialog.Dispose() }
}

function Invoke-AbmDrop {
    param([string[]]$Files)
    if ($null -eq $Files -or $Files.Count -eq 0) { return }
    try { $null = Get-AbmProvider -Path $Files[0] }
    catch { Show-AbmMessage -Owner $script:UI.Form -Message $_.Exception.Message -Icon Warning; return }
    if (-not (Confirm-AbmPendingChanges)) { return }
    $script:UI.Form.Activate()
    Open-AbmDocument -Path $Files[0]
}

function Invoke-AbmReload {
    $source = $script:State.Source
    if ($null -eq $source) { return }
    if (-not (Confirm-AbmPendingChanges)) { return }
    Open-AbmDocument -Path $source.Path -Sheet $source.Meta.Sheet
}

function Invoke-AbmSheetChange {
    $source = $script:State.Source
    if ($null -eq $source -or $source.Kind -ne 'Excel') { return }
    $name = [string]$script:UI.SheetCombo.SelectedItem
    if ($name -eq $source.Meta.Sheet) { return }
    if (-not (Confirm-AbmPendingChanges)) { $script:UI.SheetCombo.SelectedItem = $source.Meta.Sheet; return }
    Open-AbmDocument -Path $source.Path -Sheet $name
}

function Open-AbmDocument {
    param([Parameter(Mandatory)][string]$Path, [string]$Sheet)
    try {
        Set-AbmBusy -Busy $true -Text ('Abriendo {0}…' -f [System.IO.Path]::GetFileName($Path))
        $source = Open-AbmSource -Path $Path -Sheet $Sheet
        Set-AbmSource -Source $source
        $count = $source.Table.Rows.Count
        Write-AbmAudit -Action 'ABRIR' -Target (Get-AbmSourceLabel) -Detail ('{0} registros' -f $count)
        Set-AbmStatus ('Archivo abierto: {0} · {1:N0} registros' -f (Get-AbmSourceLabel), $count)
    }
    catch {
        Show-AbmError -Owner $script:UI.Form -Message "No se pudo abrir el archivo.`n$Path" -ErrorRecord $_
    }
    finally { Set-AbmBusy -Busy $false }
}

function Invoke-AbmAdd {
    $table = $script:State.Table
    if ($null -eq $table) { return }

    $values = Show-AbmRecordDialog -Owner $script:UI.Form -Title 'Alta de registro' `
        -Subtitle ('Nuevo registro en {0}' -f (Get-AbmSourceLabel)) -Columns (Get-AbmColumnNames)
    if ($null -eq $values) { return }

    $row = $table.NewRow()
    foreach ($key in $values.Keys) { $row[$key] = $values[$key] }
    $table.Rows.Add($row)

    Write-AbmAudit -Action 'ALTA' -Target (Get-AbmSourceLabel) -Detail (Format-AbmRecord -Record $values)
    Set-AbmDirty
    Select-AbmRow -Row $row
    Set-AbmStatus 'Registro agregado. Recordá guardar los cambios (Ctrl+S).'
}

function Invoke-AbmEdit {
    $rows = @(Get-AbmSelectedRows)
    if ($rows.Count -eq 0) { return }
    if ($rows.Count -gt 1) {
        Show-AbmMessage -Owner $script:UI.Form -Message 'Seleccioná un único registro para modificar.' -Icon Information
        return
    }

    $row     = $rows[0]
    $current = ConvertTo-AbmRecord -Row $row
    $values  = Show-AbmRecordDialog -Owner $script:UI.Form -Title 'Modificar registro' `
        -Subtitle (Get-AbmSourceLabel) -Columns (Get-AbmColumnNames) -Values $current
    if ($null -eq $values) { return }

    $changes = New-Object System.Collections.Generic.List[string]
    foreach ($key in $values.Keys) {
        if ($current[$key] -cne $values[$key]) {
            $changes.Add(("{0}: '{1}' → '{2}'" -f $key, $current[$key], $values[$key]))
            $row[$key] = $values[$key]
        }
    }
    if ($changes.Count -eq 0) { Set-AbmStatus 'El registro no tuvo cambios.'; return }

    Write-AbmAudit -Action 'MODIFICACION' -Target (Get-AbmSourceLabel) -Detail ($changes -join ' | ')
    Set-AbmDirty
    Set-AbmStatus ('Registro modificado ({0} campo(s)). Recordá guardar los cambios.' -f $changes.Count)
}

function Invoke-AbmDelete {
    $rows = @(Get-AbmSelectedRows)
    if ($rows.Count -eq 0) { return }

    $question = if ($rows.Count -eq 1) { '¿Confirmás la baja del registro seleccionado?' }
                else { '¿Confirmás la baja de {0} registros?' -f $rows.Count }
    $message = "$question`n`nEl cambio se aplicará al archivo cuando guardes."
    if (-not (Confirm-AbmAction -Owner $script:UI.Form -Title 'Confirmar baja' -Message $message)) { return }

    foreach ($row in $rows) {
        $detail = Format-AbmRecord -Record (ConvertTo-AbmRecord -Row $row)
        $script:State.Table.Rows.Remove($row)
        Write-AbmAudit -Action 'BAJA' -Target (Get-AbmSourceLabel) -Detail $detail
    }
    Set-AbmDirty
    Set-AbmStatus ('{0} registro(s) dado(s) de baja. Recordá guardar los cambios.' -f $rows.Count)
}

function Invoke-AbmSave {
    [OutputType([bool])]
    param()

    $source = $script:State.Source
    if ($null -eq $source) { return $false }
    if (-not $script:State.Dirty) { Set-AbmStatus 'No hay cambios pendientes de guardar.'; return $true }

    try {
        Set-AbmBusy -Busy $true -Text 'Guardando cambios…'
        $backup = Backup-AbmFile -Path $source.Path
        Save-AbmSource -Source $source -Table $script:State.Table
        $script:State.Table.AcceptChanges()
        $script:State.Dirty = $false

        $detail = '{0} registros' -f $script:State.Table.Rows.Count
        if ($backup) { $detail += ' · respaldo: ' + [System.IO.Path]::GetFileName($backup) }
        Write-AbmAudit -Action 'GUARDAR' -Target (Get-AbmSourceLabel) -Detail $detail
        Update-AbmView
        Set-AbmStatus ('Cambios guardados en {0} a las {1:HH:mm:ss}' -f (Get-AbmSourceLabel), (Get-Date))
        return $true
    }
    catch {
        Show-AbmError -Owner $script:UI.Form -Message "No se pudo guardar el archivo.`nVerificá que no esté abierto en Excel u otro programa." -ErrorRecord $_
        return $false
    }
    finally { Set-AbmBusy -Busy $false }
}

function Invoke-AbmSaveAs {
    [OutputType([bool])]
    param()

    $source = $script:State.Source
    if ($null -eq $source) { return $false }

    $target = Select-AbmSavePath -Title 'Guardar como' `
        -Filter 'Libro de Excel (*.xlsx)|*.xlsx|Archivo de texto (*.txt)|*.txt' `
        -FilterIndex $(if ($source.Kind -eq 'Excel') { 1 } else { 2 }) `
        -FileName ([System.IO.Path]::GetFileNameWithoutExtension($source.Path) + '_copia') `
        -Directory ([System.IO.Path]::GetDirectoryName($source.Path))
    if (-not $target) { return $false }

    if ([string]::Equals([System.IO.Path]::GetFullPath($target), $source.Path, [System.StringComparison]::OrdinalIgnoreCase)) {
        $script:State.Dirty = $true
        return (Invoke-AbmSave)
    }

    $options = @{
        SheetName = if ($source.Kind -eq 'Excel') { $source.Meta.Sheet } else { 'Datos' }
        Delimiter = if ($source.Kind -eq 'Txt') { $source.Meta.Delimiter } else { [string]$script:Config.DelimitadorTxtPorDefecto }
    }
    try {
        Set-AbmBusy -Busy $true -Text 'Generando archivo…'
        $null = Backup-AbmFile -Path $target
        New-AbmSource -Path $target -Table $script:State.Table -Options $options
        Write-AbmAudit -Action 'GUARDAR_COMO' -Target ([System.IO.Path]::GetFileName($target)) -Detail ('Origen: {0}' -f (Get-AbmSourceLabel))
    }
    catch {
        Show-AbmError -Owner $script:UI.Form -Message 'No se pudo generar el archivo.' -ErrorRecord $_
        return $false
    }
    finally { Set-AbmBusy -Busy $false }

    $script:State.Dirty = $false
    Open-AbmDocument -Path $target
    return $true
}

function Invoke-AbmNewFile {
    if (-not (Confirm-AbmPendingChanges)) { return }

    $spec = Show-AbmNewFileDialog -Owner $script:UI.Form -DefaultDelimiter ([string]$script:Config.DelimitadorTxtPorDefecto)
    if ($null -eq $spec) { return }

    $isExcel = $spec.Format -eq 'Excel'
    $target = Select-AbmSavePath -Title 'Crear archivo' `
        -Filter $(if ($isExcel) { 'Libro de Excel (*.xlsx)|*.xlsx' } else { 'Archivo de texto (*.txt)|*.txt' }) `
        -FilterIndex 1 -FileName 'NuevoArchivo' -Directory ([Environment]::GetFolderPath('MyDocuments'))
    if (-not $target) { return }

    try {
        $null = Backup-AbmFile -Path $target
        $table = New-AbmDataTable -Columns $spec.Columns
        New-AbmSource -Path $target -Table $table -Options @{ SheetName = $spec.SheetName; Delimiter = $spec.Delimiter }
        Write-AbmAudit -Action 'CREAR' -Target ([System.IO.Path]::GetFileName($target)) -Detail ('Columnas: {0}' -f ($spec.Columns -join ', '))
    }
    catch {
        Show-AbmError -Owner $script:UI.Form -Message 'No se pudo crear el archivo.' -ErrorRecord $_
        return
    }
    Open-AbmDocument -Path $target
}

function Confirm-AbmPendingChanges {
    [OutputType([bool])]
    param()

    if (-not $script:State.Dirty) { return $true }
    $message  = "Hay cambios sin guardar en {0}.`n`n¿Querés guardarlos antes de continuar?" -f (Get-AbmSourceLabel)
    $decision = Request-AbmSaveDecision -Owner $script:UI.Form -Message $message
    switch ($decision) {
        'Save'    { return [bool](Invoke-AbmSave) }
        'Discard' { $script:State.Dirty = $false; return $true }
        default   { return $false }
    }
}

function Select-AbmSavePath {
    param([string]$Title, [string]$Filter, [int]$FilterIndex, [string]$FileName, [string]$Directory)

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    try {
        $dialog.Title            = $Title
        $dialog.Filter           = $Filter
        $dialog.FilterIndex      = $FilterIndex
        $dialog.FileName         = $FileName
        $dialog.InitialDirectory = $Directory
        $dialog.OverwritePrompt  = $true
        $dialog.AddExtension     = $true
        if ($dialog.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
        return ''
    }
    finally { $dialog.Dispose() }
}
#endregion

#region Estado y refresco de la vista
function Set-AbmSource {
    param([Parameter(Mandatory)]$Source)

    $script:State.Source = $Source
    $script:State.Table  = $Source.Table
    $script:State.View   = [System.Data.DataView]::new($Source.Table)
    $script:State.Dirty  = $false

    $ui = $script:UI
    $ui.SearchTimer.Stop()
    $ui.Search.Text = ''
    $ui.SearchTimer.Stop()
    $ui.Grid.DataSource = $script:State.View
    Set-AbmGridColumns
    Update-AbmSheetSelector
    Update-AbmView
}

function Set-AbmDirty {
    $script:State.Dirty = $true
    Update-AbmView
}

function Set-AbmFilter {
    $view = $script:State.View
    if ($null -eq $view) { return }

    $text = $script:UI.Search.Text.Trim()
    try {
        if ($text -eq '') { $view.RowFilter = '' }
        else {
            $literal = ConvertTo-AbmLikeLiteral -Text $text
            $clauses = foreach ($column in $script:State.Table.Columns) {
                "[{0}] LIKE '%{1}%'" -f ($column.ColumnName.Replace('\', '\\').Replace(']', '\]')), $literal
            }
            $view.RowFilter = ($clauses -join ' OR ')
        }
    }
    catch { Set-AbmStatus "No se pudo aplicar el filtro: $($_.Exception.Message)" }
    Update-AbmView
}

function ConvertTo-AbmLikeLiteral {
    param([string]$Text)
    $builder = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        switch ([string]$ch) {
            "'"     { [void]$builder.Append("''") }
            '['     { [void]$builder.Append('[[]') }
            ']'     { [void]$builder.Append('[]]') }
            '*'     { [void]$builder.Append('[*]') }
            '%'     { [void]$builder.Append('[%]') }
            default { [void]$builder.Append($ch) }
        }
    }
    return $builder.ToString()
}

function Set-AbmGridColumns {
    $grid = $script:UI.Grid
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    foreach ($column in $grid.Columns) {
        $column.SortMode     = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        $column.MinimumWidth = Get-AbmScaled 70
    }
    if ($grid.Columns.Count -le 6) {
        $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
        return
    }
    $grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
    $min = Get-AbmScaled 110
    $max = Get-AbmScaled 320
    foreach ($column in $grid.Columns) {
        if ($column.Width -gt $max) { $column.Width = $max } elseif ($column.Width -lt $min) { $column.Width = $min }
    }
}

function Update-AbmSheetSelector {
    $combo  = $script:UI.SheetCombo
    $source = $script:State.Source
    $combo.BeginUpdate()
    try {
        $combo.Items.Clear()
        if ($null -ne $source -and $source.Kind -eq 'Excel') {
            foreach ($name in @($source.Meta.Sheets)) { [void]$combo.Items.Add($name) }
            $combo.SelectedItem = $source.Meta.Sheet
            $combo.Enabled = $combo.Items.Count -gt 1
        }
        else {
            [void]$combo.Items.Add('— (archivo de texto)')
            $combo.SelectedIndex = 0
            $combo.Enabled = $false
        }
    }
    finally { $combo.EndUpdate() }
}

function Update-AbmView {
    $ui     = $script:UI
    $state  = $script:State
    $source = $state.Source
    $hasData = $null -ne $state.Table

    $ui.Grid.Visible       = $hasData
    $ui.EmptyLabel.Visible = -not $hasData

    if (-not $hasData) {
        foreach ($card in @($ui.KpiRows, $ui.KpiCols, $ui.KpiSource, $ui.KpiState)) { $card.Value.Text = '—'; $card.Caption.Text = '' }
        $ui.HeaderFile.Text = 'Sin archivo abierto'
        $ui.Form.Text = $script:AppTitle
        Update-AbmActions
        return
    }

    $total   = $state.Table.Rows.Count
    $visible = $state.View.Count
    $ui.KpiRows.Value.Text   = '{0:N0}' -f $total
    $ui.KpiRows.Caption.Text = if ($visible -ne $total) { 'Filtro activo: {0:N0} visibles' -f $visible } else { 'Registros en el archivo' }
    $ui.KpiCols.Value.Text   = '{0:N0}' -f $state.Table.Columns.Count
    $ui.KpiCols.Caption.Text = 'Campos por registro'

    if ($source.Kind -eq 'Excel') {
        $ui.KpiSource.Value.Text   = 'Excel'
        $ui.KpiSource.Caption.Text = 'Hoja «{0}» · {1} hoja(s) en el libro' -f $source.Meta.Sheet, @($source.Meta.Sheets).Count
    }
    else {
        $ui.KpiSource.Value.Text   = 'TXT'
        $ui.KpiSource.Caption.Text = '{0} · {1}' -f (Get-AbmDelimiterName -Delimiter $source.Meta.Delimiter), $source.Meta.EncodingName
    }

    $stateColor = if ($state.Dirty) { Get-AbmColor 'Warning' } else { Get-AbmColor 'Success' }
    $ui.KpiState.Strip.BackColor = $stateColor
    $ui.KpiState.Value.ForeColor = $stateColor
    $ui.KpiState.Value.Text      = if ($state.Dirty) { 'Pendiente' } else { 'Al día' }
    $ui.KpiState.Caption.Text    = if ($state.Dirty) { 'Hay cambios sin guardar' } else { 'Todos los cambios guardados' }

    $ui.HeaderFile.Text = '{0}{1}{2}' -f [System.IO.Path]::GetFileName($source.Path), [Environment]::NewLine, [System.IO.Path]::GetDirectoryName($source.Path)
    $ui.Form.Text = '{0}{1} — {2}' -f $(if ($state.Dirty) { '● ' } else { '' }), (Get-AbmSourceLabel), $script:AppTitle
    Update-AbmActions
}

function Update-AbmActions {
    $ui = $script:UI
    $hasData  = $null -ne $script:State.Table
    $selected = if ($hasData) { $ui.Grid.SelectedRows.Count } else { 0 }

    $ui.BtnAdd.Enabled    = $hasData
    $ui.BtnEdit.Enabled   = $selected -eq 1
    $ui.BtnDelete.Enabled = $selected -ge 1
    $ui.BtnSave.Enabled   = $hasData -and $script:State.Dirty
    $ui.BtnSaveAs.Enabled = $hasData
    $ui.Search.Enabled    = $hasData
}

function Set-AbmStatus {
    param([string]$Text)
    $script:UI.StatusText.Text = $Text
}

function Set-AbmBusy {
    param([bool]$Busy, [string]$Text = '')
    if ($Busy) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        if ($Text) { Set-AbmStatus $Text }
        [System.Windows.Forms.Application]::DoEvents()
    }
    else {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    }
}
#endregion

#region Helpers de datos
function Get-AbmSourceLabel {
    $source = $script:State.Source
    if ($null -eq $source) { return '' }
    $name = [System.IO.Path]::GetFileName($source.Path)
    if ($source.Kind -eq 'Excel') { return '{0} [{1}]' -f $name, $source.Meta.Sheet }
    return $name
}

function Get-AbmColumnNames {
    foreach ($column in $script:State.Table.Columns) { $column.ColumnName }
}

function Get-AbmSelectedRows {
    if ($null -eq $script:State.Table) { return }
    foreach ($gridRow in $script:UI.Grid.SelectedRows) {
        $item = $gridRow.DataBoundItem
        if ($item -is [System.Data.DataRowView]) { $item.Row }
    }
}

function ConvertTo-AbmRecord {
    param([Parameter(Mandatory)][System.Data.DataRow]$Row)
    $record = [ordered]@{}
    foreach ($column in $Row.Table.Columns) { $record[$column.ColumnName] = [string]$Row[$column] }
    return $record
}

function Format-AbmRecord {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Record, [int]$MaxFields = 6)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($key in $Record.Keys) {
        if ($parts.Count -ge $MaxFields) { $parts.Add('…'); break }
        $value = [string]$Record[$key]
        if ($value.Length -gt 40) { $value = $value.Substring(0, 37) + '...' }
        $parts.Add(('{0}={1}' -f $key, $value))
    }
    return ($parts -join '; ')
}

function Select-AbmRow {
    param([Parameter(Mandatory)][System.Data.DataRow]$Row)
    $grid = $script:UI.Grid
    for ($i = $grid.Rows.Count - 1; $i -ge 0; $i--) {
        $item = $grid.Rows[$i].DataBoundItem
        if ($item -is [System.Data.DataRowView] -and [object]::ReferenceEquals($item.Row, $Row)) {
            $grid.ClearSelection()
            $grid.CurrentCell = $grid.Rows[$i].Cells[0]
            $grid.Rows[$i].Selected = $true
            try { $grid.FirstDisplayedScrollingRowIndex = $i } catch { Write-Verbose 'Fila fuera del área visible.' }
            return
        }
    }
}
#endregion

Export-ModuleMember -Function Show-AbmMainForm
