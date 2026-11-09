Set-StrictMode -Version Latest

<#
    Barra de acciones. Los botones se definen como datos (nombre, texto, estilo, ancho, ayuda);
    agregar uno nuevo es sumar una línea a la lista.
#>

$script:Buttons = @(
    @{ Name = 'BtnOpen';   Text = 'Abrir…';         Variant = 'Secondary'; Width = 96;  Tip = 'Abrir un archivo Excel o TXT (Ctrl+O)' }
    @{ Name = 'BtnNew';    Text = 'Nuevo archivo';  Variant = 'Secondary'; Width = 128; Tip = 'Crear un archivo nuevo definiendo sus columnas' }
    @{ Separator = $true }
    @{ Name = 'BtnAdd';    Text = '+  Alta';        Variant = 'Primary';   Width = 96;  Tip = 'Agregar un registro (Ctrl+N)' }
    @{ Name = 'BtnEdit';   Text = 'Modificar';      Variant = 'Secondary'; Width = 106; Tip = 'Modificar el registro seleccionado (F2 o doble clic)' }
    @{ Name = 'BtnDelete'; Text = 'Baja';           Variant = 'Danger';    Width = 90;  Tip = 'Dar de baja los registros seleccionados (Supr)' }
    @{ Separator = $true }
    @{ Name = 'BtnSave';   Text = 'Guardar';        Variant = 'Success';   Width = 104; Tip = 'Guardar los cambios en el archivo (Ctrl+S)' }
    @{ Name = 'BtnSaveAs'; Text = 'Guardar como…';  Variant = 'Secondary'; Width = 136; Tip = 'Guardar una copia en Excel o TXT (Ctrl+Shift+S)' }
)

function New-AbmToolbarView {
    [CmdletBinding()]
    param()

    $panel = New-AbmPanel -Dock Top -Height 62 -Color 'Card'
    $line  = New-AbmPanel -Dock Bottom -Color 'Border'
    $line.Height = 1

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock         = [System.Windows.Forms.DockStyle]::Fill
    $flow.WrapContents = $false
    $flow.Padding      = New-AbmPadding -Left 20 -Top 14
    $flow.BackColor    = Get-AbmColor 'Card'
    $flow.Controls.AddRange([System.Windows.Forms.Control[]]@(New-ToolbarItems))

    $panel.Controls.AddRange([System.Windows.Forms.Control[]]@($flow, $line))
    return $panel
}

function New-ToolbarItems {
    $ui = Get-AbmUi
    $ui.Tips = New-Object System.Windows.Forms.ToolTip
    foreach ($spec in $script:Buttons) {
        if ($spec.ContainsKey('Separator')) { New-AbmSeparator; continue }
        $button = New-AbmButton -Text $spec.Text -Variant $spec.Variant -Width $spec.Width
        $ui.Tips.SetToolTip($button, $spec.Tip)
        $ui[$spec.Name] = $button
        $button
    }
}

Export-ModuleMember -Function New-AbmToolbarView
