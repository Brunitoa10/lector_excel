Set-StrictMode -Version Latest

<#
    Diálogo "Nuevo archivo": formato, nombre de hoja (Excel), delimitador (TXT) y columnas.
    Devuelve @{ Format; SheetName; Delimiter; Columns } o $null si se canceló.
#>

$script:Dialog = @{}
$script:Formats    = @('Libro de Excel (.xlsx)', 'Archivo de texto (.txt)')
$script:Delimiters = @(
    @{ Text = 'Punto y coma  ( ; )';   Value = ';' }
    @{ Text = 'Tabulación';            Value = "`t" }
    @{ Text = 'Barra vertical  ( | )'; Value = '|' }
    @{ Text = 'Coma  ( , )';           Value = ',' }
)

function Show-AbmNewFileDialog {
    [CmdletBinding()]
    param($Owner, [string]$DefaultDelimiter = ';')

    $shell = New-AbmDialogShell -Title 'Nuevo archivo' -Subtitle 'Definí el formato y las columnas del archivo a crear' -BodyHeight 380
    $fields = @{
        Format    = New-ComboField -Items $script:Formats -Selected 0
        Sheet     = New-TextField -Text 'Datos' -MaxLength 31
        Delimiter = New-ComboField -Items @($script:Delimiters | ForEach-Object { $_.Text }) -Selected (Get-DelimiterIndex $DefaultDelimiter)
        Columns   = New-ColumnsField
    }
    $fields.Delimiter.Enabled = $false

    Add-AbmDialogField -Shell $shell -Caption 'Formato' -Control $fields.Format
    Add-AbmDialogField -Shell $shell -Caption 'Nombre de la hoja (Excel)' -Control $fields.Sheet
    Add-AbmDialogField -Shell $shell -Caption 'Delimitador (TXT)' -Control $fields.Delimiter
    Add-AbmDialogField -Shell $shell -Caption 'Columnas — una por línea' -Control $fields.Columns

    $script:Dialog = $fields + @{ Form = $shell.Form }
    $fields.Format.Add_SelectedIndexChanged({ Update-FormatFields })
    $shell.Ok.Add_Click({ Confirm-NewFileDialog })
    $shell.Form.Add_Shown({ [void]$script:Dialog.Columns.Focus() })

    try {
        if ((Invoke-AbmDialog -Shell $shell -Owner $Owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        return @{
            Format    = if ($fields.Format.SelectedIndex -eq 0) { 'Excel' } else { 'Txt' }
            SheetName = $fields.Sheet.Text.Trim()
            Delimiter = $script:Delimiters[$fields.Delimiter.SelectedIndex].Value
            Columns   = [string[]]@(Get-ColumnNames -Lines $fields.Columns.Lines)
        }
    }
    finally {
        $shell.Form.Dispose()
        $script:Dialog = @{}
    }
}

#region Campos
function New-ComboField {
    param([string[]]$Items, [int]$Selected)
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$combo.Items.AddRange([object[]]$Items)
    $combo.SelectedIndex = $Selected
    return $combo
}

function New-TextField {
    param([string]$Text, [int]$MaxLength = 32767)
    $box = New-Object System.Windows.Forms.TextBox
    $box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $box.MaxLength   = $MaxLength
    $box.Text        = $Text
    return $box
}

function New-ColumnsField {
    $box = New-TextField -Text ''
    $box.Multiline     = $true
    $box.AcceptsReturn = $true
    $box.ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
    $box.Height        = Get-AbmScaled 150
    return $box
}

function Get-DelimiterIndex {
    param([string]$Delimiter)
    for ($i = 0; $i -lt $script:Delimiters.Count; $i++) { if ($script:Delimiters[$i].Value -eq $Delimiter) { return $i } }
    return 0
}
#endregion

#region Comportamiento
function Update-FormatFields {
    $isExcel = $script:Dialog.Format.SelectedIndex -eq 0
    $script:Dialog.Sheet.Enabled     = $isExcel
    $script:Dialog.Delimiter.Enabled = -not $isExcel
}

function Confirm-NewFileDialog {
    $names = @(Get-ColumnNames -Lines $script:Dialog.Columns.Lines)
    $problem = Test-ColumnNames -Names $names
    if ($problem) {
        Show-AbmMessage -Owner $script:Dialog.Form -Message $problem -Icon Warning
        return
    }
    $script:Dialog.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
}

function Test-ColumnNames {
    <# Devuelve el mensaje de error, o '' si las columnas son válidas. #>
    param([AllowEmptyCollection()][string[]]$Names)
    if ($Names.Count -eq 0) { return 'Ingresá al menos una columna (una por línea).' }
    $duplicates = @($Names | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    if ($duplicates.Count -gt 0) { return ('Hay columnas repetidas: {0}' -f ($duplicates -join ', ')) }
    return ''
}

function Get-ColumnNames {
    param([AllowEmptyCollection()][string[]]$Lines)
    foreach ($line in $Lines) {
        $name = $line.Trim()
        if ($name -ne '') { $name }
    }
}
#endregion

Export-ModuleMember -Function Show-AbmNewFileDialog
