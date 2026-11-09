Set-StrictMode -Version Latest

<#
    Diálogo de alta / modificación: un campo de texto por columna.
    Devuelve un diccionario ordenado columna -> valor, o $null si se canceló.
#>

$script:Dialog = @{}

function Show-AbmRecordDialog {
    [CmdletBinding()]
    param(
        $Owner,
        [Parameter(Mandatory)][string]$Title,
        [string]$Subtitle = '',
        [Parameter(Mandatory)][string[]]$Columns,
        [System.Collections.IDictionary]$Values
    )

    $shell = New-AbmDialogShell -Title $Title -Subtitle $Subtitle -BodyHeight (Get-BodyHeight -FieldCount $Columns.Count)
    $boxes = Add-RecordFields -Shell $shell -Columns $Columns -Values $Values
    $script:Dialog = @{ Form = $shell.Form; Boxes = $boxes }

    $shell.Ok.Add_Click({ Confirm-RecordDialog })
    $shell.Form.Add_Shown({ Select-FirstField })

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

function Get-BodyHeight {
    param([int]$FieldCount)
    return [math]::Min([math]::Max($FieldCount * 58 + 24, 140), 480)
}

function Add-RecordFields {
    param($Shell, [string[]]$Columns, [System.Collections.IDictionary]$Values)

    $boxes = [ordered]@{}
    foreach ($column in $Columns) {
        $box = New-Object System.Windows.Forms.TextBox
        $box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        if ($null -ne $Values -and $Values.Contains($column)) { $box.Text = [string]$Values[$column] }
        Add-AbmDialogField -Shell $Shell -Caption $column -Control $box
        $boxes[$column] = $box
    }
    return $boxes
}

function Confirm-RecordDialog {
    <# Validación: al menos un campo con contenido. #>
    $hasValue = @($script:Dialog.Boxes.Values | Where-Object { $_.Text.Trim() -ne '' }).Count -gt 0
    if (-not $hasValue) {
        Show-AbmMessage -Owner $script:Dialog.Form -Message 'Completá al menos un campo para registrar.' -Icon Warning
        return
    }
    $script:Dialog.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
}

function Select-FirstField {
    $first = @($script:Dialog.Boxes.Values) | Select-Object -First 1
    if ($null -ne $first) { [void]$first.Focus(); $first.SelectAll() }
}

Export-ModuleMember -Function Show-AbmRecordDialog
