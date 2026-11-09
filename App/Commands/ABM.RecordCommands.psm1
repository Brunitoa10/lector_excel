Set-StrictMode -Version Latest

<#
    Comandos del ABM sobre registros: alta, modificación y baja.
    Los cambios quedan en memoria (estado "Pendiente") hasta que se guarda.
#>

function Invoke-AbmAdd {
    [CmdletBinding()]
    param()

    if (-not (Test-AbmHasData)) { return }
    $values = Show-AbmRecordDialog -Owner (Get-AbmUi).Form -Title 'Alta de registro' `
        -Subtitle ('Nuevo registro en {0}' -f (Get-AbmSourceLabel)) -Columns (Get-AbmColumnNames)
    if ($null -eq $values) { return }

    $table = (Get-AbmState).Table
    $row = $table.NewRow()
    foreach ($key in $values.Keys) { $row[$key] = $values[$key] }
    $table.Rows.Add($row)

    Write-AbmAudit -Action 'ALTA' -Target (Get-AbmSourceLabel) -Detail (Format-AbmRecord -Record $values)
    Set-AbmDirty
    Select-AbmRow -Row $row
    Set-AbmStatus 'Registro agregado. Recordá guardar los cambios (Ctrl+S).'
}

function Invoke-AbmEdit {
    [CmdletBinding()]
    param()

    $rows = @(Get-AbmSelectedRows)
    if ($rows.Count -eq 0) { return }
    if ($rows.Count -gt 1) {
        Show-AbmMessage -Owner (Get-AbmUi).Form -Message 'Seleccioná un único registro para modificar.'
        return
    }

    $row     = $rows[0]
    $current = ConvertTo-AbmRecord -Row $row
    $values  = Show-AbmRecordDialog -Owner (Get-AbmUi).Form -Title 'Modificar registro' `
        -Subtitle (Get-AbmSourceLabel) -Columns (Get-AbmColumnNames) -Values $current
    if ($null -eq $values) { return }

    $changes = @(Update-RecordValues -Row $row -Current $current -Values $values)
    if ($changes.Count -eq 0) { Set-AbmStatus 'El registro no tuvo cambios.'; return }

    Write-AbmAudit -Action 'MODIFICACION' -Target (Get-AbmSourceLabel) -Detail ($changes -join ' | ')
    Set-AbmDirty
    Set-AbmStatus ('Registro modificado ({0} campo(s)). Recordá guardar los cambios.' -f $changes.Count)
}

function Invoke-AbmDelete {
    [CmdletBinding()]
    param()

    $rows = @(Get-AbmSelectedRows)
    if ($rows.Count -eq 0) { return }
    if (-not (Confirm-AbmAction -Owner (Get-AbmUi).Form -Title 'Confirmar baja' -Message (Get-DeleteQuestion -Count $rows.Count))) { return }

    $table = (Get-AbmState).Table
    foreach ($row in $rows) {
        $detail = Format-AbmRecord -Record (ConvertTo-AbmRecord -Row $row)
        $table.Rows.Remove($row)
        Write-AbmAudit -Action 'BAJA' -Target (Get-AbmSourceLabel) -Detail $detail
    }
    Set-AbmDirty
    Set-AbmStatus ('{0} registro(s) dado(s) de baja. Recordá guardar los cambios.' -f $rows.Count)
}

#region Internos
function Update-RecordValues {
    <# Aplica los valores que cambiaron y devuelve la descripción de cada cambio. #>
    param([System.Data.DataRow]$Row, [System.Collections.IDictionary]$Current, [System.Collections.IDictionary]$Values)
    foreach ($key in $Values.Keys) {
        if ($Current[$key] -cne $Values[$key]) {
            $Row[$key] = $Values[$key]
            "{0}: '{1}' → '{2}'" -f $key, $Current[$key], $Values[$key]
        }
    }
}

function Get-DeleteQuestion {
    param([int]$Count)
    $question = if ($Count -eq 1) { '¿Confirmás la baja del registro seleccionado?' } else { '¿Confirmás la baja de {0} registros?' -f $Count }
    return "$question`n`nEl cambio se aplicará al archivo cuando guardes."
}
#endregion

Export-ModuleMember -Function Invoke-AbmAdd, Invoke-AbmEdit, Invoke-AbmDelete
