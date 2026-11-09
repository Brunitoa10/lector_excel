Set-StrictMode -Version Latest

<#
    Utilidades de tabla en memoria (System.Data.DataTable) compartidas por todos los proveedores
    y por la capa de aplicación. Sin dependencias de UI ni de formato de archivo.
#>

function Get-AbmUniqueNames {
    [CmdletBinding()]
    param([AllowEmptyString()][AllowEmptyCollection()][string[]]$Names)

    $seen   = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $result = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Names.Count; $i++) {
        $base = if ([string]::IsNullOrWhiteSpace($Names[$i])) { "Columna $($i + 1)" } else { $Names[$i].Trim() }
        $name = $base
        $suffix = 2
        while (-not $seen.Add($name)) { $name = "$base ($suffix)"; $suffix++ }
        $result.Add($name)
    }
    return ,$result.ToArray()
}

function New-AbmDataTable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Columns)

    $table = New-Object System.Data.DataTable 'Datos'
    foreach ($name in (Get-AbmUniqueNames -Names $Columns)) { [void]$table.Columns.Add($name, [string]) }
    return ,$table
}

function ConvertTo-AbmRecord {
    <# DataRow -> diccionario ordenado columna/valor. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Data.DataRow]$Row)

    $record = [ordered]@{}
    foreach ($column in $Row.Table.Columns) { $record[$column.ColumnName] = [string]$Row[$column] }
    return $record
}

function Format-AbmRecord {
    <# Resumen corto de un registro para auditoría: "Col=Valor; Col=Valor; …". #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Record,
        [int]$MaxFields = 6,
        [int]$MaxLength = 40
    )

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($key in $Record.Keys) {
        if ($parts.Count -ge $MaxFields) { $parts.Add('…'); break }
        $value = [string]$Record[$key]
        if ($value.Length -gt $MaxLength) { $value = $value.Substring(0, $MaxLength - 3) + '...' }
        $parts.Add(('{0}={1}' -f $key, $value))
    }
    return ($parts -join '; ')
}

Export-ModuleMember -Function Get-AbmUniqueNames, New-AbmDataTable, ConvertTo-AbmRecord, Format-AbmRecord
