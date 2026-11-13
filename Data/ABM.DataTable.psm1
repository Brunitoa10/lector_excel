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

#region Administración de columnas
function Test-AbmColumnName {
    <# Devuelve el mensaje de error, o '' si el nombre es válido. -ExceptIndex: columna que se está renombrando. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][System.Data.DataTable]$Table, [AllowEmptyString()][string]$Name, [int]$ExceptIndex = -1)

    $clean = $Name.Trim()
    if ($clean -eq '') { return 'Ingresá un nombre para la columna.' }
    for ($i = 0; $i -lt $Table.Columns.Count; $i++) {
        if ($i -ne $ExceptIndex -and $Table.Columns[$i].ColumnName -ieq $clean) { return ('Ya existe una columna llamada «{0}».' -f $clean) }
    }
    return ''
}

function Add-AbmTableColumn {
    <# Agrega una columna de texto vacía en la posición indicada (0 = primera). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Data.DataTable]$Table, [Parameter(Mandatory)][string]$Name, [int]$Position = -1)

    $problem = Test-AbmColumnName -Table $Table -Name $Name
    if ($problem) { throw $problem }
    $column = New-Object System.Data.DataColumn($Name.Trim(), [string])
    $column.DefaultValue = ''
    $Table.Columns.Add($column)
    $last = $Table.Columns.Count - 1
    if ($Position -ge 0 -and $Position -lt $last) { $column.SetOrdinal($Position) }
}

function Rename-AbmTableColumn {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Data.DataTable]$Table, [Parameter(Mandatory)][int]$Index, [Parameter(Mandatory)][string]$Name)

    $problem = Test-AbmColumnName -Table $Table -Name $Name -ExceptIndex $Index
    if ($problem) { throw $problem }
    $Table.Columns[$Index].ColumnName = $Name.Trim()
}

function Remove-AbmTableColumn {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Data.DataTable]$Table, [Parameter(Mandatory)][int]$Index)

    if ($Table.Columns.Count -le 1) { throw 'El archivo debe conservar al menos una columna.' }
    $Table.Columns.RemoveAt($Index)
}
#endregion

Export-ModuleMember -Function Get-AbmUniqueNames, New-AbmDataTable, ConvertTo-AbmRecord, Format-AbmRecord,
                              Test-AbmColumnName, Add-AbmTableColumn, Rename-AbmTableColumn, Remove-AbmTableColumn
