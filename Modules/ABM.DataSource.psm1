Set-StrictMode -Version Latest

<#
    Fachada de acceso a datos + Registro de estrategias (patrones Strategy / Registry / Facade).

    Cada proveedor (formato de archivo) se registra con tres operaciones:
      Read   { param($Path, $Sheet)           -> @{ Table = <DataTable>; Meta = <hashtable> } }
      Save   { param($Source, $Table)         -> sobrescribe el archivo de origen }
      Create { param($Path, $Table, $Options) -> crea un archivo nuevo con el contenido de la tabla }

    La UI sólo conoce esta fachada (Dependency Inversion): no sabe nada de OpenXML ni de delimitadores.
#>

$script:Providers = New-Object System.Collections.Generic.List[object]

#region Registro de proveedores
function Register-AbmProvider {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string[]]$Extensions,
        [Parameter(Mandatory)][scriptblock]$Read,
        [Parameter(Mandatory)][scriptblock]$Save,
        [Parameter(Mandatory)][scriptblock]$Create
    )

    $script:Providers.Add([pscustomobject]@{
        Kind        = $Kind
        Description = $Description
        Extensions  = @($Extensions | ForEach-Object { $_.ToLowerInvariant() })
        Read        = $Read
        Save        = $Save
        Create      = $Create
    })
}

function Get-AbmProvider {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    foreach ($provider in $script:Providers) {
        if ($provider.Extensions -contains $extension) { return $provider }
    }
    $valid = @($script:Providers | ForEach-Object { $_.Extensions }) -join ', '
    throw "Formato no soportado: '$extension'. Formatos válidos: $valid"
}

function Get-AbmFileFilter {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    $all   = @(foreach ($provider in $script:Providers) { foreach ($ext in $provider.Extensions) { "*$ext" } })
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add(('Archivos soportados ({0})|{0}' -f ($all -join ';')))
    foreach ($provider in $script:Providers) {
        $patterns = @($provider.Extensions | ForEach-Object { "*$_" }) -join ';'
        $parts.Add(('{0} ({1})|{1}' -f $provider.Description, $patterns))
    }
    return ($parts -join '|')
}
#endregion

#region Fachada
function Open-AbmSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Sheet
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "No existe el archivo: $Path" }
    $fullPath = (Resolve-Path -LiteralPath $Path).ProviderPath
    $provider = Get-AbmProvider -Path $fullPath
    $result   = & $provider.Read $fullPath $Sheet

    [pscustomobject]@{
        Path  = $fullPath
        Kind  = $provider.Kind
        Meta  = $result.Meta
        Table = $result.Table
    }
}

function Save-AbmSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Source,
        [Parameter(Mandatory)][System.Data.DataTable]$Table
    )
    $provider = Get-AbmProvider -Path $Source.Path
    $null = & $provider.Save $Source $Table
}

function New-AbmSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [hashtable]$Options = @{}
    )
    $provider = Get-AbmProvider -Path $Path
    $null = & $provider.Create $Path $Table $Options
}
#endregion

#region Utilidades de tabla (compartidas por todos los proveedores)
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
#endregion

Export-ModuleMember -Function Register-AbmProvider, Get-AbmProvider, Get-AbmFileFilter,
                              Open-AbmSource, Save-AbmSource, New-AbmSource,
                              Get-AbmUniqueNames, New-AbmDataTable
