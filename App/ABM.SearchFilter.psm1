Set-StrictMode -Version Latest

<#
    Búsqueda sobre todas las columnas. New-AbmRowFilter es una función pura que arma la expresión
    de DataView.RowFilter escapando comodines y comillas (no hay inyección en la expresión).
#>

function New-AbmRowFilter {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyCollection()][string[]]$Columns, [AllowEmptyString()][string]$Text)

    $term = $Text.Trim()
    if ($term -eq '' -or $Columns.Count -eq 0) { return '' }

    $literal = ConvertTo-LikeLiteral -Text $term
    $clauses = foreach ($column in $Columns) { "[{0}] LIKE '%{1}%'" -f (ConvertTo-ColumnReference $column), $literal }
    return ($clauses -join ' OR ')
}

function Set-AbmFilter {
    [CmdletBinding()]
    param()

    $view = (Get-AbmState).View
    if ($null -eq $view) { return }
    try { $view.RowFilter = New-AbmRowFilter -Columns @(Get-AbmColumnNames) -Text (Get-AbmUi).Search.Text }
    catch { Set-AbmStatus "No se pudo aplicar el filtro: $($_.Exception.Message)" }
    Update-AbmView
}

function Enter-AbmSearch {
    [CmdletBinding()]
    param()
    $search = (Get-AbmUi).Search
    if (-not $search.Enabled) { return }
    [void]$search.Focus()
    $search.SelectAll()
}

function Clear-AbmSearch {
    [CmdletBinding()]
    param()
    (Get-AbmUi).Search.Text = ''
}

#region Internos
function ConvertTo-LikeLiteral {
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

function ConvertTo-ColumnReference {
    param([string]$Name)
    return $Name.Replace('\', '\\').Replace(']', '\]')
}
#endregion

Export-ModuleMember -Function New-AbmRowFilter, Set-AbmFilter, Enter-AbmSearch, Clear-AbmSearch
