<#
.SYNOPSIS
    Pruebas automáticas de la capa de datos (sin interfaz gráfica y sin dependencias: no requiere Pester).

.DESCRIPTION
    Crea archivos temporales, ejercita lectura / alta / modificación / baja / guardado en Excel y TXT
    y verifica los resultados. Devuelve código de salida 0 si todo pasó y 1 si alguna prueba falló.

.EXAMPLE
    .\Tests\Test-Providers.ps1
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data, System.IO.Compression, System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
$modules = @(
    'Data\ABM.DataTable', 'Data\ABM.DataSource'
    'Providers\Text\ABM.TextEncoding', 'Providers\Text\ABM.DelimitedText', 'Providers\Text\ABM.TextProvider'
    'Providers\Excel\ABM.OpenXmlPackage', 'Providers\Excel\ABM.CellAddress', 'Providers\Excel\ABM.CellValue'
    'Providers\Excel\ABM.ExcelStyles', 'Providers\Excel\ABM.ExcelReader', 'Providers\Excel\ABM.ExcelWriter'
    'Providers\Excel\ABM.ExcelMaintenance', 'Providers\Excel\ABM.ExcelTemplate', 'Providers\Excel\ABM.ExcelProvider'
    'Providers\ABM.ProviderSetup'
)
foreach ($module in $modules) { Import-Module (Join-Path $root "$module.psm1") -Force -DisableNameChecking }
Register-AbmDefaultProviders

$script:Passed = 0
$script:Failed = 0
$work = Join-Path ([System.IO.Path]::GetTempPath()) ('abm_tests_' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $work)

function Assert-Equal {
    param([string]$Name, $Expected, $Actual)
    if ([string]$Expected -ceq [string]$Actual) { $script:Passed++; Write-Host "  [OK]    $Name" -ForegroundColor Green }
    else { $script:Failed++; Write-Host "  [FALLA] $Name -> esperado '$Expected', obtenido '$Actual'" -ForegroundColor Red }
}

function Get-RowText {
    param([System.Data.DataTable]$Table, [int]$Index)
    return ($Table.Rows[$Index].ItemArray -join '|')
}

try {
    Write-Host 'Utilidades' -ForegroundColor Cyan
    Assert-Equal 'Letras de columna AA'  'AA' (ConvertTo-AbmColumnLetters 26)
    Assert-Equal 'Índice de columna AB'  27   (ConvertFrom-AbmColumnLetters 'AB')
    Assert-Equal 'Nombres únicos'        'A|a (2)|Columna 3' ((Get-AbmUniqueNames -Names @('A', 'a', '')) -join '|')
    Assert-Equal 'Número con ruido'      '656.69' (Format-AbmCellNumber -Raw '656.6900000000001')
    Assert-Equal 'Código con cero'       ''       ([string](ConvertTo-AbmCellNumber -Value '007'))
    Assert-Equal 'Campo con delimitador' '"a;b"'  (Format-AbmDelimitedField -Value 'a;b' -Delimiter ';')

    Write-Host 'TXT' -ForegroundColor Cyan
    $txt = Join-Path $work 'datos.txt'
    [System.IO.File]::WriteAllText($txt, "Código;Descripción;Stock`r`nA1;`"Tornillo; largo`";10`r`nA2;`"Dice `"`"hola`"`"`";5`r`n", (New-Object System.Text.UTF8Encoding($false)))
    $source = Open-AbmSource -Path $txt
    Assert-Equal 'Delimitador detectado' ';' $source.Meta.Delimiter
    Assert-Equal 'Encabezados'           'Código|Descripción|Stock' ((@($source.Table.Columns | ForEach-Object ColumnName)) -join '|')
    Assert-Equal 'Campo con comillas'    'A2|Dice "hola"|5' (Get-RowText $source.Table 1)
    $source.Table.Rows[0]['Stock'] = '99'
    [void]$source.Table.Rows.Add([object[]]@('A3', 'Nuevo', '1'))
    $source.Table.Rows.Remove($source.Table.Rows[1])
    Save-AbmSource -Source $source -Table $source.Table
    $again = Open-AbmSource -Path $txt
    Assert-Equal 'TXT guardado: registros' 2 $again.Table.Rows.Count
    Assert-Equal 'TXT guardado: fila 1'    'A1|Tornillo; largo|99' (Get-RowText $again.Table 0)
    Assert-Equal 'TXT conserva UTF-8'      'UTF-8' $again.Meta.EncodingName

    Write-Host 'Excel' -ForegroundColor Cyan
    $xlsx  = Join-Path $work 'libro.xlsx'
    $table = New-AbmDataTable -Columns @('Serial', 'Equipo', 'Precio', 'Activo')
    [void]$table.Rows.Add([object[]]@('007', 'Notebook ñandú', '1250.5', 'VERDADERO'))
    [void]$table.Rows.Add([object[]]@('SN2', 'Monitor', '300', 'FALSO'))
    New-AbmSource -Path $xlsx -Table $table -Options @{ SheetName = 'Equipos' }
    $book = Open-AbmSource -Path $xlsx
    Assert-Equal 'Hoja'                  'Equipos' $book.Meta.Sheet
    Assert-Equal 'Fila 1 (texto, ñ, bool)' '007|Notebook ñandú|1250.5|VERDADERO' (Get-RowText $book.Table 0)
    $book.Table.Rows[1]['Equipo'] = 'Monitor 24"'
    $book.Table.Rows[0]['Equipo'] = '  Notebook <A&B> '
    [void]$book.Table.Rows.Add([object[]]@('SN3', 'Impresora', '99.9', 'VERDADERO'))
    Save-AbmSource -Source $book -Table $book.Table
    $book = Open-AbmSource -Path $xlsx
    Assert-Equal 'Excel guardado: registros' 3 $book.Table.Rows.Count
    Assert-Equal 'Excel guardado: modificación' 'SN2|Monitor 24"|300|FALSO' (Get-RowText $book.Table 1)
    Assert-Equal 'Excel guardado: alta'         'SN3|Impresora|99.9|VERDADERO' (Get-RowText $book.Table 2)
    Assert-Equal 'Excel: espacios y < & >'      '  Notebook <A&B> ' $book.Table.Rows[0]['Equipo']

    $converted = Join-Path $work 'convertido.txt'
    New-AbmSource -Path $converted -Table $book.Table -Options @{ Delimiter = "`t" }
    Assert-Equal 'Excel -> TXT (tab)' "`t" (Open-AbmSource -Path $converted).Meta.Delimiter
}
finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

$color = if ($script:Failed -eq 0) { 'Green' } else { 'Red' }
Write-Host ("`nResultado: {0} OK, {1} con falla" -f $script:Passed, $script:Failed) -ForegroundColor $color
exit ([int]($script:Failed -gt 0))
