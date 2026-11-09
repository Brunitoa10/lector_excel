Set-StrictMode -Version Latest

<#
    Estrategia Excel (.xlsx / .xlsm) - API pública del proveedor. Orquesta los módulos:
      ABM.OpenXmlPackage  (ZIP / XML / relaciones)    ABM.ExcelReader      (hoja -> tabla)
      ABM.ExcelStyles     (formatos de fecha)         ABM.ExcelWriter      (tabla -> hoja)
      ABM.ExcelMaintenance (rangos, tablas, cálculo)  ABM.ExcelTemplate    (libro nuevo)
    No necesita Excel instalado ni COM.
    Limitación conocida: las fórmulas dentro del rango editado se guardan como valores.
#>

function Read-ExcelSheet {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$SheetName)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $info   = Get-AbmWorkbookInfo -Zip $zip
        $sheet  = Select-AbmSheet -Info $info -SheetName $SheetName
        $doc    = Open-AbmZipXml -Zip $zip -EntryName $sheet.Part
        if ($null -eq $doc) { throw "No se encontró la hoja '$($sheet.Name)' dentro del archivo." }

        $shared = Get-AbmSharedStrings -Zip $zip -Part $info.SharedStringsPart
        $dates  = Get-AbmDateStyleSet -Zip $zip -StylesPart $info.StylesPart
        $rows   = Read-AbmSheetRows -Document $doc -Shared $shared -DateStyles $dates
        $layout = Get-AbmSheetLayout -Rows $rows

        return @{
            Table = ConvertTo-AbmSheetTable -Rows $rows -Layout $layout
            Meta  = @{
                Sheets    = @($info.Sheets | ForEach-Object { $_.Name })
                Sheet     = $sheet.Name
                HeaderRow = $layout.HeaderRow
                ColOffset = $layout.ColOffset
            }
        }
    }
    finally { $zip.Dispose() }
}

function Save-ExcelSheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SheetName,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [int]$HeaderRow = 1,
        [int]$ColOffset = 0,
        [string]$DefaultHeaderStyle = '',
        [string]$DefaultDataStyle = ''
    )

    if ($HeaderRow -lt 1) { $HeaderRow = 1 }
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $info  = Get-AbmWorkbookInfo -Zip $zip
        $sheet = Select-AbmSheet -Info $info -SheetName $SheetName
        $doc   = Open-AbmZipXml -Zip $zip -EntryName $sheet.Part
        $ns    = New-AbmXmlNamespace $doc
        $sheetData = $doc.SelectSingleNode('/x:worksheet/x:sheetData', $ns)
        if ($null -eq $sheetData) { throw "La hoja '$($sheet.Name)' no tiene una estructura válida." }

        $styles  = Remove-AbmTableRows -SheetData $sheetData -Ns $ns -HeaderRow $HeaderRow
        $lastRow = Write-AbmTableRows -SheetData $sheetData -Table $Table -HeaderRow $HeaderRow -ColOffset $ColOffset `
                       -Styles $styles -DateStyles (Get-AbmDateStyleSet -Zip $zip -StylesPart $info.StylesPart) `
                       -DefaultHeaderStyle $DefaultHeaderStyle -DefaultDataStyle $DefaultDataStyle

        $first = ConvertTo-AbmColumnLetters $ColOffset
        $last  = ConvertTo-AbmColumnLetters ($ColOffset + [math]::Max(1, $Table.Columns.Count) - 1)
        Update-AbmSheetRanges -Document $doc -Ns $ns -FirstColumn $first -LastColumn $last -HeaderRow $HeaderRow -LastRow $lastRow
        Save-AbmZipXml -Zip $zip -EntryName $sheet.Part -Document $doc

        $tableRange = '{0}{1}:{2}{3}' -f $first, $HeaderRow, $last, [math]::Max($lastRow, $HeaderRow + 1)
        Update-AbmSheetTables -Zip $zip -SheetPart $sheet.Part -Range $tableRange
        Remove-AbmCalcChain -Zip $zip -Info $info
    }
    finally { $zip.Dispose() }
}

function New-ExcelWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SheetName = 'Datos',
        [Parameter(Mandatory)][System.Data.DataTable]$Table
    )

    $SheetName = ConvertTo-AbmSafeSheetName $SheetName
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }

    $parts = Get-AbmNewWorkbookParts -SheetName $SheetName -ColumnCount ([math]::Max(1, $Table.Columns.Count))
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try { foreach ($name in $parts.Keys) { Save-AbmZipText -Zip $zip -EntryName $name -Content $parts[$name] } }
    finally { $zip.Dispose() }

    $styles = Get-AbmTemplateStyles
    Save-ExcelSheet -Path $Path -SheetName $SheetName -Table $Table -DefaultHeaderStyle $styles.Header -DefaultDataStyle $styles.Data
}

Export-ModuleMember -Function Read-ExcelSheet, Save-ExcelSheet, New-ExcelWorkbook
