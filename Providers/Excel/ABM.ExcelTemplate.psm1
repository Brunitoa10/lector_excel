Set-StrictMode -Version Latest

<#
    Plantilla de un libro nuevo: las partes mínimas de un .xlsx válido, con estilo corporativo
    (estilo 1 = encabezado azul marino en negrita, estilo 2 = celda con borde suave),
    panel inmovilizado y autofiltro.
#>

$script:Ns = Get-AbmSpreadsheetNamespace
$script:XmlHeader = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
$script:HeaderStyle = '1'
$script:DataStyle   = '2'

function Get-AbmTemplateStyles {
    [CmdletBinding()]
    param()
    return @{ Header = $script:HeaderStyle; Data = $script:DataStyle }
}

function ConvertTo-AbmSafeSheetName {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Name)
    $clean = ($Name -replace '[\[\]\:\*\?\/\\]', '').Trim().Trim("'")
    if (-not $clean) { $clean = 'Datos' }
    if ($clean.Length -gt 31) { $clean = $clean.Substring(0, 31) }
    return $clean
}

function Get-AbmNewWorkbookParts {
    <# Devuelve un diccionario ordenado { nombreDeParte = contenidoXml }. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SheetName, [Parameter(Mandatory)][int]$ColumnCount)

    $parts = [ordered]@{}
    $parts['[Content_Types].xml']        = Get-ContentTypesXml
    $parts['_rels/.rels']                = Get-RootRelsXml
    $parts['xl/workbook.xml']            = Get-WorkbookXml -SheetName $SheetName
    $parts['xl/_rels/workbook.xml.rels'] = Get-WorkbookRelsXml
    $parts['xl/styles.xml']              = Get-StylesXml
    $parts['xl/worksheets/sheet1.xml']   = Get-SheetXml -ColumnCount $ColumnCount
    return $parts
}

#region Partes
function Get-ContentTypesXml {
    $sml = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
    return $script:XmlHeader +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        "<Override PartName=`"/xl/workbook.xml`" ContentType=`"$sml.sheet.main+xml`"/>" +
        "<Override PartName=`"/xl/worksheets/sheet1.xml`" ContentType=`"$sml.worksheet+xml`"/>" +
        "<Override PartName=`"/xl/styles.xml`" ContentType=`"$sml.styles+xml`"/>" +
        '</Types>'
}

function Get-RootRelsXml {
    return $script:XmlHeader +
        "<Relationships xmlns=`"$($script:Ns.PkgRel)`">" +
        "<Relationship Id=`"rId1`" Type=`"$($script:Ns.DocRel)/officeDocument`" Target=`"xl/workbook.xml`"/>" +
        '</Relationships>'
}

function Get-WorkbookXml {
    param([string]$SheetName)
    $title = [System.Security.SecurityElement]::Escape($SheetName)
    return $script:XmlHeader +
        "<workbook xmlns=`"$($script:Ns.Main)`" xmlns:r=`"$($script:Ns.DocRel)`">" +
        "<sheets><sheet name=`"$title`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
}

function Get-WorkbookRelsXml {
    return $script:XmlHeader +
        "<Relationships xmlns=`"$($script:Ns.PkgRel)`">" +
        "<Relationship Id=`"rId1`" Type=`"$($script:Ns.DocRel)/worksheet`" Target=`"worksheets/sheet1.xml`"/>" +
        "<Relationship Id=`"rId2`" Type=`"$($script:Ns.DocRel)/styles`" Target=`"styles.xml`"/>" +
        '</Relationships>'
}

function Get-StylesXml {
    $border = '<color rgb="FFDDE3EC"/>'
    return $script:XmlHeader +
        "<styleSheet xmlns=`"$($script:Ns.Main)`">" +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF112240"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
        "<border><left style=`"thin`">$border</left><right style=`"thin`">$border</right>" +
        "<top style=`"thin`">$border</top><bottom style=`"thin`">$border</bottom><diagonal/></border></borders>" +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        '</styleSheet>'
}

function Get-SheetXml {
    param([int]$ColumnCount)
    $lastColumn = ConvertTo-AbmColumnLetters ($ColumnCount - 1)
    $columns = (1..$ColumnCount | ForEach-Object { '<col min="{0}" max="{0}" width="24" customWidth="1"/>' -f $_ }) -join ''
    return $script:XmlHeader +
        "<worksheet xmlns=`"$($script:Ns.Main)`" xmlns:r=`"$($script:Ns.DocRel)`">" +
        "<dimension ref=`"A1:$($lastColumn)1`"/>" +
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0">' +
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
        '<sheetFormatPr defaultRowHeight="15"/>' +
        "<cols>$columns</cols>" +
        '<sheetData/>' +
        "<autoFilter ref=`"A1:$($lastColumn)1`"/>" +
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
        '</worksheet>'
}
#endregion

Export-ModuleMember -Function Get-AbmTemplateStyles, ConvertTo-AbmSafeSheetName, Get-AbmNewWorkbookParts
