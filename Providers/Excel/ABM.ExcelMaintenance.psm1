Set-StrictMode -Version Latest

<#
    Mantenimiento del libro después de reescribir una hoja: rangos (dimensión y autofiltro),
    tabla de Excel (ListObject) y cadena de cálculo, para que Excel abra el archivo sin "reparar".
#>

function Update-AbmSheetRanges {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.Xml.XmlDocument]$Document, [Parameter(Mandatory)]$Ns,
        [string]$FirstColumn, [string]$LastColumn, [int]$HeaderRow, [int]$LastRow
    )

    $dimension = $Document.SelectSingleNode('/x:worksheet/x:dimension', $Ns)
    if ($null -ne $dimension) { $dimension.SetAttribute('ref', ('A1:{0}{1}' -f $LastColumn, $LastRow)) }

    $autoFilter = $Document.SelectSingleNode('/x:worksheet/x:autoFilter', $Ns)
    if ($null -ne $autoFilter) { $autoFilter.SetAttribute('ref', ('{0}{1}:{2}{3}' -f $FirstColumn, $HeaderRow, $LastColumn, $LastRow)) }
}

function Update-AbmSheetTables {
    <# Ajusta el rango de la tabla de Excel si la hoja tiene exactamente una. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [Parameter(Mandatory)][string]$SheetPart, [Parameter(Mandatory)][string]$Range)

    $relationships = Get-AbmRelationships -Zip $Zip -PartName $SheetPart
    $tables = @($relationships | Where-Object { $_.Type -like '*/table' -and $_.Part })
    if ($tables.Count -ne 1) { return }

    $doc = Open-AbmZipXml -Zip $Zip -EntryName $tables[0].Part
    if ($null -eq $doc) { return }
    $doc.DocumentElement.SetAttribute('ref', $Range)
    $autoFilter = $doc.SelectSingleNode('/x:table/x:autoFilter', (New-AbmXmlNamespace $doc))
    if ($null -ne $autoFilter) { $autoFilter.SetAttribute('ref', $Range) }
    Save-AbmZipXml -Zip $Zip -EntryName $tables[0].Part -Document $doc
}

function Remove-AbmCalcChain {
    <# La cadena de cálculo queda desactualizada al reescribir celdas; Excel la regenera sola. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [Parameter(Mandatory)]$Info)

    $calc = @($Info.Rels | Where-Object { $_.Type -like '*/calcChain' })
    if ($calc.Count -eq 0) { return }

    $entry = $Zip.GetEntry($calc[0].Part)
    if ($null -ne $entry) { $entry.Delete() }
    Remove-PartReference -Zip $Zip -EntryName (Get-AbmRelsPath $Info.Part) -Attribute 'Id' -Value $calc[0].Id
    Remove-PartReference -Zip $Zip -EntryName '[Content_Types].xml' -Attribute 'PartName' -Value ('/' + $calc[0].Part)
}

function Remove-PartReference {
    param([System.IO.Compression.ZipArchive]$Zip, [string]$EntryName, [string]$Attribute, [string]$Value)

    $doc = Open-AbmZipXml -Zip $Zip -EntryName $EntryName
    if ($null -eq $doc) { return }
    foreach ($node in @($doc.DocumentElement.ChildNodes)) {
        if ($node.NodeType -eq [System.Xml.XmlNodeType]::Element -and $node.GetAttribute($Attribute) -eq $Value) {
            [void]$doc.DocumentElement.RemoveChild($node)
        }
    }
    Save-AbmZipXml -Zip $Zip -EntryName $EntryName -Document $doc
}

Export-ModuleMember -Function Update-AbmSheetRanges, Update-AbmSheetTables, Remove-AbmCalcChain
