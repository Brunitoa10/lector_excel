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
    <#
        Ajusta la tabla de Excel (ListObject) si la hoja tiene exactamente una: rango y, si cambiaron
        las columnas, la lista de columnas (Excel exige que coincida con los encabezados).
        -Columns: @{ Name; SourceIndex } en el orden actual.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory)][string]$SheetPart,
        [Parameter(Mandatory)][string]$Range,
        [AllowEmptyCollection()][object[]]$Columns = @()
    )

    $relationships = Get-AbmRelationships -Zip $Zip -PartName $SheetPart
    $tables = @($relationships | Where-Object { $_.Type -like '*/table' -and $_.Part })
    if ($tables.Count -ne 1) { return }

    $doc = Open-AbmZipXml -Zip $Zip -EntryName $tables[0].Part
    if ($null -eq $doc) { return }
    $ns = New-AbmXmlNamespace $doc
    $doc.DocumentElement.SetAttribute('ref', $Range)
    $autoFilter = $doc.SelectSingleNode('/x:table/x:autoFilter', $ns)
    if ($null -ne $autoFilter) { $autoFilter.SetAttribute('ref', $Range) }
    if ($Columns.Count -gt 0 -and (Update-TableColumns -Document $doc -Ns $ns -Columns $Columns) -and $null -ne $autoFilter) {
        foreach ($filter in @($autoFilter.SelectNodes('x:filterColumn', $ns))) { [void]$autoFilter.RemoveChild($filter) }
    }
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

function Update-TableColumns {
    <# Reconstruye <tableColumns> según las columnas actuales. Devuelve $true si hubo cambios. #>
    param([System.Xml.XmlDocument]$Document, $Ns, [object[]]$Columns)

    $container = $Document.SelectSingleNode('/x:table/x:tableColumns', $Ns)
    if ($null -eq $container) { return $false }
    $existing = @($container.SelectNodes('x:tableColumn', $Ns))
    if (-not (Test-ColumnsChanged -Existing $existing -Columns $Columns)) { return $false }

    $nextId = 0
    foreach ($node in $existing) { $nextId = [math]::Max($nextId, [int]$node.GetAttribute('id')) }
    foreach ($node in $existing) { [void]$container.RemoveChild($node) }
    foreach ($column in $Columns) {
        $source = $column.SourceIndex
        if ($null -ne $source -and [int]$source -lt $existing.Count) { $node = $existing[[int]$source] }
        else {
            $nextId++
            $node = $Document.CreateElement($container.Prefix, 'tableColumn', $Document.DocumentElement.NamespaceURI)
            $node.SetAttribute('id', [string]$nextId)
        }
        $node.SetAttribute('name', $column.Name)
        [void]$container.AppendChild($node)
    }
    $container.SetAttribute('count', [string]$Columns.Count)
    return $true
}

function Test-ColumnsChanged {
    param([object[]]$Existing, [object[]]$Columns)
    if ($Existing.Count -ne $Columns.Count) { return $true }
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        $source = $Columns[$i].SourceIndex
        if ($null -eq $source -or [int]$source -ne $i -or $Existing[$i].GetAttribute('name') -cne $Columns[$i].Name) { return $true }
    }
    return $false
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
