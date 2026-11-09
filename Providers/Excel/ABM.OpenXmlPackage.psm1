Set-StrictMode -Version Latest

<#
    Acceso al paquete OpenXML (.xlsx = ZIP de partes XML): lectura/escritura de partes,
    relaciones (.rels) y estructura del libro (hojas, textos compartidos, estilos).
#>

$script:Namespaces = @{
    Main   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    DocRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    PkgRel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    Xml    = 'http://www.w3.org/XML/1998/namespace'
}

function Get-AbmSpreadsheetNamespace {
    [CmdletBinding()]
    param()
    return $script:Namespaces
}

#region Partes XML dentro del ZIP
function Open-AbmZipXml {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [AllowEmptyString()][string]$EntryName)

    $entry = $Zip.GetEntry($EntryName)
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8, $true)
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.PreserveWhitespace = $true
        $doc.XmlResolver = $null
        $doc.LoadXml($reader.ReadToEnd())
        return ,$doc
    }
    finally { $reader.Dispose() }
}

function Save-AbmZipXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory)][string]$EntryName,
        [Parameter(Mandatory)][System.Xml.XmlDocument]$Document
    )

    $existing = $Zip.GetEntry($EntryName)
    if ($null -ne $existing) { $existing.Delete() }
    $stream = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal).Open()
    try {
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
        $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
        try { $Document.Save($writer) } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Save-AbmZipText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory)][string]$EntryName,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Content
    )
    if ($Zip.Mode -ne [System.IO.Compression.ZipArchiveMode]::Create) {
        $existing = $Zip.GetEntry($EntryName)
        if ($null -ne $existing) { $existing.Delete() }
    }
    $stream = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal).Open()
    $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($Content) } finally { $writer.Dispose() }
}

function New-AbmXmlNamespace {
    <# Administrador de espacios de nombres con el prefijo 'x' para SpreadsheetML. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Xml.XmlDocument]$Document)
    $ns = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
    $ns.AddNamespace('x', $script:Namespaces.Main)
    return ,$ns
}
#endregion

#region Relaciones
function Get-AbmRelsPath {
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$PartName)
    $directory = Get-PartDirectory $PartName
    $file = $PartName.Substring($PartName.LastIndexOf('/') + 1)
    if ($directory) { return "$directory/_rels/$file.rels" }
    return "_rels/$file.rels"
}

function Get-AbmRelationships {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [AllowEmptyString()][string]$PartName)

    $list = New-Object System.Collections.Generic.List[object]
    $doc  = Open-AbmZipXml -Zip $Zip -EntryName (Get-AbmRelsPath $PartName)
    if ($null -eq $doc) { return ,$list.ToArray() }

    $baseDirectory = Get-PartDirectory $PartName
    foreach ($node in $doc.DocumentElement.ChildNodes) {
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $target   = $node.GetAttribute('Target')
        $external = $node.GetAttribute('TargetMode') -eq 'External'
        $list.Add([pscustomobject]@{
            Id     = $node.GetAttribute('Id')
            Type   = $node.GetAttribute('Type')
            Target = $target
            Part   = if ($external) { '' } else { Resolve-PartName -BaseDirectory $baseDirectory -Target $target }
        })
    }
    return ,$list.ToArray()
}

function Get-PartDirectory {
    param([AllowEmptyString()][string]$PartName)
    $index = $PartName.LastIndexOf('/')
    if ($index -lt 0) { return '' }
    return $PartName.Substring(0, $index)
}

function Resolve-PartName {
    param([AllowEmptyString()][string]$BaseDirectory, [string]$Target)

    if ($Target.StartsWith('/')) { return $Target.TrimStart('/') }
    $segments = New-Object System.Collections.Generic.List[string]
    foreach ($segment in ($BaseDirectory -split '/')) { if ($segment) { $segments.Add($segment) } }
    foreach ($segment in ($Target -split '/')) {
        if ($segment -eq '..') { if ($segments.Count -gt 0) { $segments.RemoveAt($segments.Count - 1) } }
        elseif ($segment -and $segment -ne '.') { $segments.Add($segment) }
    }
    return ($segments -join '/')
}
#endregion

#region Estructura del libro
function Get-AbmWorkbookInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip)

    $workbookPart = 'xl/workbook.xml'
    foreach ($rel in (Get-AbmRelationships -Zip $Zip -PartName '')) {
        if ($rel.Type -like '*/officeDocument') { $workbookPart = $rel.Part }
    }
    $doc = Open-AbmZipXml -Zip $Zip -EntryName $workbookPart
    if ($null -eq $doc) { throw 'El archivo no es un libro de Excel válido (no contiene workbook.xml).' }

    $rels = Get-AbmRelationships -Zip $Zip -PartName $workbookPart
    [pscustomobject]@{
        Part              = $workbookPart
        Rels              = $rels
        Sheets            = Get-WorksheetList -Document $doc -Rels $rels
        SharedStringsPart = Get-RelatedPart -Rels $rels -TypeSuffix 'sharedStrings'
        StylesPart        = Get-RelatedPart -Rels $rels -TypeSuffix 'styles'
    }
}

function Select-AbmSheet {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Info, [string]$SheetName)

    $all = @($Info.Sheets)
    if ($all.Count -eq 0) { throw 'El libro no contiene hojas de cálculo editables.' }
    if ($SheetName) { foreach ($sheet in $all) { if ($sheet.Name -eq $SheetName) { return $sheet } } }
    return $all[0]
}

function Get-WorksheetList {
    param([System.Xml.XmlDocument]$Document, [object[]]$Rels)

    $byId = @{}
    foreach ($rel in $Rels) { $byId[$rel.Id] = $rel }
    $sheets = New-Object System.Collections.Generic.List[object]
    foreach ($node in $Document.SelectNodes('/x:workbook/x:sheets/x:sheet', (New-AbmXmlNamespace $Document))) {
        $relationId = $node.GetAttribute('id', $script:Namespaces.DocRel)
        if ($byId.ContainsKey($relationId) -and $byId[$relationId].Type -like '*/worksheet') {
            $sheets.Add([pscustomobject]@{ Name = $node.GetAttribute('name'); Part = $byId[$relationId].Part })
        }
    }
    return ,$sheets.ToArray()
}

function Get-RelatedPart {
    param([object[]]$Rels, [string]$TypeSuffix)
    foreach ($rel in $Rels) { if ($rel.Type -like "*/$TypeSuffix") { return $rel.Part } }
    return ''
}
#endregion

Export-ModuleMember -Function Get-AbmSpreadsheetNamespace, Open-AbmZipXml, Save-AbmZipXml, Save-AbmZipText, New-AbmXmlNamespace,
                              Get-AbmRelsPath, Get-AbmRelationships, Get-AbmWorkbookInfo, Select-AbmSheet
