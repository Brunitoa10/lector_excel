Set-StrictMode -Version Latest

<#
    Lectura de styles.xml: determina qué índices de estilo de celda (cellXfs) tienen formato de fecha/hora,
    ya sea un formato integrado de Excel o uno personalizado.
#>

$script:BuiltInDateFormats = @(14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47)

function Get-AbmDateStyleSet {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip, [AllowEmptyString()][string]$StylesPart)

    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    if (-not $StylesPart) { return ,$set }
    $doc = Open-AbmZipXml -Zip $Zip -EntryName $StylesPart
    if ($null -eq $doc) { return ,$set }

    $ns = New-AbmXmlNamespace $doc
    $customDates = Get-CustomDateFormats -Document $doc -Ns $ns
    $index = 0
    foreach ($xf in $doc.SelectNodes('/x:styleSheet/x:cellXfs/x:xf', $ns)) {
        $formatId = 0
        [void][int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$formatId)
        if ($script:BuiltInDateFormats -contains $formatId -or $customDates.Contains($formatId)) { [void]$set.Add($index) }
        $index++
    }
    return ,$set
}

function Get-CustomDateFormats {
    param([System.Xml.XmlDocument]$Document, $Ns)

    $result = New-Object 'System.Collections.Generic.HashSet[int]'
    foreach ($format in $Document.SelectNodes('/x:styleSheet/x:numFmts/x:numFmt', $Ns)) {
        if (Test-DateFormatCode -Code $format.GetAttribute('formatCode')) {
            [void]$result.Add([int]$format.GetAttribute('numFmtId'))
        }
    }
    return ,$result
}

function Test-DateFormatCode {
    <# Quita literales, colores/locales y escapes; si quedan tokens d/m/y/h/s es fecha u hora. #>
    param([string]$Code)
    $clean = [regex]::Replace($Code, '"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', '') -ireplace 'General', ''
    return ($clean -imatch '[dymhs]')
}

Export-ModuleMember -Function Get-AbmDateStyleSet
