Set-StrictMode -Version Latest

<#
    Registro de las estrategias incluidas (Open/Closed): para soportar un formato nuevo
    se agrega un Register-AbmProvider más, sin tocar la UI ni la fachada.
#>

function Register-AbmDefaultProviders {
    [CmdletBinding()]
    param()

    Register-AbmProvider -Kind 'Excel' -Description 'Libro de Excel' -Extensions '.xlsx', '.xlsm' `
        -Read   { param($FilePath, $Sheet, $OnProgress) Read-ExcelSheet -Path $FilePath -SheetName $Sheet -OnProgress $OnProgress } `
        -Save   { param($Source, $Table)
                  Save-ExcelSheet -Path $Source.Path -SheetName $Source.Meta.Sheet -Table $Table `
                                  -HeaderRow $Source.Meta.HeaderRow -ColOffset $Source.Meta.ColOffset } `
        -Create { param($FilePath, $Table, $Options) New-ExcelWorkbook -Path $FilePath -SheetName $Options.SheetName -Table $Table }

    Register-AbmProvider -Kind 'Txt' -Description 'Archivo de texto' -Extensions '.txt' `
        -Read   { param($FilePath, $Sheet, $OnProgress) Read-TextSource -Path $FilePath -OnProgress $OnProgress } `
        -Save   { param($Source, $Table)
                  Save-TextSource -Path $Source.Path -Table $Table -Delimiter $Source.Meta.Delimiter -Encoding $Source.Meta.Encoding } `
        -Create { param($FilePath, $Table, $Options)
                  Save-TextSource -Path $FilePath -Table $Table -Delimiter $Options.Delimiter -Encoding (New-Object System.Text.UTF8Encoding($true)) }
}

Export-ModuleMember -Function Register-AbmDefaultProviders
