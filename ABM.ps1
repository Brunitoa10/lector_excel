<#
.SYNOPSIS
    ABM (Alta, Baja, Modificación) de registros en libros de Excel (.xlsx / .xlsm) y archivos de texto (.txt).

.DESCRIPTION
    Aplicación de escritorio 100 % PowerShell nativo (WinForms + .NET Framework):
      - No requiere Excel instalado (lee y escribe el formato OpenXML directamente).
      - No instala módulos, no genera .exe, no usa .bat ni código C# compilado.

    Arquitectura (punto de composición):
      ABM.Config         -> Configuración (config.json + valores por defecto).
      ABM.Logger         -> Auditoría en .txt y respaldos antes de guardar.
      ABM.DataSource     -> Registro de proveedores (Strategy + Registry) y fachada de acceso a datos.
      ABM.TextProvider   -> Estrategia para .txt (delimitado o texto plano).
      ABM.ExcelProvider  -> Estrategia para .xlsx/.xlsm (OpenXML nativo).
      ABM.Theme          -> Paleta, tipografías y componentes visuales reutilizables.
      ABM.Dialogs        -> Diálogos modales (alta/modificación, nuevo archivo, mensajes).
      ABM.MainForm       -> Pantalla principal (presentación + controlador de acciones).

.EXAMPLE
    .\ABM.ps1
.EXAMPLE
    .\ABM.ps1 -Path .\inventario.xlsx
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { Write-Verbose 'Renderizado de texto ya inicializado.' }

$moduleNames = @(
    'ABM.Config'
    'ABM.Logger'
    'ABM.DataSource'
    'ABM.TextProvider'
    'ABM.ExcelProvider'
    'ABM.Theme'
    'ABM.Dialogs'
    'ABM.MainForm'
)
foreach ($moduleName in $moduleNames) {
    Import-Module -Name (Join-Path $PSScriptRoot "Modules\$moduleName.psm1") -Force -DisableNameChecking
}

$config = Get-AbmConfig -Path (Join-Path $PSScriptRoot 'config.json')
Initialize-AbmLogger -Root $PSScriptRoot -Config $config

# --- Registro de estrategias (Open/Closed: un formato nuevo = un Register-AbmProvider más) ---
Register-AbmProvider -Kind 'Excel' -Description 'Libro de Excel' -Extensions '.xlsx', '.xlsm' `
    -Read   { param($FilePath, $Sheet) Read-ExcelSheet -Path $FilePath -SheetName $Sheet } `
    -Save   { param($Source, $Table)
              Save-ExcelSheet -Path $Source.Path -SheetName $Source.Meta.Sheet -Table $Table `
                              -HeaderRow $Source.Meta.HeaderRow -ColOffset $Source.Meta.ColOffset } `
    -Create { param($FilePath, $Table, $Options) New-ExcelWorkbook -Path $FilePath -SheetName $Options.SheetName -Table $Table }

Register-AbmProvider -Kind 'Txt' -Description 'Archivo de texto' -Extensions '.txt' `
    -Read   { param($FilePath, $Sheet) Read-TextSource -Path $FilePath } `
    -Save   { param($Source, $Table)
              Save-TextSource -Path $Source.Path -Table $Table -Delimiter $Source.Meta.Delimiter -Encoding $Source.Meta.Encoding } `
    -Create { param($FilePath, $Table, $Options)
              Save-TextSource -Path $FilePath -Table $Table -Delimiter $Options.Delimiter -Encoding (New-Object System.Text.UTF8Encoding($true)) }

$startupPath = ''
if ($Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) { $startupPath = (Resolve-Path -LiteralPath $Path).ProviderPath }
    else { Write-Warning "No se encontró el archivo indicado: $Path" }
}

Show-AbmMainForm -Config $config -InitialPath $startupPath
