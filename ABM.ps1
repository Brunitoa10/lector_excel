<#
.SYNOPSIS
    ABM (Alta, Baja, Modificación) de registros en libros de Excel (.xlsx / .xlsm) y archivos de texto (.txt).

.DESCRIPTION
    Aplicación de escritorio 100 % PowerShell nativo (WinForms + .NET Framework):
      - No requiere Excel instalado (lee y escribe el formato OpenXML directamente).
      - No instala módulos, no genera .exe, no usa .bat ni código C# compilado.

    Capas (una carpeta por capa, un módulo por responsabilidad):
      Core\       Configuración y trazabilidad (auditoría + respaldos).
      Data\       Tabla en memoria y fachada de acceso a datos (Strategy + Registry + Facade).
      Providers\  Estrategias por formato: Text\ (.txt) y Excel\ (.xlsx / .xlsm, OpenXML).
      UI\         Tema visual, controles reutilizables y diálogos.
      App\        Contexto, vistas, comandos, atajos, eventos y ventana principal.

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

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Data, System.IO.Compression, System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { Write-Verbose 'Renderizado de texto ya inicializado.' }

# Orden de carga: cada módulo después de aquellos que usa al importarse.
$modules = @(
    'Core\ABM.Config', 'Core\ABM.Logger'
    'Data\ABM.DataTable', 'Data\ABM.DataSource'
    'Providers\Text\ABM.TextEncoding', 'Providers\Text\ABM.DelimitedText', 'Providers\Text\ABM.TextProvider'
    'Providers\Excel\ABM.OpenXmlPackage', 'Providers\Excel\ABM.CellAddress', 'Providers\Excel\ABM.CellValue'
    'Providers\Excel\ABM.ExcelStyles', 'Providers\Excel\ABM.ExcelReader', 'Providers\Excel\ABM.ExcelWriter'
    'Providers\Excel\ABM.ExcelMaintenance', 'Providers\Excel\ABM.ExcelTemplate', 'Providers\Excel\ABM.ExcelProvider'
    'Providers\ABM.ProviderSetup'
    'UI\ABM.Theme', 'UI\ABM.Controls', 'UI\ABM.KpiCard', 'UI\ABM.GridStyle'
    'UI\Dialogs\ABM.Messages', 'UI\Dialogs\ABM.DialogShell', 'UI\Dialogs\ABM.RecordDialog', 'UI\Dialogs\ABM.NewFileDialog'
    'App\ABM.AppContext', 'App\ABM.Session', 'App\ABM.ViewUpdater', 'App\ABM.SearchFilter'
    'App\Views\ABM.HeaderView', 'App\Views\ABM.ToolbarView', 'App\Views\ABM.KpiView', 'App\Views\ABM.GridView', 'App\Views\ABM.StatusView'
    'App\Commands\ABM.FileCommands', 'App\Commands\ABM.SaveCommands', 'App\Commands\ABM.RecordCommands'
    'App\ABM.Shortcuts', 'App\ABM.Events', 'App\ABM.MainForm'
)
foreach ($module in $modules) {
    Import-Module -Name (Join-Path $PSScriptRoot "$module.psm1") -Force -DisableNameChecking
}

$config = Get-AbmConfig -Path (Join-Path $PSScriptRoot 'config.json')
Initialize-AbmLogger -Root $PSScriptRoot -Config $config
Register-AbmDefaultProviders

$startupPath = ''
if ($Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) { $startupPath = (Resolve-Path -LiteralPath $Path).ProviderPath }
    else { Write-Warning "No se encontró el archivo indicado: $Path" }
}

Show-AbmMainForm -Config $config -InitialPath $startupPath
