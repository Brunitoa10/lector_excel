Set-StrictMode -Version Latest

<#
    Atajos de teclado definidos como datos (Command pattern): tecla + modificadores + condición + acción.
    La misma tabla genera la ayuda que se muestra en la barra de estado. Native = lo resuelve la grilla.
#>

$script:Shortcuts = @(
    @{ Key = 'S';      Ctrl = $true; Shift = $true;  Action = { [void](Invoke-AbmSaveAs) } }
    @{ Key = 'S';      Ctrl = $true; Shift = $false; Action = { [void](Invoke-AbmSave) };  Help = 'Ctrl+S Guardar' }
    @{ Key = 'O';      Ctrl = $true;                 Action = { Invoke-AbmOpen };          Help = 'Ctrl+O Abrir' }
    @{ Key = 'N';      Ctrl = $true;                 Action = { Invoke-AbmAdd };           Help = 'Ctrl+N Alta' }
    @{ Key = 'F2';     Native = $true;                                                         Help = 'F2 Editar celda' }
    @{ Key = 'E';      Ctrl = $true;                 Action = { Invoke-AbmEdit };          Help = 'Ctrl+E Formulario' }
    @{ Key = 'Delete'; When = { Test-GridIdle };     Action = { Invoke-AbmDelete };        Help = 'Supr Baja' }
    @{ Key = 'F';      Ctrl = $true;                 Action = { Enter-AbmSearch };         Help = 'Ctrl+F Buscar' }
    @{ Key = 'Escape'; When = { (Get-AbmUi).Search.Focused }; Action = { Clear-AbmSearch } }
    @{ Key = 'F5';                                   Action = { Invoke-AbmReload };        Help = 'F5 Recargar' }
    @{ Key = 'T';      Ctrl = $true;                 Action = { Invoke-AbmToggleTheme };   Help = 'Ctrl+T Tema' }
)

function Invoke-AbmShortcut {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Windows.Forms.KeyEventArgs]$KeyEvent)

    if ((Get-AbmState).Busy) { return }
    foreach ($shortcut in $script:Shortcuts) {
        if (-not (Test-Shortcut -Shortcut $shortcut -KeyEvent $KeyEvent)) { continue }
        $KeyEvent.Handled = $true
        $KeyEvent.SuppressKeyPress = $true
        & $shortcut.Action
        return
    }
}

function Get-AbmShortcutHelp {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    return (@($script:Shortcuts | Where-Object { $_.ContainsKey('Help') } | ForEach-Object { $_.Help }) -join ' · ')
}

function Test-GridIdle {
    <# La grilla tiene el foco y no se está editando una celda (Supr debe borrar texto al editar). #>
    $grid = (Get-AbmUi).Grid
    return ($grid.Focused -and -not $grid.IsCurrentCellInEditMode)
}

function Test-Shortcut {
    param([hashtable]$Shortcut, [System.Windows.Forms.KeyEventArgs]$KeyEvent)

    if ($Shortcut.ContainsKey('Native')) { return $false }
    if ($KeyEvent.KeyCode -ne [System.Windows.Forms.Keys]$Shortcut.Key) { return $false }
    if ($KeyEvent.Control -ne [bool]$Shortcut['Ctrl']) { return $false }
    if ($Shortcut.ContainsKey('Shift') -and $KeyEvent.Shift -ne $Shortcut.Shift) { return $false }
    if ($Shortcut.ContainsKey('When') -and -not (& $Shortcut.When)) { return $false }
    return $true
}

Export-ModuleMember -Function Invoke-AbmShortcut, Get-AbmShortcutHelp
