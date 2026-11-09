Set-StrictMode -Version Latest

<#
    Atajos de teclado definidos como datos (Command pattern): tecla + modificadores + condición + acción.
    La misma tabla genera la ayuda que se muestra en la barra de estado.
#>

$script:Shortcuts = @(
    @{ Key = 'S';      Ctrl = $true; Shift = $true;  Action = { [void](Invoke-AbmSaveAs) } }
    @{ Key = 'S';      Ctrl = $true; Shift = $false; Action = { [void](Invoke-AbmSave) };  Help = 'Ctrl+S Guardar' }
    @{ Key = 'O';      Ctrl = $true;                 Action = { Invoke-AbmOpen };          Help = 'Ctrl+O Abrir' }
    @{ Key = 'N';      Ctrl = $true;                 Action = { Invoke-AbmAdd };           Help = 'Ctrl+N Alta' }
    @{ Key = 'F2';                                   Action = { Invoke-AbmEdit };          Help = 'F2 Modificar' }
    @{ Key = 'Delete'; When = { (Get-AbmUi).Grid.Focused };   Action = { Invoke-AbmDelete };  Help = 'Supr Baja' }
    @{ Key = 'F';      Ctrl = $true;                 Action = { Enter-AbmSearch };         Help = 'Ctrl+F Buscar' }
    @{ Key = 'Escape'; When = { (Get-AbmUi).Search.Focused }; Action = { Clear-AbmSearch } }
    @{ Key = 'F5';                                   Action = { Invoke-AbmReload };        Help = 'F5 Recargar' }
)

function Invoke-AbmShortcut {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Windows.Forms.KeyEventArgs]$KeyEvent)

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

function Test-Shortcut {
    param([hashtable]$Shortcut, [System.Windows.Forms.KeyEventArgs]$KeyEvent)

    if ($KeyEvent.KeyCode -ne [System.Windows.Forms.Keys]$Shortcut.Key) { return $false }
    if ($KeyEvent.Control -ne [bool]$Shortcut['Ctrl']) { return $false }
    if ($Shortcut.ContainsKey('Shift') -and $KeyEvent.Shift -ne $Shortcut.Shift) { return $false }
    if ($Shortcut.ContainsKey('When') -and -not (& $Shortcut.When)) { return $false }
    return $true
}

Export-ModuleMember -Function Invoke-AbmShortcut, Get-AbmShortcutHelp
