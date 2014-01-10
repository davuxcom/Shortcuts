
function New-Shortcut($TargetPath, $ShortcutPath) {
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}

function New-FavoriteShortcut($TargetPath, $DisplayName) {
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Links\$DisplayName.lnk")
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}