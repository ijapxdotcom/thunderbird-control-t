$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $DesktopPath "Control-T Licitacoes.lnk"
$TargetPath = "c:\PIBIC\thunderbird-control-t\Executar Control-T.bat"
$WorkingDir = "c:\PIBIC\thunderbird-control-t"
$Description = "Control-T - Compliance e Routing Engine v3.0"

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = $WorkingDir
$Shortcut.Description = $Description
$Shortcut.WindowStyle = 1
$Shortcut.Save()

Write-Host ""
Write-Host "  Atalho criado com sucesso na Area de Trabalho!" -ForegroundColor Green
Write-Host "  Local: $ShortcutPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Para executar todo dia automaticamente:" -ForegroundColor Yellow
Write-Host "  1. Abra o Agendador de Tarefas do Windows (taskschd.msc)" -ForegroundColor White
Write-Host "  2. Crie uma nova tarefa basica" -ForegroundColor White
Write-Host "  3. Gatilho: Diario, 08:00" -ForegroundColor White
Write-Host "  4. Acao: Iniciar programa" -ForegroundColor White
Write-Host "     Caminho: $TargetPath" -ForegroundColor White
Write-Host ""
