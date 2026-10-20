$ProgressPreference='SilentlyContinue'
$ErrorActionPreference='Stop'
# Explicit sign-in only. The hourly service never starts this task.
$repo=Split-Path $PSScriptRoot -Parent
$node=(Get-Command node.exe -ErrorAction Stop).Source
$root=Join-Path $env:LOCALAPPDATA 'metanav'
New-Item -ItemType Directory -Force $root | Out-Null
$wrapper=Join-Path $root 'graph-sdk-login.ps1'
$template=@'
$ProgressPreference='SilentlyContinue'
Set-Location '__REPO__'
& '__NODE__' graph.mjs login *> '__ROOT__\graph-sdk-login.log'
exit $LASTEXITCODE
'@
$template.Replace('__REPO__',$repo.Replace("'","''")).Replace('__NODE__',$node.Replace("'","''")).Replace('__ROOT__',$root.Replace("'","''")) | Set-Content $wrapper -Encoding UTF8
$action=New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -File "'+$wrapper+'"')
$principal=New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$settings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 6)
Register-ScheduledTask -TaskName 'MetaNav-GraphLogin' -Action $action -Principal $principal -Settings $settings -Force | Out-Null
$desktop=[Environment]::GetFolderPath('Desktop')
if($desktop){
 $link=(New-Object -ComObject WScript.Shell).CreateShortcut((Join-Path $desktop 'Meta-Nav Calendar Sign-in.lnk'))
 $link.TargetPath=Join-Path $env:WINDIR 'System32\schtasks.exe'
 $link.Arguments='/Run /TN "MetaNav-GraphLogin"'
 $link.Description='Sign Meta-Nav into Microsoft calendar and mail read access. No app settings needed.'
 $link.WindowStyle=7
 $link.Save()
}
Write-Output 'Installed the manual Meta-Nav Calendar Sign-in task and desktop shortcut.'
