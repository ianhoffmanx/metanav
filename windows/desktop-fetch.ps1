# Fetch in the signed-in desktop session. SSH cannot read this user's Windows Credential Manager.
# Register/run this with Task Scheduler's Interactive logon type; no password is stored or requested.
$ProgressPreference='SilentlyContinue'
$ErrorActionPreference='Stop'
$env:GIT_TERMINAL_PROMPT='0'
$root=Split-Path -Parent $PSScriptRoot
foreach($repo in @($root,(Join-Path (Split-Path -Parent $root) 'metanav-hermes'))){
 if(Test-Path (Join-Path $repo '.git')){
  git -C $repo -c credential.helper= -c 'credential.helper=!gh auth git-credential' fetch --all --prune
  if($LASTEXITCODE -ne 0){throw "Fetch failed in $repo. Run this task while signed in to Windows."}
 }
}
