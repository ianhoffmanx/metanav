import { tool, WIN } from './common.mjs';
export function notify(message) {
  if (process.platform === 'darwin') return tool('osascript', ['-e', 'on run argv', '-e', 'display notification (item 2 of argv) with title (item 1 of argv) sound name "Glass"', '-e', 'end run', 'Meta-Nav', message],{timeout:15000}).status===0;
  if (!WIN) return false;
  const ps = `$ErrorActionPreference='Stop'
[Console]::InputEncoding=[Text.Encoding]::UTF8
$t = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$x=$t.GetElementsByTagName('text'); $x.Item(0).AppendChild($t.CreateTextNode('Meta-Nav needs you')) > $null; $x.Item(1).AppendChild($t.CreateTextNode([Console]::In.ReadToEnd())) > $null
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\\WindowsPowerShell\\v1.0\\powershell.exe').Show([Windows.UI.Notifications.ToastNotification]::new($t))`;
  const result=tool('powershell', ['-NoProfile','-NonInteractive','-EncodedCommand',Buffer.from(ps,'utf16le').toString('base64')],{input:message,timeout:15000});
  return result.status===0;
}
