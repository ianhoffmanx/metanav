# Runs only on the laptop. Microsoft Azure.Identity owns encrypted token storage and refresh.
param([ValidateSet('login','read','status')][string]$Action='read')
try {[IO.File]::WriteAllText("$env:LOCALAPPDATA\metanav\graph-sdk-$Action-entry.txt",([DateTime]::UtcNow.ToString('o')+' pid='+$PID))}catch{}
$ProgressPreference='SilentlyContinue'
$ErrorActionPreference='Stop'
[Console]::InputEncoding=[Text.Encoding]::UTF8
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$recordPath=Join-Path $env:LOCALAPPDATA 'metanav\graph-sdk-authrecord.json'
$progressPath=Join-Path $env:LOCALAPPDATA 'metanav\graph-sdk-login-status.json'
$cancel=$null
$watch=[Diagnostics.Stopwatch]::StartNew()
$timings=[ordered]@{}
function Mark-Stage($stage){
 $timings[$stage]=[Math]::Round($watch.Elapsed.TotalSeconds,2)
 try {
  $path=Join-Path $env:LOCALAPPDATA ('metanav\graph-sdk-'+$Action+'-trace.json')
  @{pid=$PID;action=$Action;stage=$stage;at=[DateTime]::UtcNow.ToString('o');elapsed=$watch.Elapsed.TotalSeconds;timings=$timings}|ConvertTo-Json -Depth 4 -Compress|Set-Content $path -Encoding utf8
 } catch {} # Diagnostics must not interrupt authentication or print an exception.
}
Mark-Stage 'started'
function Report($value){$value|ConvertTo-Json -Depth 25 -Compress|Write-Output}
function Progress-State($status){@{status=$status;at=[DateTime]::UtcNow.ToString('o')}|ConvertTo-Json -Compress|Set-Content $progressPath -Encoding utf8}
try {
 $request=[Console]::In.ReadToEnd()|ConvertFrom-Json
 if(-not $request.tenant){throw 'Missing tenant configuration'}
 if($Action -ne 'login' -and -not (Test-Path $recordPath)){Report @{ok=$false;skipped=$true;code='not_signed_in';note='Graph is not signed in'};exit 0}
 Mark-Stage 'module lookup'
 $module=Get-Module -ListAvailable Microsoft.Graph.Authentication | Where-Object {$_.Version -ge [Version]'2.41.2'} | Select-Object -First 1
 if(-not $module){throw 'Microsoft Graph Authentication SDK is not installed'}
 Mark-Stage 'load SDK'
 # Load the SDK's public Azure.Identity API without its isolated cmdlet load context.
 Get-ChildItem (Join-Path $module.ModuleBase 'Dependencies/Core') -Filter '*.dll' | Where-Object {$_.Name -notlike 'msalruntime*'} | ForEach-Object {[Reflection.Assembly]::LoadFrom($_.FullName)|Out-Null}
 Get-ChildItem (Join-Path $module.ModuleBase 'Dependencies') -Filter '*.dll' | Where-Object {$_.Name -notlike 'msalruntime*'} | ForEach-Object {[Reflection.Assembly]::LoadFrom($_.FullName)|Out-Null}
 Mark-Stage 'credential options'
 $options=[Azure.Identity.InteractiveBrowserCredentialOptions]::new()
 # This is the official Microsoft Graph PowerShell public client, already consented in SSW.
 $options.ClientId='14d82eec-204b-4c2f-b7e8-296a70dab67e'
 $options.TenantId=[string]$request.tenant
 $options.RedirectUri=[Uri]'http://localhost'
 $options.DisableAutomaticAuthentication=$true
 if($request.login_hint){$options.LoginHint=[string]$request.login_hint}
 $cache=[Azure.Identity.TokenCachePersistenceOptions]::new()
 $cache.Name='metanav.graph.readonly'
 $cache.UnsafeAllowUnencryptedStorage=$false
 $options.TokenCachePersistenceOptions=$cache
 $scopes=[string[]]@('https://graph.microsoft.com/Calendars.Read','https://graph.microsoft.com/Mail.Read')
 $tokenRequest=[Azure.Core.TokenRequestContext]::new($scopes)
 if(Test-Path $recordPath){
  $stream=[IO.File]::OpenRead($recordPath)
  try{$options.AuthenticationRecord=[Azure.Identity.AuthenticationRecord]::Deserialize($stream,[Threading.CancellationToken]::None)}finally{$stream.Dispose()}
  if($options.AuthenticationRecord.TenantId -ne $request.tenant -or $options.AuthenticationRecord.ClientId -ne $options.ClientId){throw 'Saved sign-in does not match the configured tenant and application'}
 }
 Mark-Stage 'create credential'
 $credential=[Azure.Identity.InteractiveBrowserCredential]::new($options)
 $cancel=[Threading.CancellationTokenSource]::new($(if($Action -eq 'login'){300000}elseif($Action -eq 'read'){25000}else{45000}))
 if($Action -eq 'login'){
  New-Item -ItemType Directory -Force (Split-Path $recordPath)|Out-Null
  Progress-State 'waiting'
  $record=$credential.Authenticate($tokenRequest,$cancel.Token)
  if($request.login_hint -and $record.Username -ine $request.login_hint){throw 'The selected account is not the configured SSW account'}
  $stream=[IO.File]::Create($recordPath+'.tmp')
  try{$record.Serialize($stream,$cancel.Token)}finally{$stream.Dispose()}
  Move-Item ($recordPath+'.tmp') $recordPath -Force
  Progress-State 'ready'
  Report @{ok=$true;provider='microsoft-sdk';account=$record.Username;scopes=@('Calendars.Read','Mail.Read')}
  exit 0
 }
 Mark-Stage 'silent token'
 $token=$credential.GetToken($tokenRequest,$cancel.Token)
 Mark-Stage 'token acquired'
 $headers=@{Authorization=('Bearer '+$token.Token);Prefer=('outlook.timezone="'+[string]$request.timezone+'", outlook.body-content-type="text"')}
 if($Action -eq 'status'){Report @{ok=$true;provider='microsoft-sdk';account=$options.AuthenticationRecord.Username;scopes=@('Calendars.Read','Mail.Read')};exit 0}
 function Read-GraphPages([string]$url){
  $items=[Collections.Generic.List[object]]::new()
  for($page=0;$url -and $page -lt 10;$page++){
   if(([Uri]$url).Scheme -ne 'https' -or ([Uri]$url).Host -ne 'graph.microsoft.com'){throw 'Refusing an unexpected Graph pagination host'}
   $result=Invoke-RestMethod -Method Get -Uri $url -Headers $headers -TimeoutSec 15
   foreach($item in $result.value){$items.Add($item)}
   $url=$result.'@odata.nextLink'
  }
  if($url){throw 'The Graph window exceeded the collection limit'}
  return ,$items.ToArray()
 }
 $calendarOk=$true;$calendarError=$null;$events=@()
 try {
  Mark-Stage 'calendar read'
  $from=[Uri]::EscapeDataString([string]$request.from)
  $to=[Uri]::EscapeDataString([string]$request.to)
  $url='https://graph.microsoft.com/v1.0/me/calendarView?startDateTime='+$from+'&endDateTime='+$to+'&$top=100&$orderby=start/dateTime&$select=id,subject,start,end,isAllDay,isCancelled,location,organizer,showAs,responseStatus,onlineMeeting,webLink,bodyPreview'
  $events=Read-GraphPages $url
 } catch {$calendarOk=$false;$calendarError='Graph calendar collection failed.'}
 $rawMail=$null;$mailOk=$false
 if($request.mail_folders){
  $rawMail=@{ok=$true;folders=@{}}
  try {
   if(-not $request.mail_since){throw 'Mail window is missing'}
   $since=([DateTimeOffset]::Parse($request.mail_since)).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
   $select='id,conversationId,parentFolderId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,flag,importance,hasAttachments,bodyPreview,webLink'
   foreach($folder in $request.mail_folders){
    Mark-Stage ('mail '+$folder.name)
    $dateField=if($folder.sent){'sentDateTime'}else{'receivedDateTime'}
    $base='https://graph.microsoft.com/v1.0/me/mailFolders/'+[Uri]::EscapeDataString([string]$folder.id)+'/messages?$top=100&$select='+$select
    $url=$base+'&$filter='+[Uri]::EscapeDataString($dateField+' ge '+$since)+'&$orderby='+$dateField+' desc'
    $messages=Read-GraphPages $url
    if($request.first -and -not $folder.sent){
     $flagged=Read-GraphPages ($base+'&$filter='+[Uri]::EscapeDataString("flag/flagStatus eq 'flagged'"))
     $messages+=@($flagged)
    }
    $rawMail.folders[$folder.name]=@{sent=[bool]$folder.sent;messages=@($messages)}
   }
   # Previously unresolved asks can have moved to Archive or be older than the
   # lookback. Resolve only their exact collected conversation IDs, across folders.
   $known=@($rawMail.folders.Values|Where-Object {-not $_.sent}|ForEach-Object {$_.messages}|ForEach-Object {$_.conversationId})
   $missing=@($request.tracked_conversations|Where-Object {$_ -and $_ -notin $known})
   if($missing.Count){
    $blocked=@('sentitems','deleteditems','drafts')|ForEach-Object {
     (Invoke-RestMethod -Method Get -Uri ('https://graph.microsoft.com/v1.0/me/mailFolders/'+$_+'?$select=id') -Headers $headers -TimeoutSec 10).id
    }
    $tracked=@()
    foreach($conversation in $missing){
     if($conversation -notmatch '^[A-Za-z0-9+/_=-]{20,2048}$'){throw 'Invalid tracked conversation'}
     $filter=[Uri]::EscapeDataString("conversationId eq '"+$conversation+"'")
     $rows=Read-GraphPages ('https://graph.microsoft.com/v1.0/me/messages?$top=100&$select='+$select+'&$filter='+$filter)
     $tracked+=@($rows|Where-Object {-not $_.isDraft -and $_.parentFolderId -notin $blocked -and $_.from.emailAddress.address -ine $options.AuthenticationRecord.Username})
    }
    $rawMail.folders['Tracked']=@{sent=$false;messages=@($tracked)}
   }
   $mailOk=$true
  } catch {$rawMail.ok=$false;$rawMail.error='Graph mail collection failed.'}
 } else {
  Mark-Stage 'mail check'
  try {$null=Invoke-RestMethod -Method Get -Uri 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=1&$select=id' -Headers $headers -TimeoutSec 15;$mailOk=$true}catch{}
 }
 Mark-Stage 'complete'
 Report @{ok=$calendarOk;provider='microsoft-sdk';calendar=@($events);error=$calendarError;mail=$rawMail;mail_read_ok=$mailOk;account=$options.AuthenticationRecord.Username}
} catch {
 Mark-Stage 'failed'
 $chain='';$errorType='';$ex=$_.Exception
 while($ex){$chain+=' '+$ex.GetType().FullName+' '+$ex.Message;$errorType=$ex.GetType().FullName;$ex=$ex.InnerException}
 $aad=[regex]::Match($chain,'AADSTS\d+').Value
 $code='sdk_error';$note='Microsoft Graph could not be read.'
 if($chain -match 'AuthenticationRequiredException|MsalUiRequiredException'){$code='signin_required';$note='Graph sign-in is required; use the explicit Graph login command.'}
 elseif($chain -match 'AADSTS65001|AADSTS90094|AADSTS65004'){$code='consent_required';$note='SSW approval is required for the requested Graph permissions.'}
 elseif($chain -match 'TaskCanceledException|OperationCanceledException'){$code='timeout';$note='Microsoft Graph authentication or collection timed out.'}
 elseif($chain -match 'selected account|configured tenant'){$code='account_mismatch';$note='The selected Microsoft account does not match the configured SSW account.'}
 if($Action -eq 'login'){Progress-State 'failed'}
 Report @{ok=$false;code=$code;note=$note;aad_code=$aad;error_type=$errorType}
 exit 1
} finally {if($cancel){$cancel.Dispose()}}
