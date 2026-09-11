#requires -Version 5.1
<#[
DSH 极简 Agent - PowerShell runtime
将原 dsh-mini.py 的运行时功能迁移到 Windows PowerShell 5.1 / PowerShell 7。

直接运行：
  powershell -ExecutionPolicy Bypass -File .\dsh-mini.ps1
  pwsh       -File .\dsh-mini.ps1 -Prompt "列出当前目录"

远程一行启动（脚本托管后）：
  irm https://raw.githubusercontent.com/<owner>/<repo>/main/dsh-mini.ps1 | iex

安全说明：脚本只会在本机调用配置的 OpenAI 兼容 API；API Key 保存在本机配置文件，
不会上传到第三方。远程执行前请审阅 raw 文件内容，或固定到你自己的仓库/提交。
]#>

function Invoke-DshMiniScript {
    [CmdletBinding()]
    param(
        [string]$Prompt,
        [string]$Config,
        [string]$BaseUrl,
        [string]$ApiKey,
        [string]$Model,
        [string]$Cwd,
        # Validate this after binding so Windows PowerShell 5.1 irm|iex cannot
        # reject an omitted value before the script has a chance to normalize it.
        [string]$ShellMode = 'auto',
        [int]$MaxRounds = 0,
        [int]$ShellTimeout = 0,
        [switch]$NoStream,
        [switch]$NoPicker,
        [switch]$PickModel,
        [switch]$NoSave,
        [switch]$QuietTools,
        [switch]$Setup,
        [switch]$Models,
        [switch]$SelfTest,
        [switch]$ShellCheck,
        [switch]$Cli,
        [switch]$Gui,
        [switch]$Version
    )

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ShellMode)) { $ShellMode = 'auto' }
if ($ShellMode -notin @('auto','persistent','oneshot')) {
    throw "ShellMode 必须是 auto、persistent 或 oneshot；收到：$ShellMode"
}

$script:AppName = 'dsh-mini'
$script:AppTitle = 'DSH 极简 Agent'
$script:Version = '1.1.5-ps1'
$script:ConfigName = 'dsh-mini.config.json'
$script:TruncatedEditor = '<response clipped><NOTE>To save on context only part of this file has been shown to you. You should retry this tool after you have searched inside the file with `Select-String` in order to find the line numbers of what you are looking for.</NOTE>'
$script:TruncatedShell = '<response clipped><NOTE>To save on context only part of this file has been shown to you. You should retry this tool after you have searched inside the file with Select-String in order to find the line numbers of what you are looking for.</NOTE>'
$script:Persona = 'You are a helpful software engineer assistant.'
$script:GuiLogName = 'dsh-mini-gui.log'
$script:DiagLogName = 'dsh-mini-diagnose.txt'
$script:Messages = New-Object System.Collections.ArrayList
$script:ShellState = $null
$script:ConfigData = $null
$script:CurrentCancel = $null

function Get-ScriptLocation {
    if ($PSScriptRoot) { return $PSScriptRoot }
    $pathProperty = $MyInvocation.MyCommand.PSObject.Properties['Path']
    if ($pathProperty -and $pathProperty.Value) { return (Split-Path -Parent ([string]$pathProperty.Value)) }
    if ($MyInvocation.ScriptName) { return (Split-Path -Parent $MyInvocation.ScriptName) }
    return (Get-Location).Path
}

$script:BaseLocation = Get-ScriptLocation

function Get-DefaultConfig {
    return [ordered]@{
        base_url = 'https://api.deepseek.com'
        api_key = ''
        model = 'deepseek-chat'
        temperature = 0
        max_tokens = 0
        stream = $true
        request_timeout = 300
        retries = 3
        include_usage = $true
        extra_headers = [ordered]@{}
        extra_body = [ordered]@{}
        model_picker = $true
        system_prompt = $script:Persona
        include_env_note = $true
        max_output_chars = 16000
        shell_timeout_ms = 300000
        shell_exe = ''
        shell_mode = 'auto'
        shell_probe_timeout_ms = 20000
        shell_encoding = 'auto'
        cwd = ''
        max_tool_rounds = 60
        show_reasoning = $true
        save_sessions = $false
        show_live_output = $true
    }
}

function ConvertTo-HashtableDeep([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = [ordered]@{}
        foreach ($key in $Value.Keys) { $h[[string]$key] = ConvertTo-HashtableDeep $Value[$key] }
        return $h
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $h = [ordered]@{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-HashtableDeep $p.Value }
        return $h
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $a = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$a.Add((ConvertTo-HashtableDeep $item)) }
        return $a.ToArray()
    }
    return $Value
}

function Get-ObjectValue([object]$Object, [string]$Name) {
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Merge-Config([System.Collections.IDictionary]$Base, [object[]]$Layers) {
    $result = [ordered]@{}
    foreach ($key in $Base.Keys) { $result[$key] = $Base[$key] }
    foreach ($layer in $Layers) {
        if ($null -eq $layer) { continue }
        $source = ConvertTo-HashtableDeep $layer
        if ($source -isnot [System.Collections.IDictionary]) { continue }
        foreach ($key in $source.Keys) {
            if ($null -eq $source[$key]) { continue }
            if ($source[$key] -is [System.Collections.IDictionary] -and $result.Contains($key) -and $result[$key] -is [System.Collections.IDictionary]) {
                $inner = [ordered]@{}
                foreach ($innerKey in $result[$key].Keys) { $inner[$innerKey] = $result[$key][$innerKey] }
                foreach ($innerKey in $source[$key].Keys) { $inner[$innerKey] = $source[$key][$innerKey] }
                $result[$key] = $inner
            } else { $result[$key] = $source[$key] }
        }
    }
    return $result
}

function Test-WritableDirectory([string]$Path) {
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        $probe = Join-Path $Path ('.dsh-mini-write-' + [guid]::NewGuid().ToString('N'))
        [IO.File]::WriteAllText($probe, 'ok', [Text.Encoding]::UTF8)
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch { return $false }
}

function Get-DataRoot {
    if (Test-WritableDirectory $script:BaseLocation) { return $script:BaseLocation }
    $fallbackRoot = Join-Path ($(if ($env:APPDATA) { $env:APPDATA } else { [Environment]::GetFolderPath('ApplicationData') })) $script:AppName
    Test-WritableDirectory $fallbackRoot | Out-Null
    return $fallbackRoot
}

function Get-ConfigPath([string]$RequestedPath) {
    if ($RequestedPath) { return [IO.Path]::GetFullPath($RequestedPath) }
    $root = Get-DataRoot
    return (Join-Path $root $script:ConfigName)
}

function Get-SessionsDirectory {
    $path = Join-Path (Get-DataRoot) 'sessions'
    Test-WritableDirectory $path | Out-Null
    return $path
}

function Read-ConfigFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{} }
    try { return ConvertTo-HashtableDeep (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch { Write-Warning ("配置文件读取失败（已忽略）：{0} -> {1}" -f $Path, $_.Exception.Message); return @{} }
}

function Save-ConfigFile([string]$Path, [object]$Data) {
    $parent = Split-Path -Parent $Path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $json = $Data | ConvertTo-Json -Depth 30
    [IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false))
}

function Get-EnvironmentConfig {
    $h = [ordered]@{}
    foreach ($pair in @(
        @('base_url','DSH_MINI_BASE_URL','OPENAI_BASE_URL'),
        @('api_key','DSH_MINI_API_KEY','OPENAI_API_KEY'),
        @('model','DSH_MINI_MODEL','OPENAI_MODEL')
    )) {
        foreach ($name in $pair[1..($pair.Count - 1)]) {
            if (Get-Item -Path ('Env:' + $name) -ErrorAction SilentlyContinue) {
                $value = (Get-Item -Path ('Env:' + $name)).Value
                if ($value) { $h[$pair[0]] = $value; break }
            }
        }
    }
    return $h
}

function Get-EffectiveConfig {
    $path = Get-ConfigPath $Config
    $cli = [ordered]@{}
    if ($BaseUrl) { $cli.base_url = $BaseUrl }
    if ($ApiKey) { $cli.api_key = $ApiKey }
    if ($Model) { $cli.model = $Model }
    if ($Cwd) { $cli.cwd = $Cwd }
    if ($ShellMode) { $cli.shell_mode = $ShellMode }
    if ($MaxRounds -gt 0) { $cli.max_tool_rounds = $MaxRounds }
    if ($ShellTimeout -gt 0) { $cli.shell_timeout_ms = $ShellTimeout }
    if ($NoStream) { $cli.stream = $false }
    if ($NoPicker) { $cli.model_picker = $false }
    if ($PickModel) { $cli.model_picker = $true }
    if ($NoSave) { $cli.save_sessions = $false }
    $effective = Merge-Config (Get-DefaultConfig) @((Read-ConfigFile $path), (Get-EnvironmentConfig), $cli)
    if (-not $effective.cwd) { $effective.cwd = (Get-Location).Path }
    if (-not (Test-Path -LiteralPath $effective.cwd -PathType Container)) { $effective.cwd = (Get-Location).Path }
    return $effective
}

function Normalize-BaseUrl([string]$Value) {
    $url = ([string]$Value).Trim().TrimEnd('/')
    if (-not $url) { return '' }
    if ($url.EndsWith('/chat/completions')) { return $url }
    if ($url -match '/v\d+$') { return ($url + '/chat/completions') }
    try {
        $uri = [Uri]$url
        if ($uri.AbsolutePath -and $uri.AbsolutePath -ne '/') { return ($url + '/chat/completions') }
    } catch { }
    return ($url + '/v1/chat/completions')
}

function Get-ModelsEndpoint([string]$BaseUrlValue) {
    $endpoint = Normalize-BaseUrl $BaseUrlValue
    return ($endpoint -replace '/chat/completions$','') + '/models'
}

function Get-ApiHeaders([System.Collections.IDictionary]$ConfigValue, [bool]$Stream) {
    $headers = [ordered]@{
        Accept = $(if ($Stream) { 'text/event-stream' } else { 'application/json' })
        'User-Agent' = "$($script:AppName)/$($script:Version) (PowerShell)"
    }
    if ($ConfigValue.api_key) { $headers.Authorization = 'Bearer ' + $ConfigValue.api_key }
    if ($ConfigValue.extra_headers) {
        foreach ($key in $ConfigValue.extra_headers.Keys) { $headers[$key] = [string]$ConfigValue.extra_headers[$key] }
    }
    return $headers
}

function ConvertTo-PowerShellLiteral([string]$Value) {
    return "'" + ([string]$Value).Replace("'", "''") + "'"
}

function Read-WebExceptionBody([System.Net.WebException]$Exception) {
    try {
        $stream = $Exception.Response.GetResponseStream()
        $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList $stream
        $text = $reader.ReadToEnd(); $reader.Dispose(); $stream.Dispose(); return $text
    } catch { return '' }
}

function Format-HttpError([int]$Code, [string]$Detail) {
    $message = $Detail
    try {
        $parsed = $Detail | ConvertFrom-Json
        $parsedError = Get-ObjectValue $parsed 'error'
        if ($parsedError) { $errorMessage = Get-ObjectValue $parsedError 'message'; $message = if ($errorMessage) { [string]$errorMessage } else { [string]$parsedError } }
    } catch { }
    $hint = switch ($Code) { 401 {'（请检查 api_key 是否正确）'} 403 {'（权限不足或地区限制）'} 404 {'（请检查 base_url 与 model 名称）'} 429 {'（触发限流，稍后重试）'} default {''} }
    return ('HTTP {0} {1} {2}' -f $Code, ([string]$message).Trim().Substring(0,[Math]::Min(800,([string]$message).Trim().Length)), $hint).Trim()
}

function Invoke-HttpJson([string]$Uri, [string]$Method, [System.Collections.IDictionary]$Headers, [string]$Body, [int]$TimeoutSeconds = 30) {
    $request = [Net.HttpWebRequest]::Create($Uri)
    $request.Method = $Method
    $request.Timeout = $TimeoutSeconds * 1000
    $request.ReadWriteTimeout = $TimeoutSeconds * 1000
    foreach ($key in $Headers.Keys) {
        if ($key -eq 'Accept') { $request.Accept = $Headers[$key] }
        elseif ($key -eq 'User-Agent') { $request.UserAgent = $Headers[$key] }
        elseif ($key -eq 'Content-Type') { $request.ContentType = $Headers[$key] }
        elseif ($key -eq 'Authorization') { $request.Headers['Authorization'] = $Headers[$key] }
        else { $request.Headers[$key] = $Headers[$key] }
    }
    if ($Body) {
        $bytes = (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false).GetBytes($Body)
        $request.ContentLength = $bytes.Length
        $stream = $request.GetRequestStream(); $stream.Write($bytes,0,$bytes.Length); $stream.Dispose()
    }
    try {
        $response = $request.GetResponse()
        $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList $response.GetResponseStream(), (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false)
        $result = $reader.ReadToEnd(); $reader.Dispose(); $response.Dispose(); return $result
    } catch [Net.WebException] {
        $response = $_.Exception.Response
        if ($response) {
            $code = [int]$response.StatusCode
            throw (Format-HttpError $code (Read-WebExceptionBody $_.Exception))
        }
        throw ('网络错误：' + $_.Exception.Message)
    }
}

function Get-ModelIds([System.Collections.IDictionary]$ConfigValue) {
    $uri = Get-ModelsEndpoint $ConfigValue.base_url
    $body = Invoke-HttpJson $uri 'GET' (Get-ApiHeaders $ConfigValue $false) $null 15
    try { $data = $body | ConvertFrom-Json } catch { throw ('模型列表不是合法 JSON：' + $body.Substring(0,[Math]::Min(200,$body.Length))) }
    $items = Get-ObjectValue $data 'data'
    if (-not $items) { throw '接口未按 OpenAI 兼容格式返回 data 列表' }
    $ids = @($items | ForEach-Object { if ($_ -is [string]) { $_ } else { $id = Get-ObjectValue $_ 'id'; if ($id) { [string]$id } } })
    return @($ids | Sort-Object -Unique)
}

function Mask-Secret([string]$Value) {
    if (-not $Value) { return '(未设置)' }
    if ($Value.Length -le 8) { return $Value.Substring(0,[Math]::Min(2,$Value.Length)) + '***' }
    return $Value.Substring(0,4) + '***' + $Value.Substring($Value.Length - 4)
}

function Read-SecretConsole([string]$PromptText) {
    Write-Host -NoNewline $PromptText
    $chars = New-Object -TypeName System.Collections.Generic.List[char]
    while ($true) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'Enter') { Write-Host ''; break }
        if ($key.Key -eq 'Backspace') { if ($chars.Count -gt 0) { $chars.RemoveAt($chars.Count-1); Write-Host -NoNewline "`b `b" }; continue }
        if ($key.Key -eq 'Escape') { Write-Host ''; throw '已取消配置' }
        if ($key.KeyChar -and [int][char]$key.KeyChar -ge 32) { [void]$chars.Add($key.KeyChar); Write-Host -NoNewline '*' }
    }
    return (-join $chars)
}

function Invoke-Setup([System.Collections.IDictionary]$ConfigValue) {
    Write-Host "`n$('=' * 68)`n $($script:AppTitle) v$($script:Version)  配置向导`n$('=' * 68)"
    Write-Host '只支持常规 OpenAI Chat Completions 格式。直接回车使用括号内的默认值。'
    $base = Read-Host ("接口地址 base_url [$($ConfigValue.base_url)]")
    if ($base) { $ConfigValue.base_url = $base }
    $key = Read-SecretConsole ("API Key [$(Mask-Secret $ConfigValue.api_key)] ")
    if ($key) { $ConfigValue.api_key = $key }
    $model = Read-Host ("模型名 model [$($ConfigValue.model)]")
    if ($model) { $ConfigValue.model = $model }
    $path = Get-ConfigPath $Config
    Save-ConfigFile $path $ConfigValue
    Write-Host "已保存配置：$path`n"
    return $ConfigValue
}

function Select-Model([System.Collections.IDictionary]$ConfigValue) {
    try { $models = @(Get-ModelIds $ConfigValue) } catch { Write-Warning ("没能取到模型列表：$($_.Exception.Message)"); $models = @() }
    if ($models.Count -gt 0) {
        Write-Host "可用模型（$((Get-ModelsEndpoint $ConfigValue.base_url))）："
        for ($i=0; $i -lt $models.Count; $i++) { Write-Host ("  {0,2}) {1}" -f ($i+1),$models[$i]) }
    } else { Write-Host '可以直接手输模型名；没有 /models 不影响聊天功能。' }
    $answer = Read-Host "模型（回车保持 $($ConfigValue.model)）"
    if (-not $answer) { return $null }
    $answer = $answer.Trim('"').Trim("'")
    if ($answer -match '^\d+$' -and $models.Count -gt 0) {
        $index = [int]$answer - 1
        if ($index -ge 0 -and $index -lt $models.Count) { return $models[$index] }
        Write-Warning '序号超出范围，保持当前模型。'; return $null
    }
    return $answer
}

function Get-PowerShellExecutable([System.Collections.IDictionary]$ConfigValue) {
    if ($ConfigValue.shell_exe -and (Test-Path -LiteralPath $ConfigValue.shell_exe -PathType Leaf)) { return $ConfigValue.shell_exe }
    foreach ($name in @('pwsh.exe','powershell.exe','pwsh','powershell')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd.Source }
    }
    throw '没有找到可用的 PowerShell 可执行文件'
}

function New-PersistentShell([System.Collections.IDictionary]$ConfigValue) {
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'; $runspace.ThreadOptions = 'ReuseThread'; $runspace.Open()
    $ps = [PowerShell]::Create(); $ps.Runspace = $runspace
    $init = "Set-Location -LiteralPath " + (ConvertTo-PowerShellLiteral $ConfigValue.cwd)
    $ps.AddScript($init).Invoke() | Out-Null; $ps.Commands.Clear()
    return [pscustomobject]@{ Mode='persistent'; Runspace=$runspace; PowerShell=$ps; Exe='runspace'; Version=$PSVersionTable.PSVersion.ToString(); Cwd=$ConfigValue.cwd }
}

function Close-PersistentShell {
    if ($null -eq $script:ShellState) { return }
    try { $script:ShellState.PowerShell.Dispose(); $script:ShellState.Runspace.Close(); $script:ShellState.Runspace.Dispose() } catch { }
    $script:ShellState = $null
}

function Invoke-OneShotShell([string]$Command, [System.Collections.IDictionary]$ConfigValue, [int]$TimeoutMs) {
    $exe = Get-PowerShellExecutable $ConfigValue
    $tempPath = Join-Path ([IO.Path]::GetTempPath()) ('dsh-mini-' + [guid]::NewGuid().ToString('N') + '.ps1')
    [IO.File]::WriteAllText($tempPath, $Command, (New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $true))
    $start = New-Object -TypeName System.Diagnostics.ProcessStartInfo
    $quotedPath = '"' + $tempPath.Replace('"','\"') + '"'
    $start.FileName = $exe; $start.Arguments = '-NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ' + $quotedPath
    $start.WorkingDirectory = $ConfigValue.cwd; $start.UseShellExecute = $false; $start.CreateNoWindow = $true; $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true
    $process = New-Object -TypeName System.Diagnostics.Process; $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw ('无法启动 ' + $exe) }
    } catch { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue; throw }
    $outTask = $process.StandardOutput.ReadToEndAsync(); $errTask = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit($TimeoutMs)) { try { $process.Kill() } catch {}; Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue; return [pscustomobject]@{ Text='Your command timed out after the configured timeout. Below is partial output:'; Code=$null; Note='shell reset' } }
    $text = $outTask.Result + $errTask.Result
    Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    return [pscustomobject]@{ Text=$text.Trim("`r","`n"); Code=$process.ExitCode; Note='' }
}

function Invoke-PersistentShell([string]$Command, [System.Collections.IDictionary]$ConfigValue, [int]$TimeoutMs) {
    if ($null -eq $script:ShellState) { $script:ShellState = New-PersistentShell $ConfigValue }
    if ($script:ShellState.Cwd -ne $ConfigValue.cwd) { Close-PersistentShell; $script:ShellState = New-PersistentShell $ConfigValue }
    $ps = $script:ShellState.PowerShell
    $ps.Commands.Clear()
    $wrapped = "& {`n$Command`n} *>&1"
    $async = $ps.AddScript($wrapped).BeginInvoke()
    if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs)) {
        Close-PersistentShell
        return [pscustomobject]@{ Text='Your command timed out after the configured timeout. Below is partial output:'; Code=$null; Note='persistent PowerShell runspace reset' }
    }
    try { $items = @($ps.EndInvoke($async)) } catch { $items = @($_.Exception.Message) }
    $lines = @($items | ForEach-Object { if ($null -ne $_) { $_.ToString() } })
    $code = 0
    try { $last = $script:ShellState.Runspace.SessionStateProxy.GetVariable('LASTEXITCODE'); if ($null -ne $last) { $code = [int]$last } } catch { }
    return [pscustomobject]@{ Text=($lines -join [Environment]::NewLine).Trim(); Code=$code; Note='' }
}

function Invoke-ShellCommand([string]$Command, [System.Collections.IDictionary]$ConfigValue, [scriptblock]$OnOutput) {
    if (-not $Command.Trim()) { throw 'command must be a non-empty string' }
    $mode = [string]$ConfigValue.shell_mode
    $fallbackNote = ''
    if ($mode -eq 'oneshot') { $result = Invoke-OneShotShell $Command $ConfigValue ([int]$ConfigValue.shell_timeout_ms) }
    else {
        try { $result = Invoke-PersistentShell $Command $ConfigValue ([int]$ConfigValue.shell_timeout_ms) }
        catch {
            if ($mode -eq 'persistent') { throw }
            Close-PersistentShell; $ConfigValue.shell_mode = 'oneshot'; $result = Invoke-OneShotShell $Command $ConfigValue ([int]$ConfigValue.shell_timeout_ms)
            $fallbackNote = 'persistent runspace unavailable; switched to one-shot mode'
        }
    }
    $resultNote = [string](Get-ObjectValue $result 'Note')
    if ($fallbackNote) { $resultNote = $fallbackNote }
    $result = [pscustomobject]@{
        Text = Get-ResultText $result
        Code = Get-ObjectValue $result 'Code'
        Note = $resultNote
    }
    $resultText = $result.Text
    if ($OnOutput -and $resultText) { & $OnOutput $resultText }
    return $result
}

function Limit-Output([string]$Text, [int]$MaxChars, [string]$Marker) {
    if ($null -eq $Text) { return '' }
    if ($Text.Length -le $MaxChars) { return $Text }
    return $Text.Substring(0,$MaxChars) + $Marker
}

function Invoke-PwshTool([System.Collections.IDictionary]$Args, [System.Collections.IDictionary]$ConfigValue, [scriptblock]$OnOutput) {
    if (-not $Args.command) { throw 'command must be a non-empty string' }
    $result = Invoke-ShellCommand ([string]$Args.command) $ConfigValue $OnOutput
    $resultText = Get-ResultText $result
    $body = if ($resultText) { $resultText } else { '(no output)' }
    $body = Limit-Output $body ([int]$ConfigValue.max_output_chars) $script:TruncatedShell
    if ($null -ne $result.Code -and $result.Code -ne 0) { $body += "`n[exit code: $($result.Code)]" }
    if ($result.Note) { $body += "`n$($result.Note)" }
    return $body
}

function Get-TextFileInfo([string]$Path) {
    $raw = [IO.File]::ReadAllBytes($Path); $encoding = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false; $bom = $false; $offset = 0
    if ($raw.Length -ge 3 -and $raw[0] -eq 0xef -and $raw[1] -eq 0xbb -and $raw[2] -eq 0xbf) { $encoding = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false; $bom=$true; $offset=3 }
    elseif ($raw.Length -ge 2 -and $raw[0] -eq 0xff -and $raw[1] -eq 0xfe) { $encoding = New-Object -TypeName System.Text.UnicodeEncoding -ArgumentList $false,$true; $bom=$true; $offset=2 }
    elseif ($raw.Length -ge 2 -and $raw[0] -eq 0xfe -and $raw[1] -eq 0xff) { $encoding = New-Object -TypeName System.Text.UnicodeEncoding -ArgumentList $true,$true; $bom=$true; $offset=2 }
    else {
        try { $strict = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false,$true; $text = $strict.GetString($raw); $encoding = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false }
        catch { $encoding = [Text.Encoding]::Default }
    }
    $text = $encoding.GetString($raw,$offset,$raw.Length-$offset)
    $eol = if ($text.Contains("`r`n")) { "`r`n" } elseif ($text.Contains("`r")) { "`r" } else { "`n" }
    return [pscustomobject]@{ Path=$Path; Text=$text; Normalized=$text.Replace("`r`n", "`n").Replace("`r", "`n"); Encoding=$encoding; Bom=$bom; Eol=$eol }
}

function Save-TextFileInfo([object]$Info, [string]$NormalizedText) {
    $text = if ($Info.Eol -eq "`n") { $NormalizedText } else { $NormalizedText.Replace("`n",$Info.Eol) }
    $bytes = $Info.Encoding.GetBytes($text)
    if ($Info.Bom) { $preamble = $Info.Encoding.GetPreamble(); $bytes = $preamble + $bytes }
    [IO.File]::WriteAllBytes($Info.Path, $bytes)
}

function Require-AbsolutePath([string]$Path) {
    if (-not $Path -or -not [IO.Path]::IsPathRooted($Path)) { throw "The path $Path is not an absolute path." }
    return [IO.Path]::GetFullPath($Path)
}

function Invoke-EditorView([string]$Path, [object]$ViewRange, [int]$MaxChars) {
    $target = Require-AbsolutePath $Path
    if (-not (Test-Path -LiteralPath $target)) { throw "The path $target does not exist." }
    if (Test-Path -LiteralPath $target -PathType Container) {
        $rows = New-Object -TypeName System.Collections.Generic.List[string]; [void]$rows.Add("d`t$target")
        function Visit-EditorDirectory([string]$Dir,[int]$Depth) {
            foreach ($item in @(Get-ChildItem -LiteralPath $Dir -Force -ErrorAction SilentlyContinue)) {
                if ($item.Name.StartsWith('.') -or $item.Name -in @('node_modules','__pycache__')) { continue }
                $kind = if ($item.PSIsContainer) {'d'} else {'f'}; [void]$rows.Add("$kind`t$($item.FullName)")
                if ($item.PSIsContainer -and $Depth -lt 2) { Visit-EditorDirectory $item.FullName ($Depth+1) }
            }
        }
        Visit-EditorDirectory $target 1
        $listing = Limit-Output (($rows | Sort-Object | Out-String).Trim()) $MaxChars $script:TruncatedEditor
        return "Here're the files and directories up to 2 levels deep in $target, excluding hidden items, node_modules, and Python cache directories:`n$listing`n"
    }
    $info = Get-TextFileInfo $target; $all = $info.Normalized -split "`n"; $first=1; $last=$all.Count
    if ($null -ne $ViewRange) {
        $vr=@($ViewRange); if ($vr.Count -ne 2) { throw 'Invalid view_range. It should be a list of two integers.' }
        $first=[int]$vr[0]; $last=[int]$vr[1]; if ($first -lt 1 -or $first -gt $all.Count) { throw 'Invalid view_range first element.' }
        if ($last -eq -1) { $last=$all.Count } elseif ($last -lt $first -or $last -gt $all.Count) { throw 'Invalid view_range second element.' }
    }
    $selected = for ($i=$first; $i -le $last; $i++) { '{0,6}  {1}' -f $i,$all[$i-1] }
    return Limit-Output ("Here's the content of $target with line numbers (which has a total of $($all.Count) lines):`n" + ($selected -join "`n") + "`n") $MaxChars $script:TruncatedEditor
}

function Invoke-EditorCreate([string]$Path,[string]$FileText) {
    if ($null -eq $FileText) { throw 'Parameter file_text is required for create.' }
    $target=Require-AbsolutePath $Path; if (Test-Path -LiteralPath $target) { throw "File already exists at: $target." }
    $parent=Split-Path -Parent $target; if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw "Parent directory does not exist: $parent" }
    [IO.File]::WriteAllText($target,$FileText,(New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false)); return "New file created successfully at: $target"
}

function Invoke-EditorReplace([System.Collections.IDictionary]$Args) {
    if ($null -eq $Args.old_str -or $Args.old_str -eq '') { throw 'Parameter old_str is required and cannot be empty.' }
    $target=Require-AbsolutePath $Args.path; if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw "The path $target does not exist or is a directory." }
    $info=Get-TextFileInfo $target; $old=([string]$Args.old_str).Replace("`r`n","`n").Replace("`r","`n"); $new=([string]$Args.new_str).Replace("`r`n","`n").Replace("`r","`n")
    $count=([regex]::Matches($info.Normalized,[regex]::Escape($old))).Count; if ($count -eq 0) { throw "No replacement was performed, old_str did not appear verbatim in $target." }
    if ($count -gt 1) { throw "No replacement was performed. Multiple occurrences of old_str were found; ensure it is unique." }
    $updated = $info.Normalized.Replace($old,$new); Save-TextFileInfo $info $updated; return "The file $target has been edited successfully."
}

function Invoke-EditorInsert([System.Collections.IDictionary]$Args) {
    if ($null -eq $Args.insert_line -or $null -eq $Args.new_str) { throw 'insert_line and new_str are required.' }
    $target=Require-AbsolutePath $Args.path; if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw "The path $target does not exist." }
    $info=Get-TextFileInfo $target; $lines=@($info.Normalized -split "`n"); $line=[int]$Args.insert_line; if ($line -lt 0 -or $line -gt $lines.Count) { throw "insert_line must be within [0,$($lines.Count)]." }
    $insert=@(([string]$Args.new_str).Replace("`r`n","`n").Replace("`r","`n") -split "`n")
    if ($line -eq 0) { $after=@($insert + $lines) }
    elseif ($line -eq $lines.Count) { $after=@($lines + $insert) }
    else { $after=@($lines[0..($line-1)] + $insert + $lines[$line..($lines.Count-1)]) }
    Save-TextFileInfo $info ($after -join "`n"); return "The file $target has been edited successfully."
}

function Invoke-EditorTool([System.Collections.IDictionary]$Args,[int]$MaxChars) {
    switch ([string]$Args.command) {
        'view' { return Invoke-EditorView $Args.path $Args.view_range $MaxChars }
        'create' { return Invoke-EditorCreate $Args.path $Args.file_text }
        'str_replace' { return Invoke-EditorReplace $Args }
        'insert' { return Invoke-EditorInsert $Args }
        default { throw 'command must be one of: view, create, str_replace, insert' }
    }
}

$script:ToolSchemas = @(
    [ordered]@{ type='function'; function=[ordered]@{ name='pwsh'; description='Run commands in a persistent PowerShell shell. State, variables, functions and current directory persist between calls. Use native Windows paths.'; parameters=[ordered]@{ type='object'; properties=[ordered]@{ command=[ordered]@{type='string';description='PowerShell command to run'} }; required=@('command') } } },
    [ordered]@{ type='function'; function=[ordered]@{ name='str_replace_editor'; description='View, create and edit files. view supports directories and line ranges; create refuses to overwrite; str_replace requires one unique exact occurrence; insert inserts after the given line.'; parameters=[ordered]@{ type='object'; properties=[ordered]@{ command=[ordered]@{type='string';enum=@('view','create','str_replace','insert')}; path=[ordered]@{type='string';description='Absolute path'}; file_text=[ordered]@{type='string'}; insert_line=[ordered]@{type='integer'}; new_str=[ordered]@{type='string'}; old_str=[ordered]@{type='string'}; view_range=[ordered]@{type='array';items=[ordered]@{type='integer'}} }; required=@('command','path') } } }
)

function New-SystemPrompt([System.Collections.IDictionary]$ConfigValue) {
    $prompt=[string]$ConfigValue.system_prompt; if ($ConfigValue.include_env_note) { $prompt=$prompt.TrimEnd()+"`n`nRuntime: Windows; working directory: $($ConfigValue.cwd); the pwsh tool is PowerShell and keeps its state between calls." }; return $prompt
}

function New-ChatPayload([System.Collections.ArrayList]$MessageList,[System.Collections.IDictionary]$ConfigValue) {
    $payload=[ordered]@{model=$ConfigValue.model;messages=@($MessageList)}
    $payload.tools=$script:ToolSchemas; $payload.tool_choice='auto'
    if ($null -ne $ConfigValue.temperature) { $payload.temperature=$ConfigValue.temperature }
    if ([int]$ConfigValue.max_tokens -gt 0) { $payload.max_tokens=[int]$ConfigValue.max_tokens }
    if ($ConfigValue.stream) { $payload.stream=$true; if ($ConfigValue.include_usage) { $payload.stream_options=[ordered]@{include_usage=$true} } }
    if ($ConfigValue.extra_body) { foreach ($key in $ConfigValue.extra_body.Keys) { $payload[$key]=$ConfigValue.extra_body[$key] } }
    return ($payload | ConvertTo-Json -Depth 40 -Compress)
}

function Test-Cancelled([object]$Token) { return ($Token -and $Token.IsCancellationRequested) }

function Get-ResultText([object]$Value) {
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return [string]$Value }
    $textProperty = $Value.PSObject.Properties['Text']
    if ($textProperty -and $null -ne $textProperty.Value) { return [string]$textProperty.Value }
    return [string]$Value
}

function Invoke-ChatRequest([System.Collections.ArrayList]$MessageList,[System.Collections.IDictionary]$ConfigValue,[scriptblock]$OnText,[scriptblock]$OnReasoning,[scriptblock]$OnUsage,[object]$CancelToken) {
    $uri=Normalize-BaseUrl $ConfigValue.base_url; if (-not $uri) { throw '未配置 base_url，请先运行 -Setup' }
    if (-not $ConfigValue.model) { throw '未配置 model' }
    $body=New-ChatPayload $MessageList $ConfigValue; $headers=Get-ApiHeaders $ConfigValue ([bool]$ConfigValue.stream); $attempt=0; $emitted=$false
    while ($attempt -le [int]$ConfigValue.retries) {
        $attempt++
        try {
            $request=[Net.HttpWebRequest]::Create($uri); $request.Method='POST'; $request.ContentType='application/json'; $request.Accept=$headers.Accept; $request.UserAgent=$headers.'User-Agent'; $request.Timeout=[int]$ConfigValue.request_timeout*1000; $request.ReadWriteTimeout=[int]$ConfigValue.request_timeout*1000
            if ($ConfigValue.api_key) { $request.Headers['Authorization']='Bearer '+$ConfigValue.api_key }; foreach ($key in $ConfigValue.extra_headers.Keys) { $request.Headers[$key]=[string]$ConfigValue.extra_headers[$key] }
            $bytes=(New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false).GetBytes($body); $request.ContentLength=$bytes.Length; $stream=$request.GetRequestStream(); $stream.Write($bytes,0,$bytes.Length); $stream.Dispose()
            $response=$request.GetResponse(); $reader=New-Object -TypeName System.IO.StreamReader -ArgumentList $response.GetResponseStream(),(New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false)
            if (-not $ConfigValue.stream) {
                $raw=$reader.ReadToEnd(); $reader.Dispose(); $response.Dispose(); $data=$raw|ConvertFrom-Json; $errorValue=Get-ObjectValue $data 'error'; if ($errorValue) { throw ('接口错误：'+($errorValue|ConvertTo-Json -Compress)) }; $choices=@(Get-ObjectValue $data 'choices'); if ($choices.Count -eq 0) { throw '接口未返回 choices' }; $message=Get-ObjectValue $choices[0] 'message'; return [pscustomobject]@{role='assistant';content=(Get-ObjectValue $message 'content');tool_calls=(Get-ObjectValue $message 'tool_calls')}
            }
            $content=New-Object -TypeName System.Collections.Generic.List[string]; $reasoning=New-Object -TypeName System.Collections.Generic.List[string]; $calls=@{}; $finish=$null
            while ($null -ne ($line=$reader.ReadLine())) {
                if (Test-Cancelled $CancelToken) { throw '已取消' }; if (-not $line.StartsWith('data:')) { continue }; $dataText=$line.Substring(5).Trim(); if ($dataText -eq '[DONE]') { break }
                try { $chunk=$dataText|ConvertFrom-Json } catch { continue }
                $errorValue=Get-ObjectValue $chunk 'error'; if ($errorValue) { throw ('接口错误：'+($errorValue|ConvertTo-Json -Compress)) }
                $usage=Get-ObjectValue $chunk 'usage'; if ($usage -and $OnUsage) { & $OnUsage (ConvertTo-HashtableDeep $usage) | Out-Null }
                $choices=@(Get-ObjectValue $chunk 'choices'); if ($choices.Count -eq 0) { continue }; $choice=$choices[0]; $finishValue=Get-ObjectValue $choice 'finish_reason'; if ($finishValue) { $finish=$finishValue }; $delta=Get-ObjectValue $choice 'delta'
                $piece=Get-ObjectValue $delta 'content'; if ($piece) { $emitted=$true; [void]$content.Add([string]$piece); if ($OnText) { & $OnText ([string]$piece) | Out-Null } }
                $reason=Get-ObjectValue $delta 'reasoning_content'; if (-not $reason) { $reason=Get-ObjectValue $delta 'reasoning' }; if ($reason) { $emitted=$true; [void]$reasoning.Add([string]$reason); if ($OnReasoning) { & $OnReasoning ([string]$reason) | Out-Null } }
                foreach ($item in @(Get-ObjectValue $delta 'tool_calls')) { $indexValue=Get-ObjectValue $item 'index'; $idx=if ($null -eq $indexValue) {$calls.Count} else {[int]$indexValue}; if (-not $calls.ContainsKey($idx)) {$calls[$idx]=[ordered]@{id='';name='';arguments=''}}; $idValue=Get-ObjectValue $item 'id'; if ($idValue) {$calls[$idx].id=[string]$idValue}; $functionValue=Get-ObjectValue $item 'function'; $nameValue=Get-ObjectValue $functionValue 'name'; $argumentValue=Get-ObjectValue $functionValue 'arguments'; if ($nameValue) {$calls[$idx].name=[string]$nameValue}; if ($argumentValue) {$calls[$idx].arguments += [string]$argumentValue} }
            }
            $reader.Dispose(); $response.Dispose(); $toolCalls=@(); foreach ($idx in @($calls.Keys|Sort-Object)) { $slot=$calls[$idx]; $toolCalls += [ordered]@{id=$(if ($slot.id){$slot.id}else{'call_'+$idx});type='function';function=[ordered]@{name=$slot.name;arguments=$(if ($slot.arguments){$slot.arguments}else{'{}'})}} }
            return [pscustomobject]@{role='assistant';content=$(if ($content.Count -gt 0) {-join $content} else {$null});reasoning_content=$(if ($reasoning.Count -gt 0) {-join $reasoning} else {$null});tool_calls=$toolCalls;finish_reason=$finish}
        } catch [Net.WebException] {
            $response=$_.Exception.Response; $code=if($response){[int]$response.StatusCode}else{0}; $detail=if($response){Read-WebExceptionBody $_.Exception}else{$_.Exception.Message}; $errorText=if($code){Format-HttpError $code $detail}else{'网络错误：'+$detail};
            if ($attempt -gt [int]$ConfigValue.retries -or $emitted -or $errorText -notmatch 'HTTP (429|500|502|503|504)|网络错误') { throw $errorText }
        } catch {
            $errorText=$_.Exception.Message; if ($attempt -gt [int]$ConfigValue.retries -or $emitted -or $errorText -notmatch 'HTTP (429|500|502|503|504)|网络错误') { throw }
        }
        if (Test-Cancelled $CancelToken) { throw '已取消' }; Start-Sleep -Milliseconds ([int](1000 * [Math]::Min(8,[Math]::Pow(1.5,$attempt))))
        if ($OnText) { & $OnText "`n[重试 $attempt/$($ConfigValue.retries)]`n" | Out-Null }
    }
    throw '请求失败'
}

function Invoke-AgentTurn([string]$UserText,[System.Collections.IDictionary]$ConfigValue,[System.Collections.ArrayList]$MessageList,[scriptblock]$OnText,[scriptblock]$OnReasoning,[scriptblock]$OnToolStart,[scriptblock]$OnToolEnd,[scriptblock]$OnNotice,[scriptblock]$OnUsage,[object]$CancelToken) {
    [void]$MessageList.Add([ordered]@{role='user';content=$UserText}); $maxChars=[int]$ConfigValue.max_output_chars
    for ($round=0;$round -lt [int]$ConfigValue.max_tool_rounds;$round++) {
        $assistant=Invoke-ChatRequest $MessageList $ConfigValue $OnText $OnReasoning $OnUsage $CancelToken
        $toolCalls = @($assistant.tool_calls)
        $history=[ordered]@{role='assistant';content=$assistant.content}; if ($toolCalls.Count -gt 0) {$history.tool_calls=$toolCalls}; if ($null -eq $history.content -and $toolCalls.Count -eq 0) {$history.content=''}; [void]$MessageList.Add($history)
        if ($toolCalls.Count -eq 0) { return [string]$assistant.content }
        foreach ($call in $toolCalls) {
            $name=$call.function.name; $args=@{}; try {$args=ConvertTo-HashtableDeep ($call.function.arguments|ConvertFrom-Json)} catch {$args=@{}}
            $started=Get-Date; if ($OnToolStart) { & $OnToolStart $name $args | Out-Null }; $isError=$false
            try { if ($name -eq 'pwsh') {$result=Invoke-PwshTool $args $ConfigValue $(if($ConfigValue.show_live_output){$OnText}else{$null})} elseif($name -eq 'str_replace_editor'){$result=Invoke-EditorTool $args $maxChars} else {throw "Unknown tool: $name"} } catch {$result=$_.Exception.Message; $isError=$true}
            $elapsed=((Get-Date)-$started).TotalSeconds; [void]$MessageList.Add([ordered]@{role='tool';tool_call_id=$call.id;content=[string]$result}); if ($OnToolEnd) { & $OnToolEnd $name ([string]$result) $elapsed $isError | Out-Null }
        }
    }
    if ($OnNotice) { & $OnNotice "已达到最大工具调用轮数（$($ConfigValue.max_tool_rounds)），本轮停止。" | Out-Null }; return ''
}

function Get-SessionFiles { return @(Get-ChildItem -LiteralPath (Get-SessionsDirectory) -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending) }
function Save-Session([System.Collections.ArrayList]$MessageList,[System.Collections.IDictionary]$ConfigValue,[string]$Path) { if (-not $Path) {$Path=Join-Path (Get-SessionsDirectory) ('session-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.json')}; $data=[ordered]@{version=1;created=(Get-Date).ToString('o');config=[ordered]@{model=$ConfigValue.model;cwd=$ConfigValue.cwd};messages=@($MessageList)}; Save-ConfigFile $Path $data; return $Path }
function Load-Session([string]$Path) { $data=Get-Content -LiteralPath $Path -Raw -Encoding UTF8|ConvertFrom-Json; return @(ConvertTo-HashtableDeep $data.messages) }

function Show-HelpText { @'
可用命令：
  /help              显示本帮助
  /tools             列出可调用工具
  /models            扫描接口模型并切换
  /config            显示生效配置（密钥已掩码）
  /shellcheck        离线诊断 PowerShell
  /model [名称]      查看或临时切换模型
  /cwd [路径]        查看或切换工作目录并重置对话
  /clear (/new)      清空对话上下文
  /save [路径]       保存当前会话
  /resume [list|n]   恢复会话
  /exit (/quit)      退出
'@ }

function Initialize-Messages([System.Collections.IDictionary]$ConfigValue) { $script:Messages.Clear(); [void]$script:Messages.Add([ordered]@{role='system';content=(New-SystemPrompt $ConfigValue)}) }

function Invoke-ShellCheck([System.Collections.IDictionary]$ConfigValue) {
    Write-Host "`n$('='*68)`n dsh-mini PowerShell 诊断（离线，不联网）`n$('='*68)"
    Write-Host "版本：$($script:Version) / PowerShell：$($PSVersionTable.PSVersion) / 工作目录：$($ConfigValue.cwd)"
    try { $probe=Invoke-ShellCommand "Write-Output ('DSHMINI_PROBE_OK ps=' + `$PSVersionTable.PSVersion + ' cn=中文探测正常')" $ConfigValue $null; Write-Host "√ 探针通过：$(Get-ResultText $probe)"; Write-Host "模式：$($ConfigValue.shell_mode)；实现：$($probe.Note)"; return 0 }
    catch { Write-Host "× 探针失败：$($_.Exception.Message)"; return 1 }
}

function Invoke-SelfTest([System.Collections.IDictionary]$ConfigValue) {
    $tests=New-Object -TypeName System.Collections.Generic.List[object]
    function Check([string]$Name,[bool]$Ok,[string]$Detail='') { [void]$tests.Add([pscustomobject]@{Name=$Name;Ok=$Ok;Detail=$Detail}); Write-Host ("[{0}] {1}{2}" -f $(if($Ok){'√'}else{'×'}),$Name,$(if($Detail){' - '+$Detail}else{''})) }
    Check 'Normalize base_url 裸域名' ((Normalize-BaseUrl 'https://api.example.com') -eq 'https://api.example.com/v1/chat/completions')
    Check 'Normalize base_url 自定义路径' ((Normalize-BaseUrl 'https://x.com/openai') -eq 'https://x.com/openai/chat/completions')
    Check 'models endpoint' ((Get-ModelsEndpoint 'https://x.com/v1/chat/completions') -eq 'https://x.com/v1/models')
    $tmp=Join-Path ([IO.Path]::GetTempPath()) ('dsh-mini-test-'+[guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Path $tmp|Out-Null
    try {
        $file=Join-Path $tmp 'sample.txt'; [IO.File]::WriteAllText($file,"a`r`nb`r`nc`r`n",(New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false)); Write-Host '[selftest] editor view'; $view=Invoke-EditorView $file $null 16000; Check 'editor view line numbers' ($view -match '1\s+a' -and $view -match '3\s+c'); Write-Host '[selftest] editor replace'; Invoke-EditorReplace @{path=$file;old_str='b';new_str='B'}|Out-Null; Check 'editor str_replace preserves CRLF' ([IO.File]::ReadAllText($file) -eq "a`r`nB`r`nc`r`n"); Write-Host '[selftest] editor insert'; Invoke-EditorInsert @{path=$file;insert_line=1;new_str='x'}|Out-Null; Check 'editor insert' ([IO.File]::ReadAllText($file) -match 'a`r`nx`r`nB'); Write-Host '[selftest] editor directory'; $dirView=Invoke-EditorView $tmp $null 16000; Check 'editor directory view' ($dirView -match 'sample.txt')
    } catch {
        $detail = $_.Exception.Message
        if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) { $detail += " | " + $_.InvocationInfo.PositionMessage }
        if ($_.ScriptStackTrace) { $detail += " | stack: " + $_.ScriptStackTrace.Replace([Environment]::NewLine,' > ') }
        Check 'editor functions' $false $detail
    }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    try {
        $ConfigValue.cwd = (Get-Location).Path; $ConfigValue.shell_mode = 'persistent'; Close-PersistentShell
        $first = Invoke-PersistentShell "`$global:dsh_mini_selftest = 41; Write-Output '中文测试-OK'" $ConfigValue 20000
        $second = Invoke-PersistentShell "Write-Output ('x=' + `$global:dsh_mini_selftest)" $ConfigValue 20000
        Check 'PowerShell Runspace 中文输出' ((Get-ResultText $first) -match '中文测试-OK')
        Check 'PowerShell Runspace 跨调用保留变量' ((Get-ResultText $second) -match 'x=41')
    } catch { Check 'PowerShell Runspace' $false $_.Exception.Message } finally { Close-PersistentShell }
    $failed=@($tests|Where-Object {-not $_.Ok}); Write-Host "`n自检完成：$($tests.Count) 项，通过 $($tests.Count-$failed.Count) 项，失败 $($failed.Count) 项"; Write-Output "DSH_SELFTEST_RESULT total=$($tests.Count) passed=$($tests.Count-$failed.Count) failed=$($failed.Count)"; return $(if($failed.Count){1}else{0})
}

function Invoke-Interactive([System.Collections.IDictionary]$ConfigValue) {
    if (-not $ConfigValue.api_key -or $Setup) { Invoke-Setup $ConfigValue | Out-Null }
    if (($ConfigValue.model_picker -and -not $NoPicker) -or $PickModel -or $Models) { $picked=Select-Model $ConfigValue; if($picked){$ConfigValue.model=$picked} }
    Initialize-Messages $ConfigValue; Write-Host "$($script:AppTitle) v$($script:Version)"; Write-Host "工作目录：$($ConfigValue.cwd)  模型：$($ConfigValue.model)"; Write-Host '输入 /help 查看命令，/exit 退出。'
    while ($true) {
        try { $text=Read-Host '›' } catch { break }; if ($null -eq $text) {break}; if (-not $text.Trim()){continue}; $cmd=$text.Trim()
        if ($cmd.StartsWith('/')) {
            $parts=$cmd -split '\s+',2; $name=$parts[0].ToLowerInvariant(); $arg=if($parts.Count -gt 1){$parts[1]}else{''}
            switch ($name) {
                '/exit' { return }; '/quit' { return }; '/q' { return }; '/help' {Show-HelpText;continue}; '/tools' {$script:ToolSchemas|ForEach-Object{Write-Host ("$($_.function.name) - $($_.function.description)")};continue}; '/config' {$ConfigValue.GetEnumerator()|ForEach-Object{if($_.Key -eq 'api_key'){Write-Host "api_key = $(Mask-Secret $_.Value)"}elseif($_.Key -notin @('extra_body','extra_headers')){Write-Host "$($_.Key) = $($_.Value)"}};continue}; '/shellcheck' {Invoke-ShellCheck $ConfigValue|Out-Null;continue}; '/models' {$picked=Select-Model $ConfigValue;if($picked){$ConfigValue.model=$picked};continue}; '/model' {if($arg){$ConfigValue.model=$arg};Write-Host "当前模型：$($ConfigValue.model)";continue}; '/cwd' {if($arg -and (Test-Path -LiteralPath $arg -PathType Container)){$ConfigValue.cwd=[IO.Path]::GetFullPath($arg);Close-PersistentShell;Initialize-Messages $ConfigValue};Write-Host "当前工作目录：$($ConfigValue.cwd)";continue}; '/clear' {Initialize-Messages $ConfigValue;Write-Host '对话上下文已清空。';continue}; '/new' {Initialize-Messages $ConfigValue;continue}; '/save' {Write-Host (Save-Session $script:Messages $ConfigValue $arg);continue}; '/resume' {$files=Get-SessionFiles;if($arg -in @('list','ls')){$files|ForEach-Object{Write-Host $_.FullName};continue};if($files.Count){$path=if($arg -and (Test-Path $arg)){$arg}else{$files[0].FullName};$script:Messages.Clear();foreach($m in @(Load-Session $path)){[void]$script:Messages.Add($m)};Write-Host "已恢复会话：$path"};continue}; default {Write-Host "未知命令：$name（/help 查看帮助）";continue}
            }
        }
        try { Invoke-AgentTurn $text $ConfigValue $script:Messages {param($x)Write-Host -NoNewline $x} {param($x)if($ConfigValue.show_reasoning){Write-Host -NoNewline "`n…$x"}} {param($n,$a)if(-not $QuietTools){Write-Host "`n» $n $($a|ConvertTo-Json -Compress)"}} {param($n,$r,$e,$err)if(-not $QuietTools){Write-Host "`n$(if($err){'×'}else{'√'}) $n ($([Math]::Round($e,2))s)"}} {param($x)Write-Warning $x} {param($u)if($u){Write-Host "`n[usage] $($u|ConvertTo-Json -Compress)"}} $null|Out-Null } catch { Write-Warning $_.Exception.Message }
        if ($ConfigValue.save_sessions) { Save-Session $script:Messages $ConfigValue $null|Out-Null }
    }
    Close-PersistentShell
}

function Show-GuiInput([string]$Title,[string]$Label,[string]$Current,[bool]$Secret=$false) {
    Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
    $form=New-Object -TypeName System.Windows.Forms.Form; $form.Text=$Title; $form.Width=500;$form.Height=170;$form.StartPosition='CenterParent';$form.TopMost=$true
    $label=New-Object -TypeName System.Windows.Forms.Label;$label.Text=$Label;$label.Left=12;$label.Top=12;$label.Width=450
    $box=New-Object -TypeName System.Windows.Forms.TextBox;$box.Text=$Current;$box.Left=12;$box.Top=35;$box.Width=455;if($Secret){$box.UseSystemPasswordChar=$true}
    $ok=New-Object -TypeName System.Windows.Forms.Button;$ok.Text='确定';$ok.Left=300;$ok.Top=78;$ok.Width=80;$ok.DialogResult='OK';$cancel=New-Object -TypeName System.Windows.Forms.Button;$cancel.Text='取消';$cancel.Left=390;$cancel.Top=78;$cancel.Width=80;$cancel.DialogResult='Cancel'
    $form.Controls.AddRange(@($label,$box,$ok,$cancel));$form.AcceptButton=$ok;$form.CancelButton=$cancel;$form.Add_Shown({$box.Focus();$box.SelectAll()})
    if($form.ShowDialog() -eq 'OK'){return $box.Text};return $null
}

function Start-Gui([System.Collections.IDictionary]$ConfigValue) {
    Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
    if (-not $ConfigValue.api_key -or $Setup) { $b=Show-GuiInput $script:AppTitle '接口地址 base_url' $ConfigValue.base_url; if($b){$ConfigValue.base_url=$b};$k=Show-GuiInput $script:AppTitle 'API Key' $ConfigValue.api_key $true;if($k){$ConfigValue.api_key=$k};$m=Show-GuiInput $script:AppTitle '模型名 model' $ConfigValue.model;if($m){$ConfigValue.model=$m};Save-ConfigFile (Get-ConfigPath $Config) $ConfigValue }
    Initialize-Messages $ConfigValue
    $form=New-Object -TypeName System.Windows.Forms.Form;$form.Text="$($script:AppTitle) v$($script:Version)";$form.Width=950;$form.Height=700;$form.StartPosition='CenterScreen';$form.KeyPreview=$true
    $output=New-Object -TypeName System.Windows.Forms.RichTextBox;$output.Dock='Fill';$output.ReadOnly=$true;$output.Font=New-Object -TypeName System.Drawing.Font -ArgumentList 'Microsoft YaHei',10;$output.BackColor=[System.Drawing.Color]::White
    $panel=New-Object -TypeName System.Windows.Forms.Panel;$panel.Dock='Bottom';$panel.Height=92
    $inputBox=New-Object -TypeName System.Windows.Forms.TextBox;$inputBox.Multiline=$true;$inputBox.Left=8;$inputBox.Top=8;$inputBox.Width=730;$inputBox.Height=58;$inputBox.Font=$output.Font
    $send=New-Object -TypeName System.Windows.Forms.Button;$send.Text='发送';$send.Left=750;$send.Top=8;$send.Width=80;$send.Height=28;$stop=New-Object -TypeName System.Windows.Forms.Button;$stop.Text='停止';$stop.Left=840;$stop.Top=8;$stop.Width=80;$stop.Height=28;$stop.Enabled=$false
    $status=New-Object -TypeName System.Windows.Forms.Label;$status.Text='就绪';$status.Left=8;$status.Top=70;$status.Width=900
    $panel.Controls.AddRange(@($inputBox,$send,$stop,$status));$form.Controls.AddRange(@($output,$panel))
    $menu=New-Object -TypeName System.Windows.Forms.MenuStrip; $items=@{}
    function Add-Menu([string]$Title,[string]$Text,[scriptblock]$Action) {$item=New-Object -TypeName System.Windows.Forms.ToolStripMenuItem -ArgumentList $Text;$item.Add_Click($Action.GetNewClosure());$items[$Title]=$item;return $item}
    $session=New-Object -TypeName System.Windows.Forms.ToolStripMenuItem -ArgumentList '会话';$session.DropDownItems.Add((Add-Menu 'new' '新对话' {Initialize-Messages $ConfigValue;$output.Clear()}))|Out-Null;$session.DropDownItems.Add((Add-Menu 'save' '保存会话' {Save-Session $script:Messages $ConfigValue $null|Out-Null;$status.Text='会话已保存'}))|Out-Null;$session.DropDownItems.Add((Add-Menu 'exit' '退出' {$form.Close()}))|Out-Null
    $tools=New-Object -TypeName System.Windows.Forms.ToolStripMenuItem -ArgumentList '工具';$tools.DropDownItems.Add((Add-Menu 'selftest' '离线自检' {Invoke-SelfTest $ConfigValue|Out-Null}))|Out-Null;$tools.DropDownItems.Add((Add-Menu 'shellcheck' 'PowerShell 诊断' {Invoke-ShellCheck $ConfigValue|Out-Null}))|Out-Null;$tools.DropDownItems.Add((Add-Menu 'copy' '复制全部' {$output.SelectAll();$output.Copy();$output.DeselectAll()}))|Out-Null
    $setupAction={ $b=Show-GuiInput $script:AppTitle '接口地址 base_url' $ConfigValue.base_url;if($b){$ConfigValue.base_url=$b};$k=Show-GuiInput $script:AppTitle 'API Key' $ConfigValue.api_key $true;if($k){$ConfigValue.api_key=$k};$m=Show-GuiInput $script:AppTitle '模型名 model' $ConfigValue.model;if($m){$ConfigValue.model=$m};Save-ConfigFile (Get-ConfigPath $Config) $ConfigValue;Initialize-Messages $ConfigValue;$status.Text='设置已保存' }.GetNewClosure()
    $modelAction={ try {$picked=Select-Model $ConfigValue;if($picked){$ConfigValue.model=$picked;Save-ConfigFile (Get-ConfigPath $Config) $ConfigValue;$status.Text="模型已切换：$picked"}} catch {$status.Text='模型选择失败'} }.GetNewClosure()
    $settings=New-Object -TypeName System.Windows.Forms.ToolStripMenuItem -ArgumentList '设置';$settings.DropDownItems.Add((Add-Menu 'models' '选择模型' $modelAction))|Out-Null;$settings.DropDownItems.Add((Add-Menu 'setup' '接口与密钥' $setupAction))|Out-Null;$settings.DropDownItems.Add((Add-Menu 'reasoning' '显示思考过程' {$ConfigValue.show_reasoning=-not [bool]$ConfigValue.show_reasoning;$status.Text="显示思考过程：$($ConfigValue.show_reasoning)"}))|Out-Null
    $help=New-Object -TypeName System.Windows.Forms.ToolStripMenuItem -ArgumentList '帮助';$help.DropDownItems.Add((Add-Menu 'guide' '操作指南' {[System.Windows.Forms.MessageBox]::Show((Show-HelpText),$script:AppTitle)}))|Out-Null
    $menu.Items.AddRange(@($session,$tools,$settings,$help));$form.MainMenuStrip=$menu;$form.Controls.Add($menu)
    $append = { param([string]$Text) $output.AppendText($Text);$output.SelectionStart=$output.TextLength;$output.ScrollToCaret();[IO.File]::AppendAllText((Join-Path (Get-DataRoot) $script:GuiLogName),$Text,(New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false)) }.GetNewClosure()
    & $append "$($script:AppTitle) v$($script:Version)`r`n工作目录：$($ConfigValue.cwd)`r`n模型：$($ConfigValue.model)`r`n`r`n"
    $worker=New-Object -TypeName System.ComponentModel.BackgroundWorker
    $sendAction={if($worker.IsBusy){return};$text=$inputBox.Text.Trim();if(-not $text){return};$inputBox.Clear();&$append "`r`n› $text`r`n";$send.Enabled=$false;$stop.Enabled=$true;$status.Text='等待模型响应…';$script:CurrentCancel=New-Object System.Threading.CancellationTokenSource;$worker.RunWorkerAsync($text)}.GetNewClosure();$send.Add_Click($sendAction);$inputBox.Add_KeyDown({param($s,$e)if($e.KeyCode -eq 'Enter' -and -not $e.Shift){$e.SuppressKeyPress=$true;&$sendAction}}.GetNewClosure());$stop.Add_Click({if($script:CurrentCancel){$script:CurrentCancel.Cancel();$status.Text='正在取消…'}}.GetNewClosure())
    $worker.Add_DoWork({param($s,$e)$text=[string]$e.Argument;$e.Result=Invoke-AgentTurn $text $ConfigValue $script:Messages {param($x)$form.BeginInvoke([Action]{&$append $x})|Out-Null} {param($x)if($ConfigValue.show_reasoning){$form.BeginInvoke([Action]{&$append "`r`n…$x"})|Out-Null}} {param($n,$a)$form.BeginInvoke([Action]{&$append "`r`n» $n`r`n"})|Out-Null} {param($n,$r,$elapsed,$err)$form.BeginInvoke([Action]{&$append "`r`n$(if($err){'×'}else{'√'}) $n ($([Math]::Round($elapsed,2))s)`r`n"})|Out-Null} {param($x)$form.BeginInvoke([Action]{[System.Windows.Forms.MessageBox]::Show($x,$script:AppTitle)})|Out-Null} {param($u)} $script:CurrentCancel.Token}.GetNewClosure())
    $worker.Add_RunWorkerCompleted({param($s,$e)$send.Enabled=$true;$stop.Enabled=$false;$status.Text=if($e.Error){'错误：'+$e.Error.Message}else{'就绪'};if($e.Error){&$append "`r`n错误：$($e.Error.Message)`r`n"};if($ConfigValue.save_sessions){Save-Session $script:Messages $ConfigValue $null|Out-Null};$inputBox.Focus()}.GetNewClosure())
    $form.Add_FormClosed({Close-PersistentShell}.GetNewClosure());[void]$form.ShowDialog()
}

function Main {
    if ($Version) { Write-Output "$($script:AppName) $($script:Version)"; return 0 }
    $configValue=Get-EffectiveConfig; $script:ConfigData=$configValue
    if ($SelfTest) { return (Invoke-SelfTest $configValue) }
    if ($ShellCheck) { return (Invoke-ShellCheck $configValue) }
    if ($Gui -or (-not $Cli -and [Environment]::UserInteractive -and $env:OS -eq 'Windows_NT')) { Start-Gui $configValue; return 0 }
    if ($Prompt) { if(-not $configValue.api_key -and $Setup){Invoke-Setup $configValue|Out-Null};Initialize-Messages $configValue;try{Invoke-AgentTurn $Prompt $configValue $script:Messages {param($x)Write-Host -NoNewline $x} {param($x)if($configValue.show_reasoning){Write-Host -NoNewline "…$x"}} {param($n,$a)if(-not $QuietTools){Write-Host "[tool] $n"}} {param($n,$r,$e,$err)if(-not $QuietTools){Write-Host "[tool] done $n"}} {param($x)Write-Warning $x} {param($u)} $null|Out-Null;Write-Host '' ;return 0}catch{$message=$_.Exception.Message;Write-Error -Message $message -ErrorAction Continue;return 2} }
    Invoke-Interactive $configValue; return 0
}

    try { $exitCode=Main; return $exitCode } catch { $message=$_.Exception.Message; Write-Error -Message $message -ErrorAction Continue; Close-PersistentShell; return 2 }
}

try {
    Invoke-DshMiniScript @args | ForEach-Object {
        if ($_ -isnot [int]) { Write-Output $_ }
    }
} finally {
    Remove-Item -LiteralPath Function:\Invoke-DshMiniScript -Force -ErrorAction SilentlyContinue
}
