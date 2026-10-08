# Generate a JMeter 5.6.3 test plan on Windows 11.
# Requires Windows PowerShell 5.1 only. Excel and Python are not required.
# Save this file as UTF-8 with BOM so Windows PowerShell reads Japanese text.

[CmdletBinding()]
param(
    [string]$InputPath,
    [string]$OutputDir,
    [switch]$Run,
    [string]$JMeterPath,
    [int]$Threads,
    [int]$Loops,
    [int]$RampUp,
    [switch]$InitTemplates
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$Script:JMeterVersion = "5.6.3"
$Script:JmxFileName = "rest-api-test.jmx"
$Script:UsageFileName = "jmeter-usage.xlsx"
$Script:DefaultPlanName = "REST API Test Plan"
$Script:DefaultJtl = "results/listener-results.jtl"

function Show-GeneratorHelp {
    @"
usage: generate-jmeter.ps1 -InputPath FILE -OutputDir DIR [options]
       generate-jmeter.ps1 -InitTemplates -OutputDir DIR

Excel (.xlsx) または CSV から、Apache JMeter $($Script:JMeterVersion) の JMX と利用手順 Excel を生成します。
初期値では有効な各 API をスレッド 1、ループ 1 で 1 回呼び出します。

options:
  -InputPath FILE     入力 .xlsx または .csv
  -OutputDir DIR      出力ディレクトリ
  -Run                生成後に JMeter を非GUI実行し、JTL、jmeter.log、HTMLレポートを書く
  -JMeterPath PATH    jmeter.bat のパス。未指定時は JMETER_HOME と PATH
  -Threads N          スレッド数（未指定時は入力値、さらに無ければ 1）
  -Loops N            ループ回数（未指定時は入力値、さらに無ければ 1）
  -RampUp N           ランプアップ秒
  -InitTemplates      入力テンプレートを -OutputDir に書き出す
"@ | Write-Output
}

function ConvertTo-Token {
    param($Value)
    if ($null -eq $Value) { return "" }
    $text = [string]$Value
    $text = $text.Normalize([System.Text.NormalizationForm]::FormKC).Trim().ToLowerInvariant()
    return $text.Replace("-", "_").Replace(" ", "").Replace([string][char]0x3000, "")
}

function Test-IllegalXml {
    param([string]$Value)
    return $Value -match "[\x00-\x08\x0B\x0C\x0E-\x1F]"
}

function ConvertTo-XmlText {
    param($Value)
    if ($null -eq $Value) { $Value = "" }
    $text = [string]$Value
    if (Test-IllegalXml $text) { throw "使用できない制御文字が含まれています。" }
    return $text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace('"', "&quot;").Replace("'", "&apos;")
}

function ConvertTo-BoolValue {
    param($Value, [bool]$Default = $true, [string]$Label = "有効")
    if ($null -eq $Value) { return $Default }
    $text = ([string]$Value).Trim()
    if ($text -eq "") { return $Default }
    $token = ConvertTo-Token $text
    $trueSet = @("1", "true", "yes", "y", "on", "はい", "有効", "○", "◯")
    $falseSet = @("0", "false", "no", "n", "off", "いいえ", "無効", "×", "x")
    if ($trueSet -contains $token) { return $true }
    if ($falseSet -contains $token) { return $false }
    throw "$Label の値を解釈できません: $Value"
}

function ConvertTo-IntRange {
    param($Value, [string]$Label, [int]$Minimum, [int]$Maximum)
    $text = ([string]$Value).Trim()
    if ($text -notmatch "^[0-9]+$") { throw "$Label は整数で指定してください: $Value" }
    $number = [int]$text
    if ($number -lt $Minimum -or $number -gt $Maximum) {
        throw "$Label は $Minimum 以上 $Maximum 以下にしてください: $number"
    }
    return $number
}

function Test-SecretHeader {
    param([string]$Name)
    return $Name -match "(?i)(authorization|api[-_]?key|token|secret|password|passwd|cookie|credential)"
}

function Get-MaskedSecret {
    param([string]$Name, [string]$Value)
    if (Test-SecretHeader $Name) { return "********" }
    return $Value
}

function Test-HeaderName {
    param($Name)
    $text = ([string]$Name).Trim()
    if ($text -eq "" -or $text -notmatch '^[!#$%&''*+\-.0-9A-Z^_`a-z|~]+$') {
        throw "ヘッダ名が不正です: $Name"
    }
    return $text
}

function Test-HeaderValue {
    param($Value)
    $text = ([string]$Value).Trim()
    if ($text -match "[\r\n]") { throw "ヘッダ値に改行は使えません。" }
    if (Test-IllegalXml $text) { throw "ヘッダ値 に使用できない制御文字が含まれています。" }
    return $text
}

function Test-ParamName {
    param($Name, [string]$Label)
    $text = ([string]$Name).Trim()
    if ($text -eq "") { throw "$Label が空です。" }
    if ($text -match "[\r\n]") { throw "$Label に改行は使えません。" }
    if (Test-IllegalXml $text) { throw "$Label に使用できない制御文字が含まれています。" }
    return $text
}

function Test-RelativeJtl {
    param($Value)
    $text = ([string]$Value).Trim()
    if ($text -eq "") { throw "jtl_file が空です。" }
    if ([System.IO.Path]::IsPathRooted($text) -or $text -match "^[A-Za-z]:") {
        throw "jtl_file は出力ディレクトリ内の相対パスにしてください: $text"
    }
    $parts = $text -split "[\\/]"
    foreach ($part in $parts) {
        if ($part -eq "" -or $part -eq "..") { throw "jtl_file が不正です: $text" }
    }
    if (Test-IllegalXml $text) { throw "jtl_file に使用できない制御文字が含まれています。" }
    return $text.Replace("\", "/")
}

function ConvertTo-Percent {
    param([string]$Text, [string]$Safe)
    if ($null -eq $Text) { $Text = "" }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $builder = New-Object System.Text.StringBuilder
    foreach ($byte in $bytes) {
        $keep = (($byte -ge 48 -and $byte -le 57) -or ($byte -ge 65 -and $byte -le 90) -or ($byte -ge 97 -and $byte -le 122))
        if (-not $keep -and $byte -lt 128) {
            $keep = $Safe.IndexOf([char]$byte) -ge 0
        }
        if ($keep) {
            [void]$builder.Append([char]$byte)
        } else {
            [void]$builder.AppendFormat("%{0:X2}", $byte)
        }
    }
    return $builder.ToString()
}

function ConvertFrom-QueryString {
    param([string]$Raw, [string]$Label)
    Add-Type -AssemblyName System.Web -ErrorAction Stop
    $pairs = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Raw)) { return ,@() }
    foreach ($part in $Raw.Split("&")) {
        if ($part -eq "") { throw "$Label は key=value&key2=value2 の形式で指定してください。" }
        $eq = $part.IndexOf("=")
        if ($eq -lt 0) { throw "$Label は key=value&key2=value2 の形式で指定してください。" }
        $key = [System.Web.HttpUtility]::UrlDecode($part.Substring(0, $eq))
        $value = [System.Web.HttpUtility]::UrlDecode($part.Substring($eq + 1))
        $pairs.Add((New-Pair (Test-ParamName $key $Label) $value))
    }
    return ,$pairs.ToArray()
}

function Split-HeaderBlob {
    param([string]$Raw)
    $pairs = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Raw)) { return ,@() }
    foreach ($part in ($Raw -split "[|\r\n]+")) {
        $item = $part.Trim()
        if ($item -eq "") { continue }
        $colon = $item.IndexOf(":")
        if ($colon -lt 0) { throw "ヘッダは '名前: 値' です。複数指定するときは | か改行で区切ってください: $item" }
        $name = Test-HeaderName $item.Substring(0, $colon)
        $value = Test-HeaderValue $item.Substring($colon + 1)
        $pairs.Add((New-Pair $name $value))
    }
    return ,$pairs.ToArray()
}

function Box-Items {
    param($Value)
    $copy = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Value) { return ,$copy.ToArray() }
    $countProp = $Value.PSObject.Properties["Count"]
    $indexable = $null -ne $countProp -and $Value -isnot [string] -and $Value -isnot [pscustomobject]
    if ($indexable) {
        for ($i = 0; $i -lt $Value.Count; $i++) { $copy.Add($Value[$i]) }
        return ,$copy.ToArray()
    }
    $copy.Add($Value)
    return ,$copy.ToArray()
}

function Get-ItemCount {
    param($Value)
    if ($null -eq $Value) { return 0 }
    if ($Value -is [string]) { return 1 }
    if ($Value -is [System.Collections.ICollection]) { return $Value.Count }
    return 1
}

function Unwrap-Records {
    param($Value)
    $copy = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Value) { return $copy }
    foreach ($item in $Value) {
        $expand = $item -is [System.Collections.IEnumerable] -and -not ($item -is [string]) -and -not ($item -is [pscustomobject])
        if ($expand) {
            foreach ($inner in $item) { $copy.Add($inner) }
        } else {
            $copy.Add($item)
        }
    }
    return $copy
}

function Copy-Items {
    param($Value)
    $copy = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Value) { return ,$copy.ToArray() }
    if ($Value -is [string] -or -not ($Value -is [System.Collections.IEnumerable])) {
        $copy.Add($Value)
        return ,$copy.ToArray()
    }
    foreach ($item in $Value) { $copy.Add($item) }
    return ,$copy.ToArray()
}

function New-Pair {
    param($Key, $Value)
    return [pscustomobject]@{ Key = [string]$Key; Value = [string]$Value }
}

function New-ApiObject {
    param($Name, $Method, $Url, $Body, $ContentType, $BodyMode, $TimeoutMs, $ExpectedStatus, $Comment, [bool]$Enabled, $RowLabel)
    $api = [pscustomobject]@{
        Name = [string]$Name
        Method = [string]$Method
        Url = [string]$Url
        Body = [string]$Body
        ContentType = [string]$ContentType
        BodyMode = [string]$BodyMode
        TimeoutMs = [string]$TimeoutMs
        ExpectedStatus = [string]$ExpectedStatus
        Comment = [string]$Comment
        Enabled = [bool]$Enabled
        RowLabel = [string]$RowLabel
        Query = (New-Object System.Collections.Generic.List[object])
        Headers = (New-Object System.Collections.Generic.List[object])
        FormParams = (New-Object System.Collections.Generic.List[object])
    }
    return $api
}

function New-SettingsObject {
    return [pscustomobject]@{
        Threads = 1
        Loops = 1
        RampUpSec = 1
        FollowRedirects = $true
        PlanName = $Script:DefaultPlanName
        OnError = "continue"
        JtlFile = $Script:DefaultJtl
    }
}

function Get-RoleTable {
    param([string]$Kind)
    $make = {
        param($Role, $Aliases)
        return [pscustomobject]@{ Role = $Role; Aliases = $Aliases }
    }
    if ($Kind -eq "structured") {
        return @(
            (& $make "record_type" @("record_type", "レコード種別", "種別")),
            (& $make "api_name" @("api_name", "apiname", "api名", "名前", "名称", "name")),
            (& $make "enabled" @("enabled", "enable", "有効")),
            (& $make "method" @("method", "http_method", "メソッド", "httpメソッド")),
            (& $make "url" @("url", "uri", "endpoint", "アドレス")),
            (& $make "body" @("body", "request_body", "ボディ", "リクエストボディ")),
            (& $make "content_type" @("content_type", "contenttype", "コンテントタイプ")),
            (& $make "body_mode" @("body_mode", "bodymode", "ボディ形式")),
            (& $make "timeout_ms" @("timeout_ms", "timeout", "タイムアウトms", "タイムアウト")),
            (& $make "expected_status" @("expected_status", "期待ステータス", "期待するステータス")),
            (& $make "key" @("key", "キー", "パラメータ名", "ヘッダ名")),
            (& $make "value" @("value", "値")),
            (& $make "comment" @("comment", "comments", "コメント", "説明", "備考"))
        )
    }
    if ($Kind -eq "api") {
        return @(
            (& $make "enabled" @("enabled", "enable", "有効")),
            (& $make "name" @("name", "api_name", "apiname", "名前", "名称", "api名")),
            (& $make "method" @("method", "http_method", "メソッド", "httpメソッド")),
            (& $make "url" @("url", "uri", "endpoint", "アドレス")),
            (& $make "body" @("body", "request_body", "ボディ", "リクエストボディ")),
            (& $make "content_type" @("content_type", "contenttype", "コンテントタイプ")),
            (& $make "body_mode" @("body_mode", "bodymode", "ボディ形式")),
            (& $make "timeout_ms" @("timeout_ms", "timeout", "タイムアウトms", "タイムアウト")),
            (& $make "expected_status" @("expected_status", "期待ステータス", "期待するステータス")),
            (& $make "comment" @("comment", "comments", "コメント", "説明", "備考"))
        )
    }
    if ($Kind -eq "child") {
        return @(
            (& $make "api_name" @("api_name", "apiname", "api名", "対象api", "対象")),
            (& $make "key" @("key", "param_name", "name", "名前", "キー", "パラメータ名", "ヘッダ名")),
            (& $make "value" @("value", "param_value", "値")),
            (& $make "enabled" @("enabled", "enable", "有効"))
        )
    }
    if ($Kind -eq "settings") {
        return @(
            (& $make "key" @("key", "name", "名前", "キー", "項目")),
            (& $make "value" @("value", "値", "内容")),
            (& $make "comment" @("comment", "コメント", "説明", "備考"))
        )
    }
    return @(
        (& $make "name" @("name", "api_name", "名前", "名称", "api名")),
        (& $make "method" @("method", "http_method", "メソッド", "httpメソッド")),
        (& $make "url" @("url", "uri", "endpoint", "アドレス")),
        (& $make "query" @("query", "query_params", "クエリ", "クエリパラメータ")),
        (& $make "headers" @("headers", "header", "ヘッダ", "ヘッダー", "ヘッダ情報")),
        (& $make "body" @("body", "request_body", "ボディ", "リクエストボディ")),
        (& $make "form" @("form", "form_params", "フォーム", "postパラメータ")),
        (& $make "content_type" @("content_type", "contenttype", "コンテントタイプ")),
        (& $make "body_mode" @("body_mode", "bodymode", "ボディ形式")),
        (& $make "timeout_ms" @("timeout_ms", "timeout", "タイムアウトms", "タイムアウト")),
        (& $make "expected_status" @("expected_status", "期待ステータス", "期待するステータス")),
        (& $make "comment" @("comment", "comments", "コメント", "説明", "備考")),
        (& $make "enabled" @("enabled", "enable", "有効"))
    )
}

function Get-HeaderMap {
    param([object[]]$Headers, [object[]]$Roles)
    $normalized = @()
    foreach ($header in @($Headers)) { $normalized += (ConvertTo-Token $header) }
    $used = @{}
    $found = @{}
    foreach ($role in @($Roles)) {
        $chosen = $null
        foreach ($alias in @($role.Aliases)) {
            $token = ConvertTo-Token $alias
            for ($i = 0; $i -lt $normalized.Count; $i++) {
                if ($used.ContainsKey([string]$i)) { continue }
                if ($normalized[$i] -eq $token) { $chosen = $i; break }
            }
            if ($null -ne $chosen) { break }
        }
        if ($null -ne $chosen) {
            $used[[string]$chosen] = $true
            $found[$role.Role] = $chosen
        }
    }
    return $found
}

function Get-RowValue {
    param([object[]]$Row, [hashtable]$Map, [string]$Role)
    if (-not $Map.ContainsKey($Role)) { return "" }
    $index = [int]$Map[$Role]
    $items = @($Row)
    if ($index -ge $items.Count -or $null -eq $items[$index]) { return "" }
    return [string]$items[$index]
}

function Test-BlankRow {
    param([object[]]$Row)
    foreach ($cell in @($Row)) {
        if (([string]$cell).Trim() -ne "") { return $false }
    }
    return $true
}

function Get-RecordKind {
    param([string]$Value, [string]$Label)
    $token = ConvertTo-Token $Value
    if ($token -eq "") { throw "${Label}: record_type が空です。" }
    $map = @{
        "api" = "API"; "query" = "QUERY"; "header" = "HEADER"; "body" = "BODY"; "form" = "BODY"; "setting" = "SETTING"
        "クエリ" = "QUERY"; "ヘッダ" = "HEADER"; "ヘッダー" = "HEADER"; "フォーム" = "BODY"; "ボディ" = "BODY"; "設定" = "SETTING"
    }
    if (-not $map.ContainsKey($token)) { throw "${Label}: record_type が不正です: $Value" }
    return $map[$token]
}

function New-BlankRecord {
    return [pscustomobject]@{
        Kind = ""; ApiName = ""; Enabled = ""; Method = ""; Url = ""; Body = ""; ContentType = ""
        BodyMode = ""; TimeoutMs = ""; ExpectedStatus = ""; Key = ""; Value = ""; Comment = ""; RowLabel = ""
    }
}

function ConvertTo-Record {
    param([hashtable]$Map, [object[]]$Row, [string]$RowLabel, [string]$Kind)
    $record = New-BlankRecord
    $record.Kind = $Kind
    $record.RowLabel = $RowLabel
    $record.ApiName = Get-RowValue $Row $Map "api_name"
    if ($record.ApiName -eq "" -and $Map.ContainsKey("name")) { $record.ApiName = Get-RowValue $Row $Map "name" }
    $record.Enabled = Get-RowValue $Row $Map "enabled"
    $record.Method = Get-RowValue $Row $Map "method"
    $record.Url = Get-RowValue $Row $Map "url"
    $record.Body = Get-RowValue $Row $Map "body"
    $record.ContentType = Get-RowValue $Row $Map "content_type"
    $record.BodyMode = Get-RowValue $Row $Map "body_mode"
    $record.TimeoutMs = Get-RowValue $Row $Map "timeout_ms"
    $record.ExpectedStatus = Get-RowValue $Row $Map "expected_status"
    $record.Key = Get-RowValue $Row $Map "key"
    $record.Value = Get-RowValue $Row $Map "value"
    $record.Comment = Get-RowValue $Row $Map "comment"
    return $record
}

function Read-CsvTable {
    param([string]$Text)
    $rows = New-Object System.Collections.Generic.List[object]
    $field = New-Object System.Text.StringBuilder
    $row = New-Object System.Collections.Generic.List[string]
    $inQuotes = $false
    $chars = $Text.ToCharArray()
    $i = 0
    while ($i -lt $chars.Length) {
        $ch = $chars[$i]
        if ($inQuotes) {
            if ($ch -eq '"') {
                if (($i + 1) -lt $chars.Length -and $chars[$i + 1] -eq '"') {
                    [void]$field.Append('"')
                    $i += 2
                    continue
                }
                $inQuotes = $false
                $i++
                continue
            }
            [void]$field.Append($ch)
            $i++
            continue
        }
        if ($ch -eq '"') { $inQuotes = $true; $i++; continue }
        if ($ch -eq ',') { $row.Add($field.ToString()); $field.Length = 0; $i++; continue }
        if ($ch -eq "`n" -or $ch -eq "`r") {
            if ($ch -eq "`r" -and ($i + 1) -lt $chars.Length -and $chars[$i + 1] -eq "`n") { $i++ }
            $row.Add($field.ToString())
            $field.Length = 0
            $rows.Add(@($row.ToArray()))
            $row.Clear()
            $i++
            continue
        }
        [void]$field.Append($ch)
        $i++
    }
    if ($field.Length -gt 0 -or $row.Count -gt 0) {
        $row.Add($field.ToString())
        $rows.Add(@($row.ToArray()))
    }
    return ,@($rows.ToArray())
}

function Read-TextAuto {
    param([string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $utf8 = New-Object System.Text.UTF8Encoding $false, $true
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) {
        return $utf8.GetString($bytes, 3, $bytes.Length - 3)
    }
    $decoded = $null
    try { $decoded = $utf8.GetString($bytes) } catch { $decoded = $null }
    if ($null -ne $decoded -and $decoded -match ",") { return $decoded }
    $cp932 = [System.Text.Encoding]::GetEncoding(932)
    return $cp932.GetString($bytes)
}

function Get-OnErrorValue {
    param($Value)
    $raw = ([string]$Value).Trim()
    $map = @{
        "continue" = "continue"; "startnextloop" = "startnextloop"; "stopthread" = "stopthread"
        "stoptest" = "stoptest"; "stoptestnow" = "stoptestnow"; "続行" = "continue"; "停止" = "stoptest"
    }
    if ($map.ContainsKey($raw)) { return $map[$raw] }
    $token = ConvertTo-Token $raw
    if ($map.ContainsKey($token)) { return $map[$token] }
    throw "on_error の値を解釈できません: $Value"
}

function Set-PlanSetting {
    param($Settings, [hashtable]$Seen, $Record)
    $rawKey = ([string]$Record.Key).Trim()
    $label = $Record.RowLabel
    if ($rawKey -eq "") { throw "${label}: 設定名が空です。" }
    $token = ConvertTo-Token $rawKey
    $fields = @{
        "threads" = "threads"; "thread" = "threads"; "スレッド数" = "threads"; "スレッド" = "threads"
        "loops" = "loops"; "loop" = "loops"; "ループ回数" = "loops"; "ループ" = "loops"
        "ramp_up_sec" = "ramp"; "ramp_up" = "ramp"; "rampup" = "ramp"; "ランプアップ秒" = "ramp"; "ランプアップ" = "ramp"
        "follow_redirects" = "follow"; "followredirects" = "follow"; "リダイレクト追従" = "follow"
        "plan_name" = "plan"; "planname" = "plan"; "プラン名" = "plan"
        "on_error" = "error"; "onerror" = "error"; "エラー時" = "error"
        "jtl_file" = "jtl"; "jtlファイル" = "jtl"
    }
    if (-not $fields.ContainsKey($token)) { throw "${label}: 未知の設定です: $rawKey" }
    $fieldName = $fields[$token]
    if ($Seen.ContainsKey($fieldName)) { throw "設定が重複しています: $rawKey" }
    $Seen[$fieldName] = $true
    $value = $Record.Value
    if ($fieldName -eq "threads") { $Settings.Threads = ConvertTo-IntRange $value "threads" 1 1000000 }
    elseif ($fieldName -eq "loops") { $Settings.Loops = ConvertTo-IntRange $value "loops" 1 1000000 }
    elseif ($fieldName -eq "ramp") { $Settings.RampUpSec = ConvertTo-IntRange $value "ramp_up_sec" 0 86400 }
    elseif ($fieldName -eq "follow") { $Settings.FollowRedirects = ConvertTo-BoolValue $value $true "follow_redirects" }
    elseif ($fieldName -eq "plan") {
        $text = ([string]$value).Trim()
        if ($text -eq "") { throw "plan_name が空です。" }
        if (Test-IllegalXml $text) { throw "plan_name に使用できない制御文字が含まれています。" }
        if ($text.Length -gt 200) { throw "plan_name が長すぎます。" }
        $Settings.PlanName = $text
    }
    elseif ($fieldName -eq "error") { $Settings.OnError = Get-OnErrorValue $value }
    elseif ($fieldName -eq "jtl") { $Settings.JtlFile = Test-RelativeJtl $value }
}

function ConvertTo-Plan {
    param([object[]]$Records, [string]$SourcePath)
    $apis = New-Object System.Collections.Generic.List[object]
    $index = @{}
    foreach ($record in @($Records)) {
        if ($record.Kind -ne "API") { continue }
        $label = $record.RowLabel
        $name = ([string]$record.ApiName).Trim()
        if ($name -eq "") { throw "${label}: API名が空です。" }
        if ($name.Length -gt 200) { throw "${label}: API名が長すぎます。" }
        if (Test-IllegalXml $name) { throw "${label} のAPI名 に使用できない制御文字が含まれています。" }
        if ($index.ContainsKey($name)) { throw "API名が重複しています: $name" }
        $method = ([string]$record.Method).Trim().ToUpper()
        if ($method -ne "GET" -and $method -ne "POST") { throw "${label}: メソッドは GET または POST のみです: $($record.Method)" }
        $url = ([string]$record.Url).Trim()
        if ($url -eq "") { throw "${label}: URLが空です。" }
        $bodyMode = ([string]$record.BodyMode).Trim()
        if ($bodyMode -eq "") { $bodyMode = "auto" }
        $body = [string]$record.Body
        if (Test-IllegalXml $body) { throw "$name のボディ に使用できない制御文字が含まれています。" }
        $comment = ([string]$record.Comment).Trim()
        if (Test-IllegalXml $comment) { throw "$name のコメント に使用できない制御文字が含まれています。" }
        $api = New-ApiObject $name $method $url $body ([string]$record.ContentType).Trim() $bodyMode ([string]$record.TimeoutMs).Trim() ([string]$record.ExpectedStatus).Trim() $comment (ConvertTo-BoolValue $record.Enabled $true "$label の有効") $label
        $apis.Add($api)
        $index[$name] = $api
    }
    if ($apis.Count -eq 0) { throw "APIが1件もありません。" }
    $settings = New-SettingsObject
    $seen = @{}
    foreach ($record in @($Records)) {
        if ($record.Kind -eq "API") { continue }
        $label = $record.RowLabel
        if ($record.Kind -eq "SETTING") { Set-PlanSetting $settings $seen $record; continue }
        $apiName = ([string]$record.ApiName).Trim()
        if (-not $index.ContainsKey($apiName)) { throw "${label}: 対象APIが見つかりません: $apiName" }
        if (-not (ConvertTo-BoolValue $record.Enabled $true "$label の有効")) { continue }
        $api = $index[$apiName]
        if ($record.Kind -eq "QUERY") {
            $key = Test-ParamName $record.Key "$label のクエリ名"
            $value = [string]$record.Value
            if (Test-IllegalXml $value) { throw "$label のクエリ値 に使用できない制御文字が含まれています。" }
            $api.Query.Add((New-Pair $key $value))
        } elseif ($record.Kind -eq "HEADER") {
            $api.Headers.Add((New-Pair (Test-HeaderName $record.Key) (Test-HeaderValue $record.Value)))
        } elseif ($record.Kind -eq "BODY") {
            $key = Test-ParamName $record.Key "$label のフォーム名"
            $value = [string]$record.Value
            if (Test-IllegalXml $value) { throw "$label のフォーム値 に使用できない制御文字が含まれています。" }
            $api.FormParams.Add((New-Pair $key $value))
        } else {
            throw "${label}: 未知のレコードです。"
        }
    }
    return [pscustomobject]@{
        Apis = $apis
        Settings = $settings
        Warnings = (New-Object System.Collections.Generic.List[string])
        SourcePath = $SourcePath
    }
}

function Add-FlatChildren {
    param([System.Collections.Generic.List[object]]$Records, $Values, [string]$RowLabel)
    $apiName = ([string]$Values.ApiName).Trim()
    $queries = ConvertFrom-QueryString ([string]$Values.Query) "$RowLabel のクエリ"
    foreach ($pair in @($queries)) {
        $child = New-BlankRecord
        $child.Kind = "QUERY"; $child.ApiName = $apiName; $child.Key = $pair.Key; $child.Value = $pair.Value
        $child.Enabled = "true"; $child.RowLabel = $RowLabel
        $Records.Add($child)
    }
    $forms = ConvertFrom-QueryString ([string]$Values.Form) "$RowLabel のフォーム"
    foreach ($pair in @($forms)) {
        $child = New-BlankRecord
        $child.Kind = "BODY"; $child.ApiName = $apiName; $child.Key = $pair.Key; $child.Value = $pair.Value
        $child.Enabled = "true"; $child.RowLabel = $RowLabel
        $Records.Add($child)
    }
    $headers = Split-HeaderBlob ([string]$Values.Headers)
    foreach ($pair in @($headers)) {
        $child = New-BlankRecord
        $child.Kind = "HEADER"; $child.ApiName = $apiName; $child.Key = $pair.Key; $child.Value = $pair.Value
        $child.Enabled = "true"; $child.RowLabel = $RowLabel
        $Records.Add($child)
    }
}

function ConvertFrom-Table {
    param([object[]]$LabeledRows, [string]$SourcePath, [string]$Label)
    $header = $null
    $headerIndex = -1
    $rows = @($LabeledRows)
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if (-not (Test-BlankRow $rows[$i].Row)) { $header = $rows[$i]; $headerIndex = $i; break }
    }
    if ($null -eq $header) { throw "$Label にヘッダ行がありません。" }
    $structured = Get-HeaderMap $header.Row (Get-RoleTable "structured")
    $records = New-Object System.Collections.Generic.List[object]
    if ($structured.ContainsKey("record_type")) {
        for ($i = $headerIndex + 1; $i -lt $rows.Count; $i++) {
            if (Test-BlankRow $rows[$i].Row) { continue }
            $record = ConvertTo-Record $structured $rows[$i].Row $rows[$i].Label ""
            $record.Kind = Get-RecordKind (Get-RowValue $rows[$i].Row $structured "record_type") $rows[$i].Label
            $records.Add($record)
        }
        return ConvertTo-Plan $records $SourcePath
    }
    $flat = Get-HeaderMap $header.Row (Get-RoleTable "flat")
    if ($flat.ContainsKey("method") -and $flat.ContainsKey("url") -and $flat.ContainsKey("name")) {
        for ($i = $headerIndex + 1; $i -lt $rows.Count; $i++) {
            if (Test-BlankRow $rows[$i].Row) { continue }
            $record = ConvertTo-Record $flat $rows[$i].Row $rows[$i].Label "API"
            $record.ApiName = Get-RowValue $rows[$i].Row $flat "name"
            $records.Add($record)
            $holder = [pscustomobject]@{
                ApiName = $record.ApiName
                Query = (Get-RowValue $rows[$i].Row $flat "query")
                Form = (Get-RowValue $rows[$i].Row $flat "form")
                Headers = (Get-RowValue $rows[$i].Row $flat "headers")
            }
            Add-FlatChildren $records $holder $rows[$i].Label
        }
        return ConvertTo-Plan $records $SourcePath
    }
    throw "$Label のヘッダを認識できません。record_type 列か、name,method,url 列が必要です。"
}

function ConvertFrom-CsvText {
    param([string]$Text, [string]$SourcePath)
    $table = Read-CsvTable $Text
    if ($table.Count -eq 0) { throw "CSVが空です。" }
    $labeled = New-Object System.Collections.Generic.List[object]
    $number = 1
    foreach ($row in @($table)) {
        $labeled.Add([pscustomobject]@{ Label = "CSV ${number}行目"; Row = @($row) })
        $number++
    }
    return ConvertFrom-Table $labeled $SourcePath "CSV"
}

function Get-SheetKind {
    param([string]$Name)
    $token = ConvertTo-Token $Name
    $ignored = @("readme", "説明", "読み方", "使い方", "表紙", "manual")
    if ($ignored -contains $token) { return "ignore" }
    $map = @{
        "api" = "api"; "apis" = "api"; "リクエスト" = "api"
        "query" = "query"; "queryparams" = "query"; "クエリ" = "query"; "クエリパラメータ" = "query"
        "header" = "header"; "headers" = "header"; "ヘッダ" = "header"; "ヘッダー" = "header"; "ヘッダ情報" = "header"
        "body" = "form"; "bodyparams" = "form"; "form" = "form"; "forms" = "form"; "フォーム" = "form"; "postパラメータ" = "form"; "ボディパラメータ" = "form"
        "setting" = "settings"; "settings" = "settings"; "設定" = "settings"
    }
    if ($map.ContainsKey($token)) { return $map[$token] }
    return "unknown"
}

function Get-XmlLocalText {
    param($Node)
    if ($null -eq $Node) { return "" }
    if ($Node.LocalName -eq "t" -and $Node.InnerText) { return $Node.InnerText }
    $builder = New-Object System.Text.StringBuilder
    foreach ($child in @($Node.ChildNodes)) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        if ($child.LocalName -eq "rPh" -or $child.LocalName -eq "phoneticPr") { continue }
        if ($child.LocalName -eq "t") { [void]$builder.Append($child.InnerText) }
        else { [void]$builder.Append((Get-XmlLocalText $child)) }
    }
    return $builder.ToString()
}

function ConvertTo-CellText {
    param($Cell, [object[]]$Shared)
    $kind = ""
    if ($Cell.HasAttribute("t")) { $kind = $Cell.GetAttribute("t") }
    $valueNode = $null
    foreach ($child in @($Cell.ChildNodes)) {
        if ($child.LocalName -eq "v") { $valueNode = $child; break }
    }
    if ($kind -eq "inlineStr") { return (Get-XmlLocalText $Cell) }
    if ($kind -eq "s") {
        if ($null -eq $valueNode) { return "" }
        return [string]$Shared[[int]$valueNode.InnerText]
    }
    if ($kind -eq "b") {
        if ($null -eq $valueNode) { return "" }
        if ($valueNode.InnerText.Trim() -eq "1" -or $valueNode.InnerText.Trim() -eq "true") { return "true" }
        return "false"
    }
    if ($null -eq $valueNode) { return "" }
    $raw = $valueNode.InnerText.Trim()
    if ($raw -match "^-?\d+$") { return $raw }
    if ($raw -match "^-?\d+(\.\d+)?([eE][+-]?\d+)?$") {
        $number = 0.0
        if ([double]::TryParse($raw, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            if ($number -eq [math]::Round($number) -and [math]::Abs($number) -lt 1e15) { return ([int64][math]::Round($number)).ToString() }
        }
    }
    return $valueNode.InnerText
}

function Get-ColumnIndex {
    param([string]$Ref)
    $letters = ""
    foreach ($ch in $Ref.ToCharArray()) { if ([char]::IsLetter($ch)) { $letters += $ch } }
    $index = 0
    foreach ($ch in $letters.ToCharArray()) { $index = $index * 26 + ([int][char]::ToUpper($ch) - 64) }
    if ($index -le 0) { throw "セル参照が不正です: $Ref" }
    return $index
}

function Read-ZipText {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, (New-Object System.Text.UTF8Encoding $false), $true)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Resolve-ZipTarget {
    param([string]$Base, [string]$Target)
    $parts = New-Object System.Collections.Generic.List[string]
    $source = @()
    if ($Target.StartsWith("/")) { $source = $Target.TrimStart("/").Split("/") }
    else {
        if ($Base -ne "") { $source += $Base.Split("/") }
        $source += $Target.Split("/")
    }
    foreach ($part in $source) {
        if ($part -eq "" -or $part -eq ".") { continue }
        if ($part -eq "..") { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }; continue }
        $parts.Add($part)
    }
    return ($parts -join "/")
}

function Read-XlsxSheets {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $Path)) { throw "入力ファイルがありません: $Path" }
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $workbookText = Read-ZipText $zip "xl/workbook.xml"
        if ($null -eq $workbookText) { throw "workbook.xml がありません: $Path" }
        $shared = @()
        $sharedText = Read-ZipText $zip "xl/sharedStrings.xml"
        if ($null -ne $sharedText) {
            $sharedDoc = New-Object System.Xml.XmlDocument
            $sharedDoc.LoadXml($sharedText)
            $list = New-Object System.Collections.Generic.List[string]
            foreach ($node in @($sharedDoc.GetElementsByTagName("*"))) {
                if ($node.LocalName -eq "si" -and $node.ParentNode.LocalName -eq "sst") { $list.Add((Get-XmlLocalText $node)) }
            }
            $shared = $list.ToArray()
        }
        $relsText = Read-ZipText $zip "xl/_rels/workbook.xml.rels"
        $relsDoc = New-Object System.Xml.XmlDocument
        $relsDoc.LoadXml($relsText)
        $relMap = @{}
        foreach ($node in @($relsDoc.GetElementsByTagName("*"))) {
            if ($node.LocalName -ne "Relationship") { continue }
            $type = $node.GetAttribute("Type")
            if ($type -notlike "*worksheet*") { continue }
            $relMap[$node.GetAttribute("Id")] = Resolve-ZipTarget "xl" $node.GetAttribute("Target")
        }
        $book = New-Object System.Xml.XmlDocument
        $book.LoadXml($workbookText)
        $sheets = New-Object System.Collections.Generic.List[object]
        foreach ($node in @($book.GetElementsByTagName("*"))) {
            if ($node.LocalName -ne "sheet") { continue }
            $name = $node.GetAttribute("name")
            $relId = ""
            foreach ($attr in @($node.Attributes)) { if ($attr.LocalName -eq "id") { $relId = $attr.Value } }
            if (-not $relMap.ContainsKey($relId)) { throw "シートの実体がありません: $name" }
            $sheetText = Read-ZipText $zip $relMap[$relId]
            $sheetDoc = New-Object System.Xml.XmlDocument
            $sheetDoc.LoadXml($sheetText)
            $rows = New-Object System.Collections.Generic.List[object]
            $numbers = New-Object System.Collections.Generic.List[int]
            $nextNumber = 1
            foreach ($rowNode in @($sheetDoc.GetElementsByTagName("*"))) {
                if ($rowNode.LocalName -ne "row") { continue }
                $number = $nextNumber
                if ($rowNode.GetAttribute("r") -match "^[0-9]+$") { $number = [int]$rowNode.GetAttribute("r") }
                $nextNumber = $number + 1
                $cells = @{}
                $maxCol = 0
                $sequential = 1
                foreach ($cell in @($rowNode.ChildNodes)) {
                    if ($cell.NodeType -ne [System.Xml.XmlNodeType]::Element -or $cell.LocalName -ne "c") { continue }
                    $ref = $cell.GetAttribute("r")
                    if ($ref -ne "") { $column = Get-ColumnIndex $ref } else { $column = $sequential }
                    $sequential = $column + 1
                    $cells[$column] = ConvertTo-CellText $cell $shared
                    if ($column -gt $maxCol) { $maxCol = $column }
                }
                $values = @()
                for ($col = 1; $col -le $maxCol; $col++) {
                    if ($cells.ContainsKey($col)) { $values += [string]$cells[$col] } else { $values += "" }
                }
                $rows.Add($values)
                $numbers.Add($number)
            }
            $sheets.Add([pscustomobject]@{ Name = $name; Rows = $rows; Numbers = $numbers })
        }
        if ($sheets.Count -eq 0) { throw "Excelにシートがありません。" }
        return ,@($sheets.ToArray())
    } finally { $zip.Dispose() }
}

function ConvertFrom-SheetRows {
    param($Sheet, [string]$RoleKind)
    $labeled = New-Object System.Collections.Generic.List[object]
    $rows = Copy-Items $Sheet.Rows
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $number = [int]$Sheet.Numbers[$i]
        $labeled.Add([pscustomobject]@{ Label = "$($Sheet.Name) ${number}行目"; Row = @($rows[$i]) })
    }
    return $labeled
}

function ConvertFrom-Workbook {
    param([object[]]$Sheets, [string]$SourcePath)
    $grouped = @{}
    $loose = New-Object System.Collections.Generic.List[object]
    foreach ($sheet in @($Sheets)) {
        $kind = Get-SheetKind $sheet.Name
        if ($kind -eq "ignore") { continue }
        if ($kind -eq "unknown") { $loose.Add($sheet); continue }
        if ($grouped.ContainsKey($kind)) { throw "同じ種類のシートが複数あります: $($sheet.Name)" }
        $grouped[$kind] = $sheet
    }
    if (-not $grouped.ContainsKey("api")) {
        if ($loose.Count -eq 1) {
            $labeled = ConvertFrom-SheetRows $loose[0] "flat"
            return ConvertFrom-Table $labeled $SourcePath $loose[0].Name
        }
        throw "APIシートがありません。シート名は API / クエリ / ヘッダ / フォーム / 設定 にしてください。"
    }
    $apiSheet = $grouped["api"]
    $labeled = ConvertFrom-SheetRows $apiSheet "api"
    $header = $null
    $headerIndex = -1
    $rows = @($labeled)
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if (-not (Test-BlankRow $rows[$i].Row)) { $header = $rows[$i]; $headerIndex = $i; break }
    }
    if ($null -eq $header) { throw "$($apiSheet.Name) にヘッダ行がありません。" }
    $map = Get-HeaderMap $header.Row (Get-RoleTable "api")
    foreach ($required in @("name", "method", "url")) {
        if (-not $map.ContainsKey($required)) { throw "$($apiSheet.Name) に必須列がありません: $required" }
    }
    $records = New-Object System.Collections.Generic.List[object]
    for ($i = $headerIndex + 1; $i -lt $rows.Count; $i++) {
        if (Test-BlankRow $rows[$i].Row) { continue }
        $record = ConvertTo-Record $map $rows[$i].Row $rows[$i].Label "API"
        $record.ApiName = Get-RowValue $rows[$i].Row $map "name"
        $records.Add($record)
    }
    foreach ($pair in @(@("query", "QUERY"), @("header", "HEADER"), @("form", "BODY"))) {
        if (-not $grouped.ContainsKey($pair[0])) { continue }
        $childSheet = $grouped[$pair[0]]
        $childRows = @(ConvertFrom-SheetRows $childSheet "child")
        $childHeader = $null
        $childIndex = -1
        for ($i = 0; $i -lt $childRows.Count; $i++) {
            if (-not (Test-BlankRow $childRows[$i].Row)) { $childHeader = $childRows[$i]; $childIndex = $i; break }
        }
        if ($null -eq $childHeader) { continue }
        $childMap = Get-HeaderMap $childHeader.Row (Get-RoleTable "child")
        foreach ($required in @("api_name", "key", "value")) {
            if (-not $childMap.ContainsKey($required)) { throw "$($childSheet.Name) に必須列がありません: $required" }
        }
        for ($i = $childIndex + 1; $i -lt $childRows.Count; $i++) {
            if (Test-BlankRow $childRows[$i].Row) { continue }
            $record = ConvertTo-Record $childMap $childRows[$i].Row $childRows[$i].Label $pair[1]
            $records.Add($record)
        }
    }
    if ($grouped.ContainsKey("settings")) {
        $settingRows = @(ConvertFrom-SheetRows $grouped["settings"] "settings")
        $settingHeader = $null
        $settingIndex = -1
        for ($i = 0; $i -lt $settingRows.Count; $i++) {
            if (-not (Test-BlankRow $settingRows[$i].Row)) { $settingHeader = $settingRows[$i]; $settingIndex = $i; break }
        }
        if ($null -ne $settingHeader) {
            $settingMap = Get-HeaderMap $settingHeader.Row (Get-RoleTable "settings")
            for ($i = $settingIndex + 1; $i -lt $settingRows.Count; $i++) {
                if (Test-BlankRow $settingRows[$i].Row) { continue }
                $record = ConvertTo-Record $settingMap $settingRows[$i].Row $settingRows[$i].Label "SETTING"
                $records.Add($record)
            }
        }
    }
    return ConvertTo-Plan $records $SourcePath
}

function Import-ApiFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "入力ファイルがありません: $Path" }
    $full = [System.IO.Path]::GetFullPath($Path)
    $ext = [System.IO.Path]::GetExtension($full).ToLower()
    if ($ext -eq ".xlsx" -or $ext -eq ".xlsm") { return ConvertFrom-Workbook (Read-XlsxSheets $full) $full }
    if ($ext -eq ".csv" -or $ext -eq ".txt") { return ConvertFrom-CsvText (Read-TextAuto $full) $full }
    throw "入力は .xlsx または .csv にしてください: $Path"
}

function Resolve-RequestUrl {
    param($Api, [System.Collections.Generic.List[string]]$Warnings)
    try { $uri = New-Object System.Uri $Api.Url.Trim() } catch { throw "$($Api.Name): URLが不正です: $($Api.Url)" }
    $scheme = $uri.Scheme.ToLower()
    if ($scheme -ne "http" -and $scheme -ne "https") { throw "$($Api.Name): URLは http または https にしてください: $($Api.Url)" }
    if ($uri.UserInfo) { throw "$($Api.Name): URLにユーザー情報を含めないでください。認証はヘッダで指定します。" }
    if ($uri.Fragment) { $Warnings.Add("$($Api.Name): URLのフラグメントは送信されないため無視しました。") }
    $hostName = $uri.DnsSafeHost
    if ([string]::IsNullOrWhiteSpace($hostName)) { throw "$($Api.Name): URLにホスト名がありません: $($Api.Url)" }
    $portText = ""
    if (-not $uri.IsDefaultPort) {
        if ($uri.Port -lt 1 -or $uri.Port -gt 65535) { throw "$($Api.Name): ポートが不正です: $($uri.Port)" }
        $portText = [string]$uri.Port
    }
    $domain = $hostName
    if ($hostName.Contains(":")) { $domain = "[$hostName]" }
    $path = $uri.GetComponents([System.UriComponents]::Path, [System.UriFormat]::UriEscaped)
    if ($path -eq "" -or $path -eq $null) { $path = "/" }
    if (-not $path.StartsWith("/")) { $path = "/" + $path }
    $queryRaw = $uri.GetComponents([System.UriComponents]::Query, [System.UriFormat]::UriEscaped)
    $query = ConvertFrom-QueryString $queryRaw "$($Api.Name) のクエリ"
    return [pscustomobject]@{ Scheme = $scheme; Domain = $domain; Port = $portText; Path = $path; Query = $query }
}

function Get-BodyMode {
    param($Api)
    $token = "auto"
    if (-not [string]::IsNullOrWhiteSpace($Api.BodyMode)) { $token = ConvertTo-Token $Api.BodyMode }
    $modes = @{ "auto" = "auto"; "自動" = "auto"; "raw" = "raw"; "json" = "raw"; "生" = "raw"; "そのまま" = "raw"; "form" = "form"; "フォーム" = "form"; "パラメータ" = "form" }
    if (-not $modes.ContainsKey($token)) { throw "$($Api.Name): ボディ形式が不正です: $($Api.BodyMode)" }
    $mode = $modes[$token]
    $hasForm = $Api.FormParams.Count -gt 0
    $hasBody = $Api.Body -ne ""
    $contentType = ([string]$Api.ContentType).ToLower()
    if ($contentType.Contains("multipart/form-data")) { throw "$($Api.Name): multipart/form-data には対応していません。" }
    if ($Api.Method -eq "GET" -and ($hasBody -or $hasForm -or $mode -eq "raw" -or $mode -eq "form")) {
        throw "$($Api.Name): GET にリクエストボディは指定できません。"
    }
    if ($mode -eq "raw") {
        if ($hasForm) { throw "$($Api.Name): ボディ形式 raw とフォーム項目は同時に指定できません。" }
        if ($hasBody) { return "raw" }
        return "empty"
    }
    if ($mode -eq "form") {
        if ($hasBody -and $hasForm) { throw "$($Api.Name): リクエストボディとフォーム項目は同時に指定できません。" }
        if ($hasBody) { return "form_body" }
        if ($hasForm) { return "form_params" }
        throw "$($Api.Name): フォーム項目がありません。"
    }
    if ($hasForm -and $hasBody) { throw "$($Api.Name): リクエストボディとフォーム項目は同時に指定できません。" }
    if ($hasForm) {
        if ($contentType.Contains("json")) { throw "$($Api.Name): Content-Type が JSON ですがフォーム項目があります。" }
        return "form_params"
    }
    if (-not $hasBody) { return "empty" }
    if ($contentType.Contains("application/x-www-form-urlencoded")) { return "form_body" }
    return "raw"
}

function Get-MergedHeaders {
    param($Api, [string]$Mode)
    $headers = New-Object System.Collections.Generic.List[object]
    foreach ($pair in $Api.Headers) { $headers.Add($pair) }
    $contentIndexes = @()
    for ($i = 0; $i -lt $headers.Count; $i++) {
        if ($headers[$i].Key.ToLower() -eq "content-type") { $contentIndexes += $i }
    }
    if ($contentIndexes.Count -gt 1) { throw "$($Api.Name): Content-Type ヘッダが重複しています。" }
    $userType = ([string]$Api.ContentType).Trim()
    if ($userType -ne "" -and $contentIndexes.Count -eq 1) {
        $current = $headers[$contentIndexes[0]].Value.Trim()
        if ($current.ToLower() -ne $userType.ToLower()) { throw "$($Api.Name): Content-Type 列とヘッダの値が矛盾しています。" }
        return $headers
    }
    if ($userType -ne "") { $headers.Insert(0, (New-Pair "Content-Type" $userType)); return $headers }
    if ($contentIndexes.Count -eq 1) { return $headers }
    if ($Mode -eq "raw") {
        $sniff = $Api.Body.TrimStart(@([char]0xFEFF, [char]32, [char]9, [char]13, [char]10))
        if ($sniff.StartsWith("{") -or $sniff.StartsWith("[")) { $headers.Insert(0, (New-Pair "Content-Type" "application/json; charset=UTF-8")) }
    } elseif ($Mode -eq "form_body" -or $Mode -eq "form_params") {
        $headers.Insert(0, (New-Pair "Content-Type" "application/x-www-form-urlencoded; charset=UTF-8"))
    }
    return $headers
}

function Resolve-ApiRequest {
    param($Api, [System.Collections.Generic.List[string]]$Warnings)
    $mode = Get-BodyMode $Api
    $url = Resolve-RequestUrl $Api (, $Warnings)
    $path = ConvertTo-Percent $url.Path "/%:@-._~!$&'()*+,;="
    $query = New-Object System.Collections.Generic.List[object]
    foreach ($pair in (Copy-Items $url.Query)) { $query.Add($pair) }
    foreach ($pair in $Api.Query) { $query.Add($pair) }
    $headers = Get-MergedHeaders $Api $mode
    $arguments = New-Object System.Collections.Generic.List[object]
    $rawBody = $null
    $finalPath = $path
    if ($Api.Method -eq "GET") {
        foreach ($pair in $query) { $arguments.Add($pair) }
    } elseif ($mode -eq "form_body") {
        $parsed = ConvertFrom-QueryString $Api.Body "$($Api.Name) のボディ"
        if (@($parsed).Count -eq 0) { throw "$($Api.Name): フォーム項目がありません。" }
        foreach ($pair in @($parsed)) { $arguments.Add($pair) }
        $finalPath = Add-QueryPath $path $query
    } elseif ($mode -eq "form_params") {
        foreach ($pair in $Api.FormParams) { $arguments.Add($pair) }
        $finalPath = Add-QueryPath $path $query
    } elseif ($mode -eq "raw") {
        $rawBody = [string]$Api.Body
        $finalPath = Add-QueryPath $path $query
    } else {
        $finalPath = Add-QueryPath $path $query
    }
    $timeout = ""
    if (([string]$Api.TimeoutMs).Trim() -ne "") { $timeout = [string](ConvertTo-IntRange $Api.TimeoutMs "$($Api.Name) のタイムアウト" 1 3600000) }
    $status = ([string]$Api.ExpectedStatus).Trim()
    if ($status -ne "" -and $status -notmatch "^[1-5][0-9][0-9]$") { throw "$($Api.Name): 期待ステータスは 100 から 599 の整数です: $status" }
    return [pscustomobject]@{
        Name = $Api.Name; Method = $Api.Method; Protocol = $url.Scheme; Domain = $url.Domain; Port = $url.Port
        Path = $finalPath; Arguments = $arguments; RawBody = $rawBody; Headers = $headers; TimeoutMs = $timeout
        ExpectedStatus = $status; Comment = $Api.Comment; OriginalUrl = $Api.Url.Trim(); Enabled = [bool]$Api.Enabled
    }
}

function Add-QueryPath {
    param([string]$Path, [object[]]$Pairs)
    $items = @($Pairs)
    if ($items.Count -eq 0 -or ($items.Count -eq 1 -and $null -eq $items[0])) { return $Path }
    $encoded = @()
    foreach ($pair in $items) {
        if ($null -eq $pair) { continue }
        $encoded += (ConvertTo-Percent $pair.Key "-._~") + "=" + (ConvertTo-Percent $pair.Value "-._~")
    }
    if ($encoded.Count -eq 0) { return $Path }
    return $Path + "?" + ($encoded -join "&")
}

function Get-DisplayUrl {
    param($Item)
    $port = ""
    if ($Item.Port) { $port = ":" + $Item.Port }
    $base = $Item.Protocol + "://" + $Item.Domain + $port + $Item.Path
    if ($Item.Method -eq "GET" -and (Get-ItemCount $Item.Arguments) -gt 0 -and $Item.Path.IndexOf("?") -lt 0) {
        $encoded = @()
        foreach ($pair in (Get-Pairs (, $Item.Arguments))) {
            if ($null -eq $pair -or $null -eq $pair.PSObject.Properties["Key"]) { continue }
            $encoded += (ConvertTo-Percent $pair.Key "-._~") + "=" + (ConvertTo-Percent $pair.Value "-._~")
        }
        return $base + "?" + ($encoded -join "&")
    }
    return $base
}

function Get-SaveConfigXml {
    param([bool]$IncludeResponse)
    $response = "false"
    if ($IncludeResponse) { $response = "true" }
    return @"
<objProp>
              <name>saveConfig</name>
              <value class="SampleSaveConfiguration">
                <time>true</time>
                <latency>true</latency>
                <timestamp>true</timestamp>
                <success>true</success>
                <label>true</label>
                <code>true</code>
                <message>true</message>
                <threadName>true</threadName>
                <dataType>true</dataType>
                <encoding>false</encoding>
                <assertions>true</assertions>
                <subresults>true</subresults>
                <responseData>$response</responseData>
                <samplerData>false</samplerData>
                <xml>false</xml>
                <fieldNames>true</fieldNames>
                <responseHeaders>false</responseHeaders>
                <requestHeaders>false</requestHeaders>
                <responseDataOnError>false</responseDataOnError>
                <saveAssertionResultsFailureMessage>true</saveAssertionResultsFailureMessage>
                <assertionsResultsToSave>0</assertionsResultsToSave>
                <bytes>true</bytes>
                <sentBytes>true</sentBytes>
                <url>true</url>
                <threadCounts>true</threadCounts>
                <idleTime>true</idleTime>
                <connectTime>true</connectTime>
              </value>
            </objProp>
"@
}

function Get-StringPropXml { param($Name, $Value) return '<stringProp name="' + (ConvertTo-XmlText $Name) + '">' + (ConvertTo-XmlText $Value) + '</stringProp>' }
function Get-BoolPropXml {
    param($Name, [bool]$Value)
    $text = "false"; if ($Value) { $text = "true" }
    return '<boolProp name="' + (ConvertTo-XmlText $Name) + '">' + $text + '</boolProp>'
}

function Get-Pairs {
    param($Value)
    $result = New-Object System.Collections.Generic.List[object]
    $pending = New-Object System.Collections.Generic.Stack[object]
    if ($null -ne $Value) { $pending.Push($Value) }
    while ($pending.Count -gt 0) {
        $current = $pending.Pop()
        if ($null -eq $current -or $current -is [string] -or $current -is [pscustomobject]) {
            if ($null -ne $current) { $result.Add($current) }
            continue
        }
        $enumerator = $null
        try { $enumerator = $current.GetEnumerator() } catch { $enumerator = $null }
        if ($null -eq $enumerator) { continue }
        $nested = New-Object System.Collections.Generic.List[object]
        while ($enumerator.MoveNext()) { $nested.Add($enumerator.Current) }
        if ($nested.Count -eq 0) { continue }
        $only = $nested[0]
        if ($nested.Count -eq 1 -and $null -ne $only -and $only -isnot [string] -and $only -isnot [pscustomobject]) {
            $pending.Push($only)
            continue
        }
        foreach ($item in $nested) { $result.Add($item) }
    }
    return ,$result.ToArray()
}

function Get-ArgumentXml {
    param($Name, $Value, [bool]$Encode)
    $nameAttr = ""
    if ($null -ne $Name) { $nameAttr = ConvertTo-XmlText $Name }
    $lines = @(
        ('<elementProp name="' + $nameAttr + '" elementType="HTTPArgument">'),
        (Get-BoolPropXml "HTTPArgument.always_encode" $Encode),
        (Get-StringPropXml "Argument.value" $Value),
        (Get-StringPropXml "Argument.metadata" "="),
        (Get-BoolPropXml "HTTPArgument.use_equals" $true)
    )
    if ($null -ne $Name) { $lines += (Get-StringPropXml "Argument.name" $Name) }
    $lines += "</elementProp>"
    return ($lines -join "`n")
}

function Get-SamplerXml {
    param($Item, [bool]$FollowRedirects)
    $postRaw = $false
    if ($null -eq $Item.RawBody) {
        $chunks = @()
        foreach ($pair in (Get-Pairs (, $Item.Arguments))) {
            if ($null -eq $pair -or $null -eq $pair.PSObject.Properties["Key"]) { continue }
            $chunks += (Get-ArgumentXml $pair.Key $pair.Value $true)
        }
        $arguments = $chunks -join "`n"
    } else {
        $arguments = Get-ArgumentXml $null $Item.RawBody $false
        $postRaw = $true
    }
    $argumentBlock = '<elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true"><collectionProp name="Arguments.arguments">' + $arguments + '</collectionProp></elementProp>'
    return @(
        ('<HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="' + (ConvertTo-XmlText $Item.Name) + '" enabled="true">'),
        (Get-BoolPropXml "HTTPSampler.postBodyRaw" $postRaw),
        $argumentBlock,
        (Get-StringPropXml "HTTPSampler.domain" $Item.Domain),
        (Get-StringPropXml "HTTPSampler.port" $Item.Port),
        (Get-StringPropXml "HTTPSampler.protocol" $Item.Protocol),
        (Get-StringPropXml "HTTPSampler.contentEncoding" "UTF-8"),
        (Get-StringPropXml "HTTPSampler.path" $Item.Path),
        (Get-StringPropXml "HTTPSampler.method" $Item.Method),
        (Get-BoolPropXml "HTTPSampler.follow_redirects" $FollowRedirects),
        (Get-BoolPropXml "HTTPSampler.auto_redirects" $false),
        (Get-BoolPropXml "HTTPSampler.use_keepalive" $true),
        (Get-BoolPropXml "HTTPSampler.DO_MULTIPART_POST" $false),
        (Get-StringPropXml "HTTPSampler.embedded_url_re" ""),
        (Get-StringPropXml "HTTPSampler.connect_timeout" $Item.TimeoutMs),
        (Get-StringPropXml "HTTPSampler.response_timeout" $Item.TimeoutMs),
        (Get-StringPropXml "HTTPSampler.implementation" "HttpClient4"),
        "</HTTPSamplerProxy>"
    ) -join "`n"
}

function Get-HeaderManagerXml {
    param($Item)
    $headerPairs = Get-Pairs (, $Item.Headers)
    if ($headerPairs.Count -eq 0) { return "" }
    $headers = @()
    foreach ($pair in $headerPairs) {
        if ($null -eq $pair -or $null -eq $pair.PSObject.Properties["Key"]) { continue }
        $headers += @(
            '<elementProp name="" elementType="Header">',
            (Get-StringPropXml "Header.name" $pair.Key),
            (Get-StringPropXml "Header.value" $pair.Value),
            "</elementProp>"
        ) -join "`n"
    }
    return '<HeaderManager guiclass="HeaderPanel" testclass="HeaderManager" testname="HTTP Headers" enabled="true"><collectionProp name="HeaderManager.headers">' + ($headers -join "") + '</collectionProp></HeaderManager><hashTree/>'
}

function Get-AssertionXml {
    param($Item)
    if ([string]$Item.ExpectedStatus -eq "") { return "" }
    return '<ResponseAssertion guiclass="AssertionGui" testclass="ResponseAssertion" testname="Expected status ' + (ConvertTo-XmlText $Item.ExpectedStatus) + '" enabled="true"><collectionProp name="Asserion.test_strings"><stringProp name="0">' + (ConvertTo-XmlText $Item.ExpectedStatus) + '</stringProp></collectionProp>' + (Get-StringPropXml "Assertion.custom_message" ("Expected HTTP status " + $Item.ExpectedStatus)) + (Get-StringPropXml "Assertion.test_field" "Assertion.response_code") + (Get-BoolPropXml "Assertion.assume_success" $false) + '<intProp name="Assertion.test_type">8</intProp></ResponseAssertion><hashTree/>'
}

function Get-CollectorXml {
    param($Gui, $Title, $FileName, [bool]$IncludeResponse)
    return '<ResultCollector guiclass="' + $Gui + '" testclass="ResultCollector" testname="' + (ConvertTo-XmlText $Title) + '" enabled="true">' + (Get-BoolPropXml "ResultCollector.error_logging" $false) + (Get-SaveConfigXml $IncludeResponse) + (Get-StringPropXml "filename" $FileName) + '</ResultCollector><hashTree/>'
}

function New-JmxText {
    param($Plan, [object[]]$Resolved)
    $enabled = New-Object System.Collections.Generic.List[object]
    foreach ($item in (Unwrap-Records $Resolved)) { if ($item.Enabled) { $enabled.Add($item) } }
    if ($enabled.Count -eq 0) { throw "有効なAPIが1件もありません。" }
    $settings = $Plan.Settings
    $comments = "Apache JMeter $($Script:JMeterVersion) 向けに自動生成したREST APIテストです。スレッド数=$($settings.Threads)、ループ回数=$($settings.Loops)、ランプアップ=$($settings.RampUpSec)秒。初期値では有効な各APIを1回呼び出します。Simple Data Writer は作業ディレクトリ基準の $($settings.JtlFile) にJTLを書きます。JSR223 Listener は jmeter.log に API_RESULT 行を出します。非GUIで実行するときは、出力ディレクトリを作業ディレクトリにしてください。"
    $groovy = 'def urlText = ""; try { urlText = String.valueOf(prev.getURL()) } catch (Exception ignored) { urlText = "" }; log.info("API_RESULT name=" + prev.getSampleLabel() + " code=" + prev.getResponseCode() + " success=" + prev.isSuccessful() + " elapsed_ms=" + prev.getTime() + " bytes=" + prev.getBytesAsLong() + " url=" + urlText)'
    $samplers = @()
    foreach ($item in $enabled) {
        $children = @()
        $headerXml = Get-HeaderManagerXml $item
        $assertionXml = Get-AssertionXml $item
        if ($headerXml -ne "") { $children += $headerXml }
        if ($assertionXml -ne "") { $children += $assertionXml }
        $samplers += (Get-SamplerXml $item $settings.FollowRedirects) + "<hashTree>" + ($children -join "") + "</hashTree>"
    }
    $listener = '<JSR223Listener guiclass="TestBeanGUI" testclass="JSR223Listener" testname="Log sample result" enabled="true">' + (Get-StringPropXml "cacheKey" "true") + (Get-StringPropXml "filename" "") + (Get-StringPropXml "parameters" "") + (Get-StringPropXml "script" $groovy) + (Get-StringPropXml "scriptLanguage" "groovy") + '</JSR223Listener><hashTree/>'
    $xml = '<?xml version="1.0" encoding="UTF-8"?>' + "`n" +
        '<jmeterTestPlan version="1.2" properties="5.0" jmeter="' + $Script:JMeterVersion + '"><hashTree>' +
        '<TestPlan guiclass="TestPlanGui" testclass="TestPlan" testname="' + (ConvertTo-XmlText $settings.PlanName) + '" enabled="true">' +
        (Get-StringPropXml "TestPlan.comments" $comments) +
        (Get-BoolPropXml "TestPlan.functional_mode" $false) +
        (Get-BoolPropXml "TestPlan.tearDown_on_shutdown" $true) +
        (Get-BoolPropXml "TestPlan.serialize_threadgroups" $false) +
        '<elementProp name="TestPlan.user_defined_variables" elementType="Arguments" guiclass="ArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true"><collectionProp name="Arguments.arguments"/></elementProp>' +
        (Get-StringPropXml "TestPlan.user_define_classpath" "") +
        '</TestPlan><hashTree><ThreadGroup guiclass="ThreadGroupGui" testclass="ThreadGroup" testname="REST API" enabled="true">' +
        (Get-StringPropXml "ThreadGroup.on_sample_error" $settings.OnError) +
        '<elementProp name="ThreadGroup.main_controller" elementType="LoopController" guiclass="LoopControlPanel" testclass="LoopController" testname="Loop Controller" enabled="true">' +
        (Get-BoolPropXml "LoopController.continue_forever" $false) +
        (Get-StringPropXml "LoopController.loops" ([string]$settings.Loops)) +
        '</elementProp><intProp name="ThreadGroup.num_threads">' + $settings.Threads + '</intProp><intProp name="ThreadGroup.ramp_time">' + $settings.RampUpSec + '</intProp>' +
        (Get-BoolPropXml "ThreadGroup.scheduler" $false) +
        (Get-StringPropXml "ThreadGroup.duration" "") +
        (Get-StringPropXml "ThreadGroup.delay" "") +
        (Get-BoolPropXml "ThreadGroup.same_user_on_next_iteration" $true) +
        '</ThreadGroup><hashTree>' + ($samplers -join "") + $listener +
        (Get-CollectorXml "SimpleDataWriter" "Save JTL" $settings.JtlFile $false) +
        (Get-CollectorXml "SummaryReport" "Summary Report" "" $false) +
        (Get-CollectorXml "ViewResultsFullVisualizer" "View Results Tree" "" $true) +
        '</hashTree></hashTree></hashTree></jmeterTestPlan>' + "`n"
    return $xml
}

function Get-BashQuote { param([string]$Value) return "'" + $Value.Replace("'", "'\''") + "'" }
function Get-PsQuote { param([string]$Value) return "'" + $Value.Replace("'", "''") + "'" }

function New-UsageSheets {
    param($Plan, [object[]]$Resolved, [string]$OutputDir)
    $settings = $Plan.Settings
    $resolvedItems = Unwrap-Records $Resolved
    $enabledCount = 0
    foreach ($item in $resolvedItems) { if ($item.Enabled) { $enabledCount++ } }
    $jmx = Join-Path $OutputDir $Script:JmxFileName
    $results = Join-Path $OutputDir "results"
    $jtl = Join-Path $results "results.jtl"
    $logPath = Join-Path $results "jmeter.log"
    $report = Join-Path $results "html-report"
    $listener = Join-Path $results "listener-results.jtl"
    $bash = @(
        ("cd " + (Get-BashQuote $OutputDir)),
        ("mkdir -p " + (Get-BashQuote $results)),
        ("rm -f " + (Get-BashQuote $jtl) + " " + (Get-BashQuote $logPath) + " " + (Get-BashQuote $listener)),
        ("rm -rf " + (Get-BashQuote $report)),
        '"$JMETER_HOME/bin/jmeter" -n \',
        ("  -t " + (Get-BashQuote $jmx) + " \"),
        ("  -l " + (Get-BashQuote $jtl) + " \"),
        ("  -j " + (Get-BashQuote $logPath) + " \"),
        ("  -e -o " + (Get-BashQuote $report))
    ) -join "`n"
    $powershell = @(
        ("Set-Location -LiteralPath " + (Get-PsQuote $OutputDir)),
        ("New-Item -ItemType Directory -Force -Path " + (Get-PsQuote $results) + " | Out-Null"),
        ("Remove-Item -LiteralPath " + (Get-PsQuote $jtl) + "," + (Get-PsQuote $logPath) + "," + (Get-PsQuote $listener) + " -Force -ErrorAction SilentlyContinue"),
        ("if (Test-Path -LiteralPath " + (Get-PsQuote $report) + ") { Remove-Item -LiteralPath " + (Get-PsQuote $report) + " -Recurse -Force }"),
        '& "$env:JMETER_HOME\bin\jmeter.bat" -n `',
        ("  -t " + (Get-PsQuote $jmx) + ' `'),
        ("  -l " + (Get-PsQuote $jtl) + ' `'),
        ("  -j " + (Get-PsQuote $logPath) + ' `'),
        ("  -e -o " + (Get-PsQuote $report))
    ) -join "`n"
    $redirect = "false"; if ($settings.FollowRedirects) { $redirect = "true" }
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add((New-Sheet "概要" @(
        @("項目", "内容"),
        @("対象", "Apache JMeter $($Script:JMeterVersion)"),
        @("テスト計画", $settings.PlanName),
        @("JMX", $jmx),
        @("入力", $Plan.SourcePath),
        @("スレッド数", [string]$settings.Threads),
        @("ループ回数", [string]$settings.Loops),
        @("ランプアップ秒", [string]$settings.RampUpSec),
        @("呼び出し回数", "有効なAPIごとに、スレッド数 × ループ回数 です。初期値は各API 1回です。"),
        @("有効なAPI", [string]$enabledCount),
        @("JTL（CLIの -l）", $jtl),
        @("JTL（JMX内のリスナー）", $listener + " 。作業ディレクトリが出力ディレクトリのとき、この相対パスに書かれます。"),
        @("ログ", $logPath + " 。各サンプルは API_RESULT で検索できます。"),
        @("HTMLレポート", (Join-Path $report "index.html")),
        @("エラー時の動作", $settings.OnError),
        @("リダイレクト追従", $redirect),
        @("リスナーのJTL相対パス", $settings.JtlFile)
    ) @(28, 96)))
    $sheets.Add((New-Sheet "環境準備" @(
        @("項目", "内容"),
        @("Java", "JMeter 5.6.3 は Java 8、11、17 で動作します。Java 17 を推奨します。Java 22 以降では動作しません。"),
        @("RHEL 9.8 の Java", "sudo dnf install -y java-17-openjdk-headless"),
        @("Windows 11 の Java", "Eclipse Temurin 17 などの JDK 17 を入れ、java が PATH で見えるようにします。"),
        @("JMeterの入手", "https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.zip を展開します。"),
        @("JMETER_HOME", "展開先の apache-jmeter-5.6.3 ディレクトリを JMETER_HOME にします。"),
        @("RHEL 9.8 の生成スクリプト", "bash と python3 が必要です。python3 が無い場合は sudo dnf install -y python3"),
        @("Windows 11 の生成スクリプト", "Windows PowerShell 5.1 だけで動作します。Excel や Python のインストールは不要です。"),
        @("秘密情報", "Authorization などのヘッダ値は JMX に平文で入ります。この手順書では伏せています。")
    ) @(28, 96)))
    $sheets.Add((New-Sheet "GUIでの実行" @(
        @("手順", "操作"),
        @("1", "RHELでは `$JMETER_HOME/bin/jmeter 、Windowsでは %JMETER_HOME%\bin\jmeter.bat を起動します。"),
        @("2", "File、Open から生成された rest-api-test.jmx を開きます。"),
        @("3", "Thread Group REST API のスレッド数とループ回数が意図どおりか確認します。初期値はどちらも 1 です。"),
        @("4", "各 HTTP Request の Server Name、Path、Method、Parameters または Body Data、HTTP Headers を確認します。"),
        @("5", "緑の Start を押します。初期値では有効な各APIへ 1 リクエストずつ送信します。"),
        @("6", "View Results Tree と Summary Report を確認します。jmeter.log の API_RESULT 行も確認します。"),
        @("7", "Save JTL リスナーは " + $settings.JtlFile + " にCSV形式のJTLを書きます。作業ディレクトリは出力ディレクトリにしてください。")
    ) @(12, 96)))
    $sheets.Add((New-Sheet "CLI_RHEL9" @(
        @("項目", "内容"),
        @("前提", "export JMETER_HOME=/opt/apache-jmeter-5.6.3"),
        @("この出力向けのコマンド", $bash),
        @("生成と実行を一度に行う", "scripts/generate-jmeter.sh --input 入力.xlsx --output-dir 出力先 --run"),
        @("作業ディレクトリ", "JMX内のJTLパスは相対パスです。必ず出力ディレクトリで jmeter を起動します。")
    ) @(28, 110)))
    $sheets.Add((New-Sheet "CLI_Windows11" @(
        @("項目", "内容"),
        @("前提", "`$env:JMETER_HOME = 'C:\apache-jmeter-5.6.3'"),
        @("この出力向けのコマンド", $powershell),
        @("生成と実行を一度に行う", ".\scripts\generate-jmeter.ps1 -InputPath .\入力.xlsx -OutputDir .\出力先 -Run"),
        @("実行ポリシー", "powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\generate-jmeter.ps1 ...")
    ) @(28, 110)))
    $sheets.Add((New-Sheet "結果の見方" @(
        @("項目", "内容"),
        @("results.jtl", "非GUIの -l で作るCSVです。HTMLレポートの入力になります。既存ファイルは追記されるため、再実行前に削除します。"),
        @("listener-results.jtl", "JMXの Simple Data Writer が書くCSVです。"),
        @("主要列", "timeStamp, elapsed, label, responseCode, responseMessage, threadName, success, bytes, sentBytes, URL, Latency, Connect"),
        @("success", "期待ステータスを指定したAPIだけ、コード不一致で success=false になります。"),
        @("jmeter.log", "API_RESULT 行に name、code、success、elapsed_ms、bytes、url が出ます。"),
        @("html-report", "jmeter の -e -o が作るダッシュボードです。index.html を開きます。出力先は空である必要があります。")
    ) @(28, 96)))
    $sheets.Add((New-Sheet "入力定義" @(
        @("項目", "内容"),
        @("Excel", "シート名は API、クエリ、ヘッダ、フォーム、設定です。説明シートは無視します。1行目がヘッダです。"),
        @("API列", "有効、名前、メソッド、URL、ボディ、Content-Type、ボディ形式、タイムアウトms、期待ステータス、コメント"),
        @("クエリ列", "有効、API名、名前、値。"),
        @("ヘッダ列", "有効、API名、名前、値。"),
        @("フォーム列", "有効、API名、名前、値。POSTの application/x-www-form-urlencoded です。"),
        @("設定", "threads、loops、ramp_up_sec、follow_redirects、plan_name、on_error、jtl_file。初期値は threads=1、loops=1、ramp_up_sec=1 です。"),
        @("CSV", "UTF-8（BOM可）またはCP932。列は record_type,api_name,enabled,method,url,body,content_type,body_mode,timeout_ms,expected_status,key,value,comment"),
        @("メソッド", "GET と POST だけです。"),
        @("POSTのクエリ", "POSTではクエリをURLのPathに付けます。")
    ) @(28, 110)))
    $generated = New-Object System.Collections.Generic.List[object]
    $generated.Add(@("名前", "有効", "メソッド", "送信URL", "ボディ", "ヘッダ", "期待ステータス", "コメント"))
    foreach ($item in $resolvedItems) {
        $body = ""
        if ($null -ne $item.RawBody) { $body = [string]$item.RawBody }
        elseif ($item.Method -eq "POST") {
            $parts = @()
            foreach ($pair in (Get-Pairs (, $item.Arguments))) {
                if ($null -eq $pair -or $null -eq $pair.PSObject.Properties["Key"]) { continue }
                $parts += ($pair.Key + "=" + $pair.Value)
            }
            $body = $parts -join " | "
        }
        if ($body.Length -gt 4000) { $body = $body.Substring(0, 4000) + " ...(省略)" }
        $headerParts = @()
        foreach ($pair in (Get-Pairs (, $item.Headers))) {
            if ($null -eq $pair -or $null -eq $pair.PSObject.Properties["Key"]) { continue }
            $headerParts += ($pair.Key + ": " + (Get-MaskedSecret $pair.Key $pair.Value))
        }
        $enabledText = "false"; if ($item.Enabled) { $enabledText = "true" }
        $generatedRow = New-Object object[] 8
        $generatedRow[0] = [string]$item.Name
        $generatedRow[1] = [string]$enabledText
        $generatedRow[2] = [string]$item.Method
        $generatedRow[3] = [string](Get-DisplayUrl $item)
        $generatedRow[4] = [string]$body
        $generatedRow[5] = [string]($headerParts -join " | ")
        $generatedRow[6] = [string]$item.ExpectedStatus
        $generatedRow[7] = [string]$item.Comment
        $generated.Add($generatedRow)
    }
    $sheets.Add((New-Sheet "生成されたAPI" $generated.ToArray() @(18, 12, 12, 64, 42, 42, 16, 36)))
    $sheets.Add((New-Sheet "トラブルシューティング" @(
        @("症状", "対処"),
        @("python3 が見つからない", "RHEL 9.8 で sudo dnf install -y python3 を実行します。"),
        @("Java が見つからない", "java -version が 17 であることを確認します。Java 22 以降は JMeter 5.6.3 では使えません。"),
        @("jmeter が見つからない", "-JMeterPath か環境変数 JMETER_HOME を指定します。"),
        @("JTLに前回の結果が残る", "再実行前に results.jtl を削除します。-Run は実行前に削除します。"),
        @("HTMLレポートが作られない", "html-report ディレクトリを削除して再実行します。"),
        @("4xxなのに success が true", "期待ステータスを指定していないAPIは、応答を受け取れた時点で成功です。"),
        @("POSTのJSONがクエリになる", "ボディ形式を raw にし、Content-Type を application/json にします。")
    ) @(36, 96)))
    return $sheets
}

function New-Sheet {
    param([string]$Name, [object[]]$Rows, [object[]]$Widths)
    return [pscustomobject]@{ Name = $Name; Rows = $Rows; Widths = $Widths }
}

function ConvertTo-ColumnLetter {
    param([int]$Index)
    $letters = ""
    while ($Index -gt 0) {
        $Index--
        $letters = [char](65 + ($Index % 26)) + $letters
        $Index = [math]::Floor($Index / 26)
    }
    return $letters
}

function Get-SafeSheetName {
    param([string]$Name, [hashtable]$Used)
    $cleaned = [regex]::Replace($Name, "[:\\/?*\[\]]", "_").Trim().Trim("'")
    if ($cleaned -eq "") { $cleaned = "Sheet" }
    if ($cleaned.Length -gt 31) { $cleaned = $cleaned.Substring(0, 31) }
    $base = $cleaned
    $number = 2
    while ($Used.ContainsKey($cleaned.ToLower())) {
        $suffix = "_$number"
        $cleaned = $base.Substring(0, [Math]::Min($base.Length, 31 - $suffix.Length)) + $suffix
        $number++
    }
    $Used[$cleaned.ToLower()] = $true
    return $cleaned
}

function Get-WorksheetXml {
    param($Sheet)
    $rows = $Sheet.Rows
    $width = 1
    foreach ($row in $rows) {
        $count = 0
        foreach ($cell in $row) { $count++ }
        if ($count -gt $width) { $width = $count }
    }
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    $lines.Add('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    $lines.Add('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>')
    for ($index = 1; $index -le $width; $index++) {
        $size = 24
        if (($index - 1) -lt @($Sheet.Widths).Count) { $size = @($Sheet.Widths)[$index - 1] }
        $lines.Add('<col min="' + $index + '" max="' + $index + '" width="' + $size + '" customWidth="1"/>')
    }
    $lines.Add("</cols><sheetData>")
    $rowNumber = 1
    foreach ($row in $rows) {
        $lines.Add('<row r="' + $rowNumber + '">')
        $colNumber = 1
        foreach ($value in $row) {
            if ($null -ne $value -and [string]$value -ne "") {
                $style = "0"
                if ($rowNumber -eq 1) { $style = "1" } elseif ($colNumber -eq 2) { $style = "2" }
                $ref = (ConvertTo-ColumnLetter $colNumber) + $rowNumber
                $lines.Add('<c r="' + $ref + '" t="inlineStr" s="' + $style + '"><is><t xml:space="preserve">' + (ConvertTo-XmlText $value) + '</t></is></c>')
            }
            $colNumber++
        }
        $lines.Add("</row>")
        $rowNumber++
    }
    $lines.Add("</sheetData>")
    if ($rows.Count -gt 0) { $lines.Add('<autoFilter ref="A1:' + (ConvertTo-ColumnLetter $width) + ($rows.Count) + '"/>') }
    $lines.Add("</worksheet>")
    return ($lines -join "`n") + "`n"
}

function Write-XlsxBook {
    param([string]$Path, [object[]]$Sheets)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $directory = [System.IO.Path]::GetDirectoryName($Path)
    if ($directory -and -not (Test-Path -LiteralPath $directory)) { New-Item -ItemType Directory -Force -Path $directory | Out-Null }
    $used = @{}
    $safeNames = @()
    foreach ($sheet in @($Sheets)) { $safeNames += ,(Get-SafeSheetName $sheet.Name $used) }
    $sheetXml = @()
    $relXml = @('<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    for ($i = 0; $i -lt $safeNames.Count; $i++) {
        $number = $i + 1
        $sheetXml += '<sheet name="' + (ConvertTo-XmlText $safeNames[$i]) + '" sheetId="' + $number + '" r:id="rId' + $number + '"/>'
        $relXml += '<Relationship Id="rId' + $number + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $number + '.xml"/>'
    }
    $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + ($sheetXml -join "") + '</sheets></workbook>' + "`n"
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ($relXml -join "") + '</Relationships>' + "`n"
    $rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' + "`n"
    $overrides = @(
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    )
    for ($i = 1; $i -le @($Sheets).Count; $i++) {
        $overrides += '<Override PartName="/xl/worksheets/sheet' + $i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    }
    $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + ($overrides -join "") + '</Types>' + "`n"
    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor rgb="FF1F4E79"/></patternFill></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs>
</styleSheet>
'@
    $temp = $Path + ".tmp"
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
    $zip = [System.IO.Compression.ZipFile]::Open($temp, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        Add-ZipEntry $zip "[Content_Types].xml" $content
        Add-ZipEntry $zip "_rels/.rels" $rootRels
        Add-ZipEntry $zip "xl/workbook.xml" $workbook
        Add-ZipEntry $zip "xl/_rels/workbook.xml.rels" $rels
        Add-ZipEntry $zip "xl/styles.xml" ($styles + "`n")
        for ($i = 0; $i -lt @($Sheets).Count; $i++) { Add-ZipEntry $zip ("xl/worksheets/sheet" + ($i + 1) + ".xml") (Get-WorksheetXml @($Sheets)[$i]) }
    } finally { $zip.Dispose() }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    Move-Item -LiteralPath $temp -Destination $Path
}

function Add-ZipEntry {
    param($Zip, [string]$Name, [string]$Text)
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $bytes = (New-Object System.Text.UTF8Encoding $false).GetBytes($Text)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
}

function Write-Templates {
    param([string]$Directory)
    if (-not (Test-Path -LiteralPath $Directory)) { New-Item -ItemType Directory -Force -Path $Directory | Out-Null }
    $csvPath = Join-Path $Directory "api-input-template.csv"
    $xlsxPath = Join-Path $Directory "api-input-template.xlsx"
    $rows = @(
        @("record_type","api_name","enabled","method","url","body","content_type","body_mode","timeout_ms","expected_status","key","value","comment"),
        @("API","get-item","true","GET","https://example.com/v1/items","","","","","200","","","サンプルGET。実行前にURLを置き換えてください。"),
        @("QUERY","get-item","true","","","","","","","","page","1",""),
        @("QUERY","get-item","true","","","","","","","","size","10",""),
        @("HEADER","get-item","true","","","","","","","","Accept","application/json",""),
        @("HEADER","get-item","true","","","","","","","","X-Request-Id","demo-001",""),
        @("API","create-item","true","POST","https://example.com/v1/items",'{"name":"sample","count":1}',"application/json","raw","","","","","サンプルPOST(JSON)。"),
        @("HEADER","create-item","true","","","","","","","","Accept","application/json",""),
        @("API","submit-form","true","POST","https://example.com/v1/forms","","","form","","","","","サンプルPOST(フォーム)。"),
        @("BODY","submit-form","true","","","","","","","","title","hello",""),
        @("BODY","submit-form","true","","","","","","","","note","world",""),
        @("HEADER","submit-form","true","","","","","","","","Accept","application/json",""),
        @("SETTING","","true","","","","","","","","threads","1","初期値は1スレッドです。各APIを1回呼びます。"),
        @("SETTING","","true","","","","","","","","loops","1","初期値は1ループです。"),
        @("SETTING","","true","","","","","","","","ramp_up_sec","1",""),
        @("SETTING","","true","","","","","","","","follow_redirects","true",""),
        @("SETTING","","true","","","","","","","","plan_name",$Script:DefaultPlanName,""),
        @("SETTING","","true","","","","","","","","on_error","continue","continue / stopthread / stoptest / stoptestnow")
    )
    $builder = New-Object System.Text.StringBuilder
    foreach ($row in $rows) {
        $fields = @()
        foreach ($cell in @($row)) {
            $text = [string]$cell
            if ($text.Contains('"') -or $text.Contains(",") -or $text.Contains("`n") -or $text.Contains("`r")) {
                $fields += '"' + $text.Replace('"', '""') + '"'
            } else { $fields += $text }
        }
        [void]$builder.Append(($fields -join ",") + "`r`n")
    }
    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($csvPath, $builder.ToString(), $utf8Bom)
    $manual = @(
        @("項目","内容"),
        @("このファイル","JMeter 5.6.3 のテスト計画を作るための入力テンプレートです。"),
        @("編集するシート","API、クエリ、ヘッダ、フォーム、設定。この説明シートは読み込まれません。"),
        @("初期負荷","設定シートの threads=1、loops=1 が初期値です。有効な各APIを1回呼びます。")
    )
    $apiRows = @(
        @("有効","名前","メソッド","URL","ボディ","Content-Type","ボディ形式","タイムアウトms","期待ステータス","コメント"),
        @("true","get-item","GET","https://example.com/v1/items","","","","","200","サンプルGET。実行前にURLを置き換えてください。"),
        @("true","create-item","POST","https://example.com/v1/items",'{"name":"sample","count":1}',"application/json","raw","","","サンプルPOST(JSON)。"),
        @("true","submit-form","POST","https://example.com/v1/forms","","","form","","","サンプルPOST(フォーム)。")
    )
    $queries = @(@("有効","API名","名前","値"), @("true","get-item","page","1"), @("true","get-item","size","10"))
    $headers = @(
        @("有効","API名","名前","値"),
        @("true","get-item","Accept","application/json"),
        @("true","get-item","X-Request-Id","demo-001"),
        @("true","create-item","Accept","application/json"),
        @("true","submit-form","Accept","application/json")
    )
    $forms = @(@("有効","API名","名前","値"), @("true","submit-form","title","hello"), @("true","submit-form","note","world"))
    $settings = @(
        @("キー","値","コメント"),
        @("threads","1","初期値は1スレッドです。各APIを1回呼びます。"),
        @("loops","1","初期値は1ループです。"),
        @("ramp_up_sec","1",""),
        @("follow_redirects","true",""),
        @("plan_name",$Script:DefaultPlanName,""),
        @("on_error","continue","continue / stopthread / stoptest / stoptestnow")
    )
    $book = @(
        (New-Sheet "説明" $manual @(28, 88)),
        (New-Sheet "API" $apiRows @(12, 18, 12, 42, 36, 24, 14, 16, 16, 42)),
        (New-Sheet "クエリ" $queries @(12, 18, 18, 24)),
        (New-Sheet "ヘッダ" $headers @(12, 18, 22, 28)),
        (New-Sheet "フォーム" $forms @(12, 18, 18, 24)),
        (New-Sheet "設定" $settings @(24, 28, 56))
    )
    Write-XlsxBook $xlsxPath $book
    return @($csvPath, $xlsxPath)
}

function Export-PlanFiles {
    param($Plan, [string]$Destination, $ThreadOverride, $LoopOverride, $RampOverride)
    if ($null -ne $ThreadOverride) { $Plan.Settings.Threads = ConvertTo-IntRange $ThreadOverride "threads" 1 1000000 }
    if ($null -ne $LoopOverride) { $Plan.Settings.Loops = ConvertTo-IntRange $LoopOverride "loops" 1 1000000 }
    if ($null -ne $RampOverride) { $Plan.Settings.RampUpSec = ConvertTo-IntRange $RampOverride "ramp-up" 0 86400 }
    $resolved = New-Object System.Collections.Generic.List[object]
    foreach ($api in $Plan.Apis) { $resolved.Add((Resolve-ApiRequest $api (, $Plan.Warnings))) }
    $any = $false
    foreach ($item in $resolved) { if ($item.Enabled) { $any = $true } }
    if (-not $any) { throw "有効なAPIが1件もありません。" }
    if (-not (Test-Path -LiteralPath $Destination)) { New-Item -ItemType Directory -Force -Path $Destination | Out-Null }
    $full = [System.IO.Path]::GetFullPath($Destination)
    $jmxPath = Join-Path $full $Script:JmxFileName
    $usagePath = Join-Path $full $Script:UsageFileName
    $xml = New-JmxText $Plan (, $resolved)
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($jmxPath, $xml.Replace("`r`n", "`n"), $utf8)
    Write-XlsxBook $usagePath (New-UsageSheets $Plan (, $resolved) $full)
    foreach ($warning in $Plan.Warnings) { [Console]::Error.WriteLine("warning: $warning") }
    $enabledCount = 0
    foreach ($item in $resolved) { if ($item.Enabled) { $enabledCount++ } }
    [Console]::Out.WriteLine("input: " + $Plan.SourcePath)
    [Console]::Out.WriteLine("apis: " + $Plan.Apis.Count)
    [Console]::Out.WriteLine("enabled: " + $enabledCount)
    [Console]::Out.WriteLine("threads: " + $Plan.Settings.Threads)
    [Console]::Out.WriteLine("loops: " + $Plan.Settings.Loops)
    [Console]::Out.WriteLine("jmx: " + $jmxPath)
    [Console]::Out.WriteLine("usage: " + $usagePath)
    return $full
}

function Find-JMeterBinary {
    param([string]$Explicit)
    if ($Explicit) {
        if (-not (Test-Path -LiteralPath $Explicit)) { throw "JMeterの実行ファイルがありません: $Explicit" }
        return [System.IO.Path]::GetFullPath($Explicit)
    }
    if ($env:JMETER_HOME) {
        $candidate = Join-Path $env:JMETER_HOME "bin\jmeter.bat"
        if (Test-Path -LiteralPath $candidate) { return [System.IO.Path]::GetFullPath($candidate) }
    }
    $command = Get-Command jmeter.bat -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    $command = Get-Command jmeter -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw "JMeter 5.6.3 が見つかりません。-JMeterPath か環境変数 JMETER_HOME を指定してください。"
}

function Assert-Java17 {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "java"
    $psi.Arguments = "-version"
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    try { $process = [System.Diagnostics.Process]::Start($psi) } catch { throw "Java が見つかりません。JMeter 5.6.3 には Java 17 を推奨します。" }
    $stderr = $process.StandardError.ReadToEnd()
    $stdout = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()
    $text = $stderr + $stdout
    $match = [regex]::Match($text, 'version "(?:1\.)?(\d+)')
    if (-not $match.Success) { throw "java -version の結果を解釈できません。Java 17 をインストールしてください。" }
    $major = [int]$match.Groups[1].Value
    if ($major -lt 8 -or $major -gt 21) { throw "JMeter 5.6.3 は Java 8 から 21 で動作します。検出したメジャー版は $major です。Java 17 を使ってください。" }
}

function Invoke-JMeterPlan {
    param([string]$Binary, [string]$Destination)
    Assert-Java17
    $results = Join-Path $Destination "results"
    if (-not (Test-Path -LiteralPath $results)) { New-Item -ItemType Directory -Force -Path $results | Out-Null }
    $jtl = Join-Path $results "results.jtl"
    $logPath = Join-Path $results "jmeter.log"
    $report = Join-Path $results "html-report"
    $listener = Join-Path $results "listener-results.jtl"
    foreach ($path in @($jtl, $logPath, $listener)) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force } }
    if (Test-Path -LiteralPath $report) { Remove-Item -LiteralPath $report -Recurse -Force }
    $props = Join-Path $results "saveservice.properties"
    $propText = @"
jmeter.save.saveservice.output_format=csv
jmeter.save.saveservice.print_field_names=true
jmeter.save.saveservice.default_delimiter=,
jmeter.save.saveservice.bytes=true
jmeter.save.saveservice.sent_bytes=true
jmeter.save.saveservice.label=true
jmeter.save.saveservice.latency=true
jmeter.save.saveservice.response_code=true
jmeter.save.saveservice.response_message=true
jmeter.save.saveservice.successful=true
jmeter.save.saveservice.thread_counts=true
jmeter.save.saveservice.thread_name=true
jmeter.save.saveservice.time=true
jmeter.save.saveservice.connect_time=true
jmeter.save.saveservice.url=true
jmeter.save.saveservice.timestamp_format=ms
"@
    [System.IO.File]::WriteAllText($props, $propText.Replace("`r`n", "`n"), (New-Object System.Text.UTF8Encoding $false))
    $jmx = Join-Path $Destination $Script:JmxFileName
    Push-Location $Destination
    try {
        & $Binary -n -q $props -t $jmx -l $jtl -j $logPath -e -o $report
        $code = $LASTEXITCODE
    } finally { Pop-Location }
    if (-not (Test-Path -LiteralPath $jtl)) {
        [Console]::Error.WriteLine("error: JTLが出力されませんでした。jmeter.log を確認してください。")
        if ($code) { return $code }
        return 1
    }
    $rows = @(Import-Csv -LiteralPath $jtl)
    $success = 0
    $failed = 0
    foreach ($row in $rows) {
        if ([string]$row.success -eq "true") { $success++ } else { $failed++ }
    }
    [Console]::Out.WriteLine("samples: " + $rows.Count)
    [Console]::Out.WriteLine("success: " + $success)
    [Console]::Out.WriteLine("failed: " + $failed)
    [Console]::Out.WriteLine("jtl: " + $jtl)
    [Console]::Out.WriteLine("log: " + $logPath)
    $index = Join-Path $report "index.html"
    if (Test-Path -LiteralPath $index) { [Console]::Out.WriteLine("html: " + $index) }
    else { [Console]::Error.WriteLine("warning: HTMLレポートがありません。") }
    if ($null -eq $code) { return 0 }
    return $code
}

$threadOverride = $null
$loopOverride = $null
$rampOverride = $null
if ($PSBoundParameters.ContainsKey("Threads")) { $threadOverride = $Threads }
if ($PSBoundParameters.ContainsKey("Loops")) { $loopOverride = $Loops }
if ($PSBoundParameters.ContainsKey("RampUp")) { $rampOverride = $RampUp }

try {
    if ($InitTemplates) {
        if (-not $OutputDir) { throw "テンプレートの出力先を -OutputDir で指定してください。" }
        $written = Write-Templates $OutputDir
        Write-Output ("template_csv: " + [System.IO.Path]::GetFullPath($written[0]))
        Write-Output ("template_xlsx: " + [System.IO.Path]::GetFullPath($written[1]))
        exit 0
    }
    if (-not $InputPath -or -not $OutputDir) {
        Show-GeneratorHelp
        exit 2
    }
    $plan = Import-ApiFile $InputPath
    $destination = Export-PlanFiles $plan $OutputDir $threadOverride $loopOverride $rampOverride
    if ($Run) {
        $binary = Find-JMeterBinary $JMeterPath
        $runCode = Invoke-JMeterPlan $binary $destination
        exit $runCode
    }
    exit 0
} catch {
    [Console]::Error.WriteLine("error: " + $_.Exception.Message)
    if ($_.InvocationInfo.PositionMessage) { [Console]::Error.WriteLine($_.InvocationInfo.PositionMessage) }
    exit 2
}
