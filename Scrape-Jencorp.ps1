# Scrape-Jencorp.ps1
# - 輸出：data\jencorp_master.csv（依ロット番号更新/追加，不覆蓋）
# - 快照：data\snapshots\jencorp_yyyyMMdd_HHmmss.csv
# - 時序：data\jencorp_timeseries.csv（每次執行逐筆追加）
$ErrorActionPreference = "Stop"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try { [System.Threading.Thread]::CurrentThread.CurrentCulture=[System.Globalization.CultureInfo]::GetCultureInfo("ja-JP"); [System.Threading.Thread]::CurrentThread.CurrentUICulture=[System.Globalization.CultureInfo]::GetCultureInfo("ja-JP") } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$csvEncoding = if ($PSVersionTable.PSVersion.Major -ge 7) { "utf8BOM" } else { "utf8" }

# 參數
$BaseUrl = "https://www.jencorp.net/jp/net/"
$QueryCommon = @{ id=""; MODEL_ITEM_LIMIT="30"; MODEL_ITEM_PAGE="10"; MODEL_ITEM_SORT="7"; MODEL_ITEM_S_AUCNUM=""; MODEL_ITEM_S_DELIYARD=""; MODEL_ITEM_S_MODEL=""; MODEL_ITEM_S_AUCEND_Y=""; MODEL_ITEM_S_AUCEND_M=""; MODEL_ITEM_S_AUCEND_D=""; MODEL_ITEM_S_ICONTYPE=""; MODEL_ITEM_S_THUMBNAIL=""; MODEL_ITEM_S_FAVORITE=""; MODEL_ITEM_S_BIDDING="" }
$MaxPages=999; $DelayMs=350; $Concurrency=7
$headersReq=@{ "Accept-Language"="ja,en;q=0.8"; "Referer"="https://www.jencorp.net/" }

# 監測（最後 5 分鐘）
$EnableSnipe=$true; $SnipeWindowMin=5; $SnipePollSec=10; $SnipeStableRounds=3

# 路徑
$RepoRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$DataDir  = Join-Path $RepoRoot "data"; $SnapDir = Join-Path $DataDir "snapshots"
if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Force -Path $DataDir | Out-Null }
if (-not (Test-Path $SnapDir)) { New-Item -ItemType Directory -Force -Path $SnapDir | Out-Null }
$MasterCsv = Join-Path $DataDir "jencorp_master.csv"
$TimeSeriesCsv = Join-Path $DataDir "jencorp_timeseries.csv"
$RunStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$RunCsv = Join-Path $SnapDir ("jencorp_{0}.csv" -f $RunStamp)
$NowStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

# 工具
Add-Type -AssemblyName System.Web | Out-Null
function Decode($s){ [System.Web.HttpUtility]::HtmlDecode($s) }
function Get-WebString([string]$Url,[hashtable]$Headers=$null,[int]$TimeoutSec=30){
  $resp = Invoke-WebRequest -Uri $Url -Headers $Headers -UseBasicParsing -TimeoutSec $TimeoutSec -MaximumRedirection 5 -UserAgent "Mozilla/5.0"
  $stream=$resp.RawContentStream
  if ($null -eq $stream) { return [string]$resp.Content }
  if ($stream.CanSeek) { $stream.Position=0 }
  $ms=New-Object System.IO.MemoryStream; $stream.CopyTo($ms); $bytes=$ms.ToArray()
  $charset=""; $ct=$resp.Headers["Content-Type"]
  if ($ct -and $ct -match "(?i)charset=([A-Za-z0-9\-_]+)") { $charset=$Matches[1] }
  $sample=[System.Text.Encoding]::ASCII.GetString($bytes,0,[Math]::Min($bytes.Length,4096))
  if (-not $charset -and $sample -match "(?i)<meta[^>]+charset=['""]?([\w\-]+)") { $charset=$Matches[1] }
  if (-not $charset -and $sample -match "(?i)content=['"'][^'""]*charset=([\w\-]+)") { $charset=$Matches[1] }
  switch -Regex ($charset) {
    "shift[_\-]?jis|sjis|cp932" { $enc=[System.Text.Encoding]::GetEncoding(932) }
    "euc[_\-]?jp"               { $enc=[System.Text.Encoding]::GetEncoding("euc-jp") }
    "utf[_\-]?8"                { $enc=[System.Text.Encoding]::UTF8 }
    default                     { $enc=[System.Text.Encoding]::UTF8 }
  }
  $text=$enc.GetString($bytes); if ($text -match "Ã.|ã.") { $text=[System.Text.Encoding]::UTF8.GetString($bytes) }
  return $text
}
function CleanHtmlText([string]$html){ if ([string]::IsNullOrWhiteSpace($html)) { return "" }; $t=$html -replace "(?i)<br\s*/?>","`n"; $t=$t -replace "<[^>]*>",""; $t=Decode $t; $t=$t -replace "\s+"," "; $t.Trim() }
function Rx1([string]$text,[string]$pattern){ $m=[regex]::Match($text,$pattern,"Singleline,IgnoreCase"); if($m.Success){$m.Groups[1].Value.Trim()} else {""} }
function Build-Url([int]$offset,[string]$baseUrl,[hashtable]$queryCommon){ $q=[ordered]@{}; foreach($k in $queryCommon.Keys){ $q[$k]=$queryCommon[$k] }; $q["MODEL_ITEM_OFFSET"]=[string]$offset; $qs=($q.GetEnumerator()|%{ '{0}={1}' -f $_.Key,[uri]::EscapeDataString([string]$_.Value) }) -join '&'; $ub=[System.UriBuilder]::new($baseUrl); $ub.Query=$qs; $ub.Uri.AbsoluteUri }
function Parse-RemainingToSeconds([string]$text){ if ([string]::IsNullOrWhiteSpace($text)) { return $null }; $d=0;$h=0;$m=0; if ($text -match '(?i)(\d+)\s*d'){ $d=[int]$Matches[1] } ; if ($text -match '(?i)(\d+)\s*h'){ $h=[int]$Matches[1] } ; if ($text -match '(?i)(\d+)\s*m'){ $m=[int]$Matches[1] } ; ($d*86400+$h*3600+$m*60) }
function Get-DetailFields([string]$detailUrl,[hashtable]$headers){
  $result=[ordered]@{ MAKE=""; "START PRICE (JPY)"=""; "機種分類"=""; "結標金額"=""; "現在価格_from_detail"=""; "入札數_from_detail"=""; "残り時間_from_detail"="" }
  if ([string]::IsNullOrWhiteSpace($detailUrl)) { return $result }
  try{
    $html = Get-WebString $detailUrl $headers 30
    $result.MAKE = CleanHtmlText (Rx1 $html "<p\s+class=""maker"">\s*([^<]+)\s*</p>")
    $result."START PRICE (JPY)" = (Rx1 $html "(?:開始価格|出品開始価格|Start\s*Price)[^<]*</[^>]*>\s*<[^>]*>\s*JPY?\s*([\d,]+)") -replace "[^\d]",""
    $result."機種分類" = CleanHtmlText (Rx1 $html "(?:カテゴリ|カテゴリー|機種)[^<]*</[^>]*>\s*<[^>]*>\s*([^<]+)<")
    $result."結標金額" = (Rx1 $html "(?:落札価格|結標金額|Final\s*Price)[^<]*</[^>]*>\s*<[^>]*>\s*JPY?\s*([\d,]+)") -replace "[^\d]",""
    $curD  = Rx1 $html "(?:現在価格|Current\s*Bid)[^<]*</[^>]*>\s*<[^>]*>\s*JPY?\s*([\d,]+)"; if ($curD){ $result."現在価格_from_detail" = ($curD -replace "[^\d]","") }
    $bidsD = Rx1 $html "(?:入札件数|Bids?)[^<]*</[^>]*>\s*<[^>]*>\s*([\d,]+)"; if ($bidsD){ $result."入札數_from_detail" = ($bidsD -replace "[^\d]","") }
    $remainD = Rx1 $html "<dd\s+class=""time_left"">\s*([^<]+)"; if ($remainD){ $result."残り時間_from_detail" = CleanHtmlText $remainD }
  } catch {}
  $result
}

# 抓頁（Start-Job 平行）
$allPageResults = New-Object System.Collections.Generic.List[object]
$stop=$false; $pagesCompleted=0
for ($start=1; $start -le $MaxPages -and -not $stop; $start += $Concurrency) {
  $batch = $start..([Math]::Min($MaxPages,$start+$Concurrency-1))
  Write-Host ("批次頁碼：{0}" -f ($batch -join ', '))
  $jobs=@()
  foreach($offset in $batch){
    $jobs += Start-Job -ScriptBlock {
      param($offset,$BaseUrl,$QueryCommon,$headersReq,$DelayMs)
      Add-Type -AssemblyName System.Web | Out-Null
      function Decode([string]$s){ [System.Web.HttpUtility]::HtmlDecode($s) }
      function Get-WebString([string]$Url,[hashtable]$Headers,[int]$TimeoutSec){ ${function:Get-WebString}.Invoke($Url,$Headers,$TimeoutSec) }
      function CleanHtmlText([string]$html){ if ([string]::IsNullOrWhiteSpace($html)) { return "" }; $t=$html -replace "(?i)<br\s*/?>","`n"; $t=$t -replace "<[^>]*>",""; $t=[System.Web.HttpUtility]::HtmlDecode($t); $t=$t -replace "\s+"," "; $t.Trim() }
      function Rx1([string]$text,[string]$pattern){ $m=[regex]::Match($text,$pattern,"Singleline,IgnoreCase"); if($m.Success){$m.Groups[1].Value.Trim()} else {""} }
      function Build-Url([int]$o,[string]$baseUrl,[hashtable]$queryCommon){ ${function:Build-Url}.Invoke($o,$baseUrl,$queryCommon) }
      function Get-DetailFields([string]$detailUrl,[hashtable]$headers){ ${function:Get-DetailFields}.Invoke($detailUrl,$headers) }
      function Parse-RemainingToSeconds([string]$text){ ${function:Parse-RemainingToSeconds}.Invoke($text) }
      $url = (Build-Url -offset $offset -baseUrl $BaseUrl -queryCommon $QueryCommon)
      try { $page = (Invoke-WebRequest -Uri $url -Headers $headersReq -UseBasicParsing -TimeoutSec 30 -MaximumRedirection 5 -UserAgent "Mozilla/5.0").Content } catch {
        Start-Sleep -Milliseconds $DelayMs
        return [pscustomobject]@{ Offset=$offset; PageCount=0; Items=@(); Url=$url; Error=$_.Exception.Message }
      }
      if ([string]::IsNullOrWhiteSpace($page)) {
        Start-Sleep -Milliseconds $DelayMs
        return [pscustomobject]@{ Offset=$offset; PageCount=0; Items=@(); Url=$url; Error="Empty page" }
      }
      $pattern = '<div\s+class="main_info_cont">.*?(?=<div\s+class="main_info_cont"|</main>|\Z)'
      $matches = [regex]::Matches($page,$pattern,'Singleline,IgnoreCase')
      if ($matches.Count -le 0) {
        Start-Sleep -Milliseconds $DelayMs
        return [pscustomobject]@{ Offset=$offset; PageCount=0; Items=@(); Url=$url; Error=$null }
      }
      $rows = New-Object System.Collections.Generic.List[object]
      foreach($m in $matches){
        $b=$m.Value
        $model = Rx1 $b "<h3\s+class=""model"">\s*(.*?)\s*</h3>"
        $lot   = Rx1 $b "ロット番号.*?<dd\s+class=""ac_num""[^>]*>.*?>(.*?)</a>"
        if (-not $lot) { continue } # ロット番号作為唯一鍵
        $year  = Rx1 $b "<dd\s+class=""year"">\s*([^<]+)\s*</dd>"
        $yardBlock = Rx1 $b "<dl\s+class=""delivery_yard_wrap"">.*?<dd\s+class=""delivery_yard"">\s*(.*?)\s*</dd>.*?</dl>"
        if (-not $yardBlock){ $yardBlock = Rx1 $b "<dd\s+class=""delivery_yard"">\s*(.*?)\s*</dd>" }
        $yard = CleanHtmlText $yardBlock
        $specHtml = Rx1 $b "<p\s+class=""feature_and_comment"">\s*(.*?)\s*</p>"
        $spec = CleanHtmlText $specHtml
        $meter = (Rx1 $b "<dd\s+class=""Meter_reads"">\s*([^<]+)\s*</dd>") -replace "[^\d]",""
        $current = Rx1 $b "<dd\s+class=""current_bid"">.*?JPY\s*([\d,]+)"
        $bids    = Rx1 $b "<dd\s+class=""bids"">\s*([\d,]+)"
        $remain  = Rx1 $b "<dd\s+class=""time_left"">\s*([^<]+)"
        $detail1 = Rx1 $b "<a\s+href=""([^""]+?/model/detail\.html\?[^""]+)""\s+class=""model_link"""
        $detail2 = Rx1 $b "<dd\s+class=""ac_num""><a\s+href=""([^""]+?/model/detail\.html\?[^""]+)"""
        $detailUrl = if ($detail1) { $detail1 } elseif ($detail2) { $detail2 } else { "" }
        $more = Get-DetailFields $detailUrl $headersReq
        $remainRaw = CleanHtmlText $remain
        $remainSec = Parse-RemainingToSeconds $remainRaw
        $row = [pscustomobject]([ordered]@{
          "渡地 (EX-YARD)"=$yard; "MAKE"=$more.MAKE; "MODEL"=(CleanHtmlText $model); "YEAR"=($year -replace "[^\d]","");
          "SPECIFICATION"=$spec; "メーター"=$meter; "START PRICE (JPY)"=$more."START PRICE (JPY)"; "機種分類"=$more."機種分類"; "結標金額"=$more."結標金額";
          "ロット番号"=$lot; "詳細URL"=$detailUrl; "現在価格 (JPY)"=($current -replace "[^\d]",""); "入札數"=($bids -replace "[^\d]","");
          "残り時間_raw"=$remainRaw; "残り秒"=$remainSec; "現在価格_from_detail"=$more."現在価格_from_detail"; "入札數_from_detail"=$more."入札數_from_detail"; "残り時間_from_detail"=$more."残り時間_from_detail"
        })
        $rows.Add($row) | Out-Null
      }
      Start-Sleep -Milliseconds $DelayMs
      [pscustomobject]@{ Offset=$offset; PageCount=$matches.Count; Items=$rows; Url=$url; Error=$null }
    } -ArgumentList $offset,$BaseUrl,$QueryCommon,$headersReq,$DelayMs
  }
  while ( ($jobs | ? { $_.State -notin @('Completed','Failed','Stopped') }).Count -gt 0 ) {
    $done =   ($jobs | ? { $_.State -in     @('Completed','Failed','Stopped') }).Count
    $pctBatch = if ($jobs.Count -gt 0){ [math]::Min(100,[math]::Round(($done/$jobs.Count)*100,0)) } else { 100 }
    Write-Progress -Id 1 -Activity "批次進度" -Status ("完成 {0}/{1}" -f $done,$jobs.Count) -PercentComplete $pctBatch
    Start-Sleep -Milliseconds 200
  }
  Wait-Job -Job $jobs | Out-Null
  $batchResults = foreach($j in $jobs){ try { Receive-Job -Job $j -ErrorAction Stop } catch { $null } finally { Remove-Job -Job $j -Force -ErrorAction SilentlyContinue } }
  foreach($r in (@($batchResults) | ? { $_ -ne $null } | Sort-Object Offset)){ $allPageResults.Add($r) | Out-Null }
  if ($batchResults | ? { $_.PageCount -le 0 }) { $stop=$true }
}

# 全部商品
$items = New-Object System.Collections.Generic.List[object]
foreach($r in ($allPageResults | Sort-Object Offset)){ foreach($row in $r.Items){ $items.Add($row) | Out-Null } }

# 短期監測（<=5 分）
if ($EnableSnipe){
  $targets = $items | ? { $_."詳細URL" -and $_."残り秒" -ne $null -and [int]$_."残り秒" -le ($SnipeWindowMin*60) }
  if ($targets){
    $jobsS=@()
    foreach($row in $targets){
      $jobsS += Start-Job -ScriptBlock {
        param($detailUrl,$pollSec,$stable,$headers)
        function Rx1([string]$text,[string]$pattern){ $m=[regex]::Match($text,$pattern,"Singleline,IgnoreCase"); if($m.Success){$m.Groups[1].Value.Trim()} else {""} }
        $last=""; $same=0; $max=0; $ended=$false
        while($true){
          try{$html=(Invoke-WebRequest -Uri $detailUrl -Headers $headers -UseBasicParsing -MaximumRedirection 5 -UserAgent "Mozilla/5.0").Content}catch{ Start-Sleep -Seconds $pollSec; continue }
          $price = (Rx1 $html "(?:落札価格|Final\s*Price|現在価格|Current\s*Bid)[^<]*</[^>]*>\s*<[^>]*>\s*JPY?\s*([\d,]+)") -replace "[^\d]",""
          if ($price){ $max=[math]::Max($max,[int]$price) }
          $remain = Rx1 $html "<dd\s+class=""time_left"">\s*([^<]+)"
          if ($price -eq $last -and $price){ $same++ } else { $same=0; $last=$price }
          if ($html -match "(落札|終了|closed|ended)" -or [string]::IsNullOrWhiteSpace($remain) -or $remain -match '^\s*0m\s*$' -or ($same -ge $stable -and $last)){ break }
          Start-Sleep -Seconds $pollSec
        }
        [pscustomobject]@{ Url=$detailUrl; FinalPrice = (if($last){$last}else{$max.ToString()}); Ended=$true }
      } -ArgumentList $row."詳細URL",$SnipePollSec,$SnipeStableRounds,$headersReq
    }
    Wait-Job -Job $jobsS | Out-Null
    $results = foreach($j in $jobsS){ try{ Receive-Job -Job $j -ErrorAction Stop } catch{} finally{ Remove-Job -Job $j -Force -ErrorAction SilentlyContinue } }
    $map=@{}; foreach($r in $results){ $map[$r.Url]=$r }
    foreach($row in $items){
      $u=$row."詳細URL"; if ($u -and $map.ContainsKey($u)){ $r=$map[$u]; if ($r.FinalPrice){ $row | Add-Member -NotePropertyName "最終監測價 (JPY)" -NotePropertyValue $r.FinalPrice -Force }; $row | Add-Member -NotePropertyName "是否結標(推測)" -NotePropertyValue ($r.Ended ? "是" : "否") -Force }
    }
  }
}

# Master 合併（ロット番号為鍵）
function KeyFrom($row){ if ($row."ロット番号"){ "LOT:"+[string]$row."ロット番号" } else { "" } }
$dynamicCols = @("現在価格 (JPY)","入札數","残り時間_raw","残り秒","結標金額","現在価格_from_detail","入札數_from_detail","残り時間_from_detail","最終監測價 (JPY)","是否結標(推測)")
$fixedCols   = @("渡地 (EX-YARD)","MAKE","MODEL","YEAR","SPECIFICATION","メーター","START PRICE (JPY)","機種分類","ロット番号","詳細URL")
$masterRowsList = New-Object System.Collections.Generic.List[psobject]; $masterDict=@{}
if (Test-Path $MasterCsv){ try{ $loaded=Import-Csv $MasterCsv; foreach($mr in $loaded){ [void]$masterRowsList.Add($mr); if ($mr."ロット番号"){ $masterDict["LOT:"+[string]$mr."ロット番号"]=$mr } } }catch{} }
foreach($row in $items){
  $key = KeyFrom $row; if (-not $key) { continue }
  if ($masterDict.ContainsKey($key)){
    $mr=$masterDict[$key]
    foreach($col in $dynamicCols){ if ($row.PSObject.Properties.Name -contains $col){ $val=$row.$col; if ($val -ne $null -and $val -ne ""){ $mr | Add-Member -NotePropertyName $col -NotePropertyValue $val -Force } } }
    $mr | Add-Member -NotePropertyName "最後更新時間" -NotePropertyValue $NowStr -Force
  } else {
    $nr=[ordered]@{}; foreach($c in $fixedCols + $dynamicCols){ $nr[$c] = ($row.PSObject.Properties.Name -contains $c) ? $row.$c : $null }
    $nr["首次抓取時間"]=$NowStr; $nr["最後更新時間"]=$NowStr; $nr["狀態"]= if ($row."結標金額"){ "結標" } else { "進行中" }
    $obj=New-Object psobject -Property $nr; [void]$masterRowsList.Add($obj); $masterDict[$key]=$obj
  }
}
$allHeaders = @("首次抓取時間","最後更新時間","狀態") + $fixedCols + $dynamicCols
$masterOut = foreach($r in $masterRowsList){ $o=[ordered]@{}; foreach($h in $allHeaders){ $o[$h] = ($r.PSObject.Properties.Name -contains $h) ? $r.$h : $null }; New-Object psobject -Property $o }
$masterOut | Export-Csv -Path $MasterCsv -NoTypeInformation -Encoding $csvEncoding
$items | Export-Csv -Path $RunCsv -NoTypeInformation -Encoding $csvEncoding

# Timeseries 追加
$tsHeader = @("TimestampJST","ロット番号","MODEL","現在価格 (JPY)","入札數","渡地 (EX-YARD)","MAKE","YEAR","詳細URL")
if (-not (Test-Path $TimeSeriesCsv)) { Set-Content -Encoding UTF8 $TimeSeriesCsv ($tsHeader -join ",") }
try { $tz=[System.TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time") } catch { $tz=$null }
$nowJst = if ($tz) { [System.TimeZoneInfo]::ConvertTimeFromUtc((Get-Date).ToUniversalTime(), $tz).ToString("yyyy-MM-dd HH:mm:ss") } else { $NowStr }
$tsLines = foreach($row in $items){ if (-not $row."ロット番号") { continue }; ($nowJst, ($row."ロット番号" -replace ","," "), ($row."MODEL" -replace ","," "), $row."現在価格 (JPY)", $row."入札數", ($row."渡地 (EX-YARD)" -replace ","," "), ($row."MAKE" -replace ","," "), $row."YEAR", ($row."詳細URL" -replace ","," ")) -join "," }
Add-Content -Encoding UTF8 $TimeSeriesCsv $tsLines
