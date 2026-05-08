param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "config.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# StrictMode でも Write-Log が安全に動くよう、最初に初期化する。
$script:LogPath = $null

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
    }

    Write-Host $line
}

function Get-ConfigValue {
    param(
        [object]$Config,
        [string]$Name,
        [object]$DefaultValue
    )

    if ($null -ne $Config.PSObject.Properties[$Name] -and $null -ne $Config.$Name) {
        return $Config.$Name
    }

    return $DefaultValue
}

function Get-ComText {
    param(
        [object]$Object,
        [string]$PropertyName,
        [string]$DefaultValue = ""
    )

    try {
        $value = $Object.$PropertyName
        if ($null -eq $value) { return $DefaultValue }
        return [string]$value
    } catch {
        return $DefaultValue
    }
}

function Get-ComBool {
    param(
        [object]$Object,
        [string]$PropertyName,
        [bool]$DefaultValue = $false
    )

    try {
        $value = $Object.$PropertyName
        if ($null -eq $value) { return $DefaultValue }
        return [bool]$value
    } catch {
        return $DefaultValue
    }
}

function Get-ComInt {
    param(
        [object]$Object,
        [string]$PropertyName,
        [int]$DefaultValue = 0
    )

    try {
        $value = $Object.$PropertyName
        if ($null -eq $value) { return $DefaultValue }
        return [int]$value
    } catch {
        return $DefaultValue
    }
}

function Normalize-Text {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Get-BodyPreview {
    param(
        [AllowNull()][string]$Body,
        [int]$MaxLength = 700
    )

    $text = Normalize-Text $Body
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }

    $text = $text -replace "[ \t]+", " "
    $text = $text -replace "\n{3,}", "`n`n"

    if ($MaxLength -gt 0 -and $text.Length -gt $MaxLength) {
        return $text.Substring(0, $MaxLength) + "..."
    }

    return $text
}

function To-JsJsonArray {
    param([object[]]$Rows)

    if ($null -eq $Rows) { return "[]" }

    $array = @($Rows)
    $json = ConvertTo-Json -InputObject $array -Depth 10 -Compress

    if ([string]::IsNullOrWhiteSpace($json)) { return "[]" }

    # 件名や本文に </script> が含まれても HTML が壊れないようにする。
    return ($json -replace '</', '<\/')
}

try {
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Config file not found: $ConfigPath"
    }

    $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $outDir = [string](Get-ConfigValue -Config $config -Name "OutputDirectory" -DefaultValue ".")
    $htmlFileName = [string](Get-ConfigValue -Config $config -Name "HtmlFileName" -DefaultValue "schedule.html")
    $csvFileName = [string](Get-ConfigValue -Config $config -Name "CsvFileName" -DefaultValue "schedule.csv")
    $logFileName = [string](Get-ConfigValue -Config $config -Name "LogFileName" -DefaultValue "schedule.log")
    $daysAhead = [int](Get-ConfigValue -Config $config -Name "DaysAhead" -DefaultValue 14)
    $title = [string](Get-ConfigValue -Config $config -Name "Title" -DefaultValue "Outlook Schedule")
    $writeCsv = [bool](Get-ConfigValue -Config $config -Name "WriteCsv" -DefaultValue $true)
    $maskPrivateItems = [bool](Get-ConfigValue -Config $config -Name "MaskPrivateItems" -DefaultValue $true)
    $includeLocation = [bool](Get-ConfigValue -Config $config -Name "IncludeLocation" -DefaultValue $true)
    $includeBusyStatus = [bool](Get-ConfigValue -Config $config -Name "IncludeBusyStatus" -DefaultValue $true)
    $includeOrganizer = [bool](Get-ConfigValue -Config $config -Name "IncludeOrganizer" -DefaultValue $true)
    $includeAttendees = [bool](Get-ConfigValue -Config $config -Name "IncludeAttendees" -DefaultValue $true)
    $includeBodyPreview = [bool](Get-ConfigValue -Config $config -Name "IncludeBodyPreview" -DefaultValue $false)
    $bodyPreviewMaxLength = [int](Get-ConfigValue -Config $config -Name "BodyPreviewMaxLength" -DefaultValue 700)
    $includeCategories = [bool](Get-ConfigValue -Config $config -Name "IncludeCategories" -DefaultValue $true)
    $includeReminder = [bool](Get-ConfigValue -Config $config -Name "IncludeReminder" -DefaultValue $true)

    if ($daysAhead -lt 0) { throw "DaysAhead must be greater than or equal to 0." }

    New-Item -ItemType Directory -Force -Path $outDir | Out-Null

    $script:LogPath = Join-Path $outDir $logFileName
    $htmlPath = Join-Path $outDir $htmlFileName
    $csvPath = Join-Path $outDir $csvFileName
    $tmpHtmlPath = Join-Path $outDir ($htmlFileName + ".tmp")
    $tmpCsvPath = Join-Path $outDir ($csvFileName + ".tmp")

    Write-Log "Start exporting Outlook calendar."

    $rangeStart = (Get-Date).Date
    $rangeEnd = $rangeStart.AddDays($daysAhead + 1)

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9) # 9 = olFolderCalendar

    $items = $calendar.Items
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $startText = $rangeStart.ToString("MM/dd/yyyy hh:mm tt", $culture)
    $endText = $rangeEnd.ToString("MM/dd/yyyy hh:mm tt", $culture)

    # 期間に重なる予定を取得する。
    $filter = "[End] >= '$startText' AND [Start] < '$endText'"
    $restricted = $items.Restrict($filter)

    $busyMap = @{
        0 = "空き"
        1 = "仮予定"
        2 = "予定あり"
        3 = "外出中"
        4 = "他の場所で作業中"
    }

    $importanceMap = @{
        0 = "低"
        1 = "標準"
        2 = "高"
    }

    $meetingStatusMap = @{
        0 = "通常予定"
        1 = "会議"
        3 = "受信した会議"
        5 = "キャンセルされた会議"
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $index = 0

    foreach ($appointment in $restricted) {
        try {
            $startDt = [datetime]$appointment.Start
            $endDt = [datetime]$appointment.End

            # 繰り返し予定では Restrict 後も範囲外が混ざることがある。
            if ($endDt -lt $rangeStart -or $startDt -ge $rangeEnd) { continue }

            $isPrivate = (Get-ComInt -Object $appointment -PropertyName "Sensitivity" -DefaultValue 0) -eq 2
            $shouldMask = $maskPrivateItems -and $isPrivate
            $allDay = Get-ComBool -Object $appointment -PropertyName "AllDayEvent" -DefaultValue $false

            $rawSubject = Get-ComText -Object $appointment -PropertyName "Subject" -DefaultValue ""
            $rawLocation = Get-ComText -Object $appointment -PropertyName "Location" -DefaultValue ""
            $rawOrganizer = Get-ComText -Object $appointment -PropertyName "Organizer" -DefaultValue ""
            $rawRequired = Get-ComText -Object $appointment -PropertyName "RequiredAttendees" -DefaultValue ""
            $rawOptional = Get-ComText -Object $appointment -PropertyName "OptionalAttendees" -DefaultValue ""
            $rawCategories = Get-ComText -Object $appointment -PropertyName "Categories" -DefaultValue ""
            $rawBody = Get-ComText -Object $appointment -PropertyName "Body" -DefaultValue ""

            if ($shouldMask) {
                $subject = "非公開予定"
                $location = ""
                $organizer = ""
                $requiredAttendees = ""
                $optionalAttendees = ""
                $categories = ""
                $bodyPreview = ""
            } else {
                $subject = if ([string]::IsNullOrWhiteSpace($rawSubject)) { "(件名なし)" } else { $rawSubject }
                $location = if ($includeLocation) { $rawLocation } else { "" }
                $organizer = if ($includeOrganizer) { $rawOrganizer } else { "" }
                $requiredAttendees = if ($includeAttendees) { $rawRequired } else { "" }
                $optionalAttendees = if ($includeAttendees) { $rawOptional } else { "" }
                $categories = if ($includeCategories) { $rawCategories } else { "" }
                $bodyPreview = if ($includeBodyPreview) { Get-BodyPreview -Body $rawBody -MaxLength $bodyPreviewMaxLength } else { "" }
            }

            $busyStatus = ""
            if ($includeBusyStatus) {
                $busyKey = Get-ComInt -Object $appointment -PropertyName "BusyStatus" -DefaultValue -1
                if ($busyMap.ContainsKey($busyKey)) { $busyStatus = $busyMap[$busyKey] }
            }

            $importance = ""
            $importanceKey = Get-ComInt -Object $appointment -PropertyName "Importance" -DefaultValue -1
            if ($importanceMap.ContainsKey($importanceKey)) { $importance = $importanceMap[$importanceKey] }

            $meetingStatus = ""
            $meetingStatusKey = Get-ComInt -Object $appointment -PropertyName "MeetingStatus" -DefaultValue -1
            if ($meetingStatusMap.ContainsKey($meetingStatusKey)) { $meetingStatus = $meetingStatusMap[$meetingStatusKey] }

            $reminderText = ""
            if ($includeReminder -and (Get-ComBool -Object $appointment -PropertyName "ReminderSet" -DefaultValue $false)) {
                $reminderMinutes = Get-ComInt -Object $appointment -PropertyName "ReminderMinutesBeforeStart" -DefaultValue 0
                $reminderText = "$reminderMinutes 分前"
            }

            $durationMinutes = [int][Math]::Round(($endDt - $startDt).TotalMinutes)
            $isRecurring = Get-ComBool -Object $appointment -PropertyName "IsRecurring" -DefaultValue $false
            $attendeeCount = 0
            foreach ($text in @($requiredAttendees, $optionalAttendees)) {
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    $attendeeCount += @($text -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
                }
            }

            $rows.Add([PSCustomObject]@{
                id                = $index
                startISO          = $startDt.ToString("yyyy-MM-ddTHH:mm:ss")
                endISO            = $endDt.ToString("yyyy-MM-ddTHH:mm:ss")
                date              = $startDt.ToString("yyyy-MM-dd")
                startHM           = if ($allDay) { "終日" } else { $startDt.ToString("HH:mm") }
                endHM             = if ($allDay) { "" } else { $endDt.ToString("HH:mm") }
                subject           = Normalize-Text $subject
                location          = Normalize-Text $location
                organizer         = Normalize-Text $organizer
                requiredAttendees = Normalize-Text $requiredAttendees
                optionalAttendees = Normalize-Text $optionalAttendees
                attendeeCount     = $attendeeCount
                categories        = Normalize-Text $categories
                bodyPreview       = Normalize-Text $bodyPreview
                isPrivate         = $isPrivate
                isMasked          = $shouldMask
                allDay            = $allDay
                busyStatus        = $busyStatus
                importance        = $importance
                meetingStatus     = $meetingStatus
                reminder          = $reminderText
                isRecurring       = $isRecurring
                durationMinutes   = $durationMinutes
            })

            $index += 1
        } catch {
            Write-Log "Skipped one appointment: $($_.Exception.Message)" "WARN"
            continue
        }
    }

    $sortedRows = @($rows | Sort-Object startISO)
    $jsonSafe = To-JsJsonArray -Rows $sortedRows

    if ($writeCsv) {
        $sortedRows |
            Select-Object `
                @{Name="開始";Expression={$_.startISO}},
                @{Name="終了";Expression={$_.endISO}},
                @{Name="件名";Expression={$_.subject}},
                @{Name="場所";Expression={$_.location}},
                @{Name="主催者";Expression={$_.organizer}},
                @{Name="必須参加者";Expression={$_.requiredAttendees}},
                @{Name="任意参加者";Expression={$_.optionalAttendees}},
                @{Name="本文プレビュー";Expression={$_.bodyPreview}},
                @{Name="状態";Expression={$_.busyStatus}},
                @{Name="会議状態";Expression={$_.meetingStatus}},
                @{Name="重要度";Expression={$_.importance}},
                @{Name="リマインダー";Expression={$_.reminder}},
                @{Name="分類";Expression={$_.categories}},
                @{Name="終日";Expression={$_.allDay}},
                @{Name="繰り返し";Expression={$_.isRecurring}},
                @{Name="非公開";Expression={$_.isPrivate}} |
            Export-Csv -Path $tmpCsvPath -NoTypeInformation -Encoding UTF8

        Move-Item -LiteralPath $tmpCsvPath -Destination $csvPath -Force
    }

    $updatedText = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $safeTitle = [System.Net.WebUtility]::HtmlEncode($title)

    $htmlTemplate = @'
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>__TITLE__</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{--bg:#f5f6f8;--card:#fff;--text:#1f2937;--muted:#6b7280;--line:#e5e7eb;--accent:#2563eb;--accent-soft:#dbeafe;--private:#f3f4f6;--today:#fff7ed;--shadow:0 20px 50px rgba(15,23,42,.18)}
*{box-sizing:border-box}body{margin:0;font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:var(--bg);color:var(--text)}
header{position:sticky;top:0;z-index:10;background:rgba(245,246,248,.96);backdrop-filter:blur(8px);border-bottom:1px solid var(--line)}
.header-inner{max-width:1120px;margin:0 auto;padding:18px}h1{margin:0 0 6px;font-size:22px}.meta{color:var(--muted);font-size:13px}
.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:14px 0}.summary-card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px}.summary-label{font-size:12px;color:var(--muted)}.summary-value{margin-top:4px;font-weight:700;font-size:18px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.controls{display:flex;gap:8px;flex-wrap:wrap}button,input{font:inherit}button{border:1px solid var(--line);background:var(--card);border-radius:999px;padding:8px 12px;cursor:pointer}button.active{border-color:var(--accent);background:var(--accent-soft);color:var(--accent);font-weight:700}input[type=search]{flex:1;min-width:220px;border:1px solid var(--line);border-radius:999px;padding:8px 13px;background:white}
main{max-width:1120px;margin:0 auto;padding:18px}.day{margin-bottom:18px}.day-title{display:flex;align-items:baseline;gap:10px;margin:20px 0 8px}.day-title h2{margin:0;font-size:18px}.count{color:var(--muted);font-size:13px}.day.today{background:var(--today);border:1px solid #fed7aa;border-radius:16px;padding:10px 12px 4px}
.event{display:grid;grid-template-columns:92px 1fr auto;gap:12px;align-items:start;background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px;margin:8px 0;cursor:pointer;transition:.12s}.event:hover{border-color:#bfdbfe;transform:translateY(-1px)}.event.private{background:var(--private)}.time{text-align:center;color:var(--accent);font-weight:700;white-space:nowrap}.end{display:block;color:var(--muted);font-weight:500;font-size:12px;margin-top:3px}.subject{font-weight:700;font-size:15px;line-height:1.35}.subline{margin-top:5px;color:var(--muted);font-size:13px}.details{margin-top:7px;display:flex;flex-wrap:wrap;gap:7px;color:var(--muted);font-size:12px}.badge{display:inline-flex;align-items:center;border:1px solid var(--line);border-radius:999px;padding:2px 8px;background:white}.open{color:var(--accent);font-weight:700;font-size:13px;margin-top:3px;white-space:nowrap}.empty{background:var(--card);border:1px dashed var(--line);border-radius:14px;padding:28px;text-align:center;color:var(--muted)}
.backdrop{position:fixed;inset:0;background:rgba(15,23,42,.32);z-index:20;display:none}.backdrop.opened{display:block}.panel{position:fixed;right:18px;top:18px;bottom:18px;width:min(520px,calc(100vw - 36px));background:white;border:1px solid var(--line);border-radius:20px;box-shadow:var(--shadow);z-index:21;display:none;overflow:hidden}.panel.opened{display:flex;flex-direction:column}.panel-head{padding:18px;border-bottom:1px solid var(--line);display:flex;gap:12px;justify-content:space-between;align-items:flex-start}.panel-title{font-size:18px;font-weight:800;line-height:1.35}.panel-body{padding:18px;overflow:auto}.kv{margin:0 0 14px}.kv dt{font-size:12px;color:var(--muted);margin-bottom:3px}.kv dd{margin:0;white-space:pre-wrap;line-height:1.5}.close{border-radius:10px;padding:6px 10px}.hint{font-size:12px;color:var(--muted);margin-top:8px}
@media(max-width:720px){.summary{grid-template-columns:1fr}.event{grid-template-columns:1fr}.time{text-align:left}.open{display:none}.panel{left:8px;right:8px;top:auto;bottom:8px;width:auto;max-height:88vh}}
@media print{header{position:static;background:white}.controls,.backdrop,.panel,.open{display:none!important}body{background:white}.event,.summary-card{break-inside:avoid}.event{grid-template-columns:92px 1fr}}
</style>
</head>
<body>
<header><div class="header-inner"><h1>__TITLE__</h1><div class="meta">Last updated: __UPDATED__</div><div class="summary"><div class="summary-card"><div class="summary-label">今日の予定</div><div class="summary-value" id="todayCount">-</div></div><div class="summary-card"><div class="summary-label">期間内の予定</div><div class="summary-value" id="totalCount">-</div></div><div class="summary-card"><div class="summary-label">次の予定</div><div class="summary-value" id="nextEvent">-</div></div></div><div class="controls"><button id="btnToday" class="active">今日</button><button id="btnAll">すべて</button><button id="btnPrint">印刷</button><input id="search" type="search" placeholder="件名・場所・参加者・本文を検索"></div></div></header>
<main id="app"></main>
<div id="backdrop" class="backdrop"></div><aside id="panel" class="panel" aria-label="予定詳細"><div class="panel-head"><div><div id="panelTitle" class="panel-title"></div><div id="panelHint" class="hint"></div></div><button id="btnClose" class="close">閉じる</button></div><div id="panelBody" class="panel-body"></div></aside>
<script id="schedule-data" type="application/json">__DATA__</script>
<script>
const events=JSON.parse(document.getElementById('schedule-data').textContent||'[]');
const app=document.getElementById('app'),search=document.getElementById('search'),btnToday=document.getElementById('btnToday'),btnAll=document.getElementById('btnAll'),btnPrint=document.getElementById('btnPrint'),panel=document.getElementById('panel'),backdrop=document.getElementById('backdrop'),btnClose=document.getElementById('btnClose');
let mode='today';
const pad=n=>String(n).padStart(2,'0');
const today=()=>{const d=new Date();return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`};
const fmtDate=s=>new Intl.DateTimeFormat('ja-JP',{month:'numeric',day:'numeric',weekday:'short'}).format(new Date(s+'T00:00:00'));
const el=(tag,cls,text)=>{const x=document.createElement(tag);if(cls)x.className=cls;if(text!==undefined&&text!==null)x.textContent=text;return x};
const oneLine=s=>(s||'').replace(/\s+/g,' ').trim();
function nextEvent(){const now=new Date();return events.filter(e=>new Date(e.endISO)>=now).sort((a,b)=>new Date(a.startISO)-new Date(b.startISO))[0]}
function updateSummary(){document.getElementById('todayCount').textContent=events.filter(e=>e.date===today()).length+'件';document.getElementById('totalCount').textContent=events.length+'件';const n=nextEvent();document.getElementById('nextEvent').textContent=n?`${n.startHM} ${n.subject}`:'なし'}
function filtered(){const q=search.value.trim().toLowerCase(),t=today();return events.filter(e=>{if(mode==='today'&&e.date!==t)return false;if(!q)return true;return [e.subject,e.location,e.organizer,e.requiredAttendees,e.optionalAttendees,e.categories,e.bodyPreview,e.busyStatus].join(' ').toLowerCase().includes(q)})}
function render(){app.innerHTML='';btnToday.classList.toggle('active',mode==='today');btnAll.classList.toggle('active',mode==='all');const list=filtered();if(!list.length){app.appendChild(el('div','empty','表示する予定がありません。'));return}const grouped=new Map();for(const e of list){if(!grouped.has(e.date))grouped.set(e.date,[]);grouped.get(e.date).push(e)}const t=today();for(const [date,items]of grouped){const sec=el('section','day'+(date===t?' today':''));const title=el('div','day-title');title.appendChild(el('h2','',fmtDate(date)));title.appendChild(el('span','count',`${items.length}件`));sec.appendChild(title);for(const e of items)sec.appendChild(card(e));app.appendChild(sec)}}
function card(e){const c=el('article','event'+(e.isPrivate?' private':''));c.tabIndex=0;c.title='クリックして詳細を表示';c.onclick=()=>openDetail(e);c.onkeydown=ev=>{if(ev.key==='Enter'||ev.key===' '){ev.preventDefault();openDetail(e)}};const tm=el('div','time');tm.appendChild(document.createTextNode(e.startHM));if(!e.allDay&&e.endHM)tm.appendChild(el('span','end','〜 '+e.endHM));const body=el('div','body');body.appendChild(el('div','subject',e.subject));const sub=[e.location?`📍 ${e.location}`:'',e.organizer?`主催: ${e.organizer}`:''].filter(Boolean).join(' / ');if(sub)body.appendChild(el('div','subline',sub));const det=el('div','details');for(const b of badges(e))det.appendChild(el('span','badge',b));body.appendChild(det);c.appendChild(tm);c.appendChild(body);c.appendChild(el('div','open','詳細'));return c}
function badges(e){const a=[];if(e.busyStatus)a.push(e.busyStatus);if(e.meetingStatus)a.push(e.meetingStatus);if(e.attendeeCount)a.push(`参加者 ${e.attendeeCount}名`);if(e.allDay)a.push('終日');if(e.isRecurring)a.push('繰り返し');if(e.importance==='高')a.push('重要度 高');return a}
function openDetail(e){document.getElementById('panelTitle').textContent=e.subject;document.getElementById('panelHint').textContent=e.isMasked?'非公開予定のため詳細はマスクされています。':'カード外をクリック、または Esc で閉じます。';const rows=[['日時',`${fmtDate(e.date)} ${e.startHM}${e.endHM?' 〜 '+e.endHM:''}`],['所要時間',e.allDay?'終日':`${e.durationMinutes}分`],['場所',e.location],['主催者',e.organizer],['必須参加者',e.requiredAttendees],['任意参加者',e.optionalAttendees],['状態',e.busyStatus],['会議状態',e.meetingStatus],['重要度',e.importance],['リマインダー',e.reminder],['分類',e.categories],['本文プレビュー',e.bodyPreview]].filter(x=>x[1]);const dl=document.createElement('dl');dl.className='kv';for(const [k,v]of rows){dl.appendChild(el('dt','',k));dl.appendChild(el('dd','',v))}const box=document.getElementById('panelBody');box.innerHTML='';box.appendChild(dl);panel.classList.add('opened');backdrop.classList.add('opened')}
function closeDetail(){panel.classList.remove('opened');backdrop.classList.remove('opened')}
btnToday.onclick=()=>{mode='today';render()};btnAll.onclick=()=>{mode='all';render()};btnPrint.onclick=()=>window.print();search.oninput=render;btnClose.onclick=closeDetail;backdrop.onclick=closeDetail;document.addEventListener('keydown',e=>{if(e.key==='Escape')closeDetail()});
updateSummary();render();
</script>
</body>
</html>
'@

    $html = $htmlTemplate.
        Replace("__TITLE__", $safeTitle).
        Replace("__UPDATED__", [System.Net.WebUtility]::HtmlEncode($updatedText)).
        Replace("__DATA__", $jsonSafe)

    $html | Out-File -FilePath $tmpHtmlPath -Encoding UTF8
    Move-Item -LiteralPath $tmpHtmlPath -Destination $htmlPath -Force

    Write-Log "Exported $($sortedRows.Count) events to $htmlPath"
    if ($writeCsv) { Write-Log "CSV exported to $csvPath" }
    exit 0
}
catch {
    Write-Log $_.Exception.Message "ERROR"
    exit 1
}
