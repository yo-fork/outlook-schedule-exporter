param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "config.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$script:LogPath = $null

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        Add-Content -Path $script:LogPath -Value $line -Encoding UTF8
    }
    Write-Host $line
}

function Get-ConfigValue {
    param($Config, [string]$Name, $DefaultValue)
    if ($null -ne $Config.PSObject.Properties[$Name] -and $null -ne $Config.$Name) { return $Config.$Name }
    return $DefaultValue
}

function HtmlEncode {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Normalize-Text {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function Get-ComProperty {
    param($Object, [string]$Name, $DefaultValue = $null)
    try {
        $value = $Object.GetType().InvokeMember($Name, [System.Reflection.BindingFlags]::GetProperty, $null, $Object, @())
        if ($null -eq $value) { return $DefaultValue }
        return $value
    } catch {
        try {
            $value = $Object.$Name
            if ($null -eq $value) { return $DefaultValue }
            return $value
        } catch {
            return $DefaultValue
        }
    }
}

function Get-ComText {
    param($Object, [string]$Name, [string]$DefaultValue = "")
    $value = Get-ComProperty -Object $Object -Name $Name -DefaultValue $DefaultValue
    if ($null -eq $value) { return $DefaultValue }
    return [string]$value
}

function Get-ComBool {
    param($Object, [string]$Name, [bool]$DefaultValue = $false)
    $value = Get-ComProperty -Object $Object -Name $Name -DefaultValue $DefaultValue
    try { return [bool]$value } catch { return $DefaultValue }
}

function Get-ComInt {
    param($Object, [string]$Name, [int]$DefaultValue = 0)
    $value = Get-ComProperty -Object $Object -Name $Name -DefaultValue $DefaultValue
    try { return [int]$value } catch { return $DefaultValue }
}

function Get-ComDateTime {
    param($Object, [string]$Name)
    $value = Get-ComProperty -Object $Object -Name $Name -DefaultValue $null
    if ($null -eq $value) { return $null }
    try { return [datetime]$value } catch { return $null }
}

function Get-BodyPreview {
    param([AllowNull()][string]$Body, [int]$MaxLength = 700)
    $text = Normalize-Text $Body
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    $text = $text -replace "[ \t]+", " "
    $text = $text -replace "\n{3,}", "`n`n"
    if ($MaxLength -gt 0 -and $text.Length -gt $MaxLength) { return $text.Substring(0, $MaxLength) + "..." }
    return $text
}

function ConvertTo-SafeJson {
    param($Rows)
    $json = ConvertTo-Json -InputObject @($Rows) -Depth 10 -Compress
    if ([string]::IsNullOrWhiteSpace($json)) { return "[]" }
    return ($json -replace '</', '<\/')
}

function Get-CalendarFolderTypeCandidates {
    $items = New-Object System.Collections.ArrayList
    try {
        Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop | Out-Null
        [void]$items.Add([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    } catch {
        Write-Log "Outlook interop enum is not available. Fallback to int calendar folder type." "WARN"
    }
    [void]$items.Add([int]9)
    return @($items)
}

function Get-DefaultCalendarFolder {
    param($Namespace)
    foreach ($folderType in (Get-CalendarFolderTypeCandidates)) {
        try { return $Namespace.GetDefaultFolder($folderType) }
        catch { Write-Log "GetDefaultFolder failed with type '$folderType': $($_.Exception.Message)" "WARN" }
    }
    throw "Could not open default calendar folder."
}

function Get-SharedCalendarFolder {
    param($Namespace, $Recipient)
    foreach ($folderType in (Get-CalendarFolderTypeCandidates)) {
        try { return $Namespace.GetSharedDefaultFolder($Recipient, $folderType) }
        catch { Write-Log "GetSharedDefaultFolder failed with type '$folderType': $($_.Exception.Message)" "WARN" }
    }
    throw "Could not open shared calendar folder. Check sharing permission."
}

function Get-SharedEntryField {
    param($Entry, [string]$Name)
    if ($Entry -is [string]) {
        if ($Name -eq "Email") { return [string]$Entry }
        if ($Name -eq "Name") { return [string]$Entry }
    }
    if ($null -ne $Entry.PSObject.Properties[$Name]) { return [string]$Entry.$Name }
    return ""
}

function Get-CalendarTargets {
    param($Config, $Namespace)
    $targets = New-Object System.Collections.ArrayList
    $includeDefault = [bool](Get-ConfigValue -Config $Config -Name "IncludeDefaultCalendar" -DefaultValue $true)
    $defaultName = [string](Get-ConfigValue -Config $Config -Name "DefaultCalendarName" -DefaultValue "自分")
    $shared = Get-ConfigValue -Config $Config -Name "SharedCalendars" -DefaultValue @()

    if ($includeDefault) {
        try {
            [void]$targets.Add([PSCustomObject]@{ Name=$defaultName; Email=""; Key="default"; Kind="default"; Folder=(Get-DefaultCalendarFolder -Namespace $Namespace) })
            Write-Log "Default calendar added: $defaultName"
        } catch {
            Write-Log "Failed to open default calendar: $($_.Exception.Message)" "WARN"
        }
    }

    foreach ($entry in @($shared)) {
        $email = Get-SharedEntryField -Entry $entry -Name "Email"
        $name = Get-SharedEntryField -Entry $entry -Name "Name"
        if ([string]::IsNullOrWhiteSpace($email)) {
            Write-Log "Skipped shared calendar entry without Email." "WARN"
            continue
        }

        try {
            $recipient = $Namespace.CreateRecipient($email)
            if (-not [bool]$recipient.Resolve()) {
                Write-Log "Could not resolve shared calendar recipient: $email" "WARN"
                continue
            }
            $folder = Get-SharedCalendarFolder -Namespace $Namespace -Recipient $recipient
            if ([string]::IsNullOrWhiteSpace($name)) { $name = [string](Get-ComProperty -Object $recipient -Name "Name" -DefaultValue $email) }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = $email }
            [void]$targets.Add([PSCustomObject]@{ Name=$name; Email=$email; Key=("shared:" + $email.ToLowerInvariant()); Kind="shared"; Folder=$folder })
            Write-Log "Shared calendar added: $name <$email>"
        } catch {
            Write-Log "Failed to open shared calendar '$email': $($_.Exception.Message)" "WARN"
        }
    }
    return @($targets)
}

function Add-AppointmentRow {
    param($Rows, $Appointment, $Target, $RangeStart, $RangeEnd, $Index, [hashtable]$Settings)

    $startDt = Get-ComDateTime -Object $Appointment -Name "Start"
    $endDt = Get-ComDateTime -Object $Appointment -Name "End"
    if ($null -eq $startDt -or $null -eq $endDt) { return $Index }
    if ($endDt -lt $RangeStart -or $startDt -ge $RangeEnd) { return $Index }

    $busyMap = @{ 0="空き"; 1="仮予定"; 2="予定あり"; 3="外出中"; 4="他の場所で作業中" }
    $importanceMap = @{ 0="低"; 1="標準"; 2="高" }
    $meetingStatusMap = @{ 0="通常予定"; 1="会議"; 3="受信した会議"; 5="キャンセルされた会議" }

    $isPrivate = (Get-ComInt -Object $Appointment -Name "Sensitivity" -DefaultValue 0) -eq 2
    $shouldMask = ([bool]$Settings["MaskPrivateItems"]) -and $isPrivate
    $allDay = Get-ComBool -Object $Appointment -Name "AllDayEvent" -DefaultValue $false

    $rawSubject = Get-ComText -Object $Appointment -Name "Subject"
    $rawLocation = Get-ComText -Object $Appointment -Name "Location"
    $rawOrganizer = Get-ComText -Object $Appointment -Name "Organizer"
    $rawRequired = Get-ComText -Object $Appointment -Name "RequiredAttendees"
    $rawOptional = Get-ComText -Object $Appointment -Name "OptionalAttendees"
    $rawCategories = Get-ComText -Object $Appointment -Name "Categories"
    $rawBody = Get-ComText -Object $Appointment -Name "Body"

    if ($shouldMask) {
        $subject="非公開予定"; $location=""; $organizer=""; $requiredAttendees=""; $optionalAttendees=""; $categories=""; $bodyPreview=""
    } else {
        $subject = if ([string]::IsNullOrWhiteSpace($rawSubject)) { "(件名なし)" } else { $rawSubject }
        $location = if ([bool]$Settings["IncludeLocation"]) { $rawLocation } else { "" }
        $organizer = if ([bool]$Settings["IncludeOrganizer"]) { $rawOrganizer } else { "" }
        $requiredAttendees = if ([bool]$Settings["IncludeAttendees"]) { $rawRequired } else { "" }
        $optionalAttendees = if ([bool]$Settings["IncludeAttendees"]) { $rawOptional } else { "" }
        $categories = if ([bool]$Settings["IncludeCategories"]) { $rawCategories } else { "" }
        $bodyPreview = if ([bool]$Settings["IncludeBodyPreview"]) { Get-BodyPreview -Body $rawBody -MaxLength ([int]$Settings["BodyPreviewMaxLength"]) } else { "" }
    }

    $busyStatus = ""
    if ([bool]$Settings["IncludeBusyStatus"]) {
        $busyKey = Get-ComInt -Object $Appointment -Name "BusyStatus" -DefaultValue -1
        if ($busyMap.ContainsKey($busyKey)) { $busyStatus = $busyMap[$busyKey] }
    }

    $importance = ""
    $importanceKey = Get-ComInt -Object $Appointment -Name "Importance" -DefaultValue -1
    if ($importanceMap.ContainsKey($importanceKey)) { $importance = $importanceMap[$importanceKey] }

    $meetingStatus = ""
    $meetingStatusKey = Get-ComInt -Object $Appointment -Name "MeetingStatus" -DefaultValue -1
    if ($meetingStatusMap.ContainsKey($meetingStatusKey)) { $meetingStatus = $meetingStatusMap[$meetingStatusKey] }

    $reminderText = ""
    if (([bool]$Settings["IncludeReminder"]) -and (Get-ComBool -Object $Appointment -Name "ReminderSet" -DefaultValue $false)) {
        $reminderText = "{0} 分前" -f (Get-ComInt -Object $Appointment -Name "ReminderMinutesBeforeStart" -DefaultValue 0)
    }

    $attendeeCount = 0
    foreach ($text in @($requiredAttendees, $optionalAttendees)) {
        if (-not [string]::IsNullOrWhiteSpace($text)) { $attendeeCount += @($text -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count }
    }

    [void]$Rows.Add([PSCustomObject]@{
        id=$Index; calendarName=[string]$Target.Name; calendarEmail=[string]$Target.Email; calendarKey=[string]$Target.Key; calendarKind=[string]$Target.Kind
        startISO=$startDt.ToString("yyyy-MM-ddTHH:mm:ss"); endISO=$endDt.ToString("yyyy-MM-ddTHH:mm:ss"); date=$startDt.ToString("yyyy-MM-dd")
        startHM=if($allDay){"終日"}else{$startDt.ToString("HH:mm")}; endHM=if($allDay){""}else{$endDt.ToString("HH:mm")}
        subject=Normalize-Text $subject; location=Normalize-Text $location; organizer=Normalize-Text $organizer
        requiredAttendees=Normalize-Text $requiredAttendees; optionalAttendees=Normalize-Text $optionalAttendees; attendeeCount=$attendeeCount
        categories=Normalize-Text $categories; bodyPreview=Normalize-Text $bodyPreview
        isPrivate=$isPrivate; isMasked=$shouldMask; allDay=$allDay; busyStatus=$busyStatus; importance=$importance; meetingStatus=$meetingStatus; reminder=$reminderText
        isRecurring=(Get-ComBool -Object $Appointment -Name "IsRecurring" -DefaultValue $false); durationMinutes=[int][Math]::Round(($endDt - $startDt).TotalMinutes)
    })
    return ($Index + 1)
}

function New-HtmlContent {
    param([string]$Title, [string]$UpdatedText, [string]$DataJson)
    $template = @'
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>__TITLE__</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{--bg:#f3f5f9;--card:#fff;--text:#172033;--muted:#6b7280;--line:#e5e7eb;--accent:#2563eb;--accent-soft:#e8f0ff;--accent-strong:#1d4ed8;--today-bg:#fff7ed;--today-line:#fdba74;--selected-bg:#eef4ff;--private-bg:#f5f5f5;--shadow:0 18px 40px rgba(15,23,42,.08)}*{box-sizing:border-box}html,body{margin:0;padding:0}body{font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:linear-gradient(180deg,#f8fbff 0%,var(--bg) 180px);color:var(--text)}.header-shell{position:sticky;top:0;z-index:10;background:rgba(248,251,255,.88);backdrop-filter:blur(12px);border-bottom:1px solid rgba(229,231,235,.9)}.container{max-width:1240px;margin:0 auto;padding:20px}h1{margin:0;font-size:28px}.meta{margin-top:6px;color:var(--muted);font-size:13px}.summary{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-top:18px}.summary-card,.panel,.agenda-card,.empty{background:var(--card);border:1px solid rgba(229,231,235,.9);border-radius:18px;box-shadow:var(--shadow)}.summary-card{padding:14px 16px}.summary-label{color:var(--muted);font-size:12px}.summary-value{margin-top:6px;font-size:20px;font-weight:800}.summary-sub{margin-top:4px;color:var(--muted);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;align-items:center}.search-wrap{flex:1;min-width:260px;position:relative}.search-wrap input{width:100%;border:1px solid var(--line);border-radius:999px;padding:11px 15px 11px 42px;font:inherit;background:#fff}.search-wrap .icon{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--muted)}button{border:1px solid var(--line);border-radius:999px;padding:10px 14px;background:#fff;cursor:pointer;font:inherit}button.primary{background:var(--accent);color:#fff;border-color:var(--accent)}button.active{background:var(--accent-soft);color:var(--accent-strong);border-color:#bfd4ff;font-weight:700}main.container{display:grid;grid-template-columns:minmax(0,1.2fr) minmax(360px,.9fr);gap:18px;align-items:start}.panel{padding:16px}.panel-title-row{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:12px}.panel-title{font-size:18px;font-weight:800}.panel-sub{color:var(--muted);font-size:12px}.month-nav{display:flex;align-items:center;gap:8px}.month-label{min-width:130px;text-align:center;font-weight:700}.weekdays,.calendar-grid{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:8px}.weekday{padding:8px 4px;color:var(--muted);text-align:center;font-size:12px;font-weight:700}.day-cell{min-height:104px;padding:10px;border:1px solid var(--line);border-radius:16px;background:#fff;cursor:pointer;display:flex;flex-direction:column;gap:8px;transition:transform .12s ease,box-shadow .12s ease,border-color .12s ease}.day-cell:hover{transform:translateY(-1px);box-shadow:0 10px 24px rgba(15,23,42,.06)}.day-cell.other-month{opacity:.45;background:#fafafa}.day-cell.today{background:var(--today-bg);border-color:var(--today-line)}.day-cell.selected{background:var(--selected-bg);border-color:#8eb3ff;box-shadow:0 12px 28px rgba(37,99,235,.14)}.day-top{display:flex;justify-content:space-between;align-items:center}.day-num{font-size:15px;font-weight:800}.day-count{min-width:24px;padding:2px 7px;border-radius:999px;background:#eff6ff;color:var(--accent-strong);font-size:11px;text-align:center;font-weight:700}.day-preview{display:flex;flex-direction:column;gap:4px;margin-top:auto}.day-preview-item{font-size:11px;color:var(--muted);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.day-dot-row{display:flex;gap:5px;flex-wrap:wrap}.day-dot{width:7px;height:7px;border-radius:999px;background:var(--accent)}.day-dot.private{background:#9ca3af}.agenda-header{margin-bottom:12px}.agenda-title{font-size:22px;font-weight:800}.agenda-sub{color:var(--muted);font-size:13px;margin-top:4px}.agenda-list{display:flex;flex-direction:column;gap:12px}.agenda-card{padding:14px}.agenda-card.private{background:var(--private-bg)}.event-top{display:flex;justify-content:space-between;gap:12px;align-items:start}.event-time{color:var(--accent-strong);font-weight:800;font-size:15px;white-space:nowrap}.event-subject{font-size:16px;font-weight:800;line-height:1.4}.event-meta{margin-top:8px;color:var(--muted);font-size:13px;display:flex;flex-wrap:wrap;gap:8px}.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}.chip{display:inline-flex;align-items:center;gap:4px;padding:4px 9px;border-radius:999px;border:1px solid var(--line);background:#fff;color:var(--muted);font-size:12px}.chip.owner{background:#eef4ff;color:var(--accent-strong);border-color:#bfdbfe}.chip.strong{background:#fef3c7;color:#92400e;border-color:#fde68a}.event-body{margin-top:10px;color:#334155;font-size:13px;line-height:1.55;white-space:pre-wrap}.empty{padding:28px;text-align:center;color:var(--muted)}.legend{display:flex;gap:12px;flex-wrap:wrap;margin-top:12px;color:var(--muted);font-size:12px}.legend span{display:inline-flex;align-items:center;gap:6px}.legend i{width:8px;height:8px;border-radius:999px;display:inline-block}.legend .normal{background:var(--accent)}.legend .priv{background:#9ca3af}@media(max-width:980px){.summary{grid-template-columns:repeat(2,minmax(0,1fr))}main.container{grid-template-columns:1fr}}@media(max-width:640px){.container{padding:14px}.summary{grid-template-columns:1fr}.day-cell{min-height:88px;padding:8px}.day-preview-item{display:none}}@media print{.header-shell{position:static;background:#fff}.toolbar,.month-nav button{display:none!important}body{background:#fff}main.container{display:block}.panel,.agenda-card,.summary-card{box-shadow:none}}
</style>
</head>
<body>
<div class="header-shell"><div class="container"><h1>__TITLE__</h1><div class="meta">Last updated: __UPDATED__</div><div class="summary"><div class="summary-card"><div class="summary-label">今日の予定</div><div id="todayCount" class="summary-value">-</div><div class="summary-sub" id="todaySub">-</div></div><div class="summary-card"><div class="summary-label">選択日の予定</div><div id="selectedCount" class="summary-value">-</div><div class="summary-sub" id="selectedSub">-</div></div><div class="summary-card"><div class="summary-label">表示範囲の予定</div><div id="totalCount" class="summary-value">-</div><div class="summary-sub">検索条件を反映</div></div><div class="summary-card"><div class="summary-label">次の予定</div><div id="nextEvent" class="summary-value">-</div><div class="summary-sub" id="nextEventSub">-</div></div></div><div class="toolbar"><div class="search-wrap"><span class="icon">🔎</span><input id="search" type="search" placeholder="件名・場所・主催者・参加者・本文を検索"></div><button id="btnToday" class="primary">今日へ</button><button id="btnClear">検索クリア</button><button id="btnPrint">印刷</button></div></div></div>
<main class="container"><section class="panel"><div class="panel-title-row"><div><div class="panel-title">月カレンダー</div><div class="panel-sub">日付をクリックすると右側に予定一覧を表示</div></div><div class="month-nav"><button id="btnPrevMonth">◀</button><div id="monthLabel" class="month-label">-</div><button id="btnNextMonth">▶</button></div></div><div class="weekdays"><div class="weekday">日</div><div class="weekday">月</div><div class="weekday">火</div><div class="weekday">水</div><div class="weekday">木</div><div class="weekday">金</div><div class="weekday">土</div></div><div id="calendarGrid" class="calendar-grid"></div><div class="legend"><span><i class="normal"></i>通常予定</span><span><i class="priv"></i>非公開予定</span></div></section><section class="panel"><div class="agenda-header"><div id="agendaTitle" class="agenda-title">-</div><div id="agendaSub" class="agenda-sub">-</div></div><div id="agendaList" class="agenda-list"></div></section></main>
<script id="schedule-data" type="application/json">__DATA__</script>
<script>
const events=JSON.parse(document.getElementById('schedule-data').textContent||'[]');const search=document.getElementById('search'),calendarGrid=document.getElementById('calendarGrid'),monthLabel=document.getElementById('monthLabel'),agendaTitle=document.getElementById('agendaTitle'),agendaSub=document.getElementById('agendaSub'),agendaList=document.getElementById('agendaList');function pad(n){return String(n).padStart(2,'0')}function todayString(){const d=new Date();return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}function toDateObj(s){return new Date(s+'T00:00:00')}function fmtDate(s){return new Intl.DateTimeFormat('ja-JP',{year:'numeric',month:'long',day:'numeric',weekday:'short'}).format(toDateObj(s))}function fmtMonth(d){return new Intl.DateTimeFormat('ja-JP',{year:'numeric',month:'long'}).format(d)}function startOfMonth(d){return new Date(d.getFullYear(),d.getMonth(),1)}function addMonths(d,n){return new Date(d.getFullYear(),d.getMonth()+n,1)}function ymd(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}function node(t,c,x){const e=document.createElement(t);if(c)e.className=c;if(x!==undefined&&x!==null)e.textContent=x;return e}const sortedEvents=events.slice().sort((a,b)=>new Date(a.startISO)-new Date(b.startISO));const dates=sortedEvents.map(e=>e.date);const initialDate=dates.indexOf(todayString())>=0?todayString():(dates[0]||todayString());let selectedDate=initialDate,currentMonth=startOfMonth(toDateObj(selectedDate));function searchableText(e){return [e.calendarName||'',e.subject||'',e.location||'',e.organizer||'',e.requiredAttendees||'',e.optionalAttendees||'',e.categories||'',e.bodyPreview||'',e.busyStatus||''].join(' ').toLowerCase()}function filteredEvents(){const q=search.value.trim().toLowerCase();return q?sortedEvents.filter(e=>searchableText(e).includes(q)):sortedEvents.slice()}function eventsByDate(){const m=new Map();filteredEvents().forEach(e=>{if(!m.has(e.date))m.set(e.date,[]);m.get(e.date).push(e)});return m}function monthCells(ms){const first=new Date(ms.getFullYear(),ms.getMonth(),1),start=new Date(first);start.setDate(first.getDate()-first.getDay());const a=[];for(let i=0;i<42;i++){const d=new Date(start);d.setDate(start.getDate()+i);a.push({dateText:ymd(d),inMonth:d.getMonth()===ms.getMonth()})}return a}function updateSummary(){const m=eventsByDate(),te=m.get(todayString())||[],se=m.get(selectedDate)||[],next=filteredEvents().filter(e=>new Date(e.endISO)>=new Date())[0];document.getElementById('todayCount').textContent=te.length+'件';document.getElementById('todaySub').textContent=te.length?te.map(e=>e.subject).slice(0,2).join(' / '):'予定なし';document.getElementById('selectedCount').textContent=se.length+'件';document.getElementById('selectedSub').textContent=fmtDate(selectedDate);document.getElementById('totalCount').textContent=filteredEvents().length+'件';document.getElementById('nextEvent').textContent=next?(next.startHM+' '+next.subject):'なし';document.getElementById('nextEventSub').textContent=next?fmtDate(next.date):'今後の予定なし'}function renderCalendar(){const m=eventsByDate();monthLabel.textContent=fmtMonth(currentMonth);calendarGrid.innerHTML='';monthCells(currentMonth).forEach(c=>{const items=m.get(c.dateText)||[],b=document.createElement('button');b.type='button';b.className='day-cell'+(c.inMonth?'':' other-month')+(c.dateText===todayString()?' today':'')+(c.dateText===selectedDate?' selected':'');const top=node('div','day-top');top.appendChild(node('div','day-num',String(Number(c.dateText.slice(8,10)))));const count=node('div','day-count',items.length?String(items.length):'');if(!items.length)count.style.visibility='hidden';top.appendChild(count);b.appendChild(top);const dots=node('div','day-dot-row');items.slice(0,5).forEach(e=>{const dot=node('span','day-dot'+(e.isPrivate?' private':''));dots.appendChild(dot)});b.appendChild(dots);const preview=node('div','day-preview');items.slice(0,2).forEach(e=>preview.appendChild(node('div','day-preview-item',(e.startHM||'')+' '+(e.subject||''))));b.appendChild(preview);b.onclick=()=>{selectedDate=c.dateText;currentMonth=startOfMonth(toDateObj(selectedDate));renderAll()};calendarGrid.appendChild(b)})}function renderAgenda(){const m=eventsByDate(),items=(m.get(selectedDate)||[]).slice().sort((a,b)=>new Date(a.startISO)-new Date(b.startISO));agendaTitle.textContent=fmtDate(selectedDate);agendaSub.textContent=items.length?(items.length+' 件の予定'):'予定はありません';agendaList.innerHTML='';if(!items.length){agendaList.appendChild(node('div','empty','この日に表示できる予定はありません。'));return}items.forEach(e=>{const card=node('article','agenda-card'+(e.isPrivate?' private':'')),top=node('div','event-top'),left=document.createElement('div');left.appendChild(node('div','event-subject',e.subject||'(件名なし)'));const meta=node('div','event-meta');if(e.location)meta.appendChild(node('span','', '📍 '+e.location));if(e.organizer)meta.appendChild(node('span','', '主催: '+e.organizer));left.appendChild(meta);top.appendChild(left);top.appendChild(node('div','event-time',e.allDay?'終日':((e.startHM||'')+(e.endHM?' - '+e.endHM:''))));card.appendChild(top);const chips=node('div','chips');if(e.calendarName)chips.appendChild(node('span','chip owner','👤 '+e.calendarName));if(e.busyStatus)chips.appendChild(node('span','chip',e.busyStatus));if(e.meetingStatus)chips.appendChild(node('span','chip',e.meetingStatus));if(e.allDay)chips.appendChild(node('span','chip','終日'));if(e.isRecurring)chips.appendChild(node('span','chip','繰り返し'));if(e.importance==='高')chips.appendChild(node('span','chip strong','重要度 高'));if(e.attendeeCount)chips.appendChild(node('span','chip','参加者 '+e.attendeeCount+'名'));card.appendChild(chips);if(e.requiredAttendees)card.appendChild(node('div','event-body','必須参加者: '+e.requiredAttendees));if(e.optionalAttendees)card.appendChild(node('div','event-body','任意参加者: '+e.optionalAttendees));if(e.bodyPreview)card.appendChild(node('div','event-body',e.bodyPreview));agendaList.appendChild(card)})}function renderAll(){updateSummary();renderCalendar();renderAgenda()}document.getElementById('btnToday').onclick=()=>{selectedDate=todayString();currentMonth=startOfMonth(toDateObj(selectedDate));renderAll()};document.getElementById('btnClear').onclick=()=>{search.value='';renderAll()};document.getElementById('btnPrint').onclick=()=>window.print();document.getElementById('btnPrevMonth').onclick=()=>{currentMonth=addMonths(currentMonth,-1);renderCalendar()};document.getElementById('btnNextMonth').onclick=()=>{currentMonth=addMonths(currentMonth,1);renderCalendar()};search.oninput=()=>renderAll();renderAll();
</script>
</body>
</html>
'@
    return $template.Replace('__TITLE__', (HtmlEncode $Title)).Replace('__UPDATED__', (HtmlEncode $UpdatedText)).Replace('__DATA__', $DataJson)
}

try {
    if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "Config file not found: $ConfigPath" }
    $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

    $outDir = [string](Get-ConfigValue -Config $config -Name "OutputDirectory" -DefaultValue ".")
    $htmlFileName = [string](Get-ConfigValue -Config $config -Name "HtmlFileName" -DefaultValue "schedule.html")
    $csvFileName = [string](Get-ConfigValue -Config $config -Name "CsvFileName" -DefaultValue "schedule.csv")
    $logFileName = [string](Get-ConfigValue -Config $config -Name "LogFileName" -DefaultValue "schedule.log")
    $daysAhead = [int](Get-ConfigValue -Config $config -Name "DaysAhead" -DefaultValue 14)
    $title = [string](Get-ConfigValue -Config $config -Name "Title" -DefaultValue "Outlook Schedule")
    $writeCsv = [bool](Get-ConfigValue -Config $config -Name "WriteCsv" -DefaultValue $true)

    $settings = @{
        MaskPrivateItems = [bool](Get-ConfigValue -Config $config -Name "MaskPrivateItems" -DefaultValue $true)
        IncludeLocation = [bool](Get-ConfigValue -Config $config -Name "IncludeLocation" -DefaultValue $true)
        IncludeBusyStatus = [bool](Get-ConfigValue -Config $config -Name "IncludeBusyStatus" -DefaultValue $true)
        IncludeOrganizer = [bool](Get-ConfigValue -Config $config -Name "IncludeOrganizer" -DefaultValue $true)
        IncludeAttendees = [bool](Get-ConfigValue -Config $config -Name "IncludeAttendees" -DefaultValue $true)
        IncludeBodyPreview = [bool](Get-ConfigValue -Config $config -Name "IncludeBodyPreview" -DefaultValue $false)
        BodyPreviewMaxLength = [int](Get-ConfigValue -Config $config -Name "BodyPreviewMaxLength" -DefaultValue 700)
        IncludeCategories = [bool](Get-ConfigValue -Config $config -Name "IncludeCategories" -DefaultValue $true)
        IncludeReminder = [bool](Get-ConfigValue -Config $config -Name "IncludeReminder" -DefaultValue $true)
    }

    if ($daysAhead -lt 0) { throw "DaysAhead must be greater than or equal to 0." }
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null

    $script:LogPath = Join-Path $outDir $logFileName
    $htmlPath = Join-Path $outDir $htmlFileName
    $csvPath = Join-Path $outDir $csvFileName
    $tmpHtmlPath = Join-Path $outDir ($htmlFileName + ".tmp")
    $tmpCsvPath = Join-Path $outDir ($csvFileName + ".tmp")

    Write-Log "Start exporting Outlook calendars."
    $rangeStart = (Get-Date).Date
    $rangeEnd = $rangeStart.AddDays($daysAhead + 1)

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendarTargets = Get-CalendarTargets -Config $config -Namespace $namespace
    if (@($calendarTargets).Count -eq 0) { throw "No calendar target is available. Check IncludeDefaultCalendar and SharedCalendars." }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $startText = $rangeStart.ToString("MM/dd/yyyy hh:mm tt", $culture)
    $endText = $rangeEnd.ToString("MM/dd/yyyy hh:mm tt", $culture)
    $filter = "[End] >= '$startText' AND [Start] < '$endText'"

    $rows = New-Object System.Collections.ArrayList
    $index = 0
    foreach ($target in $calendarTargets) {
        Write-Log "Reading calendar: $($target.Name)"
        try {
            $items = $target.Folder.Items
            $items.Sort("[Start]")
            $items.IncludeRecurrences = $true
            $restricted = $items.Restrict($filter)
        } catch {
            Write-Log "Failed to enumerate calendar '$($target.Name)': $($_.Exception.Message)" "WARN"
            continue
        }
        foreach ($appointment in $restricted) {
            try { $index = Add-AppointmentRow -Rows $rows -Appointment $appointment -Target $target -RangeStart $rangeStart -RangeEnd $rangeEnd -Index $index -Settings $settings }
            catch { Write-Log "Skipped one appointment in '$($target.Name)': $($_.Exception.Message)" "WARN" }
        }
    }

    $sortedRows = @($rows | Sort-Object startISO, calendarName)
    if ($writeCsv) {
        $sortedRows | Select-Object `
            @{Name="予定表";Expression={$_.calendarName}}, @{Name="予定表メール";Expression={$_.calendarEmail}}, @{Name="開始";Expression={$_.startISO}}, @{Name="終了";Expression={$_.endISO}}, @{Name="件名";Expression={$_.subject}}, @{Name="場所";Expression={$_.location}}, @{Name="主催者";Expression={$_.organizer}}, @{Name="必須参加者";Expression={$_.requiredAttendees}}, @{Name="任意参加者";Expression={$_.optionalAttendees}}, @{Name="本文プレビュー";Expression={$_.bodyPreview}}, @{Name="状態";Expression={$_.busyStatus}}, @{Name="会議状態";Expression={$_.meetingStatus}}, @{Name="重要度";Expression={$_.importance}}, @{Name="リマインダー";Expression={$_.reminder}}, @{Name="分類";Expression={$_.categories}}, @{Name="終日";Expression={$_.allDay}}, @{Name="繰り返し";Expression={$_.isRecurring}}, @{Name="非公開";Expression={$_.isPrivate}} |
            Export-Csv -Path $tmpCsvPath -NoTypeInformation -Encoding UTF8
        Move-Item -LiteralPath $tmpCsvPath -Destination $csvPath -Force
    }

    $html = New-HtmlContent -Title $title -UpdatedText (Get-Date -Format "yyyy/MM/dd HH:mm:ss") -DataJson (ConvertTo-SafeJson -Rows $sortedRows)
    $html | Out-File -FilePath $tmpHtmlPath -Encoding UTF8
    Move-Item -LiteralPath $tmpHtmlPath -Destination $htmlPath -Force

    Write-Log "Exported $($sortedRows.Count) events from $(@($calendarTargets).Count) calendars to $htmlPath"
    if ($writeCsv) { Write-Log "CSV exported to $csvPath" }
    exit 0
} catch {
    $line = if ($_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { "unknown" }
    Write-Log ("Line {0} : {1}" -f $line, $_.Exception.Message) "ERROR"
    if ($_.ScriptStackTrace) { Write-Log $_.ScriptStackTrace "ERROR" }
    exit 1
}
