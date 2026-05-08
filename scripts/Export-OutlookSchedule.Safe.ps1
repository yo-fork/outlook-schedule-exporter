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

function Get-SettingBool {
    param([hashtable]$Settings, [string]$Name, [bool]$DefaultValue = $false)
    if ($Settings.ContainsKey($Name)) { return [bool]$Settings[$Name] }
    return $DefaultValue
}

function Get-SettingInt {
    param([hashtable]$Settings, [string]$Name, [int]$DefaultValue = 0)
    if ($Settings.ContainsKey($Name)) { return [int]$Settings[$Name] }
    return $DefaultValue
}

function Normalize-Text {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return (($Text -replace "`r`n", "`n") -replace "`r", "`n").Trim()
}

function HtmlEncode {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
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
    if ($MaxLength -gt 0 -and $text.Length -gt $MaxLength) {
        return $text.Substring(0, $MaxLength) + "..."
    }
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
            [void]$targets.Add([PSCustomObject]@{
                Name = $defaultName
                Email = ""
                Key = "default"
                Kind = "default"
                Folder = (Get-DefaultCalendarFolder -Namespace $Namespace)
            })
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

            [void]$targets.Add([PSCustomObject]@{
                Name = $name
                Email = $email
                Key = "shared:" + $email.ToLowerInvariant()
                Kind = "shared"
                Folder = $folder
            })
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

    $busyMap = @{ 0 = "空き"; 1 = "仮予定"; 2 = "予定あり"; 3 = "外出中"; 4 = "他の場所で作業中" }
    $importanceMap = @{ 0 = "低"; 1 = "標準"; 2 = "高" }
    $meetingStatusMap = @{ 0 = "通常予定"; 1 = "会議"; 3 = "受信した会議"; 5 = "キャンセルされた会議" }

    $isPrivate = (Get-ComInt -Object $Appointment -Name "Sensitivity" -DefaultValue 0) -eq 2
    $shouldMask = (Get-SettingBool -Settings $Settings -Name "MaskPrivateItems" -DefaultValue $true) -and $isPrivate
    $allDay = Get-ComBool -Object $Appointment -Name "AllDayEvent" -DefaultValue $false

    $rawSubject = Get-ComText -Object $Appointment -Name "Subject"
    $rawLocation = Get-ComText -Object $Appointment -Name "Location"
    $rawOrganizer = Get-ComText -Object $Appointment -Name "Organizer"
    $rawRequired = Get-ComText -Object $Appointment -Name "RequiredAttendees"
    $rawOptional = Get-ComText -Object $Appointment -Name "OptionalAttendees"
    $rawCategories = Get-ComText -Object $Appointment -Name "Categories"
    $rawBody = Get-ComText -Object $Appointment -Name "Body"

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
        $location = if (Get-SettingBool -Settings $Settings -Name "IncludeLocation" -DefaultValue $true) { $rawLocation } else { "" }
        $organizer = if (Get-SettingBool -Settings $Settings -Name "IncludeOrganizer" -DefaultValue $true) { $rawOrganizer } else { "" }
        $requiredAttendees = if (Get-SettingBool -Settings $Settings -Name "IncludeAttendees" -DefaultValue $true) { $rawRequired } else { "" }
        $optionalAttendees = if (Get-SettingBool -Settings $Settings -Name "IncludeAttendees" -DefaultValue $true) { $rawOptional } else { "" }
        $categories = if (Get-SettingBool -Settings $Settings -Name "IncludeCategories" -DefaultValue $true) { $rawCategories } else { "" }
        $bodyPreview = if (Get-SettingBool -Settings $Settings -Name "IncludeBodyPreview" -DefaultValue $false) {
            Get-BodyPreview -Body $rawBody -MaxLength (Get-SettingInt -Settings $Settings -Name "BodyPreviewMaxLength" -DefaultValue 700)
        } else { "" }
    }

    $busyStatus = ""
    if (Get-SettingBool -Settings $Settings -Name "IncludeBusyStatus" -DefaultValue $true) {
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
    if ((Get-SettingBool -Settings $Settings -Name "IncludeReminder" -DefaultValue $true) -and (Get-ComBool -Object $Appointment -Name "ReminderSet" -DefaultValue $false)) {
        $reminderText = "{0} 分前" -f (Get-ComInt -Object $Appointment -Name "ReminderMinutesBeforeStart" -DefaultValue 0)
    }

    $attendeeCount = 0
    foreach ($text in @($requiredAttendees, $optionalAttendees)) {
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $attendeeCount += @($text -split ";" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
        }
    }

    [void]$Rows.Add([PSCustomObject]@{
        id = $Index
        calendarName = [string]$Target.Name
        calendarEmail = [string]$Target.Email
        calendarKey = [string]$Target.Key
        calendarKind = [string]$Target.Kind
        startISO = $startDt.ToString("yyyy-MM-ddTHH:mm:ss")
        endISO = $endDt.ToString("yyyy-MM-ddTHH:mm:ss")
        date = $startDt.ToString("yyyy-MM-dd")
        startHM = if ($allDay) { "終日" } else { $startDt.ToString("HH:mm") }
        endHM = if ($allDay) { "" } else { $endDt.ToString("HH:mm") }
        subject = Normalize-Text $subject
        location = Normalize-Text $location
        organizer = Normalize-Text $organizer
        requiredAttendees = Normalize-Text $requiredAttendees
        optionalAttendees = Normalize-Text $optionalAttendees
        attendeeCount = $attendeeCount
        categories = Normalize-Text $categories
        bodyPreview = Normalize-Text $bodyPreview
        isPrivate = $isPrivate
        isMasked = $shouldMask
        allDay = $allDay
        busyStatus = $busyStatus
        importance = $importance
        meetingStatus = $meetingStatus
        reminder = $reminderText
        isRecurring = (Get-ComBool -Object $Appointment -Name "IsRecurring" -DefaultValue $false)
        durationMinutes = [int][Math]::Round(($endDt - $startDt).TotalMinutes)
    })

    return ($Index + 1)
}

function New-HtmlContent {
    param([string]$Title, [string]$UpdatedText, [string]$DataJson)
    $safeTitle = HtmlEncode $Title
    $safeUpdated = HtmlEncode $UpdatedText
    return @"
<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>$safeTitle</title><meta name="viewport" content="width=device-width,initial-scale=1"><style>
body{margin:0;font-family:system-ui,-apple-system,"Segoe UI",sans-serif;background:#f5f6f8;color:#1f2937}header{position:sticky;top:0;background:#f5f6f8eF;border-bottom:1px solid #e5e7eb;z-index:10}.wrap{max-width:1120px;margin:auto;padding:18px}h1{margin:0 0 6px;font-size:22px}.meta{font-size:13px;color:#6b7280}.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:14px 0}.box,.event{background:white;border:1px solid #e5e7eb;border-radius:14px}.box{padding:12px}.label{font-size:12px;color:#6b7280}.value{margin-top:4px;font-size:18px;font-weight:700}.controls,.filters{display:flex;gap:8px;flex-wrap:wrap;margin-top:8px}button,input{font:inherit}button{border:1px solid #e5e7eb;background:white;border-radius:999px;padding:8px 12px;cursor:pointer}button.active{border-color:#2563eb;background:#dbeafe;color:#2563eb;font-weight:700}input{flex:1;min-width:230px;border:1px solid #e5e7eb;border-radius:999px;padding:8px 13px}main{max-width:1120px;margin:auto;padding:18px}.day{margin-bottom:18px}.day.today{background:#fff7ed;border:1px solid #fed7aa;border-radius:16px;padding:10px 12px 4px}.day-title{display:flex;gap:10px;align-items:baseline;margin:20px 0 8px}.day-title h2{margin:0;font-size:18px}.count{font-size:13px;color:#6b7280}.event{display:grid;grid-template-columns:92px 1fr auto;gap:12px;align-items:start;padding:12px;margin:8px 0;cursor:pointer}.event.private{background:#f3f4f6}.time{text-align:center;color:#2563eb;font-weight:700}.end{display:block;color:#6b7280;font-size:12px;font-weight:500;margin-top:3px}.subject{font-weight:700;line-height:1.35}.sub{margin-top:5px;color:#6b7280;font-size:13px}.badges{display:flex;flex-wrap:wrap;gap:7px;margin-top:7px;font-size:12px}.badge{border:1px solid #e5e7eb;border-radius:999px;padding:2px 8px;background:white;color:#6b7280}.owner{border-color:#bfdbfe;background:#eff6ff;color:#1d4ed8}.open{color:#2563eb;font-weight:700;font-size:13px}.empty{background:white;border:1px dashed #e5e7eb;border-radius:14px;padding:28px;text-align:center;color:#6b7280}.backdrop{position:fixed;inset:0;background:rgba(15,23,42,.32);z-index:20;display:none}.backdrop.on{display:block}.panel{position:fixed;right:18px;top:18px;bottom:18px;width:min(560px,calc(100vw - 36px));background:white;border:1px solid #e5e7eb;border-radius:20px;box-shadow:0 20px 50px rgba(15,23,42,.18);z-index:21;display:none;overflow:hidden}.panel.on{display:flex;flex-direction:column}.panel-head{padding:18px;border-bottom:1px solid #e5e7eb;display:flex;justify-content:space-between;gap:12px}.panel-title{font-size:18px;font-weight:800;line-height:1.35}.hint{font-size:12px;color:#6b7280;margin-top:8px}.panel-body{padding:18px;overflow:auto}.kv dt{font-size:12px;color:#6b7280;margin-bottom:3px}.kv dd{margin:0 0 14px;white-space:pre-wrap;line-height:1.5}.close{border-radius:10px}@media(max-width:720px){.summary{grid-template-columns:1fr}.event{grid-template-columns:1fr}.time{text-align:left}.open{display:none}.panel{left:8px;right:8px;top:auto;bottom:8px;width:auto;max-height:88vh}}@media print{header{position:static}.controls,.filters,.backdrop,.panel,.open{display:none!important}body{background:white}.event,.box{break-inside:avoid}}</style></head><body><header><div class="wrap"><h1>$safeTitle</h1><div class="meta">Last updated: $safeUpdated</div><div class="summary"><div class="box"><div class="label">今日の予定</div><div class="value" id="todayCount">-</div></div><div class="box"><div class="label">表示中の予定</div><div class="value" id="totalCount">-</div></div><div class="box"><div class="label">次の予定</div><div class="value" id="nextEvent">-</div></div></div><div class="controls"><button id="btnToday" class="active">今日</button><button id="btnAll">すべて</button><button id="btnPrint">印刷</button><input id="search" type="search" placeholder="予定表・件名・場所・参加者・本文を検索"></div><div id="filters" class="filters"></div></div></header><main id="app"></main><div id="backdrop" class="backdrop"></div><aside id="panel" class="panel"><div class="panel-head"><div><div id="panelTitle" class="panel-title"></div><div id="panelHint" class="hint"></div></div><button id="btnClose" class="close">閉じる</button></div><div id="panelBody" class="panel-body"></div></aside><script id="schedule-data" type="application/json">$DataJson</script><script>
const events=JSON.parse(document.getElementById('schedule-data').textContent||'[]');const app=document.getElementById('app'),search=document.getElementById('search'),filters=document.getElementById('filters'),panel=document.getElementById('panel'),backdrop=document.getElementById('backdrop');let mode='today',cal='all';const pad=n=>String(n).padStart(2,'0'),today=()=>{const d=new Date();return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`},fmt=s=>new Intl.DateTimeFormat('ja-JP',{month:'numeric',day:'numeric',weekday:'short'}).format(new Date(s+'T00:00:00')),el=(t,c,x)=>{const e=document.createElement(t);if(c)e.className=c;if(x!=null)e.textContent=x;return e};function calendarList(){const m=new Map();for(const e of events){const k=e.calendarKey||'default';if(!m.has(k))m.set(k,{key:k,name:e.calendarName||'予定表',count:0});m.get(k).count++}return [...m.values()].sort((a,b)=>a.name.localeCompare(b.name,'ja'))}function drawFilters(){filters.innerHTML='';const all=el('button','',`全員 (${events.length})`);all.onclick=()=>{cal='all';drawFilters();update();render()};filters.appendChild(all);for(const c of calendarList()){const b=el('button','',`${c.name} (${c.count})`);b.onclick=()=>{cal=c.key;drawFilters();update();render()};b.classList.toggle('active',cal===c.key);filters.appendChild(b)}all.classList.toggle('active',cal==='all')}function pool(){return events.filter(e=>cal==='all'||e.calendarKey===cal)}function filtered(){const q=search.value.trim().toLowerCase(),t=today();return pool().filter(e=>{if(mode==='today'&&e.date!==t)return false;if(!q)return true;return [e.calendarName,e.subject,e.location,e.organizer,e.requiredAttendees,e.optionalAttendees,e.categories,e.bodyPreview,e.busyStatus].join(' ').toLowerCase().includes(q)})}function update(){const p=pool(),f=filtered(),n=p.filter(e=>new Date(e.endISO)>=new Date()).sort((a,b)=>new Date(a.startISO)-new Date(b.startISO))[0];document.getElementById('todayCount').textContent=p.filter(e=>e.date===today()).length+'件';document.getElementById('totalCount').textContent=f.length+'件';document.getElementById('nextEvent').textContent=n?`${n.startHM} ${n.calendarName}: ${n.subject}`:'なし'}function render(){app.innerHTML='';document.getElementById('btnToday').classList.toggle('active',mode==='today');document.getElementById('btnAll').classList.toggle('active',mode==='all');const list=filtered();if(!list.length){app.appendChild(el('div','empty','表示する予定がありません。'));return}const g=new Map();for(const e of list){if(!g.has(e.date))g.set(e.date,[]);g.get(e.date).push(e)}for(const [d,items] of g){const sec=el('section','day'+(d===today()?' today':'')),h=el('div','day-title');h.appendChild(el('h2','',fmt(d)));h.appendChild(el('span','count',items.length+'件'));sec.appendChild(h);for(const e of items)sec.appendChild(card(e));app.appendChild(sec)}}function card(e){const c=el('article','event'+(e.isPrivate?' private':''));c.tabIndex=0;c.onclick=()=>detail(e);c.onkeydown=x=>{if(x.key==='Enter'||x.key===' '){x.preventDefault();detail(e)}};const tm=el('div','time',e.startHM);if(!e.allDay&&e.endHM)tm.appendChild(el('span','end','〜 '+e.endHM));const b=el('div','');b.appendChild(el('div','subject',e.subject));const sub=[e.location?'📍 '+e.location:'',e.organizer?'主催: '+e.organizer:''].filter(Boolean).join(' / ');if(sub)b.appendChild(el('div','sub',sub));const bs=el('div','badges');bs.appendChild(el('span','badge owner','👤 '+(e.calendarName||'予定表')));for(const x of [e.busyStatus,e.meetingStatus,e.attendeeCount?`参加者 ${e.attendeeCount}名`:'',e.allDay?'終日':'',e.isRecurring?'繰り返し':'',e.importance==='高'?'重要度 高':''].filter(Boolean))bs.appendChild(el('span','badge',x));b.appendChild(bs);c.appendChild(tm);c.appendChild(b);c.appendChild(el('div','open','詳細'));return c}function detail(e){document.getElementById('panelTitle').textContent=e.subject;document.getElementById('panelHint').textContent=e.isMasked?'非公開予定のため詳細はマスクされています。':'カード外または Esc で閉じます。';const rows=[['予定表',e.calendarName],['予定表メール',e.calendarEmail],['日時',`${fmt(e.date)} ${e.startHM}${e.endHM?' 〜 '+e.endHM:''}`],['所要時間',e.allDay?'終日':`${e.durationMinutes}分`],['場所',e.location],['主催者',e.organizer],['必須参加者',e.requiredAttendees],['任意参加者',e.optionalAttendees],['状態',e.busyStatus],['会議状態',e.meetingStatus],['重要度',e.importance],['リマインダー',e.reminder],['分類',e.categories],['本文プレビュー',e.bodyPreview]].filter(r=>r[1]);const dl=el('dl','kv');for(const [k,v]of rows){dl.appendChild(el('dt','',k));dl.appendChild(el('dd','',v))}document.getElementById('panelBody').innerHTML='';document.getElementById('panelBody').appendChild(dl);panel.classList.add('on');backdrop.classList.add('on')}function close(){panel.classList.remove('on');backdrop.classList.remove('on')}document.getElementById('btnToday').onclick=()=>{mode='today';update();render()};document.getElementById('btnAll').onclick=()=>{mode='all';update();render()};document.getElementById('btnPrint').onclick=()=>window.print();document.getElementById('btnClose').onclick=close;backdrop.onclick=close;search.oninput=()=>{update();render()};document.addEventListener('keydown',e=>{if(e.key==='Escape')close()});drawFilters();update();render();
</script></body></html>
"@
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
            try {
                $index = Add-AppointmentRow -Rows $rows -Appointment $appointment -Target $target -RangeStart $rangeStart -RangeEnd $rangeEnd -Index $index -Settings $settings
            } catch {
                Write-Log "Skipped one appointment in '$($target.Name)': $($_.Exception.Message)" "WARN"
            }
        }
    }

    $sortedRows = @($rows | Sort-Object startISO, calendarName)
    if ($writeCsv) {
        $sortedRows | Select-Object `
            @{Name="予定表";Expression={$_.calendarName}},
            @{Name="予定表メール";Expression={$_.calendarEmail}},
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

    $html = New-HtmlContent -Title $title -UpdatedText (Get-Date -Format "yyyy/MM/dd HH:mm:ss") -DataJson (ConvertTo-SafeJson -Rows $sortedRows)
    $html | Out-File -FilePath $tmpHtmlPath -Encoding UTF8
    Move-Item -LiteralPath $tmpHtmlPath -Destination $htmlPath -Force

    Write-Log "Exported $($sortedRows.Count) events from $(@($calendarTargets).Count) calendars to $htmlPath"
    if ($writeCsv) { Write-Log "CSV exported to $csvPath" }
    exit 0
} catch {
    $line = if ($_.InvocationInfo) { $_.InvocationInfo.ScriptLineNumber } else { "unknown" }
    Write-Log "Line $line: $($_.Exception.Message)" "ERROR"
    if ($_.ScriptStackTrace) { Write-Log $_.ScriptStackTrace "ERROR" }
    exit 1
}
