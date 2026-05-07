param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "config.json")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    if ($script:LogPath) {
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

function HtmlEncode {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function To-JsJsonArray {
    param([object[]]$Rows)

    if ($null -eq $Rows) {
        return "[]"
    }

    $array = @($Rows)
    $json = ConvertTo-Json -InputObject $array -Depth 8 -Compress

    if ([string]::IsNullOrWhiteSpace($json)) {
        return "[]"
    }

    # Prevent accidental script closing if a subject contains </script>.
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

    if ($daysAhead -lt 0) {
        throw "DaysAhead must be greater than or equal to 0."
    }

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

    # Pick appointments that overlap the output window.
    $filter = "[End] >= '$startText' AND [Start] < '$endText'"
    $restricted = $items.Restrict($filter)

    $busyMap = @{
        0 = "空き"
        1 = "仮予定"
        2 = "予定あり"
        3 = "外出中"
        4 = "他の場所で作業中"
    }

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($appointment in $restricted) {
        try {
            $startDt = [datetime]$appointment.Start
            $endDt = [datetime]$appointment.End

            # Some recurring items may still fall outside the requested range.
            if ($endDt -lt $rangeStart -or $startDt -ge $rangeEnd) {
                continue
            }

            $isPrivate = ($appointment.Sensitivity -eq 2)
            $allDay = [bool]$appointment.AllDayEvent

            $rawSubject = [string]$appointment.Subject
            $rawLocation = [string]$appointment.Location

            if ($maskPrivateItems -and $isPrivate) {
                $subject = "非公開予定"
                $location = ""
            } else {
                if ([string]::IsNullOrWhiteSpace($rawSubject)) {
                    $subject = "(件名なし)"
                } else {
                    $subject = $rawSubject
                }

                $location = if ($includeLocation) { $rawLocation } else { "" }
            }

            $busyStatus = ""
            if ($includeBusyStatus -and $null -ne $appointment.BusyStatus) {
                $busyKey = [int]$appointment.BusyStatus
                if ($busyMap.ContainsKey($busyKey)) {
                    $busyStatus = $busyMap[$busyKey]
                }
            }

            $rows.Add([PSCustomObject]@{
                startISO   = $startDt.ToString("yyyy-MM-ddTHH:mm:ss")
                endISO     = $endDt.ToString("yyyy-MM-ddTHH:mm:ss")
                date       = $startDt.ToString("yyyy-MM-dd")
                startHM    = if ($allDay) { "終日" } else { $startDt.ToString("HH:mm") }
                endHM      = if ($allDay) { "" } else { $endDt.ToString("HH:mm") }
                subject    = $subject
                location   = $location
                isPrivate  = $isPrivate
                allDay     = $allDay
                busyStatus = $busyStatus
            })
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
                @{Name="状態";Expression={$_.busyStatus}},
                @{Name="終日";Expression={$_.allDay}} |
            Export-Csv -Path $tmpCsvPath -NoTypeInformation -Encoding UTF8

        Move-Item -LiteralPath $tmpCsvPath -Destination $csvPath -Force
    }

    $updatedText = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $safeTitle = HtmlEncode $title

    $htmlTemplate = @'
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>__TITLE__</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root {
  --bg: #f5f6f8;
  --card: #ffffff;
  --text: #1f2937;
  --muted: #6b7280;
  --line: #e5e7eb;
  --accent: #2563eb;
  --accent-soft: #dbeafe;
  --private: #f3f4f6;
  --today: #fff7ed;
  --warn: #b45309;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  background: var(--bg);
  color: var(--text);
}
header {
  position: sticky;
  top: 0;
  z-index: 10;
  background: rgba(245, 246, 248, 0.95);
  backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--line);
}
.header-inner {
  max-width: 1080px;
  margin: 0 auto;
  padding: 18px 18px 14px;
}
h1 { margin: 0 0 6px; font-size: 22px; }
.meta { color: var(--muted); font-size: 13px; }
.summary {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 10px;
  margin: 14px 0;
}
.summary-card {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 12px;
}
.summary-label { font-size: 12px; color: var(--muted); }
.summary-value { margin-top: 4px; font-weight: 700; font-size: 18px; }
.controls { display: flex; gap: 8px; flex-wrap: wrap; }
button, input { font: inherit; }
button {
  border: 1px solid var(--line);
  background: var(--card);
  border-radius: 999px;
  padding: 8px 12px;
  cursor: pointer;
}
button.active {
  border-color: var(--accent);
  background: var(--accent-soft);
  color: var(--accent);
  font-weight: 700;
}
input[type="search"] {
  flex: 1;
  min-width: 220px;
  border: 1px solid var(--line);
  border-radius: 999px;
  padding: 8px 13px;
  background: white;
}
main {
  max-width: 1080px;
  margin: 0 auto;
  padding: 18px;
}
.day { margin-bottom: 18px; }
.day-title {
  display: flex;
  align-items: baseline;
  gap: 10px;
  margin: 20px 0 8px;
}
.day-title h2 { margin: 0; font-size: 18px; }
.day-title .count { color: var(--muted); font-size: 13px; }
.day.today {
  background: var(--today);
  border: 1px solid #fed7aa;
  border-radius: 16px;
  padding: 10px 12px 4px;
}
.event {
  display: grid;
  grid-template-columns: 92px 1fr;
  gap: 12px;
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 12px;
  margin: 8px 0;
}
.event.private { background: var(--private); }
.time {
  text-align: center;
  color: var(--accent);
  font-weight: 700;
  white-space: nowrap;
}
.time .end {
  display: block;
  color: var(--muted);
  font-weight: 500;
  font-size: 12px;
  margin-top: 3px;
}
.subject { font-weight: 700; font-size: 15px; line-height: 1.35; }
.details {
  margin-top: 6px;
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  color: var(--muted);
  font-size: 13px;
}
.badge {
  display: inline-flex;
  align-items: center;
  border: 1px solid var(--line);
  border-radius: 999px;
  padding: 2px 8px;
  background: white;
}
.empty {
  background: var(--card);
  border: 1px dashed var(--line);
  border-radius: 14px;
  padding: 28px;
  text-align: center;
  color: var(--muted);
}
@media (max-width: 720px) {
  .summary { grid-template-columns: 1fr; }
  .event { grid-template-columns: 1fr; }
  .time { text-align: left; }
}
@media print {
  header { position: static; background: white; }
  .controls { display: none; }
  body { background: white; }
  .event, .summary-card { break-inside: avoid; }
}
</style>
</head>
<body>
<header>
  <div class="header-inner">
    <h1>__TITLE__</h1>
    <div class="meta">Last updated: __UPDATED__</div>

    <div class="summary">
      <div class="summary-card">
        <div class="summary-label">今日の予定</div>
        <div class="summary-value" id="todayCount">-</div>
      </div>
      <div class="summary-card">
        <div class="summary-label">期間内の予定</div>
        <div class="summary-value" id="totalCount">-</div>
      </div>
      <div class="summary-card">
        <div class="summary-label">次の予定</div>
        <div class="summary-value" id="nextEvent">-</div>
      </div>
    </div>

    <div class="controls">
      <button id="btnToday" class="active">今日</button>
      <button id="btnAll">すべて</button>
      <button id="btnPrint">印刷</button>
      <input id="search" type="search" placeholder="件名・場所を検索">
    </div>
  </div>
</header>

<main id="app"></main>

<script id="schedule-data" type="application/json">
__DATA__
</script>

<script>
const events = JSON.parse(document.getElementById("schedule-data").textContent || "[]");

const app = document.getElementById("app");
const search = document.getElementById("search");
const btnToday = document.getElementById("btnToday");
const btnAll = document.getElementById("btnAll");
const btnPrint = document.getElementById("btnPrint");

let mode = "today";

function pad(n) {
  return String(n).padStart(2, "0");
}

function localDateString(d = new Date()) {
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
}

function formatDateLabel(dateText) {
  const d = new Date(dateText + "T00:00:00");
  return new Intl.DateTimeFormat("ja-JP", {
    month: "numeric",
    day: "numeric",
    weekday: "short"
  }).format(d);
}

function createEl(tag, className, text) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (text !== undefined && text !== null) el.textContent = text;
  return el;
}

function getTodayEvents() {
  const today = localDateString();
  return events.filter(e => e.date === today);
}

function getNextEvent() {
  const now = new Date();
  return events
    .filter(e => new Date(e.endISO) >= now)
    .sort((a, b) => new Date(a.startISO) - new Date(b.startISO))[0];
}

function updateSummary() {
  const todayEvents = getTodayEvents();
  document.getElementById("todayCount").textContent = todayEvents.length + "件";
  document.getElementById("totalCount").textContent = events.length + "件";

  const next = getNextEvent();
  document.getElementById("nextEvent").textContent = next
    ? `${next.startHM} ${next.subject}`
    : "なし";
}

function filteredEvents() {
  const q = search.value.trim().toLowerCase();
  const today = localDateString();

  return events.filter(e => {
    if (mode === "today" && e.date !== today) return false;
    if (!q) return true;

    const text = [
      e.subject || "",
      e.location || "",
      e.busyStatus || ""
    ].join(" ").toLowerCase();

    return text.includes(q);
  });
}

function render() {
  app.innerHTML = "";

  btnToday.classList.toggle("active", mode === "today");
  btnAll.classList.toggle("active", mode === "all");

  const list = filteredEvents();

  if (list.length === 0) {
    app.appendChild(createEl("div", "empty", "表示する予定がありません。"));
    return;
  }

  const grouped = new Map();
  for (const e of list) {
    if (!grouped.has(e.date)) grouped.set(e.date, []);
    grouped.get(e.date).push(e);
  }

  const today = localDateString();

  for (const [date, items] of grouped.entries()) {
    const section = createEl("section", "day" + (date === today ? " today" : ""));

    const title = createEl("div", "day-title");
    title.appendChild(createEl("h2", "", formatDateLabel(date)));
    title.appendChild(createEl("span", "count", `${items.length}件`));
    section.appendChild(title);

    for (const e of items) {
      const card = createEl("article", "event" + (e.isPrivate ? " private" : ""));

      const time = createEl("div", "time");
      time.appendChild(document.createTextNode(e.startHM));
      if (!e.allDay && e.endHM) {
        time.appendChild(createEl("span", "end", "〜 " + e.endHM));
      }

      const body = createEl("div", "body");
      body.appendChild(createEl("div", "subject", e.subject));

      const details = createEl("div", "details");

      if (e.location) {
        details.appendChild(createEl("span", "badge", "📍 " + e.location));
      }

      if (e.busyStatus) {
        details.appendChild(createEl("span", "badge", e.busyStatus));
      }

      if (e.allDay) {
        details.appendChild(createEl("span", "badge", "終日"));
      }

      body.appendChild(details);

      card.appendChild(time);
      card.appendChild(body);
      section.appendChild(card);
    }

    app.appendChild(section);
  }
}

btnToday.addEventListener("click", () => {
  mode = "today";
  render();
});

btnAll.addEventListener("click", () => {
  mode = "all";
  render();
});

btnPrint.addEventListener("click", () => {
  window.print();
});

search.addEventListener("input", render);

updateSummary();
render();
</script>
</body>
</html>
'@

    $html = $htmlTemplate.
        Replace("__TITLE__", $safeTitle).
        Replace("__UPDATED__", (HtmlEncode $updatedText)).
        Replace("__DATA__", $jsonSafe)

    $html | Out-File -FilePath $tmpHtmlPath -Encoding UTF8
    Move-Item -LiteralPath $tmpHtmlPath -Destination $htmlPath -Force

    Write-Log "Exported $($sortedRows.Count) events to $htmlPath"
    if ($writeCsv) {
        Write-Log "CSV exported to $csvPath"
    }

    exit 0
}
catch {
    if (-not $script:LogPath) {
        $script:LogPath = $null
    }

    Write-Log $_.Exception.Message "ERROR"
    exit 1
}
