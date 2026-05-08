# Outlook Schedule Exporter

Outlook の予定表を読み取り、単体 HTML と CSV を出力するための PowerShell スクリプトです。

## 目的

- Outlook が入っていない、または Outlook Web にアクセスできない端末から予定を確認する
- 共有フォルダ等に `schedule.html` を置き、ブラウザだけで見られるようにする
- 単体HTMLなので、外部CDNやインターネット接続なしで使える
- 今日の予定、全期間表示、検索、印刷、予定詳細ドロワーに対応する
- Outlookで閲覧権限がある共有予定表もまとめて表示する

## UI

`schedule.html` では、予定カードをクリックすると詳細ドロワーが開きます。

詳細ドロワーには、設定に応じて以下を表示できます。

- 予定表名
- 日時
- 所要時間
- 場所
- 主催者
- 必須参加者
- 任意参加者
- 空き時間状態
- 会議状態
- 重要度
- リマインダー
- 分類
- 本文プレビュー

検索欄では、予定表名、件名、場所、主催者、参加者、分類、本文プレビューをまとめて検索できます。

複数予定表を取得している場合は、上部に `全員` / 各予定表名のフィルタボタンが出ます。

## 出力されるもの

```text
schedule.html  # 普段見る用。検索、今日表示、全期間表示、詳細ドロワー、印刷に対応
schedule.csv   # Excel確認、簡易検索、バックアップ用
schedule.log   # 実行ログ
```

## 前提

- classic Outlook が入っている Windows 端末
- Outlook に予定表が同期されていること
- PowerShell が利用できること
- 出力先フォルダへの書き込み権限があること
- 共有予定表を取得する場合、Outlook上でその予定表を閲覧できる権限があること

new Outlook では COM/VBA 系の自動化が使えない場合があります。基本的には classic Outlook 前提です。

## 使い方

### 1. 設定ファイルを作る

`config.sample.json` を `config.json` にコピーします。

```powershell
Copy-Item .\config.sample.json .\config.json
```

`config.json` の `OutputDirectory` を自分の出力先に変更します。

JSONでは `\` をエスケープする必要があります。

```json
{
  "OutputDirectory": "\\\\fileserver\\share\\schedule",
  "DaysAhead": 14,
  "Title": "Outlook Schedule",
  "WriteCsv": true,
  "IncludeDefaultCalendar": true,
  "DefaultCalendarName": "自分",
  "SharedCalendars": [
    {
      "Name": "田中さん",
      "Email": "tanaka@example.com"
    }
  ],
  "MaskPrivateItems": true,
  "IncludeLocation": true,
  "IncludeBusyStatus": true,
  "IncludeOrganizer": true,
  "IncludeAttendees": true,
  "IncludeBodyPreview": false,
  "BodyPreviewMaxLength": 700,
  "IncludeCategories": true,
  "IncludeReminder": true
}
```

ローカルフォルダで試す場合は、スラッシュ `/` を使うとエスケープミスを避けやすいです。

```json
{
  "OutputDirectory": "C:/temp/outlook-schedule"
}
```

### 2. 共有予定表を追加する

`SharedCalendars` に、Outlookで閲覧権限がある相手のメールアドレスを追加します。

```json
{
  "SharedCalendars": [
    {
      "Name": "田中さん",
      "Email": "tanaka@example.com"
    },
    {
      "Name": "佐藤さん",
      "Email": "sato@example.com"
    }
  ]
}
```

文字列だけでも指定できます。

```json
{
  "SharedCalendars": [
    "tanaka@example.com",
    "sato@example.com"
  ]
}
```

自分の予定を出さず、共有予定表だけ出したい場合は以下にします。

```json
{
  "IncludeDefaultCalendar": false
}
```

注意: ここに書いても、Outlook側で閲覧権限がない予定表は取得できません。その場合はログに警告が出て、その予定表はスキップされます。

### 3. 手動実行

PowerShellで以下を実行します。

```powershell
.\scripts\Export-OutlookSchedule.ps1 -ConfigPath .\config.json
```

実行後、出力先に `schedule.html` が作成されます。

### 4. タスクスケジューラで定期実行

管理者権限が不要な範囲で、現在ユーザーのタスクとして登録する例です。

```powershell
.\scripts\Register-ScheduledTask.ps1 -ConfigPath .\config.json -IntervalMinutes 30
```

登録後は、Outlook が入っている端末にログオンしている間、30分ごとに予定表をHTML/CSVへ出力します。

## 詳細情報の出力設定

詳細情報は `config.json` で制御できます。

| 設定 | 既定 | 内容 |
|---|---:|---|
| `IncludeDefaultCalendar` | `true` | 自分の既定予定表を出力する |
| `DefaultCalendarName` | `自分` | 自分の予定表の表示名 |
| `SharedCalendars` | `[]` | 追加で取得する共有予定表 |
| `IncludeOrganizer` | `true` | 主催者を出力する |
| `IncludeAttendees` | `true` | 必須参加者・任意参加者を出力する |
| `IncludeBodyPreview` | `false` | 本文プレビューを出力する |
| `BodyPreviewMaxLength` | `700` | 本文プレビューの最大文字数 |
| `IncludeCategories` | `true` | 分類を出力する |
| `IncludeReminder` | `true` | リマインダーを出力する |
| `MaskPrivateItems` | `true` | 非公開予定の詳細をマスクする |

本文は機密情報やTeams URLを含みやすいため、既定では `IncludeBodyPreview` を `false` にしています。必要な場合だけ `true` にしてください。

## 文字コード

PowerShell 5.1 系の環境を考慮して、`.ps1` は UTF-8 with BOM で保存しています。

## セキュリティ上の考え方

このツールは便利ですが、予定表には機密情報が含まれます。

詳細表示をONにすると、主催者、参加者、本文プレビューなどがHTML/CSVに保存されます。さらに、共有予定表を追加すると、他者の予定情報もHTML/CSVに保存されます。出力先は、自分専用または必要最小限の権限に絞ったフォルダにしてください。

特に注意が必要な情報:

- 参加者名、メールアドレス
- 会議本文
- Teams URL
- 取引先名、案件名
- 非公開予定
- 共有予定表の内容

非公開予定は、`MaskPrivateItems: true` の場合、件名・場所・主催者・参加者・本文プレビューをマスクします。

## よくあるトラブル

### 共有予定表が取得されない

以下を確認してください。

- Outlookでその人の予定表を手動で開けるか
- `SharedCalendars` のメールアドレスが正しいか
- 予定表の共有権限が付与されているか
- Outlookがオンライン/同期済みか

権限がない場合、スクリプトはログに警告を出して、その予定表をスキップします。

### 認識できないエスケープシーケンスと表示される

`config.json` のWindowsパスで `\` が1個だけになっている可能性があります。

ダメな例:

```json
{
  "OutputDirectory": "C:\Users\yourname\schedule"
}
```

良い例:

```json
{
  "OutputDirectory": "C:\\Users\\yourname\\schedule"
}
```

または:

```json
{
  "OutputDirectory": "C:/Users/yourname/schedule"
}
```

### スクリプトの実行が無効です

会社の実行ポリシーにより `.ps1` の実行が制限されている可能性があります。

```powershell
Get-ExecutionPolicy -List
```

`MachinePolicy` または `UserPolicy` が設定されている場合は、GPOで制御されています。無理に回避せず、情シスに「署名済みスクリプト」または「許可済みタスク」として申請してください。

### Outlook が起動していないと失敗する

環境によっては Outlook が起動済み、またはプロファイル設定済みである必要があります。最初は Outlook を起動した状態で手動実行してください。

### 共有フォルダに書けない

PowerShellを実行しているユーザーで出力先フォルダに書き込みできるか確認してください。

```powershell
New-Item -ItemType Directory -Force "\\fileserver\share\schedule"
```

## リポジトリ構成

```text
.
├── README.md
├── config.sample.json
├── scripts
│   ├── Export-OutlookSchedule.ps1
│   └── Register-ScheduledTask.ps1
└── docs
    └── SECURITY-NOTES.md
```
