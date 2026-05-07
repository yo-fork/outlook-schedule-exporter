# Outlook Schedule Exporter

Outlook の予定表を読み取り、単体 HTML と CSV を出力するための PowerShell スクリプトです。

## 目的

- Outlook が気に入らないけど、Outlook の予定を見ないといけない時
- 共有フォルダ等に `schedule.html` を置き、ブラウザだけで見られるようにする
- 予定表の本文、参加者、Teams URL、添付などを出力しない
- 非公開予定は `非公開予定` としてマスクする

## 出力されるもの

```text
schedule.html  # 普段見る用。検索、今日表示、全期間表示、印刷に対応
schedule.csv   # Excel確認、簡易検索、バックアップ用
schedule.log   # 実行ログ
```

## 前提

- classic Outlook が入っている Windows 端末
- Outlook に予定表が同期されていること
- PowerShell が利用できること
- 出力先フォルダへの書き込み権限があること

new Outlook では COM/VBA 系の自動化が使えない場合があります。基本的には classic Outlook 前提です。

## 使い方

### 1. 設定ファイルを作る

`config.sample.json` を `config.json` にコピーします。

```powershell
Copy-Item .\config.sample.json .\config.json
```

`config.json` の `OutputDirectory` を自分の出力先に変更します。

```json
{
  "OutputDirectory": "\\\\fileserver\\share\\schedule",
  "DaysAhead": 14,
  "Title": "Outlook Schedule",
  "WriteCsv": true,
  "MaskPrivateItems": true,
  "IncludeLocation": true,
  "IncludeBusyStatus": true
}
```

### 2. 手動実行

PowerShellで以下を実行します。

```powershell
.\scripts\Export-OutlookSchedule.ps1 -ConfigPath .\config.json
```

実行後、出力先に `schedule.html` が作成されます。

### 3. タスクスケジューラで定期実行

管理者権限が不要な範囲で、現在ユーザーのタスクとして登録する例です。

```powershell
.\scripts\Register-ScheduledTask.ps1 -ConfigPath .\config.json -IntervalMinutes 30
```

登録後は、Outlook が入っている端末にログオンしている間、30分ごとに予定表をHTML/CSVへ出力します。

## セキュリティ上の考え方

このツールは、最初から出力情報を絞る設計にしています。

出力する情報:

- 開始時刻
- 終了時刻
- 件名
- 場所
- 空き時間種別

出力しない情報:

- 本文
- 参加者
- メールアドレス
- Teams URL
- 添付
- 非公開予定の詳細

それでも、予定の件名だけで機密情報になる場合があります。出力先は、自分専用または必要最小限の権限に絞ったフォルダにしてください。

## よくあるトラブル

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
