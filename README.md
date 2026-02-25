# Get-FlattenedGroupMembers v3.0

Outlookのグループアドレス（配布リスト）を再帰的に展開して、個人のメールアドレスのリストを取得するPowerShellスクリプトです。

## 概要

このスクリプトは以下の機能を提供します：

- 📧 **統一入力処理**: 個人名、グループ名、メールアドレスを同時処理
- 🌐 **日本語対応**: 日本語の個人名から自動的にメールアドレスを解決
- 🔄 **再帰的展開**: ネストしたグループアドレスの完全展開
- 🚫 **重複排除**: 重複メールアドレスの自動排除
- 🔍 **フィルタリング**: 内部/外部メールアドレスのフィルタリング
- ✅ **アクティブユーザー**: アクティブユーザーのみの抽出
- 📊 **複数出力形式**: Array、CSV、JSON形式での出力
- 🛡️ **循環参照検出**: 循環参照の検出と回避
- 📝 **詳細ログ**: 詳細なログ出力

## v3.0の新機能

### 🎯 統一パラメータ設計
- 単一の `-Inputs` パラメータで全ての入力タイプを処理
- 複雑なParameterSetを廃止し、シンプルな使用方法を実現

### 🌏 混合入力対応
- 個人名、グループ名、メールアドレスを同時に指定可能
- 日本語の個人名から自動的にメールアドレスを解決

## 前提条件

- PowerShell 7.x
- Microsoft Outlook がインストールされ、正しく設定されている
- MAPIプロファイルが設定されている
- Global Address List への読み取りアクセス権限

### Outlook COM版の利点
- ✅ Exchange Online接続が不要
- ✅ 認証プロセスが簡素化
- ✅ ローカルキャッシュを使用可能
- ✅ オフライン環境での動作
- ✅ Exchange管理者権限が不要

## インストール

1. このリポジトリをクローンまたはダウンロードします
2. PowerShellでスクリプトディレクトリに移動します
3. 実行ポリシーを確認します（必要に応じて変更）

```powershell
Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 基本的な使用方法

### v3.0統一構文

```powershell
# 単一グループの展開
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("team@company.com")

# 複数グループの展開
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("team1@company.com", "team2@company.com")

# 混合入力の処理（v3.0の新機能）
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("田中太郎", "team@company.com", "john.doe@company.com")

# 日本語名の処理
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("田中 太郎", "佐藤 花子 （営業部）", "山田 次郎")
```

### オプション付きの実行

```powershell
# 外部メールアドレスを除外
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("all-staff@company.com") -ExcludeExternal

# アクティブユーザーのみを含める
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("team@company.com") -OnlyActiveUsers

# CSV形式で出力
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("team@company.com") -OutputFormat CSV | Out-File "emails.csv"

# 詳細ログを有効にする
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @("team@company.com") -LogLevel Debug
```

## パラメータ

| パラメータ | 型 | 必須 | 説明 |
|-----------|----|----|------|
| `Inputs` | string[] | ○ | 処理する入力のリスト（個人名、グループ名、メールアドレス） |
| `MaxDepth` | int | × | 再帰展開の最大深度（デフォルト: 10） |
| `ExcludeExternal` | switch | × | 外部メールアドレスを除外 |
| `OnlyActiveUsers` | switch | × | アクティブなユーザーのみを含める |
| `OutputFormat` | string | × | 出力形式（Array, CSV, JSON）（デフォルト: Array） |
| `LogLevel` | string | × | ログレベル（None, Error, Warning, Info, Debug）（デフォルト: Info） |

## 実際の使用例

### 混合入力の実行例

```powershell
.\src\Get-FlattenedGroupMembers.ps1 -Inputs @(
    'Maeda Akie （前田 明恵）',
    'Ienaka Michinori （家中 孔憲）',
    'Sakamoto Kenji （坂本 賢治）',
    'RSI利用テーマのPM/設計L',
    'Noah developers',
    'RiDP Toolbox Team',
    'zjp_legal_doc_align_poc_v2@jp.ricoh.com',
    'ryoh.shimomoto@jp.ricoh.com'
) -LogLevel Info
```

**実行結果:**
- 入力数: 8個
- 解決成功: 7個（87.5%）
- 最終出力: 354個のユニークなメールアドレス
- 処理時間: 約39秒

詳細な実行例は [`examples/mixed-input-example.md`](examples/mixed-input-example.md) を参照してください。

## 出力例

### Array形式（デフォルト）
```
akie.maeda@jp.ricoh.com
kenji_sakamoto@jp.ricoh.com
taisuke.hosokawa@jp.ricoh.com
shingo.tamura@jp.ricoh.com
...
```

### CSV形式
```csv
EmailAddress
akie.maeda@jp.ricoh.com
kenji_sakamoto@jp.ricoh.com
taisuke.hosokawa@jp.ricoh.com
...
```

### JSON形式
```json
[
  "akie.maeda@jp.ricoh.com",
  "kenji_sakamoto@jp.ricoh.com",
  "taisuke.hosokawa@jp.ricoh.com"
]
```

## 高度な使用例

### 複合条件での実行

```powershell
.\src\Get-FlattenedGroupMembers.ps1 `
    -Inputs @("team1@company.com", "田中太郎", "team2@company.com") `
    -ExcludeExternal `
    -OnlyActiveUsers `
    -MaxDepth 8 `
    -LogLevel Info
```

### バッチ処理

```powershell
$inputs = @("team1@company.com", "田中太郎", "team2@company.com")
$allEmails = .\src\Get-FlattenedGroupMembers.ps1 -Inputs $inputs
$uniqueEmails = $allEmails | Sort-Object | Get-Unique
$uniqueEmails | Export-Csv -Path "all_team_emails.csv" -NoTypeInformation
```

### エラーハンドリング付きの実行

```powershell
try {
    $emails = .\src\Get-FlattenedGroupMembers.ps1 -Inputs @("team@company.com", "田中太郎") -LogLevel Info
    Write-Host "成功: $($emails.Count) 個のメールアドレスを取得しました"
    
    # 統計情報を表示
    $domains = $emails | ForEach-Object { ($_ -split '@')[1] } | Group-Object | Sort-Object Count -Descending
    Write-Host "ドメイン別統計:"
    $domains | ForEach-Object { Write-Host "  $($_.Name): $($_.Count) 個" }
    
} catch {
    Write-Error "エラーが発生しました: $($_.Exception.Message)"
}
```

## トラブルシューティング

### よくある問題と解決方法

#### 1. "Outlook COM接続に失敗しました"

**解決方法:**
- Microsoft Outlookがインストールされているか確認
- Outlookが正しく設定されているか確認
- MAPIプロファイルが設定されているか確認

#### 2. "個人名が解決できません"

**原因:**
- 個人名の表記が間違っている
- Global Address Listに登録されていない

**解決方法:**
- 個人名の表記を確認（姓名の順序、スペース、括弧など）
- Outlookのアドレス帳で検索可能か確認

#### 3. "グループが見つかりません"

**原因:**
- グループ名またはメールアドレスが間違っている
- グループへのアクセス権限がない

**解決方法:**
- グループ名/アドレスを確認
- 管理者に権限を確認

#### 4. "処理が遅い"

**対策:**
- `MaxDepth` パラメータで再帰深度を制限
- 大きなグループは分割して処理
- `LogLevel` を `Error` に設定してログを削減

#### 5. "メモリ不足"

