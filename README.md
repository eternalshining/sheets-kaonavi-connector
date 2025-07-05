# Googleスプレッドシートアドオン「カオナビデータ連携」

## 概要

カオナビに蓄積された従業員データやカスタムシートの情報を、Googleスプレッドシートへ簡単に出力できるアドオンです。

## 機能

### 1. メンバー情報の取得
- カオナビのメンバー情報を取得し、スプレッドシートに出力
- 新規シート作成または既存シートへの上書きを選択可能

### 2. シート情報の取得
- 特定のカスタムシートの情報を取得し、スプレッドシートに出力
- シート選択UIで対象シートを選択

### 3. カスタムシートの作成
- 複数のデータソースから必要な項目のみを抽出
- 定義シート（kaonavi_custom_def）で取得項目を設定

### 4. 認証情報の管理
- カオナビAPI認証情報の安全な保存
- 認証情報の更新機能

## セットアップ

### 前提条件
- Google Workspaceアカウント
- カオナビAPI認証情報（Consumer Key、Consumer Secret）
- Node.js（clasp CLIツール用）

### インストール手順

1. **プロジェクトのクローン**
   ```bash
   git clone https://github.com/eternalshining/sheets-kaonavi-connector.git
   cd sheets-kaonavi-connector
   ```

2. **依存関係のインストール**
   ```bash
   npm install
   ```

3. **クラスプのインストール（グローバル）**
   ```bash
   npm install -g @google/clasp
   ```

4. **Google Apps Scriptにログイン**
   ```bash
   clasp login
   ```

5. **Google Apps Scriptプロジェクトの作成**
   ```bash
   clasp create --type sheets --title "カオナビデータ連携"
   ```
   
   **注意**: 既に`.clasp.json`ファイルが存在する場合、以下のエラーが発生する可能性があります：
   ```
   Project file already exists.
   ```
   
   この場合は、以下のコマンドで既存の設定ファイルを削除してから再実行してください：
   ```bash
   rm .clasp.json
   clasp create --type sheets --title "カオナビデータ連携"
   ```
   
   または、`.clasp.json.template`ファイルをコピーして設定してください：
   ```bash
   cp .clasp.json.template .clasp.json
   # .clasp.jsonファイルを編集して、YOUR_SCRIPT_IDとYOUR_PARENT_IDを実際の値に置き換え
   ```

6. **rootDirの設定確認**
   作成された`.clasp.json`ファイルを開き、`rootDir`が`"./src"`に設定されていることを確認してください：
   ```json
   {
     "scriptId": "YOUR_SCRIPT_ID",
     "rootDir": "./src",
     "parentId": "YOUR_PARENT_ID"
   }
   ```

7. **コードのアップロード**
   ```bash
   clasp push
   ```
   
   マニフェストファイルの更新を確認するダイアログが表示されたら、`y`を選択してください。

8. **スプレッドシートでの動作確認**
   - `clasp create`コマンドで表示されたスプレッドシートのURLを開く
   - ページを更新すると「カオナビ連携」メニューが表示されます

## 使用方法

### 初回設定
1. 作成されたスプレッドシートを開く
2. 「カオナビ連携」メニューから任意の機能を選択
3. 認証情報入力ダイアログが表示されるので、Consumer KeyとConsumer Secretを入力
4. 認証情報は安全に保存されます

### メンバー情報の取得
1. 「カオナビ連携」→「メンバー情報を取得」を選択
2. 出力方法を選択（新規シート作成：「はい」/現在のシートに上書き：「いいえ」）
3. データが自動的にスプレッドシートに出力されます

### シート情報の取得
1. 「カオナビ連携」→「シート情報を取得」を選択
2. 取得したいシートを選択（番号を入力）
3. 出力方法を選択（新規シート作成：「はい」/現在のシートに上書き：「いいえ」）
4. データが自動的にスプレッドシートに出力されます

### カスタムシートの作成
1. スプレッドシート内に「kaonavi_custom_def」という名前のシートを作成
2. A列にデータソース（「基本情報」またはシート名）、B列に項目名を入力
3. 「カオナビ連携」→「カスタムシートを作成」を選択
4. 定義に基づいてカスタムシートが作成されます

#### 定義シートの例
```
A列（データソース） | B列（項目名）
基本情報            | 氏名
基本情報            | 部署
評価シート          | 評価点
スキルシート        | プログラミング
```

### 認証情報の更新
1. 「カオナビ連携」→「認証情報を更新」を選択
2. 確認ダイアログで「はい」を選択
3. 新しい認証情報を入力

## 開発

### ローカル開発
```bash
# 依存関係のインストール
npm install

# テストの実行
npm test

# コードのアップロード
npm run push
```

### ファイル構成
```
src/
├── main.js                 # メインアプリケーション
├── auth-service.js         # 認証情報管理
├── kaonavi-service.js      # カオナビAPI通信
├── ui-service.js           # UI関連機能
├── sheet-service.js        # シート操作
├── custom-sheet-service.js # カスタムシート作成
└── appsscript.json         # Google Apps Script設定
```

## 注意事項

- カオナビAPIの認証情報は、Google Apps Scriptの`PropertiesService`を使用して安全に保存されます
- 大量のデータを処理する場合、Google Apps Scriptの実行時間制限（6分）に注意してください
- APIレート制限に注意し、必要以上の頻繁なリクエストは避けてください
- 実際のカオナビAPIの仕様に応じて、認証トークンの生成ロジックの調整が必要になる場合があります

## トラブルシューティング

### セットアップ時の問題

1. **"Project file already exists" エラー**
   ```bash
   rm .clasp.json
   clasp create --type sheets --title "カオナビデータ連携"
   ```

2. **"rootDir" 設定の問題**
   - `.clasp.json`ファイルで`"rootDir": "./src"`が設定されていることを確認
   - 設定後、`clasp push`を再実行

3. **認証に関する問題**
   - `clasp login`でGoogleアカウントにログインしているか確認
   - 必要に応じて`clasp logout`してから再度ログイン

### 実行時の問題

1. **認証エラー**
   - Consumer KeyとConsumer Secretが正しいことを確認
   - 認証情報を更新してみてください

2. **シートが見つからない**
   - カオナビでシートへのアクセス権限があることを確認
   - シート名が正確であることを確認

3. **データが表示されない**
   - カオナビでデータが存在することを確認
   - Apps Scriptエディタで実行ログを確認

4. **実行時間エラー**
   - 大量のデータを処理する場合は、データを分割して処理することを検討

## プロジェクトの確認

デプロイ後、以下のURLから作成されたプロジェクトを確認できます：
- **スプレッドシート**: `clasp create`実行時に表示されるGoogle DriveのURL
- **Apps Scriptエディタ**: `clasp create`実行時に表示されるscript.google.comのURL

## ライセンス

MIT License