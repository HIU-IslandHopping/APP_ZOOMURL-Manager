# ZOOM ROOM自動作成GAS - セットアップガイド

## 前提条件

- Googleアカウント
- ZOOM Proアカウント
- ZOOM Marketplace Developerアカウント

## セットアップ手順

### Step 1: ZOOM API認証情報の取得

1. **ZOOM Marketplace にアクセス**
   - https://marketplace.zoom.us/ にアクセス
   - ZOOMアカウントでログイン

2. **アプリの作成**
   - 「Develop」→「Build App」をクリック
   - 「Server-to-Server OAuth」を選択
   - アプリ情報を入力：
     - **App Name**: `HIU冬合宿 ZOOM ROOM管理`（任意の名前）
     - **Company name**: `HIU冬合宿実行委員会`（任意の名前）
     - **Developer contact information**: 連絡先メールアドレス
   - 「Create」をクリック

3. **認証情報の取得**

   **Server-to-Server OAuth**（推奨）の場合：
   - 「App Credentials」タブを開く
   - 以下の3つの値をコピーします：
     
     | 項目 | 説明 | どこから取得 |
     |------|------|-------------|
     | **Account ID** | ZOOMアカウントID |
     | **Client ID** | アプリのクライアントID  |
     | **Client Secret** | アプリのシークレットキー  |
   
   ⚠️ **重要**: Client SecretとAPI Secretは機密情報です。安全に管理してください。
   
   **💡 Tokenについて**:
   - App作成時にTokenが表示される場合がありますが、**これは不要です**
   - Server-to-Server OAuthでは、アクセストークンを動的に生成します
   - コード内で`Account ID`、`Client ID`、`Client Secret`を使用して自動でトークンを取得します

4. **アクセス権限の設定（必須）**
   - 「Scopes」タブを開く
   - 「Add Scopes」ボタンをクリック
   - 検索欄でスコープを検索して追加：
     
     **最低限必要なスコープ**：
     - `meeting:write` - ミーティングの作成（**必須**）
     
     **推奨スコープ**：
     - `meeting:read` - ミーティング情報の取得
     - `user:read` - ユーザー情報の取得
   
   ⚠️ **注意**: スコープは少なくとも1つ以上追加する必要があります。`meeting:write`は必ず追加してください。

5. **アプリの有効化**
   - 「Activation」タブを開く
   - 「Activate your app」をクリック
   - 利用条件に同意して有効化

### Step 2: Googleスプレッドシートの準備

1. **スプレッドシートを確認**
   - スプレッドシートID: `1zYJa2d8vEqByHcesi5C-OUyhexweZR28Iia0mkUaFqs`
   - シートID: `1942334831`

2. **カラムを確認**
   - 必要なカラムが揃っているか確認
   - ヘッダー: `add`, `user`, `title`, `password`, `star`

### Step 3: GASスクリプトの作成

1. **Apps Scriptエディタを開く**
   - Googleスプレッドシートを開く
   - 「拡張機能」→「Apps Script」を選択

2. **コードをコピー**
   - `zoom/create-zoom-rooms.gs` の内容をすべてコピー
   - `CreateMenu.js` の内容をすべてコピー
   - Apps Scriptエディタに2つのファイルとして貼り付け
   - 「保存」をクリック（プロジェクト名: `ZOOM ROOM自動作成`）
   
   **ファイル構成**:
   - `Code.gs` (または `create-zoom-rooms.gs`): メインの処理コード
   - `CreateMenu.js` (または任意の名前): カスタムメニューのコード

3. **認証情報の設定**

#### 方法A: setupConfiguration() 関数を使用（推奨）

1. `setupConfiguration()` 関数を開く
2. 認証方式に応じて以下を編集：

   **Server-to-Server OAuth**の場合：
   ```javascript
   const configOAuth2 = {
     'ZOOM_AUTH_TYPE': 'OAUTH2',
     'ZOOM_ACCOUNT_ID': '取得したAccount ID',           // Step 1で取得
     'ZOOM_CLIENT_ID': '取得したClient ID',             // Step 1で取得
     'ZOOM_CLIENT_SECRET': '取得したClient Secret',     // Step 1で取得
     'ZOOM_USER_EMAIL': 'test2025@gmail.com'       // あなたのZOOMアカウントのメールアドレス
   };
   const config = configOAuth2;
   ```
   
   **ZOOM_USER_EMAILについて**:
   - ZOOMに登録しているメールアドレスを指定します
   - 例: `test2025@gmail.com`
   - このメールアドレスは、作成されたミーティングのホストとなります
3. `setupConfiguration` 関数を実行
4. 実行後、関数内の認証情報を削除
5. 再保存

#### 方法B: プロジェクトプロパティから設定

1. Apps Scriptエディタで「プロジェクトの設定」を開く
2. 「スクリプトのプロパティ」セクションで以下を追加：

   **Server-to-Server OAuth**の場合：
   - プロパティ: `ZOOM_AUTH_TYPE`, 値: `OAUTH2`
   - プロパティ: `ZOOM_ACCOUNT_ID`, 値: `[取得したAccount ID]`
   - プロパティ: `ZOOM_CLIENT_ID`, 値: `[取得したClient ID]`
   - プロパティ: `ZOOM_CLIENT_SECRET`, 値: `[取得したClient Secret]`
   - プロパティ: `ZOOM_USER_EMAIL`, 値: `[ZOOMアカウントのメールアドレス]`
   
3. 「スクリプトのプロパティを保存」をクリック

### Step 4: 動作確認

1. **テスト実行**
   - `testCreateZoomRoom` 関数を選択
   - 「実行」ボタンをクリック
   - 初回は認証が求められます
   - 実行ログを確認

2. **設定確認**
   - `checkConfiguration` 関数を実行
   - ログで設定値が正しく表示されるか確認

3. **実際の実行**
   - スプレッドシートにテストデータを1件追加
   - starカラムは空白のまま
   - `createZoomRooms` 関数を実行
   - スプレッドシートに結果が記録されるか確認

### Step 5: 自動実行の設定（オプション）

1. **トリガーの追加**
   - Apps Scriptエディタで「トリガー」をクリック
   - 「トリガーを追加」をクリック
   - 以下を設定：
     - 実行する関数: `createZoomRooms`
     - イベントソース: 時間主導型
     - 時間ベースのトリガー: 1時間おき
   - 「保存」をクリック

2. **確認**
   - トリガーが正常に作成されたか確認
   - 次回の実行時間を確認

## トラブルシューティング

### 認証エラー

**問題**: 「認証に失敗しました」というエラー

**解決方法**:
1. API KeyとAPI Secretが正しいか確認
2. JWT生成部分のコードを確認
3. トークンの有効期限を確認（1時間）

### 権限エラー

**問題**: 「スプレッドシートにアクセスできません」というエラー

**解決方法**:
1. Apps Scriptの実行権限を確認
2. スプレッドシートの共有設定を確認
3. 必要に応じて権限を付与

### API制限エラー

**問題**: 「Rate limit exceeded」というエラー

**解決方法**:
1. リクエスト間隔を調整
2. 処理する件数を制限
3. トリガーの実行頻度を調整

## 次のステップ

- [ ] 実際のZOOM ROOM作成をテスト
- [ ] トリガーを設定して自動実行
- [ ] エラーメール通知の実装
- [ ] ログ管理の改善
- [ ] **Googleカレンダー連携機能の実装**（予定）
  - ZOOM ROOM作成後に自動的にカレンダーイベントを追加
  - 特定のカレンダーIDを設定
  - ZOOMリンク、パスワードを含めた詳細情報を自動入力

---

**作成日**: 2025年10月22日  
**最終更新**: 2025年10月22日
