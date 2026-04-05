# GitHub Copilot Instructions for boothcsv

## 言語ポリシー
- 回答は日本語で行う。
- PRタイトル・PR本文・レビューコメント・説明コメント・必要なコードコメントは、ユーザーから別指定がない限り日本語で統一する。

## プロジェクト概要
- BOOTH の宛名印刷用 CSV を読み込み、注文票・ラベル・QR 関連情報をブラウザだけで扱うクライアントサイドアプリ。
- 通常の Web / file:// 直開き / Chrome 拡張の 3 モードで動作する。
- サーバー必須ではないが、ローカルサーバー起動用に `boothcsv.ps1` と `boothcsv.sh` がある。
- 永続化は IndexedDB を使い、注文・設定・フォント・カスタムラベル・商品画像をローカル保存する。

## 現在の主要ファイル構成

### エントリポイントとUI
- `boothcsv.html` - メイン画面。固定ヘッダー、サイドバー、処理済み注文一覧、各種テンプレートを持つ。
- `help.html` - 使い方ガイド。
- `css/boothcsv.css`, `css/sidebar.css` - メイン UI とサイドバーのスタイル。

### アプリ本体
- `js/boothcsv.js` - メイン制御。CSV 読み込み、印刷プレビュー生成、アプリモード判定、設定反映、バックアップ/リストア、Yamato 連携設定、各機能の配線を担当する。
- `js/storage.js` - `UnifiedDatabase` と `StorageManager`。IndexedDB スキーマは現在 v11。
- `js/order-repository.js` - 注文データのキャッシュと DB 同期を担う `OrderRepository`。
- `js/processed-orders-panel.js` - 処理済み注文一覧の描画、ソート、フィルタ、複数選択、発送通知導線。
- `js/custom-labels.js` - カスタムラベル編集 UI、保存、集計、複数面計算。
- `js/custom-labels-font.js` - カスタムフォント管理。
- `js/docs-capture.js` - ドキュメント用スクリーンショット撮影の補助。

### Chrome 拡張関連
- `manifest.json` - MV3 マニフェスト。
- `js/extension-bridge.js` - 拡張 UI から既存アプリへの橋渡し。
- `js/extension-background.js` - BOOTH 画面との通信、CSV/QR/発送通知の自動化。
- `js/booth-extension-content.js` - BOOTH 注文詳細ページで QR や発送状態を取得・操作する content script。
- `docs/chrome-extension-dev-notes.md` - 拡張開発時の注意点と実装メモ。

### 外部ライブラリ
- `js/papaparse.min.js` - CSV 解析。
- `js/jsQR.js` - QR 読み取り。

## 実行モードと分岐方針
- 実行モードは `extension` / `web` / `file` の 3 種類。
- モード判定は `js/boothcsv.js` の `getForcedAppMode()` / `detectActualAppMode()` / `resolveAppMode()` / `getAppMode()` に集約されている。
- UI の見た目差分は `document.body.dataset.appMode` を単一の真実源として扱い、できるだけ CSS と共通の表示更新処理で分岐する。
- `?appMode=extension|web|file` による強制上書きは、デバッグやスクリーンショット撮影用として維持する。
- 強制 `extension` モードは見た目確認用であり、Chrome 拡張 API が本当に使えることを意味しない。
- 実際に拡張 API を呼べるかどうかは `canUseExtensionBridge()` や `isRealChromeExtensionApp()` で別途判定する。
- `?extension=1` のような旧式の別パラメータは増やさず、`appMode` に統一する。

## 主要データフロー
1. CSV ファイル選択または拡張経由で CSV 取得。
2. PapaParse でパースし、`OrderRepository.bulkUpsert()` で注文を IndexedDB とメモリキャッシュへ反映。
3. `generateLabels()` などのプレビュー生成処理で注文票やラベルを DOM に描画。
4. 印刷完了後は `updateSkipCount()` がラベルスキップ数を更新し、次回プレビューへ引き継ぐ。
5. 処理済み注文一覧は `processed-orders-panel.js` が `OrderRepository` を参照して再プレビュー・削除・発送通知へつなぐ。

## 永続化とストレージの重要事項
- IndexedDB 名は `BoothCSVStorage`、スキーマバージョンは現在 v11。
- 主なストアは `settings` / `fonts` / `orders` / `customLabels` / `productImages`。
- バックアップ/リストアは `StorageManager.exportAllData()` / `StorageManager.importAllData()` を使い、上記ストアをまとめて JSON 化する。
- カスタムラベルは設定 JSON ではなく独立ストアで扱う前提で考える。
- 商品画像は `productImages` ストアで管理する。注文単位画像とは別の層なので混同しない。

## このリポジトリで特に重要な実装パターン

### 設定変更時の同期
- 設定を変更するときは、`StorageManager.set(...)` だけで終わらせず `window.settingsCache` も必ず更新する。
- 片方だけ更新すると、再プレビューや CSV 再読み込み後に不整合が起きやすい。

```javascript
await StorageManager.set(StorageManager.KEYS.LABEL_SKIP, newValue);
settingsCache.labelskip = newValue;
```

### 注文データアクセス
- 注文データは `window.orderRepository` を経由して扱う。
- 直接 IndexedDB を読みに行く前に、既存の repository API で足りるか確認する。
- 注文番号は `OrderRepository.normalize()` 相当の正規化を前提に扱う。

### 処理済み注文一覧との整合
- 注文の印刷済み・発送済み状態を変える変更は、処理済み注文一覧の UI と repository 更新通知まで含めて考える。
- 単発プレビューだけ直して一覧を置き去りにしない。

### カスタムラベルの安全性
- リッチテキストを扱うときは、既存の `sanitizeCustomLabelHTML()` / `sanitizeCustomLabelStyle()` の方針に合わせる。
- カスタムラベルは内部モデルと差分保存を前提にしているので、DOM を直接いじるだけの実装を増やさない。

### 画像・Blob URL
- Blob URL を作る処理では、既存の `revokeBlobUrl()` / `revokeBlobImages()` の流れを崩さない。
- 画像保存形式が `ArrayBuffer` か Blob URL かを確認し、既存の変換経路を再利用する。

## Chrome 拡張まわりの注意
- 見た目が拡張モードでも、bridge が使えないケースを常に考慮する。
- CSV 自動取得、QR 自動取得、発送通知は `extension-background.js`・`booth-extension-content.js`・`extension-bridge.js` の役割分担を崩さず実装する。
- 発送通知はボタン押下だけで成功扱いにせず、最終的に発送日時などの状態確認まで取れてから保存する。
- QR 取得は画像 URL 抽出、data URL 化、jsQR 解析、受付番号/パスワードのフォールバック保存まで含めて考える。
- Chrome 拡張機能は「ゆるく」非公開の扱いとし、`README.md` や `help.html` などの公開寄りドキュメントでは積極的に案内しない。
- ここでいう「ゆるく」非公開とは、ユーザーが自力で使う工夫まで積極的に妨げる意味ではないが、BOOTH 側サーバーに自動処理で関与する都合上、他者に強く勧めたり表立って宣伝したりしない、という運用方針を指す。
- 厳密には BOOTH 側の規約違反になり得る前提で扱い、公開文書では通常機能のように紹介しない。必要なメモや実装知見は開発者向け文脈に閉じて記述する。

## UI / CSS の作法
- 新しいボタンは既存の共有クラスを優先して組み合わせる。
- 基本は `btn-base` + サイズ系 + 色系で構成し、必要なら `btn-header` を重ねる。
- 既存色 (`btn-primary`, `btn-success`, `btn-danger`, `btn-info`, `btn-warning`, `btn-download`, `btn-print-accent`) を先に使う。
- app mode による見た目差分は、個別要素ごとの場当たり的な `hidden` 切り替えより、ルート状態と CSS セレクタを優先する。

## ドキュメント・スクリーンショット関連
- `docs/images/` 配下の画像と `help.html` / `README.md` は連動している前提で扱う。
- UI のスクリーンショット採取には `js/docs-capture.js` と `scripts/docs-capture/` の定義が使われる。
- スクリーンショットやヘルプ更新時は、必要に応じて `?debug=1&appMode=...` を使い、再現可能な状態で撮る。

## ローカル開発と確認方法
- 簡易確認は `boothcsv.html` を直接開いてもよいが、Web モード確認や docs-capture 利用時はローカルサーバー起動を優先する。
- Windows は `.\boothcsv.ps1`、macOS / Linux は `./boothcsv.sh` を使える。
- サンプルデータは `sample/` 配下を使う。
- デバッグログは `?debug=1` で有効化され、`[csv]` / `[repo]` / `[image]` などのカテゴリを確認できる。
- IndexedDB の確認はブラウザ DevTools の Application タブを使う。

## テスト方針
- 現在、このリポジトリに自動テスト基盤はない前提で進める。
- 変更時は、対象機能に応じて手動で確認手順を組み立てる。
- とくに以下は壊しやすいので影響範囲に含める。
  - CSV 読み込み
  - ラベルスキップ数の反映
  - 処理済み注文一覧の表示と操作
  - カスタムラベル保存
  - バックアップ/リストア
  - 拡張モードの表示差分と bridge 可否

## よくある変更の観点
- 新しい設定を追加するなら、`StorageManager.KEYS`、デフォルト値、`settingsCache`、UI バインドを揃えて更新する。
- DB スキーマ変更時は `UnifiedDatabase.version` を上げ、`createObjectStores()` と必要な移行処理を追加する。
- 注文データの項目追加時は、repository キャッシュ・一覧 UI・バックアップ形式まで確認する。
- 拡張機能の変更時は、実ページの DOM 依存と既存フォールバックの両方を確認する。
