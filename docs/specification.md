# Mermaid for Excel 仕様書

本書は現行実装（`src/taskpane/taskpane.ts`, `src/taskpane/taskpane.html`, `manifest.xml`, `build.js`, `Installer/MermaidForExcel.iss`）から起こした仕様書である。記述は実装の挙動を正とする。

- 対象バージョン: 1.0.0（`package.json` / マニフェスト `1.0.0.0`）
- 作成日: 2026-10-21

---

## 1. 概要

Mermaid 記法のテキストからダイアグラムを生成し、Excel のアクティブシートへ画像（PNG / JPEG / SVG）として挿入する Office タスクペインアドイン。

- レンダリングはタスクペイン内（ブラウザ）で完結し、入力内容を外部へ送信しない。
- 挿入した図形の代替テキストに Mermaid コードを保持し、後から読み出して再編集・差し替えできる。

### 1.1 動作環境

| 項目 | 内容 |
|---|---|
| ホスト | Excel（`Workbook`）のみ |
| フォームファクタ | デスクトップ（`DesktopFormFactor`）のみ |
| 権限 | `ReadWriteDocument` |
| 要件セット（マニフェスト） | `ExcelApi 1.9` |
| 要件セット（実行時判定） | `ImageCoercion 1.2`（SVG 挿入時のみ。§4.4） |
| 主要ライブラリ | mermaid `^11.0.0`（バンドル同梱）、Office.js（Microsoft CDN から読み込み） |

ImageCoercion 1.2 をマニフェストに宣言しないのは、非対応環境でアドイン全体が読み込めなくなるのを避け、PNG / JPEG は使えるようにするため。

---

## 2. アドイン起動

| 項目 | 内容 |
|---|---|
| リボン | [ホーム] タブにグループ「Mermaid For Excel」を追加 |
| ボタン | 「Open Mermaid」（`OpenPaneButton`）。押下でタスクペイン `MermaidPane` を表示 |
| タスクペイン URL | 開発: `https://localhost:3000/taskpane.html` / 本番: `https://msmsrep.github.io/MermaidForExcel/dist/taskpane.html` |
| アドイン ID | `4f7e8c1b-2a3d-4e5f-9071-a2b3c4d5e6f7` |

---

## 3. 画面仕様（タスクペイン）

タスクペインの高さいっぱい（`100dvh`）に、上から順に以下の要素を縦に配置する。プレビューが残りの高さを占める。

| # | 要素 | ID | 内容 |
|---|---|---|---|
| 1 | タイトル | — | 「Mermaid for Excel」 |
| 2 | シート図形選択 | `sheet-diagram-select` | ラベル「Sheet diagram」。先頭は `(No diagram selected)`（値は空）。以降にアクティブシート上の Mermaid 図形を図形名で列挙 |
| 3 | Mermaid 入力欄 | `mermaid-input` | 等幅フォントの textarea（初期高さ 100px、スペルチェック無効）。初期値はサンプルのフローチャート（§3.1） |
| 4 | スプリッタ | `splitter` | 入力欄とプレビューの境界。ドラッグで入力欄の高さを変更（§4.7） |
| 5 | 形式選択 | `format-select` | `PNG`（既定）/ `JPEG` / `SVG` |
| 6 | 挿入ボタン | `insert-btn` | 「Insert to Excel」。初期状態は無効 |
| 7 | ダウンロードボタン | `download-btn` | 「Download」。初期状態は無効 |
| 8 | エラー表示 | `error` | エラーメッセージを赤字で表示 |
| 9 | プレビュー | `preview` | レンダリング結果の SVG を表示（最小高さ 120px、はみ出しはスクロール）。入力が空のときは「Type Mermaid code」を表示 |

形式選択は挿入とダウンロードの両方に適用される。

### 3.1 入力欄の初期値

```
graph LR
    A[Start] --> B{Condition}
    B -->|Yes| C[Process A]
    B -->|No| D[Process B]
    C --> E[End]
    D --> E
```

---

## 4. 機能仕様

### 4.1 初期化

`Office.onReady` 後に以下を行う。

1. mermaid を初期化する。

   | 設定 | 値 | 目的 |
   |---|---|---|
   | `startOnLoad` | `false` | 自動描画しない |
   | `theme` | `default` | — |
   | `htmlLabels` / `flowchart.htmlLabels` | `false` | ラベルを `<foreignObject>` ではなく SVG `<text>` で出力する（Excel 挿入時・画像化時に文字が消えるのを防ぐ） |
   | `flowchart.useMaxWidth` | `false` | 図を実寸で出力する |

2. 入力欄の初期値でプレビューを描画する（§4.2）。
3. シート図形一覧を読み込む（§4.3）。
4. ワークシートの `onActivated` イベントを登録する（§4.3）。

### 4.2 プレビュー描画

- **トリガー**: 入力欄の `input` イベント。最後の入力から **250ms** のデバウンス後に描画する。シート図形を選択したときも即時描画する。
- **処理**:
  1. エラー表示をクリアし、挿入・ダウンロードボタンを無効化する。
  2. 入力をトリムし、空ならヒント表示に戻して終了する。
  3. 既存の `#mermaid-graph` 要素を削除してから `mermaid.render("mermaid-graph", code)` を実行する。
  4. 成功時はプレビューに SVG を表示し、両ボタンを有効化する。
  5. 失敗時はプレビューを空にし、`Render error: <メッセージ>` を表示する。
- **競合制御**: 描画ごとにリクエスト ID を採番し、完了時に最新 ID でなければ結果を破棄する（古い描画結果で上書きしない）。

### 4.3 シート図形一覧

- **対象**: アクティブシートの図形のうち、`type === "Image"` かつ `altTextTitle === "Mermaid Diagram"` のもの。画像形式（PNG / JPEG / SVG）は問わない。
- **表示**: 図形名（`shape.name`）。値は図形 ID。
- **読み込みタイミング**:
  - タスクペイン起動時
  - ドロップダウンにフォーカスしたとき
  - 挿入処理の完了後（挿入した図形を選択状態にする）
- **選択状態の保持**: 再読み込み時、直前の選択 ID が一覧に存在すれば選択を維持し、なければ `(No diagram selected)` に戻す。
- **シート切替時**: 一覧を空（`(No diagram selected)` のみ）にリセットする。再読み込みは次のフォーカス時。
- **読み込み失敗時**: エラーは表示せず、一覧を空にする。
- **選択時の動作**: 選択した図形の代替テキスト（説明）を入力欄に設定し、即座にプレビューを描画する。`(No diagram selected)` を選んだ場合は何もしない（入力欄は変更しない）。

### 4.4 Excel への挿入

挿入ボタン押下時、プレビューの SVG を形式に応じて挿入する。プレビューに SVG がない場合は `Please render first` を表示する。

**配置**:

| シート図形選択 | 位置 | 図形名 | 元図形 |
|---|---|---|---|
| なし | アクティブセルの左上 | Excel の既定 | — |
| あり | 選択図形と同じ位置 | 選択図形の名前を引き継ぐ | 新しい図形の挿入後に削除 |

**形式別の処理**:

- **PNG / JPEG**: §4.6 で画像化し、`shapes.addImage` で追加する。表示サイズは SVG の実寸（px × 0.75pt）に設定する。追加・プロパティ設定・元図形の削除は 1 回の `sync` で行う。
- **SVG**:
  1. `ImageCoercion 1.2` 非対応の場合は `SVG insertion is not supported in this version of Excel. Use PNG or JPEG.` エラーとする。
  2. 挿入前にシート上の図形 ID 一覧と配置先を取得する。
  3. `setSelectedDataAsync`（`CoercionType.XmlSvg`）で挿入する。
  4. 挿入前後の図形 ID の差分から新しい図形を特定する。ちょうど 1 個でなければ `Could not identify the inserted shape.` エラーとする。
  5. 新しい図形に位置・名前・代替テキストを設定する。
  6. 差し替えの場合、元図形がまだ存在すれば削除する。

  挿入が失敗した場合、元図形は削除されずに残る。

**共通の付与情報**:

| 図形プロパティ | 値 |
|---|---|
| `altTextTitle` | `Mermaid Diagram` |
| `altTextDescription` | 入力欄の Mermaid コード（前後空白をトリム） |

**エラー**: `Insert error: <メッセージ>` を表示する。差し替え対象の図形が既に存在しない場合のメッセージは `Selected diagram no longer exists.`。

### 4.5 ダウンロード

ダウンロードボタン押下時、選択形式でファイルを生成し、`<a download>` のクリックで保存させる。プレビューに SVG がない場合は `Please render first` を表示する。

| 形式 | ファイル名 | 生成方法 |
|---|---|---|
| PNG | `diagram.png` | §4.6 |
| JPEG | `diagram.jpg` | §4.6 |
| SVG | `diagram.svg` | プレビュー SVG を `XMLSerializer` で文字列化 |

失敗時は `Download error: <メッセージ>` を表示する。

### 4.6 画像変換（SVG → ラスタ）

1. SVG を `XMLSerializer` で文字列化し、`data:image/svg+xml` URL として `Image` に読み込む。
2. 基準サイズ（CSS px）を決定する。
   - 幅: `viewBox.width`（> 0 の場合）→ 画像の `naturalWidth` → `800`
   - 高さ: `viewBox.height`（> 0 の場合）→ `naturalHeight` → `600`
3. 倍率を決定する。既定は **2 倍**。ただし Canvas の面積が 4096 × 4096 px を超える場合は、超えない倍率まで下げる。
4. Canvas（基準サイズ × 倍率）に描画し、Base64 を得る。

| 形式 | 背景 | 品質 |
|---|---|---|
| PNG | 透過 | — |
| JPEG | 白 (`#ffffff`) で塗りつぶし | 0.92 |

Excel 挿入時は §4.4 のとおり基準サイズで表示するため、見た目の大きさは倍率によらず一定で、解像度だけが上がる。

### 4.7 スプリッタ

- ドラッグ（Pointer Events）で入力欄の高さを変更する。上下矢印キーでも 10px 単位で変更できる（`role="separator"`、フォーカス可能）。
- 入力欄の高さは最小 40px。プレビューが 120px を下回らない範囲までしか広げない。
- 変更した高さは保存しない（タスクペインを開き直すと初期値に戻る）。

---

## 5. データ仕様

本アドインは独自の保存領域を持たず、**図形の代替テキストを唯一の永続データ**とする。

| 保存先 | 内容 | 用途 |
|---|---|---|
| 図形の `altTextTitle` | 固定文字列 `Mermaid Diagram` | Mermaid 図形の識別（§4.3 の一覧の抽出条件） |
| 図形の `altTextDescription` | Mermaid ソースコード | シート図形選択時に入力欄へ復元 |

利用者が代替テキストのタイトルを変更すると、その図形は一覧に表示されなくなる。

Cookie、localStorage、外部送信は使用しない。

---

## 6. ビルド・開発環境

### 6.1 npm スクリプト

| コマンド | 実体 | 内容 |
|---|---|---|
| `npm run build` | `node build.js` | 本番ビルド。マニフェストの URL を本番 URL に置換して `dist/` に出力 |
| `npm run build:dev` | `node build.js dev` | 開発ビルド。マニフェストの URL は `https://localhost:3000` |
| `npm run watch` | `node build.js watch` | ファイル監視ビルド（インラインソースマップ、開発 URL） |
| `npm run dev-server` | `node build.js serve` | esbuild の HTTP サーバを HTTPS プロキシ（ポート 3000）で公開 |
| `npm start` | `office-addin-debugging start manifest.xml` | Office へのサイドロード＋デバッグ起動 |
| `npm stop` | `office-addin-debugging stop manifest.xml` | デバッグ停止 |

### 6.2 build.js の処理

1. `assets/icon-{16,32,64,80}.png` が無ければ 1×1 のプレースホルダー PNG を生成する。
2. `taskpane.html` と `assets/` 配下を `dist/` にコピーする。
3. `manifest.xml` 内の `https://localhost:3000` をモードに応じた URL に置換して `dist/manifest.xml` に出力する。

   | モード | ベース URL |
   |---|---|
   | `build` | `https://msmsrep.github.io/MermaidForExcel/dist` |
   | `dev` / `watch` / `serve` | `https://localhost:3000` |

4. esbuild で `src/taskpane/taskpane.ts` を `dist/taskpane.js` にバンドルする（IIFE、ターゲット ES2017、ソースマップあり）。
5. `serve` では `office-addin-dev-certs` の証明書で HTTPS サーバを起動する。証明書が無い場合はエラーメッセージを出して終了する。

### 6.3 配布

- 本番資産は GitHub Pages（`https://msmsrep.github.io/MermaidForExcel/dist/`）でホストする。`dist/` はリポジトリにコミットする。
- コミット前は `npm run build`（本番 URL）で `dist/` を生成すること。

---

## 7. インストーラ（Inno Setup）

`Installer/MermaidForExcel.iss` により、共有フォルダカタログ方式でマニフェストを登録する。

| 項目 | 内容 |
|---|---|
| 出力 | `MermaidForExcelSetup.exe` |
| 実行権限 | ログオン中のユーザー（`PrivilegesRequired=lowest`）。SMB 共有の作成・削除のみ UAC で昇格 |
| 対応アーキテクチャ | x64 互換 |
| 言語 | 日本語 |
| 発行元 / URL | `msmsrep` / `https://github.com/msmsrep/MermaidForExcel` |
| 配置ファイル | `dist/manifest.xml` のみ |
| 配置先 | `%USERPROFILE%\Documents\OfficeAddIns\Excel\MermaidForExcel`（`{userdocs}` は OneDrive にリダイレクトされるため使わない） |

**インストール時**:

1. マニフェストを配置先にコピーする。
2. 共有名 `MermaidForExcel` が未作成なら、`net share MermaidForExcel="<配置先>"` を UAC 昇格して実行する。実行後に共有が存在しなければエラーダイアログを表示する。
3. `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{0e1e7db9-24b4-4064-8d0c-995f6f80a03f}` に以下を登録する。

   | 値 | 型 | データ |
   |---|---|---|
   | `Id` | 文字列 | `{0e1e7db9-24b4-4064-8d0c-995f6f80a03f}` |
   | `Url` | 文字列 | `\\<コンピュータ名>\MermaidForExcel` |
   | `Flags` | DWORD | `1`（メニューに表示） |

**アンインストール時**: 共有 `MermaidForExcel` が存在すれば UAC 昇格して削除し、上記レジストリキーを削除する。

インストール後、利用者は Excel の [アドイン] → [共有フォルダー] からアドインを追加する。

---

## 8. 既知の制約・課題

1. **SVG 挿入の一時的な位置**: SVG は Excel の選択範囲に挿入された後で目的の位置へ移動する。
2. **同名の既存共有**: 共有 `MermaidForExcel` が別のフォルダを指して既に存在する場合、インストーラは作り直さない。
3. **型チェック**: `npx tsc` で `node_modules`（type-fest）の型定義エラーが 3 件出る。`src` 側のエラーではない（`skipLibCheck` 未設定のため）。
