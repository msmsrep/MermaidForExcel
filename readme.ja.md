# Text-to-SVG Diagrams for Excel

[English](./readme.md) | [日本語](./readme.ja.md)


Mermaid記法でダイアグラムを作成し、Excelのシートに画像として挿入するOfficeアドインです。

テキストベースの記法でフローチャート・シーケンス図・ER図などを素早く描き、PNG/JPEG/SVG画像として現在のシートに貼り付けられます。
ローカル完結処理のため、入力内容が外部に送信されることはありません。

紹介ページと導入方法: https://msmsrep.github.io/TextToSvgDiagramsForExcel/

Mermaid プロジェクトとは関係のない非公式のツールです。旧名は「Mermaid for Excel」です。旧名の版からの移行は [docs/migration.md](./docs/migration.md) を参照してください。

## 主な機能

- **リアルタイムプレビュー** — Mermaid記法を入力すると、タスクペイン内で自動的にダイアグラムが描画されます
- **Excelへの挿入** — プレビューした画像をワンクリックでアクティブセルの位置に挿入します
- **再編集** — 挿入した図の Mermaid コードは代替テキストに保存され、[Sheet diagram] から選んで編集・差し替えできます
- **PNG/JPEG/SVG出力対応** — 保存形式を選択してダウンロードすることもできます
- **編集できる図形（試験的）** — フローチャートを Excel の図形とコネクタで挿入する `Shapes (editable)` と、Excel の「図形に変換」向けに整えた SVG を挿入する `SVG (for shape conversion)` を選べます。アドインを持たない人も図を編集できます
- **多様なダイアグラム種別** — フローチャート、シーケンス図、クラス図、ER図、ガントチャートなど[Mermaid](https://github.com/mermaid-js/mermaid)がサポートするすべての記法が利用可能です
- **プライバシー保護** — すべての処理はタスクペイン内で完結し、外部サーバーへのデータ送信はありません

## 構成

```
TextToSvgDiagramsForExcel/
  build.js                 ← esbuild ビルドスクリプト（build / dev / watch / serve）
  manifest.xml             ← Office アドインマニフェスト
  package.json
  tsconfig.json
  assets/                  ← アイコン置き場
  scripts/sync-version.js  ← package.json のバージョンをマニフェストとインストーラに反映
  site/index.html          ← GitHub Pages のトップに置く紹介ページ（導入方法を含む）
  src/taskpane/
    taskpane.html / .css
    taskpane.ts            ← UI の結線（エントリポイント）
    render.ts              ← Mermaid のレンダリング
    image.ts               ← SVG → PNG/JPEG 変換
    officeSvg.ts           ← Excel の「図形に変換」向けの SVG 変換
    flowchartModel.ts      ← フローチャートの SVG → Excel 図形モデル
    svgUtil.ts             ← 行列・色の共通処理
    excelShapes.ts         ← Excel 図形の挿入・差し替え・一覧取得
    splitter.ts            ← 入力欄とプレビューのスプリッタ
  test/                    ← Vitest のテスト（Office.js・mermaid・Canvas はモック）
  dist/                    ← ビルド成果物（コミットしない。GitHub Actions でデプロイ）
```

## セットアップ

```bash
cd TextToSvgDiagramsForExcel
npm install

# 開発用 HTTPS 証明書のインストール（初回のみ・管理者権限が必要）
npx office-addin-dev-certs install
```

## 開発サーバの起動

```bash
npm run dev-server
# → https://localhost:3000/taskpane.html
```

サイドロードしてデバッグする場合は、開発サーバを起動したまま `npm start` を実行します。

## Excel へのアドイン登録

1. Excel を起動
2. **[挿入]** → **[アドイン]** → **[個人用アドイン]** → **[アドインをアップロード]**
3. `dist/manifest.xml` を選択（開発サーバを使う場合は `npm run build:dev` 後のもの）

ホームタブに **[Open Mermaid]** ボタンが追加されます。

## スクリプト

| コマンド | 説明 |
|---|---|
| `npm run dev-server` | ビルド + HTTPS サーバ起動（ポート 3000） |
| `npm start` | Excel にサイドロードしてデバッグ起動 |
| `npm stop` | デバッグ終了 |
| `npm run build` | `dist/` への一回限りのビルド（本番 URL のマニフェスト） |
| `npm run build:dev` | `dist/` への一回限りのビルド（`https://localhost:3000` のマニフェスト） |
| `npm run watch` | ファイル変更を監視して自動リビルド |
| `npm run typecheck` | `tsc` でソースとテストを型チェック |
| `npm test` | Vitest でテストを実行 |
| `npm run test:watch` | テストを監視モードで実行 |
| `npm version <patch\|minor\|major>` | バージョンを上げて `manifest.xml` とインストーラに反映 |

## アイコンの差し替え

`assets/icon-{16,32,64,80}.png` を任意の PNG 画像に置き換えてください。 


## ライセンス

このプロジェクトは [MIT License](LICENSE.txt)のもとで公開しています。

### サードパーティライセンス

このソフトウェアは [mermaid](https://github.com/mermaid-js/mermaid)（MIT License）を使用しています。

## プライバシーポリシー

最終更新日：2026年4月8日

### データの収集について

Text-to-SVG Diagrams for Excelは、ユーザーの個人情報およびデータを**一切収集しません**。

### 処理の仕組み

- ユーザーが入力したMermaid記法のテキストは、エクセルの**タスクペイン内のみで処理**されます
- ダイアグラムのレンダリングはローカルで完結し、外部サーバーへのデータ送信は行いません
- Excelシートへの画像挿入も、Microsoft Office JavaScript API を通じてローカルで行われます

### 外部サービスへのアクセス

本アドインは、アドイン本体（HTML/JavaScript）をホスティングサーバーから読み込みます。  
この通信にはユーザーのデータは含まれません。

### Cookieおよびトラッキング

本アドインはCookie、ローカルストレージ、トラッキング技術を使用しません。

## お問い合わせ

プライバシー・その他ご質問は、[GitHub Issues](https://github.com/msmsrep/TextToSvgDiagramsForExcel/issues) までお寄せください。
