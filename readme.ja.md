# Mermaid for Excel

<a href="./README.md">English</a> | <a href="./README.ja.md">日本語</a>

Mermaid記法でダイアグラムを作成し、Excelのシートに画像として挿入するOfficeアドインです。

テキストベースの記法でフローチャート・シーケンス図・ER図などを素早く描き、PNG/JPEG画像として現在のシートに貼り付けられます。
ローカル完結処理のため、入力内容が外部に送信されることはありません。

## 主な機能

- **リアルタイムプレビュー** — Mermaid記法を入力してレンダリングボタンを押すと、タスクペイン内でダイアグラムを即座に確認できます
- **Excelへの挿入** — プレビューした画像をワンクリックでアクティブシートに挿入します
- **PNG/JPEG出力対応** — 保存形式を選択してダウンロードすることもできます
- **多様なダイアグラム種別** — フローチャート、シーケンス図、クラス図、ER図、ガントチャートなど[Mermaid](https://github.com/mermaid-js/mermaid)がサポートするすべての記法が利用可能です
- **プライバシー保護** — すべての処理はタスクペイン内で完結し、外部サーバーへのデータ送信はありません

## 構成

```
MermaidForExcel/
  build.js                 ← esbuild ビルドスクリプト（serve / watch / build）
  manifest.xml             ← Office アドインマニフェスト
  package.json
  tsconfig.json
  assets/                  ← アイコン置き場（初回ビルド時にプレースホルダーを自動生成）
  src/taskpane/
    taskpane.html
    taskpane.ts
  dist/                    ← ビルド成果物（git 管理外）
```

## セットアップ

```bash
cd MermaidForExcel
npm install

# 開発用 HTTPS 証明書のインストール（初回のみ・管理者権限が必要）
npx office-addin-dev-certs install
```

## 開発サーバの起動

```bash
npm start
# → https://localhost:3000/taskpane.html
```

## Excel へのアドイン登録

1. Excel を起動
2. **[挿入]** → **[アドイン]** → **[個人用アドイン]** → **[アドインをアップロード]**
3. `dist/manifest.xml` を選択

ホームタブに **[Open Mermaid]** ボタンが追加されます。

## スクリプト

| コマンド | 説明 |
|---|---|
| `npm start` | ビルド + HTTPS サーバ起動（ポート 3000） |
| `npm run build` | `dist/` への一回限りのビルド |
| `npm run watch` | ファイル変更を監視して自動リビルド |

## アイコンの差し替え

`assets/icon-{16,32,80}.png` を任意の PNG 画像に置き換えてください。  
初回ビルド時に 1×1px のプレースホルダーが自動生成されます。


## License

This project is licensed under the [MIT License](LICENSE.txt).

### Third-party licenses

This software uses [mermaid](https://github.com/mermaid-js/mermaid) (MIT License).

## プライバシーポリシー

最終更新日：2026年4月6日

### データの収集について

Mermaid for Excelは、ユーザーの個人情報およびデータを**一切収集しません**。

### 処理の仕組み

- ユーザーが入力したMermaid記法のテキストは、エクセルの**タスクペイン内のみで処理**されます
- ダイアグラムのレンダリングはローカルで完結し、外部サーバーへのデータ送信は行いません
- Excelシートへの画像挿入も、Microsoft Office JavaScript API を通じてローカルで行われます

### 外部サービスへのアクセス

本アドインは、アドイン本体（HTML/JavaScript）をホスティングサーバーから読み込みます。  
この通信にはユーザーのデータは含まれません。

### Cookieおよびトラッキング

本アドインはCookie、ローカルストレージ、トラッキング技術を使用しません。

### お問い合わせ

プライバシーに関するご質問は、[GitHub Issues](https://github.com/msmsrep/MermaidForExcel/issues) までお寄せください。
