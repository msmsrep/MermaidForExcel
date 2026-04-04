# Mermaid for Excel

Mermaid 記法でダイアグラムを作成し、Excel のシートに画像として挿入する Office アドイン。  
ビルドツールは **esbuild** のみを使用したミニマル構成。

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
3. `manifest.xml` を選択

ホームタブに **[Mermaid を開く]** ボタンが追加されます。

## スクリプト

| コマンド | 説明 |
|---|---|
| `npm start` | ビルド + HTTPS サーバ起動（ポート 3000） |
| `npm run build` | `dist/` への一回限りのビルド |
| `npm run watch` | ファイル変更を監視して自動リビルド |

## アイコンの差し替え

`assets/icon-{16,32,80}.png` を任意の PNG 画像に置き換えてください。  
初回ビルド時に 1×1px のプレースホルダーが自動生成されます。
