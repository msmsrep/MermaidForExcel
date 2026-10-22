# 旧名（Mermaid for Excel）からの移行

アドインの名前を **Mermaid for Excel** から **Text-to-SVG Diagrams for Excel** に変えた。あわせてリポジトリ名と本番の URL も変わる。

| 項目 | 旧 | 新 |
|---|---|---|
| リポジトリ | `msmsrep/MermaidForExcel` | `msmsrep/TextToSvgDiagramsForExcel` |
| 紹介ページ | `https://msmsrep.github.io/MermaidForExcel/` | `https://msmsrep.github.io/TextToSvgDiagramsForExcel/` |
| タスクペイン | `…/MermaidForExcel/dist/taskpane.html` | `…/TextToSvgDiagramsForExcel/dist/taskpane.html` |
| インストーラ | `MermaidForExcelSetup.exe` | `TextToSvgDiagramsForExcelSetup.exe` |
| 配置先・共有名 | `MermaidForExcel` | `TextToSvgDiagramsForExcel` |

次の値は変えない。旧版と同じアドインとして扱われ、上書きで更新される。

- マニフェストの `<Id>`（`4f7e8c1b-…`）
- インストーラの `AppId`（`b03396e2-…`）
- 信頼できるカタログのレジストリ GUID（`0e1e7db9-…`）
- マニフェスト内部の ID（`MermaidGroup`、`MermaidPane`）。画面には表示されない

## 作者の手順

GitHub Pages の URL は、リポジトリ名を変えても旧 URL から転送されない。旧 URL を参照している manifest.xml が動かなくなるため、旧 URL に転送用のページを置く（手順 6）。手順 3〜6 の間は旧 URL が 404 になるので、続けて行う。

1. バージョンを上げる。マニフェストの `<Version>` が変わると、Office がマニフェストを読み直す。

   ```bash
   npm version minor
   ```

2. 変更を `main` にマージする（まだ push しない）。
3. GitHub でリポジトリ名を変える: Settings → General → Repository name を `TextToSvgDiagramsForExcel` にする。
4. ローカルの remote を新しい URL に変える。

   ```bash
   git remote set-url origin https://github.com/msmsrep/TextToSvgDiagramsForExcel.git
   ```

5. `main` を push する。GitHub Actions が新しい URL にデプロイしたら、`https://msmsrep.github.io/TextToSvgDiagramsForExcel/` と `…/dist/manifest.xml` が開けることを確認する。
6. 旧 URL 用のリポジトリを作る。
   1. GitHub で新しい公開リポジトリ `msmsrep/MermaidForExcel` を作る。
   2. このリポジトリの [legacy-pages/](../legacy-pages/) の中身をルートに置く。
   3. 新しいビルド結果から `dist/manifest.xml` と `dist/assets/` を、旧 URL 用リポジトリの `dist/` にコピーする。旧 manifest.xml のリボンのアイコンと、旧紹介ページから manifest.xml をダウンロードするリンクに使われる。
   4. Settings → Pages で、`main` ブランチのルートを公開する。
   5. `https://msmsrep.github.io/MermaidForExcel/dist/taskpane.html` が新しいタスクペインに転送されることを確認する。

   旧名のリポジトリを作ると、GitHub による旧リポジトリ名から新リポジトリ名への自動転送（`git clone` や Issue のリンク）は効かなくなる。

7. `npm run build` のあと、Inno Setup で `Installer/TextToSvgDiagramsForExcel.iss` をコンパイルし、`TextToSvgDiagramsForExcelSetup.exe` を配布する。
8. ローカルのフォルダ名（`MermaidForExcel`）は変えなくても動く。変える場合は Claude Code などの作業を終えてから行う。

## 利用者の手順

### manifest.xml をアップロードしている場合（方法 1）

旧 manifest.xml のままでも、転送ページによって引き続き使える。ただし、なるべく次の手順で登録し直す。

1. [紹介ページ](https://msmsrep.github.io/TextToSvgDiagramsForExcel/#install) から新しい manifest.xml をダウンロードする。
2. Excel の [個人用アドイン] で、旧アドインを削除してから新しい manifest.xml をアップロードする。

### 共有フォルダーカタログを自分で設定している場合（方法 2）

1. 共有フォルダーの manifest.xml を新しいものに置き換える。
2. Excel を再起動する。マニフェストの `Id` が同じなので、同じアドインとして更新される。

### インストーラで導入している場合

1. 新しい `TextToSvgDiagramsForExcelSetup.exe` を実行する。旧版のアンインストールは不要。
   - 管理者権限の確認が 1 回表示される。旧名の共有 `MermaidForExcel` の削除と、新しい共有 `TextToSvgDiagramsForExcel` の作成を行う。
   - 旧名の配置先（`Documents\OfficeAddIns\Excel\MermaidForExcel`）は自動で削除される。
2. Excel を再起動する。

### 名前が古いまま表示される場合

Excel を終了し、Office のアドインのキャッシュ（`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`）の中身を削除してから、[アドイン] → [共有フォルダー]（または [個人用アドイン]）から追加し直す。

## 転送ページを残す期間

旧 URL 用のリポジトリは、旧 manifest.xml の利用者が移行し終えるまで残す（目安として 1 年）。削除すると、旧 manifest.xml のアドインは開けなくなる。
