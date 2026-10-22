# 旧名（Mermaid for Excel）からの移行

アドインの名前を **Mermaid for Excel** から **Text-to-SVG Diagrams for Excel** に変えた。あわせてリポジトリ名と本番の URL も変わった。

| 項目 | 旧 | 新 |
|---|---|---|
| リポジトリ | `msmsrep/MermaidForExcel` | `msmsrep/TextToSvgDiagramsForExcel` |
| 紹介ページ | `https://msmsrep.github.io/MermaidForExcel/` | `https://msmsrep.github.io/TextToSvgDiagramsForExcel/` |
| タスクペイン | `…/MermaidForExcel/dist/taskpane.html` | `…/TextToSvgDiagramsForExcel/dist/taskpane.html` |
| インストーラ | `MermaidForExcelSetup.exe` | `TextToSvgDiagramsForExcelSetup.exe` |
| 配置先・共有名 | `MermaidForExcel` | `TextToSvgDiagramsForExcel` |

旧 URL は転送しない。旧 manifest.xml のアドインは開けなくなるため、新しい manifest.xml を登録し直す必要がある。

次の値は変えていない。旧版と同じアドインとして扱われ、上書きで更新される。

- マニフェストの `<Id>`（`4f7e8c1b-…`）
- インストーラの `AppId`（`b03396e2-…`）
- 信頼できるカタログのレジストリ GUID（`0e1e7db9-…`）
- マニフェスト内部の ID（`MermaidGroup`、`MermaidPane`）。画面には表示されない

## 利用者の手順

### manifest.xml をアップロードしている場合（方法 1）

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
