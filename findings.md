# findings.md

ffmpeg-wrapper-gui の実現手段、共通指針の適用条件と調査知見の正本。作業前に全文読む。利用者向けの導入・使用方法はREADME、更新の手順はhow-to-update.md、実施結果と未検証事項はverification.mdを参照する。移した記述の確認日・条件・留保は元のまま保持し、文書移動を再実測と解釈しない。

## 共通指針の適用条件と代替検証

- 対象の共通ルール: 設計と実装方針／検証／生成物とビルド成果物
- 理由: Windows向けPython GUIをPyInstallerで配布するため、TypeScript/Bunの一律コマンドは適用しない。
- 適用範囲: Python実装とexe配布物
- 同等の品質を保つ手段・検証: 既存のmypy・Python構文確認とspecによるビルドを行う。操作確認と未実施範囲はverification.mdへ記録する。

新たな免除や操作許可を追加するものではない。AGENTS.mdの禁止・承認条件と共通指針は引き続き有効。

<a id="readme-1"></a>

## 配布版の作成

PyInstaller の spec ファイルを使って exe を作成します。

```powershell
pyinstaller .\ffmpeg-wrapper-gui.spec --noconfirm --clean
```

生成物は `dist\ffmpeg-wrapper-gui.exe` です。

移動元: `README.md`。

<a id="readme-2"></a>

## 開発メモ

- メイン実装: [`ffmpeg_wrapper_gui.pyw`](./ffmpeg_wrapper_gui.pyw)
- PyInstaller 設定: [`ffmpeg-wrapper-gui.spec`](./ffmpeg-wrapper-gui.spec)

移動元: `README.md`。

<a id="readme-3"></a>

## 依存更新の自動処理

Dependabot は対象の依存関係を毎週確認します。patch／minor 更新は PR のチェック（CI）が成功した後に自動で squash merge されます。CI の失敗ジョブは 1 回だけ再実行します。再失敗した PR は残して手動で修正します。major 更新は手動で確認します。

移動元: `README.md`。

<a id="update-1"></a>

## Dependabot PR の更新

前提は `.github/dependabot.yml` と PR 用 CI（CI）です。更新 PR の head SHA と `gh pr checks <PR番号>` の結果を確認してください。patch／minor は全チェック成功後に自動取り込みされます。初回 CI 失敗は failed jobs のみを 1 回再実行し、再失敗した PR は残して手動で修正します。

設定を変えたときは `actionlint .github/workflows/dependabot-automation.yml` と実際の PR の Actions 結果を確認します。問題があれば呼び出し先の共通 workflow SHA を直前の検証済み値へ戻すコミットを push します。取り込まれた依存更新に問題があれば通常の revert コミットで復旧します。

CI 完了より Dependabot の分類が遅れる場合は、`callback_workflow_file` が指す呼び出し側 workflow を `workflow_dispatch` し、同じ PR 番号・head SHA・全チェックを再確認する。呼び出し側のファイル名を変える際はこの入力も一緒に更新する。

移動元: `how-to-update.md`。

<a id="verification-1"></a>

## Dependabot 自動処理（2026-09-23）

`.github/workflows/dependabot-automation.yml` を actionlint で検査し、PR 用 workflow 名（CI）と一致することを確認する。Dependabot の patch／minor かつ全 PR チェック成功の場合だけ取り込み、major・古い SHA・再失敗は残す。

実際の Dependabot PR がまだない場合、動作経路は未検証として扱う。実 PR 発生後に自動化ジョブ、CI の再試行、マージ結果を確認する。

大量の Dependabot PR により CI 完了より分類が遅れる場合でも、分類後の `workflow_dispatch` が現在の PR 番号と head SHA を照合して再評価する。別の作成者、古い SHA、未完了の CI はマージしない。

移動元: `verification.md`。

## 関連文書

- [利用者向けREADME](README.md)
- [更新と復旧](how-to-update.md)
- [検証結果と未検証範囲](verification.md)
- [変更履歴](CHANGELOG.md)
