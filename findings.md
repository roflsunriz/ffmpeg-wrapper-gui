# 調査知見

- 配布exeは個別オプションの再構成ではなく `ffmpeg-wrapper-gui.spec` を正本にPyInstallerで生成する。ビルド成功とGUI操作・ffmpeg実行の成功は分けて確認する。
- Dependabotの分類がCI完了より遅れた場合は、`callback_workflow_file` の呼出し側を `workflow_dispatch` して同じPR番号・head SHA・全チェックを再照合する。workflow名変更時はこの入力も更新する。
