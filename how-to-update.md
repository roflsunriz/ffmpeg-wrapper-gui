1. mypyでエラーがないことを確認
2. pythonで構文エラーがないことを確認
3. コミット＆プッシュ
4. pyinstaller ffmpeg-wrapper-gui.specでバイナリ更新
5. git tag vX.X.X & git push origin main vX.X.Xでリリース作成

## Dependabot PR の更新



Dependabot version update は毎月28日00:00、検証済み PR の自動マージは同日18:00（いずれも `Asia/Tokyo`）にまとめます。28日より前に自動マージせず、現在の head SHA と全 PR チェックを月末処理で再確認します。
詳細は [findings.md](findings.md#update-1) に移しました。
