# ffmpeg Wrapper GUI

`ffmpeg` を GUI から操作する Windows 向けの軽量ツールです。音声変換、動画から音声抽出、動画圧縮をまとめて扱えます。

## 機能

- ファイルの追加、削除、複数ファイルの一括処理
- ドラッグ & ドロップ入力
- 音声変換
- 動画から音声抽出
- 動画圧縮
- 出力先、出力名ポリシー、速度、ビットレート、解像度、FPS、CRF の指定
- 進捗表示とログ表示
- 設定の保存と再読み込み

## 必要なもの

- Windows
- `ffmpeg`
- `ffprobe`

アプリは起動時に以下の順で `ffmpeg` / `ffprobe` を探します。

1. 実行ファイルと同じフォルダ
2. `resources` フォルダ
3. `PATH`

つまり、配布版をそのまま使う場合は、`ffmpeg.exe` と `ffprobe.exe` を exe と同じフォルダに置くか、`PATH` に通してください。

### 依存関係の導入

`setup-scripts/install-dependencies.ps1` を使うと、`winget` 経由で FFmpeg をまとめて導入できます。

```powershell
iex "& { $(iwr -useb 'https://raw.githubusercontent.com/roflsunriz/ffmpeg-wrapper-gui/main/setup-scripts/install-dependencies.ps1') }"
```

このスクリプトは `Gyan.FFMpeg` をインストールします。`winget` が使える環境で実行してください。

## 使い方

1. 入力ファイルを追加します。
2. 変換モードを選びます。
3. 出力先と各種設定を調整します。
4. `実行` を押します。

`設定保存` を押すと、次回起動時に前回の設定を復元します。

## 配布版の作成

PyInstaller の spec ファイルを使って exe を作成します。

```powershell
pyinstaller .\ffmpeg-wrapper-gui.spec --noconfirm --clean
```

生成物は `dist\ffmpeg-wrapper-gui.exe` です。

## 開発メモ

- メイン実装: [`ffmpeg_wrapper_gui.pyw`](./ffmpeg_wrapper_gui.pyw)
- PyInstaller 設定: [`ffmpeg-wrapper-gui.spec`](./ffmpeg-wrapper-gui.spec)
