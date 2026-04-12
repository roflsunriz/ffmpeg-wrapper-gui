from __future__ import annotations

import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from tkinter import END, HORIZONTAL, BOTH, LEFT, RIGHT, X, filedialog, messagebox, StringVar, Tk, Toplevel
from tkinter import ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore

    HAS_DND = True
except Exception:
    HAS_DND = False
    DND_FILES = None
    TkinterDnD = None


SETTINGS_PATH = Path(__file__).resolve().parent / "ffmpeg-wrapper-settings.json"

AUDIO_FORMATS = ("mp3", "m4a", "flac", "wav", "ogg", "opus")
AUDIO_CODECS = ("aac", "libmp3lame", "libopus", "pcm_s16le")
AUDIO_BITRATES = ("64k", "96k", "128k", "192k", "256k")
VIDEO_FORMATS = ("mp4", "mkv", "mov", "webm")
VIDEO_CODECS = ("h264", "h265", "mpeg4", "vp9")
VIDEO_AUDIO_CODECS = ("aac", "libmp3lame", "libopus", "vorbis")
VIDEO_CODEC_LABELS = {
    "h264": "H.264",
    "h265": "H.265 / HEVC",
    "mpeg4": "MPEG-4 Part 2",
    "vp9": "VP9",
}
VIDEO_ENCODER_OPTIONS: dict[str, tuple[str, ...]] = {
    "h264": ("libx264", "h264_nvenc", "h264_qsv", "h264_amf"),
    "h265": ("libx265", "hevc_nvenc", "hevc_qsv", "hevc_amf"),
    "mpeg4": ("mpeg4",),
    "vp9": ("libvpx-vp9",),
}
VIDEO_ENCODER_DESCRIPTIONS = {
    "libx264": "ソフトウェアエンコードです。処理は比較的遅めですが、H.264では品質と互換性のバランスが良好です。",
    "h264_nvenc": "NVIDIA GPU のハードウェアアクセラレーションを利用できます。高速ですが、同等ビットレートでは libx264 より画質が落ちやすいです。",
    "h264_qsv": "Intel Quick Sync Video を利用できます。高速で消費電力も抑えやすい一方、画質傾向は libx264 よりやや不利です。",
    "h264_amf": "AMD GPU のハードウェアアクセラレーションを利用できます。高速ですが、画質調整の自由度はソフトウェアエンコードより低めです。",
    "libx265": "ソフトウェアエンコードです。高圧縮ですが処理負荷が高く、速度は遅めです。容量重視の HEVC 向けです。",
    "hevc_nvenc": "NVIDIA GPU の HEVC ハードウェアエンコードです。高速ですが、同条件では libx265 より細部が甘くなりやすいです。",
    "hevc_qsv": "Intel Quick Sync Video の HEVC エンコードです。速度重視ですが、画質は libx265 より控えめです。",
    "hevc_amf": "AMD GPU の HEVC ハードウェアエンコードです。高速ですが、圧縮効率と画質は libx265 に及ばないことがあります。",
    "mpeg4": "旧式のソフトウェアエンコードです。互換性重視向けですが、圧縮効率は H.264/H.265 よりかなり劣ります。",
    "libvpx-vp9": "VP9 のソフトウェアエンコードです。圧縮効率は高い一方、処理時間は長くなりやすいです。",
}
VIDEO_FORMAT_COMPATIBILITY: dict[str, dict[str, tuple[str, ...]]] = {
    "mp4": {
        "video_codecs": ("h264", "h265", "mpeg4"),
        "audio_codecs": ("aac", "libmp3lame"),
    },
    "mkv": {
        "video_codecs": VIDEO_CODECS,
        "audio_codecs": VIDEO_AUDIO_CODECS,
    },
    "mov": {
        "video_codecs": ("h264", "h265", "mpeg4"),
        "audio_codecs": ("aac", "libmp3lame"),
    },
    "webm": {
        "video_codecs": ("vp9",),
        "audio_codecs": ("libopus", "vorbis"),
    },
}
MODE_OPTIONS = ("audio_convert", "video_extract_audio", "video_compress")
NAME_POLICY_OPTIONS = ("source_name", "source_name_with_suffix", "custom_name")
PRESET_LABELS = (
    "High Quality",
    "Standard",
    "Balanced",
    "Mobile Optimized",
    "Web Friendly",
    "Economy",
    "Minimum",
)

PRESETS: dict[str, dict[str, str]] = {
    "High Quality": {"scale": "1920:1080", "fps": "30", "crf": "18", "audio_bitrate": "192k"},
    "Standard": {"scale": "1366:630", "fps": "30", "crf": "22", "audio_bitrate": "96k"},
    "Balanced": {"scale": "1366:630", "fps": "30", "crf": "20", "audio_bitrate": "96k"},
    "Mobile Optimized": {"scale": "854:480", "fps": "24", "crf": "24", "audio_bitrate": "96k"},
    "Web Friendly": {"scale": "1280:720", "fps": "30", "crf": "22", "audio_bitrate": "128k"},
    "Economy": {"scale": "1024:472", "fps": "24", "crf": "24", "audio_bitrate": "96k"},
    "Minimum": {"scale": "854:394", "fps": "24", "crf": "26", "audio_bitrate": "64k"},
}

WINDOWS_CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


@dataclass
class ProgressEvent:
    file_path: Path
    percent: float
    text: str


class OverwriteResolver:
    def __init__(self, root: Tk) -> None:
        self.root = root

    def ask(self, output_path: Path) -> str:
        choice = {"value": "skip"}
        done = threading.Event()

        def _show() -> None:
            dialog = Toplevel(self.root)
            dialog.title("出力ファイルの確認")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()

            frame = ttk.Frame(dialog, padding=12)
            frame.pack(fill=BOTH, expand=True)

            ttk.Label(
                frame,
                text=f"出力ファイルが既に存在します:\n{output_path}\n\nどうしますか？",
                justify=LEFT,
            ).pack(fill=X, pady=(0, 12))

            def set_choice(value: str) -> None:
                choice["value"] = value
                dialog.destroy()
                done.set()

            btn_frame = ttk.Frame(frame)
            btn_frame.pack(fill=X)
            ttk.Button(btn_frame, text="上書き", command=lambda: set_choice("overwrite")).pack(side=LEFT, padx=4)
            ttk.Button(btn_frame, text="スキップ", command=lambda: set_choice("skip")).pack(side=LEFT, padx=4)
            ttk.Button(btn_frame, text="連番リネーム", command=lambda: set_choice("rename")).pack(side=LEFT, padx=4)

            def close_as_skip() -> None:
                set_choice("skip")

            dialog.protocol("WM_DELETE_WINDOW", close_as_skip)

        self.root.after(0, _show)
        done.wait()
        return choice["value"]


class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("ffmpeg Wrapper GUI")
        self.root.geometry("980x760")
        self.root.minsize(980, 760)

        self.file_list: list[Path] = []
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.progress_queue: queue.Queue[ProgressEvent] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.pause_event = threading.Event()
        self.current_process: subprocess.Popen[str] | None = None
        self.remaining_queue: list[Path] = []
        self._tool_resolution_logged: set[str] = set()
        self.available_ffmpeg_encoders: set[str] = set()

        self.mode_var = StringVar(value="audio_convert")
        self.output_dir_var = StringVar(value="")
        self.name_policy_var = StringVar(value="source_name")
        self.custom_name_var = StringVar(value="")
        self.playback_speed_var = StringVar(value="1.0")
        self.audio_format_var = StringVar(value="m4a")
        self.audio_codec_var = StringVar(value="aac")
        self.audio_bitrate_var = StringVar(value="128k")
        self.sample_rate_var = StringVar(value="")
        self.preset_var = StringVar(value="Standard")
        self.scale_var = StringVar(value=PRESETS["Standard"]["scale"])
        self.fps_var = StringVar(value=PRESETS["Standard"]["fps"])
        self.crf_var = StringVar(value=PRESETS["Standard"]["crf"])
        self.video_format_var = StringVar(value="mp4")
        self.video_codec_var = StringVar(value="h264")
        self.video_encoder_var = StringVar(value="libx264")
        self.video_audio_codec_var = StringVar(value="aac")
        self.video_audio_bitrate_var = StringVar(value=PRESETS["Standard"]["audio_bitrate"])
        self.video_encoder_description_var = StringVar(value="")
        self.dnd_status_var = StringVar(value="D&D: 初期化中")
        self.progress_text_var = StringVar(value="待機中")

        self.overwrite_resolver = OverwriteResolver(root)
        self.available_ffmpeg_encoders = self._detect_ffmpeg_encoders()
        self._build_ui()
        self._bind_events()
        self._load_settings()
        self._start_polling()

    def _build_ui(self) -> None:
        wrapper = ttk.Frame(self.root, padding=10)
        wrapper.pack(fill=BOTH, expand=True)

        file_frame = ttk.LabelFrame(wrapper, text="入力ファイル", padding=8)
        file_frame.pack(fill=BOTH, expand=False, pady=(0, 8))

        self.file_listbox = ttk.Treeview(file_frame, columns=("path",), show="headings", height=2)
        self.file_listbox.heading("path", text="Path")
        self.file_listbox.column("path", anchor="w", stretch=True, width=900)
        self.file_listbox.pack(fill=BOTH, expand=True)
        self._refresh_file_list_height()

        file_btn_frame = ttk.Frame(file_frame)
        file_btn_frame.pack(fill=X, pady=(8, 0))
        ttk.Button(file_btn_frame, text="ファイル追加", command=self._add_files).pack(side=LEFT, padx=(0, 6))
        ttk.Button(file_btn_frame, text="選択削除", command=self._remove_selected_files).pack(side=LEFT, padx=(0, 6))
        ttk.Button(file_btn_frame, text="全削除", command=self._clear_files).pack(side=LEFT)
        ttk.Label(file_btn_frame, textvariable=self.dnd_status_var).pack(side=RIGHT)

        mode_frame = ttk.LabelFrame(wrapper, text="変換モード", padding=8)
        mode_frame.pack(fill=X, pady=(0, 8))
        for mode in MODE_OPTIONS:
            ttk.Radiobutton(mode_frame, text=mode, variable=self.mode_var, value=mode, command=self._toggle_mode_sections).pack(
                side=LEFT, padx=(0, 10)
            )

        output_frame = ttk.LabelFrame(wrapper, text="出力設定", padding=8)
        output_frame.pack(fill=X, pady=(0, 8))
        ttk.Label(output_frame, text="出力ディレクトリ").grid(row=0, column=0, sticky="w")
        ttk.Entry(output_frame, textvariable=self.output_dir_var, width=70).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Button(output_frame, text="参照", command=self._pick_output_dir).grid(row=0, column=2)

        ttk.Label(output_frame, text="出力名ポリシー").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.name_policy_combo = ttk.Combobox(output_frame, state="readonly", values=NAME_POLICY_OPTIONS, textvariable=self.name_policy_var)
        self.name_policy_combo.grid(row=1, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(output_frame, text="カスタム名").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.custom_name_entry = ttk.Entry(output_frame, textvariable=self.custom_name_var, width=40)
        self.custom_name_entry.grid(row=2, column=1, sticky="w", padx=6, pady=(6, 0))
        output_frame.columnconfigure(1, weight=1)

        self.mode_settings_container = ttk.Frame(wrapper)
        self.mode_settings_container.pack(fill=X, pady=(0, 8))

        self.audio_frame = ttk.LabelFrame(self.mode_settings_container, text="音声変換設定", padding=8)
        self.audio_frame.pack(fill=X, pady=(0, 8))
        ttk.Label(self.audio_frame, text="出力形式").grid(row=0, column=0, sticky="w")
        ttk.Combobox(self.audio_frame, state="readonly", values=AUDIO_FORMATS, textvariable=self.audio_format_var, width=8).grid(
            row=0, column=1, sticky="w", padx=6
        )
        ttk.Label(self.audio_frame, text="再生速度(atempo)").grid(row=0, column=2, sticky="w")
        ttk.Entry(self.audio_frame, textvariable=self.playback_speed_var, width=10).grid(row=0, column=3, sticky="w", padx=6)
        ttk.Label(self.audio_frame, text="コーデック").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(self.audio_frame, state="readonly", values=AUDIO_CODECS, textvariable=self.audio_codec_var, width=12).grid(
            row=1, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(self.audio_frame, text="ビットレート").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Combobox(self.audio_frame, state="readonly", values=AUDIO_BITRATES, textvariable=self.audio_bitrate_var, width=10).grid(
            row=1, column=3, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(self.audio_frame, text="サンプリングレート(任意)").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self.audio_frame, textvariable=self.sample_rate_var, width=12).grid(row=2, column=1, sticky="w", padx=6, pady=(6, 0))

        self.video_frame = ttk.LabelFrame(self.mode_settings_container, text="動画圧縮設定", padding=8)
        self.video_frame.pack(fill=X, pady=(0, 8))
        ttk.Label(self.video_frame, text="プリセット").grid(row=0, column=0, sticky="w")
        self.preset_combo = ttk.Combobox(self.video_frame, state="readonly", values=PRESET_LABELS, textvariable=self.preset_var, width=16)
        self.preset_combo.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(self.video_frame, text="出力形式").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.video_format_combo = ttk.Combobox(
            self.video_frame, state="readonly", values=VIDEO_FORMATS, textvariable=self.video_format_var, width=12
        )
        self.video_format_combo.grid(
            row=1, column=1, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(self.video_frame, text="ビデオコーデック").grid(row=1, column=2, sticky="w", pady=(6, 0))
        self.video_codec_combo = ttk.Combobox(
            self.video_frame, state="readonly", values=VIDEO_CODECS, textvariable=self.video_codec_var, width=14
        )
        self.video_codec_combo.grid(
            row=1, column=3, sticky="w", padx=6, pady=(6, 0)
        )
        ttk.Label(self.video_frame, text="解像度(W:H)").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self.video_frame, textvariable=self.scale_var, width=16).grid(row=2, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(self.video_frame, text="FPS").grid(row=2, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(self.video_frame, textvariable=self.fps_var, width=10).grid(row=2, column=3, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(self.video_frame, text="CRF").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self.video_frame, textvariable=self.crf_var, width=10).grid(row=3, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(self.video_frame, text="エンコーダ").grid(row=3, column=2, sticky="w", pady=(6, 0))
        self.video_encoder_combo = ttk.Combobox(
            self.video_frame, state="readonly", values=(), textvariable=self.video_encoder_var, width=14
        )
        self.video_encoder_combo.grid(row=3, column=3, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(self.video_frame, text="オーディオコーデック").grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.video_audio_codec_combo = ttk.Combobox(
            self.video_frame, state="readonly", values=VIDEO_AUDIO_CODECS, textvariable=self.video_audio_codec_var, width=14
        )
        self.video_audio_codec_combo.grid(row=4, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(self.video_frame, text="Audio Bitrate").grid(row=4, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(self.video_frame, textvariable=self.video_audio_bitrate_var, width=12).grid(row=4, column=3, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(
            self.video_frame,
            textvariable=self.video_encoder_description_var,
            justify=LEFT,
            wraplength=720,
        ).grid(row=5, column=0, columnspan=4, sticky="w", pady=(8, 0))

        action_frame = ttk.Frame(wrapper)
        action_frame.pack(fill=X, pady=(0, 8))
        ttk.Button(action_frame, text="実行", command=self._start).pack(side=LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="停止", command=self._stop).pack(side=LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="再開", command=self._resume).pack(side=LEFT, padx=(0, 6))
        ttk.Button(action_frame, text="設定保存", command=self._save_settings).pack(side=LEFT, padx=(0, 6))

        progress_frame = ttk.LabelFrame(wrapper, text="進捗", padding=8)
        progress_frame.pack(fill=X, pady=(0, 8))
        self.progress_bar = ttk.Progressbar(progress_frame, orient=HORIZONTAL, mode="determinate")
        self.progress_bar.pack(fill=X, expand=True)
        ttk.Label(progress_frame, textvariable=self.progress_text_var).pack(anchor="w", pady=(6, 0))

        log_frame = ttk.LabelFrame(wrapper, text="ログ", padding=8)
        log_frame.pack(fill=X, expand=False)
        self.log_text = ttk.Treeview(log_frame, columns=("log",), show="headings", height=5)
        self.log_text.heading("log", text="message")
        self.log_text.column("log", anchor="w", stretch=True, width=930)
        self.log_text.pack(fill=X, expand=False)

    def _bind_events(self) -> None:
        self.preset_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_preset())
        self.name_policy_combo.bind("<<ComboboxSelected>>", lambda _e: self._toggle_mode_sections())
        self.video_format_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_video_codec_options())
        self.video_codec_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_video_encoder_options())
        self.video_encoder_combo.bind("<<ComboboxSelected>>", lambda _e: self._update_video_encoder_description())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if HAS_DND and TkinterDnD is not None and DND_FILES is not None:
            try:
                self.file_listbox.drop_target_register(DND_FILES)
                self.file_listbox.dnd_bind("<<Drop>>", self._handle_drop)
                self.dnd_status_var.set("D&D: 有効")
            except Exception:
                self.dnd_status_var.set("D&D: 初期化失敗")
        else:
            self.dnd_status_var.set("D&D: tkinterdnd2未導入")

    def _toggle_mode_sections(self) -> None:
        mode = self.mode_var.get()
        if mode == "video_compress":
            self.audio_frame.pack_forget()
            self.video_frame.pack(fill=X, pady=(0, 8))
        else:
            self.video_frame.pack_forget()
            self.audio_frame.pack(fill=X, pady=(0, 8))
        if self.name_policy_var.get() == "custom_name":
            self.custom_name_entry.configure(state="normal")
        else:
            self.custom_name_entry.configure(state="disabled")

    def _apply_preset(self) -> None:
        preset = self.preset_var.get()
        data = PRESETS.get(preset)
        if not data:
            return
        self.scale_var.set(data["scale"])
        self.fps_var.set(data["fps"])
        self.crf_var.set(data["crf"])
        self.video_audio_bitrate_var.set(data["audio_bitrate"])

    def _sync_video_codec_options(self) -> None:
        compatibility = VIDEO_FORMAT_COMPATIBILITY.get(self.video_format_var.get().strip(), VIDEO_FORMAT_COMPATIBILITY["mp4"])
        video_codecs = compatibility["video_codecs"]
        audio_codecs = compatibility["audio_codecs"]

        self.video_codec_combo.configure(values=video_codecs)
        self.video_audio_codec_combo.configure(values=audio_codecs)

        if self.video_codec_var.get().strip() not in video_codecs:
            self.video_codec_var.set(video_codecs[0])
        if self.video_audio_codec_var.get().strip() not in audio_codecs:
            self.video_audio_codec_var.set(audio_codecs[0])
        self._sync_video_encoder_options()

    def _sync_video_encoder_options(self) -> None:
        codec = self.video_codec_var.get().strip()
        known_candidates = VIDEO_ENCODER_OPTIONS.get(codec, ())
        available_candidates = tuple(encoder for encoder in known_candidates if self._is_encoder_available(encoder))
        encoder_values = available_candidates or known_candidates

        self.video_encoder_combo.configure(values=encoder_values)
        if encoder_values and self.video_encoder_var.get().strip() not in encoder_values:
            self.video_encoder_var.set(encoder_values[0])
        self._update_video_encoder_description()

    def _update_video_encoder_description(self) -> None:
        encoder = self.video_encoder_var.get().strip()
        if not encoder:
            self.video_encoder_description_var.set("")
            return
        description = VIDEO_ENCODER_DESCRIPTIONS.get(encoder, "このエンコーダの説明は未定義です。")
        if self.available_ffmpeg_encoders and encoder not in self.available_ffmpeg_encoders:
            description = f"{description} 現在の ffmpeg では利用できないため選択候補からは通常除外されます。"
        self.video_encoder_description_var.set(description)

    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="入力ファイルを選択")
        self._append_paths(paths)

    def _append_paths(self, paths: tuple[str, ...] | list[str]) -> None:
        for raw in paths:
            p = Path(raw).expanduser().resolve()
            if p not in self.file_list:
                self.file_list.append(p)
                self.file_listbox.insert("", END, values=(str(p),))
        self._refresh_file_list_height()

    def _remove_selected_files(self) -> None:
        selected = self.file_listbox.selection()
        if not selected:
            return
        paths_to_remove: set[Path] = set()
        for item in selected:
            values = self.file_listbox.item(item).get("values", [])
            if values:
                paths_to_remove.add(Path(values[0]))
        self.file_list = [f for f in self.file_list if f not in paths_to_remove]
        for item in selected:
            self.file_listbox.delete(item)
        self._refresh_file_list_height()

    def _clear_files(self) -> None:
        self.file_list = []
        for item in self.file_listbox.get_children():
            self.file_listbox.delete(item)
        self._refresh_file_list_height()

    def _refresh_file_list_height(self) -> None:
        row_count = len(self.file_list)
        target_rows = max(2, min(row_count, 10))
        self.file_listbox.configure(height=target_rows)

    def _pick_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="出力ディレクトリを選択")
        if selected:
            self.output_dir_var.set(selected)

    def _handle_drop(self, event: object) -> None:
        data = getattr(event, "data", "")
        if not isinstance(data, str) or not data:
            return
        paths = self.root.tk.splitlist(data)
        normalized = [p.strip("{}") for p in paths]
        self._append_paths(normalized)

    def _on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno("確認", "変換中です。停止して終了しますか？"):
                return
            self._stop()
        self.root.destroy()

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("実行中", "既に処理中です。")
            return
        if not self.file_list:
            messagebox.showwarning("入力不足", "入力ファイルを追加してください。")
            return
        valid, msg = self._validate_inputs()
        if not valid:
            messagebox.showerror("入力エラー", msg)
            return

        self.stop_event.clear()
        self.pause_event.clear()
        self.remaining_queue = list(self.file_list)
        self._set_progress(0.0, "処理開始")
        self.worker = threading.Thread(target=self._run_batch, args=(list(self.file_list),), daemon=True)
        self.worker.start()

    def _resume(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("実行中", "現在実行中です。")
            return
        if not self.remaining_queue:
            messagebox.showinfo("再開", "再開対象のキューはありません。")
            return
        self.stop_event.clear()
        self.pause_event.clear()
        self.worker = threading.Thread(target=self._run_batch, args=(list(self.remaining_queue),), daemon=True)
        self.worker.start()
        self._log("停止時点の残りキューから再開しました。")

    def _stop(self) -> None:
        self.stop_event.set()
        self.pause_event.set()
        proc = self.current_process
        if proc and proc.poll() is None:
            try:
                proc.terminate()
            except Exception:
                pass
        self._log("停止要求を受け付けました。")

    def _validate_inputs(self) -> tuple[bool, str]:
        try:
            speed = float(self.playback_speed_var.get())
            if speed < 0.5 or speed > 2.0:
                return False, "再生速度(atempo)は0.5〜2.0の範囲で指定してください。"
        except ValueError:
            return False, "再生速度(atempo)は数値で指定してください。"

        sample_rate = self.sample_rate_var.get().strip()
        if sample_rate and not sample_rate.isdigit():
            return False, "サンプリングレートは数字のみで指定してください。"

        if self.mode_var.get() == "video_compress":
            if not re.fullmatch(r"\d+:\d+", self.scale_var.get().strip()):
                return False, "解像度は WxH ではなく W:H 形式（例 1366:630）で指定してください。"
            if not self.fps_var.get().strip().isdigit():
                return False, "FPSは整数で指定してください。"
            if not self.crf_var.get().strip().isdigit():
                return False, "CRFは整数で指定してください。"
            compatibility = VIDEO_FORMAT_COMPATIBILITY.get(self.video_format_var.get().strip())
            if not compatibility:
                return False, "動画の出力形式が不正です。"
            if self.video_codec_var.get().strip() not in compatibility["video_codecs"]:
                return False, "選択した出力形式では、そのビデオコーデックは使用できません。"
            allowed_encoders = VIDEO_ENCODER_OPTIONS.get(self.video_codec_var.get().strip(), ())
            if self.video_encoder_var.get().strip() not in allowed_encoders:
                return False, "選択したビデオコーデックでは、そのエンコーダは使用できません。"
            if self.video_audio_codec_var.get().strip() not in compatibility["audio_codecs"]:
                return False, "選択した出力形式では、そのオーディオコーデックは使用できません。"
            if self.available_ffmpeg_encoders and self.video_encoder_var.get().strip() not in self.available_ffmpeg_encoders:
                return False, "選択したエンコーダは現在の ffmpeg では利用できません。"

        if self.name_policy_var.get() == "custom_name":
            if len(self.file_list) != 1:
                return False, "custom_name は単一入力時のみ指定できます。"
            if not self.custom_name_var.get().strip():
                return False, "custom_name が空です。"

        return True, ""

    def _run_batch(self, files: list[Path]) -> None:
        success = 0
        failed = 0
        skipped = 0
        self.remaining_queue = list(files)
        total = len(files)

        for index, input_file in enumerate(files, start=1):
            if self.stop_event.is_set():
                self._log("停止要求により処理を中断しました。")
                break

            self.remaining_queue = files[index - 1 :]
            if not input_file.exists():
                skipped += 1
                self._log(f"[SKIP] 入力ファイルが存在しません: {input_file}")
                continue

            output_file = self._build_output_path(input_file)
            output_file = self._resolve_output_collision(output_file)
            if output_file is None:
                skipped += 1
                self._log(f"[SKIP] 出力衝突によりスキップ: {input_file.name}")
                continue

            cmd = self._build_ffmpeg_command(input_file, output_file)
            self._log(f"[RUN] {input_file.name} -> {output_file.name}")
            duration_sec = self._probe_duration(input_file)

            ok, stderr_lines = self._run_ffmpeg_with_progress(cmd, input_file, duration_sec, index, total)
            if ok:
                success += 1
                self._log(f"[OK] {output_file}")
            else:
                failed += 1
                self._log(f"[NG] {input_file.name} の変換に失敗しました。")
                for line in stderr_lines[-10:]:
                    self._log(f"  {line}")

            self.remaining_queue = files[index:]

        self._set_progress(0.0, f"完了: success={success} failed={failed} skipped={skipped}")
        self._log(f"完了: success={success}, failed={failed}, skipped={skipped}")

    def _build_output_path(self, input_file: Path) -> Path:
        output_dir = Path(self.output_dir_var.get()).expanduser() if self.output_dir_var.get().strip() else input_file.parent
        output_dir.mkdir(parents=True, exist_ok=True)

        mode = self.mode_var.get()
        if mode == "video_compress":
            ext = f".{self.video_format_var.get().strip()}"
        else:
            ext = f".{self.audio_format_var.get()}"

        policy = self.name_policy_var.get()
        if policy == "source_name":
            name = input_file.stem
        elif policy == "source_name_with_suffix":
            name = f"{input_file.stem}_compressed"
        else:
            name = self.custom_name_var.get().strip()

        return output_dir / f"{name}{ext}"

    def _resolve_output_collision(self, output_file: Path) -> Path | None:
        if not output_file.exists():
            return output_file
        choice = self.overwrite_resolver.ask(output_file)
        if choice == "overwrite":
            return output_file
        if choice == "skip":
            return None

        base = output_file.stem
        suffix = output_file.suffix
        parent = output_file.parent
        i = 1
        while True:
            candidate = parent / f"{base}_{i}{suffix}"
            if not candidate.exists():
                return candidate
            i += 1

    def _build_ffmpeg_command(self, input_file: Path, output_file: Path) -> list[str]:
        ffmpeg_cmd = self._resolve_tool_command("ffmpeg")
        mode = self.mode_var.get()
        if mode == "video_compress":
            cmd = [
                ffmpeg_cmd,
                "-y",
                "-i",
                str(input_file),
                "-vf",
                f"scale={self.scale_var.get().strip()}",
                "-r",
                self.fps_var.get().strip(),
                "-c:v",
                self.video_encoder_var.get().strip(),
                "-crf",
                self.crf_var.get().strip(),
                "-preset",
                "medium",
                "-c:a",
                self.video_audio_codec_var.get().strip(),
                "-b:a",
                self.video_audio_bitrate_var.get().strip(),
                "-progress",
                "pipe:1",
                "-nostats",
            ]
            if self.video_format_var.get().strip() in {"mp4", "mov"}:
                cmd.extend(["-movflags", "+faststart"])
            cmd.append(str(output_file))
            return cmd

        cmd = [ffmpeg_cmd, "-y", "-i", str(input_file)]
        if mode == "video_extract_audio":
            cmd.extend(["-vn"])
        cmd.extend(["-af", f"atempo={self.playback_speed_var.get().strip()}"])
        cmd.extend(["-c:a", self.audio_codec_var.get().strip(), "-b:a", self.audio_bitrate_var.get().strip()])
        sample_rate = self.sample_rate_var.get().strip()
        if sample_rate:
            cmd.extend(["-ar", sample_rate])
        cmd.extend(["-progress", "pipe:1", "-nostats", str(output_file)])
        return cmd

    def _probe_duration(self, input_file: Path) -> float:
        ffprobe_cmd = self._resolve_tool_command("ffprobe")
        cmd = [
            ffprobe_cmd,
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(input_file),
        ]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=True,
                creationflags=WINDOWS_CREATE_NO_WINDOW,
            )
            value = result.stdout.strip()
            return float(value) if value else 0.0
        except Exception:
            return 0.0

    def _detect_ffmpeg_encoders(self) -> set[str]:
        ffmpeg_cmd = self._resolve_tool_command("ffmpeg")
        cmd = [ffmpeg_cmd, "-hide_banner", "-encoders"]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                check=True,
                creationflags=WINDOWS_CREATE_NO_WINDOW,
            )
        except Exception:
            return set()

        encoders: set[str] = set()
        for line in result.stdout.splitlines():
            match = re.match(r"^\s*[A-Z\.]{6}\s+([0-9A-Za-z_\-]+)\s", line)
            if match:
                encoders.add(match.group(1))
        return encoders

    def _is_encoder_available(self, encoder_name: str) -> bool:
        if not self.available_ffmpeg_encoders:
            return True
        return encoder_name in self.available_ffmpeg_encoders

    def _resolve_tool_command(self, tool_name: str) -> str:
        candidates = self._tool_candidates(tool_name)
        for candidate in candidates:
            candidate_path = Path(candidate)
            if candidate_path.is_file():
                resolved = str(candidate_path)
                self._log_tool_resolution(tool_name, resolved)
                return resolved

        path_resolved = shutil.which(tool_name)
        if path_resolved:
            self._log_tool_resolution(tool_name, path_resolved)
            return path_resolved

        return tool_name

    def _tool_candidates(self, tool_name: str) -> list[Path]:
        executable_name = f"{tool_name}.exe" if os.name == "nt" else tool_name
        base_dirs: list[Path] = []

        if getattr(sys, "frozen", False):
            base_dirs.append(Path(sys.executable).resolve().parent)

        if "__file__" in globals():
            base_dirs.append(Path(__file__).resolve().parent)

        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            base_dirs.append(Path(meipass))

        candidates: list[Path] = []
        seen: set[Path] = set()
        for base_dir in base_dirs:
            for candidate in (base_dir / executable_name, base_dir / "resources" / executable_name):
                if candidate not in seen:
                    seen.add(candidate)
                    candidates.append(candidate)
        return candidates

    def _log_tool_resolution(self, tool_name: str, resolved: str) -> None:
        if tool_name in self._tool_resolution_logged:
            return
        self._tool_resolution_logged.add(tool_name)
        self._log(f"{tool_name} を使用します: {resolved}")

    def _run_ffmpeg_with_progress(
        self, cmd: list[str], input_file: Path, duration_sec: float, index: int, total: int
    ) -> tuple[bool, list[str]]:
        stderr_lines: list[str] = []
        progress_state: dict[str, str] = {}
        last_progress_log_at = 0.0
        try:
            self.current_process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=WINDOWS_CREATE_NO_WINDOW,
            )
        except FileNotFoundError:
            self._log("ffmpeg が見つかりません。PATH を確認してください。")
            return False, stderr_lines

        assert self.current_process.stdout is not None
        assert self.current_process.stderr is not None

        def _consume_stderr() -> None:
            assert self.current_process is not None
            assert self.current_process.stderr is not None
            for raw_line in self.current_process.stderr:
                line = raw_line.strip()
                if not line:
                    continue
                stderr_lines.append(line)
                # ffmpeg banner/警告/エラーを少し見せる
                if line.startswith("Input #") or line.startswith("Output #") or "Error" in line or "error" in line:
                    self._log(f"[ffmpeg] {line}")

        stderr_thread = threading.Thread(target=_consume_stderr, daemon=True)
        stderr_thread.start()

        while True:
            if self.stop_event.is_set():
                proc = self.current_process
                if proc and proc.poll() is None:
                    proc.terminate()
                break

            line = self.current_process.stdout.readline()
            if not line:
                if self.current_process.poll() is not None:
                    break
                continue
            stripped = line.strip()
            if "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            progress_state[key] = value

            if key == "out_time_ms":
                try:
                    out_time_ms = int(value)
                    out_sec = out_time_ms / 1_000_000.0
                    if duration_sec > 0:
                        file_percent = min(100.0, (out_sec / duration_sec) * 100.0)
                    else:
                        file_percent = 0.0
                    overall_percent = ((index - 1) + (file_percent / 100.0)) / total * 100.0
                    self._set_progress(overall_percent, f"{input_file.name}: {file_percent:.1f}% ({index}/{total})")
                except ValueError:
                    continue
                continue

            if key == "progress":
                now = time.monotonic()
                should_log = (now - last_progress_log_at) >= 1.0 or value == "end"
                if should_log:
                    out_time = progress_state.get("out_time", "N/A")
                    speed = progress_state.get("speed", "N/A")
                    bitrate = progress_state.get("bitrate", "N/A")
                    self._log(
                        f"[ffmpeg] {input_file.name} progress={value} time={out_time} speed={speed} bitrate={bitrate}"
                    )
                    last_progress_log_at = now

        stderr_thread.join(timeout=0.5)
        return self.current_process.returncode == 0, stderr_lines

    def _log(self, message: str) -> None:
        self.log_queue.put(message)

    def _set_progress(self, percent: float, text: str) -> None:
        self.progress_queue.put(ProgressEvent(file_path=Path("."), percent=percent, text=text))

    def _start_polling(self) -> None:
        self._poll_queues()

    def _poll_queues(self) -> None:
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_text.insert("", END, values=(msg,))
            children = self.log_text.get_children()
            if children:
                self.log_text.see(children[-1])

        while True:
            try:
                ev = self.progress_queue.get_nowait()
            except queue.Empty:
                break
            self.progress_bar["value"] = ev.percent
            self.progress_text_var.set(ev.text)

        self.root.after(150, self._poll_queues)

    def _settings_payload(self) -> dict[str, str]:
        return {
            "mode": self.mode_var.get(),
            "output_dir": self.output_dir_var.get(),
            "name_policy": self.name_policy_var.get(),
            "custom_name": self.custom_name_var.get(),
            "playback_speed": self.playback_speed_var.get(),
            "audio_format": self.audio_format_var.get(),
            "audio_codec": self.audio_codec_var.get(),
            "audio_bitrate": self.audio_bitrate_var.get(),
            "sample_rate": self.sample_rate_var.get(),
            "preset": self.preset_var.get(),
            "scale": self.scale_var.get(),
            "fps": self.fps_var.get(),
            "crf": self.crf_var.get(),
            "video_format": self.video_format_var.get(),
            "video_codec": self.video_codec_var.get(),
            "video_encoder": self.video_encoder_var.get(),
            "video_audio_codec": self.video_audio_codec_var.get(),
            "video_audio_bitrate": self.video_audio_bitrate_var.get(),
        }

    def _save_settings(self) -> None:
        payload = self._settings_payload()
        with SETTINGS_PATH.open("w", encoding="utf-8", newline="\n") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        self._log(f"設定を保存しました: {SETTINGS_PATH}")

    def _load_settings(self) -> None:
        defaults = self._settings_payload()
        loaded: dict[str, str] = defaults
        if SETTINGS_PATH.exists():
            try:
                with SETTINGS_PATH.open("r", encoding="utf-8") as f:
                    raw = json.load(f)
                if isinstance(raw, dict):
                    loaded = {**defaults, **{k: str(v) for k, v in raw.items() if k in defaults}}
                else:
                    loaded = defaults
            except Exception:
                loaded = defaults
                self._log("設定ファイルが読み込めなかったため、初期値に戻して上書きします。")
                with SETTINGS_PATH.open("w", encoding="utf-8", newline="\n") as f:
                    json.dump(defaults, f, ensure_ascii=False, indent=2)
        else:
            with SETTINGS_PATH.open("w", encoding="utf-8", newline="\n") as f:
                json.dump(defaults, f, ensure_ascii=False, indent=2)

        self.mode_var.set(loaded["mode"])
        self.output_dir_var.set(loaded["output_dir"])
        self.name_policy_var.set(loaded["name_policy"])
        self.custom_name_var.set(loaded["custom_name"])
        self.playback_speed_var.set(loaded["playback_speed"])
        self.audio_format_var.set(loaded["audio_format"])
        self.audio_codec_var.set(loaded["audio_codec"])
        self.audio_bitrate_var.set(loaded["audio_bitrate"])
        self.sample_rate_var.set(loaded["sample_rate"])
        self.preset_var.set(loaded["preset"])
        self.scale_var.set(loaded["scale"])
        self.fps_var.set(loaded["fps"])
        self.crf_var.set(loaded["crf"])
        self.video_format_var.set(loaded["video_format"])
        legacy_video_codec = loaded["video_codec"]
        if legacy_video_codec in VIDEO_ENCODER_DESCRIPTIONS:
            codec_from_encoder = next(
                (codec for codec, encoders in VIDEO_ENCODER_OPTIONS.items() if legacy_video_codec in encoders),
                "h264",
            )
            self.video_codec_var.set(codec_from_encoder)
            self.video_encoder_var.set(legacy_video_codec)
        else:
            self.video_codec_var.set(legacy_video_codec)
            self.video_encoder_var.set(loaded["video_encoder"])
        self.video_audio_codec_var.set(loaded["video_audio_codec"])
        self.video_audio_bitrate_var.set(loaded["video_audio_bitrate"])
        self._sync_video_codec_options()
        self._toggle_mode_sections()


def create_root() -> Tk:
    if HAS_DND and TkinterDnD is not None:
        return TkinterDnD.Tk()
    return Tk()


def main() -> None:
    root = create_root()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
