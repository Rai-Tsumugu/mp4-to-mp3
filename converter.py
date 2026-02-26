"""
MP4 to MP3 Converter
Drag & Drop GUI application using ffmpeg
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import subprocess
import threading
import os
import re
import queue
from pathlib import Path


class MP4toMP3Converter(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("MP4 → MP3 Converter")
        self.geometry("750x660")
        self.resizable(True, True)
        self.configure(bg="#1e1e2e")

        self._conversion_queue = queue.Queue()
        self._current_proc = None
        self._cancel_flag = False
        self._is_converting = False
        self._file_items: dict[str, str] = {}  # path -> iid in treeview

        self._build_ui()
        self._setup_dnd()

    # ------------------------------------------------------------------ UI --

    def _build_ui(self):
        PAD = 12

        # ── Drop zone ──────────────────────────────────────────────────────
        drop_frame = tk.Frame(self, bg="#313244", bd=0, highlightthickness=2,
                              highlightbackground="#7c7f9c", relief="flat")
        drop_frame.pack(fill="x", padx=PAD, pady=(PAD, 6))

        self._drop_label = tk.Label(
            drop_frame,
            text="ここに MP4 ファイルをドラッグ＆ドロップ\n（複数ファイル対応）",
            font=("Meiryo", 13),
            fg="#cdd6f4", bg="#313244",
            pady=22,
        )
        self._drop_label.pack(fill="both")

        # ── Output directory row ───────────────────────────────────────────
        row = tk.Frame(self, bg="#1e1e2e")
        row.pack(fill="x", padx=PAD, pady=(0, 6))

        tk.Label(row, text="出力先:", font=("Meiryo", 10),
                 fg="#a6adc8", bg="#1e1e2e").pack(side="left")

        self._out_var = tk.StringVar(value="入力ファイルと同じフォルダ")
        self._out_entry = tk.Entry(row, textvariable=self._out_var,
                                   font=("Consolas", 9), fg="#cdd6f4",
                                   bg="#313244", insertbackground="white",
                                   relief="flat", bd=4)
        self._out_entry.pack(side="left", fill="x", expand=True, padx=(6, 4))

        tk.Button(row, text="選択…", command=self._choose_output,
                  font=("Meiryo", 9), bg="#45475a", fg="#cdd6f4",
                  activebackground="#585b70", relief="flat",
                  cursor="hand2", padx=6).pack(side="left")

        tk.Button(row, text="クリア", command=self._clear_output,
                  font=("Meiryo", 9), bg="#45475a", fg="#cdd6f4",
                  activebackground="#585b70", relief="flat",
                  cursor="hand2", padx=6).pack(side="left", padx=(4, 0))

        # ── Quality row ────────────────────────────────────────────────────
        qrow = tk.Frame(self, bg="#1e1e2e")
        qrow.pack(fill="x", padx=PAD, pady=(0, 6))

        tk.Label(qrow, text="品質 (kbps):", font=("Meiryo", 10),
                 fg="#a6adc8", bg="#1e1e2e").pack(side="left")

        self._bitrate_var = tk.StringVar(value="192")
        for br in ("128", "192", "256", "320"):
            tk.Radiobutton(
                qrow, text=br, variable=self._bitrate_var, value=br,
                font=("Meiryo", 10), fg="#cdd6f4", bg="#1e1e2e",
                activebackground="#1e1e2e", activeforeground="#89b4fa",
                selectcolor="#313244", relief="flat",
            ).pack(side="left", padx=6)

        # ── File list ──────────────────────────────────────────────────────
        list_frame = tk.Frame(self, bg="#1e1e2e")
        list_frame.pack(fill="both", expand=True, padx=PAD, pady=(0, 6))

        cols = ("file", "status")
        self._tree = ttk.Treeview(list_frame, columns=cols, show="headings",
                                   selectmode="extended", height=10)
        self._tree.heading("file", text="ファイル名")
        self._tree.heading("status", text="状態")
        self._tree.column("file", width=480, stretch=True)
        self._tree.column("status", width=120, anchor="center", stretch=False)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#313244",
                         foreground="#cdd6f4", fieldbackground="#313244",
                         rowheight=26, font=("Meiryo", 10))
        style.configure("Treeview.Heading", background="#45475a",
                         foreground="#cdd6f4", font=("Meiryo", 10, "bold"))
        style.map("Treeview", background=[("selected", "#585b70")])

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical",
                                   command=self._tree.yview)
        self._tree.configure(yscrollcommand=scrollbar.set)

        self._tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        tk.Button(self, text="選択ファイルを削除", command=self._remove_selected,
                  font=("Meiryo", 9), bg="#45475a", fg="#cdd6f4",
                  activebackground="#585b70", relief="flat",
                  cursor="hand2", padx=8).pack(anchor="e", padx=PAD)

        # ── Progress bar ───────────────────────────────────────────────────
        self._progress_var = tk.DoubleVar(value=0)
        self._progress = ttk.Progressbar(self, variable=self._progress_var,
                                          maximum=100, length=200)
        style.configure("TProgressbar", troughcolor="#313244",
                         background="#89b4fa", thickness=14)
        self._progress.pack(fill="x", padx=PAD, pady=(4, 2))

        # ── Status label ───────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="ファイルをドロップしてください")
        tk.Label(self, textvariable=self._status_var,
                 font=("Meiryo", 9), fg="#a6adc8",
                 bg="#1e1e2e", anchor="w").pack(fill="x", padx=PAD)

        # ── Action buttons ─────────────────────────────────────────────────
        btn_row = tk.Frame(self, bg="#1e1e2e")
        btn_row.pack(pady=(6, PAD))

        self._convert_btn = tk.Button(
            btn_row, text="  変換開始  ", command=self._start_conversion,
            font=("Meiryo", 12, "bold"), bg="#89b4fa", fg="#1e1e2e",
            activebackground="#74c7ec", relief="flat",
            cursor="hand2", padx=16, pady=6,
        )
        self._convert_btn.pack(side="left", padx=8)

        self._cancel_btn = tk.Button(
            btn_row, text="  キャンセル  ", command=self._cancel_conversion,
            font=("Meiryo", 12), bg="#f38ba8", fg="#1e1e2e",
            activebackground="#eba0ac", relief="flat",
            cursor="hand2", padx=16, pady=6, state="disabled",
        )
        self._cancel_btn.pack(side="left", padx=8)

    # --------------------------------------------------------- Drag & Drop --

    def _setup_dnd(self):
        for widget in (self, self._drop_label,
                       self._drop_label.master):
            widget.drop_target_register(DND_FILES)  # type: ignore
            widget.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore

    def _on_drop(self, event):
        raw = event.data
        # parse: paths may be space-separated or brace-wrapped on Windows
        paths = self._parse_dnd_paths(raw)
        added = 0
        for p in paths:
            if p.lower().endswith(".mp4") and p not in self._file_items:
                iid = self._tree.insert("", "end",
                                        values=(Path(p).name, "待機中"),
                                        tags=("pending",))
                self._file_items[p] = iid
                self._conversion_queue.put(p)
                added += 1
        if added:
            self._status_var.set(f"{added} ファイルを追加しました（合計 {len(self._file_items)} 件）")
        else:
            self._status_var.set("MP4 ファイルのみ対応しています")
        self._tree.tag_configure("pending", foreground="#cdd6f4")
        self._tree.tag_configure("done",    foreground="#a6e3a1")
        self._tree.tag_configure("error",   foreground="#f38ba8")
        self._tree.tag_configure("running", foreground="#89b4fa")

    @staticmethod
    def _parse_dnd_paths(raw: str) -> list[str]:
        paths = []
        raw = raw.strip()
        i = 0
        while i < len(raw):
            if raw[i] == "{":
                end = raw.index("}", i)
                paths.append(raw[i + 1:end])
                i = end + 2
            else:
                sp = raw.find(" ", i)
                if sp == -1:
                    paths.append(raw[i:])
                    break
                # check if next token starts with { or is another path
                chunk = raw[i:sp]
                paths.append(chunk)
                i = sp + 1
        return paths

    # --------------------------------------------------------- UI helpers --

    def _choose_output(self):
        d = filedialog.askdirectory(title="出力フォルダを選択")
        if d:
            self._out_var.set(d)

    def _clear_output(self):
        self._out_var.set("入力ファイルと同じフォルダ")

    def _remove_selected(self):
        if self._is_converting:
            return
        selected = self._tree.selection()
        for iid in selected:
            # find path by iid
            path = next((p for p, i in self._file_items.items() if i == iid), None)
            if path:
                del self._file_items[path]
            self._tree.delete(iid)

    def _resolve_output_dir(self, input_path: str) -> str:
        val = self._out_var.get().strip()
        if val == "入力ファイルと同じフォルダ" or not val:
            return str(Path(input_path).parent)
        return val

    # ------------------------------------------------------ Conversion ------

    def _start_conversion(self):
        if self._is_converting:
            return
        if self._tree.get_children() == ():
            messagebox.showinfo("情報", "変換するファイルがありません。\nMP4 ファイルをドロップしてください。")
            return

        self._is_converting = True
        self._cancel_flag = False
        self._convert_btn.configure(state="disabled")
        self._cancel_btn.configure(state="normal")

        # rebuild queue from pending items
        while not self._conversion_queue.empty():
            try:
                self._conversion_queue.get_nowait()
            except queue.Empty:
                break

        for path, iid in self._file_items.items():
            status = self._tree.item(iid, "values")[1]
            if status in ("待機中", "エラー"):
                self._tree.item(iid, values=(Path(path).name, "待機中"),
                                tags=("pending",))
                self._conversion_queue.put(path)

        threading.Thread(target=self._worker, daemon=True).start()

    def _cancel_conversion(self):
        self._cancel_flag = True
        if self._current_proc and self._current_proc.poll() is None:
            self._current_proc.kill()
        self._status_var.set("キャンセルしました")

    def _worker(self):
        total = self._conversion_queue.qsize()
        done = 0

        while not self._conversion_queue.empty() and not self._cancel_flag:
            try:
                path = self._conversion_queue.get_nowait()
            except queue.Empty:
                break

            iid = self._file_items.get(path)
            if not iid:
                continue
                
            self.after(0, lambda i=iid, n=Path(path).name:
                       self._tree.item(i, values=(n, "変換中…"),
                                       tags=("running",)))

            out_dir = self._resolve_output_dir(path)
            out_path = str(Path(out_dir) / (Path(path).stem + ".mp3"))
            bitrate = self._bitrate_var.get()

            success = self._run_ffmpeg(path, out_path, bitrate, iid)

            if success and not self._cancel_flag:
                done += 1
                self.after(0, lambda i=iid, n=Path(path).name:
                           self._tree.item(i, values=(n, "完了 ✓"),
                                           tags=("done",)))
            elif not self._cancel_flag:
                self.after(0, lambda i=iid, n=Path(path).name:
                           self._tree.item(i, values=(n, "エラー"),
                                           tags=("error",)))

            self.after(0, self._progress_var.set, 0)

        self.after(0, self._on_worker_done, done, total)

    def _run_ffmpeg(self, input_path: str, output_path: str,
                    bitrate: str, iid) -> bool:
        # First pass: get duration
        duration = self._get_duration(input_path)

        cmd = [
            "ffmpeg", "-y",
            "-i", input_path,
            "-vn",
            "-acodec", "libmp3lame",
            "-ab", f"{bitrate}k",
            "-ar", "44100",
            "-progress", "pipe:2",
            output_path,
        ]

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
            self._current_proc = proc

            name = Path(input_path).name
            out_us = 0  # microseconds processed

            if proc.stderr is None:
                proc.wait()
                return proc.returncode == 0
                
            for line in proc.stderr:
                if self._cancel_flag:
                    proc.kill()
                    return False
                line = line.strip()

                # parse progress output
                m = re.match(r"out_time_us=(\d+)", line)
                if m:
                    out_us = int(m.group(1))
                    if duration and duration > 0:
                        pct = min(out_us / (duration * 1_000_000) * 100, 99)
                        self.after(0, self._progress_var.set, pct)
                        elapsed_s = out_us / 1_000_000
                        self.after(0, self._status_var.set,
                                   f"{name}  {self._fmt_time(elapsed_s)} / {self._fmt_time(duration)}")
                    else:
                        self.after(0, self._status_var.set,
                                   f"変換中: {name}")

            proc.wait()
            return proc.returncode == 0

        except FileNotFoundError:
            messagebox.showerror("エラー",
                                 "ffmpeg が見つかりません。\nPATH に ffmpeg を追加してください。")
            return False

    @staticmethod
    def _get_duration(path: str) -> float | None:
        try:
            result = subprocess.run(
                ["ffprobe", "-v", "error", "-show_entries",
                 "format=duration", "-of", "default=noprint_wrappers=1:nokey=1",
                 path],
                capture_output=True, text=True,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
            return float(result.stdout.strip())
        except Exception:
            return None

    @staticmethod
    def _fmt_time(seconds: float) -> str:
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{h:02d}:{m:02d}:{s:02d}"

    def _on_worker_done(self, done: int, total: int):
        self._is_converting = False
        self._current_proc = None
        self._convert_btn.configure(state="normal")
        self._cancel_btn.configure(state="disabled")
        self._progress_var.set(100 if not self._cancel_flag else 0)
        if not self._cancel_flag:
            self._status_var.set(f"完了: {done}/{total} ファイルを変換しました")


# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = MP4toMP3Converter()
    app.mainloop()


if __name__ == "__main__":
    main()
