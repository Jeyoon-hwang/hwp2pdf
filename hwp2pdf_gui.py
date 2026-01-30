#!/usr/bin/env python3
"""HWP/HWPX → PDF 대량 변환 GUI 앱"""

import os
import platform
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

IS_WINDOWS = platform.system() == "Windows"

SOFFICE_PATHS_MAC = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
]

SOFFICE_PATHS_WIN = [
    os.path.join(os.environ.get("PROGRAMFILES", r"C:\Program Files"), "LibreOffice", "program", "soffice.exe"),
    os.path.join(os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"), "LibreOffice", "program", "soffice.exe"),
]


def find_soffice():
    candidates = SOFFICE_PATHS_WIN if IS_WINDOWS else SOFFICE_PATHS_MAC
    for p in candidates:
        if os.path.isfile(p):
            return p
    # PATH에서 찾기
    try:
        cmd = "where" if IS_WINDOWS else "which"
        result = subprocess.run(
            [cmd, "soffice"], capture_output=True, text=True, timeout=5
        )
        if result.returncode == 0:
            return result.stdout.strip().splitlines()[0]
    except Exception:
        pass
    return None


class HwpToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HWP → PDF 대량 변환기")
        self.root.geometry("700x520")
        self.root.minsize(600, 450)
        self.root.resizable(True, True)

        self.files = []
        self.output_dir = ""
        self.converting = False

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- 파일 선택 ---
        frame_files = ttk.LabelFrame(self.root, text="1. HWP/HWPX 파일 선택")
        frame_files.pack(fill="x", **pad)

        btn_frame = ttk.Frame(frame_files)
        btn_frame.pack(fill="x", padx=8, pady=5)

        ttk.Button(btn_frame, text="파일 추가", command=self._add_files).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(btn_frame, text="폴더 추가", command=self._add_folder).pack(
            side="left", padx=(0, 5)
        )
        ttk.Button(btn_frame, text="목록 초기화", command=self._clear_files).pack(
            side="left"
        )
        self.lbl_count = ttk.Label(btn_frame, text="0개 파일")
        self.lbl_count.pack(side="right")

        list_frame = ttk.Frame(frame_files)
        list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.listbox = tk.Listbox(list_frame, height=8, selectmode="extended")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # --- 출력 폴더 ---
        frame_out = ttk.LabelFrame(self.root, text="2. 출력 폴더")
        frame_out.pack(fill="x", **pad)

        out_inner = ttk.Frame(frame_out)
        out_inner.pack(fill="x", padx=8, pady=8)

        self.var_outdir = tk.StringVar(value="(원본 파일과 같은 폴더)")
        ttk.Entry(out_inner, textvariable=self.var_outdir, state="readonly").pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        ttk.Button(out_inner, text="폴더 선택", command=self._select_output).pack(side="left")
        ttk.Button(out_inner, text="초기화", command=self._reset_output).pack(side="left", padx=(5, 0))

        # --- 진행 상태 ---
        frame_prog = ttk.LabelFrame(self.root, text="3. 변환")
        frame_prog.pack(fill="x", **pad)

        prog_inner = ttk.Frame(frame_prog)
        prog_inner.pack(fill="x", padx=8, pady=8)

        self.progress = ttk.Progressbar(prog_inner, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 5))

        self.lbl_status = ttk.Label(prog_inner, text="대기 중")
        self.lbl_status.pack(anchor="w")

        # --- 로그 ---
        log_frame = ttk.Frame(frame_prog)
        log_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        log_font = ("Consolas", 10) if IS_WINDOWS else ("Menlo", 11)
        self.log_text = tk.Text(log_frame, height=6, state="disabled", font=log_font)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        # --- 변환 버튼 ---
        self.btn_convert = ttk.Button(
            self.root, text="변환 시작", command=self._start_convert
        )
        self.btn_convert.pack(pady=10)

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="HWP/HWPX 파일 선택",
            filetypes=[("한글 파일", "*.hwp *.hwpx *.HWP *.HWPX"), ("모든 파일", "*.*")],
        )
        self._append_files(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="HWP 파일이 있는 폴더 선택")
        if not folder:
            return
        found = []
        for f in os.listdir(folder):
            if f.lower().endswith((".hwp", ".hwpx")):
                found.append(os.path.join(folder, f))
        found.sort()
        self._append_files(found)

    def _append_files(self, paths):
        existing = set(self.files)
        for p in paths:
            if p not in existing:
                self.files.append(p)
                self.listbox.insert("end", os.path.basename(p))
        self.lbl_count.config(text=f"{len(self.files)}개 파일")

    def _clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.lbl_count.config(text="0개 파일")

    def _select_output(self):
        folder = filedialog.askdirectory(title="PDF 출력 폴더 선택")
        if folder:
            self.output_dir = folder
            self.var_outdir.set(folder)

    def _reset_output(self):
        self.output_dir = ""
        self.var_outdir.set("(원본 파일과 같은 폴더)")

    def _start_convert(self):
        if self.converting:
            return
        if not self.files:
            messagebox.showwarning("알림", "변환할 파일을 추가해 주세요.")
            return

        soffice = find_soffice()
        if not soffice:
            if IS_WINDOWS:
                install_msg = (
                    "LibreOffice가 설치되어 있지 않습니다.\n\n"
                    "https://www.libreoffice.org 에서 다운로드하세요."
                )
            else:
                install_msg = (
                    "LibreOffice가 설치되어 있지 않습니다.\n\n"
                    "터미널에서 다음 명령어로 설치하세요:\n"
                    "  brew install --cask libreoffice"
                )
            messagebox.showerror("오류", install_msg)
            return

        self.converting = True
        self.btn_convert.config(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        thread = threading.Thread(target=self._convert_worker, args=(soffice,), daemon=True)
        thread.start()

    def _convert_worker(self, soffice):
        total = len(self.files)
        success = 0
        fail = 0
        failed_names = []

        self.root.after(0, lambda: self.progress.config(maximum=total, value=0))
        self.root.after(0, lambda: self.lbl_status.config(text=f"변환 중... 0/{total}"))

        for i, filepath in enumerate(self.files):
            name = os.path.basename(filepath)
            outdir = self.output_dir if self.output_dir else os.path.dirname(filepath)

            self.root.after(0, lambda n=name, idx=i: (
                self.lbl_status.config(text=f"변환 중... {idx + 1}/{total}  -  {n}"),
                self._log(f"[{idx + 1}/{total}] {n} ..."),
            ))

            try:
                result = subprocess.run(
                    [soffice, "--headless", "--convert-to", "pdf", "--outdir", outdir, filepath],
                    capture_output=True, text=True, timeout=120,
                )
                if result.returncode == 0:
                    success += 1
                    self.root.after(0, lambda n=name: self._log(f"  -> 완료"))
                else:
                    fail += 1
                    failed_names.append(name)
                    self.root.after(0, lambda n=name: self._log(f"  -> 실패"))
            except subprocess.TimeoutExpired:
                fail += 1
                failed_names.append(name)
                self.root.after(0, lambda n=name: self._log(f"  -> 시간 초과"))
            except Exception as e:
                fail += 1
                failed_names.append(name)
                self.root.after(0, lambda n=name, e=e: self._log(f"  -> 오류: {e}"))

            self.root.after(0, lambda v=i + 1: self.progress.config(value=v))

        # 완료
        summary = f"완료! 성공: {success}, 실패: {fail} / 총 {total}개"
        self.root.after(0, lambda: self.lbl_status.config(text=summary))
        self.root.after(0, lambda: self._log(f"\n{'=' * 40}"))
        self.root.after(0, lambda: self._log(summary))

        if failed_names:
            self.root.after(0, lambda: self._log("실패한 파일:"))
            for fn in failed_names:
                self.root.after(0, lambda fn=fn: self._log(f"  - {fn}"))

        self.root.after(0, lambda: self.btn_convert.config(state="normal"))
        self.root.after(0, lambda: messagebox.showinfo("변환 완료", summary))
        self.converting = False


def main():
    root = tk.Tk()
    HwpToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
