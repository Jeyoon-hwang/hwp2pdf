#!/usr/bin/env python3
"""HWP/HWPX → PDF 대량 변환 GUI 앱"""

import os
import platform
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.request
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


LIBREOFFICE_VERSION = "25.2.7"
LIBREOFFICE_MSI_URL = (
    f"https://download.documentfoundation.org/libreoffice/stable/"
    f"{LIBREOFFICE_VERSION}/win/x86_64/"
    f"LibreOffice_{LIBREOFFICE_VERSION}_Win_x86-64.msi"
)


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


def download_and_install_libreoffice(status_callback=None):
    """Windows에서 LibreOffice MSI를 다운로드하고 자동 설치합니다."""
    if not IS_WINDOWS:
        return False

    msi_path = os.path.join(tempfile.gettempdir(), f"LibreOffice_{LIBREOFFICE_VERSION}_Win_x86-64.msi")

    # 다운로드
    if status_callback:
        status_callback("LibreOffice 다운로드 중... (약 350MB, 잠시 기다려주세요)")

    try:
        def _reporthook(block_num, block_size, total_size):
            downloaded = block_num * block_size
            if total_size > 0:
                pct = min(downloaded * 100 // total_size, 100)
                mb_down = downloaded // (1024 * 1024)
                mb_total = total_size // (1024 * 1024)
                if status_callback:
                    status_callback(f"LibreOffice 다운로드 중... {mb_down}MB / {mb_total}MB ({pct}%)")

        urllib.request.urlretrieve(LIBREOFFICE_MSI_URL, msi_path, reporthook=_reporthook)
    except Exception as e:
        if status_callback:
            status_callback(f"다운로드 실패: {e}")
        return False

    # 자동 설치 (msiexec /passive = 진행률만 표시, 사용자 입력 불필요)
    if status_callback:
        status_callback("LibreOffice 설치 중... (자동 설치, 잠시 기다려주세요)")

    try:
        result = subprocess.run(
            ["msiexec", "/i", msi_path, "/passive", "/norestart"],
            timeout=600,
        )
        if result.returncode != 0:
            if status_callback:
                status_callback(f"설치 실패 (코드: {result.returncode})")
            return False
    except Exception as e:
        if status_callback:
            status_callback(f"설치 오류: {e}")
        return False
    finally:
        try:
            os.remove(msi_path)
        except OSError:
            pass

    if status_callback:
        status_callback("LibreOffice 설치 완료!")
    return True


class MultiFolderDialog:
    """여러 폴더를 한 번에 선택할 수 있는 커스텀 다이얼로그"""

    def __init__(self, parent, title="폴더 선택"):
        self.result = []

        self.dlg = tk.Toplevel(parent)
        self.dlg.title(title)
        self.dlg.geometry("600x500")
        self.dlg.minsize(400, 300)
        self.dlg.transient(parent)
        self.dlg.grab_set()

        # --- 상단: 현재 경로 & 이동 ---
        nav_frame = ttk.Frame(self.dlg)
        nav_frame.pack(fill="x", padx=8, pady=(8, 4))

        ttk.Label(nav_frame, text="경로:").pack(side="left")
        self.var_path = tk.StringVar(value=os.path.expanduser("~"))
        self.entry_path = ttk.Entry(nav_frame, textvariable=self.var_path)
        self.entry_path.pack(side="left", fill="x", expand=True, padx=(4, 4))
        ttk.Button(nav_frame, text="이동", command=self._navigate).pack(side="left")
        ttk.Button(nav_frame, text="상위 폴더", command=self._go_up).pack(side="left", padx=(4, 0))

        # --- 중간 왼쪽: 폴더 트리 ---
        mid_frame = ttk.Frame(self.dlg)
        mid_frame.pack(fill="both", expand=True, padx=8, pady=4)

        left_frame = ttk.LabelFrame(mid_frame, text="폴더 탐색 (Ctrl/Shift+클릭으로 다중 선택)")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))

        self.tree = tk.Listbox(left_frame, selectmode="extended", height=15)
        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._on_double_click)

        # --- 중간 버튼 ---
        btn_mid = ttk.Frame(mid_frame)
        btn_mid.pack(side="left", padx=4)
        ttk.Button(btn_mid, text="추가 →", command=self._add_selected).pack(pady=2)
        ttk.Button(btn_mid, text="← 제거", command=self._remove_selected).pack(pady=2)
        ttk.Button(btn_mid, text="현재 폴더\n추가 →", command=self._add_current).pack(pady=(10, 2))

        # --- 중간 오른쪽: 선택된 폴더 목록 ---
        right_frame = ttk.LabelFrame(mid_frame, text="선택된 폴더")
        right_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))

        self.selected_list = tk.Listbox(right_frame, selectmode="extended", height=15)
        sel_scroll = ttk.Scrollbar(right_frame, orient="vertical", command=self.selected_list.yview)
        self.selected_list.configure(yscrollcommand=sel_scroll.set)
        self.selected_list.pack(side="left", fill="both", expand=True)
        sel_scroll.pack(side="right", fill="y")

        # --- 하단 버튼 ---
        bottom = ttk.Frame(self.dlg)
        bottom.pack(fill="x", padx=8, pady=8)
        ttk.Button(bottom, text="확인", command=self._ok).pack(side="right", padx=(4, 0))
        ttk.Button(bottom, text="취소", command=self._cancel).pack(side="right")

        self._selected_folders = []
        self._populate(self.var_path.get())

        self.dlg.wait_window()

    def _populate(self, path):
        """지정 경로의 하위 폴더 목록을 표시"""
        self.tree.delete(0, "end")
        self.var_path.set(path)
        try:
            entries = sorted(os.listdir(path), key=str.lower)
            for e in entries:
                full = os.path.join(path, e)
                if os.path.isdir(full) and not e.startswith("."):
                    self.tree.insert("end", e)
        except PermissionError:
            self.tree.insert("end", "(접근 권한 없음)")

    def _navigate(self):
        p = self.var_path.get().strip()
        if os.path.isdir(p):
            self._populate(p)

    def _go_up(self):
        cur = self.var_path.get()
        parent = os.path.dirname(cur)
        if parent and parent != cur:
            self._populate(parent)

    def _on_double_click(self, event):
        sel = self.tree.curselection()
        if not sel:
            return
        name = self.tree.get(sel[0])
        if name == "(접근 권한 없음)":
            return
        full = os.path.join(self.var_path.get(), name)
        if os.path.isdir(full):
            self._populate(full)

    def _add_selected(self):
        """왼쪽에서 선택한 폴더를 오른쪽 목록에 추가"""
        sel = self.tree.curselection()
        base = self.var_path.get()
        for i in sel:
            name = self.tree.get(i)
            if name == "(접근 권한 없음)":
                continue
            full = os.path.join(base, name)
            if full not in self._selected_folders:
                self._selected_folders.append(full)
                self.selected_list.insert("end", full)

    def _add_current(self):
        """현재 경로 자체를 오른쪽 목록에 추가"""
        cur = self.var_path.get()
        if cur and cur not in self._selected_folders:
            self._selected_folders.append(cur)
            self.selected_list.insert("end", cur)

    def _remove_selected(self):
        """오른쪽 목록에서 선택한 항목 제거"""
        sel = list(self.selected_list.curselection())
        sel.reverse()
        for i in sel:
            self._selected_folders.pop(i)
            self.selected_list.delete(i)

    def _ok(self):
        self.result = list(self._selected_folders)
        self.dlg.destroy()

    def _cancel(self):
        self.result = []
        self.dlg.destroy()


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
        dialog = MultiFolderDialog(self.root, title="HWP 파일이 있는 폴더 선택")
        if not dialog.result:
            return
        found = []
        for folder in dialog.result:
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
                answer = messagebox.askyesno(
                    "LibreOffice 설치 필요",
                    "LibreOffice가 설치되어 있지 않습니다.\n\n"
                    "자동으로 다운로드하고 설치할까요?\n"
                    "(약 350MB 다운로드, 자동 설치)",
                )
                if answer:
                    self.converting = True
                    self.btn_convert.config(state="disabled")
                    self.log_text.configure(state="normal")
                    self.log_text.delete("1.0", "end")
                    self.log_text.configure(state="disabled")
                    thread = threading.Thread(target=self._install_and_convert, daemon=True)
                    thread.start()
                return
            else:
                messagebox.showerror(
                    "오류",
                    "LibreOffice가 설치되어 있지 않습니다.\n\n"
                    "터미널에서 다음 명령어로 설치하세요:\n"
                    "  brew install --cask libreoffice",
                )
                return

        self.converting = True
        self.btn_convert.config(state="disabled")
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        thread = threading.Thread(target=self._convert_worker, args=(soffice,), daemon=True)
        thread.start()

    def _install_and_convert(self):
        """LibreOffice를 설치한 뒤 변환을 이어서 실행합니다."""
        def _status(msg):
            self.root.after(0, lambda: self.lbl_status.config(text=msg))
            self.root.after(0, lambda: self._log(msg))

        _status("LibreOffice 자동 설치를 시작합니다...")
        ok = download_and_install_libreoffice(status_callback=_status)

        if not ok:
            self.root.after(0, lambda: messagebox.showerror(
                "설치 실패",
                "LibreOffice 자동 설치에 실패했습니다.\n\n"
                "https://www.libreoffice.org 에서 직접 설치해 주세요.",
            ))
            self.root.after(0, lambda: self.btn_convert.config(state="normal"))
            self.converting = False
            return

        soffice = find_soffice()
        if not soffice:
            self.root.after(0, lambda: messagebox.showerror(
                "오류", "설치 후에도 LibreOffice를 찾을 수 없습니다."
            ))
            self.root.after(0, lambda: self.btn_convert.config(state="normal"))
            self.converting = False
            return

        self.root.after(0, lambda: self._log(""))
        self._convert_worker(soffice)

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
