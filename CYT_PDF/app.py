"""
CYT PDF 工具 - 主程式骨架
架構：
  - PDFApp        : 主視窗 (繼承 CTk)
  - Sidebar       : 左側導覽列
  - BasePage      : 所有頁面的基礎類別
  - ThreadedTask  : 後台執行裝飾器 + 基礎類別 (防止 GUI 凍結)
"""

from __future__ import annotations

import threading
import functools
import queue
import traceback
from typing import Callable, Any

import customtkinter as ctk

# ─────────────────────────────────────────────
# 專案資訊 (由 release_helper.py 讀取)
# ─────────────────────────────────────────────
APP_VERSION = "1.1.1"
GITHUB_REPO = "ced-cyt/CYT_PDF" # 請根據實際 GitHub 帳號修改



# ─────────────────────────────────────────────
# 全域設定
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE  = "CYT PDF 工具"
APP_WIDTH  = 1100
APP_HEIGHT = 700  # 稍微增加高度以容納更多內容
SIDEBAR_W  = 200


# ═══════════════════════════════════════════════════════════════
# ThreadedTask — 裝飾器 & 基礎類別
# ═══════════════════════════════════════════════════════════════

def threaded_task(func: Callable) -> Callable:
    """
    裝飾器：將函式放到背景執行緒執行，防止 GUI 凍結。

    用法::

        @threaded_task
        def run_heavy_pdf_work(self):
            ...  # PDF 處理邏輯

    回傳值透過 result_queue 傳回（若有需要）。
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        t = threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True)
        t.start()
        return t
    return wrapper


class ThreadedTaskMixin:
    """
    基礎類別 Mixin：提供結構化的後台任務管理。
    繼承此類別的 Page 可直接呼叫 self.run_in_thread()。

    Attributes:
        _result_queue : 子執行緒回傳結果的佇列
        _task_thread  : 目前執行中的執行緒
    """

    def __init__(self):
        self._result_queue: queue.Queue = queue.Queue()
        self._task_thread: threading.Thread | None = None

    # ── 公開 API ──────────────────────────────────────────────

    def run_in_thread(
        self,
        target: Callable,
        *args,
        on_success: Callable[[Any], None] | None = None,
        on_error:   Callable[[Exception], None] | None = None,
        **kwargs,
    ) -> threading.Thread:
        """
        在背景執行緒中執行 target，完成後透過 CTk after() 回主執行緒。

        Args:
            target     : 要執行的函式（通常是 PDF 處理邏輯）
            *args      : 傳入 target 的位置參數
            on_success : 成功後在主執行緒呼叫，接收回傳值
            on_error   : 失敗後在主執行緒呼叫，接收 Exception
            **kwargs   : 傳入 target 的關鍵字參數
        """
        def _worker():
            try:
                result = target(*args, **kwargs)
                self._result_queue.put(("ok", result))
            except Exception as exc:
                self._result_queue.put(("err", exc))
                traceback.print_exc()

        self._task_thread = threading.Thread(target=_worker, daemon=True)
        self._task_thread.start()

        # 使用 CTk after() 輪詢結果，保持在主執行緒更新 UI
        self._poll_result(on_success, on_error)
        return self._task_thread

    def is_running(self) -> bool:
        """回傳目前是否有後台任務執行中。"""
        return self._task_thread is not None and self._task_thread.is_alive()

    # ── 內部 ──────────────────────────────────────────────────

    def _poll_result(
        self,
        on_success: Callable | None,
        on_error:   Callable | None,
        interval_ms: int = 100,
    ):
        """每 interval_ms 毫秒輪詢一次佇列，有結果才觸發回呼。"""
        try:
            status, payload = self._result_queue.get_nowait()
            if status == "ok" and on_success:
                on_success(payload)
            elif status == "err" and on_error:
                on_error(payload)
        except queue.Empty:
            # 檢查執行緒是否意外死掉 (防止 GUI 凍結等待)
            if self._task_thread and not self._task_thread.is_alive() and self._result_queue.empty():
                if on_error:
                    on_error(RuntimeError("背景任務異常終止 (沒有回傳任何結果)"))
                return

            # 尚未完成，繼續等待
            if hasattr(self, "after"):
                self.after(interval_ms, self._poll_result, on_success, on_error, interval_ms)


# ═══════════════════════════════════════════════════════════════
# BasePage — 所有頁面的基礎類別
# ═══════════════════════════════════════════════════════════════

class BasePage(ctk.CTkFrame, ThreadedTaskMixin):
    """
    所有功能頁面的基礎類別。
    """

    def __init__(self, parent: ctk.CTkFrame, app: "PDFApp", **kwargs):
        ctk.CTkFrame.__init__(self, parent, corner_radius=0, **kwargs)
        ThreadedTaskMixin.__init__(self)
        self.app = app
        self.build_ui()

    def build_ui(self) -> None:
        raise NotImplementedError

    def reset_state(self):
        pass

    def on_show(self) -> None:
        pass

    def on_hide(self) -> None:
        pass


class MergePage(BasePage):
    """
    PDF 合併頁面：支援檔案選取、清單排序、刪除、輸出目錄選取與自訂檔名。
    """

    def build_ui(self):
        self.files: list[str] = []
        self._selected_index: int | None = None
        self._list_item_widgets: list[ctk.CTkFrame] = []
        self.output_folder: str = ""

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # 1. 標題
        ctk.CTkLabel(self, text="PDF 合併", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, pady=(20, 10), sticky="w", padx=30)

        # 2. 選取區域
        self.drop_zone = ctk.CTkFrame(self, height=100, border_width=2, border_color=("gray70", "gray30"))
        self.drop_zone.grid(row=1, column=0, padx=30, pady=10, sticky="nsew")
        self.drop_zone.grid_propagate(False)
        self.drop_zone.columnconfigure(0, weight=1)
        self.drop_zone.rowconfigure(0, weight=1)

        self.drop_label = ctk.CTkLabel(self.drop_zone, text="點擊此處選擇要合併的 PDF 檔案", text_color=("gray40", "gray60"))
        self.drop_label.grid(row=0, column=0)
        self.drop_zone.bind("<Button-1>", lambda e: self._select_files())
        self.drop_label.bind("<Button-1>", lambda e: self._select_files())

        # 3. 檔案清單
        self.list_container = ctk.CTkFrame(self, fg_color="transparent")
        self.list_container.grid(row=2, column=0, padx=30, pady=10, sticky="nsew")
        self.list_container.columnconfigure(0, weight=1)
        self.list_container.rowconfigure(0, weight=1)

        self.scroll_frame = ctk.CTkScrollableFrame(self.list_container, label_text="待合併檔案清單")
        self.scroll_frame.grid(row=0, column=0, sticky="nsew")

        self.action_bar = ctk.CTkFrame(self.list_container, fg_color="transparent")
        self.action_bar.grid(row=0, column=1, padx=(10, 0), sticky="ns")
        
        ctk.CTkButton(self.action_bar, text="▲", width=40, command=self._move_up).pack(pady=5)
        ctk.CTkButton(self.action_bar, text="▼", width=40, command=self._move_down).pack(pady=5)
        ctk.CTkButton(self.action_bar, text="✕", width=40, fg_color="#E74C3C", hover_color="#C0392B", command=self._remove_selected).pack(pady=5)

        # 4. 輸出設定
        self.output_frame = ctk.CTkFrame(self)
        self.output_frame.grid(row=3, column=0, padx=30, pady=10, sticky="ew")
        self.output_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self.output_frame, text="輸出目錄:").grid(row=0, column=0, padx=10, pady=5)
        self.folder_label = ctk.CTkLabel(self.output_frame, text="預設為第一個檔案所在目錄...", text_color="gray")
        self.folder_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkButton(self.output_frame, text="瀏覽...", width=80, command=self._select_folder).grid(row=0, column=2, padx=10, pady=5)

        ctk.CTkLabel(self.output_frame, text="輸出檔名:").grid(row=1, column=0, padx=10, pady=5)
        self.filename_entry = ctk.CTkEntry(self.output_frame, placeholder_text="預設為 merged_result")
        self.filename_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")
        self.filename_entry.insert(0, "merged_result")

        # 5. 底部區域
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=4, column=0, padx=30, pady=(10, 20), sticky="ew")
        self.footer.columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(self.footer)
        self.progress_bar.grid(row=0, column=0, padx=(0, 20), sticky="ew")
        self.progress_bar.set(0)

        self.start_btn = ctk.CTkButton(self.footer, text="開始合併", width=120, height=40, font=ctk.CTkFont(weight="bold"), command=self._start_merge)
        self.start_btn.grid(row=0, column=1)

    def reset_state(self):
        self.files = []
        self._selected_index = None
        self.output_folder = ""
        self.folder_label.configure(text="預設為第一個檔案所在目錄...", text_color="gray")
        self.filename_entry.delete(0, "end")
        self.filename_entry.insert(0, "merged_result")
        self.progress_bar.set(0)
        self._refresh_list_ui()

    def _select_files(self):
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(filetypes=[("PDF 檔案", "*.pdf")])
        if paths:
            self.files.extend(list(paths))
            self._refresh_list_ui()

    def _select_folder(self):
        from tkinter import filedialog
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.folder_label.configure(text=path, text_color=("gray10", "gray90"))

    def _refresh_list_ui(self):
        for widget in self._list_item_widgets:
            widget.destroy()
        self._list_item_widgets.clear()
        import os
        for i, path in enumerate(self.files):
            item = ctk.CTkFrame(self.scroll_frame, fg_color=("gray85", "gray25") if i == self._selected_index else "transparent")
            item.pack(fill="x", pady=2)
            lbl = ctk.CTkLabel(item, text=f"{i+1}. {os.path.basename(path)}", anchor="w")
            lbl.pack(side="left", padx=10, fill="x", expand=True)
            for w in [item, lbl]:
                w.bind("<Button-1>", lambda e, idx=i: self._on_item_click(idx))
            self._list_item_widgets.append(item)

    def _on_item_click(self, index: int):
        self._selected_index = index
        self._refresh_list_ui()

    def _move_up(self):
        idx = self._selected_index
        if idx is not None and idx > 0:
            self.files[idx], self.files[idx-1] = self.files[idx-1], self.files[idx]
            self._selected_index = idx - 1
            self._refresh_list_ui()

    def _move_down(self):
        idx = self._selected_index
        if idx is not None and idx < len(self.files) - 1:
            self.files[idx], self.files[idx+1] = self.files[idx+1], self.files[idx]
            self._selected_index = idx + 1
            self._refresh_list_ui()

    def _remove_selected(self):
        idx = self._selected_index
        if idx is not None:
            self.files.pop(idx)
            self._selected_index = None
            self._refresh_list_ui()

    def _start_merge(self):
        if not self.files:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先添加要合併的檔案")
            return

        import os
        out_f = self.output_folder if self.output_folder else os.path.dirname(self.files[0])
        out_n = self.filename_entry.get().strip()
        if not out_n: out_n = "merged_result"

        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)
        
        import pdf_utils
        self.run_in_thread(
            target=pdf_utils.merge_pdfs,
            pdf_list=self.files,
            output_folder=out_f,
            output_filename=out_n,
            callback=self._update_progress,
            on_success=self._on_success,
            on_error=self._on_error
        )

    def _update_progress(self, progress: float):
        self.after(0, lambda: self.progress_bar.set(progress))

    def _on_success(self, result):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(1)
        success, msg = result
        import tkinter.messagebox as messagebox
        if success: messagebox.showinfo("完成", msg)
        else: messagebox.showerror("錯誤", msg)

    def _on_error(self, exc):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(0)
        import tkinter.messagebox as messagebox
        messagebox.showerror("錯誤", f"合併失敗：\n{exc}")



class ConvertPage(BasePage):
    """
    PDF 轉圖片頁面：將 PDF 頁面轉換為高畫質 JPG。
    """

    def build_ui(self):
        self.input_file: str = ""
        self.output_folder: str = ""

        self.grid_columnconfigure(0, weight=1)

        # 1. 標題
        ctk.CTkLabel(self, text="PDF 轉圖片", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, pady=(20, 10), sticky="w", padx=30)

        # 2. 來源檔案
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=1, column=0, padx=30, pady=10, sticky="ew")
        self.file_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(self.file_frame, text="來源檔案:").grid(row=0, column=0, padx=10, pady=10)
        self.file_label = ctk.CTkLabel(self.file_frame, text="尚未選擇檔案...", text_color="gray")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkButton(self.file_frame, text="選擇 PDF", width=100, command=self._select_file).grid(row=0, column=2, padx=10, pady=10)

        # 3. 輸出設定 (目錄 & 檔名)
        self.output_frame = ctk.CTkFrame(self)
        self.output_frame.grid(row=2, column=0, padx=30, pady=10, sticky="ew")
        self.output_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self.output_frame, text="輸出目錄:").grid(row=0, column=0, padx=10, pady=5)
        self.folder_label = ctk.CTkLabel(self.output_frame, text="預設為檔案所在目錄...", text_color="gray")
        self.folder_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkButton(self.output_frame, text="瀏覽...", width=80, command=self._select_folder).grid(row=0, column=2, padx=10, pady=5)

        ctk.CTkLabel(self.output_frame, text="自訂檔名:").grid(row=1, column=0, padx=10, pady=5)
        self.filename_entry = ctk.CTkEntry(self.output_frame, placeholder_text="預設為來源檔名")
        self.filename_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

        # 4. 處理範圍模式 (全部 vs 自訂)
        self.scope_frame = ctk.CTkFrame(self)
        self.scope_frame.grid(row=3, column=0, padx=30, pady=10, sticky="ew")
        ctk.CTkLabel(self.scope_frame, text="處理範圍:").grid(row=0, column=0, padx=10, pady=10)
        
        self.scope_var = ctk.StringVar(value="all")
        self.scope_switch = ctk.CTkSegmentedButton(self.scope_frame, values=["全部頁面", "自訂頁面"], 
                                                  command=self._on_scope_change)
        self.scope_switch.set("全部頁面")
        self.scope_switch.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 5. 頁碼選取區域 (預設隱藏)
        self.page_select_frame = ctk.CTkFrame(self)
        self.page_select_frame.columnconfigure(2, weight=1)

        ctk.CTkButton(self.page_select_frame, text="🔍 視覺化選取", width=120, command=self._open_visual_selector).grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkLabel(self.page_select_frame, text="頁碼範圍:").grid(row=0, column=1, padx=10, pady=10)
        self.range_entry = ctk.CTkEntry(self.page_select_frame, placeholder_text="例如: 1-3, 5", width=250)
        self.range_entry.grid(row=0, column=2, padx=10, pady=10, sticky="w")

        # 6. 轉圖設定 (DPI & Quality)
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=4, column=0, padx=30, pady=10, sticky="ew")
        self.settings_frame.columnconfigure((1, 4), weight=1)

        ctk.CTkLabel(self.settings_frame, text="解析度 (DPI):").grid(row=0, column=0, padx=10, pady=10)
        self.dpi_slider = ctk.CTkSlider(self.settings_frame, from_=100, to=400, number_of_steps=6)
        self.dpi_slider.set(200)
        self.dpi_slider.grid(row=0, column=1, columnspan=3, padx=10, pady=10, sticky="ew")
        self.dpi_val = ctk.CTkLabel(self.settings_frame, text="200")
        self.dpi_val.grid(row=0, column=4, padx=10, pady=10)
        self.dpi_slider.configure(command=lambda v: self.dpi_val.configure(text=str(int(v))))

        ctk.CTkLabel(self.settings_frame, text="圖片品質:").grid(row=1, column=0, padx=10, pady=10)
        self.quality_slider = ctk.CTkSlider(self.settings_frame, from_=50, to=100, number_of_steps=10)
        self.quality_slider.set(85)
        self.quality_slider.grid(row=1, column=1, columnspan=3, padx=10, pady=10, sticky="ew")
        self.quality_val = ctk.CTkLabel(self.settings_frame, text="85%")
        self.quality_val.grid(row=1, column=4, padx=10, pady=10)
        self.quality_slider.configure(command=lambda v: self.quality_val.configure(text=f"{int(v)}%"))

        # 7. 預覽與進度區域
        self.bottom_container = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_container.grid(row=5, column=0, padx=30, pady=10, sticky="nsew")
        self.bottom_container.columnconfigure(0, weight=1)

        self.preview_frame = ctk.CTkFrame(self.bottom_container, width=200, height=240)
        self.preview_frame.grid(row=0, column=0, padx=(0, 20), sticky="n")
        self.preview_frame.grid_propagate(False)
        self.preview_label = ctk.CTkLabel(self.preview_frame, text="等待預覽...", text_color="gray")
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")

        self.controls_frame = ctk.CTkFrame(self.bottom_container, fg_color="transparent")
        self.controls_frame.grid(row=0, column=1, sticky="nsew")

        self.start_btn = ctk.CTkButton(self.controls_frame, text="開始轉換", width=200, height=50, font=ctk.CTkFont(weight="bold"), command=self._start_convert)
        self.start_btn.pack(pady=(0, 20))

        self.info_label = ctk.CTkLabel(self.controls_frame, text="請選擇檔案以查看資訊", text_color="gray")
        self.info_label.pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(self.controls_frame, width=200)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

    def _on_scope_change(self, mode_name):
        if mode_name == "自訂頁面":
            self.page_select_frame.grid(row=4, column=0, padx=30, pady=10, sticky="ew")
            self.settings_frame.grid(row=5, column=0, padx=30, pady=10, sticky="ew")
            self.bottom_container.grid(row=6, column=0, padx=30, pady=10, sticky="nsew")
        else:
            self.page_select_frame.grid_forget()
            self.settings_frame.grid(row=4, column=0, padx=30, pady=10, sticky="ew")
            self.bottom_container.grid(row=5, column=0, padx=30, pady=10, sticky="nsew")

    def reset_state(self):
        self.input_file = ""
        self.output_folder = ""
        self.file_label.configure(text="尚未選擇檔案...", text_color="gray")
        self.folder_label.configure(text="預設為檔案所在目錄...", text_color="gray")
        self.info_label.configure(text="請選擇檔案以查看資訊", text_color="gray")
        self.range_entry.delete(0, "end")
        self.filename_entry.delete(0, "end")
        self.scope_switch.set("全部頁面")
        self._on_scope_change("全部頁面")
        self.progress_bar.set(0)
        self.preview_label.configure(text="等待預覽...", image=None)

    def _select_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if path:
            self.input_file = path
            import os
            import pypdf
            base = os.path.splitext(os.path.basename(path))[0]
            self.file_label.configure(text=os.path.basename(path), text_color=("gray10", "gray90"))
            self.filename_entry.delete(0, "end")
            self.filename_entry.insert(0, base)
            
            try:
                reader = pypdf.PdfReader(path)
                total = len(reader.pages)
                self.info_label.configure(text=f"總頁數：{total} 頁", text_color=("gray10", "gray90"))
                self.range_entry.delete(0, "end")
                self.range_entry.insert(0, f"1-{total}")
            except:
                self.info_label.configure(text="無法讀取頁數資訊", text_color="red")

            self.preview_label.configure(text="正在產生預覽...", image=None)
            self.run_in_thread(self._get_preview_image, path, on_success=self._show_preview)

    def _get_preview_image(self, path):
        from pdf2image import convert_from_path
        from pdf_utils import POPPLER_PATH
        try:
            pages = convert_from_path(path, first_page=1, last_page=1, dpi=72, poppler_path=POPPLER_PATH)
            if pages: return pages[0]
        except: return None

    def _show_preview(self, pil_img):
        if pil_img:
            img_w, img_h = pil_img.size
            ratio = min(180/img_w, 240/img_h)
            new_size = (int(img_w * ratio), int(img_h * ratio))
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=new_size)
            self.preview_label.configure(text="", image=ctk_img)
        else:
            self.preview_label.configure(text="預覽不可用")

    def _select_folder(self):
        from tkinter import filedialog
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.folder_label.configure(text=path, text_color=("gray10", "gray90"))

    def _start_convert(self):
        if not self.input_file:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return

        import os
        out_f = self.output_folder if self.output_folder else os.path.dirname(self.input_file)
        out_n = self.filename_entry.get().strip()
        
        if self.scope_switch.get() == "全部頁面":
            ranges = "" 
        else:
            ranges = self.range_entry.get().strip()
            if not ranges:
                import tkinter.messagebox as messagebox
                messagebox.showwarning("提示", "請輸入頁碼範圍或使用視覺化選取")
                return

        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)

        import pdf_utils
        self.run_in_thread(
            target=pdf_utils.pdf_to_jpg,
            input_path=self.input_file,
            output_folder=out_f,
            dpi=int(self.dpi_slider.get()),
            quality=int(self.quality_slider.get()),
            ranges=ranges,
            custom_name=out_n,
            callback=self._update_progress,
            on_success=self._on_success,
            on_error=self._on_error
        )

    def _update_progress(self, progress: float):
        self.after(0, lambda: self.progress_bar.set(progress))

    def _on_success(self, msg):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(1)
        import tkinter.messagebox as messagebox
        messagebox.showinfo("完成", msg)

    def _on_error(self, exc):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(0)
        import tkinter.messagebox as messagebox
        messagebox.showerror("錯誤", f"轉換失敗：\n{exc}")

    def _open_visual_selector(self):
        if not self.input_file:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return
        VisualPageSelector(self, self.input_file, self._on_pages_selected)

    def _on_pages_selected(self, pages_str: str):
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, pages_str)
        self.info_label.configure(text=f"已選取特定頁面: {pages_str}")

    def _open_visual_selector(self):
        """開啟視覺化分頁選取器"""
        if not self.input_file:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return
        
        selector = VisualPageSelector(self, self.input_file, self._on_pages_selected)

    def _on_pages_selected(self, pages_str: str):
        """當使用者在視覺選取器完成選取後的回呼"""
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, pages_str)
        self.info_label.configure(text=f"已選取特定頁面: {pages_str}")


class VisualPageSelector(ctk.CTkToplevel):
    """
    彈出式視覺化分頁選取器。
    """
    def __init__(self, parent, pdf_path, on_selected_callback):
        super().__init__(parent)
        self.title("視覺化分頁選取")
        self.geometry("900x700")
        self.attributes("-topmost", True)
        
        self.pdf_path = pdf_path
        self.on_selected = on_selected_callback
        self.selected_vars = {} # page_idx: BooleanVar
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 1. 頂部工具列
        self.toolbar = ctk.CTkFrame(self, height=50)
        self.toolbar.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        ctk.CTkButton(self.toolbar, text="全選", width=80, command=self._select_all).pack(side="left", padx=5)
        ctk.CTkButton(self.toolbar, text="全不選", width=80, command=self._deselect_all).pack(side="left", padx=5)
        
        self.info_label = ctk.CTkLabel(self.toolbar, text="正在載入預覽圖...")
        self.info_label.pack(side="left", padx=20)
        ctk.CTkButton(self.toolbar, text="確認選取", fg_color="green", hover_color="darkgreen", 
                      command=self._confirm).pack(side="right", padx=5)

        # 2. 滾動網格區域
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        # 設定網格列數
        self.columns = 4
        for i in range(self.columns):
            self.scroll_frame.grid_columnconfigure(i, weight=1)

        # 3. 啟動載入
        self._load_pages()

    def _load_pages(self):
        import pypdf
        try:
            reader = pypdf.PdfReader(self.pdf_path)
            self.total_pages = len(reader.pages)
            self.info_label.configure(text=f"總計 {self.total_pages} 頁，請勾選要處理的頁面")
            
            # 建立所有 Checkbox 變數
            for i in range(self.total_pages):
                self.selected_vars[i] = ctk.BooleanVar(value=False)
            
            # 分批產生縮圖以免卡死
            self.after(100, lambda: self._render_batch(0))
        except Exception as e:
            self.info_label.configure(text=f"載入失敗: {e}", text_color="red")

    def _render_batch(self, start_idx):
        if start_idx >= self.total_pages:
            return
        
        end_idx = min(start_idx + 4, self.total_pages) # 一次轉 4 頁
        
        from pdf2image import convert_from_path
        from pdf_utils import POPPLER_PATH
        from PIL import Image
        
        try:
            # 批量轉換
            pages = convert_from_path(self.pdf_path, first_page=start_idx+1, last_page=end_idx, 
                                    dpi=50, poppler_path=POPPLER_PATH)
            
            for i, pil_img in enumerate(pages):
                page_idx = start_idx + i
                
                # 建立容器
                item_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
                row = page_idx // self.columns
                col = page_idx % self.columns
                item_frame.grid(row=row, column=col, padx=10, pady=10)
                
                # 縮圖顯示
                img_w, img_h = pil_img.size
                ratio = min(150/img_w, 200/img_h)
                new_size = (int(img_w * ratio), int(img_h * ratio))
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=new_size)
                
                img_label = ctk.CTkLabel(item_frame, text="", image=ctk_img)
                img_label.pack()
                
                # 勾選框
                cb = ctk.CTkCheckBox(item_frame, text=f"第 {page_idx+1} 頁", variable=self.selected_vars[page_idx])
                cb.pack(pady=5)

            # 繼續下一批
            self.after(10, lambda: self._render_batch(end_idx))
        except:
            self.info_label.configure(text="部分縮圖載入失敗 (Poppler 問題)", text_color="orange")

    def _select_all(self):
        for v in self.selected_vars.values(): v.set(True)

    def _deselect_all(self):
        for v in self.selected_vars.values(): v.set(False)

    def _confirm(self):
        selected = [i+1 for i, v in self.selected_vars.items() if v.get()]
        if not selected:
            self.destroy()
            return
        
        # 轉換成範圍字串，例如 [1,2,3,5] -> "1-3, 5"
        pages_str = self._to_range_str(selected)
        self.on_selected(pages_str)
        self.destroy()

    def _to_range_str(self, nums):
        if not nums: return ""
        nums = sorted(list(set(nums)))
        ranges = []
        if not nums: return ""
        
        start = nums[0]
        end = nums[0]
        
        for i in range(1, len(nums)):
            if nums[i] == end + 1:
                end = nums[i]
            else:
                ranges.append(f"{start}-{end}" if start != end else f"{start}")
                start = nums[i]
                end = nums[i]
        ranges.append(f"{start}-{end}" if start != end else f"{start}")
        return ", ".join(ranges)


class SplitPage(BasePage):
    """
    PDF 拆分頁面：支援一頁一檔或多頁合併。
    """

    def build_ui(self):
        self.input_file: str = ""
        self.output_folder: str = ""

        self.grid_columnconfigure(0, weight=1)

        # 1. 標題
        ctk.CTkLabel(self, text="PDF 拆分", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, pady=(20, 10), sticky="w", padx=30)

        # 2. 來源檔案
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=1, column=0, padx=30, pady=10, sticky="ew")
        self.file_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(self.file_frame, text="來源檔案:").grid(row=0, column=0, padx=10, pady=10)
        self.file_label = ctk.CTkLabel(self.file_frame, text="尚未選擇檔案...", text_color="gray")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkButton(self.file_frame, text="選擇 PDF", width=100, command=self._select_file).grid(row=0, column=2, padx=10, pady=10)

        # 3. 輸出設定 (目錄 & 檔名)
        self.output_frame = ctk.CTkFrame(self)
        self.output_frame.grid(row=2, column=0, padx=30, pady=10, sticky="ew")
        self.output_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self.output_frame, text="輸出目錄:").grid(row=0, column=0, padx=10, pady=5)
        self.folder_label = ctk.CTkLabel(self.output_frame, text="預設為檔案所在目錄...", text_color="gray")
        self.folder_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkButton(self.output_frame, text="瀏覽...", width=80, command=self._select_folder).grid(row=0, column=2, padx=10, pady=5)

        ctk.CTkLabel(self.output_frame, text="自訂檔名:").grid(row=1, column=0, padx=10, pady=5)
        self.filename_entry = ctk.CTkEntry(self.output_frame, placeholder_text="預設為來源檔名")
        self.filename_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

        # 4. 處理範圍模式 (全部 vs 自訂)
        self.scope_frame = ctk.CTkFrame(self)
        self.scope_frame.grid(row=3, column=0, padx=30, pady=10, sticky="ew")
        ctk.CTkLabel(self.scope_frame, text="處理範圍:").grid(row=0, column=0, padx=10, pady=10)
        
        self.scope_var = ctk.StringVar(value="all")
        self.scope_switch = ctk.CTkSegmentedButton(self.scope_frame, values=["全部頁面", "自訂頁面"], 
                                                  command=self._on_scope_change)
        self.scope_switch.set("全部頁面")
        self.scope_switch.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 5. 頁碼選取區域 (預設隱藏)
        self.page_select_frame = ctk.CTkFrame(self)
        self.page_select_frame.columnconfigure(2, weight=1)

        ctk.CTkButton(self.page_select_frame, text="🔍 視覺化選取", width=120, command=self._open_visual_selector).grid(row=0, column=0, padx=10, pady=10)
        ctk.CTkLabel(self.page_select_frame, text="頁碼範圍:").grid(row=0, column=1, padx=10, pady=10)
        self.range_entry = ctk.CTkEntry(self.page_select_frame, placeholder_text="例如: 1-3, 5", width=250)
        self.range_entry.grid(row=0, column=2, padx=10, pady=10, sticky="w")

        # 6. 拆分設定 (模式)
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=4, column=0, padx=30, pady=10, sticky="ew")
        self.settings_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self.settings_frame, text="拆分模式:").grid(row=0, column=0, padx=10, pady=10)
        self.mode_var = ctk.StringVar(value="single")
        self.mode_switch = ctk.CTkSegmentedButton(self.settings_frame, values=["一頁一檔案", "多頁合成一檔"], 
                                                 command=self._on_mode_change_internal)
        self.mode_switch.set("一頁一檔案")
        self.mode_switch.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 7. 預覽與進度區域
        self.bottom_container = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_container.grid(row=5, column=0, padx=30, pady=10, sticky="nsew")
        self.bottom_container.columnconfigure(0, weight=1)

        self.preview_frame = ctk.CTkFrame(self.bottom_container, width=160, height=220)
        self.preview_frame.grid(row=0, column=0, padx=(0, 20), sticky="n")
        self.preview_frame.grid_propagate(False)
        self.preview_label = ctk.CTkLabel(self.preview_frame, text="等待預覽...", text_color="gray")
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")

        self.controls_frame = ctk.CTkFrame(self.bottom_container, fg_color="transparent")
        self.controls_frame.grid(row=0, column=1, sticky="nsew")

        self.start_btn = ctk.CTkButton(self.controls_frame, text="開始拆分", width=200, height=50, font=ctk.CTkFont(weight="bold"), command=self._start_split)
        self.start_btn.pack(pady=(0, 20))

        self.info_label = ctk.CTkLabel(self.controls_frame, text="請選擇檔案以查看資訊", text_color="gray")
        self.info_label.pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(self.controls_frame, width=200)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

    def _on_scope_change(self, mode_name):
        if mode_name == "自訂頁面":
            self.page_select_frame.grid(row=4, column=0, padx=30, pady=10, sticky="ew")
            self.settings_frame.grid(row=5, column=0, padx=30, pady=10, sticky="ew")
            self.bottom_container.grid(row=6, column=0, padx=30, pady=10, sticky="nsew")
        else:
            self.page_select_frame.grid_forget()
            self.settings_frame.grid(row=4, column=0, padx=30, pady=10, sticky="ew")
            self.bottom_container.grid(row=5, column=0, padx=30, pady=10, sticky="nsew")

    def reset_state(self):
        self.input_file = ""
        self.output_folder = ""
        self.file_label.configure(text="尚未選擇檔案...", text_color="gray")
        self.folder_label.configure(text="預設為檔案所在目錄...", text_color="gray")
        self.info_label.configure(text="請選擇檔案以查看資訊", text_color="gray")
        self.range_entry.delete(0, "end")
        self.filename_entry.delete(0, "end")
        self.scope_switch.set("全部頁面")
        self._on_scope_change("全部頁面")
        self.mode_switch.set("一頁一檔案")
        self.mode_var.set("single")
        self.progress_bar.set(0)
        self.preview_label.configure(text="等待預覽...", image=None)

    def _on_mode_change_internal(self, mode_name):
        self.mode_var.set("range" if mode_name == "多頁合成一檔" else "single")

    def _select_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if path:
            self.input_file = path
            import os
            import pypdf
            base = os.path.splitext(os.path.basename(path))[0]
            self.file_label.configure(text=os.path.basename(path), text_color=("gray10", "gray90"))
            self.filename_entry.delete(0, "end")
            self.filename_entry.insert(0, base)
            
            try:
                reader = pypdf.PdfReader(path)
                total = len(reader.pages)
                self.info_label.configure(text=f"總頁數：{total} 頁", text_color=("gray10", "gray90"))
            except:
                self.info_label.configure(text="無法讀取頁數資訊", text_color="red")

            self.preview_label.configure(text="正在產生預覽...", image=None)
            self.run_in_thread(self._get_preview_image, path, on_success=self._show_preview)

    def _get_preview_image(self, path):
        from pdf2image import convert_from_path
        from pdf_utils import POPPLER_PATH
        try:
            pages = convert_from_path(path, first_page=1, last_page=1, dpi=72, poppler_path=POPPLER_PATH)
            if pages: return pages[0]
        except: return None

    def _show_preview(self, pil_img):
        if pil_img:
            img_w, img_h = pil_img.size
            ratio = min(140/img_w, 200/img_h)
            new_size = (int(img_w * ratio), int(img_h * ratio))
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=new_size)
            self.preview_label.configure(text="", image=ctk_img)
        else:
            self.preview_label.configure(text="預覽不可用")

    def _select_folder(self):
        from tkinter import filedialog
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.folder_label.configure(text=path, text_color=("gray10", "gray90"))

    def _start_split(self):
        if not self.input_file:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return

        import os
        out_f = self.output_folder if self.output_folder else os.path.dirname(self.input_file)
        out_n = self.filename_entry.get().strip()
        mode = self.mode_var.get()
        
        if self.scope_switch.get() == "全部頁面":
            ranges = "" 
        else:
            ranges = self.range_entry.get().strip()
            if not ranges:
                import tkinter.messagebox as messagebox
                messagebox.showwarning("提示", "請輸入頁碼範圍或使用視覺化選取")
                return

        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)

        import pdf_utils
        self.run_in_thread(
            target=pdf_utils.split_pdf,
            input_path=self.input_file,
            output_folder=out_f,
            mode=mode,
            ranges=ranges,
            custom_name=out_n,
            callback=self._update_progress,
            on_success=self._on_success,
            on_error=self._on_error
        )

    def _update_progress(self, progress: float):
        self.after(0, lambda: self.progress_bar.set(progress))

    def _on_success(self, msg):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(1)
        import tkinter.messagebox as messagebox
        messagebox.showinfo("完成", msg)

    def _on_error(self, exc):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(0)
        import tkinter.messagebox as messagebox
        messagebox.showerror("錯誤", f"拆分失敗：\n{exc}")

    def _open_visual_selector(self):
        if not self.input_file:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return
        VisualPageSelector(self, self.input_file, self._on_pages_selected)

    def _on_pages_selected(self, pages_str: str):
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, pages_str)
        self.info_label.configure(text=f"已選取特定頁面: {pages_str}")


class CompressPage(BasePage):
    """
    PDF 壓縮頁面：縮減 PDF 檔案大小。
    """

    def build_ui(self):
        self.input_file: str = ""
        self.output_folder: str = ""

        self.grid_columnconfigure(0, weight=1)

        # 1. 標題
        ctk.CTkLabel(self, text="PDF 壓縮", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, pady=(20, 10), sticky="w", padx=30)

        # 2. 來源檔案
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=1, column=0, padx=30, pady=10, sticky="ew")
        self.file_frame.columnconfigure(1, weight=1)
        ctk.CTkLabel(self.file_frame, text="來源檔案:").grid(row=0, column=0, padx=10, pady=10)
        self.file_label = ctk.CTkLabel(self.file_frame, text="尚未選擇檔案...", text_color="gray")
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkButton(self.file_frame, text="選擇 PDF", width=100, command=self._select_file).grid(row=0, column=2, padx=10, pady=10)

        # 3. 輸出設定 (目錄 & 檔名)
        self.output_frame = ctk.CTkFrame(self)
        self.output_frame.grid(row=2, column=0, padx=30, pady=10, sticky="ew")
        self.output_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self.output_frame, text="輸出目錄:").grid(row=0, column=0, padx=10, pady=5)
        self.folder_label = ctk.CTkLabel(self.output_frame, text="預設為檔案所在目錄...", text_color="gray")
        self.folder_label.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        ctk.CTkButton(self.output_frame, text="瀏覽...", width=80, command=self._select_folder).grid(row=0, column=2, padx=10, pady=5)

        ctk.CTkLabel(self.output_frame, text="自訂檔名:").grid(row=1, column=0, padx=10, pady=5)
        self.filename_entry = ctk.CTkEntry(self.output_frame, placeholder_text="預設為來源檔名_compressed")
        self.filename_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=5, sticky="ew")

        # 4. 壓縮強度
        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=3, column=0, padx=30, pady=10, sticky="ew")
        self.settings_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(self.settings_frame, text="壓縮強度:").grid(row=0, column=0, padx=10, pady=10)
        self.quality_var = ctk.StringVar(value="medium")
        self.quality_switch = ctk.CTkSegmentedButton(
            self.settings_frame, 
            values=["基本壓縮", "建議壓縮", "極致壓縮"],
            command=self._on_quality_change
        )
        self.quality_switch.set("建議壓縮")
        self.quality_switch.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # 預估大小顯示
        self.est_label = ctk.CTkLabel(self.settings_frame, text="", text_color="#3B8ED0", font=ctk.CTkFont(size=12, slant="italic"))
        self.est_label.grid(row=1, column=1, padx=10, pady=(0, 10), sticky="w")

        # 5. 預覽與進度區域
        self.bottom_container = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_container.grid(row=4, column=0, padx=30, pady=20, sticky="nsew")
        self.bottom_container.columnconfigure(0, weight=1)

        self.preview_frame = ctk.CTkFrame(self.bottom_container, width=160, height=220)
        self.preview_frame.grid(row=0, column=0, padx=(0, 20), sticky="n")
        self.preview_frame.grid_propagate(False)
        self.preview_label = ctk.CTkLabel(self.preview_frame, text="等待預覽...", text_color="gray")
        self.preview_label.place(relx=0.5, rely=0.5, anchor="center")

        self.controls_frame = ctk.CTkFrame(self.bottom_container, fg_color="transparent")
        self.controls_frame.grid(row=0, column=1, sticky="nsew")

        self.start_btn = ctk.CTkButton(self.controls_frame, text="開始壓縮", width=200, height=50, font=ctk.CTkFont(weight="bold"), command=self._start_compress)
        self.start_btn.pack(pady=(0, 20))

        self.info_label = ctk.CTkLabel(self.controls_frame, text="請選擇檔案以查看資訊", text_color="gray")
        self.info_label.pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(self.controls_frame, width=200)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

    def _on_quality_change(self, val):
        mapping = {"基本壓縮": "low", "建議壓縮": "medium", "極致壓縮": "high"}
        self.quality_var.set(mapping.get(val, "medium"))
        self._update_estimate_display()

    def _update_estimate_display(self):
        """更新 UI 上的預估大小文字"""
        if not hasattr(self, "orig_size_kb") or self.orig_size_kb <= 0:
            self.est_label.configure(text="")
            return

        q = self.quality_var.get()
        # 預估比例 (僅供參考，實際取決於 PDF 內容)
        ratios = {"low": 0.90, "medium": 0.60, "high": 0.35}
        ratio = ratios.get(q, 0.60)
        est_kb = self.orig_size_kb * ratio
        
        saved_pct = (1 - ratio) * 100
        if est_kb > 1024:
            size_str = f"{est_kb/1024:.2f} MB"
        else:
            size_str = f"{est_kb:.1f} KB"
            
        self.est_label.configure(text=f"✨ 預計大小：約 {size_str} (節省約 {saved_pct:.0f}%)")

    def reset_state(self):
        self.input_file = ""
        self.output_folder = ""
        self.orig_size_kb = 0
        self.file_label.configure(text="尚未選擇檔案...", text_color="gray")
        self.folder_label.configure(text="預設為檔案所在目錄...", text_color="gray")
        self.info_label.configure(text="請選擇檔案以查看資訊", text_color="gray")
        self.filename_entry.delete(0, "end")
        self.quality_switch.set("建議壓縮")
        self.quality_var.set("medium")
        self.est_label.configure(text="")
        self.progress_bar.set(0)
        self.preview_label.configure(text="等待預覽...", image=None)

    def _select_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(filetypes=[("PDF 檔案", "*.pdf")])
        if path:
            self.input_file = path
            import os
            base = os.path.splitext(os.path.basename(path))[0]
            self.file_label.configure(text=os.path.basename(path), text_color=("gray10", "gray90"))
            self.filename_entry.delete(0, "end")
            self.filename_entry.insert(0, f"{base}_compressed")
            
            import pypdf
            try:
                reader = pypdf.PdfReader(path)
                size_bytes = os.path.getsize(path)
                self.orig_size_kb = size_bytes / 1024
                size_mb = self.orig_size_kb / 1024
                self.info_label.configure(text=f"總頁數：{len(reader.pages)} 頁 | 大小：{size_mb:.2f} MB", text_color=("gray10", "gray90"))
                self._update_estimate_display()
            except:
                self.info_label.configure(text="無法讀取檔案資訊", text_color="red")

            self.preview_label.configure(text="正在產生預覽...", image=None)
            self.run_in_thread(self._get_preview_image, path, on_success=self._show_preview)

    def _get_preview_image(self, path):
        from pdf2image import convert_from_path
        from pdf_utils import POPPLER_PATH
        try:
            pages = convert_from_path(path, first_page=1, last_page=1, dpi=72, poppler_path=POPPLER_PATH)
            if pages: return pages[0]
        except: return None

    def _show_preview(self, pil_img):
        if pil_img:
            img_w, img_h = pil_img.size
            ratio = min(140/img_w, 200/img_h)
            new_size = (int(img_w * ratio), int(img_h * ratio))
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=new_size)
            self.preview_label.configure(text="", image=ctk_img)
        else:
            self.preview_label.configure(text="預覽不可用")

    def _select_folder(self):
        from tkinter import filedialog
        path = filedialog.askdirectory()
        if path:
            self.output_folder = path
            self.folder_label.configure(text=path, text_color=("gray10", "gray90"))

    def _start_compress(self):
        if not self.input_file:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("提示", "請先選擇 PDF 檔案")
            return

        import os
        out_f = self.output_folder if self.output_folder else os.path.dirname(self.input_file)
        out_n = self.filename_entry.get().strip()
        
        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)

        import pdf_utils
        self.run_in_thread(
            target=pdf_utils.compress_pdf,
            input_path=self.input_file,
            output_folder=out_f,
            quality=self.quality_var.get(),
            custom_name=out_n,
            callback=self._update_progress,
            on_success=self._on_success,
            on_error=self._on_error
        )

    def _update_progress(self, progress: float):
        self.after(0, lambda: self.progress_bar.set(progress))

    def _on_success(self, result):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(1)
        success, path_or_msg = result
        import tkinter.messagebox as messagebox
        if success:
            import os
            try:
                old_size = os.path.getsize(self.input_file) / 1024
                new_size = os.path.getsize(path_or_msg) / 1024
                ratio = (1 - new_size/old_size) * 100
                final_msg = f"壓縮成功！檔案已儲存至：\n{path_or_msg}\n\n原始大小: {old_size:.1f} KB\n壓縮後大小: {new_size:.1f} KB\n實際壓縮率: {ratio:.1f}%"
                messagebox.showinfo("完成", final_msg)
            except:
                messagebox.showinfo("完成", f"壓縮成功！已儲存至 {path_or_msg}")
        else:
            messagebox.showerror("錯誤", path_or_msg)

    def _on_error(self, exc):
        self.start_btn.configure(state="normal")
        self.progress_bar.set(0)
        import tkinter.messagebox as messagebox
        messagebox.showerror("錯誤", f"壓縮失敗：\n{exc}")


class WatermarkPage(BasePage):
    """PDF 浮水印頁面（待實作）"""

    def build_ui(self):
        ctk.CTkLabel(self, text="加入浮水印", font=ctk.CTkFont(size=28, weight="bold")).pack(pady=30)
        ctk.CTkLabel(self, text="文字或圖片浮水印", text_color="gray").pack()


class SettingsPage(BasePage):
    """設定頁面（待實作）"""

    def build_ui(self):
        ctk.CTkLabel(self, text="設定", font=ctk.CTkFont(size=28, weight="bold")).pack(pady=30)
        # 外觀模式切換範例
        ctk.CTkLabel(self, text="外觀模式").pack()
        self.mode_menu = ctk.CTkOptionMenu(
            self,
            values=["dark", "light", "system"],
            command=lambda m: ctk.set_appearance_mode(m),
        )
        self.mode_menu.pack(pady=8)


# ═══════════════════════════════════════════════════════════════
# Sidebar — 左側導覽列
# ═══════════════════════════════════════════════════════════════

# 導覽項目：(顯示名稱, 對應 PDFApp 中的 page key)
NAV_ITEMS: list[tuple[str, str]] = [
    ("📄  PDF 合併",   "merge"),
    ("🖼️  PDF 轉圖",   "convert"),
    ("✂️  PDF 拆分",   "split"),
    ("🗜️  PDF 壓縮",   "compress"),
]


class Sidebar(ctk.CTkFrame):
    """
    左側固定導覽列。

    Args:
        parent  : 父容器
        on_nav  : 切換頁面的回呼函式，接收 page_key: str
    """

    def __init__(self, parent, on_nav: Callable[[str], None], **kwargs):
        super().__init__(parent, width=SIDEBAR_W, corner_radius=0, **kwargs)
        self.pack_propagate(False)
        self._on_nav      = on_nav
        self._nav_buttons : dict[str, ctk.CTkButton] = {}
        self._active_key  : str | None = None

        self._build()

    # ── 建立 UI ──────────────────────────────────────────────

    def _build(self):
        # 應用程式 Logo / 標題
        logo = ctk.CTkLabel(
            self,
            text="CYT\nPDF 工具",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        logo.pack(pady=(24, 16))

        ctk.CTkFrame(self, height=1, fg_color="gray30").pack(fill="x", padx=12, pady=4)

        # 導覽按鈕
        for label, key in NAV_ITEMS:
            btn = ctk.CTkButton(
                self,
                text=label,
                anchor="w",
                height=42,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray80", "gray25"),
                corner_radius=8,
                command=lambda k=key: self._on_nav(k),
            )
            btn.pack(fill="x", padx=10, pady=3)
            self._nav_buttons[key] = btn

        # 底部：設定按鈕
        self.pack(side="left", fill="y")
        settings_btn = ctk.CTkButton(
            self,
            text="⚙️  設定",
            anchor="w",
            height=42,
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray80", "gray25"),
            corner_radius=8,
            command=lambda: self._on_nav("settings"),
        )
        settings_btn.pack(side="bottom", fill="x", padx=10, pady=(4, 16))
        self._nav_buttons["settings"] = settings_btn

    # ── 公開 API ─────────────────────────────────────────────

    def set_active(self, key: str):
        """更新按鈕高亮狀態，標示目前所在頁面。"""
        if self._active_key and self._active_key in self._nav_buttons:
            self._nav_buttons[self._active_key].configure(fg_color="transparent")

        if key in self._nav_buttons:
            self._nav_buttons[key].configure(fg_color=("gray75", "gray30"))

        self._active_key = key


# ═══════════════════════════════════════════════════════════════
# PDFApp — 主視窗
# ═══════════════════════════════════════════════════════════════

class PDFApp(ctk.CTk):
    """
    PDF 工具主視窗。

    頁面切換邏輯：
        - self._pages  : dict[str, BasePage]  所有已建立的頁面實例
        - self._current : str | None          目前顯示的頁面 key
        - navigate(key) : 切換頁面的統一入口
    """

    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{APP_WIDTH}x{APP_HEIGHT}")
        self.minsize(800, 500)

        self._pages:   dict[str, BasePage] = {}
        self._current: str | None = None

        self._build_layout()
        self._register_pages()

        # 預設顯示第一個頁面
        self.navigate("merge")

    # ── 版面建立 ─────────────────────────────────────────────

    def _build_layout(self):
        # 側邊欄
        self.sidebar = Sidebar(self, on_nav=self.navigate)

        # 右側內容區（所有 Page 都放在這裡，透過 pack/pack_forget 切換）
        self.content_area = ctk.CTkFrame(self, corner_radius=0)
        self.content_area.pack(side="left", fill="both", expand=True)

    def _register_pages(self):
        """
        在這裡集中建立所有頁面實例並存入 self._pages。
        頁面初始時全部隱藏（不 pack），由 navigate() 控制顯示。
        """
        page_classes: dict[str, type[BasePage]] = {
            "merge":     MergePage,
            "convert":   ConvertPage,
            "split":     SplitPage,
            "compress":  CompressPage,
            "settings":  SettingsPage,
        }
        for key, cls in page_classes.items():
            page = cls(self.content_area, app=self, fg_color="transparent")
            self._pages[key] = page

    # ── 頁面切換 ─────────────────────────────────────────────

    def navigate(self, key: str):
        """
        切換到指定頁面。

        Args:
            key : 頁面識別碼，對應 _pages 中的鍵值
        """
        if key not in self._pages:
            print(f"[PDFApp] 未知的頁面 key：{key}")
            return

        if key == self._current:
            return  # 已在此頁面，不重複切換

        # 隱藏目前頁面
        if self._current and self._current in self._pages:
            self._pages[self._current].pack_forget()
            self._pages[self._current].on_hide()

        # 顯示新頁面
        self._pages[key].pack(fill="both", expand=True)
        self._pages[key].on_show()
        self._current = key

        # 更新側邊欄高亮
        self.sidebar.set_active(key)

    # ── 工具方法 ─────────────────────────────────────────────

    def show_status(self, message: str, duration_ms: int = 3000):
        """
        （預留）在狀態列顯示訊息，可在未來加入底部 StatusBar 後實作。
        """
        print(f"[Status] {message}")
        # TODO: self.statusbar.set(message)


# ═══════════════════════════════════════════════════════════════
# 程式進入點
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = PDFApp()
    app.mainloop()
