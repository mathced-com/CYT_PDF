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
APP_VERSION = "1.0.4"
GITHUB_REPO = "ced-cyt/CYT_PDF" # 請根據實際 GitHub 帳號修改



# ─────────────────────────────────────────────
# 全域設定
# ─────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

APP_TITLE  = "CYT PDF 工具"
APP_WIDTH  = 1100
APP_HEIGHT = 680
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
            # 注意：self 必須是 CTk widget，才有 after()
            if hasattr(self, "after"):
                self.after(interval_ms, self._poll_result, on_success, on_error, interval_ms)


# ═══════════════════════════════════════════════════════════════
# BasePage — 所有頁面的基礎類別
# ═══════════════════════════════════════════════════════════════

class BasePage(ctk.CTkFrame, ThreadedTaskMixin):
    """
    所有功能頁面的基礎類別。
    繼承 CTkFrame（作為容器）與 ThreadedTaskMixin（後台執行緒支援）。

    子類別只需實作：
        build_ui(self) -> None
    """

    def __init__(self, parent: ctk.CTkFrame, app: "PDFApp", **kwargs):
        ctk.CTkFrame.__init__(self, parent, corner_radius=0, **kwargs)
        ThreadedTaskMixin.__init__(self)
        self.app = app          # 可透過 self.app 存取主視窗
        self.build_ui()

    def build_ui(self) -> None:
        """子類別覆寫此方法，建立頁面 UI。"""
        raise NotImplementedError

    def on_show(self) -> None:
        """每次頁面被切換顯示時觸發，子類別可選擇性覆寫。"""

    def on_hide(self) -> None:
        """每次頁面被隱藏時觸發，子類別可選擇性覆寫。"""


# ═══════════════════════════════════════════════════════════════
# 功能頁面（佔位符）
# ═══════════════════════════════════════════════════════════════

class MergePage(BasePage):
    """
    PDF 合併頁面：支援檔案選取、清單排序、刪除與進度顯示。
    """

    def build_ui(self):
        self.files: list[str] = []
        self._selected_index: int | None = None
        self._list_item_widgets: list[ctk.CTkFrame] = []

        # 設定頁面內邊距
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # 讓清單區域佔據剩餘空間

        # 1. 標題
        self.title_label = ctk.CTkLabel(
            self, text="PDF 合併", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.grid(row=0, column=0, columnspan=2, pady=(20, 10), sticky="w", padx=30)

        # 2. 拖放區域 (模擬)
        self.drop_zone = ctk.CTkFrame(self, height=120, border_width=2, border_color=("gray70", "gray30"), dash=(10, 5))
        self.drop_zone.grid(row=1, column=0, columnspan=2, padx=30, pady=10, sticky="nsew")
        self.drop_zone.grid_propagate(False)
        self.drop_zone.columnconfigure(0, weight=1)
        self.drop_zone.rowconfigure(0, weight=1)

        self.drop_label = ctk.CTkLabel(
            self.drop_zone, 
            text="點擊此處 或 拖曳 PDF 檔案至此",
            text_color=("gray40", "gray60"),
            font=ctk.CTkFont(size=14)
        )
        self.drop_label.grid(row=0, column=0)
        
        # 點擊觸發檔案選擇
        self.drop_zone.bind("<Button-1>", lambda e: self._select_files())
        self.drop_label.bind("<Button-1>", lambda e: self._select_files())

        # 3. 檔案清單區域 (清單 + 側邊按鈕)
        self.list_container = ctk.CTkFrame(self, fg_color="transparent")
        self.list_container.grid(row=2, column=0, columnspan=2, padx=30, pady=10, sticky="nsew")
        self.list_container.columnconfigure(0, weight=1)
        self.list_container.rowconfigure(0, weight=1)

        # 檔案清單 (Scrollable Frame)
        self.scroll_frame = ctk.CTkScrollableFrame(self.list_container, label_text="待合併檔案清單")
        self.scroll_frame.grid(row=0, column=0, sticky="nsew")

        # 側邊按鈕欄
        self.action_bar = ctk.CTkFrame(self.list_container, fg_color="transparent")
        self.action_bar.grid(row=0, column=1, padx=(10, 0), sticky="ns")

        btn_style = {"width": 40, "font": ctk.CTkFont(size=16)}
        
        self.up_btn = ctk.CTkButton(self.action_bar, text="▲", command=self._move_up, **btn_style)
        self.up_btn.pack(pady=5)
        
        self.down_btn = ctk.CTkButton(self.action_bar, text="▼", command=self._move_down, **btn_style)
        self.down_btn.pack(pady=5)
        
        self.remove_btn = ctk.CTkButton(self.action_bar, text="✕", fg_color="#E74C3C", hover_color="#C0392B", command=self._remove_selected, **btn_style)
        self.remove_btn.pack(pady=5)

        # 4. 底部區域 (進度條 + 開始按鈕)
        self.footer = ctk.CTkFrame(self, fg_color="transparent")
        self.footer.grid(row=3, column=0, columnspan=2, padx=30, pady=(10, 20), sticky="ew")
        self.footer.columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(self.footer)
        self.progress_bar.grid(row=0, column=0, padx=(0, 20), sticky="ew")
        self.progress_bar.set(0)

        self.start_btn = ctk.CTkButton(
            self.footer, text="開始合併", 
            width=120, height=40, font=ctk.CTkFont(weight="bold"),
            command=self._start_merge_process
        )
        self.start_btn.grid(row=0, column=1)

    # ── 邏輯實作 ──────────────────────────────────────────────

    def _select_files(self):
        """開啟對話框選取檔案。"""
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(
            title="選擇 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf")]
        )
        if paths:
            self.files.extend(list(paths))
            self._refresh_list_ui()

    def _refresh_list_ui(self):
        """根據 self.files 重新繪製清單。"""
        # 清除舊項
        for widget in self._list_item_widgets:
            widget.destroy()
        self._list_item_widgets.clear()

        import os
        for i, path in enumerate(self.files):
            filename = os.path.basename(path)
            
            # 建立每一列的容器
            item = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            item.pack(fill="x", pady=2)
            
            # 檔案名稱
            lbl = ctk.CTkLabel(item, text=f"{i+1}. {filename}", anchor="w")
            lbl.pack(side="left", padx=10, fill="x", expand=True)
            
            # 點擊選取邏輯
            bg_color = ("gray85", "gray25") if i == self._selected_index else "transparent"
            item.configure(fg_color=bg_color)
            
            # 綁定選取事件 (包含子元件)
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

    def _start_merge_process(self):
        """啟動合併流程：詢問存檔位置並交由背景執行緒處理。"""
        if not self.files:
            return
            
        from tkinter import filedialog
        output_path = filedialog.asksaveasfilename(
            title="儲存合併後的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf")]
        )
        if not output_path:
            return
            
        self._output_path = output_path
        self._run_merge_task()
        
    def _run_merge_task(self):
        self.start_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        
        # 準備密碼字典，若是先前已輸入過的就繼續用
        if not hasattr(self, "_pdf_passwords"):
            self._pdf_passwords = {}
            
        import pdf_utils
        self.run_in_thread(
            target=pdf_utils.merge_pdfs,
            input_paths=self.files,
            output_path=self._output_path,
            passwords=self._pdf_passwords,
            on_success=self._on_merge_success,
            on_error=self._on_merge_error
        )

    def _on_merge_success(self, result):
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1)
        self.start_btn.configure(state="normal")
        
        success, msg = result
        import tkinter.messagebox as messagebox
        if success:
            messagebox.showinfo("合併完成", msg)
        else:
            messagebox.showerror("合併失敗", msg)

    def _on_merge_error(self, exc: Exception):
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(0)
        self.start_btn.configure(state="normal")
        
        import pdf_utils
        if isinstance(exc, pdf_utils.EncryptedPDFError):
            # 彈出視窗要求輸入密碼
            import os
            filename = os.path.basename(exc.filepath)
            dialog = ctk.CTkInputDialog(
                text=f"檔案受密碼保護，請輸入密碼解鎖：\n{filename}", 
                title="需要密碼"
            )
            pwd = dialog.get_input()
            if pwd: # 如果使用者輸入了密碼
                self._pdf_passwords[exc.filepath] = pwd
                # 重新嘗試 (不會再問存檔位置)
                self._run_merge_task()
            else:
                import tkinter.messagebox as messagebox
                messagebox.showwarning("已取消", "因缺少密碼，合併已取消。")
        else:
            import tkinter.messagebox as messagebox
            messagebox.showerror("錯誤", f"合併發生異常：\n{exc}")



class SplitPage(BasePage):
    """PDF 拆分頁面（待實作）"""

    def build_ui(self):
        ctk.CTkLabel(self, text="PDF 拆分", font=ctk.CTkFont(size=28, weight="bold")).pack(pady=30)
        ctk.CTkLabel(self, text="選擇頁碼範圍進行拆分", text_color="gray").pack()


class CompressPage(BasePage):
    """PDF 壓縮頁面（待實作）"""

    def build_ui(self):
        ctk.CTkLabel(self, text="PDF 壓縮", font=ctk.CTkFont(size=28, weight="bold")).pack(pady=30)
        ctk.CTkLabel(self, text="降低 PDF 檔案大小", text_color="gray").pack()


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
    ("✂️  PDF 拆分",   "split"),
    ("🗜️  PDF 壓縮",   "compress"),
    ("🖼️  加浮水印",   "watermark"),
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
            "split":     SplitPage,
            "compress":  CompressPage,
            "watermark": WatermarkPage,
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
