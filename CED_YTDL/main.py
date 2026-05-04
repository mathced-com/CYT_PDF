import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import yt_dlp
import threading
import os
import sys
import re
import time
import urllib.request
import io

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 將工作目錄設定為程式所在資料夾
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class YouTubeDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CED_YouTube 下載器")
        self.root.geometry("750x650")
        self.root.resizable(False, False)
        
        self.download_path = tk.StringVar(value=os.path.join(os.getcwd(), "Downloads"))
        self.format_choice = tk.StringVar(value="mp4")
        self.quality_choice = tk.StringVar()
        
        self.video_info = None
        self.is_playlist = False
        
        self.playlist_vars = []
        self.playlist_entries = []
        
        # 暫停與取消狀態標記
        self.is_paused = False
        self.is_cancelled = False
        
        self.create_widgets()
        self.update_quality_options()
        
        if not HAS_PIL:
            messagebox.showwarning("缺少套件", "系統缺少 Pillow 套件，將無法顯示影片封面。")

    def create_widgets(self):
        title_label = tk.Label(self.root, text="CED_YouTube 下載器", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        url_frame = tk.Frame(self.root)
        url_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(url_frame, text="網址：", font=("Arial", 12)).pack(side="left")
        
        self.url_entry = tk.Entry(url_frame, width=35, font=("Arial", 10))
        self.url_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        self.analyze_btn = tk.Button(url_frame, text="解析網址", command=self.start_analyze, bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
        self.analyze_btn.pack(side="left", padx=2)
        
        self.clear_btn = tk.Button(url_frame, text="清除網址", command=self.clear_url, font=("Arial", 10))
        self.clear_btn.pack(side="left", padx=2)
        
        # 步驟提示
        hint_label = tk.Label(url_frame, text="💡 步驟：1.解析網址 ➔ 2.開始下載", fg="#E91E63", font=("Arial", 9, "bold"))
        hint_label.pack(side="left", padx=5)
        
        self.info_frame = tk.LabelFrame(self.root, text="影片預覽 / 播放清單", font=("Arial", 10))
        self.info_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        self.title_label = tk.Label(self.info_frame, text="請輸入網址並點選「解析網址」", fg="gray", wraplength=650, justify="left")
        self.title_label.pack(pady=5, padx=10)
        
        self.list_frame = ScrollableFrame(self.info_frame)
        
        self.select_btn_frame = tk.Frame(self.info_frame)
        tk.Button(self.select_btn_frame, text="全部勾選", command=self.select_all).pack(side="left", padx=5)
        tk.Button(self.select_btn_frame, text="取消全選", command=self.deselect_all).pack(side="left", padx=5)
        
        format_frame = tk.Frame(self.root)
        format_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(format_frame, text="格式：", font=("Arial", 12)).pack(side="left")
        tk.Radiobutton(format_frame, text="MP4", variable=self.format_choice, value="mp4", command=self.update_quality_options).pack(side="left", padx=2)
        tk.Radiobutton(format_frame, text="MP3", variable=self.format_choice, value="mp3", command=self.update_quality_options).pack(side="left", padx=2)
        
        tk.Label(format_frame, text="   品質：", font=("Arial", 12)).pack(side="left")
        self.quality_combo = ttk.Combobox(format_frame, textvariable=self.quality_choice, state="readonly", width=18)
        self.quality_combo.pack(side="left", padx=5)
        
        path_frame = tk.Frame(self.root)
        path_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(path_frame, text="儲存：", font=("Arial", 12)).pack(side="left")
        self.path_entry = tk.Entry(path_frame, textvariable=self.download_path, width=40, state="readonly", font=("Arial", 10))
        self.path_entry.pack(side="left", padx=5, fill="x", expand=True)
        tk.Button(path_frame, text="選擇資料夾", command=self.browse_folder).pack(side="left")
        
        status_frame = tk.Frame(self.root)
        status_frame.pack(fill="x", padx=20, pady=5)
        self.progress_bar = ttk.Progressbar(status_frame, orient="horizontal", length=700, mode="determinate")
        self.progress_bar.pack(pady=2)
        self.status_label = tk.Label(status_frame, text="等待解析...", fg="blue", font=("Arial", 10))
        self.status_label.pack(pady=2)
        
        # 執行與控制按鈕區
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=5)
        self.download_btn = tk.Button(btn_frame, text="開始下載", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", width=12, command=self.start_download, state="disabled")
        self.download_btn.pack(side="left", padx=5)
        
        self.pause_btn = tk.Button(btn_frame, text="暫停", font=("Arial", 10), command=self.toggle_pause, state="disabled", width=8)
        self.pause_btn.pack(side="left", padx=5)
        
        self.cancel_btn = tk.Button(btn_frame, text="取消", font=("Arial", 10), command=self.cancel_download, state="disabled", bg="#f44336", fg="white", width=8)
        self.cancel_btn.pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="檢查更新 yt-dlp", command=self.update_ytdlp).pack(side="left", padx=15)

    def select_all(self):
        for var in self.playlist_vars:
            var.set(True)
            
    def deselect_all(self):
        for var in self.playlist_vars:
            var.set(False)

    def toggle_pause(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_btn.config(text="暫停", bg="SystemButtonFace")
            self.update_progress_ui(self.progress_bar['value'], "繼續下載...", "blue")
        else:
            self.is_paused = True
            self.pause_btn.config(text="繼續", bg="#FFC107")
            self.update_progress_ui(self.progress_bar['value'], "下載已暫停", "orange")

    def cancel_download(self):
        if messagebox.askyesno("確認取消", "確定要取消目前的下載任務嗎？"):
            self.is_cancelled = True
            self.is_paused = False # 釋放可能在暫停狀態的迴圈
            self.update_progress_ui(self.progress_bar['value'], "正在終止下載程序，請稍候...", "red")
            self.cancel_btn.config(state="disabled")
            self.pause_btn.config(state="disabled")

    def update_quality_options(self):
        if self.format_choice.get() == "mp4":
            options = ["最高畫質 (自動)", "1080p", "720p", "480p", "360p"]
        else:
            options = ["最高音質 (320k)", "標準音質 (192k)", "普通音質 (128k)"]
        self.quality_combo['values'] = options
        self.quality_combo.current(0)

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.download_path.get())
        if folder:
            self.download_path.set(folder)

    def update_ytdlp(self):
        self.update_progress_ui(0, "正在更新 yt-dlp... 請稍候", "orange")
        def run_update():
            result = os.system(f"{sys.executable} -m pip install -U yt-dlp")
            if result == 0:
                self.root.after(0, lambda: messagebox.showinfo("更新成功", "yt-dlp 已更新至最新版！"))
                self.root.after(0, lambda: self.update_progress_ui(0, "準備就緒", "blue"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(0, "更新失敗", "red"))
        threading.Thread(target=run_update, daemon=True).start()

    def clear_url(self):
        self.url_entry.delete(0, tk.END)
        self.title_label.config(text="請輸入網址並點選「解析網址」")
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
        self.list_frame.pack_forget()
        self.select_btn_frame.pack_forget()
        self.download_btn.config(state="disabled")
        self.video_info = None
        self.is_playlist = False
        self.update_progress_ui(0, "等待解析...", "blue")

    def start_analyze(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("警告", "請輸入 YouTube 網址！")
            return
            
        self.analyze_btn.config(state="disabled")
        self.download_btn.config(state="disabled")
        self.update_progress_ui(0, "正在解析網址與抓取標題，請稍候...", "blue")
        self.title_label.config(text="解析中...")
        self.list_frame.pack_forget()
        self.select_btn_frame.pack_forget()
        
        threading.Thread(target=self.process_analyze, args=(url,), daemon=True).start()
        
    def process_analyze(self, url):
        import urllib.parse
        parsed_url = urllib.parse.urlparse(url)
        query_params = urllib.parse.parse_qs(parsed_url.query)
        if 'list' in query_params:
            playlist_id = query_params['list'][0]
            url = f"https://www.youtube.com/playlist?list={playlist_id}"

        ydl_opts = {
            'extract_flat': True, 
            'quiet': True
        }
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)
                
            self.video_info = info
            
            if 'entries' in info:
                self.is_playlist = True
                entries = list(info.get('entries') or [])
                self.root.after(0, lambda: self.show_playlist(info.get('title', '播放清單'), entries))
            else:
                self.is_playlist = False
                title = info.get('title', '未知影片標題')
                dur_str = self.format_duration(info.get('duration'))
                thumb_url = info.get('thumbnail')
                self.root.after(0, lambda: self.show_single_video(title, dur_str, thumb_url))
                
        except Exception as e:
            self.root.after(0, lambda: self.title_label.config(text="解析失敗，請確認網址是否正確。"))
            self.root.after(0, lambda: self.update_progress_ui(0, "發生錯誤", "red"))
            self.root.after(0, lambda: messagebox.showerror("錯誤", f"解析失敗：\n{str(e)}"))
        finally:
            self.root.after(0, lambda: self.analyze_btn.config(state="normal"))

    def format_duration(self, seconds):
        if not seconds:
            return ""
        m, s = divmod(int(seconds), 60)
        h, m = divmod(m, 60)
        return f" [{h}:{m:02d}:{s:02d}]" if h else f" [{m:02d}:{s:02d}]"

    def show_single_video(self, title, dur_str, thumb_url):
        self.title_label.config(text="【單一影片解析結果】")
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=5)
        row_frame.pack(fill="x", anchor="w")
        
        thumb_label = tk.Label(row_frame, text="無圖片", bg="#e0e0e0", width=14, height=3)
        thumb_label.pack(side="left", padx=5)
        
        txt_label = tk.Label(row_frame, text=f"{title}\n時間: {dur_str.strip() if dur_str else '未知'}", justify="left", wraplength=500, font=("Arial", 10))
        txt_label.pack(side="left", anchor="w", padx=10)
        
        if HAS_PIL and thumb_url:
            threading.Thread(target=self.load_thumbnail, args=(thumb_url, thumb_label), daemon=True).start()
        
        self.update_progress_ui(0, "解析完成！請確認資訊後點擊「開始下載」", "green")
        self.download_btn.config(state="normal")

    def show_playlist(self, title, entries):
        self.title_label.config(text=f"【播放清單】\n{title} (共 {len(entries)} 部影片)")
        
        for widget in self.list_frame.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.playlist_vars.clear()
        self.playlist_entries = entries
        
        for i, entry in enumerate(entries):
            var = tk.BooleanVar(value=True)
            self.playlist_vars.append(var)
            
            row_frame = tk.Frame(self.list_frame.scrollable_frame, pady=3)
            row_frame.pack(fill="x", anchor="w")
            
            chk = tk.Checkbutton(row_frame, variable=var)
            chk.pack(side="left", padx=5)
            
            thumb_label = tk.Label(row_frame, text="無圖片", bg="#e0e0e0", width=14, height=3)
            thumb_label.pack(side="left", padx=5)
            
            dur_str = self.format_duration(entry.get('duration'))
            title_text = entry.get('title', f'隱藏影片 {i+1}')
            txt_label = tk.Label(row_frame, text=f"{i+1}. {title_text}\n時間: {dur_str.strip() if dur_str else '未知'}", justify="left", wraplength=450, font=("Arial", 10))
            txt_label.pack(side="left", anchor="w")
            
            if HAS_PIL:
                url = entry.get('thumbnail')
                if not url and entry.get('thumbnails'):
                    url = entry['thumbnails'][0].get('url')
                if url:
                    threading.Thread(target=self.load_thumbnail, args=(url, thumb_label), daemon=True).start()

        self.list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.select_btn_frame.pack(pady=5)
        
        self.update_progress_ui(0, "解析完成！請勾選想下載的集數，點擊「開始下載」", "green")
        self.download_btn.config(state="normal")

    def load_thumbnail(self, url, label):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            raw_data = urllib.request.urlopen(req, timeout=5).read()
            im = Image.open(io.BytesIO(raw_data))
            im.thumbnail((100, 56))
            photo = ImageTk.PhotoImage(im)
            self.root.after(0, lambda: self._set_image(label, photo))
        except Exception:
            pass
            
    def _set_image(self, label, photo):
        label.config(image=photo, text="", width=100, height=56)
        label.image = photo

    def update_progress_ui(self, value, text, color="blue"):
        self.progress_bar['value'] = value
        self.status_label.config(text=text, fg=color)

    def progress_hook(self, d):
        # 攔截下載封包，實現暫停與取消
        while self.is_paused:
            if self.is_cancelled:
                raise ValueError("USER_CANCELLED")
            time.sleep(0.5)
            
        if self.is_cancelled:
            raise ValueError("USER_CANCELLED")

        if d['status'] == 'downloading':
            downloaded = d.get('downloaded_bytes', 0)
            total = d.get('total_bytes') or d.get('total_bytes_estimate', 0)
            percent_val = (downloaded / total * 100) if total > 0 else 0.0
            
            ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
            percent_str = ansi_escape.sub('', d.get('_percent_str', f'{percent_val:.1f}%')).strip()
            speed = ansi_escape.sub('', d.get('_speed_str', 'N/A')).strip()
            eta = ansi_escape.sub('', d.get('_eta_str', 'N/A')).strip()
            
            self.root.after(0, lambda: self.update_progress_ui(percent_val, f"下載進度: {percent_str} (速度: {speed}, 剩餘: {eta})", "blue"))
            
        elif d['status'] == 'finished':
            self.root.after(0, lambda: self.update_progress_ui(100.0, "單檔下載完成！正在合併影像或轉檔... (此階段無法暫停)", "orange"))

    def start_download(self):
        save_dir = self.download_path.get()
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
            
        fmt = self.format_choice.get()
        quality = self.quality_combo.get()
        
        urls_to_download = []
        if self.is_playlist:
            selected_indices = [i for i, var in enumerate(self.playlist_vars) if var.get()]
            if not selected_indices:
                messagebox.showwarning("提示", "請至少在清單中勾選一部影片！")
                return
            entries = self.playlist_entries
            for i in selected_indices:
                vid_url = entries[i].get('url') or entries[i].get('webpage_url')
                if vid_url:
                    urls_to_download.append(vid_url)
                else:
                    vid_id = entries[i].get('id')
                    if vid_id:
                        urls_to_download.append(f"https://www.youtube.com/watch?v={vid_id}")
        else:
            urls_to_download.append(self.url_entry.get().strip())

        self.download_btn.config(state="disabled")
        self.analyze_btn.config(state="disabled")
        
        # 啟用控制按鈕並重置狀態
        self.is_cancelled = False
        self.is_paused = False
        self.pause_btn.config(state="normal", text="暫停", bg="SystemButtonFace")
        self.cancel_btn.config(state="normal")
        self.update_progress_ui(0, "準備開始下載...", "blue")
        
        threading.Thread(target=self.process_download, args=(urls_to_download, save_dir, fmt, quality), daemon=True).start()

    def process_download(self, urls, save_dir, fmt, quality):
        if fmt == "mp4":
            if "最高畫質" in quality:
                format_str = 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best'
            elif "1080" in quality:
                format_str = 'bestvideo[height<=1080][ext=mp4]+bestaudio[ext=m4a]/best[height<=1080][ext=mp4]/best'
            elif "720" in quality:
                format_str = 'bestvideo[height<=720][ext=mp4]+bestaudio[ext=m4a]/best[height<=720][ext=mp4]/best'
            elif "480" in quality:
                format_str = 'bestvideo[height<=480][ext=mp4]+bestaudio[ext=m4a]/best[height<=480][ext=mp4]/best'
            else:
                format_str = 'bestvideo[height<=360][ext=mp4]+bestaudio[ext=m4a]/best[height<=360][ext=mp4]/best'
                
            ydl_opts = {
                'outtmpl': os.path.join(save_dir, '%(title)s.%(ext)s'),
                'format': format_str,
                'merge_output_format': 'mp4',
                'progress_hooks': [self.progress_hook],
                'ffmpeg_location': os.getcwd(),
                'color': 'no_color'
            }
        else:
            if "320" in quality:
                kbps = '320'
            elif "192" in quality:
                kbps = '192'
            else:
                kbps = '128'
                
            ydl_opts = {
                'outtmpl': os.path.join(save_dir, '%(title)s.%(ext)s'),
                'format': 'bestaudio/best',
                'postprocessors': [{
                    'key': 'FFmpegExtractAudio',
                    'preferredcodec': 'mp3',
                    'preferredquality': kbps,
                }],
                'progress_hooks': [self.progress_hook],
                'ffmpeg_location': os.getcwd(),
                'color': 'no_color'
            }
            
        ydl_opts['noplaylist'] = True

        try:
            total = len(urls)
            for i, url in enumerate(urls):
                if self.is_cancelled:
                    break
                    
                if total > 1:
                    self.root.after(0, lambda idx=i: self.update_progress_ui(0, f"即將下載清單第 {idx+1}/{total} 部，請稍候...", "blue"))
                else:
                    self.root.after(0, lambda: self.update_progress_ui(0, "連線中，準備開始下載...", "blue"))
                    
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    ydl.download([url])
                    
            if self.is_cancelled:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載任務已取消", "red"))
                self.root.after(0, lambda: messagebox.showinfo("取消", "已成功取消下載任務。\n(未完成的暫存檔已保留，未來重新下載可自動接續進度)"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(100.0, "所有任務皆已處理完成！", "green"))
                self.root.after(0, lambda: messagebox.showinfo("成功", f"全部下載完畢！\n檔案已成功儲存至：\n{save_dir}"))
                
        except Exception as e:
            # 判斷是否為我們主動拋出的取消例外
            if "USER_CANCELLED" in str(e):
                self.root.after(0, lambda: self.update_progress_ui(0, "下載任務已取消", "red"))
                self.root.after(0, lambda: messagebox.showinfo("取消", "已成功取消下載任務。\n(未完成的暫存檔已保留，未來重新下載可自動接續進度)"))
            else:
                self.root.after(0, lambda: self.update_progress_ui(0, "下載過程發生錯誤", "red"))
                self.root.after(0, lambda: messagebox.showerror("錯誤", f"下載失敗，可能是網路問題或影片遭版權封鎖：\n{str(e)}"))
        finally:
            self.root.after(0, lambda: self.download_btn.config(state="normal"))
            self.root.after(0, lambda: self.analyze_btn.config(state="normal"))
            self.root.after(0, lambda: self.pause_btn.config(state="disabled", text="暫停", bg="SystemButtonFace"))
            self.root.after(0, lambda: self.cancel_btn.config(state="disabled"))

if __name__ == "__main__":
    root = tk.Tk()
    app = YouTubeDownloaderGUI(root)
    root.mainloop()
