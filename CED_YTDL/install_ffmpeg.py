import os
import urllib.request
import zipfile
import shutil
import ssl

# 解決部分 Windows Python 環境的 SSL 憑證驗證錯誤
ssl._create_default_https_context = ssl._create_unverified_context

# 採用 gyan.dev 提供的 Windows 穩定精簡版 FFmpeg (足夠 yt-dlp 使用)
FFMPEG_URL = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
ZIP_FILE = "ffmpeg.zip"

def download_ffmpeg():
    if os.path.exists("ffmpeg.exe") and os.path.exists("ffprobe.exe"):
        print("FFmpeg 相關檔案已存在，不需重新下載。")
        return

    print("開始下載 FFmpeg (檔案約 40MB，可能需要幾分鐘，請稍候)...")
    try:
        # 下載檔案
        urllib.request.urlretrieve(FFMPEG_URL, ZIP_FILE)
        print("下載完成，開始解壓縮與提取...")
        
        with zipfile.ZipFile(ZIP_FILE, 'r') as zip_ref:
            # 在壓縮檔內尋找 bin 資料夾的路徑
            bin_path = None
            for name in zip_ref.namelist():
                if name.endswith('bin/ffmpeg.exe'):
                    bin_path = os.path.dirname(name)
                    break
            
            if bin_path:
                # 僅提取我們需要的兩個核心執行檔，放到當前目錄
                for exe in ['ffmpeg.exe', 'ffprobe.exe']:
                    source = f"{bin_path}/{exe}"
                    target = os.path.join(os.getcwd(), exe)
                    with zip_ref.open(source) as zf, open(target, 'wb') as f:
                        shutil.copyfileobj(zf, f)
                print("FFmpeg 提取成功！環境配置完成。")
            else:
                print("錯誤：在壓縮檔內找不到 FFmpeg 執行檔。")

    except Exception as e:
        print(f"下載或解壓縮時發生錯誤: {e}")
    finally:
        # 清理暫存的壓縮檔
        if os.path.exists(ZIP_FILE):
            os.remove(ZIP_FILE)
            print("已清理下載暫存檔。")

if __name__ == "__main__":
    download_ffmpeg()
