from PIL import Image
import os

# 圖片來源 (剛生成的 PNG)
png_path = r'C:\Users\ced\.gemini\antigravity\brain\6cb231d3-f90a-4773-a42c-269da5babd14\cyt_pdf_icon_1777904617902.png'
# 輸出目標
ico_path = r'd:\Antigravity\CYT_PDF\icon.ico'

try:
    if not os.path.exists(png_path):
        print(f"錯誤：找不到來源圖片 {png_path}")
    else:
        img = Image.open(png_path)
        # 設定標準 ICO 的多種尺寸
        icon_sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (255, 255)]
        img.save(ico_path, format='ICO', sizes=icon_sizes)
        print(f"成功將圖示轉換為：{ico_path}")
except ImportError:
    print("錯誤：系統中未安裝 Pillow 套件。請執行 pip install Pillow")
except Exception as e:
    print(f"發生非預期錯誤：{e}")
