import os
import pypdf
import io
from PIL import Image
from typing import Callable, Dict
from pdf2image import convert_from_path, pdfinfo_from_path
from pdf2image.exceptions import (
    PDFInfoNotInstalledError, 
    PDFPageCountError, 
    PDFSyntaxError
)
# Poppler 執行檔路徑 (用於 PDF 轉圖)
# 建議放在專案目錄下，方便打包
POPPLER_PATH = os.path.join(os.path.dirname(__file__), "poppler-26.02.0", "Library", "bin")

class EncryptedPDFError(Exception):
    """自訂例外：當 PDF 受密碼保護且未能解密時拋出"""
    def __init__(self, filepath: str):
        self.filepath = filepath
        super().__init__(f"檔案受密碼保護: {filepath}")

def merge_pdfs(pdf_list: list[str], output_folder: str, output_filename: str, callback: Callable[[float], None] | None = None) -> tuple[bool, str]:
    """
    合併多個 PDF 檔案。
    """
    try:
        writer = pypdf.PdfWriter()
        total = len(pdf_list)
        
        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"
            
        output_path = os.path.join(output_folder, output_filename)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)

        for i, pdf in enumerate(pdf_list):
            writer.append(pdf)
            if callback:
                callback((i + 1) / total)
        
        with open(output_path, "wb") as f:
            writer.write(f)
            
        return True, f"合併成功！檔案已儲存至：\n{output_path}"
    except Exception as e:
        return False, f"合併失敗: {str(e)}"
    finally:
        try:
            writer.close()
        except:
            pass

def parse_range_string(ranges: str, total_pages: int) -> list[int]:
    """
    解析範圍字串（如 "1-5, 8, 10-12"）並回傳 0-indexed 的頁碼列表。
    """
    target_pages = set()
    if not ranges.strip():
        return list(range(total_pages))
        
    parts = ranges.replace(" ", "").split(",")
    for part in parts:
        try:
            if "-" in part:
                start_str, end_str = part.split("-")
                s = int(start_str)
                e = int(end_str)
                for p in range(s, e + 1):
                    if 1 <= p <= total_pages:
                        target_pages.add(p - 1)
            else:
                p = int(part)
                if 1 <= p <= total_pages:
                    target_pages.add(p - 1)
        except ValueError:
            continue # 忽略格式錯誤的部分
            
    return sorted(list(target_pages))

def split_pdf(
    input_path: str, 
    output_folder: str, 
    mode: str = "single",  # "single" (一頁一檔案), "range" (多頁合成一檔)
    ranges: str = "",      
    custom_name: str = "",
    callback: Callable[[float], None] | None = None
) -> tuple[bool, str]:
    """
    拆分 PDF 檔案。
    """
    try:
        reader = pypdf.PdfReader(input_path)
        total_pages = len(reader.pages)
        base_name = custom_name if custom_name else os.path.splitext(os.path.basename(input_path))[0]

        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)

        target_indices = parse_range_string(ranges, total_pages)
        if not target_indices:
            return False, "未偵測到有效的頁碼範圍。"

        if mode == "single":
            # 一頁一檔案
            for i, p_idx in enumerate(target_indices):
                writer = pypdf.PdfWriter()
                writer.add_page(reader.pages[p_idx])
                output_filename = f"{base_name}_page_{p_idx+1:03d}.pdf"
                with open(os.path.join(output_folder, output_filename), "wb") as f:
                    writer.write(f)
                
                if callback:
                    callback((i + 1) / len(target_indices))
            return True, f"成功！已將選取的 {len(target_indices)} 頁分別儲存為單一檔案。"

        elif mode == "range":
            # 多頁合成一檔
            writer = pypdf.PdfWriter()
            for i, p_idx in enumerate(target_indices):
                writer.add_page(reader.pages[p_idx])
                if callback:
                    callback((i + 1) / len(target_indices))
            
            output_filename = f"{base_name}_extracted.pdf"
            with open(os.path.join(output_folder, output_filename), "wb") as f:
                writer.write(f)
            
            return True, f"成功！已將選取的 {len(target_indices)} 頁合併提取至 {output_filename}"

    except Exception as e:
        return False, f"拆分失敗：\n{str(e)}"

def pdf_to_jpg(
    input_path: str, 
    output_folder: str, 
    dpi: int = 200, 
    quality: int = 85, 
    ranges: str = "", 
    custom_name: str = "",
    callback: Callable[[float], None] | None = None
) -> tuple[bool, str]:
    """
    將 PDF 轉換為一系列 JPG 圖片。
    """
    try:
        # 1. 取得 PDF 資訊
        info = pdfinfo_from_path(input_path, poppler_path=POPPLER_PATH)
        total_pages = info["Pages"]
        
        if total_pages == 0:
            return False, "該 PDF 檔案沒有任何頁面。"

        # 2. 建立輸出目錄
        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)
            
        base_name = custom_name if custom_name else os.path.splitext(os.path.basename(input_path))[0]
        
        target_indices = parse_range_string(ranges, total_pages)
        if not target_indices:
            return False, "未偵測到有效的頁碼範圍。"

        # 3. 逐頁轉換以回傳進度
        for i, p_idx in enumerate(target_indices):
            page_num = p_idx + 1
            # 轉換單一頁面
            pages = convert_from_path(
                input_path, 
                first_page=page_num, 
                last_page=page_num, 
                dpi=dpi,
                fmt="jpeg",
                poppler_path=POPPLER_PATH
            )
            
            if pages:
                # 存檔
                output_filename = f"{base_name}_page_{page_num:03d}.jpg"
                save_path = os.path.join(output_folder, output_filename)
                pages[0].save(save_path, "JPEG", quality=quality)
            
            # 4. 回傳進度 (0.0 ~ 1.0)
            if callback:
                callback((i + 1) / len(target_indices))
                
        return True, f"成功！已將選取的 {len(target_indices)} 頁轉換為圖片並儲存至 {output_folder}"

    except PDFInfoNotInstalledError:
        return False, (
            "【錯誤】找不到 Poppler 執行檔。\n\n"
            "原因：pdf2image 依賴於 Poppler 軟體進行轉換。\n"
            "解決方案：\n"
            "1. 下載 Poppler for Windows (例如從 Release 網站)。\n"
            "2. 解壓縮後將 bin 資料夾的路徑加入系統的 PATH 環境變數中。\n"
            "3. 重新啟動 IDE 或終端機。"
        )
    except PDFPageCountError:
        return False, "無法讀取 PDF 頁數，檔案可能毀損或受密碼保護。"
    except PDFSyntaxError:
        return False, "PDF 語法錯誤，無法處理該檔案。"
    except Exception as e:
        return False, f"轉換過程中發生非預期錯誤：\n{str(e)}"

def compress_pdf(
    input_path: str,
    output_folder: str,
    quality: str = "medium", # "low", "medium", "high"
    custom_name: str = "",
    callback: Callable[[float], None] | None = None
) -> tuple[bool, str]:
    """
    壓縮 PDF 檔案大小。
    """
    try:
        reader = pypdf.PdfReader(input_path)
        writer = pypdf.PdfWriter()
        
        base_name = custom_name if custom_name else os.path.splitext(os.path.basename(input_path))[0]
        output_filename = f"{base_name}_compressed.pdf"
        output_path = os.path.join(output_folder, output_filename)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)

        # 根據等級決定圖片品質與縮放限制
        img_quality = 65 if quality == "low" else (40 if quality == "medium" else 20)
        max_dim = 1800 if quality == "low" else (1200 if quality == "medium" else 800)

        processed_ids = set()

        total_pages = len(reader.pages)
        for i, page in enumerate(reader.pages):
            new_page = writer.add_page(page)
            new_page.compress_content_streams() # 壓縮文字與路徑
            
            try:
                for img_proxy in new_page.images:
                    try:
                        img_id = img_proxy.indirect_reference.idnum
                    except:
                        continue 

                    if img_id not in processed_ids:
                        # 取得原始資料大小
                        original_data_size = len(img_proxy.data)
                        pil_img = img_proxy.image
                        
                        # 轉為 RGB 以確保 JPEG 壓縮效率 (移除透明通道)
                        if pil_img.mode in ("RGBA", "P"):
                            pil_img = pil_img.convert("RGB")
                        
                        w, h = pil_img.size
                        if max(w, h) > max_dim:
                            ratio = max_dim / max(w, h)
                            pil_img = pil_img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
                        
                        # 在記憶體中模擬壓縮後的結果
                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format="JPEG", quality=img_quality)
                        compressed_data_size = img_byte_arr.tell()

                        # 終極保換：只有當壓縮後真的變小，才進行替換
                        if compressed_data_size < original_data_size:
                            img_proxy.replace(pil_img, quality=img_quality)
                        
                        processed_ids.add(img_id)
            except Exception as e:
                pass
                
            if callback:
                callback((i + 1) / total_pages)

        # 寫入檔案時，嘗試啟動內部壓縮
        # pypdf 會在 write 時自動對未壓縮的物件進行 FlateDecode
        with open(output_path, "wb") as f:
            writer.write(f)

        return True, output_path
    except Exception as e:
        return False, f"壓縮失敗: {str(e)}"


if __name__ == "__main__":
    # 測試程式碼（僅供開發參考）
    def my_progress(p):
        print(f"目前進度: {p*100:.1f}%")
        
    # success, msg = pdf_to_jpg("test.pdf", "output", progress_callback=my_progress)
    # print(msg)
    pass
