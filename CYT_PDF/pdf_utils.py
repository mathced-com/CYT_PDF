import os
import pypdf
from typing import Callable, Dict
from pdf2image import convert_from_path, pdfinfo_from_path
from pdf2image.exceptions import (
    PDFInfoNotInstalledError, 
    PDFPageCountError, 
    PDFSyntaxError
)

)

class EncryptedPDFError(Exception):
    """自訂例外：當 PDF 受密碼保護且未能解密時拋出"""
    def __init__(self, filepath: str):
        self.filepath = filepath
        super().__init__(f"檔案受密碼保護: {filepath}")

def merge_pdfs(
    input_paths: list[str], 
    output_path: str, 
    passwords: Dict[str, str] | None = None
) -> tuple[bool, str]:
    """
    合併多個 PDF 檔案。
    - 處理了資源釋放 (防止檔案佔用)
    - 處理密碼加密問題
    """
    if passwords is None:
        passwords = {}
        
    merger = pypdf.PdfWriter()
    file_handles = []
    
    try:
        for path in input_paths:
            f = open(path, "rb")
            file_handles.append(f)
            reader = pypdf.PdfReader(f)
            
            # 檢查是否加密
            if reader.is_encrypted:
                pwd = passwords.get(path, "")
                if not reader.decrypt(pwd):
                    # 解密失敗拋出例外，供 GUI 捕捉並要求輸入密碼
                    raise EncryptedPDFError(path)
                    
            merger.append(reader)
            
        with open(output_path, "wb") as out_f:
            merger.write(out_f)
            
        return True, f"成功合併至:\n{output_path}"
    except EncryptedPDFError as e:
        raise e
    except Exception as e:
        return False, f"合併失敗: {str(e)}"
    finally:
        # 確保所有檔案被關閉，避免資源洩漏
        for f in file_handles:
            try:
                f.close()
            except:
                pass
        try:
            merger.close()
        except:
            pass


def pdf_to_jpg(
    input_path: str, 
    output_folder: str, 
    dpi: int = 200, 
    quality: int = 80, 
    progress_callback: Callable[[float], None] | None = None
) -> tuple[bool, str]:
    """
    將 PDF 轉換為一系列 JPG 圖片。
    
    Args:
        input_path: PDF 檔案路徑
        output_folder: 輸出的資料夾路徑
        dpi: 解析度 (Dots Per Inch)
        quality: JPG 品質 (1-100)
        progress_callback: 進度回呼函式，接收一個 0.0 到 1.0 的浮點數
        
    Returns:
        (是否成功, 訊息)
    """
    try:
        # 1. 取得 PDF 資訊（主要為了取得總頁數）
        # 如果 poppler 未安裝，這裡通常會噴出 PDFInfoNotInstalledError
        info = pdfinfo_from_path(input_path)
        total_pages = info["Pages"]
        
        if total_pages == 0:
            return False, "該 PDF 檔案沒有任何頁面。"

        # 2. 建立輸出目錄
        if not os.path.exists(output_folder):
            os.makedirs(output_folder, exist_ok=True)
            
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        
        # 3. 逐頁轉換以回傳進度
        for i in range(1, total_pages + 1):
            # 轉換單一頁面
            pages = convert_from_path(
                input_path, 
                first_page=i, 
                last_page=i, 
                dpi=dpi,
                fmt="jpeg"
            )
            
            if pages:
                # 存檔
                output_filename = f"{base_name}_page_{i:03d}.jpg"
                save_path = os.path.join(output_folder, output_filename)
                pages[0].save(save_path, "JPEG", quality=quality)
            
            # 4. 回傳進度 (0.0 ~ 1.0)
            if progress_callback:
                progress_callback(i / total_pages)
                
        return True, f"成功！已將 {total_pages} 頁轉換並儲存至 {output_folder}"

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

if __name__ == "__main__":
    # 測試程式碼（僅供開發參考）
    def my_progress(p):
        print(f"目前進度: {p*100:.1f}%")
        
    # success, msg = pdf_to_jpg("test.pdf", "output", progress_callback=my_progress)
    # print(msg)
    pass
