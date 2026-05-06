import os
import re
import sys
import subprocess
import webbrowser

def get_next_version(current):
    parts = current.split('.')
    if len(parts) == 3 and parts[-1].isdigit():
        parts[-1] = str(int(parts[-1]) + 1)
        return '.'.join(parts)
    return current + "_new"

def check_gh_login():
    try:
        result = subprocess.run(["gh", "auth", "status"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
        if result.returncode != 0:
            print("\n[!] 偵測到您尚未登入 GitHub 命令列工具。")
            print("為了能自動上傳檔案，現在將啟動一次性登入流程：")
            print("請在接下來的提示中選擇：")
            print("1. 選擇 GitHub.com")
            print("2. 選擇 HTTPS")
            print("3. 選擇 Y (Authenticate Git with your GitHub credentials)")
            print("4. 選擇 Login with a web browser")
            print("5. 複製一次性驗證碼並在瀏覽器貼上\n")
            subprocess.run(["gh", "auth", "login"])
            
            result2 = subprocess.run(["gh", "auth", "status"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
            if result2.returncode != 0:
                print("\n登入失敗或取消，無法自動上傳 Release。")
                return False
        return True
    except FileNotFoundError:
        print("\n[!] 系統中找不到 gh 指令 (GitHub CLI)。請先安裝 GitHub CLI 才能完全自動化。")
        return False

def convert_md_to_txt(md_path, txt_path):
    try:
        with open(md_path, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 簡單的去標記化：去除 GitHub 警示框語法
        content = re.sub(r'> \[!IMPORTANT\]', '【重要聲明】', content)
        content = re.sub(r'> \[!NOTE\]', '【備註】', content)
        content = re.sub(r'> \[!TIP\]', '【提示】', content)
        content = re.sub(r'# ', '', content)
        content = re.sub(r'## ', '■ ', content)
        content = re.sub(r'### ', '  - ', content)
        content = re.sub(r'\*\*', '', content)
        
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"轉換說明檔失敗: {e}")
        return False

def main():
    print("==============================================")
    print("      專案一鍵發布新版本助手 (全自動版)")
    print("==============================================\n")
    
    # 請修改此處以符合您的新專案主程式名稱
    main_script = "app.py" 
    github_repo = "未知"

    try:
        with open(main_script, "r", encoding="utf-8") as f:
            content = f.read()
            # 讀取版本號
            v_match = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', content)
            if v_match:
                current_version = v_match.group(1)
            # 讀取 GitHub 倉庫路徑
            r_match = re.search(r'GITHUB_REPO\s*=\s*"([^"]+)"', content)
            if r_match:
                github_repo = r_match.group(1)
    except Exception as e:
        print(f"讀取 {main_script} 失敗 (請確認檔案存在且包含 APP_VERSION 變數): {e}")
        input("請按 Enter 鍵結束...")
        return

    suggested_version = get_next_version(current_version)
    print(f"目前版本為: {current_version}")
    
    new_version = input(f"請輸入新的版本號 [直接按 Enter 預設為 {suggested_version}]: ").strip()
    if not new_version:
        new_version = suggested_version
        
    print(f"\n[OK] 將發布新版本: {new_version}")
    
    update_notes = input("\n請簡單輸入這次更新的內容: ").strip()
    if not update_notes:
        update_notes = "一般更新與修復"
        
    print("\n[1/6] 正在檢查 GitHub 授權狀態...")
    has_gh = check_gh_login()
    if not has_gh:
        print("\n無法使用自動上傳，請取消這次發布，或改用手動發布。")
        input("請按 Enter 鍵結束...")
        return
        
    print(f"\n[2/6] 正在更新版本號...")
    try:
        # 更新主程式
        new_content = re.sub(r'APP_VERSION\s*=\s*"[^"]+"', f'APP_VERSION = "{new_version}"', content)
        with open(main_script, "w", encoding="utf-8") as f:
            f.write(new_content)
            
        # 同步更新 使用說明.md
        if os.path.exists("使用說明.md"):
            with open("使用說明.md", "r", encoding="utf-8") as f:
                md_content = f.read()
            md_content = re.sub(r'使用說明 \(v[^)]+\)', f'使用說明 (v{new_version})', md_content)
            md_content = re.sub(r'Version [0-9.]+', f'Version {new_version}', md_content)
            with open("使用說明.md", "w", encoding="utf-8") as f:
                f.write(md_content)
            print("      [OK] 使用說明.md 版本號已同步。")
            
    except Exception as e:
        print(f"更新版本號失敗: {e}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[3/6] 正在打包成執行檔...")
    # 注意：這裡的 --name 會影響 exe 檔名
    subprocess.run(["py", "-3", "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed", "--icon=icon.ico", "--add-data", "icon.ico;.", "--name=CYT_PDF_Tool", "app.py"])



    
    # 這裡請手動修改為您在 PyInstaller 中設定的名稱
    exe_name = "CYT_PDF_Tool" 
    exe_path = os.path.join("dist", f"{exe_name}.exe")
    zip_path = os.path.join("dist", f"{exe_name}.zip")
    txt_path = os.path.join("dist", "程式說明.txt")
    
    if not os.path.exists(exe_path):
        print(f"\n[Error] 打包失敗，找不到 {exe_path}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[3.3/6] 正在產生文字版說明...")
    convert_md_to_txt("使用說明.md", txt_path)

    print("\n[3.5/6] 正在建立 ZIP 壓縮包...")
    try:
        import zipfile
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 寫入執行檔與說明檔
            zipf.write(exe_path, os.path.basename(exe_path))
            if os.path.exists(txt_path):
                zipf.write(txt_path, os.path.basename(txt_path))
            
            # 新增：打包 poppler 資料夾
            poppler_dir = "poppler-26.02.0"
            if os.path.exists(poppler_dir):
                print(f"      [OK] 偵測到 Poppler，正在打包...")
                for root, dirs, files in os.walk(poppler_dir):
                    for file in files:
                        file_full_path = os.path.join(root, file)
                        # 保持相對目錄結構
                        zipf.write(file_full_path, file_full_path)
            else:
                print(f"      [Warning] 找不到 {poppler_dir}，壓縮包可能不完整！")
                
        print(f"      [OK] 壓縮完成: {zip_path}")
    except Exception as e:
        print(f"      [Error] 壓縮失敗: {e}")
        input("請按 Enter 鍵結束...")
        return

    print("\n[4/6] 正在推送代碼至 GitHub...")
    subprocess.run(["git", "add", "."])
    subprocess.run(["git", "commit", "-m", f"Release v{new_version}: {update_notes}"])
    subprocess.run(["git", "push"])
    
    print(f"\n[5/6] 正在建立 GitHub Release 並上傳至 {github_repo}...")
    release_cmd = [
        "gh", "release", "create", f"v{new_version}", 
        exe_path, 
        zip_path,
        txt_path,
        "--title", f"v{new_version}", 
        "--notes", update_notes
    ]
    
    # 如果有抓到 GITHUB_REPO，就明確指定
    if github_repo != "未知":
        release_cmd.extend(["--repo", github_repo])
        
    result = subprocess.run(release_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
    
    if result.returncode == 0:
        print("\n[Success] 發布成功！")
    else:
        print(f"\n[Error] 自動發布失敗: {result.stderr}")
    
    input("\n請按 Enter 鍵關閉視窗...")

if __name__ == "__main__":
    main()
