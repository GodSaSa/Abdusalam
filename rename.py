import os
import win32com.client # 使用 Windows 原生組件，打包更輕量

def rename_excel_files():
    print("--- ⚡ 寶貝專屬：自動重命名工具 ⚡ ---")
    current_dir = os.getcwd()
    
    try:
        # 調用 Windows 原生 Excel 接口
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
    except:
        print("❌ 錯誤：這台電腦似乎沒有安裝 Excel")
        return

    files = [f for f in os.listdir(current_dir) if f.endswith(('.xlsx', '.xls'))]
    
    for filename in files:
        if filename.startswith('~$'): continue # 跳過臨時文件
        
        file_path = os.path.join(current_dir, filename)
        try:
            wb = excel.Workbooks.Open(file_path)
            # 1. 嘗試讀取文檔屬性中的標題
            new_title = wb.BuiltinDocumentProperties("Title").Value
            
            # 2. 如果屬性為空，讀取第一張表的第一個單元格 (A1)
            if not new_title or str(new_title).strip() == "":
                new_title = wb.Sheets(1).Cells(1, 1).Value
            
            wb.Close(False)

            if new_title:
                # 過濾 Windows 不允許的特殊字符
                safe_title = "".join([c for c in str(new_title) if c not in r'\/:*?"<>|']).strip()
                ext = os.path.splitext(filename)[1]
                new_filename = f"{safe_title}{ext}"
                
                new_path = os.path.join(current_dir, new_filename)
                
                if not os.path.exists(new_path):
                    os.rename(file_path, new_path)
                    print(f"✅ 成功: {filename} -> {new_filename}")
                else:
                    print(f"⚠️ 跳過: {new_filename} 已存在")
        except Exception as e:
            print(f"❌ 處理 {filename} 時出錯")

    excel.Quit()
    print("\n✨ 全部處理完成！按任意鍵關閉 ✨")
    input()

if __name__ == "__main__":
    rename_excel_files()