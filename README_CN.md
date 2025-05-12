# Record_calender

開發日誌：Python 待辦事項與行事曆工具

專案目標： 開發一個具備待辦事項管理、日期關聯及基本行事曆查看功能的 Python 應用程式。

開發階段：

階段 1：文字介面原型開發

日期： (假設起始日期)
目標： 建立核心功能，包括新增、列出、標記完成待辦事項，並能關聯到期日，以文字介面呈現。
實現：
使用 Python 內建資料結構（列表中的字典）儲存待辦事項。
使用 datetime 模組處理日期。
使用 json 模組實現資料的儲存和載入，保持程式持久性。
實作文字命令選單。
基本的「行事曆」功能以按日期分組列出待辦事項的形式呈現。
結果： 完成一個基礎的文字介面待辦事項管理工具。
階段 2：轉換為 GUI 介面 (使用 Tkinter)

日期： (階段 1 後)
目標： 將文字介面改為圖形使用者介面 (GUI)，支援滑鼠點擊操作。
實現：
選擇標準的 tkinter 和 tkinter.ttk 作為 GUI 庫。
設計 GUI 佈局，包括輸入區、按鈕、列表顯示區。
使用 tkinter.ttk.Treeview 來顯示結構化的待辦事項列表（內容、到期日、狀態）。
將後端邏輯（新增、完成、儲存、載入）綁定到 GUI 按鈕和事件上。
日期輸入暫時使用文字輸入框。
結果： 應用程式具備基本的 GUI 操作能力，列表顯示更直觀。
階段 3：引入現代 GUI 風格 (CustomTkinter Hybrid)

日期： (階段 2 後)
目標： 提升應用程式的視覺外觀，採用更現代的風格（推測使用者意指 customtkinter）。
實現：
引入 customtkinter 庫，用於視窗、框架、按鈕、輸入框等元件。
挑戰： 發現 customtkinter 沒有內建 Treeview 元件。
解決方案： 採用混合模式，繼續使用 tkinter.ttk.Treeview 來顯示列表，將其嵌入到 customtkinter.CTkFrame 中。
引入唯一的任務 ID，以便在 Treeview 中選中項目時能精確對應到後端資料列表中的任務。
更新儲存/載入邏輯以處理新的任務 ID。
結果： 應用程式界面風格更現代化，列表顯示功能保持。
階段 4：功能擴展 (日曆選擇、多狀態、備註/超連結)

日期： (階段 3 後)
目標： 增加更完善的日曆互動、更多任務狀態選項以及帶有超連結的備註欄位。
實現：
引入 tkcalendar 庫實現圖形化日期選擇器（作為新的依賴）。
增加任務狀態下拉選單 (ttk.Combobox) 和「變更狀態」按鈕。
新增備註輸入框 (customtkinter.CTkTextbox)。
挑戰： CTkTextbox 不支援 tkinter.Text 的標籤和綁定功能，無法直接實現可點擊超連結。
解決方案： 在列表下方增加一個「詳細資訊」顯示區域，該區域的備註欄位使用標準 tkinter.Text 元件。
實現 Treeview 選中事件綁定，在選中任務時將詳細資訊載入到下方區域。
實現 find_and_tag_urls 方法，在 tkinter.Text 備註框中查找網址並應用標籤，綁定點擊事件打開網址 (webbrowser 模組）。
初步調整 Enter 鍵綁定，嘗試區分換行和儲存（後續繼續優化）。
結果： 應用程式具備日曆彈窗選日期、多狀態管理、帶超連結的備註詳細顯示功能。
階段 5：除錯與優化 (Hyperlink, Shortcut, Editing UI)

日期： (階段 4 後)
目標： 解決前期遇到的問題並完善使用者互動體驗。
問題 1： CTkTextbox 無法使用 tag_configure 報錯。
除錯： 確認 CTkTextbox 不支援 tkinter.Text 標籤功能。
解決： 將詳細資訊區域的備註元件替換回標準 tkinter.Text，並確保在狀態切換時正確處理其狀態。
問題 2： Ctrl+S 儲存快捷鍵無效。
除錯： 檢查事件綁定方式。
解決： 更改綁定事件字串為 <Control-KeyPress-s>，並在處理函數中返回 "break"。
問題 3： 沒有看到修改任務的按鈕，備註輸入框單獨按 Enter 未換行。
除錯： 確認編輯模式的觸發邏輯和按鈕的顯示/隱藏狀態。調整 Enter 鍵綁定處理。
解決：
增加 Treeview 的雙擊綁定 (<Double-1>) 到 load_task_for_editing 方法，實現載入任務到輸入區進行編輯。
修改「新增」按鈕 (self.save_button) 的文字，使其在編輯模式下顯示為「儲存修改」。
創建「取消編輯」按鈕 (self.cancel_edit_button)，並在進入編輯模式時顯示，退出時隱藏。
修改 handle_return_key 方法，精確判斷焦點所在的元件和是否按下 Ctrl 鍵，實現：描述欄位/日期相關按 Enter 觸發儲存；備註欄位按 Enter 實現換行；備註欄位按 Ctrl+Enter 觸發儲存。
階段 6：新增匯出功能

日期： (階段 5 後)
目標： 增加將待辦事項匯出到 Excel 檔案 (.xlsx) 的功能。
實現：
引入 openpyxl 庫（新的依賴）。
使用 tkinter.filedialog.asksaveasfilename 打開檔案儲存對話框。
創建 export_to_excel 方法：獲取檔案路徑，創建 openpyxl.Workbook，寫入表頭，遍歷任務資料寫入行，儲存檔案，並提供成功/失敗提示。
在 GUI 中新增「匯出為 Excel」按鈕，並綁定到 export_to_excel 方法。
結果： 應用程式具備將資料匯出為標準 Excel 檔案的能力。
最終狀態： 應用程式是一個使用 customtkinter 和 tkinter.ttk 混合構建的 GUI 工具，具備待辦事項的新增、列表顯示（帶到期日、狀態、備註預覽）、多狀態管理、雙擊編輯、圖形化日曆選取日期、帶超連結的備註詳細資訊查看、Ctrl+S 快捷鍵儲存以及匯出為 Excel 檔案等功能。支援在備註欄位使用 Enter 鍵進行換行，Ctrl+Enter 進行儲存。
