# Cody Wealth AppScript

## 日常流程
1. 編輯程式碼。
2. 使用 `clasp push` 將本地變更推送到 Google Apps Script。
3. 使用 `git commit` 和 `git push` 將變更提交並推送到遠端 Git 儲存庫。
4. 將專案資料夾壓縮備份（zip）。

## 開機後該怎麼做
1. 開啟終端機或命令提示字元。
2. 執行 `cd ~/cody-wealth-appscript` 進入專案目錄。
3. 開始進行編輯或其他工作。

## cw alias 的作用與設定
- **作用**：`cw` 是一個自訂的 shell alias，用來快速切換到專案目錄，節省輸入完整路徑的時間。
- **設定方法**：
  - 在你的 shell 設定檔（例如 `.bashrc`、`.zshrc`）中加入以下一行：
    ```bash
    alias cw='cd ~/cody-wealth-appscript'
    ```
  - 設定後，執行 `source ~/.bashrc` 或 `source ~/.zshrc`（依你的 shell 而定）使設定生效。
  - 之後只要輸入 `cw`，即可快速切換到專案目錄。

## 指令範例與平台差異
- **clasp push**：將本地程式碼推送到 Google Apps Script。
- **git commit -m "message"**：提交變更。
- **git push**：推送到遠端 Git 儲存庫。
- **zip 備份**：
  - Mac/Linux：
    ```bash
    zip -r backup.zip ~/cody-wealth-appscript
    ```
  - Windows（PowerShell）：
    ```powershell
    Compress-Archive -Path C:\Users\cody\cody-wealth-appscript -DestinationPath backup.zip
    ```
