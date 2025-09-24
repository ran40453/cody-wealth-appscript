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

## VS Code 快捷鍵整合
- 你可以在 `.vscode/tasks.json` 設定一個任務，讓 **⌘⇧B (Mac)** 或 **Ctrl+Shift+B (Windows/Linux)** 直接執行 `npx clasp push`。
- 範例設定：
  ```json
  {
    "version": "2.0.0",
    "tasks": [
      {
        "label": "Clasp Push",
        "type": "shell",
        "command": "npx clasp push",
        "group": {
          "kind": "build",
          "isDefault": true
        },
        "problemMatcher": []
      }
    ]
  }
  ```
- 設定完成後，只要在 VS Code 按下快捷鍵，就會自動推送到 Google Apps Script。

## 恢復專案的方法
如果修改後版面壞掉或需要回復舊版本，可以參考以下方式：

### 1. 使用 Git
- 回到最後一次提交：
  ```bash
  git checkout -- .
  ```
- 回到特定版本：
  ```bash
  git log   # 找到 commit id
  git checkout <commit_id> -- .
  ```

### 2. 使用 VS Code Local History
- 在 VS Code 中，右鍵檔案 → **Open Timeline**，可以選擇任意歷史版本還原。

### 3. 使用雲端備份
- 如果專案資料夾有同步 Google Drive / iCloud / Dropbox 等，可從雲端歷史版本還原。

### 4. 沒有備份的情況
- 嘗試在 VS Code 中使用 **Undo (⌘Z / Ctrl+Z)** 一步步退回。
- 或者檢查 macOS Time Machine / Windows 檔案歷史紀錄。

### 建議
將專案交給 Git 管理，哪怕只是自己使用，也能快速回復：
```bash
git init
git add .
git commit -m "init backup"
```
