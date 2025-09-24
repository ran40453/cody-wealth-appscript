# Cody Wealth AppScript Project

## 日常流程

1. **編輯程式碼**
   - 在本地使用你喜歡的編輯器（如 VSCode）修改 Apps Script 相關檔案。

2. **將程式碼推送到 Google Apps Script**
   - 使用 `clasp` 工具將本地修改同步到 Google Apps Script。

3. **提交並推送到 GitHub**
   - 將修改提交到 Git 本地倉庫，並推送到遠端 GitHub。

4. **備份成 Zip 檔案**
   - 將專案資料夾壓縮成 Zip 檔，方便備份。

---

## 每步驟指令範例

### 1. 編輯程式碼

使用你喜歡的編輯器開啟專案資料夾，進行編輯。

---

### 2. 推送到 Apps Script

```bash
clasp push
```

---

### 3. Git 提交與推送

```bash
git add .
git commit -m "Update script"
git push origin main
```

---

### 4. 備份 Zip 檔

#### Mac/Linux

```bash
zip -r cody-wealth-appscript-backup-$(date +%Y%m%d).zip /Users/cody/cody-wealth-appscript
```

#### Windows (PowerShell)

```powershell
Compress-Archive -Path C:\Users\cody\cody-wealth-appscript\* -DestinationPath C:\Users\cody\cody-wealth-appscript-backup-$(Get-Date -Format yyyyMMdd).zip
```

---

## 環境需求

- **Node.js & npm**  
  需先安裝 Node.js（包含 npm），可從官方網站下載並安裝。

- **clasp**  
  Google Apps Script 的命令列工具，安裝指令：
  ```bash
  npm install -g @google/clasp
  ```

- **git**  
  版本控制工具，請先安裝並設定好。

---

## Windows 與 macOS 差異備註

- **Zip 備份指令不同**  
  macOS/Linux 使用 `zip` 指令，Windows PowerShell 使用 `Compress-Archive`。

- **剪貼簿指令**  
  macOS 可使用 `pbcopy`，Windows 則無法使用，需使用其他工具或手動複製。

- **路徑格式差異**  
  Windows 使用反斜線 `\`，macOS/Linux 使用斜線 `/`。

- **環境變數設定**  
  Windows 與 macOS 設定環境變數方式不同，請依照系統調整。

---

以上為本專案的基本使用與備份流程說明。
