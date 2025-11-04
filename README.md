# Cody Wealth – Apps Script Project

## 🚀 開發流程（短版）

1. 在 VS Code 修改程式。
2. `npm run push` 同步到 Apps Script。
3. 進入 Apps Script 使用「Test deployments」測試。
4. 可選：`npm run ver --msg="..."` 建版本快照。
5. `git add . && git commit -m "..." && git push` 推上 GitHub。
6. 可選：`npm run zip` 打包備份。

---

## 💻 常用指令

| 指令 | 說明 |
|------|------|
| `npm run pull` | 從雲端同步到本機。 |
| `npm run push` | 將本機推送到雲端。 |
| `npm run ver --msg="..."` | 建立雲端版本（帶訊息）。 |
| `npm run zip` | 壓縮專案備份。 |

---

## 📦 專案結構
```
/.vscode/          # VS Code 設定（含 tasks.json）
/src/              # 前端頁面與樣式
/app.js            # Apps Script 主後端
/page_*.html       # 各功能頁
/style.html        # 全域樣式
/clasp.json        # Clasp 專案設定
```

---

## 🧩 附註
- 所有 Google Apps Script 檔案皆以 HTML 模組形式分頁維護。
- 主要分頁：`page_input`, `page_routines`, `page_dashmain`, `page_record`, `page_acc`。
- 建議使用 `VS Code + clasp + npm script` 一致開發。
# Cody Wealth – Apps Script Project

## 🚀 開發流程（標準流程）

1. 在 VS Code 修改程式。
2. `npm run push` 同步更新到 Google Apps Script。
3. 到 Apps Script 後台使用「Test deployments」測試。
4. 可選：`npm run ver --msg="..."` 建立雲端版本快照。
5. `git add . && git commit -m "..." && git push` 推上 GitHub。
6. 可選：`npm run zip` 打包專案備份。

---

## 💻 常用指令對照表

| 類別 | 指令 | 說明 |
|------|------|------|
| **雲端同步** | `npm run pull` | 從 Apps Script 雲端同步最新程式到本機。 |
|  | `npm run push` | 將本機更新推送到 Apps Script。 |
|  | `npm run ver --msg="..."` | 建立新的 Apps Script 版本（可附註說明）。 |
|  | `npm run zip` | 壓縮整個專案為 ZIP 備份。 |
| **Apps Script 操作** | `npx @google/clasp status` | 查看本機與雲端的差異。 |
|  | `npx @google/clasp open` | 直接在瀏覽器開啟對應的 Apps Script 專案。 |
|  | `npx @google/clasp deploy --description "deploy"` | 發布新版 Web App。 |
| **GitHub 管理** | `git add .` | 暫存所有修改。 |
|  | `git commit -m "update"` | 建立版本紀錄。 |
|  | `git push` | 推送到 GitHub 遠端。 |
|  | `git push -u origin main` | 首次推送到遠端（只需一次）。 |
|  | `git fetch origin` | 抓取遠端最新資訊。 |
|  | `git pull --rebase origin main --allow-unrelated-histories` | 與遠端同步（避免重疊歷史）。 |
|  | `git reset --hard origin/main` | 強制對齊遠端版本（會覆蓋本機修改）。 |

---

## 📦 專案結構
```
/\.vscode/          # VS Code 設定（含 tasks.json）
/src/              # 前端頁面與樣式
/app.js            # Apps Script 主後端
/page_*.html       # 各功能頁
/style.html        # 全域樣式
/clasp.json        # Clasp 專案設定（每個專案各自一份）
```

---

## 🧩 附註

- 每個 Apps Script 專案都需擁有自己的 `.clasp.json`（指向各自的 scriptId）。
- 所有 Google Apps Script 檔案以 HTML 模組形式分頁維護。
- 主要分頁：`page_input`, `page_routines`, `page_dashmain`, `page_record`, `page_acc`。
- 建議使用 `VS Code + Clasp + npm script` 統一開發流程。
- 若需快速部署新版，可在根目錄執行：
  ```
  npm run push && npx @google/clasp deploy --description "update"
  ```
- Git 指令、Clasp 指令、npm script 皆與其他 Apps Script 專案相容，可通用於多個專案。