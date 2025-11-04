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