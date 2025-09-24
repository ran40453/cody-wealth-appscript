# Cody Wealth – Apps Script Project

## 開發流程（短版）
1) 在 VS Code 修改
2) `npm run push` 同步到 Apps Script
3) 用 Test deployments 測試
4) `npm run ver --msg="..."` 建版本快照（可選）
5) `git add . && git commit -m "..." && git push`
6) `npm run zip` 打包備份（可選）

## 常用指令
- 同步雲端到本機：`npm run pull`
- 推本機到雲端：`npm run push`
- 建雲端版本：`npm run ver --msg="..."` 
- 壓縮備份：`npm run zip`
