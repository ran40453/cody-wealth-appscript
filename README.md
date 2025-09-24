# README（給未來的你/同事快速上手）
cat > README.md << 'EOF'
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
EOF

# 小抄（md 版，VS Code 好讀）
cat > CHEATSHEET.md << 'EOF'
# Cody Apps Script 工作流小抄 (macOS)

## 日常
```bash
cd ~/cody-wealth-appscript
npm run push
# Test deployments 測試
npm run ver --msg="修改說明"   # 可選
git add .
git commit -m "feat: 修改說明"
npm run zip                   # 可選

##從雲端拉回
npm run pull
EOF

git add README.md CHEATSHEET.md
git commit -m “docs: add README & workflow cheatsheet”
git push



