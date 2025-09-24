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
npm run pull
cd ~/cody-wealth-appscript
npm run push
cd ~/cody-wealth-appscript
grep -n '<div class="title">' page_dashmain.html
