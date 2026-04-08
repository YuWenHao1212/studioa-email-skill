# Email Skill — 安裝指南

這個 repo 是 email skill 的安裝源。

## 安裝

當使用者說「安裝 email skill」、「install email」或任何安裝相關的請求時：

1. 把本 repo 下的 `email/` 資料夾整個複製到 `~/.claude/skills/email/`：
   ```bash
   cp -r email/ ~/.claude/skills/email/
   ```

2. 確認複製成功：
   ```bash
   ls ~/.claude/skills/email/SKILL.md
   ```

3. 告訴使用者：
   ```
   Email skill 已安裝完成。

   接下來請在任何目錄開啟 Claude Code，說「幫我設定 email」，
   我會引導你完成 email 帳號設定。
   ```

## 注意

- 如果 `~/.claude/skills/email/` 已存在，詢問使用者是否要覆蓋
- 安裝完成後，這個 repo 資料夾可以刪除
- Skill 安裝到全域（`~/.claude/skills/`），在任何目錄都能使用
