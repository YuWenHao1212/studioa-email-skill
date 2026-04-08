# STUDIO A Email Skill

**STUDIO A 專用的 email skill** — fork 自 [claude-email-skill](https://github.com/YuWenHao1212/claude-email-skill)，差別在於 **草稿模式**：

- **主線 claude-email-skill**：用 IMAP `APPEND` 把草稿寫到 server 的 `Drafts` folder
- **本 repo（STUDIO A）**：產 `.eml` 檔到本機 + 自動開 Apple Mail compose 視窗，讓使用者按 Cmd+S 存進**本機草稿匣**

## 為什麼分開

STUDIO A 學員的工作環境：
- Mail server：Nusoft MLS-2500（自建）
- 本機 client：Apple Mail，且**草稿存本機**（不從 server 同步）
- → 主線 skill 寫到 server `Drafts` 學員看不到

這個 fork 改 `cmd_draft` / `cmd_reply` 不再 IMAP `APPEND`，改產 .eml + `open -a Mail`。

## 已知限制

| 內容類型 | Cmd+S 行為 | 實際動作 |
|---|---|---|
| 純文字 | 一鍵存進本機草稿匣 | 1 步 |
| HTML（含表格） | 跳「另存新檔」對話框 | 選 folder + 關視窗 = 多 2 步 |

HTML 場景的多 2 步是 Apple Mail 對外部 .eml 檔處理機制的限制（已驗證走不通的方案：AppleScript `htmlcontent`、剪貼簿 hack、`make new outgoing message` + `content` 都不支援帶樣式 HTML）。詳見 STUDIO A 內部 issue 文件。

## 安裝

```bash
git clone https://github.com/YuWenHao1212/studioa-email-skill.git /tmp/studioa-email-skill && cp -r /tmp/studioa-email-skill/email ~/.claude/skills/email
```

或在 Claude Code 內讓 Sonnet 跑 STUDIO A 課程的 `setup-email-skill-for-sonnet.md` prompt（已內建一次到位的 local draft 模式，不需要再跑 upgrade）。

## Setup

安裝後在任何目錄開 Claude Code，說「幫我設定 email」，Sonnet 會引導完成帳號設定。

## 功能

| Feature | 行為 |
|---|---|
| 讀信 | 跟主線一樣，IMAP fetch |
| 寫草稿 (`draft`) | 產 .eml + 開 Apple Mail（**STUDIO A 特有**） |
| 回信 (`reply`) | 讀原信 + 產含引用的 .eml + 開 Apple Mail（**STUDIO A 特有**） |
| 搜尋 / 看信 / 標已讀 | 跟主線一樣 |
| HTML theme (`--html --theme`) | 跟主線一樣，移動裝置安全模板 |

## 安全規則

| 規則 | 怎麼確保 |
|---|---|
| 從不發送 email | **Code-level** — 沒有 `send` 指令 |
| 從不刪除 email | **Code-level** — 沒有 `delete` 指令 |
| Draft-first | 所有產出都到草稿，使用者 review 後手動寄出 |
| 密碼保護 | Sonnet 引導使用者自己編輯 .env.email，從不讀取密碼 |

## 與主線的差異一覽

| 檔案 | 改動 |
|---|---|
| `email/scripts/email_ops.py` | 新增 `save_local_draft_and_open()` helper、`_load_user_address()` helper；改寫 `cmd_draft` 不連 IMAP；改寫 `cmd_reply` 在讀完原信後 logout 並改走 local 模式 |

## License

MIT
