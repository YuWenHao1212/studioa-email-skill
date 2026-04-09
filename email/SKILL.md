---
name: email
description: Email assistant — read, reply, draft, search emails via IMAP (Gmail & Outlook). This skill should be used when the user wants to read, write, search, reply, or manage emails. Triggers on "check email", "read email", "draft", "reply", "help me with this email", or any email-related request.
---

# Email 工作流

## 環境設定

讀取同目錄下的 `env.md` 取得 python 指令、email_ops.py 路徑、帳號名稱。

如果 `env.md` 不存在，執行下方的「首次設定」。如果存在，跳到「工具指令」直接開始工作。

---

## 首次設定（僅在 env.md 不存在時執行）

### Step 1：確認 Python

執行 `python3 --version`。如果失敗，試 `python --version`。
記住能用的指令。如果都沒有，告訴使用者安裝 Python 3.8+：
- Mac：通常已內建
- Windows：到 https://www.python.org/downloads/ 下載，安裝時勾選「Add Python to PATH」

### Step 2：確認 email_ops.py 路徑

email_ops.py 在此 SKILL.md 同目錄下的 `scripts/email_ops.py`。
用此 SKILL.md 所在目錄拼出絕對路徑，確認檔案存在。

### Step 3：建立 .env.email

檢查 `scripts/` 目錄下是否有 `.env.email`（email_ops.py 讀的是同目錄下的 .env.email）。

如果不存在，引導使用者：

```
需要先設定你的 email 帳號：

1. 複製模板：
   cp {此skill目錄}/assets/env.email.template {此skill目錄}/scripts/.env.email

2. 用文字編輯器（VS Code、記事本等）打開 scripts/.env.email，填入：
   - work_PROVIDER：填 gmail 或 outlook
   - work_USER：你的 email 地址
   - work_PASSWORD：你的 App Password

   App Password 不是登入密碼，取得方式見 .env.email 裡的說明。

3. 儲存後告訴我「設定好了」。
```

**安全規則：不要讓使用者在對話中輸入密碼。引導他們自己用文字編輯器填寫。不要用 Read 工具讀取 .env.email。**

### Step 4：測試連線

```bash
{python} {email_ops_path} status
```

確認回傳 `"status": "ok"`。
如果失敗，讀取 `references/test-commands.md` 的常見錯誤段落協助排除。

### Step 5：確認草稿匣

```bash
{python} {email_ops_path} list_folders {account}
```

確認能找到草稿匣資料夾。email_ops.py 有自動偵測功能（支援 Gmail 中文介面）。
如果仍然有問題，引導使用者在 `scripts/.env.email` 加上 `work_DRAFTS_FOLDER=正確名稱`。

### Step 6：產出 env.md

在此 SKILL.md 同目錄下建立 `env.md`：

```
python: {確認過的 python 指令}
email_ops: {email_ops.py 的絕對路徑}
account: {帳號名稱}
```

完成後告訴使用者：「Email 設定完成。以後說『幫我看這封信』就能開始。」

---

## 工具指令

讀取 `env.md` 取得 python、email_ops、account 的值。

執行格式：
```bash
{python} {email_ops} <command> [args]
```

### 指令速查

| 指令 | 用途 |
|------|------|
| `status` | 查看各帳號未讀數 |
| `check {account} [limit]` | 列出未讀信件 |
| `recent {account} [limit]` | 列出最近 N 封信（已讀＋未讀，預設 3） |
| `read {account} <id>` | 讀一封信的完整內容 |
| `search {account} <query> [limit]` | 搜尋信件（支援中文） |
| `draft {account} <to> <subject> <body> [cc] [--html] [--theme] [--attach file]` | 產草稿。`to` 和 `cc` 都支援多收件人（逗號或分號分隔） |
| `reply {account} <id> <body> [--all] [--html] [--theme] [--attach file]` | 回覆（自動帶入原信引用：plain 用 `> ` 前綴，HTML 用 `<div>` 縮排） |
| `forward {account} <id> <to> [note] [--cc cc] [--html] [--theme] [--attach file]` | 轉寄一封信給其他人。`to` / `cc` 支援多收件人。⚠️ 目前只轉寄文字內容，不會帶原信附件（見下方 Forward 限制） |
| `mark_read {account} <id> [id...]` | 標記已讀 |
| `list_folders {account}` | 列出所有信箱資料夾 |

> **兩種寫法都接受**：positional 或 `--flag` 皆可，也可混用。
> - Positional：`email_ops.py draft work "user@email.com" "主旨" "內文"`
> - Flag：`email_ops.py draft --account work --to "user@email.com" --subject "主旨" --body "內文"`
> - 混合：`email_ops.py draft work --to "user@email.com" --subject "主旨" --body "內文"`
>
> draft 的 `to` 和 `cc` 會驗證 email 格式，寫反會報錯。
>
> 查看完整說明：`email_ops.py --help` 或 `email_ops.py draft --help`

## Forward（轉寄）

### 使用時機

`forward` 跟 `reply` 是不同動作：

- **reply** = 回給原寄件人（繼續同一個對話 thread）
- **forward** = 把現有 email 轉發給**新的收件人**（開新 thread）

常見 forward 場景：
- 講師寄課程提案 → 轉寄給老闆 / 部門決策者
- 客戶投訴 → 轉寄給相關同事處理
- 會議通知 → 轉寄給沒被加進去的同事
- 學員報名確認 → 轉寄給教務 / 會計

使用者說「**把這封信轉給 X**」「**幫我把這封寄給 Y 看**」「**forward 給 ___**」時用這個指令。

### 指令格式

```bash
# 最簡：轉寄給單一人，無額外說明
email_ops.py forward {account} <msg_id> <to>

# 多收件人（to / cc 都支援逗號或分號分隔）
email_ops.py forward {account} <msg_id> "a@x.com,b@y.com" --cc "c@z.com;d@w.com"

# 加轉寄說明（會顯示在原信內容上方）
email_ops.py forward {account} <msg_id> <to> "請幫忙看一下這個"
# 或用 flag
email_ops.py forward {account} <msg_id> <to> --note "請幫忙看一下這個"

# HTML 格式
email_ops.py forward {account} <msg_id> <to> "說明" --html --theme
```

### Subject 自動處理

- 原主旨 `課程提案` → forward 後 `Fwd: 課程提案`
- 原主旨已經有 `Fwd: ` 前綴 → 不會重複加

### ⚠️ Tier 1 限制：不帶原信附件

**目前版本只轉寄文字內容**，**不會**把原信的附件一起轉過去。如果原信有附件：

1. script 會在輸出中列出原信的附件名稱（`original_attachments` 欄位）
2. 會附上一則 `attachment_hint`：建議使用者在 Apple Mail 草稿視窗**手動拖曳附件進去**

> **注意**：`forward` 的 `--attach` 參數是掛**新的檔案**到轉寄草稿（例如你想在轉寄時加一份補充資料），**不是**轉寄原信附件。原信附件在 Tier 1 永遠要手動處理。

使用者流程範例：
- 使用者說「把講師寄的課程提案轉給老闆看」
- 原信有 `proposal.pdf` 附件
- 執行 `forward` → 產出草稿（只有文字），輸出告知「原信有 1 個附件 proposal.pdf 未轉寄」
- 告訴使用者：「草稿已建好，原信的 proposal.pdf 請在 Apple Mail 手動拖進去再送出」

### Forward body 格式

轉寄的 body 結構：

```
（使用者提供的 note，如果有）

---------- Forwarded message ----------
From: 原寄件人 <foo@bar.com>
Date: 原信日期
Subject: 原主旨
To: 原收件人
Cc: 原 cc（如果有）

[原信完整內容]
```

這是業界標準（Gmail / Outlook / Apple Mail 都這樣）。

---

## STUDIO A 草稿模式（重要：跟主線 claude-email-skill 不同）

本 fork 的 `draft` 和 `reply` **不再用 IMAP `APPEND` 寫到 server `Drafts`**，因為 STUDIO A 學員的 Apple Mail 設定為「草稿存本機」，看不到 server 上的草稿。

改為**混合 dispatch**：

| 內容 | 走的路徑 | 使用者體驗 |
|---|---|---|
| 純文字（無附件） | AppleScript `make new outgoing message` | Cmd+S **一鍵存草稿匣** |
| HTML（含表格） | 寫 `.eml` → `open -a Mail` | Cmd+S 跳「另存新檔」對話框 → 選 folder → 關視窗（多 2 步） |
| 純文字 + 附件 | `.eml` 路徑 | 同 HTML，多 2 步 |

**為什麼 HTML 場景多 2 步**：Apple Mail 對 .eml 檔案是當外部檔案開啟（read-only viewer），不是 draft 狀態。所有試過的 Apple-supported 替代方案（AppleScript `htmlcontent`、剪貼簿 hack、`make new outgoing message` + `content`）都無法同時做到「HTML 樣式正確 + 直接存草稿匣」。詳見 STUDIO A 內部 issue。

**對 Claude 的影響**：
- 純文字場景跟學員說「按 Cmd+S 存草稿」即可
- HTML 場景要明確告訴學員「會跳對話框，請選 folder + 關視窗才存得進去」
- 不要承諾「一鍵存草稿」這種話如果是 HTML 場景

## 多帳號 dispatch（重要）

`.env.email` 每個帳號都需要 `<NAME>_APPLE_SENDER` 欄位，格式為 `Full Name <email@example.com>`，必須跟 Apple Mail 帳號設定的 Full Name 和 Email Address 完全一致。

範例 `.env.email`（雙帳號）：
```
ACCOUNTS=personal,training

personal_PROVIDER=gmail
personal_HOST=studioa.com.tw
personal_PORT=143
personal_USER=kat.chang
personal_PASSWORD=...
personal_APPLE_SENDER=Kat Chang <kat.chang@studioa.com.tw>

training_PROVIDER=gmail
training_HOST=studioa.com.tw
training_PORT=143
training_USER=training
training_PASSWORD=...
training_APPLE_SENDER=STUDIO A Training <training@studioa.com.tw>
```

**dispatch 邏輯**：
- 純文字路徑（AppleScript）：把 `<NAME>_APPLE_SENDER` 帶進 `make new outgoing message` 的 `sender` 屬性 → 草稿直接存到對應帳號的本機草稿匣
- HTML 路徑（.eml）：`From:` header 寫 `<NAME>_USER` → Apple Mail 開啟 .eml 時識別 From → Cmd+S 對話框會預設展開到對應帳號的草稿匣

**對 Claude 的影響**：使用者說「幫我用 personal 寄信給 X」或「用 training 帳號發通知」時，把對應的 account_name 傳給 `email_ops.py draft <name> ...`。

## Sender 選擇邏輯（回信 / 寄信時怎麼決定用哪個帳號）

當使用者有多個帳號時，**不要默默選 sender，也不要每次都空白問**。按以下順序判斷：

### 1. 使用者明確指定

使用者說「**用 personal 回**」「**用 training 帳號寫**」「**幫我用部門信箱回覆**」— 直接用指定的帳號，不用問。

### 2. 回信情境（reply）

使用者說「**回這封信**」「**幫我回覆**」「**對這封信寫回應**」時：

1. 先看**原信是從哪個帳號讀到的**（看 recent / read / search 時用的 account_name，或看原信的收件人地址）
2. **建議**用同一個帳號回，但**要跟使用者確認**
3. 話術範例：
   > 「這封信是從 **personal** 帳號讀到的，我建議用 personal 回。要用 personal 嗎，還是妳要改用 training？」

**不要默默選**。即使 99% 的情況是「讀什麼就從什麼回」，還是要確認 — 因為反例很關鍵：
- 個人信箱收到部門業務相關的信（講師寄錯地方）→ 應該用 training 回（代表部門立場）
- 部門信箱收到個人推薦的諮詢 → 可能用 personal 回（建立個人關係）
- 寄錯帳號會影響對外信譽、法律身分、信件 threading

### 3. 新寫信情境（不是 reply）

使用者說「**幫我寫一封信給 X**」「**幫我擬一封通知**」而沒指定帳號時：

- **主動問**：「要用哪個帳號寄？personal 還是 training？」
- **不要預設用第一個帳號**
- 如果有明確語境線索可以推論（例如「幫我用我個人名義寫一封推薦信」），可以建議但仍要確認

### 4. 為什麼這樣設計

這套邏輯符合 **Human-in-the-loop (HITL)** 原則：
- AI 有 context awareness（知道原信來自哪）
- AI 提供 sane default（建議同一帳號回）
- **但使用者永遠是最後決策點**（確認 sender）

寄錯帳號的代價（對外信譽、法律身分、thread 斷掉）**遠高於問一句話的成本**。

## HTML 信件規則

產 HTML body 時**必須遵守**：

- **不要用 `<blockquote>`** — iOS Mail 會渲染成帶色條的縮排引用區塊，手機看起來格式錯亂
- **不要用 `<h1>`~`<h6>`** — 各 mail client 渲染差異大，用 `<strong>` 或 `<p style="font-size:18px;font-weight:bold;">` 代替
- **所有樣式用 inline style** — mail client 不吃 `<style>` block 或 class
- **保持扁平結構** — 不要巢狀 `<div>`，用 `<p>` + `<br>` 控制排版
- **加 `--theme`** 會自動套用 email-safe 的 table layout 模板，建議 HTML 信件都加
- 程式會自動把 `<blockquote>` 替換成 `<div>`（防呆），但產 body 時就不要用

## 自建 Mail Server（非 Gmail / Outlook）

如果使用者的 email 不是 Gmail 或 Outlook（例如公司自建 mail server），需要手動設定 `.env.email`：

```
work_PROVIDER=gmail
work_HOST=mail.company.com
work_PORT=143
work_SECURITY=starttls
work_USER=帳號
work_PASSWORD=密碼
```

- `work_PROVIDER=gmail` 不用改（只是 fallback 預設值，實際用 HOST/PORT）
- `work_SECURITY`：port 993 自動用 SSL，port 143 自動用 STARTTLS，通常不用設
- 帳號格式依 server 而定（可能是 `user`、`user@domain`、或其他）
- 密碼是登入 webmail 的密碼，不是 Google App Password

## HTML Sanitization（重要：攻擊者控制的 HTML）

本 fork 會把**原始信件的 HTML body** 透過 `bleach` 過濾後才嵌進轉寄或回覆草稿。原因：惡意寄件者可能在 HTML 裡藏 `<script>`、`<img onerror=>`、`javascript:` URL、`<iframe>`、CSS expression 等 payload — 如果原封不動轉寄給客戶，你會變成攻擊者的跳板，payload 沿著「你 → 客戶」的信任鏈傳播。

### 運作流程

- **使用者/助手生成的 HTML**（`cmd_draft` 的 body、`cmd_reply` 的 body、`cmd_forward` 的 note）→ 只跑 `sanitize_html`（= `rewrite_blockquotes_for_ios`），做 iOS 樣式修補
- **原信 HTML**（`orig_html` 來自 IMAP fetch）→ 跑 `sanitize_external_html`，用 bleach + tinycss2 白名單過濾：
  - **允許 tag**：`p br div span strong em b i u font a img ul ol li blockquote pre code table thead tbody tfoot tr th td caption col colgroup h1-h6 hr`
  - **允許 presentational attr**：`width height align valign bgcolor dir style`（幾乎每個 HTML 元素都有）+ `colspan rowspan scope headers` on `td/th` + `border cellpadding cellspacing summary` on `table` + `href title` on `a` + `src alt title` on `img` + `color face size` on `font`
  - **允許 protocol**：`http https mailto`（不含 `cid:`，因為 Tier 1 forward 不重組 inline MIME parts，留著會是 broken image reference）
  - **允許 CSS property**：`color background-color` + font 系列 + text 系列 + `border*` + `padding*` + `margin*` + `width height` + `display visibility` + list/table/vertical-align 等。**排除** `position float background-image behavior expression` 等攻擊媒介
- **原信純文字**（`orig_plain`）→ `html_lib.escape()` 後包 `<pre>` 再嵌入

### 依賴

安裝方式(從 repo):
```bash
pip3 install -r ~/GitHub/studioa-email-skill/email/requirements.txt
```

或手動:
```bash
pip3 install 'bleach>=6.0' 'tinycss2>=1.2'
```

**降級策略**:
- **bleach 未安裝** → `HAS_BLEACH=False`,`sanitize_external_html` 全部 fallback 成 `<pre>{html_lib.escape(...)}</pre>`。安全但 forward HTML 完全變純文字
- **bleach 有裝但 tinycss2 沒裝** → `HAS_CSS_SANITIZER=False`,bleach 會警告並剝除所有 `style=` 屬性。安全但失去所有 CSS 樣式(顏色、字型、邊距)
- **都有裝** → 完整過濾,保留合法 presentation

**驗證**:
```bash
python3 -c "import bleach, tinycss2; print('bleach', bleach.__version__, '+', 'tinycss2', tinycss2.__version__, 'OK')"
```

### 為什麼重要

STUDIO A 教學主管會轉寄各種外部來信（講師提案、客戶投訴、廠商通知）給同事、客戶、老闆。如果不 sanitize，任何一封帶 XSS payload 的來信都能變成攻擊武器。這是企業級 email 工具的基本安全要求。

### 已知未覆蓋

- `cmd_draft --html` 的 body 參數沒過 bleach(目前只過 iOS rewrite)— 假設 assistant 生成的 HTML 是可信的。如果直接把外部 HTML 當 draft body 會繞過 sanitize。第三堂課前會修成 defence-in-depth
- `.eml` 檔案權限（未加 chmod 600/700，目前 umask 預設）
- `account_name` path traversal whitelist（目前沒檢查 `../`）
- AppleScript control char 防禦縱深
- `.env.email` 權限檢查
- `--attach` 路徑白名單

上述六項已列 tech debt，第三堂課前補完。

---

## 安全規則

- **草稿優先**：所有信件先存草稿匣，不自動寄出（email_ops.py 沒有 send 指令）
- **人工確認**：產完草稿後告訴使用者去草稿匣確認，不要說「已寄出」
- **不刪信**：只標記已讀，不刪除（email_ops.py 沒有 delete 指令）
- **不洩漏**：不讀取、不顯示 .env.email 內容

## 工作流程

### 回信（任何信件）

使用者說「幫我看這封信」或貼了信件內容時：

1. 讀信：用 `read` 或直接讀使用者貼的內容
2. 摘要：告訴使用者這封信在說什麼（重點、對方要什麼、截止日）
3. 討論：問使用者想怎麼回（哪些答應、哪些拒絕、語氣偏好）
4. 出草稿：根據討論寫完整回覆信
5. 微調：使用者提修改意見 → 修改 → 直到滿意
6. 存草稿：用 `draft` 或 `reply` 存到草稿匣
7. 告知：「草稿已存到草稿匣，請確認後寄出。」

### 批次（重複信件）

使用者有大量結構相同的信件時：

1. 確認模板：信件格式、必填欄位
2. 確認資料來源：使用者提供資料表或口述
3. 產第一封：先做一封讓使用者確認
4. 批次產出：確認後逐封存草稿匣
5. 回報：產了幾封、收件人各是誰

### 看最近的信

使用者說「最近的信」「最近幾封信」「幫我看信」（未指定未讀）時：

1. 用 `recent` 列出最近信件（預設 3 封）
2. 簡要告訴使用者有哪些信
3. 使用者選一封 → 用 `read` 讀取

### 查信

使用者問「有沒有某人寄來的信」時：

1. 用 `search` 搜尋
2. 列出結果（寄件人、主旨、日期）
3. 使用者選一封 → 用 `read` 讀取

## 信件風格指南

使用者可以請你在此段落加入：語氣偏好、簽名檔、常用信件模板。
