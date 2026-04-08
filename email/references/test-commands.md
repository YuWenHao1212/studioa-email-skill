# Email 工作流 — 測試指令

設定完成後，依序執行以下指令驗證每個功能都正常。

> Mac 用 `python3`，Windows 用 `python`。下面統一寫 `python3`。

## Step 0：確認 Python 已安裝

```bash
python3 --version
```

預期：`Python 3.8.x` 以上。

如果找不到 `python3`，試試 `python --version`。
如果都沒有，到 https://www.python.org/downloads/ 安裝。
Windows 安裝時勾選「Add Python to PATH」。

## Step 1：確認連線

```bash
python3 email_ops.py status
```

預期回應：
```json
{
  "work": {"unread": 5, "status": "ok"}
}
```

看到 `"status": "ok"` 表示連線成功。

## Step 2：列出信箱資料夾

```bash
python3 email_ops.py list_folders work
```

預期：會列出你信箱中所有資料夾。
找到含有 "Draft" 或 "草稿" 的資料夾 — 這就是你的草稿匣名稱。

如果後續草稿存不進去，把這個名稱填到 `.env.email`：
```
work_DRAFTS_FOLDER=你看到的資料夾名稱
```

## Step 3：列出未讀信件

```bash
python3 email_ops.py check work 5
```

會列出最近 5 封未讀信件的寄件人、主旨、日期。

## Step 4：列出最近的信

```bash
python3 email_ops.py recent work 3
```

會列出最近 3 封信件（不分已讀未讀），最新的排最前面。

## Step 5：讀一封信

�� Step 3 或 Step 4 的結果中選一封，用它的 id：

```bash
python3 email_ops.py read work <id>
```

會顯示完整信件內容。如果信件有附件，會列出附件檔名。

## Step 6：搜尋信件

```bash
python3 email_ops.py search work "報價" 5
```

支援中文搜尋。會按主旨和寄件人搜尋。

## Step 7：產一封測試草稿

```bash
python3 email_ops.py draft work "你自己的email" "測試信件" "這是一封測試草稿，可以刪除。"
```

> 收件人填你自己的 email，這樣測試草稿不會寄給別人。

去 Gmail / Outlook 的草稿匣確認有沒有出現。

## Step 8：帶附件的草稿

準備一個小檔案（例如桌面上的任何 PDF），然後：

```bash
python3 email_ops.py draft work "你自己的email" "附件測試" "請見附件。" --attach /path/to/file.pdf
```

去草稿匣確認附件有沒有出現。

## 測試完成後

1. 去 Gmail / Outlook 草稿匣，手動刪除 Step 7 和 Step 8 產生的測試草稿
2. 所有 Step 都通過 = 設定完成，可以開始使用

## 常見錯誤

| 錯誤訊息 | 原因 | 解法 |
|---------|------|------|
| `LOGIN failed` | 帳號或密碼錯誤 | 確認用的是 App Password，不是登入密碼 |
| `AUTHENTICATIONFAILED` | Gmail 未開 App Password 或未開 2FA | 到 Google 帳號設定開啟兩步驟驗證，再建 App Password |
| `IMAP is disabled` | Gmail IMAP 被關閉 | Gmail 設定 > 轉寄和 POP/IMAP > 啟用 IMAP |
| `.env.email not found` | 設定檔不存在 | 確認已從 .env.email.template 複製為 .env.email |
| `Unknown provider` | .env.email 的 PROVIDER 打錯 | 只能填 `gmail` 或 `outlook` |
| `Connection timed out` | 網路問題或防火牆擋 IMAP | 確認網路連線，公司防火牆可能擋 port 993 |
| 草稿沒出現在草稿匣 | 草稿匣路徑不對（中文介面常見） | 跑 `list_folders` 找到正確名稱，設定 `work_DRAFTS_FOLDER` |
| `python3: command not found` | Python 未安裝或未加入 PATH | 安裝 Python 並勾選 Add to PATH，或改用 `python` |
