#!/usr/bin/env python3
"""
Email operations for Claude Code — Gmail & Outlook unified.
Reads account config from .env.email in the same directory as this script.

Supported providers: Gmail, Google Workspace, Outlook, Microsoft 365.
All operations use standard IMAP protocol.
"""

import argparse
import imaplib
import email
import html as html_lib
import re
import sys
import os
import json
import mimetypes
import subprocess

# Optional HTML sanitizer dependency. If bleach is available we use it to
# strip dangerous tags/attrs/URL schemes from attacker-controlled HTML
# (original message bodies forwarded to third parties). If bleach is not
# installed the code falls back to plain-text rendering (safe but loses
# styling). See SKILL.md for the security rationale.
try:
  import bleach
  HAS_BLEACH = True
except ImportError:
  HAS_BLEACH = False

# HTML whitelist — conservative, tuned for quoted mail content
SANITIZE_ALLOWED_TAGS = [
  'p', 'br', 'div', 'span',
  'strong', 'em', 'b', 'i', 'u',
  'a', 'img',
  'ul', 'ol', 'li',
  'blockquote', 'pre', 'code',
  'table', 'thead', 'tbody', 'tr', 'th', 'td',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'hr',
]
# Note: 'style' attribute is intentionally NOT whitelisted. bleach 6.x
# requires a separate css_sanitizer (tinycss2) to safely allow style attrs;
# without it style= is stripped entirely. We accept that trade-off — loss
# of inline styling on forwarded HTML is preferable to risking CSS-based
# attacks (expression(), url(javascript:...), data: URLs in background, etc.).
# A future upgrade can add tinycss2 and whitelist safe CSS properties.
SANITIZE_ALLOWED_ATTRS = {
  'a': ['href', 'title'],
  'img': ['src', 'alt', 'title'],
  'table': ['border', 'cellpadding', 'cellspacing'],
  'td': ['colspan', 'rowspan'],
  'th': ['colspan', 'rowspan'],
}
SANITIZE_ALLOWED_PROTOCOLS = ['http', 'https', 'mailto', 'cid']
from datetime import datetime
from pathlib import Path
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate, parseaddr, getaddresses

# --- Configuration ---

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_FILE = os.environ.get("EMAIL_ENV_FILE", os.path.join(SCRIPT_DIR, ".env.email"))
TEMPLATE_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "templates")

# Provider presets
PROVIDERS = {
  "gmail": {
    "host": "imap.gmail.com",
    "port": 993,
    "drafts_folder": "[Gmail]/Drafts",
  },
  "outlook": {
    "host": "outlook.office365.com",
    "port": 993,
    "drafts_folder": "Drafts",
  },
}


def load_env():
  """Load key=value pairs from .env.email file."""
  env = {}
  if not os.path.exists(ENV_FILE):
    print(json.dumps({"error": f".env.email not found at {ENV_FILE}"}))
    sys.exit(1)
  with open(ENV_FILE) as f:
    for line in f:
      line = line.strip()
      if "=" in line and not line.startswith("#"):
        key, val = line.split("=", 1)
        env[key.strip()] = val.strip()
  return env


def get_accounts():
  """Build account registry from .env.email.

  Expected format in .env.email:
    ACCOUNTS=work,personal       (comma-separated account names)
    work_PROVIDER=gmail          (gmail or outlook)
    work_USER=you@gmail.com
    work_PASSWORD=xxxx-xxxx-xxxx-xxxx
    personal_PROVIDER=outlook
    personal_USER=you@company.com
    personal_PASSWORD=xxxx
  """
  env = load_env()
  account_names = [a.strip() for a in env.get("ACCOUNTS", "default").split(",")]
  accounts = {}
  for name in account_names:
    provider_key = env.get(f"{name}_PROVIDER", "gmail").lower()
    provider = PROVIDERS.get(provider_key)
    if not provider:
      print(json.dumps({"error": f"Unknown provider '{provider_key}' for account '{name}'"}))
      sys.exit(1)
    user = env.get(f"{name}_USER", "")
    password = env.get(f"{name}_PASSWORD", "")
    if not user or not password:
      print(json.dumps({"error": f"Missing USER or PASSWORD for account '{name}'"}))
      sys.exit(1)
    # Allow per-account override of drafts folder
    drafts = env.get(f"{name}_DRAFTS_FOLDER", provider["drafts_folder"])
    port = int(env.get(f"{name}_PORT", provider["port"]))
    # Security: explicit override or auto-detect from port
    security = env.get(f"{name}_SECURITY", "").lower()
    if not security:
      security = "ssl" if port == 993 else "starttls"
    accounts[name] = {
      "host": env.get(f"{name}_HOST", provider["host"]),
      "port": port,
      "security": security,
      "user": user,
      "password": password,
      "drafts_folder": drafts,
    }
  return accounts


def connect(account_name):
  """Connect and login to IMAP server. Returns (connection, drafts_folder, user)."""
  accounts = get_accounts()
  if account_name not in accounts:
    available = ", ".join(accounts.keys())
    print(json.dumps({"error": f"Account '{account_name}' not found. Available: {available}"}))
    sys.exit(1)
  acct = accounts[account_name]
  security = acct.get("security", "ssl")
  if security == "none":
    print(json.dumps({"warning": "IMAP connection uses plain text (no encryption). Credentials sent unencrypted."}), file=sys.stderr)
  if security == "ssl":
    try:
      m = imaplib.IMAP4_SSL(acct["host"], acct["port"], timeout=30)
    except TypeError:
      # Python < 3.9: timeout not supported
      m = imaplib.IMAP4_SSL(acct["host"], acct["port"])
      m.socket().settimeout(30)
  else:
    try:
      m = imaplib.IMAP4(acct["host"], acct["port"], timeout=30)
    except TypeError:
      # Python < 3.9: timeout not supported
      m = imaplib.IMAP4(acct["host"], acct["port"])
      m.socket().settimeout(30)
    if security == "starttls":
      m.starttls()
  m.login(acct["user"], acct["password"])
  return m, acct["drafts_folder"], acct["user"]


def detect_drafts_folder(m, configured_folder):
  """Try configured drafts folder, fallback to common alternatives.
  Gmail Chinese UI uses UTF-7 encoded folder names.
  Returns the working drafts folder name."""
  # Common drafts folder names across providers and languages
  candidates = [
    configured_folder,
    "[Gmail]/Drafts",
    "[Gmail]/&g0l6Pw-",       # Gmail Chinese UI (UTF-7 encoded)
    "Drafts",
    "&g0l6Pw-",               # Outlook Chinese
    "INBOX.Drafts",
  ]
  # Deduplicate while preserving order
  seen = set()
  unique = []
  for c in candidates:
    if c not in seen:
      seen.add(c)
      unique.append(c)
  for folder in unique:
    try:
      status, _ = m.select(folder)
      if status == "OK":
        m.select("INBOX")  # Reset selection
        return folder
    except Exception:
      continue
  print(json.dumps({"warning": f"No drafts folder found. Tried: {unique}. Falling back to '{configured_folder}'."}), file=sys.stderr)
  return configured_folder  # Fallback to configured


# --- Helpers ---

def decode_subject(raw_subject):
  """Decode email subject header to string."""
  if not raw_subject:
    return "(no subject)"
  parts = decode_header(raw_subject)
  decoded = []
  for part, charset in parts:
    if isinstance(part, bytes):
      decoded.append(part.decode(charset or "utf-8", errors="replace"))
    else:
      decoded.append(part)
  return "".join(decoded)


def decode_addr(raw_header):
  """Decode a MIME-encoded address header to readable string."""
  if not raw_header:
    return ""
  parts = decode_header(raw_header)
  decoded = []
  for part, charset in parts:
    if isinstance(part, bytes):
      decoded.append(part.decode(charset or "utf-8", errors="replace"))
    else:
      decoded.append(part)
  return "".join(decoded)


MAX_ATTACH_SIZE = 25 * 1024 * 1024  # 25MB per file


def attach_files(msg, file_paths):
  """Attach files to a MIMEMultipart message."""
  for fpath in file_paths:
    resolved = os.path.realpath(fpath)
    # Block dotfiles and dot-directories (e.g. .ssh/, .env, .gnupg/)
    path_parts = resolved.split(os.sep)
    has_dotfile = any(p.startswith(".") and p not in (".", "..") for p in path_parts if p)
    if has_dotfile:
      print(json.dumps({"error": f"Refused to attach dotfile/dotdir path: {fpath}"}), file=sys.stderr)
      continue
    if not os.path.exists(resolved):
      print(json.dumps({"warning": f"File not found: {fpath}"}), file=sys.stderr)
      continue
    file_size = os.path.getsize(resolved)
    if file_size > MAX_ATTACH_SIZE:
      print(json.dumps({"error": f"File too large ({file_size // 1024 // 1024}MB > 25MB): {fpath}"}), file=sys.stderr)
      continue
    mime_type, _ = mimetypes.guess_type(resolved)
    if mime_type is None:
      mime_type = "application/octet-stream"
    main_type, sub_type = mime_type.split("/", 1)
    with open(resolved, "rb") as f:
      part = MIMEBase(main_type, sub_type)
      part.set_payload(f.read())
    encoders.encode_base64(part)
    filename = os.path.basename(resolved)
    part.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(part)


def rewrite_blockquotes_for_ios(html_body):
  """Replace <blockquote> with <div> to avoid iOS Mail rendering bars.

  This is a STYLE fix only — it does NOT sanitize attacker-controlled
  HTML. For that, use sanitize_external_html() which runs bleach.
  """
  html_body = re.sub(
    r'<blockquote[^>]*>',
    '<div style="margin:0;padding:0;">',
    html_body,
    flags=re.IGNORECASE,
  )
  html_body = re.sub(
    r'</blockquote>',
    '</div>',
    html_body,
    flags=re.IGNORECASE,
  )
  return html_body


# Backwards-compat alias — cmd_draft/cmd_reply/cmd_forward call this name
# for locally-produced HTML (assistant-generated body, not external).
sanitize_html = rewrite_blockquotes_for_ios


def sanitize_external_html(html_body):
  """Sanitize attacker-controlled HTML before forwarding or quoting it.

  Use this on the HTML body of an ORIGINAL message (orig_html from IMAP)
  before embedding it in a new draft. Without this a malicious sender can
  inject <script>, javascript: URLs, onerror handlers, etc., that would
  execute in the recipient's mail client when the victim forwards the
  message to a trusted third party.

  Falls back to plain-text escape if bleach is not installed (safe but
  loses styling). See SKILL.md Security section.
  """
  if not html_body:
    return ""
  if HAS_BLEACH:
    return bleach.clean(
      html_body,
      tags=SANITIZE_ALLOWED_TAGS,
      attributes=SANITIZE_ALLOWED_ATTRS,
      protocols=SANITIZE_ALLOWED_PROTOCOLS,
      strip=True,
      strip_comments=True,
    )
  # Fallback: escape the whole thing — no HTML rendering, but safe
  return f"<pre>{html_lib.escape(html_body)}</pre>"


def strip_html_tags(html_body):
  """Convert HTML to readable plain text. No external dependencies."""
  text = re.sub(r'<style[^>]*>.*?</style>', '', html_body, flags=re.DOTALL | re.IGNORECASE)
  text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
  text = re.sub(r'<br\s*/?\s*>', '\n', text, flags=re.IGNORECASE)
  text = re.sub(r'</(p|div|tr|li|h[1-6])>', '\n', text, flags=re.IGNORECASE)
  text = re.sub(r'<[^>]+>', '', text)
  text = text.replace('&nbsp;', ' ').replace('&amp;', '&')
  text = text.replace('&lt;', '<').replace('&gt;', '>')
  text = text.replace('&quot;', '"').replace('&#39;', "'")
  text = re.sub(r'\n{3,}', '\n\n', text)
  return text.strip()


def load_theme():
  """Load HTML email theme template."""
  theme_path = os.path.join(TEMPLATE_DIR, "default.html")
  if os.path.exists(theme_path):
    with open(theme_path) as f:
      return f.read()
  return None


def apply_theme(body_html):
  """Wrap body HTML in theme template. Returns full HTML."""
  template = load_theme()
  if template and "{{BODY}}" in template:
    return template.replace("{{BODY}}", body_html)
  return body_html


# --- Commands ---

def cmd_status(accounts=None):
  """Print unread count for each account. JSON output."""
  all_accounts = get_accounts()
  targets = accounts or list(all_accounts.keys())
  results = {}
  for name in targets:
    m = None
    try:
      m, _, _ = connect(name)
      m.select("INBOX")
      _, data = m.search(None, "UNSEEN")
      count = len(data[0].split()) if data[0] else 0
      results[name] = {"unread": count, "status": "ok"}
    except Exception as e:
      results[name] = {"unread": -1, "status": str(e)}
    finally:
      if m:
        try:
          m.logout()
        except Exception:
          pass
  print(json.dumps(results, indent=2))


def cmd_check(account_name, limit=10, mailbox="INBOX"):
  """List unread emails for an account. JSON output."""
  m, _, _ = connect(account_name)
  try:
    m.select(mailbox)
    _, data = m.search(None, "UNSEEN")
    ids = data[0].split() if data[0] else []
    results = []
    for uid in ids[-limit:]:
      _, msg_data = m.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
      if msg_data and isinstance(msg_data[0], tuple):
        header = email.message_from_bytes(msg_data[0][1])
        results.append({
          "id": uid.decode(),
          "from": decode_addr(header.get("From", "")),
          "subject": decode_subject(header.get("Subject")),
          "date": header.get("Date", ""),
        })
    print(json.dumps(results, indent=2, ensure_ascii=False))
  finally:
    try:
      m.logout()
    except Exception:
      pass


def cmd_recent(account_name, limit=3, mailbox="INBOX"):
  """List the most recent N emails (read or unread). JSON output."""
  m, _, _ = connect(account_name)
  try:
    m.select(mailbox, readonly=True)
    _, data = m.search(None, "ALL")
    ids = data[0].split() if data[0] else []
    results = []
    for uid in reversed(ids[-limit:]):
      _, msg_data = m.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
      if msg_data and isinstance(msg_data[0], tuple):
        header = email.message_from_bytes(msg_data[0][1])
        results.append({
          "id": uid.decode(),
          "from": decode_addr(header.get("From", "")),
          "subject": decode_subject(header.get("Subject")),
          "date": header.get("Date", ""),
        })
    print(json.dumps(results, indent=2, ensure_ascii=False))
  finally:
    try:
      m.logout()
    except Exception:
      pass


def fetch_original_for_quote(m, msg_id, mailbox):
  """Fetch original message body for quoting in reply.
  Returns (plain_body, html_body, sender_display, date_str).
  Any field may be empty string on failure — caller must fail-safe."""
  try:
    m.select(mailbox)
    _, msg_data = m.fetch(msg_id.encode(), "(BODY.PEEK[])")
    if not msg_data or not isinstance(msg_data[0], tuple):
      return "", "", "", ""
    msg = email.message_from_bytes(msg_data[0][1])

    plain_body = ""
    html_body = ""
    if msg.is_multipart():
      for part in msg.walk():
        ct = part.get_content_type()
        cd = str(part.get("Content-Disposition", ""))
        if "attachment" in cd:
          continue
        if ct == "text/plain" and not plain_body:
          payload = part.get_payload(decode=True)
          if payload:
            charset = part.get_content_charset() or "utf-8"
            plain_body = payload.decode(charset, errors="replace")
        elif ct == "text/html" and not html_body:
          payload = part.get_payload(decode=True)
          if payload:
            charset = part.get_content_charset() or "utf-8"
            html_body = payload.decode(charset, errors="replace")
    else:
      payload = msg.get_payload(decode=True)
      if payload:
        charset = msg.get_content_charset() or "utf-8"
        ct = msg.get_content_type()
        decoded = payload.decode(charset, errors="replace")
        if ct == "text/html":
          html_body = decoded
        else:
          plain_body = decoded

    # Fallbacks: ensure both formats are populated when possible
    if not plain_body and html_body:
      plain_body = strip_html_tags(html_body)

    sender_display = decode_addr(msg.get("From", ""))
    date_str = msg.get("Date", "")
    return plain_body, html_body, sender_display, date_str
  except Exception:
    return "", "", "", ""


def cmd_read(account_name, msg_id, mailbox="INBOX"):
  """Read full email content. JSON output."""
  m, _, _ = connect(account_name)
  try:
    m.select(mailbox)
    _, msg_data = m.fetch(msg_id.encode(), "(BODY.PEEK[])")
    if not msg_data or not isinstance(msg_data[0], tuple):
      print(json.dumps({"error": "Message not found"}))
      return
    msg = email.message_from_bytes(msg_data[0][1])

    # Extract body (prefer plain text, fallback to HTML)
    body = ""
    html_body = ""
    attachments = []
    if msg.is_multipart():
      for part in msg.walk():
        ct = part.get_content_type()
        cd = str(part.get("Content-Disposition", ""))
        if "attachment" in cd:
          filename = part.get_filename()
          if filename:
            attachments.append(decode_subject(filename))
          continue
        if ct == "text/plain" and not body:
          payload = part.get_payload(decode=True)
          if payload:
            charset = part.get_content_charset() or "utf-8"
            body = payload.decode(charset, errors="replace")
        elif ct == "text/html" and not html_body:
          payload = part.get_payload(decode=True)
          if payload:
            charset = part.get_content_charset() or "utf-8"
            html_body = payload.decode(charset, errors="replace")
    else:
      payload = msg.get_payload(decode=True)
      if payload:
        charset = msg.get_content_charset() or "utf-8"
        body = payload.decode(charset, errors="replace")

    # Use HTML body if no plain text available (strip tags for readability)
    if not body and html_body:
      body = strip_html_tags(html_body)

    result = {
      "id": msg_id,
      "from": decode_addr(msg.get("From", "")),
      "to": decode_addr(msg.get("To", "")),
      "cc": decode_addr(msg.get("Cc", "")),
      "subject": decode_subject(msg.get("Subject")),
      "date": msg.get("Date", ""),
      "message_id": msg.get("Message-ID", ""),
      "body": body,
    }
    if attachments:
      result["attachments"] = attachments
    print(json.dumps(result, indent=2, ensure_ascii=False))
  finally:
    try:
      m.logout()
    except Exception:
      pass


def cmd_list_folders(account_name):
  """List all IMAP folders for an account. Useful for finding the correct drafts folder name."""
  m, _, _ = connect(account_name)
  try:
    status, folders = m.list()
    results = []
    if status == "OK":
      for f in folders:
        # Parse IMAP LIST response: (flags) "delimiter" "name"
        decoded = f.decode("utf-8", errors="replace")
        # Extract folder name — match any single-char delimiter
        match = re.match(r'\(.*?\)\s+"(.)"\s+(.*)', decoded)
        if match:
          name = match.group(2).strip('"')
          results.append(name)
        else:
          results.append(decoded)
    print(json.dumps({"account": account_name, "folders": results}, indent=2, ensure_ascii=False))
  finally:
    try:
      m.logout()
    except Exception:
      pass


def _applescript_quote(s):
  """Escape a Python string for safe embedding inside an AppleScript double-quoted literal."""
  return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "")


def _save_via_applescript(to_addr, subject, body, cc=None, sender=None):
  """Use AppleScript to create a draft directly in Apple Mail.

  This is the PLAIN-TEXT path. Apple Mail's `make new outgoing message` API
  creates a real draft state — Cmd+S saves directly into the local drafts
  mailbox without any "save as" dialog.

  When `sender` is provided (format: "Full Name <email@example.com>"), the
  draft is dispatched to that account's local drafts mailbox. The sender
  string must match an account that Apple Mail already has configured. If
  omitted, Apple Mail uses the default account.

  Limitation: Apple Mail's `content` property is plain text only. HTML markup
  will appear as raw tags. For HTML, use _save_via_eml() instead.

  Returns (mode_string, opened_in_mail_bool).
  """
  recipient_blocks = []
  for addr in [a.strip() for a in re.split(r'[,;]', to_addr) if a.strip()]:
    recipient_blocks.append(
      f'make new to recipient with properties {{address:"{_applescript_quote(addr)}"}}'
    )
  if cc:
    for addr in [a.strip() for a in re.split(r'[,;]', cc) if a.strip()]:
      recipient_blocks.append(
        f'make new cc recipient with properties {{address:"{_applescript_quote(addr)}"}}'
      )
  recipients_script = "\n    ".join(recipient_blocks)

  # Build the properties block — include sender only if provided
  props = [
    f'subject:"{_applescript_quote(subject)}"',
    f'content:"{_applescript_quote(body)}"',
    "visible:true",
  ]
  if sender:
    props.append(f'sender:"{_applescript_quote(sender)}"')
  props_block = ", ".join(props)

  script = f'''tell application "Mail"
  set newMessage to make new outgoing message with properties {{{props_block}}}
  tell newMessage
    {recipients_script}
  end tell
  activate
end tell'''

  try:
    subprocess.run(["osascript", "-e", script], check=True, capture_output=True)
    return "applescript", True
  except subprocess.CalledProcessError:
    return "applescript", False


def _save_via_eml(msg, account_name):
  """Write MIME message as .eml and open with Apple Mail.

  This is the HTML path. The .eml is opened by Apple Mail as an external file
  (not draft state), so HTML renders correctly with full styling — but Cmd+S
  triggers a "save as" dialog instead of saving to drafts. Users must select
  a folder and close the window to save (2 extra clicks vs the AppleScript
  path).

  This 2-click cost is unavoidable: every Apple-supported alternative we tried
  (AppleScript htmlcontent, clipboard HTML hack, make new outgoing message
  with content) either fails to render HTML or doesn't reach draft state. See
  studio-a/issues/email/20260408_apple-mail-html-draft-limitation.md.

  Returns (eml_path_str, opened_in_mail_bool).
  """
  output_dir = Path.home() / "Documents" / "ai-workspace" / "output"
  output_dir.mkdir(parents=True, exist_ok=True)

  ts = datetime.now().strftime("%Y%m%d-%H%M%S")
  filename = f"draft-{account_name}-{ts}.eml"
  eml_path = output_dir / filename

  with open(eml_path, "wb") as f:
    f.write(msg.as_bytes())

  try:
    subprocess.run(["open", "-a", "Mail", str(eml_path)], check=True)
    opened = True
  except Exception:
    opened = False

  return str(eml_path), opened


def _load_user_address(account_name):
  """Read EMAIL_<ACCOUNT>_USER from .env.email without connecting to IMAP.

  cmd_draft no longer needs IMAP, so we just need the From address.
  """
  user = os.environ.get(f"EMAIL_{account_name.upper()}_USER", "")
  if user:
    return user
  if os.path.exists(ENV_FILE):
    with open(ENV_FILE) as f:
      for line in f:
        line = line.strip()
        if line.startswith(f"EMAIL_{account_name.upper()}_USER="):
          return line.split("=", 1)[1].strip().strip('"').strip("'")
  return ""


def _load_apple_sender(account_name):
  """Read EMAIL_<ACCOUNT>_APPLE_SENDER from .env.email.

  Format: "Full Name <email@example.com>" — must match an account that Apple
  Mail already has configured. Used by AppleScript dispatch to route plain-
  text drafts to the correct account's local drafts mailbox.

  Returns empty string if not configured (caller falls back to Apple Mail's
  default account).
  """
  sender = os.environ.get(f"EMAIL_{account_name.upper()}_APPLE_SENDER", "")
  if sender:
    return sender
  if os.path.exists(ENV_FILE):
    with open(ENV_FILE) as f:
      for line in f:
        line = line.strip()
        if line.startswith(f"EMAIL_{account_name.upper()}_APPLE_SENDER="):
          return line.split("=", 1)[1].strip().strip('"').strip("'")
  return ""


def cmd_draft(account_name, to_addr, subject, body, cc=None, html=False, theme=False, attachments=None):
  """Create a draft email in Apple Mail.

  STUDIO A mode (hybrid):
    - Plain text → AppleScript `make new outgoing message`. Cmd+S saves to
      drafts mailbox in one keystroke.
    - HTML → write .eml + `open -a Mail`. HTML renders correctly but saving
      requires "select folder + close window" (2 extra clicks). Apple Mail
      limitation, see studio-a/issues/email/20260408_apple-mail-html-draft-limitation.md.

  Does NOT connect to IMAP server. Server-side draft sync is unavailable on
  Nusoft (their Apple Mail stores drafts locally).
  """
  user = _load_user_address(account_name)

  has_attachments = attachments and len(attachments) > 0

  if html:
    # HTML path: build MIME and write .eml
    body = sanitize_html(body)
    if theme:
      body = apply_theme(body)

    needs_multipart = has_attachments or theme
    if needs_multipart:
      msg = MIMEMultipart("mixed")
      msg.attach(MIMEText(body, "html", "utf-8"))
      if has_attachments:
        attach_files(msg, attachments)
    else:
      msg = MIMEText(body, "html", "utf-8")

    msg["From"] = user
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    if cc:
      msg["Cc"] = cc

    eml_path, opened = _save_via_eml(msg, account_name)
    output = {
      "mode": "html-eml",
      "path": eml_path,
      "opened_in_mail": opened,
      "account": account_name,
      "to": to_addr,
      "subject": subject,
      "hint": "Apple Mail 已彈出 HTML 郵件視窗。要存草稿請按 Cmd+S → 選 folder → 關視窗（HTML 場景的已知限制，多 2 步）。",
    }
    if has_attachments:
      output["attachments"] = [os.path.basename(f) for f in attachments]
  else:
    # Plain text path: AppleScript direct draft (one-click save)
    if has_attachments:
      # AppleScript draft path can't easily attach files. Fall back to .eml.
      msg = MIMEMultipart("mixed")
      msg.attach(MIMEText(body, "plain", "utf-8"))
      attach_files(msg, attachments)
      msg["From"] = user
      msg["To"] = to_addr
      msg["Subject"] = subject
      msg["Date"] = formatdate(localtime=True)
      if cc:
        msg["Cc"] = cc
      eml_path, opened = _save_via_eml(msg, account_name)
      output = {
        "mode": "plain-eml-with-attachments",
        "path": eml_path,
        "opened_in_mail": opened,
        "account": account_name,
        "to": to_addr,
        "subject": subject,
        "attachments": [os.path.basename(f) for f in attachments],
        "hint": "因為帶附件，走 .eml 路徑。要存草稿請按 Cmd+S → 選 folder → 關視窗。",
      }
    else:
      sender = _load_apple_sender(account_name)
      mode, opened = _save_via_applescript(to_addr, subject, body, cc=cc, sender=sender)
      output = {
        "mode": "plain-applescript",
        "opened_in_mail": opened,
        "account": account_name,
        "sender": sender or "(default)",
        "to": to_addr,
        "subject": subject,
        "hint": "Apple Mail 已彈出新郵件視窗。按 Cmd+S 直接存進本機草稿匣（一鍵）。",
      }

  print(json.dumps(output, ensure_ascii=False))


def cmd_reply(account_name, msg_id, body, reply_all=False, html=False, theme=False, attachments=None, mailbox="INBOX"):
  """Create a reply draft with proper threading headers."""
  m, drafts_folder, user = connect(account_name)
  try:
    drafts_folder = detect_drafts_folder(m, drafts_folder)

    m.select(mailbox)
    _, msg_data = m.fetch(msg_id.encode(), "(BODY.PEEK[HEADER.FIELDS (FROM TO CC SUBJECT MESSAGE-ID REFERENCES)])")
    if not msg_data or not isinstance(msg_data[0], tuple):
      print(json.dumps({"error": "Original message not found"}))
      return

    orig = email.message_from_bytes(msg_data[0][1])
    orig_from = orig.get("From", "")
    orig_to = orig.get("To", "")
    orig_cc = orig.get("Cc", "")
    orig_subject = decode_subject(orig.get("Subject"))
    orig_msg_id = orig.get("Message-ID", "")
    orig_refs = orig.get("References", "")

    reply_subject = orig_subject if orig_subject.startswith("Re:") else f"Re: {orig_subject}"
    _, reply_to_addr = parseaddr(orig_from)

    cc_addrs = ""
    if reply_all:
      all_addrs = []
      for header_val in [orig_to, orig_cc]:
        if header_val:
          all_addrs.append(decode_addr(header_val))
      combined = ", ".join(all_addrs)
      parsed = getaddresses([combined])
      filtered = [
        f"{name} <{addr}>" if name else addr
        for name, addr in parsed
        if addr.lower() != user.lower() and addr.lower() != reply_to_addr.lower()
      ]
      cc_addrs = ", ".join(filtered) if filtered else ""

    references = f"{orig_refs} {orig_msg_id}".strip() if orig_refs else orig_msg_id

    # Fetch original body for quoting (fail-safe: empty strings if anything goes wrong)
    orig_plain, orig_html, sender_display, date_str = fetch_original_for_quote(m, msg_id, mailbox)
    quote_header = f"On {date_str}, {sender_display} wrote:" if date_str or sender_display else "Original message:"

    if html:
      body = sanitize_html(body)
      if theme:
        body = apply_theme(body)
      # Build HTML quote block — use <div> not <blockquote> (iOS Mail renders blockquote poorly)
      # Escape quote_header (contains sender display name) and run
      # attacker-controlled orig_html through bleach.
      esc_quote_header = html_lib.escape(quote_header)
      if orig_html:
        quoted_html = sanitize_external_html(orig_html)
      elif orig_plain:
        quoted_html = f"<pre>{html_lib.escape(orig_plain)}</pre>"
      else:
        quoted_html = ""
      if quoted_html:
        body = (
          f"{body}<br><br>"
          f'<div style="border-left:2px solid #ccc;padding-left:10px;color:#666;">'
          f"<p>{esc_quote_header}</p>{quoted_html}</div>"
        )
    else:
      # Plain text quote block — RFC 3676 "> " prefix
      if orig_plain:
        quoted_lines = "\n".join("> " + line for line in orig_plain.split("\n"))
        body = f"{body}\n\n{quote_header}\n{quoted_lines}"

    has_attachments = attachments and len(attachments) > 0
    needs_multipart = has_attachments or (html and theme)

    if needs_multipart:
      msg = MIMEMultipart()
      content_type = "html" if html else "plain"
      msg.attach(MIMEText(body, content_type, "utf-8"))
      if has_attachments:
        attach_files(msg, attachments)
    else:
      content_type = "html" if html else "plain"
      msg = MIMEText(body, content_type, "utf-8")

    msg["From"] = user
    msg["To"] = reply_to_addr
    msg["Subject"] = reply_subject
    msg["Date"] = formatdate(localtime=True)
    if cc_addrs:
      msg["Cc"] = cc_addrs
    if orig_msg_id:
      msg["In-Reply-To"] = orig_msg_id
    if references:
      msg["References"] = references

    # STUDIO A mode: logout from IMAP (we already read the original message)
    # then dispatch to AppleScript (plain text) or .eml (HTML) path.
    try:
      m.logout()
    except Exception:
      pass

    if html or has_attachments:
      # HTML or attachments → .eml path (multi-step save)
      eml_path, opened = _save_via_eml(msg, account_name)
      output = {
        "mode": "html-eml" if html else "plain-eml-with-attachments",
        "path": eml_path,
        "opened_in_mail": opened,
        "account": account_name,
        "to": reply_to_addr,
        "cc": cc_addrs,
        "subject": reply_subject,
        "hint": "Apple Mail 已彈出回信視窗。要存草稿請按 Cmd+S → 選 folder → 關視窗（HTML 或附件場景的已知限制，多 2 步）。",
      }
      if has_attachments:
        output["attachments"] = [os.path.basename(f) for f in attachments]
    else:
      # Plain text reply → AppleScript direct draft (one-click save)
      sender = _load_apple_sender(account_name)
      mode, opened = _save_via_applescript(
        reply_to_addr, reply_subject, body,
        cc=cc_addrs if cc_addrs else None,
        sender=sender,
      )
      output = {
        "mode": "plain-applescript",
        "opened_in_mail": opened,
        "account": account_name,
        "sender": sender or "(default)",
        "to": reply_to_addr,
        "cc": cc_addrs,
        "subject": reply_subject,
        "hint": "Apple Mail 已彈出回信視窗。按 Cmd+S 直接存進本機草稿匣（一鍵）。",
      }
    print(json.dumps(output, ensure_ascii=False))
    return
  finally:
    try:
      m.logout()
    except Exception:
      pass


def cmd_forward(account_name, msg_id, to_addr, cc=None, note=None, html=False, theme=False, attachments=None, mailbox="INBOX"):
  """Create a forward draft.

  First version (Tier 1): forwards text content only, no attachments from the
  original message. If the original had attachments, a note in the output hints
  the user to drag them in manually via Apple Mail.

  Supports multiple recipients (to / cc) via comma or semicolon separators;
  the caller is expected to have run them through validate_email first so they
  arrive here already normalised to comma-separated form.
  """
  m, _, user = connect(account_name)
  try:
    m.select(mailbox)
    _, msg_data = m.fetch(msg_id.encode(), "(BODY.PEEK[HEADER.FIELDS (FROM TO CC SUBJECT MESSAGE-ID DATE)])")
    if not msg_data or not isinstance(msg_data[0], tuple):
      print(json.dumps({"error": "Original message not found"}))
      return

    orig = email.message_from_bytes(msg_data[0][1])
    orig_from = decode_addr(orig.get("From", ""))
    orig_to = decode_addr(orig.get("To", ""))
    orig_cc = decode_addr(orig.get("Cc", ""))
    orig_subject = decode_subject(orig.get("Subject"))
    orig_date = orig.get("Date", "")

    # Detect existing forward prefixes from any major mail client to avoid
    # double-prefixing (e.g. "Fwd: FW: ..." when forwarding an Outlook message).
    _subject_lower = orig_subject.lstrip().lower()
    if _subject_lower.startswith(("fwd:", "fw:", "轉寄:", "轉寄：")):
      fwd_subject = orig_subject
    else:
      fwd_subject = f"Fwd: {orig_subject}"

    # Fetch original body for inclusion (fail-safe: empty strings if anything goes wrong)
    orig_plain, orig_html, _, _ = fetch_original_for_quote(m, msg_id, mailbox)

    # Detect attachments in the original (for hint only — Tier 1 does not forward them)
    orig_attachments = []
    try:
      _, full_data = m.fetch(msg_id.encode(), "(BODY.PEEK[])")
      if full_data and isinstance(full_data[0], tuple):
        full_msg = email.message_from_bytes(full_data[0][1])
        if full_msg.is_multipart():
          for part in full_msg.walk():
            cd = str(part.get("Content-Disposition", ""))
            if "attachment" in cd:
              fn = part.get_filename()
              if fn:
                orig_attachments.append(decode_subject(fn))
    except Exception:
      pass

    note_text = note or ""

    # Build forwarded header block
    fwd_header_text = (
      "---------- Forwarded message ----------\n"
      f"From: {orig_from}\n"
      f"Date: {orig_date}\n"
      f"Subject: {orig_subject}\n"
      f"To: {orig_to}\n"
    )
    if orig_cc:
      fwd_header_text += f"Cc: {orig_cc}\n"

    if html:
      body = sanitize_html(note_text)
      if theme:
        body = apply_theme(body)
      # Escape original header values to prevent display names like
      # "Foo <bar@x.com>" from being swallowed as HTML tags.
      esc_from = html_lib.escape(orig_from)
      esc_date = html_lib.escape(orig_date)
      esc_subject = html_lib.escape(orig_subject)
      esc_to = html_lib.escape(orig_to)
      esc_cc = html_lib.escape(orig_cc) if orig_cc else ""
      fwd_header_html = (
        '<div style="border-top:1px solid #ccc;margin-top:20px;padding-top:10px;color:#666;">'
        "<p><strong>---------- Forwarded message ----------</strong></p>"
        f"<p>From: {esc_from}<br>"
        f"Date: {esc_date}<br>"
        f"Subject: {esc_subject}<br>"
        f"To: {esc_to}"
      )
      if esc_cc:
        fwd_header_html += f"<br>Cc: {esc_cc}"
      fwd_header_html += "</p></div>"
      # orig_html is attacker-controlled — run through bleach before embed.
      # If orig has no HTML, fall back to escaped plain text in <pre>.
      if orig_html:
        quoted_html = sanitize_external_html(orig_html)
      elif orig_plain:
        quoted_html = f"<pre>{html_lib.escape(orig_plain)}</pre>"
      else:
        quoted_html = ""
      body = f"{body}<br><br>{fwd_header_html}{quoted_html}"
    else:
      quoted_plain = orig_plain or ""
      body = (
        f"{note_text}\n\n" if note_text else ""
      ) + fwd_header_text + "\n" + quoted_plain

    has_attachments = attachments and len(attachments) > 0
    needs_multipart = has_attachments or (html and theme)

    if needs_multipart:
      msg = MIMEMultipart()
      content_type = "html" if html else "plain"
      msg.attach(MIMEText(body, content_type, "utf-8"))
      if has_attachments:
        attach_files(msg, attachments)
    else:
      content_type = "html" if html else "plain"
      msg = MIMEText(body, content_type, "utf-8")

    msg["From"] = user
    msg["To"] = to_addr
    msg["Subject"] = fwd_subject
    msg["Date"] = formatdate(localtime=True)
    if cc:
      msg["Cc"] = cc

    try:
      m.logout()
    except Exception:
      pass

    if html or has_attachments:
      eml_path, opened = _save_via_eml(msg, account_name)
      output = {
        "mode": "html-eml" if html else "plain-eml-with-attachments",
        "path": eml_path,
        "opened_in_mail": opened,
        "account": account_name,
        "to": to_addr,
        "cc": cc or "",
        "subject": fwd_subject,
        "hint": "Apple Mail 已彈出轉寄視窗。要存草稿請按 Cmd+S → 選 folder → 關視窗（HTML 或附件場景的已知限制，多 2 步）。",
      }
      if has_attachments:
        output["attachments"] = [os.path.basename(f) for f in attachments]
    else:
      sender = _load_apple_sender(account_name)
      mode, opened = _save_via_applescript(
        to_addr, fwd_subject, body,
        cc=cc if cc else None,
        sender=sender,
      )
      output = {
        "mode": "plain-applescript",
        "opened_in_mail": opened,
        "account": account_name,
        "sender": sender or "(default)",
        "to": to_addr,
        "cc": cc or "",
        "subject": fwd_subject,
        "hint": "Apple Mail 已彈出轉寄視窗。按 Cmd+S 直接存進本機草稿匣（一鍵）。",
      }

    if orig_attachments:
      output["original_attachments"] = orig_attachments
      output["attachment_hint"] = (
        f"原信有 {len(orig_attachments)} 個附件未轉寄（Tier 1 限制）：{', '.join(orig_attachments)}。"
        "如需附件，請在 Apple Mail 草稿視窗手動拖曳進去。"
      )

    print(json.dumps(output, ensure_ascii=False))
    return
  finally:
    try:
      m.logout()
    except Exception:
      pass


def cmd_mark_read(account_name, msg_ids, mailbox="INBOX"):
  """Mark messages as read."""
  m, _, _ = connect(account_name)
  try:
    m.select(mailbox)
    for uid in msg_ids:
      m.store(uid.encode(), "+FLAGS", "\\Seen")
    print(json.dumps({"marked_read": len(msg_ids), "account": account_name}))
  finally:
    try:
      m.logout()
    except Exception:
      pass


def cmd_search(account_name, query, limit=10, mailbox="INBOX"):
  """Search emails by subject or from. Supports non-ASCII via client-side filter."""
  m, _, _ = connect(account_name)
  try:
    m.select(mailbox, readonly=True)

    is_ascii = True
    try:
      query.encode("ascii")
    except UnicodeEncodeError:
      is_ascii = False

    if is_ascii:
      safe_query = query.replace('\\', '\\\\').replace('"', '\\"')
      _, data = m.search(None, f'(SUBJECT "{safe_query}")')
      ids = data[0].split() if data[0] else []
      if not ids:
        _, data = m.search(None, f'(FROM "{safe_query}")')
        ids = data[0].split() if data[0] else []
    else:
      scan_count = min(max(limit * 20, 200), 500)
      _, data = m.search(None, "ALL")
      all_ids = data[0].split() if data[0] else []
      candidate_ids = all_ids[-scan_count:]
      ids = []
      for uid in candidate_ids:
        _, msg_data = m.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])")
        if msg_data and isinstance(msg_data[0], tuple):
          header = email.message_from_bytes(msg_data[0][1])
          subj = decode_subject(header.get("Subject"))
          from_decoded = decode_addr(header.get("From", ""))
          if query in subj or query in from_decoded:
            ids.append(uid)
        if len(ids) >= limit:
          break

    results = []
    for uid in ids[-limit:]:
      _, msg_data = m.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
      if msg_data and isinstance(msg_data[0], tuple):
        header = email.message_from_bytes(msg_data[0][1])
        results.append({
          "id": uid.decode(),
          "from": decode_addr(header.get("From", "")),
          "subject": decode_subject(header.get("Subject")),
          "date": header.get("Date", ""),
        })
    print(json.dumps(results, indent=2, ensure_ascii=False))
  finally:
    try:
      m.logout()
    except Exception:
      pass


# --- CLI ---


class JsonErrorParser(argparse.ArgumentParser):
  """ArgumentParser that outputs errors as JSON instead of plain text."""

  def error(self, message):
    print(json.dumps({"error": message, "usage": self.format_usage().strip()}))
    sys.exit(1)


def resolve(*values):
  """Return the first non-None value, or None if all are None."""
  for v in values:
    if v is not None:
      return v
  return None


def validate_email(addr, label="to"):
  """Validate email format. Exit with JSON error if invalid.

  Supports multiple recipients separated by comma or semicolon. Returns the
  normalised comma-separated string (stripped, semicolons converted to commas)
  so callers can pass it downstream unchanged.
  """
  if not addr:
    print(json.dumps({"error": f"{label} is empty."}))
    sys.exit(1)
  # Split on either , or ; then strip whitespace
  parts = [p.strip() for p in re.split(r'[,;]', addr) if p.strip()]
  if not parts:
    print(json.dumps({"error": f"{label} '{addr[:50]}' has no valid addresses."}))
    sys.exit(1)
  for p in parts:
    if not re.match(r'^[^@\s,;]+@[^@\s,;]+\.[a-zA-Z]{2,}$', p):
      print(json.dumps({"error": f"{label} contains invalid email '{p[:50]}'."}))
      sys.exit(1)
  return ", ".join(parts)


def build_parser():
  """Build the argparse parser with all subcommands."""
  parser = JsonErrorParser(
    prog="email_ops.py",
    description="Email operations for Claude Code",
  )
  sub = parser.add_subparsers(dest="command")

  # --- status ---
  p = sub.add_parser("status", help="Check unread counts")
  p.add_argument("accounts", nargs="*", default=None, help="Account names (default: all)")

  # --- check ---
  p = sub.add_parser("check", help="List unread emails")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("limit_pos", nargs="?", default=None, type=int)
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--limit", dest="limit_flag", default=None, type=int)

  # --- recent ---
  p = sub.add_parser("recent", help="List most recent N emails (read+unread)")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("limit_pos", nargs="?", default=None, type=int)
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--limit", dest="limit_flag", default=None, type=int)

  # --- read ---
  p = sub.add_parser("read", help="Read full email content")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("msg_id_pos", nargs="?", default=None)
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--id", dest="id_flag", default=None)

  # --- draft ---
  p = sub.add_parser("draft", help="Create email draft")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("to_pos", nargs="?", default=None)
  p.add_argument("subject_pos", nargs="?", default=None)
  p.add_argument("body_pos", nargs="?", default=None)
  p.add_argument("cc_pos", nargs="?", default=None)
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--to", dest="to_flag", default=None)
  p.add_argument("--subject", dest="subject_flag", default=None)
  p.add_argument("--body", dest="body_flag", default=None)
  p.add_argument("--cc", dest="cc_flag", default=None)
  p.add_argument("--html", action="store_true")
  p.add_argument("--theme", action="store_true")
  p.add_argument("--attach", action="append", default=None, metavar="FILE")

  # --- reply ---
  p = sub.add_parser("reply", help="Reply to an email")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("msg_id_pos", nargs="?", default=None)
  p.add_argument("body_pos", nargs="?", default=None)
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--id", dest="id_flag", default=None)
  p.add_argument("--body", dest="body_flag", default=None)
  p.add_argument("--all", dest="reply_all", action="store_true")
  p.add_argument("--html", action="store_true")
  p.add_argument("--theme", action="store_true")
  p.add_argument("--attach", action="append", default=None, metavar="FILE")

  # --- forward ---
  p = sub.add_parser("forward", help="Forward an email")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("msg_id_pos", nargs="?", default=None)
  p.add_argument("to_pos", nargs="?", default=None)
  p.add_argument("note_pos", nargs="?", default=None, help="Optional note prepended to the forwarded content")
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--id", dest="id_flag", default=None)
  p.add_argument("--to", dest="to_flag", default=None)
  p.add_argument("--cc", dest="cc_flag", default=None)
  p.add_argument("--note", dest="note_flag", default=None)
  p.add_argument("--html", action="store_true")
  p.add_argument("--theme", action="store_true")
  p.add_argument("--attach", action="append", default=None, metavar="FILE")

  # --- mark_read ---
  p = sub.add_parser("mark_read", help="Mark messages as read")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("msg_ids", nargs="*", default=None)
  p.add_argument("--account", dest="account_flag", default=None)

  # --- search ---
  p = sub.add_parser("search", help="Search emails by subject or sender")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("query_pos", nargs="?", default=None)
  p.add_argument("limit_pos", nargs="?", default=None, type=int)
  p.add_argument("--account", dest="account_flag", default=None)
  p.add_argument("--query", dest="query_flag", default=None)
  p.add_argument("--limit", dest="limit_flag", default=None, type=int)

  # --- list_folders ---
  p = sub.add_parser("list_folders", help="List all mailbox folders")
  p.add_argument("account_pos", nargs="?", default=None)
  p.add_argument("--account", dest="account_flag", default=None)

  return parser


if __name__ == "__main__":
  parser = build_parser()
  args = parser.parse_args()

  if not args.command:
    parser.print_help()
    sys.exit(1)

  cmd = args.command

  if cmd == "status":
    accts = args.accounts if args.accounts else None
    cmd_status(accts)

  elif cmd == "check":
    account = resolve(args.account_flag, args.account_pos) or "default"
    limit = resolve(args.limit_flag, args.limit_pos) or 10
    cmd_check(account, int(limit))

  elif cmd == "recent":
    account = resolve(args.account_flag, args.account_pos) or "default"
    limit = resolve(args.limit_flag, args.limit_pos) or 3
    cmd_recent(account, int(limit))

  elif cmd == "read":
    account = resolve(args.account_flag, args.account_pos)
    msg_id = resolve(args.id_flag, args.msg_id_pos)
    if not account or not msg_id:
      print(json.dumps({"error": "Usage: read <account> <msg_id>  or  read --account <name> --id <id>"}))
      sys.exit(1)
    cmd_read(account, msg_id)

  elif cmd == "draft":
    account = resolve(args.account_flag, args.account_pos)
    to_addr = resolve(args.to_flag, args.to_pos)
    subject = resolve(args.subject_flag, args.subject_pos)
    body = resolve(args.body_flag, args.body_pos)
    cc = resolve(args.cc_flag, args.cc_pos)
    if not all([account, to_addr, subject, body]):
      missing = [n for n, v in [("account", account), ("to", to_addr), ("subject", subject), ("body", body)] if not v]
      print(json.dumps({"error": f"Missing required arguments: {', '.join(missing)}",
                         "usage": "draft <account> <to> <subject> <body> [cc] [--html] [--theme] [--attach file]"}))
      sys.exit(1)
    to_addr = validate_email(to_addr, "to")
    if cc:
      cc = validate_email(cc, "cc")
    cmd_draft(account, to_addr, subject, body, cc,
              html=args.html, theme=args.theme,
              attachments=args.attach)

  elif cmd == "reply":
    account = resolve(args.account_flag, args.account_pos)
    msg_id = resolve(args.id_flag, args.msg_id_pos)
    body = resolve(args.body_flag, args.body_pos)
    if not all([account, msg_id, body]):
      missing = [n for n, v in [("account", account), ("id", msg_id), ("body", body)] if not v]
      print(json.dumps({"error": f"Missing required arguments: {', '.join(missing)}",
                         "usage": "reply <account> <msg_id> <body> [--all] [--html] [--theme] [--attach file]"}))
      sys.exit(1)
    cmd_reply(account, msg_id, body,
              reply_all=args.reply_all, html=args.html, theme=args.theme,
              attachments=args.attach)

  elif cmd == "forward":
    account = resolve(args.account_flag, args.account_pos)
    msg_id = resolve(args.id_flag, args.msg_id_pos)
    to_addr = resolve(args.to_flag, args.to_pos)
    cc = args.cc_flag
    note = resolve(args.note_flag, args.note_pos)
    if not all([account, msg_id, to_addr]):
      missing = [n for n, v in [("account", account), ("id", msg_id), ("to", to_addr)] if not v]
      print(json.dumps({"error": f"Missing required arguments: {', '.join(missing)}",
                         "usage": "forward <account> <msg_id> <to> [note] [--cc cc] [--html] [--theme] [--attach file]"}))
      sys.exit(1)
    to_addr = validate_email(to_addr, "to")
    if cc:
      cc = validate_email(cc, "cc")
    cmd_forward(account, msg_id, to_addr, cc=cc, note=note,
                html=args.html, theme=args.theme,
                attachments=args.attach)

  elif cmd == "mark_read":
    account = resolve(args.account_flag, args.account_pos)
    msg_ids = args.msg_ids if args.msg_ids else []
    if not account or not msg_ids:
      print(json.dumps({"error": "Usage: mark_read <account> <msg_id> [msg_id ...]"}))
      sys.exit(1)
    cmd_mark_read(account, msg_ids)

  elif cmd == "search":
    account = resolve(args.account_flag, args.account_pos) or "default"
    query = resolve(args.query_flag, args.query_pos) or ""
    limit = resolve(args.limit_flag, args.limit_pos) or 10
    cmd_search(account, query, int(limit))

  elif cmd == "list_folders":
    account = resolve(args.account_flag, args.account_pos) or "default"
    cmd_list_folders(account)
