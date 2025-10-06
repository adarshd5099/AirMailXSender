# airmailx_fixed.py
import re
import json
import os
import smtplib
import random
import time
import threading
import webview
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ==============================
# Default settings
# ==============================
settings = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "",
    "password": "",
    "subject": "Test Email from AirMailX Sender",
    "body": "<html>Paste your HTML email template here</html>",
    "min_delay": 2,
    "max_delay": 5,
    "receiver_list": [],
}

window = None

# Simple email validation (not perfect, but good enough)
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ==============================
# Logging helper
# ==============================
def timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def safe_eval_js(js_expression: str):
    """Call window.evaluate_js safely (catch exceptions, check window)."""
    global window
    try:
        if window:
            return window.evaluate_js(js_expression)
    except Exception as e:
        # Can't log to GUI (might be shutting down) => fallback to print
        print("evaluate_js failed:", e)

def log_message(msg: str):
    """Log message with timestamp to both console and webview log box."""
    ts_msg = f"[{timestamp()}] {msg}"
    print(ts_msg)
    # Use json.dumps to safely quote string for JS call
    safe_eval_js(f"addLog({json.dumps(ts_msg)})")

# ==============================
# Email sending worker
# ==============================
def send_bulk_emails():
    """Worker that sends emails sequentially with random delay between min and max."""
    try:
        log_message("Connecting to SMTP server...")
        # Create SMTP connection
        server = smtplib.SMTP(settings["smtp_server"], int(settings["smtp_port"]), timeout=30)
        server.ehlo()
        try:
            server.starttls()
            server.ehlo()
        except Exception:
            # Some SMTP servers may use SSL on port 465 (handled below by user settings)
            pass

        server.login(settings["sender_email"], settings["password"])
        log_message("Logged in successfully!")

        total = len(settings["receiver_list"])
        if total == 0:
            log_message("No recipient emails found. Aborting send.")
            return

        for i, recipient in enumerate(settings["receiver_list"], start=1):
            recipient = recipient.strip()
            try:
                msg = MIMEMultipart()
                msg["From"] = settings["sender_email"]
                msg["To"] = recipient
                msg["Subject"] = settings["subject"]
                msg.attach(MIMEText(settings["body"], "html"))


                server.sendmail(settings["sender_email"], recipient, msg.as_string())
                log_message(f"Sent to {recipient} ({i}/{total})")
            except Exception as e:
                log_message(f"Failed to send to {recipient}: {e}")

            # delay
            try:
                mind = float(settings.get("min_delay", 2))
                maxd = float(settings.get("max_delay", 5))
                if maxd < mind:
                    maxd = mind
                delay = random.uniform(mind, maxd)
            except Exception:
                delay = random.uniform(2, 5)
            log_message(f"Waiting {delay:.2f}s before next email...")
            time.sleep(delay)

        try:
            server.quit()
        except Exception:
            pass

        log_message("All emails processed.")
    except Exception as e:
        log_message(f"Error in send_bulk_emails: {e}")

# ==============================
# Utility: parse receivers string
# ==============================
def parse_receivers(text: str):
    """
    Accepts text with emails separated by commas, newlines, semicolons.
    Returns cleaned, deduped list (preserving order).
    """
    if not text:
        return []
    # split on commas, semicolons, newlines
    parts = re.split(r"[,\n;]+", text)
    seen = set()
    out = []
    for p in parts:
        e = p.strip()
        if not e:
            continue
        # basic validation
        if EMAIL_RE.match(e):
            if e.lower() not in seen:
                seen.add(e.lower())
                out.append(e)
        else:
            # keep invalid ones too but log a warning (so user sees mistake)
            if e.lower() not in seen:
                seen.add(e.lower())
                out.append(e)
                log_message(f"Warning: '{e}' doesn't look like a valid email address.")
    return out

# ==============================
# PyWebView API
# ==============================
class API:
    def send_emails(self, new_settings):
        """
        Called from JS when user clicks Send.
        Updates settings, parses receivers, starts worker thread.
        Returns an empty string (avoid popup showing 'None').
        """
        # update allowed keys
        for key in ("smtp_server","smtp_port","sender_email","password","subject","body","min_delay","max_delay"):
            if key in new_settings:
                settings[key] = new_settings[key]

        # parse receivers
        receivers_text = new_settings.get("receiver_emails", "")
        emails = parse_receivers(receivers_text)
        settings["receiver_list"] = emails
        log_message(f"Loaded {len(emails)} receiver emails from textarea.")

        # Start worker thread as daemon so it won't block process exit
        t = threading.Thread(target=send_bulk_emails, daemon=True)
        t.start()
        log_message("Emails are being sent in background!")
        # Return an empty string to JS (avoid 'None' popups). JS won't alert.
        return ""

    # Optional: let JS ask for current settings (if you want prefill)
    def get_settings(self):
        return settings

# ==============================
# HTML UI
html_template = """<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>AirMailX Sender</title>
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body {
      padding: 8px;
      background: #f8f9fa;
      font-size: 13px;
    }
    .container {
      max-width: 780px;
    }
    label {
      font-size: 13px;
      margin-bottom: 2px;
    }
    input, textarea, select {
      font-size: 13px;
      padding: 4px 6px;
    }
    textarea {
      resize: vertical;
    }
    .nav-tabs .nav-link {
      padding: 4px 10px;
      font-size: 13px;
    }
    .btn {
      font-size: 13px;
      padding: 5px 10px;
    }
    #logBox {
      background: #0b0b0b;
      color: #c7f9c7;
      font-family: monospace;
      font-size: 12px;
      height: 160px;
      overflow-y: auto;
      padding: 6px;
      border-radius: 6px;
      margin-top: 10px;
      border: 1px solid #333;
    }
    h3, h5 {
      font-size: 16px;
      margin-top: 10px;
      margin-bottom: 8px;
    }
    p, li, ol {
      font-size: 13px;
      line-height: 1.4em;
    }
  </style>
</head>
<body>
<div class="container">
  <h4 class="mb-2 text-center">📧 AirMailX Sender</h4>

  <!-- Tabs -->
  <ul class="nav nav-tabs mb-2" id="mainTabs" role="tablist">
    <li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#home">Home</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#instructions">Instructions</button></li>
    <li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#about">About</button></li>
  </ul>

  <div class="tab-content">
    <div class="tab-pane fade show active" id="home">
      <form id="settingsForm">
        <h5>SMTP Settings</h5>
        <div class="row g-2 mb-2">
          <div class="col-md-7">
            <label>SMTP Server</label>
            <input class="form-control" name="smtp_server" value="smtp.gmail.com">
          </div>
          <div class="col-md-3">
            <label>Port</label>
            <input class="form-control" name="smtp_port" value="587" type="number">
          </div>
        </div>

        <h5>Sender Account</h5>
        <div class="mb-2">
          <label>Email</label>
          <input type="text" class="form-control" name="sender_email" placeholder="you@example.com">
        </div>
        <div class="mb-2">
          <label>App Password</label>
          <input type="password" class="form-control" name="password" placeholder="App Password">
        </div>

        <h5>Email Content</h5>
        <div class="mb-2">
          <label>Subject</label>
          <input class="form-control" name="subject" value="Test Email from AirMailX Sender">
        </div>
        <div class="mb-2">
          <label>Body (HTML supported)</label>
          <textarea class="form-control" rows="4" name="body">&lt;html&gt;Paste your HTML email template here&lt;/html&gt;</textarea>
        </div>

        <h5>Delay Settings</h5>
        <div class="row g-2 mb-2">
          <div class="col-md-6">
            <label>Min Delay (s)</label>
            <input class="form-control" name="min_delay" value="2" type="number" min="0">
          </div>
          <div class="col-md-6">
            <label>Max Delay (s)</label>
            <input class="form-control" name="max_delay" value="5" type="number" min="0">
          </div>
        </div>

        <h5>Receiver Emails</h5>
        <div class="mb-2">
          <label>Paste emails separated by commas / newlines / semicolons:</label>
          <textarea class="form-control" name="receiver_emails" rows="6" placeholder="email1@example.com, email2@example.com, ..."></textarea>
        </div>

        <div class="text-center mb-2">
          <button id="sendBtn" type="button" class="btn btn-success btn-sm" onclick="sendEmails()">🚀 Send Emails</button>
        </div>
      </form>

      <h5 class="mt-2">📜 Logs</h5>
      <div id="logBox"></div>
    </div>

    <!-- Instructions Tab -->
    <div class="tab-pane fade" id="instructions" role="tabpanel">
      <h5>How to use AirMailX Sender</h5>
      <ol>
        <li>Enter SMTP details:
          <ul>
            <li><b>Server:</b> e.g. smtp.gmail.com, smtp.office365.com</li>
            <li><b>Port:</b> 587 (TLS) or 465 (SSL)</li>
          </ul>
        </li>
        <li>Enter your sender email and app password.</li>
        <li>Fill subject & HTML body for the email.</li>
        <li>Paste receiver emails separated by commas or newlines.</li>
        <li>Set delay range to prevent spam flagging.</li>
        <li>Click "Send Emails" and watch logs below.</li>
      </ol>
      <p><b>Tip:</b> Make sure SMTP access is enabled and app password is used if required.</p>
    </div>

    <!-- About Tab -->
    <div class="tab-pane fade" id="about">
      <h5>About AirMailX Sender</h5>
      <p><b>Version:</b> 1.0</p>
      <p><b>Developer:</b> Code with ❤️ Adarsh</p>
      <p>This is a simple email sending app for sending bulk marketing emails safely and efficiently.</p>
      <p><b>Disclaimer:</b> Do not use this app for spam or malicious activity.</p>
      <p><b>Connect:</b> <a href="https://www.linkedin.com/in/adarshd5099" target="_blank">LinkedIn</a></p>
    </div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
async function sendEmails(){
  const form = document.getElementById('settingsForm');
  const data = {};
  for(const el of form.elements){
    if(el.name) data[el.name] = el.value;
  }

  const btn = document.getElementById('sendBtn');
  btn.disabled = true;
  btn.innerText = "⏳ Sending...";
  try {
    await window.pywebview.api.send_emails(data);
  } catch (e) {
    addLog("[JS] Error calling send_emails: " + String(e));
  } finally {
    setTimeout(() => {
      btn.disabled = false;
      btn.innerText = "🚀 Send Emails";
    }, 1200);
  }
}

function addLog(msg){
  const box = document.getElementById("logBox");
  box.innerHTML += msg + "<br>";
  box.scrollTop = box.scrollHeight;
}
</script>
</body>
</html>
"""

# ==============================
# Start PyWebView
# ==============================
if __name__ == "__main__":
    api = API()
    window = webview.create_window("AirMailX Sender", html=html_template, width=980, height=860, js_api=api)
    webview.start()
