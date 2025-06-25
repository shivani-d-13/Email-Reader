# 📧 Gmail Excel Reader

A Python script that connects to your Gmail account, searches for **unread emails** with a specific subject, downloads any `.xlsx` attachment, reads and prints its contents, and then marks the email as **read**.

---

## 🚀 Features

- ✅ Authenticates with Gmail using OAuth 2.0
- 📥 Searches for **unread** emails by subject
- 📎 Downloads `.xlsx` attachments
- 📊 Reads and prints Excel data using `openpyxl`
- ✉️ Marks the processed emails as **read**

---

## 🧰 Requirements

- Python 3.7+
- Gmail account with API access enabled
- A valid `credentials.json` (OAuth 2.0 client ID)

### 📦 Python Libraries

Install required dependencies:

```bash
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dotenv openpyxl
```

