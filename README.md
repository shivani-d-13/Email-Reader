# 📧 Gmail Excel Reader

A Python script that connects to your Gmail account, searches for **unread emails** with a specific subject, downloads any `.xlsx` attachment, reads and prints its contents, and then marks the email as **read**.

---

## 🚀 Features

- ✅ Authenticates with Gmail using OAuth 2.0
- 📥 Searches for **unread** emails by subject
- 📎 Downloads `.xlsx` attachments
- 📊 Reads and prints Excel data using `openpyxl`
- ✉️ Marks the processed emails as **read**

## 🧰 Requirements

- Python 3.7+
- Gmail account with API access enabled
- A valid `credentials.json` (OAuth 2.0 client ID)

### 📦 Python Libraries

Install required dependencies:

```bash
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dotenv openpyxl
```

## 🔧 Setup
### 1. Enable the Gmail API
Go to Google Cloud Console → APIs & Services
- Create a project
- Enable Gmail API
- Create OAuth 2.0 credentials
- Download credentials.json and place it in the project directory

### 2. Set up your environment
Create a .env file in the root directory:
```ini
EMAIL_ADDRESS=your-email@gmail.com
```

### 3. Run the script
```bash
python read_email_excel.py
```
The first time, it will prompt for Google authentication in your browser. This will save a token.json file for future runs.

## 📂 Project Structure
```bash
.
├── script_name.py
├── credentials.json
├── token.json           # created after first auth
├── .env
└── attachment.xlsx      # downloaded Excel file
```

## 🧠 Notes
- The script currently processes only the latest unread email with the target subject.
- Make sure SCOPES is set to ['https://www.googleapis.com/auth/gmail.modify'] so that it can update email read status.
- If you change SCOPES, you must delete token.json and re-authenticate.
