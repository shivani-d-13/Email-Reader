# Gmail Excel Reader

A secure automation tool to reduce manual email tasks.
This Python script connects to your **Gmail inbox**, searches for emails with a **specific subject**, downloads any **attached `.xlsx` file**, and **processes the content**.

## 🚀 Features

- Gmail OAuth2 authentication
- Filters unread emails by subject
- Downloads and parses `.xlsx` attachments
- Prints the content in a readable format

## 🛠️ Setup Instructions

1. Clone this repo
```bash
git clone https://github.com/shivani-d-13/Email-Reader.git
cd Email-Reader
```
2. Setup virtual environment
```bash
python -m venv venv
source venv/bin/activate    # On Windows: venv\Scripts\activate
```
3. Install Required Dependencies
```bash
pip install -r requirements.txt
```
4. Follow [this guide](https://developers.google.com/gmail/api/quickstart/python) to get `credentials.json`
5. Create .env File

    In the root of your project, create a .env file with the following content:
    ```ini
   EMAIL_USER=your_email@gmail.com
    ```
7. Run the script
```bash
python read_email_excel.py
```

On the first run:
1. A browser window will open to let you log into your Gmail account
2. Grant access to read your emails
3. A token.json file will be generated for future use

## ✅ What the Script Does
- Authenticates using OAuth2
- Connects to your Gmail inbox
- Finds the most recent email that matches the specified subject
- Downloads all .xlsx attachments from that email to your selected folder

## ❗ Security Notes
- This app uses OAuth2 in test mode (personal use only)
- .env, token.json, credentials.json, and downloaded files should not be committed
- .gitignore is pre-configured to exclude these files

## 💡 Future Updates
- Parse and extract data from Excel files
- Support .pdf and .csv attachments
- Add filters by sender or date
- Schedule the script using cron or Task Scheduler
- Export data to Google Sheets or a database



## NOTE
[NOTE: THIS IS ONLY A BASIC VERSION. Involves: Reading Gmail attachment (.xlsx ONLY) and displaying the contents]
