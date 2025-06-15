# Email-Reader
# 📧 Gmail Excel Reader

This Python script connects to your **Gmail inbox**, searches for emails with a **specific subject**, downloads any **attached `.xlsx` file**, and **processes the content**.

## 🚀 Features

- Gmail OAuth2 authentication
- Filters unread emails by subject
- Downloads and parses `.xlsx` attachments
- Prints the content in a readable format

## 📁 Folder Structure

Email-Reader/
├── read_email_excel.py
├── .env
├── token.json
├── credentials.json


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
        EMAIL_USER=youremail@gmail.com
7. Run the script
```bash
python read_email_excel.py
```







## NOTE
[NOTE: THIS IS ONLY A BASIC VERSION. Involves: Reading Gmail attachment (.xlsx ONLY) and displaying the contents]
