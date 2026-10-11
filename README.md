# GST Return Downloader (GSTR-1 / GSTR-3B)

A Python + Selenium tool that bulk-downloads GSTR-1 and GSTR-3B PDFs from the GST portal for multiple logins, over any month range.

## Features
- Reads multiple GST logins from an Excel file (`Username | Password | Company Name`)
- Auto-fills credentials; you only solve the CAPTCHA/OTP in the browser
- Choose one inclusive month range (e.g. April 2022 to July 2024) and GSTR-1, GSTR-3B or both
- Auto-calculates the correct Financial Year and Quarter for each month
- Saves PDFs into per-company folders; already-downloaded returns are skipped, so re-runs only fetch what is missing

## Setup
1. Python 3.9+ and Google Chrome
2. `pip install -r requirements.txt`
3. Save a copy of `GST_Credentials_Template.xlsx` under a new name (e.g. `my_logins.xlsx`), fill in your own logins and keep it next to the script. The demo rows in the template are placeholders and will not log in. Never type real passwords into the template itself.
4. `python gst_return_downloader.py`

## Output
```
<Company Name>/gstr 1/GSTR1_of_April_2025.pdf
<Company Name>/gstr 3b/GSTR3B_of_April_2025.pdf
```

## Security & limitations
- The Excel file holds plain-text passwords. Keep it out of Git and out of cloud-synced folders (OneDrive, Dropbox, Google Drive). Delete it after use.
- Debug screenshots/HTML (`_debug/`) may contain GSTIN and session data. Delete them after troubleshooting.
- CAPTCHA is deliberately left manual.
- The portal's HTML changes from time to time; a selector may need updating if a step fails.
- Use only for logins you are authorised to access. This is an independent tool, not affiliated with GSTN.

## Author
Shubham - CA Finalist, articleship at a CA firm. Interested in finance and automation.
