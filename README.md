# AI SmartGuard — local Streamlit edition

A ChatGPT-style spam/scam checker built with Streamlit. It needs **no API key** and makes **no calls to OpenAI or another AI service**. Detection is performed by transparent, local rules and patterns—not by a generative AI model—so it is inexpensive and works without a model service, but it can miss scams or flag unusual legitimate messages.

## Run it

Python 3.10 or newer is recommended.

```bash
cd ai_smartguard_local
python -m venv .venv
```

Activate the environment, then install the dependencies and start Streamlit:

```bash
# macOS / Linux
source .venv/bin/activate

# Windows PowerShell: use this instead of the command above
# .venv\Scripts\Activate.ps1

python -m pip install -r requirements.txt
streamlit run app.py
```

Streamlit will print a local address to open in your browser.

## What it checks

- Text pasted into the chat
- Images (`png`, `jpg`, `jpeg`, `webp`) using optional local OCR
- PDF, DOCX, TXT, CSV, Markdown, RTF, XLSX, and PPTX text
- Common scam signals such as requests to disclose OTPs/PINs/passwords, UPI-PIN tricks, account-block threats paired with action links, advance-fee job/loan offers, fake prizes/refunds, remote-access apps, URL shorteners, raw-IP links, and punycode domains
- A local SQLite history, usage counts, and a history page

URLs are inspected as text only: **the app never opens or visits them**. Uploaded files are read in memory for extraction; file contents are not saved as files by the app. The history database stores the text you paste (and a filename for file-only submissions), plus the result and timestamp. It does not store the uploaded file or OCR/document text. Delete `smartguard_history.db` to remove the local history.

## Optional OCR for images

The `pytesseract` Python package is installed with `requirements.txt`, but it also requires the Tesseract OCR application:

- **Windows:** install Tesseract OCR, then ensure `tesseract.exe` is on `PATH`.
- **macOS:** `brew install tesseract`
- **Debian/Ubuntu:** `sudo apt install tesseract-ocr`

Without the Tesseract application, an image is marked **UNCERTAIN** if no text can be extracted; it is not silently treated as safe. Scanned PDFs are not OCR'd in this version; PDFs need selectable text.

## Privacy and deployment

Analysis has no external API dependency or outbound model call. “Local” means processing happens in the machine/process running Streamlit. If you deploy the app to a hosted server, uploads and pasted text reach that server, so use a trusted deployment and protect its history database. The app does not embed or ask for any API key.

## Important limitation

The displayed percentage is a rough rule-based estimate, not a statistically calibrated probability. A **SAFE** result means only that these rules found no strong scam indicators; it does not prove a message is genuine. Verify unexpected requests using an official app, website, or phone number you already trust. Never share OTPs, passwords, PINs, or card security codes.
