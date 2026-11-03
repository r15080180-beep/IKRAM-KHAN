# AI-Powered CV Analyzer

An AI-powered web app that reviews your CV and gives you practical advice to improve it. Upload a PDF or DOCX, optionally enter the job you want, and get a dashboard with scores, strengths, weaknesses, missing skills and a personalised improvement roadmap. It is built with **Python**, **Streamlit** and the **Google Gemini API**.

> **Important:** ATS (Applicant Tracking System) scores are AI-generated *estimates*. They are not guarantees of recruitment success, because every employer's system works differently. AI can also make mistakes, so always review its advice.

## Main features

- Upload a CV in **PDF** or **DOCX** format (up to 5 MB)
- Overall CV score out of 100
- ATS compatibility score (estimate) with issues and tips
- Strengths, weaknesses and missing CV sections
- Analysis of education, technical skills, work experience and projects
- Grammar and professional language check
- Prioritised improvement suggestions
- Target job title comparison with a list of missing relevant skills
- Personalised improvement roadmap
- Clean dashboard with tabs, metrics, progress bars and expanders
- Download the report as **TXT** or **PDF**

## Privacy and security

- Uploaded files are processed **in memory only** and are never saved to disk.
- Only `.pdf` and `.docx` files are accepted. The extension, MIME type and file signature are checked, and the file size is limited to 5 MB.
- Uploaded content is only read as text and is never executed.
- E-mail addresses and phone numbers are removed before the CV text is sent to the AI, and they never appear in the dashboard or the reports.
- The AI is instructed not to invent skills, qualifications or experience.
- API keys are read from Streamlit secrets or environment variables and are never hardcoded.

Note: the remaining CV text (without e-mail and phone) is sent to Google's Gemini API for analysis. Check Google's terms and privacy policy for the Gemini API, and avoid uploading CVs you are not allowed to share.

## Technologies used

| Technology | Purpose |
|---|---|
| Python 3.10+ | Programming language |
| Streamlit | Web interface |
| google-genai | Official Google Gemini Python SDK |
| PyMuPDF (`fitz`) | PDF text extraction |
| python-docx | DOCX text extraction |
| Pandas | Tables and chart data |
| python-dotenv | Loading a local `.env` file |
| fpdf2 | Creating the PDF report |

## Project structure

```
cv-analyzer/
├── app.py             # The complete Streamlit application
├── requirements.txt   # Python dependencies
├── README.md          # This file
├── .gitignore         # (you create this) keeps your secrets out of GitHub
└── .env               # (you create this, optional) local API key - never upload it
```

## Installation

1. Install **Python 3.10 or newer** from [python.org](https://www.python.org/downloads/).
2. Put `app.py`, `requirements.txt` and `README.md` in a folder, for example `cv-analyzer`.
3. Open a terminal in that folder and create a virtual environment:

   ```bash
   python -m venv venv
   ```

4. Activate it:

   - Windows: `venv\Scripts\activate`
   - macOS / Linux: `source venv/bin/activate`

5. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## How to create a Google AI Studio API key

1. Go to [Google AI Studio](https://aistudio.google.com/) and sign in with your Google account.
2. Click **Get API key** (in the left menu).
3. Click **Create API key** and choose a project (or create a new one).
4. Copy the key and keep it private. Anyone with this key can use your quota.

The free tier is usually enough for testing, but it has rate limits.

## How to configure the API key securely

Never write your key inside `app.py` and never upload it to GitHub. Use **one** of these options.

**Option A: `.env` file (simple)**

Create a file named `.env` in the project folder:

```
GEMINI_API_KEY=your_api_key_here
```

**Option B: Streamlit secrets file**

Create the folder `.streamlit` and inside it a file named `secrets.toml`:

```toml
GEMINI_API_KEY = "your_api_key_here"
```

**Optional: choose a model.** Google retires old Gemini models over time. The app tries `gemini-3.8-flash` first and automatically falls back to other models. To force a specific model, add this line to your `.env` or secrets file (see the model list in [Google AI Studio](https://aistudio.google.com/) or the [Gemini API docs](https://ai.google.dev/gemini-api/docs/models)):

```
GEMINI_MODEL=gemini-3.8-flash
```

(In `secrets.toml` write `GEMINI_MODEL = "gemini-3.8-flash"`.)

**Create a `.gitignore` file** with this content so your secrets are never uploaded:

```
.env
.streamlit/secrets.toml
venv/
__pycache__/
```

## How to run the application locally

With the virtual environment active, run:

```bash
streamlit run app.py
```

Your browser opens at `http://localhost:8501`. Upload a CV, optionally type a target job title, and click **Analyze CV**.

## How to upload the project to GitHub

1. Create a free account at [github.com](https://github.com) and install [Git](https://git-scm.com/downloads).
2. Make sure `.gitignore` exists (see above) so `.env` and `secrets.toml` are not uploaded.
3. On GitHub, click **New repository**, name it (for example `cv-analyzer`), and click **Create repository**. Leave it empty.
4. In your project folder run:

   ```bash
   git init
   git add app.py requirements.txt README.md .gitignore
   git commit -m "Initial commit: AI CV Analyzer"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/cv-analyzer.git
   git push -u origin main
   ```

5. Refresh the GitHub page and check that `.env` is **not** listed.

If you ever upload your API key by mistake, delete that key in Google AI Studio and create a new one.

## How to deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **Create app** and choose to deploy from an existing GitHub repository.
3. Select your repository, the branch `main`, and the main file path `app.py`.
4. Open **Advanced settings** and choose a Python version (3.11 or 3.12 is recommended).
5. Add your secrets (next section) in **Advanced settings -> Secrets**, then click **Deploy**.
6. Wait a few minutes. Your app gets a public URL that you can share.

## How to configure Streamlit secrets during deployment

In the Streamlit Community Cloud dashboard:

1. Open your app and click **Settings** (or use **Advanced settings** while deploying).
2. Open the **Secrets** tab.
3. Paste this (with your real key) in TOML format:

   ```toml
   GEMINI_API_KEY = "your_api_key_here"
   ```

4. Click **Save**. The app restarts automatically.

## Common errors and solutions

| Problem | Solution |
|---|---|
| "No Gemini API key was found" | Create `.env` or `.streamlit/secrets.toml` (see above), or add the secret in Streamlit Cloud, then restart the app. |
| "The Gemini API key was rejected" | The key is wrong, deleted or restricted. Create a new key in Google AI Studio and copy it without extra spaces or quotes. |
| "Rate limit or quota was reached" | You sent too many requests. Wait a minute and try again, or enable billing in Google AI Studio. |
| "The selected Gemini model is not available" | The model was renamed or retired. Set `GEMINI_MODEL` to a current model from the Gemini docs. |
| "Gemini service is temporarily unavailable" | The service is overloaded. Try again in a few minutes. |
| "Could not reach the Gemini API" | Check your internet connection, VPN or firewall. |
| "Looks like a scanned or image-based PDF" | The PDF contains pictures of text instead of real text. Export your CV from Word or Google Docs as a text PDF, upload a DOCX, or run an OCR tool first. |
| "PDF is password-protected" | Remove the password and upload again. |
| "File is corrupted" or "not a valid PDF/DOCX" | Re-save or re-export the file. Renaming a file's extension does not convert it. |
| "The file is too large" | Reduce the file below 5 MB (for example by compressing images). |
| `ModuleNotFoundError` | Activate your virtual environment and run `pip install -r requirements.txt`. |
| `streamlit: command not found` | Activate the virtual environment, or run `python -m streamlit run app.py`. |
| Deployment fails with a dependency error | Check that `requirements.txt` is in the repository root and select Python 3.11 or 3.12 in Advanced settings. |
| The AI returns incomplete results | The app retries automatically. If it keeps happening, click **Analyze CV** again. |

## Limitations

- Scanned PDFs (images) are not supported because OCR is not included.
- Only the first 15 pages of a PDF and the first 30,000 characters of text are analyzed.
- Scores are estimates generated by AI and can vary between runs.

## License

Use and modify this project freely for learning and personal projects.# IKRAM-KHAN
