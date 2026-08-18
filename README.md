# Real-Company Ratio Comparison Tool (with student login)

This is your existing ratio-analysis Streamlit app (`AI Gemini.py`) with a
login screen added in front of it: students must sign in with a
username/password before they can use the tool, and each generated
CSV/Excel/PDF report is stamped "Prepared by: <student name>".

## Files

- `app.py` — the app itself (login gate + your original ratio comparison tool, unchanged below the login).
- `build_roster.py` — turns a plain-text roster CSV into `config.yaml` (hashed passwords).
- `roster.csv` — **sample** roster with 2 demo logins. Replace with your real class list.
- `config.yaml` — generated login file (hashed passwords). Already built from the sample roster so you can try it immediately.
- `requirements.txt` — everything needed to run the app.

## Try it locally right now

```bash
pip install -r requirements.txt
streamlit run app.py
```

Log in with one of the sample accounts from `roster.csv`:

| username | password |
|---|---|
| demo_student | ChangeMe123! |
| jdoe | ChangeMe123! |

## Setting up your real class roster

1. Open `roster.csv` and replace the sample rows with your actual students —
   one row per student: `username,name,email,password`.
   - `username` is what they type to log in (e.g. a student ID or their
     first initial + last name). Must be unique per student.
   - `password` is a starting password you assign. Students can't change it
     from inside the app — to reset someone's password, edit their row and
     rerun the build script.
2. Regenerate the login file:
   ```bash
   python build_roster.py
   ```
   This overwrites `config.yaml` with freshly hashed passwords for everyone
   in `roster.csv`.
3. **Never share or commit `roster.csv` or `config.yaml` publicly** —
   `roster.csv` has plain-text passwords, and `config.yaml` has the hashed
   versions plus a cookie secret. Keep both private (e.g. add them to
   `.gitignore` if this becomes a git repo).

## Deploying to Streamlit Community Cloud

Don't upload `config.yaml` to a public repo. Instead:

1. Push `app.py`, `build_roster.py`, and `requirements.txt` to a repo
   (leave `roster.csv`/`config.yaml` out, or keep the repo private).
2. Deploy the app on share.streamlit.io.
3. Run `python build_roster.py` locally, and copy the TOML block it prints
   at the end into the app's **Settings → Secrets** box on Streamlit Cloud.
   That gives the deployed app the same login roster without it ever
   touching your public repo.
4. Redeploy/reboot the app after saving secrets.

## AI interpretation: bring-your-own-chat, no API key

This app does **not** call any AI API and never asks for a key — nothing
about a student's AI account touches this app or your server. Consumer
chat products (ChatGPT, Gemini, Claude.ai, Perplexity) don't allow being
embedded in another site, and there's no supported way for a third-party
app to drive a student's already-logged-in chat session for them, so
instead the Summary Dashboard tab gives students a three-step manual
bridge:

1. **Open your own AI chat** — buttons link out to chatgpt.com,
   gemini.google.com, claude.ai, and perplexity.ai, each opening in a new
   browser tab using whatever account the student is already logged into.
2. **Copy the prompt** — the app builds a prompt from the already-computed
   ratio table (never raw financials) and shows it in a copyable code box.
3. **Paste the answer back** — the student pastes the AI's response into a
   text box in the app, which renders it alongside the ratios, includes the
   original prompt for disclosure, and requires the student to write a
   short verification note before treating it as done.

This keeps everything on the student's own AI account and quota — no
shared instructor key, no API cost tracking, no dependency on any one AI
provider.

## What changed from the original app

- Added a login screen (`streamlit-authenticator`) gating the whole tool —
  nothing past the title screen runs until a valid username/password is
  entered.
- Sidebar now shows "Logged in as `<name>`" with a **Log out** button.
- CSV/Excel/PDF exports now include a "Prepared by: `<student name>`" line.
- Replaced the Gemini-API-key AI feature with the bring-your-own-chat
  workflow described above — no `google-genai` dependency, no API key
  fields anywhere in the app.
- Everything else (ratio calculations, yfinance fetch logic, benchmarks,
  tabs, downloads) is unchanged from your original file.
