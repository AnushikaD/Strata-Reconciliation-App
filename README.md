# Strata Reconciliation App

Local web app for analysing strata/property management reports and generating levy arrears and advance breakdowns.

## Run

```powershell
python -m pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`.

## Railway Hosting

Use this start command:

```bash
python -m waitress --host=0.0.0.0 --port=$PORT app:app
```

Railway will install dependencies from `requirements.txt`.

## Vercel Hosting

This app includes `vercel.json` and can be imported into Vercel from GitHub.

Vercel will detect the Flask `app` in `app.py` and install dependencies from `requirements.txt`.

Important Vercel notes:

- Uploaded files are processed in temporary serverless storage.
- Large PDF uploads may hit Vercel request/runtime limits.
- Export downloads use the last processed result in the current serverless runtime, so Railway or local office hosting is still better for heavier accounting reports.

## Inputs

- Balance Sheet Report: PDF, Excel, or CSV
- Levy Positions Report: PDF, Excel, or CSV
- Owner Transaction Summary Report: PDF, Excel, or CSV

The first version is tuned to the sample PDF reports supplied for CTS 32563 and includes fallback parsing for Excel/CSV table uploads.
