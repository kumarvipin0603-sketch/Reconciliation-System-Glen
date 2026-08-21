# Glen Reconciliation Tower — Production Setup

## One-time setup

1. Keep this project in the GitHub repository already connected to the Streamlit Cloud app.
2. In Streamlit Cloud open **App → Settings → Secrets**.
3. Add the production PostgreSQL/Supabase connection string:

```toml
DATABASE_URL = "postgresql://<user>:<password>@<host>:<port>/<database>"
```

4. Save Secrets and reboot the app once.
5. Confirm the app shows **Storage: Supabase PostgreSQL — persistent cloud database**.

Do not place the real database password in `app.py`, GitHub, `.env`, README, or the BAT file.

## Normal future deployment

Double-click:

`DEPLOY_WEBSITE_ONE_CLICK.bat`

The script validates the project, commits local code changes, syncs GitHub `main`, validates again, pushes, and opens the live app. Streamlit Cloud then auto-redeploys from GitHub.

A code deployment does not require Amazon/Flipkart source data to be uploaded again. The latest verified source is stored in the database and remains active until an Admin uploads a newer source workbook.

## Team-working persistence

Team values are kept separately from source-derived reconciliation data. Source refreshes do not intentionally rewrite existing Team Remarks, Ticket information, MIR/TEI values, Working/Completed status or completion dates. An unchanged team workbook upload is treated as a no-op.

## First production source load

If the Supabase database is new/empty, upload the current Amazon and/or Flipkart source once. After that, use a new source upload only when the marketplace source itself changes.

## Troubleshooting

If the app shows **Local SQLite fallback** on Streamlit Cloud, the production database secret is missing/not readable. Fix `DATABASE_URL` before relying on the app for persistent production history.

If the one-click BAT stops on pre-flight, do not bypass it. Run `py preflight_check.py` from the project folder and resolve the reported error.
