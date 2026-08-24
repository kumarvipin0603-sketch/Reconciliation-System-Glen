# v15.18a — Project Review & Database Resilience Fix

## Reviewed
- app.py
- amazon_engine.py
- flipkart_engine.py
- requirements.txt
- preflight_check.py
- local start BAT
- one-click deployment BAT
- deployment/README guidance
- Excel workflow templates
- packaged ZIP structure

## Findings / fixes
1. The uploaded loose `app(9).py` was newer than the `app.py` inside the uploaded Management MIS ZIP. The reviewed package now uses the newer build under the correct production filename `app.py`.
2. A PostgreSQL/Supabase failure occurred during the unguarded module-level `init_db()` call, which caused Streamlit to show a raw traceback before the dashboard could load. Startup is now fail-closed with a clean retry screen.
3. PostgreSQL connections now enforce SSL, disable GSS encryption negotiation, request UTF-8 client encoding, retry short-lived failures, and attempt Supabase transaction-pooler port 6543 only after the configured Session Pooler fails.
4. The app never silently switches from a configured cloud database to SQLite. This protects production history from being split across databases.
5. Opening/restarting the app no longer runs the task auto-completion maintenance function merely as a side effect of startup.
6. Added `db_diagnostic.py`, which checks local secret/env discovery, DNS, TCP reachability and PostgreSQL login without printing the password.
7. Local start BAT now runs preflight and database diagnostics from its own folder, so moving the project to another drive/folder does not break relative paths.
8. `.streamlit/secrets.toml.example` documents the correct local secret location. The real `secrets.toml` remains ignored by Git.

## Important external finding
If Supabase SQL Editor can execute `select 1` but the Session Pooler returns `{:error, :nxdomain}`, Supabase documents this as the pooler being unable to connect to the customer database. Code cannot repair that backend route; the revised app handles it safely and can retry/fallback to the transaction-pooler route where compatible.
