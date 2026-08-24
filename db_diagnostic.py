from __future__ import annotations
import os, socket, sys
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

try:
    import tomllib
except Exception:
    tomllib = None


def read_url():
    value = os.getenv("DATABASE_URL", "").strip()
    source = "Windows environment"
    if value:
        return value, source
    secret = Path('.streamlit') / 'secrets.toml'
    if secret.exists() and tomllib:
        try:
            data = tomllib.loads(secret.read_text(encoding='utf-8'))
            value = str(data.get('DATABASE_URL', '')).strip()
            if value:
                return value, str(secret)
        except Exception as exc:
            print(f"ERROR: could not parse {secret}: {exc}")
            return "", ""
    return "", ""

url, source = read_url()
if not url:
    print("DATABASE CHECK FAILED")
    print(" - DATABASE_URL was not found in the current CMD environment or .streamlit\\secrets.toml")
    sys.exit(2)

p = urlsplit(url)
host = p.hostname or ''
port = p.port or 5432
print("DATABASE CONFIG FOUND")
print(f" - Source: {source}")
print(f" - Host: {host}")
print(f" - Port: {port}")
print(f" - Database: {(p.path or '/postgres').lstrip('/')}")
print(" - Password: hidden")

try:
    addrs = sorted({x[4][0] for x in socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)})
    print("DNS: OK -> " + ", ".join(addrs))
except Exception as exc:
    print(f"DNS: FAILED -> {exc}")
    sys.exit(3)

try:
    sock = socket.create_connection((host, port), timeout=8)
    sock.close()
    print(f"TCP {port}: OK")
except Exception as exc:
    print(f"TCP {port}: FAILED -> {exc}")

try:
    import psycopg2
except Exception as exc:
    print(f"PSYCOPG2: NOT AVAILABLE -> {exc}")
    sys.exit(4)


def test(test_url, label):
    try:
        conn = psycopg2.connect(test_url, connect_timeout=20, sslmode='require', gssencmode='disable')
        cur = conn.cursor(); cur.execute('select 1'); cur.fetchone(); cur.close(); conn.close()
        print(f"{label}: CONNECTED SUCCESSFULLY")
        return True
    except Exception as exc:
        text = str(exc).replace(url, '<DATABASE_URL>')
        print(f"{label}: FAILED -> {type(exc).__name__}: {text}")
        return False

if test(url, f"POSTGRES {port}"):
    sys.exit(0)

if host.endswith('.pooler.supabase.com') and port == 5432:
    try:
        netloc = p.netloc.rsplit(':',1)[0] + ':6543'
        tx = urlunsplit((p.scheme, netloc, p.path, p.query, p.fragment))
        if test(tx, "TRANSACTION POOLER 6543"):
            print("NOTE: Session Pooler failed but Transaction Pooler is reachable.")
            sys.exit(0)
    except Exception as exc:
        print(f"TRANSACTION POOLER BUILD FAILED -> {exc}")

print("DATABASE CHECK FAILED")
print(" - Project files compile independently of this failure.")
print(" - If SQL Editor SELECT 1 works but the pooler returns NXDOMAIN, the fault is the Supabase pooler-to-database route, not the project folder.")
sys.exit(1)
