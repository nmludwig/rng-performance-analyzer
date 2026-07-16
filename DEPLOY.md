# Deploying rng-performance-analyzer

The app is a Flask + gunicorn service. It runs on the CElab server
(`matthewludwig.celab.ringcentral.com`, `216.249.107.252`) behind nginx, and is
reached at **https://rng-performance-analyzer.celab.ringcentral.com**.

- **No system packages needed** — the deck is built with `python-pptx`; there is
  no LibreOffice/soffice dependency at runtime.
- **Python 3.12** is required. Newer (3.14) has no stable pandas/numpy wheels and
  crashes the worker with a native SIGSEGV.

## Required secrets (`.env` in the app directory)

The app calls `load_dotenv(override=True)`, so it reads `/opt/rng-performance-analyzer/.env`.
Keep it `chmod 600`; it is git-ignored and must never be committed.

```
SECRET_KEY=<random; python -c "import secrets; print(secrets.token_hex(32))">
ANTHROPIC_API_KEY=sk-ant-...
FIRECRAWL_API_KEY=fc-...        # optional; website crawl degrades gracefully if absent

# RingCentral corporate SSO (login). Register the redirect URI below in the
# RingCentral OAuth app, then reuse that app's client id/secret here.
RC_CLIENT_ID=...
RC_CLIENT_SECRET=...
RC_REDIRECT_URI=https://rng-performance-analyzer.celab.ringcentral.com/oauth/callback
# RC_SERVER_URL=https://platform.ringcentral.com   # optional; defaults to production
```

`SECRET_KEY` and `ANTHROPIC_API_KEY` are required — the app fails to start
without them. Login is RingCentral OAuth: without `RC_CLIENT_ID`,
`RC_CLIENT_SECRET` and `RC_REDIRECT_URI` the app still runs but the login page
shows "single sign-on is not configured" and no one can sign in. Only
`@ringcentral.com` accounts are admitted. `FIRECRAWL_API_KEY` is optional.

The `RC_REDIRECT_URI` above must be registered as an allowed redirect URI in the
RingCentral OAuth app (the same app ACE Engine uses, or a new one). The
authorization-code flow is: `/oauth/start` → RingCentral login →
`/oauth/callback` reads the extension's contact email → employee check → session.

## First-time setup on the server

```bash
# 1. Clone
cd /opt
sudo git clone https://github.com/nmludwig/rng-performance-analyzer.git
sudo chown -R "$USER:" rng-performance-analyzer
cd rng-performance-analyzer

# 2. Python 3.12 venv + deps
python3.12 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

# 3. Create .env (see "Required secrets" above), then chmod 600 .env

# 4. Smoke test
.venv/bin/gunicorn app:app --bind 127.0.0.1:8004 &
sleep 3 && curl -I http://127.0.0.1:8004/   # expect 302 -> /login
kill %1
```

## systemd service

`/etc/systemd/system/rng-analyzer.service`:

```ini
[Unit]
Description=RNG Performance Analyzer (Flask/gunicorn)
After=network.target

[Service]
User=matthew.ludwig
Group=domain users
WorkingDirectory=/opt/rng-performance-analyzer
Environment="PATH=/opt/rng-performance-analyzer/.venv/bin"
ExecStart=/opt/rng-performance-analyzer/.venv/bin/gunicorn --bind 127.0.0.1:8004 app:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now rng-analyzer
sudo systemctl status rng-analyzer --no-pager
```

`gunicorn.conf.py` (auto-loaded from `WorkingDirectory`) sets a single worker and
a 300s timeout. The single worker is deliberate: parsing a full-month ~40 MB Calls
export plus building the deck peaks around 400 MB, so a second concurrent job would
risk OOM. Do not add `--workers 2` unless the box has ample memory.

## nginx site

`/etc/nginx/sites-available/rng-performance-analyzer` (the `*.celab` wildcard cert
already covers this host):

```nginx
server {
    listen 443 ssl;
    server_name rng-performance-analyzer.celab.ringcentral.com;

    ssl_certificate     /etc/ssl/certs/STAR_celab_ringcentral_com.pem;
    ssl_certificate_key /etc/ssl/private/star-celab-ringcentral-com.key;

    client_max_body_size 60M;        # PR Calls .xlsx can be ~40 MB

    location / {
        proxy_pass http://127.0.0.1:8004;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 600s;     # deck-generation SSE stream runs long
        proxy_send_timeout 600s;
    }
}

server {
    listen 80;
    server_name rng-performance-analyzer.celab.ringcentral.com;
    return 301 https://$host$request_uri;
}
```

```bash
sudo ln -s /etc/nginx/sites-available/rng-performance-analyzer /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

## DNS

Each celab app is a CNAME to the server's own hostname (managed by the CElab DNS
owner — Tim McKee):

```
rng-performance-analyzer.celab.ringcentral.com  CNAME  matthewludwig.celab.ringcentral.com
```

(An A record to `216.249.107.252` works too.) The server can't resolve these
internal names itself; test resolution from a machine on the corporate VPN.

## Updating (replaces Render auto-deploy)

```bash
cd /opt/rng-performance-analyzer
git pull
.venv/bin/pip install -r requirements.txt   # only when requirements changed
sudo systemctl restart rng-analyzer
```

## Verify

```bash
# Bypass DNS, prove nginx + app + TLS locally:
curl -I --resolve rng-performance-analyzer.celab.ringcentral.com:443:127.0.0.1 \
  https://rng-performance-analyzer.celab.ringcentral.com/

# Once DNS resolves (from VPN):
curl -I https://rng-performance-analyzer.celab.ringcentral.com/   # expect 302 -> /login
```

## Logs / troubleshooting

```bash
sudo systemctl status rng-analyzer --no-pager
journalctl -u rng-analyzer -n 100 --no-pager      # app / gunicorn logs
sudo tail -f /var/log/nginx/error.log              # proxy / TLS errors
```

- `KeyError: 'SECRET_KEY'` (or `APP_PASSWORD` / `ANTHROPIC_API_KEY`) on start →
  `.env` is missing a value or the service isn't running from
  `WorkingDirectory=/opt/rng-performance-analyzer`.
- HTTP 413 on upload → `client_max_body_size` too low in nginx.
- "connection lost" mid-generation → nginx `proxy_read_timeout` too low.
