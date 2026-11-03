# One origin for FlowLab Pro and E&I Lab

Both frameworks run inside `integrated_app:application`, one WSGI process and one
port. Do not deploy `web_app:app` alone or run a second Django server for E&I.
Independent module sessions, existing accounts, databases and uploaded files are
unchanged by this hosting fix. There are no database migrations for this change.

## Ngrok / Windows executable

1. Stop the old executable after saving open work. Keep its existing runtime data.
2. Put `configuration.env` next to the executable. The prepared
   `deployment/ngrok.env.example` contains the review tunnel address; copy it as
   `configuration.env`. Do not overwrite another deployment's database/secret
   settings: merge the hosting entries if a file already exists.
3. Start the updated executable. Its fixed port is 5050 when public hosting is
   configured; it will report an occupied port instead of silently changing it
   and leaving ngrok connected to a different application.
4. Start ngrok against `http://127.0.0.1:5050`. Preserve the public Host header;
   don't rewrite it to localhost. Both `/login` and `/ei-lab/login/` belong to
   the same public URL. `/ei-lab` also redirects to the prefixed login.
5. Use the **HTTPS public URL** to sign in when secure cookies are enabled.
   Plain local HTTP is not a supported sign-in address in that public profile;
   use a separate local profile with `DJANGO_SECURE_COOKIES=0` when not sharing.

If the ngrok hostname changes, update `EATL_PUBLIC_URL`, `DJANGO_ALLOWED_HOSTS`
and `DJANGO_CSRF_TRUSTED_ORIGINS`, then restart. No rebuild is needed just for a
hostname change. Existing process/service environment values take precedence
over `configuration.env`; remove stale overrides from the service if necessary.

## Source server / reverse proxy

The source runner reads the same project-root `configuration.env`. Start with
`Start-EATL.ps1` or `.venv/Scripts/python.exe scripts/run_integrated_server.py`.
`EATL_BIND_HOST` defaults to `127.0.0.1`; `EATL_PORT` defaults to 5050.
Before exposing a source deployment, set distinct, strong persistent
`FLOWLAB_SECRET_KEY` and `DJANGO_SECRET_KEY` values in private service
configuration. Do not use development defaults or commit those values to Git.
For direct LAN use choose a bind address explicitly and list the actual IP/host
in `DJANGO_ALLOWED_HOSTS`. Do not expose the database or development server.

For nginx proxying from the same machine, use a location for the **entire app**:

```nginx
location / {
    proxy_pass http://127.0.0.1:5050;
    proxy_set_header Host $http_host;
    proxy_set_header X-Forwarded-Proto $scheme;
}
```

Waitress trusts only `X-Forwarded-Proto` from the exact `EATL_TRUSTED_PROXY` IP
(127.0.0.1 by default). Other forwarding headers are discarded. If the proxy is
on another machine, set its actual IP and firewall the application port to that
proxy. Never configure a wildcard trusted proxy or wildcard public hostname.
Proxy software must replace, not blindly pass through, client scheme headers.
If another WSGI server is used, its trusted-proxy policy must be configured too.

## Diagnostics

- **400 / DisallowedHost**: the public Host is not in the allow-list, or another
  service/environment is still supplying an old allow-list.
- **403 on login / CSRF**: origin, public HTTPS scheme or cookie settings differ.
  Do not disable CSRF protection to work around this.
- **404 only for E&I**: the wrong Flask-only WSGI entry point is running, or the
  proxy is stripping `/ei-lab` instead of forwarding the complete URL path.
- **Login repeatedly returns to login**: clear old site cookies and verify the
  public URL uses HTTPS when secure cookies are enabled.

The executable logs to `%LOCALAPPDATA%/EATL-Integrated-Lab/runtime.log` and
`runtime-console.log`. Keep logs private: they may contain record identifiers.
Automated tests use disposable databases, simulated public hosts, real CSRF
tokens and independent authentication. A live public ngrok endpoint must still
be checked on the actual deployed machine after applying its configuration.
