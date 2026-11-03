"""Shared, explicit hosting configuration for both WSGI applications."""
from __future__ import annotations

import ipaddress
import os
from pathlib import Path
from urllib.parse import urlsplit


def load_configuration(root: Path) -> None:
    config = root / 'configuration.env'
    if not config.is_file():
        return
    for raw_line in config.read_text(encoding='utf-8-sig').splitlines():
        line = raw_line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, value = line.split('=', 1)
        if key.strip():
            # Process/service configuration always takes precedence over a file.
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def public_origin(value: str) -> str:
    parsed = urlsplit(value.strip())
    if (parsed.scheme not in ('http', 'https') or not parsed.hostname
            or parsed.username is not None or parsed.password is not None
            or parsed.path not in ('', '/') or parsed.query or parsed.fragment):
        raise ValueError('EATL_PUBLIC_URL must be an http(s) origin, without credentials, paths or query strings.')
    host = parsed.hostname.encode('idna').decode('ascii').lower()
    if '*' in host or any(character.isspace() for character in host):
        raise ValueError('EATL_PUBLIC_URL must name one specific server, not a wildcard.')
    port = parsed.port  # Also rejects malformed/out-of-range port numbers.
    authority = f'[{host}]' if ':' in host else host
    if port is not None and (parsed.scheme, port) not in (('https', 443), ('http', 80)):
        authority += f':{port}'
    return f'{parsed.scheme}://{authority}'


def _append_setting(name: str, additions: list[str], defaults: str = '') -> None:
    values = [item.strip() for item in os.environ.get(name, defaults).split(',') if item.strip()]
    os.environ[name] = ','.join(dict.fromkeys(values + additions))


def configure_hosting(root: Path) -> str:
    """Configure Django before import; never infer trust from request headers."""
    load_configuration(root)
    public_url = os.environ.get('EATL_PUBLIC_URL', '').strip()
    origin = public_origin(public_url) if public_url else ''
    additions = []
    if origin:
        hostname = urlsplit(origin).hostname
        additions.append(f'[{hostname}]' if ':' in hostname else hostname)
        _append_setting('DJANGO_CSRF_TRUSTED_ORIGINS', [origin])
        os.environ['EATL_PUBLIC_URL'] = origin
    _append_setting('DJANGO_ALLOWED_HOSTS', additions, '127.0.0.1,localhost,[::1]')
    os.environ.setdefault('DJANGO_DEBUG', '0')
    # Local desktop use remains HTTP-compatible. Public HTTPS defaults to secure
    # session/CSRF cookies; existing explicit settings are preserved.
    os.environ.setdefault('DJANGO_SECURE_COOKIES', '1' if origin.startswith('https://') else '0')
    return origin


def waitress_proxy_options() -> dict:
    """Trust scheme headers only from the explicitly configured reverse proxy."""
    peer = os.environ.get('EATL_TRUSTED_PROXY', '127.0.0.1').strip()
    options = {'clear_untrusted_proxy_headers': True}
    if peer:
        # Do not allow '*' to make every internet client a trusted proxy.
        ipaddress.ip_address(peer)
        options.update(trusted_proxy=peer, trusted_proxy_count=1,
                       trusted_proxy_headers={'x-forwarded-proto'})
    return options


def server_address() -> tuple[str, int]:
    host = os.environ.get('EATL_BIND_HOST', '127.0.0.1').strip()
    if not host:
        raise ValueError('EATL_BIND_HOST cannot be empty.')
    port = int(os.environ.get('EATL_PORT', '5050'))
    if not 1 <= port <= 65535:
        raise ValueError('EATL_PORT must be between 1 and 65535.')
    return host, port
