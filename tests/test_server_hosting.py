"""Public-origin and HTTPS proxy regressions; no live accounts or data."""
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from deployment.server_config import configure_hosting, public_origin, server_address, waitress_proxy_options

HOST = 'supersolemn-leontine-honorable.ngrok-free.dev'
ORIGIN = 'https://' + HOST


class HostingConfigurationTests(unittest.TestCase):
    def test_one_public_url_configures_exact_host_and_csrf_origin(self):
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {'EATL_PUBLIC_URL': ORIGIN}, clear=True):
            self.assertEqual(configure_hosting(Path(folder)), ORIGIN)
            self.assertEqual(os.environ['DJANGO_ALLOWED_HOSTS'], '127.0.0.1,localhost,[::1],' + HOST)
            self.assertEqual(os.environ['DJANGO_CSRF_TRUSTED_ORIGINS'], ORIGIN)
            self.assertEqual(os.environ['DJANGO_SECURE_COOKIES'], '1')
            configure_hosting(Path(folder))
            self.assertEqual(os.environ['DJANGO_ALLOWED_HOSTS'].count(HOST), 1)

    def test_source_and_executable_use_same_file_without_overriding_service_settings(self):
        from deployment.flowlab_launcher import load_configuration
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {'EATL_PORT':'5051'}, clear=True):
            Path(folder, 'configuration.env').write_text('EATL_PUBLIC_URL=' + ORIGIN + '\nEATL_PORT=5050\nDJANGO_ALLOWED_HOSTS=private.example\n', encoding='utf-8')
            load_configuration(Path(folder))
            self.assertEqual(server_address(), ('127.0.0.1', 5051))
            self.assertEqual(os.environ['DJANGO_ALLOWED_HOSTS'], 'private.example,' + HOST)

    def test_local_desktop_defaults_are_not_wildcard_or_secure_cookie_only(self):
        with tempfile.TemporaryDirectory() as folder, patch.dict(os.environ, {}, clear=True):
            configure_hosting(Path(folder))
            self.assertNotIn('*', os.environ['DJANGO_ALLOWED_HOSTS'])
            self.assertEqual(os.environ['DJANGO_SECURE_COOKIES'], '0')
            self.assertEqual(server_address(), ('127.0.0.1', 5050))

    def test_public_origin_validation(self):
        self.assertEqual(public_origin('https://LAB.example:443/'), 'https://lab.example')
        self.assertEqual(public_origin('http://[::1]:5050/'), 'http://[::1]:5050')
        for value in ('*', 'https://*.ngrok-free.dev', 'https://name:password@lab.example',
                      'https://lab.example/ei-lab', 'https://lab.example?x=1', 'https://lab.example:99999'):
            with self.subTest(value=value), self.assertRaises(ValueError):
                public_origin(value)

    def test_proxy_trust_is_exact_ip_and_scheme_only(self):
        with patch.dict(os.environ, {}, clear=True):
            options = waitress_proxy_options()
            self.assertEqual(options['trusted_proxy'], '127.0.0.1')
            self.assertEqual(options['trusted_proxy_headers'], {'x-forwarded-proto'})
        with patch.dict(os.environ, {'EATL_TRUSTED_PROXY':'*'}, clear=True), self.assertRaises(ValueError):
            waitress_proxy_options()
        with patch.dict(os.environ, {'EATL_TRUSTED_PROXY':''}, clear=True):
            self.assertEqual(waitress_proxy_options(), {'clear_untrusted_proxy_headers':True})


class PublicIntegratedHostingTests(unittest.TestCase):
    def setUp(self):
        from django.test import override_settings
        from integrated_app import application
        from web_app import app
        from waitress.proxy_headers import proxy_headers_middleware
        from werkzeug.test import Client
        from werkzeug.wrappers import Response
        previous_testing = app.config.get('TESTING', False)
        app.config['TESTING'] = False
        self.addCleanup(app.config.__setitem__, 'TESTING', previous_testing)
        self.settings = override_settings(ALLOWED_HOSTS=['127.0.0.1', HOST],
            CSRF_TRUSTED_ORIGINS=[ORIGIN], SESSION_COOKIE_SECURE=True, CSRF_COOKIE_SECURE=True)
        self.settings.enable()
        self.addCleanup(self.settings.disable)
        self.client = Client(proxy_headers_middleware(application, trusted_proxy='127.0.0.1',
            trusted_proxy_headers={'x-forwarded-proto'}, clear_untrusted=True), Response)

    def request(self, path, method='GET', data=None, origin=None):
        headers = {'X-Forwarded-Proto':'https'}
        if origin:
            headers.update(Origin=origin, Referer=origin + '/ei-lab/login/')
        return self.client.open(path, method=method, data=data, base_url='http://' + HOST,
            headers=headers, environ_overrides={'REMOTE_ADDR':'127.0.0.1'})

    def test_public_routes_and_assets_share_one_host_without_local_port_redirects(self):
        for path in ('/', '/login', '/ei-lab/login/', '/ei-lab/static/core/ei-module-theme.css'):
            with self.subTest(path=path):
                self.assertEqual(self.request(path).status_code, 200)
        response = self.request('/ei-lab/')
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers['Location'].startswith('/ei-lab/login/'))
        response = self.request('/ei-lab')
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers['Location'].startswith('/ei-lab/login/'))

    def test_ei_login_accepts_public_https_csrf_and_retains_own_session(self):
        from django.contrib.auth.models import User
        user = User.objects.create_user('public-hosting-only', password='Disposable-Hosting-Test-925!')
        self.addCleanup(user.delete)
        response = self.request('/ei-lab/login/')
        self.assertIn('Secure', response.headers.get('Set-Cookie', ''))
        token = self.client.get_cookie('ei_lab_csrftoken', domain=HOST, path='/ei-lab/').value
        result = self.request('/ei-lab/login/', 'POST', dict(username=user.username,
            password='Disposable-Hosting-Test-925!', csrfmiddlewaretoken=token), ORIGIN)
        self.assertEqual(result.status_code, 302, result.data[:500])
        self.assertIsNotNone(self.client.get_cookie('ei_lab_sessionid', domain=HOST, path='/ei-lab/'))
        self.assertEqual(self.request('/ei-lab/').status_code, 200)
        # E&I login must not silently log the same person into FlowLab.
        self.assertEqual(self.request('/dashboard').status_code, 302)

    def test_untrusted_host_and_forged_origin_are_still_rejected(self):
        response = self.client.get('/ei-lab/login/', base_url='https://attacker.example',
            environ_overrides={'REMOTE_ADDR':'127.0.0.1'})
        self.assertEqual(response.status_code, 400)
        self.request('/ei-lab/login/')
        token = self.client.get_cookie('ei_lab_csrftoken', domain=HOST, path='/ei-lab/').value
        response = self.request('/ei-lab/login/', 'POST', {'csrfmiddlewaretoken':token}, 'https://attacker.example')
        self.assertEqual(response.status_code, 403)

    def test_untrusted_peer_cannot_claim_https_or_replace_host(self):
        from waitress.proxy_headers import proxy_headers_middleware
        from werkzeug.test import Client
        from werkzeug.wrappers import Response
        captured = {}
        def echo(environ, start):
            captured.update(environ)
            start('200 OK', [('Content-Type','text/plain')])
            return [b'ok']
        client = Client(proxy_headers_middleware(echo, trusted_proxy='127.0.0.1',
            trusted_proxy_headers={'x-forwarded-proto'}, clear_untrusted=True), Response)
        client.get('/', base_url='http://localhost', headers={'X-Forwarded-Proto':'https',
            'X-Forwarded-Host':'attacker.example'}, environ_overrides={'REMOTE_ADDR':'203.0.113.10'})
        self.assertEqual(captured['wsgi.url_scheme'], 'http')
        self.assertNotIn('HTTP_X_FORWARDED_PROTO', captured)
        self.assertNotIn('HTTP_X_FORWARDED_HOST', captured)
