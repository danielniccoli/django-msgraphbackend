"""
Tests for the certificate credentials that replace the client secret.

A self-signed certificate is generated once for the module, and every token
request is answered by the fake Graph of the backend tests, so the tests check
the assertion the backend would send without any network traffic.
"""

from __future__ import annotations

import base64
import datetime
import json
import os
import tempfile
import unittest
import urllib.parse
from unittest import mock

from django.core.exceptions import ImproperlyConfigured
from django.test import override_settings

from msgraphbackend import MSGraphBackend
from msgraphbackend.certificate import CLIENT_ASSERTION_TYPE, ClientCertificate

from .test_backend import TOKEN_URL, FakeGraph

try:
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec, padding, rsa
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography.x509.oid import NameOID
except ImportError:  # pragma: no cover
    x509 = None

PASSWORD = "pfx-password"


def b64url_decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def make_pfx(key=None, password: str = PASSWORD, expired: bool = False):
    """Returns a PKCS#12 file with a self-signed certificate, and the certificate."""
    key = key or rsa.generate_private_key(public_exponent=65537, key_size=2048)
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "msgraphbackend")])
    now = datetime.datetime.now(datetime.timezone.utc)
    start = now - datetime.timedelta(days=30 if expired else 1)
    end = (
        now - datetime.timedelta(days=1)
        if expired
        else now + datetime.timedelta(days=30)
    )
    certificate = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(start)
        .not_valid_after(end)
        .sign(key, hashes.SHA256())
    )
    encryption = (
        serialization.BestAvailableEncryption(password.encode())
        if password
        else serialization.NoEncryption()
    )
    data = pkcs12.serialize_key_and_certificates(
        b"msgraphbackend", key, certificate, None, encryption
    )
    return data, certificate


def token_form(graph: FakeGraph) -> dict[str, str]:
    """Returns the form the backend posted to the token endpoint."""
    (request,) = graph.requests_to(TOKEN_URL)
    return dict(urllib.parse.parse_qsl(request.data.decode("utf-8")))


def make_backend(graph: FakeGraph, **kwargs) -> MSGraphBackend:
    with mock.patch("urllib.request.urlopen", graph):
        return MSGraphBackend(
            tenant_id="tenant-id", client_id="client-id", user_id="user-id", **kwargs
        )


@unittest.skipIf(x509 is None, "cryptography is not installed")
class ClientAssertionTests(unittest.TestCase):
    """The JWT that proves to Microsoft Entra that the key is at hand."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.pfx, cls.certificate = make_pfx()

    def assertion(self) -> tuple[dict, dict, bytes, bytes]:
        graph = FakeGraph()
        make_backend(
            graph, client_certificate=self.pfx, client_certificate_password=PASSWORD
        )
        form = token_form(graph)
        self.assertEqual(form["client_assertion_type"], CLIENT_ASSERTION_TYPE)
        header, claims, signature = form["client_assertion"].split(".")
        return (
            json.loads(b64url_decode(header)),
            json.loads(b64url_decode(claims)),
            f"{header}.{claims}".encode("ascii"),
            b64url_decode(signature),
        )

    def test_certificate_replaces_the_secret(self):
        graph = FakeGraph()
        make_backend(
            graph,
            client_secret="client-secret",
            client_certificate=self.pfx,
            client_certificate_password=PASSWORD,
        )
        form = token_form(graph)
        self.assertNotIn("client_secret", form)
        self.assertEqual(form["grant_type"], "client_credentials")
        self.assertIn("client_assertion", form)

    def test_header_names_the_certificate(self):
        header, _, _, _ = self.assertion()
        thumbprint = base64.urlsafe_b64encode(
            self.certificate.fingerprint(hashes.SHA256())
        ).rstrip(b"=")
        self.assertEqual(header["alg"], "PS256")
        self.assertEqual(header["x5t#S256"], thumbprint.decode("ascii"))

    def test_claims_name_the_client_and_the_token_endpoint(self):
        _, claims, _, _ = self.assertion()
        self.assertEqual(claims["aud"], TOKEN_URL)
        self.assertEqual(claims["iss"], "client-id")
        self.assertEqual(claims["sub"], "client-id")
        self.assertTrue(claims["jti"])
        self.assertLessEqual(claims["exp"] - claims["nbf"], 600)

    def test_signature_verifies_with_the_public_key(self):
        _, _, signed, signature = self.assertion()
        # verify() raises InvalidSignature on a mismatch.
        self.certificate.public_key().verify(
            signature,
            signed,
            padding.PSS(
                mgf=padding.MGF1(hashes.SHA256()),
                salt_length=padding.PSS.DIGEST_LENGTH,
            ),
            hashes.SHA256(),
        )

    def test_every_token_request_gets_a_new_assertion(self):
        backend = make_backend(
            FakeGraph(),
            client_certificate=self.pfx,
            client_certificate_password=PASSWORD,
        )
        first = backend.client_certificate.client_assertion("client-id", TOKEN_URL)
        second = backend.client_certificate.client_assertion("client-id", TOKEN_URL)
        self.assertNotEqual(first, second)


@unittest.skipIf(x509 is None, "cryptography is not installed")
class CertificateSettingsTests(unittest.TestCase):
    """Where the certificate is read from, and what is refused."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.pfx, _ = make_pfx()

    def test_base64_setting(self):
        graph = FakeGraph()
        encoded = base64.b64encode(self.pfx).decode("ascii")
        # A value wrapped over several lines is still read.
        wrapped = "\n".join(encoded[i : i + 76] for i in range(0, len(encoded), 76))
        with override_settings(
            MSGRAPH_CLIENT_CERTIFICATE=wrapped,
            MSGRAPH_CLIENT_CERTIFICATE_PASSWORD=PASSWORD,
        ):
            make_backend(graph)
        self.assertIn("client_assertion", token_form(graph))

    def test_path_setting(self):
        graph = FakeGraph()
        with tempfile.NamedTemporaryFile(suffix=".pfx", delete=False) as file:
            file.write(self.pfx)
        self.addCleanup(os.remove, file.name)
        with override_settings(
            MSGRAPH_CLIENT_CERTIFICATE_PATH=file.name,
            MSGRAPH_CLIENT_CERTIFICATE_PASSWORD=PASSWORD,
        ):
            make_backend(graph)
        self.assertIn("client_assertion", token_form(graph))

    def test_unreadable_path_is_refused(self):
        with (
            override_settings(MSGRAPH_CLIENT_CERTIFICATE_PATH="/nonexistent/cert.pfx"),
            self.assertRaises(ImproperlyConfigured),
        ):
            make_backend(FakeGraph())

    def test_secret_is_used_without_a_certificate(self):
        graph = FakeGraph()
        make_backend(graph, client_secret="client-secret")
        form = token_form(graph)
        self.assertEqual(form["client_secret"], "client-secret")
        self.assertNotIn("client_assertion", form)

    def test_no_credentials_are_refused(self):
        with self.assertRaises(ImproperlyConfigured):
            make_backend(FakeGraph())

    def test_wrong_password_is_refused(self):
        with self.assertRaises(ImproperlyConfigured):
            ClientCertificate.from_pfx(self.pfx, "wrong-password")

    def test_invalid_base64_is_refused(self):
        with self.assertRaises(ImproperlyConfigured):
            ClientCertificate.from_base64("not base64!")

    def test_pfx_without_password(self):
        pfx, _ = make_pfx(password="")
        self.assertTrue(ClientCertificate.from_pfx(pfx).thumbprint)

    def test_non_rsa_key_is_refused(self):
        pfx, _ = make_pfx(key=ec.generate_private_key(ec.SECP256R1()))
        with self.assertRaises(ImproperlyConfigured):
            ClientCertificate.from_pfx(pfx, PASSWORD)

    def test_expired_certificate_is_logged(self):
        pfx, _ = make_pfx(expired=True)
        with self.assertLogs("msgraphbackend.certificate", "WARNING"):
            ClientCertificate.from_pfx(pfx, PASSWORD)
