"""
Certificate credentials, the alternative to a client secret.

In the client credentials flow Microsoft Entra accepts, instead of the secret, a
JWT signed with the private key of a certificate whose public part is uploaded
to the app registration. Only the PKCS#12 form (a .pfx or .p12 file) is read,
because it holds the key and the certificate together.

The private key never leaves the server and a stolen assertion expires within
minutes, which is what certificate credentials have over a secret. Signing needs
the `cryptography` package, an optional dependency installed with the
`certificate` extra.
"""

from __future__ import annotations

import base64
import datetime
import json
import logging
import time
import uuid
from dataclasses import dataclass
from typing import TYPE_CHECKING

from django.core.exceptions import ImproperlyConfigured

if TYPE_CHECKING:
    from cryptography.hazmat.primitives.asymmetric.rsa import RSAPrivateKey


logger = logging.getLogger(__name__)

CLIENT_ASSERTION_TYPE = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
# Microsoft Entra rejects an assertion that is valid for more than ten minutes.
ASSERTION_LIFETIME = 600


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


@dataclass(frozen=True)
class ClientCertificate:
    """The private key and thumbprint that sign a client assertion."""

    private_key: RSAPrivateKey
    # The base64url encoded SHA-256 hash of the DER encoded certificate.
    thumbprint: str

    @classmethod
    def from_pfx(
        cls, data: bytes, password: str | bytes | None = None
    ) -> ClientCertificate:
        """Reads the key and the certificate out of the contents of a PKCS#12 file."""
        try:
            from cryptography.hazmat.primitives import hashes
            from cryptography.hazmat.primitives.asymmetric.rsa import RSAPrivateKey
            from cryptography.hazmat.primitives.serialization import pkcs12
        except ImportError as err:
            raise ImproperlyConfigured(
                "Certificate credentials need the cryptography package: "
                "install django-msgraphbackend[certificate]."
            ) from err
        if isinstance(password, str):
            password = password.encode("utf-8")
        try:
            key, certificate, _ = pkcs12.load_key_and_certificates(
                data, password or None
            )
        except ValueError as err:
            raise ImproperlyConfigured(
                "The Microsoft Graph client certificate could not be opened: "
                "the password is wrong or the file is not PKCS#12 (.pfx)."
            ) from err
        if key is None or certificate is None:
            raise ImproperlyConfigured(
                "The Microsoft Graph client certificate must hold both the "
                "private key and the certificate."
            )
        if not isinstance(key, RSAPrivateKey):
            raise ImproperlyConfigured(
                "Microsoft Entra only accepts client assertions signed with an RSA key."
            )
        expires = certificate.not_valid_after_utc
        if expires < datetime.datetime.now(datetime.timezone.utc):
            logger.warning(
                "The Microsoft Graph client certificate expired on %s.",
                expires.isoformat(),
            )
        return cls(key, _b64url(certificate.fingerprint(hashes.SHA256())))

    @classmethod
    def from_base64(
        cls, value: str, password: str | bytes | None = None
    ) -> ClientCertificate:
        """Reads a PKCS#12 file that was base64 encoded to fit in a setting."""
        try:
            data = base64.b64decode("".join(value.split()), validate=True)
        except ValueError as err:
            raise ImproperlyConfigured(
                "The Microsoft Graph client certificate is not valid base64."
            ) from err
        return cls.from_pfx(data, password)

    def client_assertion(self, client_id: str, audience: str) -> str:
        """
        Returns a JWT that proves the possession of the key to the token
        endpoint given as the audience.
        """
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.asymmetric import padding

        now = int(time.time())
        header = {"alg": "PS256", "typ": "JWT", "x5t#S256": self.thumbprint}
        claims = {
            "aud": audience,
            "iss": client_id,
            "sub": client_id,
            "jti": str(uuid.uuid4()),
            "iat": now,
            "nbf": now,
            "exp": now + ASSERTION_LIFETIME,
        }
        signing_input = ".".join(
            _b64url(json.dumps(part, separators=(",", ":")).encode("utf-8"))
            for part in (header, claims)
        ).encode("ascii")
        signature = self.private_key.sign(
            signing_input,
            padding.PSS(
                mgf=padding.MGF1(hashes.SHA256()),
                salt_length=padding.PSS.DIGEST_LENGTH,
            ),
            hashes.SHA256(),
        )
        return f"{signing_input.decode('ascii')}.{_b64url(signature)}"
