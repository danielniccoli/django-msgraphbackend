from __future__ import annotations

import base64
import json
import time
import typing
import urllib.parse
import urllib.request
from dataclasses import dataclass

from django.conf import settings
from django.core.exceptions import ImproperlyConfigured
from django.core.mail.backends.base import BaseEmailBackend

if typing.TYPE_CHECKING:
    from django.core.mail.message import EmailMessage


@dataclass(frozen=True)
class MSGraphToken:
    token_type: str
    expires_in: int
    ext_expires_in: int
    access_token: str

    def __post_init__(self):
        expires_in = int(time.time() + self.expires_in)
        ext_expires_in = int(time.time() + self.ext_expires_in)
        object.__setattr__(self, "expires_in", expires_in)
        object.__setattr__(self, "ext_expires_in", ext_expires_in)

    @property
    def authorization_value(self):
        return f"{self.token_type} {self.access_token}"

    @property
    def is_valid(self):
        return self.expires_in > time.time()


class MSGraphBackend(BaseEmailBackend):
    def __init__(
        self,
        tenant_id=None,
        client_id=None,
        client_secret=None,
        user_id=None,
        fail_silently=False,
        **kwargs,
    ) -> None:
        super().__init__(fail_silently=fail_silently)
        if not tenant_id and not hasattr(settings, "MSGRAPH_TENANT_ID"):
            raise ImproperlyConfigured("The MSGRAPH_TENANT_ID setting must be set.")
        if not client_id and not hasattr(settings, "MSGRAPH_CLIENT_ID"):
            raise ImproperlyConfigured("The MSGRAPH_CLIENT_ID setting must be set.")
        if not client_secret and not hasattr(settings, "MSGRAPH_CLIENT_SECRET"):
            raise ImproperlyConfigured("The MSGRAPH_CLIENT_SECRET setting must be set.")
        self.tenant_id = tenant_id or settings.MSGRAPH_TENANT_ID
        self.client_id = client_id or settings.MSGRAPH_CLIENT_ID
        self.client_secret = client_secret or settings.MSGRAPH_CLIENT_SECRET
        self.user_id = getattr(settings, "MSGRAPH_USER_ID", user_id)
        self._token: None | MSGraphToken = None
        self.open()

    def open(self) -> None:
        """Gets a Microsoft Graph token."""
        if self._token and self._token.is_valid:
            return
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = urllib.parse.urlencode(
            {
                "client_id": self.client_id,
                "scope": "https://graph.microsoft.com/.default",
                "client_secret": self.client_secret,
                "grant_type": "client_credentials",
            }
        ).encode("utf-8")
        request = urllib.request.Request(url, data, headers)
        try:
            response = urllib.request.urlopen(request)
        except urllib.error.HTTPError:
            if not self.fail_silently:
                raise
        response_body = response.read().decode("utf-8")
        self._token = MSGraphToken(**json.loads(response_body))

    def send_messages(self, email_messages: list[EmailMessage]) -> int:
        """
        Send one or more EmailMessage objects and return the number of email
        messages sent.
        """
        num_sent = 0
        if not email_messages:
            return num_sent
        self.open()
        for message in email_messages:
            sent = self._send(message)
            if sent:
                num_sent += 1
        return num_sent

    def _send(self, email_message: EmailMessage) -> bool:
        """A helper method that does the actual sending."""
        if not email_message.recipients():
            return False
        user_id = self.user_id or self._get_user(email_message.from_email)
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/sendMail"
        message = base64.b64encode(email_message.message().as_bytes())
        headers = {
            "Content-Type": "text/plain",
            "Authorization": self._token.authorization_value,
        }
        request = urllib.request.Request(url, data=message, headers=headers)
        try:
            urllib.request.urlopen(request)
        except urllib.error.HTTPError as err:
            if self.fail_silently:
                return False
            error_details = json.load(err)
            code = error_details["error"]["code"]
            message = error_details["error"]["message"]
            err.add_note(f"{code}: {message}")
            raise
        return True

    def _get_user(self, from_address: str) -> str:
        """Gets the user id who is assigned the from_address."""
        url = (
            "https://graph.microsoft.com/v1.0/users"
            f"?$filter=proxyAddresses/any(x:x%20eq%20'smtp:{from_address}')&$select=id"
        )
        headers = {
            "Authorization": f"{self._token.authorization_value}",
        }
        request = urllib.request.Request(url, headers=headers)
        try:
            response = urllib.request.urlopen(request)
        except urllib.error.HTTPError as err:
            error_details = json.load(err)
            code = error_details["error"]["code"]
            message = error_details["error"]["message"]
            err.add_note(f"{code}: {message}")
            raise
        response_body = response.read().decode("utf-8")
        users = json.loads(response_body)
        if len(users["value"]) == 0:
            raise ValueError(
                f"No user found in Entra ID with the smtp address '{from_address}'."
            )

        return users["value"][0]["id"]
