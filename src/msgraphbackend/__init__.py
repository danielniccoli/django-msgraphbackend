from __future__ import annotations

import base64
import json
import time
import typing
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass

from django.conf import settings
from django.core.exceptions import ImproperlyConfigured
from django.core.mail.backends.base import BaseEmailBackend
from django.core.mail.message import EmailMultiAlternatives

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
        use_json_api=False,
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
        self.use_json_api = getattr(settings, "MSGRAPH_USE_JSON_API", use_json_api)
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
        headers = self._prepare_headers()
        data = self._prepare_request_payload(email_message)
        request = urllib.request.Request(url, data=data, headers=headers)
        try:
            urllib.request.urlopen(request)
        except urllib.error.HTTPError as err:
            if self.fail_silently:
                return False
            # Error handling for Graph API responses
            response_body = err.read().decode('utf-8')
            try:
                error_details = json.loads(response_body)
                code = error_details.get("error", {}).get("code", "UNKNOWN_CODE")
                message = error_details.get("error", {}).get("message", "UNKNOWN_MESSAGE")
                err.add_note(f"Graph API Error: {code}: {message}")
            except json.JSONDecodeError:
                err.add_note(f"Graph API HTTP Error (Non-JSON Response): {response_body}")
            raise err

        return True

    def _prepare_headers(self) -> dict:
        """Prepare the headers for the request."""
        if not self._token:
            raise ValueError("The Microsoft Graph token is not set.")
        headers = {
            "Authorization": self._token.authorization_value,
        }
        if self.use_json_api:
            headers["Content-Type"] = "application/json"
        else:
            headers["Content-Type"] = "text/plain"
        return headers

    def _prepare_request_payload(self, email_message: EmailMessage) -> bytes:
        """
        Prepare the payload for the sendMail request.
        If use_json_api is True, it will convert the EmailMessage to a JSON format
        suitable for Microsoft Graph API. Otherwise, it will return the raw MIME
        message as a byte string.
        """
        if not self.use_json_api:
            # If not using JSON API, return the raw MIME message
            return base64.b64encode(email_message.message().as_bytes())

        # Build the message payload for Graph API
        message_payload = {
            "subject": email_message.subject,
            "toRecipients": [{"emailAddress": {"address": recipient}} for recipient in email_message.to],
            "from": {"emailAddress": {"address": email_message.from_email}},
            "body": {},
            "attachments": [],
        }

        # Handle body content (plain text and HTML)
        # EmailMultiAlternatives stores the plain text in .body and HTML in .alternatives
        # Graph API prefers HTML if both are present
        html_content = None
        if isinstance(email_message, EmailMultiAlternatives):
            for alt_content, alt_mimetype in email_message.alternatives:
                if alt_mimetype == 'text/html':
                    html_content = alt_content
                    break

        if html_content:
            message_payload["body"] = {
                "contentType": "html",
                "content": html_content
            }
        else:
            message_payload["body"] = {
                "contentType": "text",
                "content": email_message.body
            }

        # Handle CC recipients
        if email_message.cc:
            message_payload["ccRecipients"] = [{"emailAddress": {"address": cc}} for cc in email_message.cc]

        # Handle BCC recipients
        if email_message.bcc:
            message_payload["bccRecipients"] = [{"emailAddress": {"address": bcc}} for bcc in email_message.bcc]

        # Handle attachments
        for attachment in email_message.attachments:
            if isinstance(attachment, tuple):
                filename, content, mimetype = attachment
                # Graph API expects contentBytes to be base64 encoded
                encoded_content = base64.b64encode(content).decode('utf-8')
                message_payload["attachments"].append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentType": mimetype,
                    "contentBytes": encoded_content
                })
            # Handle here other attachment types if needed (e.g., Django's File objects)

        # Set the overall payload for the sendMail endpoint
        send_mail_payload = {
            "message": message_payload,
            "saveToSentItems": "true"  # Save a copy to the sender's Sent Items folder
        }

        return json.dumps(send_mail_payload).encode('utf-8')

    def _get_user(self, from_address: str) -> str:
        """Gets the user id who is assigned the from_address."""
        url = (
            "https://graph.microsoft.com/v1.0/users"
            f"?$filter=proxyAddresses/any(x:x%20eq%20'smtp:{from_address}')&$select=id"
        )
        if not self._token:
            raise ValueError("The Microsoft Graph token is not set.")
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
