from __future__ import annotations

import base64
import json
import logging
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import TYPE_CHECKING

from django import VERSION as DJANGO_VERSION
from django.conf import settings
from django.core.exceptions import ImproperlyConfigured
from django.core.mail.backends.base import BaseEmailBackend

from .attachments import (
    GraphAttachment,
    attachments_size,
    iter_attachments,
    iter_chunks,
)

if TYPE_CHECKING:
    import http.client
    from collections.abc import Sequence

    from django.core.mail.message import EmailMessage


logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class MSGraphToken:
    token_type: str
    expires_in: int
    ext_expires_in: int
    access_token: str
    expires_at: float = 0.0  # computed deadline
    ext_expires_at: float = 0.0  # computed deadline

    def __post_init__(self) -> None:
        expires_in = int(time.time() + self.expires_in)
        ext_expires_in = int(time.time() + self.ext_expires_in)
        object.__setattr__(self, "expires_at", expires_in)
        object.__setattr__(self, "ext_expires_at", ext_expires_in)

    @property
    def authorization_value(self) -> str:
        return f"{self.token_type} {self.access_token}"

    @property
    def is_valid(self) -> bool:
        return self.expires_at > time.time()


class MSGraphBackend(BaseEmailBackend):
    """A Django email backend that sends emails through the Microsoft Graph API."""

    # Graph rejects a sendMail request whose body exceeds about 4 MB. Messages
    # above this limit are sent as a draft with separately uploaded attachments.
    MAX_SENDMAIL_SIZE = 3_500_000
    # Graph only accepts an attachment that is posted in one request if the file
    # is smaller than 3 MB. Anything else needs an upload session.
    MAX_INLINE_ATTACHMENT_SIZE = 3 * 1024 * 1024
    # The largest byte range Microsoft allows per upload session request.
    UPLOAD_CHUNK_SIZE = 4 * 1024 * 1024

    def __init__(
        self,
        tenant_id: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
        user_id: str | None = None,
        fail_silently: bool = False,
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
        self._token: MSGraphToken | None = None
        self.open()

    def open(self) -> bool | None:
        """Gets a Microsoft Graph API token."""
        if self._token and self._token.is_valid:
            return True
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
        response = self._urlopen(request, "Failed to obtain Microsoft Graph API token.")
        if response is None:
            return None
        response_body = response.read().decode("utf-8")
        self._token = MSGraphToken(**json.loads(response_body))
        return True

    def send_messages(self, email_messages: Sequence[EmailMessage]) -> int:
        """
        Send one or more EmailMessage objects and return the number of email
        messages sent.
        """
        num_sent = 0
        if not email_messages:
            return num_sent
        if self.open() is None or self._token is None:
            return num_sent
        for message in email_messages:
            sent = self._send(message)
            if sent:
                num_sent += 1
        return num_sent

    def _send(self, email_message: EmailMessage) -> bool:
        """A helper method that does the actual sending."""
        if not email_message.recipients():
            return False
        # open() is called by send_messages() and sets self._token on success.
        # Assert documents that runtime invariant and narrows Optional for type checkers.
        assert self._token is not None
        user_id = self.user_id or self._get_user(email_message.from_email)
        if user_id is None:
            return False
        message = self._encode_for_sendmail(email_message)
        if message is None:
            # Too large for a single request, so send it the long way around.
            return self._send_large(email_message, user_id)
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/sendMail"
        headers = self._prepare_headers("text/plain")
        request = urllib.request.Request(url, data=message, headers=headers)
        response = self._urlopen(
            request, "Failed to send email via Microsoft Graph API."
        )
        return response is not None

    def _send_large(self, email_message: EmailMessage, user_id: str) -> bool:
        """
        A helper method that sends a message exceeding MAX_SENDMAIL_SIZE.

        Such a message cannot be sent in a single request, so it is created as a
        draft without its attachments first. Every attachment is then uploaded
        to that draft on its own, and finally the draft is sent. If uploading an
        attachment fails, the draft is deleted again, so that no unsent message
        is left behind in the mailbox.

        Unlike the single request, this requires the application permission
        Mail.ReadWrite, because the message exists in the mailbox while it is
        being assembled.
        """
        message_id = self._create_draft(email_message, user_id)
        if message_id is None:
            return False
        try:
            for attachment in iter_attachments(email_message):
                if not self._attach(attachment, message_id, user_id):
                    self._delete_draft(message_id, user_id)
                    return False
        except Exception:
            self._delete_draft(message_id, user_id)
            raise
        return self._send_draft(message_id, user_id)

    def _create_draft(self, email_message: EmailMessage, user_id: str) -> str | None:
        """Creates the message without its attachments and returns its id."""
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages"
        headers = self._prepare_headers("text/plain")
        data = self._prepare_draft_message(email_message)
        request = urllib.request.Request(url, data=data, headers=headers)
        response = self._urlopen(
            request, "Failed to create a draft message via Microsoft Graph API."
        )
        if response is None:
            return None
        return json.loads(response.read().decode("utf-8"))["id"]

    def _attach(
        self, attachment: GraphAttachment, message_id: str, user_id: str
    ) -> bool:
        """Adds a single attachment to a draft message."""
        if attachment.size >= self.MAX_INLINE_ATTACHMENT_SIZE:
            return self._upload_attachment(attachment, message_id, user_id)
        url = (
            f"https://graph.microsoft.com/v1.0/users/{user_id}"
            f"/messages/{message_id}/attachments"
        )
        headers = self._prepare_headers("application/json")
        data = json.dumps(attachment.as_file_attachment()).encode("utf-8")
        request = urllib.request.Request(url, data=data, headers=headers)
        response = self._urlopen(
            request,
            "Failed to attach '%s' via Microsoft Graph API.",
            attachment.name,
        )
        return response is not None

    def _upload_attachment(
        self, attachment: GraphAttachment, message_id: str, user_id: str
    ) -> bool:
        """Adds a large attachment to a draft message using an upload session."""
        url = (
            f"https://graph.microsoft.com/v1.0/users/{user_id}"
            f"/messages/{message_id}/attachments/createUploadSession"
        )
        headers = self._prepare_headers("application/json")
        data = json.dumps({"AttachmentItem": attachment.as_attachment_item()}).encode(
            "utf-8"
        )
        request = urllib.request.Request(url, data=data, headers=headers)
        response = self._urlopen(
            request,
            "Failed to create an upload session for '%s' via Microsoft Graph API.",
            attachment.name,
        )
        if response is None:
            return False
        upload_url = json.loads(response.read().decode("utf-8"))["uploadUrl"]
        for first_byte, last_byte, chunk in iter_chunks(
            attachment.content, self.UPLOAD_CHUNK_SIZE
        ):
            # The upload url is pre-authenticated and must be called without an
            # Authorization header.
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {first_byte}-{last_byte}/{attachment.size}",
            }
            request = urllib.request.Request(
                upload_url, data=chunk, headers=headers, method="PUT"
            )
            response = self._urlopen(
                request,
                "Failed to upload the bytes %s to %s of '%s' via Microsoft Graph API.",
                first_byte,
                last_byte,
                attachment.name,
            )
            if response is None:
                return False
        return True

    def _send_draft(self, message_id: str, user_id: str) -> bool:
        """Sends a draft message once all of its attachments are in place."""
        url = (
            f"https://graph.microsoft.com/v1.0/users/{user_id}"
            f"/messages/{message_id}/send"
        )
        headers = self._prepare_headers("application/json")
        request = urllib.request.Request(url, data=b"", headers=headers)
        response = self._urlopen(
            request, "Failed to send a draft message via Microsoft Graph API."
        )
        return response is not None

    def _delete_draft(self, message_id: str, user_id: str) -> None:
        """
        Deletes a draft message that could not be completed.

        This is a best effort cleanup that runs after something else already
        went wrong, so it never raises on its own and only reports what is left
        behind in the mailbox.
        """
        url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}"
        headers = self._prepare_headers("application/json")
        request = urllib.request.Request(url, headers=headers, method="DELETE")
        self._urlopen(
            request,
            "Failed to delete the draft message '%s' via Microsoft Graph API. "
            "It is left behind unsent in the mailbox.",
            message_id,
            fail_silently=True,
        )

    def _prepare_headers(self, content_type: str) -> dict:
        """Prepare the headers for a request against the Microsoft Graph API."""
        if not self._token:
            raise ValueError("The Microsoft Graph token is not set.")
        return {
            "Content-Type": content_type,
            "Authorization": self._token.authorization_value,
        }

    def _encode_message(self, email_message: EmailMessage) -> bytes:
        """
        Returns the MIME message of an EmailMessage, base64 encoded.

        The message is serialized with the CRLF line endings that RFC 5322
        requires, because Exchange does not decode a quoted-printable soft
        line break that ends in a bare LF: the `=` stays in the text and the
        character after it is lost.
        """
        if DJANGO_VERSION >= (6, 0):
            from email.policy import SMTPUTF8

            return base64.b64encode(
                email_message.message(policy=SMTPUTF8).as_bytes()  # pyrefly: ignore
            )
        return base64.b64encode(
            email_message.message().as_bytes(linesep="\r\n")  # pyrefly: ignore
        )

    def _encode_for_sendmail(self, email_message: EmailMessage) -> bytes | None:
        """
        Returns the encoded message if it fits into a single sendMail request.

        Encoding the message grows it by a third, and its attachments are in it
        at least at their own size, so once they alone reach three quarters of
        the limit the message cannot fit. Checking that first spares a large
        attachment from being encoded only to find out that it has to be
        uploaded separately anyway. Everything else is measured exactly.
        """
        if attachments_size(email_message) * 4 // 3 > self.MAX_SENDMAIL_SIZE:
            return None
        message = self._encode_message(email_message)
        if len(message) > self.MAX_SENDMAIL_SIZE:
            return None
        return message

    def _prepare_draft_message(self, email_message: EmailMessage) -> bytes:
        """
        Returns the MIME message that creates the draft of a large email.

        The draft holds everything but the attachments, which are uploaded to it
        separately afterwards. Serializing the message once more with its
        attachments temporarily removed leaves that work to Django.
        """
        original_attachments = email_message.attachments
        email_message.attachments = []
        try:
            return self._encode_message(email_message)
        finally:
            email_message.attachments = original_attachments

    def _get_user(self, from_address: str) -> str | None:
        """Gets the user id who is assigned the from_address and returns the user id."""
        # _get_user() is only reached from _send() after token acquisition.
        # Keep this assert so Optional is narrowed at this access site too.
        assert self._token is not None
        # Escape the quote (') -> ('') so input can't break out of the OData literal, then url-encode.
        proxy_address = "smtp:" + from_address.replace("'", "''")
        filter_expr = f"proxyAddresses/any(x:x eq '{proxy_address}')"
        query = urllib.parse.urlencode({"$filter": filter_expr, "$select": "id"})
        url = f"https://graph.microsoft.com/v1.0/users?{query}"
        headers = {
            "Authorization": f"{self._token.authorization_value}",
        }
        request = urllib.request.Request(url, headers=headers)
        response = self._urlopen(
            request, "Failed to query for Microsoft Entra ID user."
        )
        if response is None:
            return None
        response_body = response.read().decode("utf-8")
        users = json.loads(response_body)
        if len(users["value"]) == 0:
            if self.fail_silently:
                logger.error(
                    "No user found in Microsoft Entra ID with the smtp address '%s'.",
                    from_address,
                )
                return None
            else:
                raise ValueError(
                    f"No user found in Microsoft Entra ID with the smtp address '{from_address}'."
                )
        return users["value"][0]["id"]

    def _urlopen(
        self,
        request: urllib.request.Request,
        error_message: str,
        *args,
        fail_silently: bool | None = None,
    ) -> http.client.HTTPResponse | None:
        """
        Perform a request against the API and return the response.

        An error is logged and None is returned when the backend fails
        silently, and is raised with the error message of the API attached
        otherwise. A call that must not raise, such as a cleanup after an error,
        can ask to fail silently on its own.
        """
        if fail_silently is None:
            fail_silently = self.fail_silently
        try:
            return urllib.request.urlopen(request)
        except urllib.error.URLError as err:
            if isinstance(err, urllib.error.HTTPError):
                msgraph_error = err.read().decode("utf-8", errors="replace")
                # BaseException.add_note() needs Python 3.11 or newer.
                if hasattr(err, "add_note"):
                    err.add_note(f"Microsoft Graph API error: {msgraph_error}")
            else:
                msgraph_error = str(err)
            if fail_silently:
                logger.exception(
                    error_message, *args, extra={"msgraph_error": msgraph_error}
                )
                return None
            else:
                raise
