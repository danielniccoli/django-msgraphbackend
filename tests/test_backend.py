"""
Tests for the email backend itself.

Every test replaces urllib.request.urlopen with a fake that records the requests
it is given, so the tests describe what the backend sends to the Microsoft Graph
API without any network traffic.
"""

from __future__ import annotations

import base64
import contextlib
import io
import json
import unittest
import urllib.error
from email import message_from_bytes
from email.mime.image import MIMEImage
from typing import TYPE_CHECKING
from unittest import mock

from django.core.mail import EmailMessage, EmailMultiAlternatives

from msgraphbackend import MSGraphBackend

if TYPE_CHECKING:
    from collections.abc import Iterator

TOKEN_URL = "https://login.microsoftonline.com/tenant-id/oauth2/v2.0/token"
USER_URL = "https://graph.microsoft.com/v1.0/users/user-id"
MESSAGES_URL = f"{USER_URL}/messages"
DRAFT_URL = f"{MESSAGES_URL}/draft-id"
UPLOAD_URL = "https://upload.example.com/session"


class FakeResponse:
    """The part of an HTTP response that the backend reads."""

    def __init__(self, payload: dict) -> None:
        self._body = json.dumps(payload).encode("utf-8")

    def read(self) -> bytes:
        return self._body


class FakeGraph:
    """
    A replacement for urlopen that answers as the Graph API would.

    The requests it is given are recorded, so that a test can assert what the
    backend sent. A request whose url contains fail_on fails instead, which is
    how the tests provoke an error in the middle of a sequence of requests.
    """

    def __init__(self, fail_on: str | None = None) -> None:
        self.requests: list = []
        self.fail_on = fail_on

    def __call__(self, request, *args, **kwargs) -> FakeResponse:
        self.requests.append(request)
        url = request.full_url
        if self.fail_on and self.fail_on in url:
            raise urllib.error.HTTPError(
                url, 400, "Bad Request", {}, io.BytesIO(b'{"error": "invalid"}')
            )
        return FakeResponse(self._payload_for(url))

    def _payload_for(self, url: str) -> dict:
        if url == TOKEN_URL:
            return {
                "token_type": "Bearer",
                "expires_in": 3600,
                "ext_expires_in": 3600,
                "access_token": "access-token",
            }
        if url.startswith("https://graph.microsoft.com/v1.0/users?"):
            return {"value": [{"id": "user-id"}]}
        if url.endswith("/createUploadSession"):
            return {"uploadUrl": UPLOAD_URL}
        if url.endswith("/messages"):
            return {"id": "draft-id"}
        return {}

    @property
    def urls(self) -> list[str]:
        return [request.full_url for request in self.requests]

    @property
    def methods(self) -> list[str]:
        return [request.get_method() for request in self.requests]

    def requests_to(self, url: str) -> list:
        return [request for request in self.requests if request.full_url == url]


def make_backend(**kwargs) -> MSGraphBackend:
    """Returns a backend that already holds a token."""
    with mock.patch("urllib.request.urlopen", FakeGraph()):
        return MSGraphBackend(
            tenant_id="tenant-id",
            client_id="client-id",
            client_secret="client-secret",
            user_id=kwargs.pop("user_id", "user-id"),
            **kwargs,
        )


def send(message: EmailMessage, graph: FakeGraph, **kwargs) -> int:
    """Sends a message through a backend that talks to the given fake."""
    backend = make_backend(**kwargs)
    with mock.patch("urllib.request.urlopen", graph):
        return backend.send_messages([message])


def make_message(*attachments) -> EmailMessage:
    """Returns a message with the given attachments."""
    message = EmailMessage(
        subject="Subject",
        body="Body",
        from_email="sender@example.com",
        to=["recipient@example.com"],
    )
    for attachment in attachments:
        message.attach(*attachment)
    return message


def payload_of(request) -> dict:
    """Returns the JSON body of a request."""
    return json.loads(request.data.decode("utf-8"))


def header_of(request, name: str) -> str | None:
    """Returns a header of a request, which urllib stores capitalized."""
    return request.headers.get(name.capitalize())


def body_of(part) -> bytes:
    """Returns the decoded text of a MIME part without its trailing line break."""
    # Django 6 ends a text part with a line break, Django 5 does not.
    return part.get_payload(decode=True).rstrip(b"\r\n")


@contextlib.contextmanager
def encoded_attachment_counts() -> Iterator[list[int]]:
    """
    Records how many attachments each message had when the backend encoded it.

    The draft of a large message is encoded with its attachments removed and
    restored afterwards, so the count has to be taken while it is encoded.
    """
    counts: list[int] = []
    encode_message = MSGraphBackend._encode_message

    def spy(backend, email_message, **kwargs):
        counts.append(len(email_message.attachments))
        return encode_message(backend, email_message, **kwargs)

    with mock.patch.object(MSGraphBackend, "_encode_message", spy):
        yield counts


class SendMailTests(unittest.TestCase):
    """The single request that sends everything that is small enough."""

    def test_message_is_sent_as_mime_in_one_request(self):
        graph = FakeGraph()
        message = make_message(("notes.txt", b"a note", "text/plain"))

        self.assertEqual(send(message, graph), 1)

        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])
        request = graph.requests[0]
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(header_of(request, "Authorization"), "Bearer access-token")
        self.assertEqual(header_of(request, "Content-Type"), "text/plain")
        # The whole message travels in one request, attachment included.
        sent = message_from_bytes(base64.b64decode(request.data))
        self.assertEqual(sent["Subject"], "Subject")
        self.assertEqual(sent["From"], "sender@example.com")
        self.assertEqual(sent["To"], "recipient@example.com")
        body, attachment = sent.get_payload()
        self.assertEqual(body_of(body), b"Body")
        self.assertEqual(attachment.get_filename(), "notes.txt")
        self.assertEqual(body_of(attachment), b"a note")

    def test_message_is_encoded_once(self):
        graph = FakeGraph()
        message = make_message(("notes.txt", b"a note", "text/plain"))

        with encoded_attachment_counts() as counts:
            self.assertEqual(send(message, graph), 1)

        self.assertEqual(counts, [1])
        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])

    def test_message_is_sent_with_crlf_line_endings(self):
        # A line longer than 998 bytes makes Django encode the text parts as
        # quoted-printable, and Exchange decodes the soft line breaks of that
        # encoding only when they end in CRLF.
        graph = FakeGraph()
        text = "<p>" + "word " * 300 + "</p>"
        message = EmailMultiAlternatives(
            subject="Subject",
            body=text,
            from_email="sender@example.com",
            to=["recipient@example.com"],
            alternatives=[(text, "text/html")],
        )

        self.assertEqual(send(message, graph), 1)

        raw = base64.b64decode(graph.requests[0].data)
        self.assertEqual(raw.replace(b"\r\n", b"").count(b"\n"), 0)
        plain, html = message_from_bytes(raw).get_payload()
        for part in (plain, html):
            self.assertEqual(part["Content-Transfer-Encoding"], "quoted-printable")
            self.assertEqual(body_of(part), text.encode("utf-8"))

    def test_sender_is_looked_up_when_no_user_id_is_configured(self):
        graph = FakeGraph()

        self.assertEqual(send(make_message(), graph, user_id=None), 1)

        self.assertEqual(len(graph.urls), 2)
        self.assertIn("proxyAddresses", graph.urls[0])
        self.assertIn("sender%40example.com", graph.urls[0])
        self.assertEqual(graph.urls[1], f"{USER_URL}/sendMail")

    def test_message_without_recipients_is_not_sent(self):
        graph = FakeGraph()
        message = EmailMessage(subject="Subject", body="Body", to=[])

        self.assertEqual(send(message, graph), 0)

        self.assertEqual(graph.requests, [])

    def test_error_is_raised_when_not_failing_silently(self):
        graph = FakeGraph(fail_on="/sendMail")

        with self.assertRaises(urllib.error.HTTPError):
            send(make_message(), graph)

    def test_error_is_swallowed_when_failing_silently(self):
        graph = FakeGraph(fail_on="/sendMail")

        with self.assertLogs("msgraphbackend", level="ERROR"):
            self.assertEqual(send(make_message(), graph, fail_silently=True), 0)


class JsonSendMailTests(unittest.TestCase):
    """The same requests, carrying the message as a Graph JSON resource."""

    def test_message_is_sent_as_json_in_one_request(self):
        graph = FakeGraph()
        message = EmailMultiAlternatives(
            subject="Subject",
            body="Body",
            from_email="Sender <sender@example.com>",
            to=["recipient@example.com"],
            alternatives=[("<p>Body</p>", "text/html")],
        )
        message.attach("notes.txt", b"a note", "text/plain")

        self.assertEqual(send(message, graph, use_json_api=True), 1)

        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])
        request = graph.requests[0]
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(header_of(request, "Authorization"), "Bearer access-token")
        self.assertEqual(header_of(request, "Content-Type"), "application/json")
        payload = payload_of(request)
        self.assertEqual(list(payload), ["message"])
        sent = payload["message"]
        self.assertEqual(sent["subject"], "Subject")
        self.assertEqual(
            sent["from"],
            {"emailAddress": {"address": "sender@example.com", "name": "Sender"}},
        )
        self.assertEqual(
            sent["toRecipients"],
            [{"emailAddress": {"address": "recipient@example.com"}}],
        )
        self.assertEqual(
            sent["body"], {"contentType": "html", "content": "<p>Body</p>"}
        )
        (attachment,) = sent["attachments"]
        self.assertEqual(attachment["name"], "notes.txt")
        self.assertEqual(base64.b64decode(attachment["contentBytes"]), b"a note")

    def test_large_message_is_drafted_as_json_without_attachments(self):
        graph = FakeGraph()
        message = make_message(
            ("first.bin", b"\x01" * 2_000_000, "application/octet-stream"),
            ("second.bin", b"\x02" * 2_000_000, "application/octet-stream"),
        )

        with encoded_attachment_counts() as counts:
            self.assertEqual(send(message, graph, use_json_api=True), 1)

        # The attachments alone are too large, so only the draft is encoded.
        self.assertEqual(counts, [0])
        self.assertEqual(
            graph.urls,
            [
                MESSAGES_URL,
                f"{DRAFT_URL}/attachments",
                f"{DRAFT_URL}/attachments",
                f"{DRAFT_URL}/send",
            ],
        )
        draft_request = graph.requests[0]
        self.assertEqual(header_of(draft_request, "Content-Type"), "application/json")
        # The draft is the message resource itself, not the sendMail payload.
        draft = payload_of(draft_request)
        self.assertEqual(draft["subject"], "Subject")
        self.assertNotIn("message", draft)
        self.assertNotIn("attachments", draft)
        self.assertEqual(
            [
                payload_of(request)["name"]
                for request in graph.requests_to(f"{DRAFT_URL}/attachments")
            ],
            ["first.bin", "second.bin"],
        )
        self.assertEqual(
            [attachment.filename for attachment in message.attachments],
            ["first.bin", "second.bin"],
        )

    def test_json_message_near_the_limit_is_measured_exactly(self):
        graph = FakeGraph()
        # A JSON payload carries its attachments base64 encoded as well, so the
        # attachments alone fill the limit exactly, and the rest of the payload
        # tips the message over it. Only encoding the whole message shows that.
        message = make_message(
            ("first.bin", b"\x01" * 1_312_500, "application/octet-stream"),
            ("second.bin", b"\x02" * 1_312_500, "application/octet-stream"),
        )

        with encoded_attachment_counts() as counts:
            self.assertEqual(send(message, graph, use_json_api=True), 1)

        self.assertEqual(counts, [2, 0])
        self.assertEqual(graph.urls[0], MESSAGES_URL)


class SendMailLimitTests(unittest.TestCase):
    """The settings that move or remove the limit of the single request."""

    def make_large_message(self) -> EmailMessage:
        """Returns a message that is too large for the default limit."""
        return make_message(
            ("first.bin", b"\x01" * 2_000_000, "application/octet-stream"),
            ("second.bin", b"\x02" * 2_000_000, "application/octet-stream"),
        )

    def test_default_limit_is_the_class_attribute(self):
        self.assertEqual(
            make_backend().max_sendmail_size, MSGraphBackend.MAX_SENDMAIL_SIZE
        )

    def test_raised_limit_keeps_a_larger_message_in_one_request(self):
        graph = FakeGraph()
        message = self.make_large_message()

        sent = send(message, graph, max_sendmail_size=10_000_000)

        self.assertEqual(sent, 1)
        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])
        raw = message_from_bytes(base64.b64decode(graph.requests[0].data))
        self.assertEqual(
            [part.get_filename() for part in raw.get_payload()[1:]],
            ["first.bin", "second.bin"],
        )

    def test_lowered_limit_sends_a_small_message_as_a_draft(self):
        graph = FakeGraph()
        message = make_message(("notes.txt", b"a note", "text/plain"))

        self.assertEqual(send(message, graph, max_sendmail_size=100), 1)

        self.assertEqual(graph.urls[0], MESSAGES_URL)
        self.assertEqual(graph.urls[-1], f"{DRAFT_URL}/send")

    def test_limit_is_read_from_the_settings(self):
        graph = FakeGraph()

        with mock.patch(
            "django.conf.settings.MSGRAPH_MAX_SENDMAIL_SIZE", 10_000_000, create=True
        ):
            self.assertEqual(send(self.make_large_message(), graph), 1)

        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])

    def test_always_use_sendmail_skips_the_draft(self):
        graph = FakeGraph()
        # Far too large for the limit, yet the message is neither measured nor
        # sent as a draft.
        message = make_message(
            ("huge.bin", b"\x03" * 5_000_000, "application/octet-stream")
        )

        with encoded_attachment_counts() as counts:
            self.assertEqual(send(message, graph, always_use_sendmail=True), 1)

        self.assertEqual(counts, [1])
        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])

    def test_always_use_sendmail_with_json(self):
        graph = FakeGraph()
        message = self.make_large_message()

        sent = send(message, graph, use_json_api=True, always_use_sendmail=True)

        self.assertEqual(sent, 1)
        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])
        self.assertEqual(
            [
                a["name"]
                for a in payload_of(graph.requests[0])["message"]["attachments"]
            ],
            ["first.bin", "second.bin"],
        )

    def test_always_use_sendmail_surfaces_the_rejection(self):
        graph = FakeGraph(fail_on="/sendMail")

        with self.assertRaises(urllib.error.HTTPError):
            send(self.make_large_message(), graph, always_use_sendmail=True)

        # Nothing was drafted, so nothing had to be cleaned up.
        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])

    def test_always_use_sendmail_is_read_from_the_settings(self):
        graph = FakeGraph()

        with mock.patch(
            "django.conf.settings.MSGRAPH_ALWAYS_USE_SENDMAIL", True, create=True
        ):
            self.assertEqual(send(self.make_large_message(), graph), 1)

        self.assertEqual(graph.urls, [f"{USER_URL}/sendMail"])

    def test_json_api_is_read_from_the_settings(self):
        graph = FakeGraph()

        with mock.patch("django.conf.settings.MSGRAPH_USE_JSON_API", True, create=True):
            self.assertEqual(send(make_message(), graph), 1)

        self.assertEqual(
            header_of(graph.requests[0], "Content-Type"), "application/json"
        )


class SendLargeMailTests(unittest.TestCase):
    """The draft that carries what does not fit into a single request."""

    def make_large_message(self) -> EmailMessage:
        """Returns a message that is too large for a single request."""
        return make_message(
            ("first.bin", b"\x01" * 2_000_000, "application/octet-stream"),
            ("second.bin", b"\x02" * 2_000_000, "application/octet-stream"),
        )

    def test_draft_is_created_without_attachments_and_sent(self):
        graph = FakeGraph()
        message = self.make_large_message()

        self.assertEqual(send(message, graph), 1)

        self.assertEqual(
            graph.urls,
            [
                MESSAGES_URL,
                f"{DRAFT_URL}/attachments",
                f"{DRAFT_URL}/attachments",
                f"{DRAFT_URL}/send",
            ],
        )
        self.assertEqual(graph.methods, ["POST"] * 4)
        draft = message_from_bytes(base64.b64decode(graph.requests[0].data))
        self.assertEqual(header_of(graph.requests[0], "Content-Type"), "text/plain")
        self.assertEqual(draft["Subject"], "Subject")
        self.assertEqual(draft["To"], "recipient@example.com")
        self.assertFalse(draft.is_multipart())
        self.assertEqual(body_of(draft), b"Body")

    def test_oversized_attachments_are_not_encoded_for_the_single_request(self):
        graph = FakeGraph()
        # The attachments alone are too large, so only the draft is encoded.
        message = self.make_large_message()

        with encoded_attachment_counts() as counts:
            self.assertEqual(send(message, graph), 1)

        self.assertEqual(counts, [0])
        self.assertEqual(graph.urls[0], MESSAGES_URL)

    def test_message_near_the_limit_is_measured_exactly(self):
        graph = FakeGraph()
        # The attachments alone leave room for the message, but encoded, the
        # message exceeds the limit, so it is encoded whole before the draft.
        message = make_message(
            ("first.bin", b"\x01" * 1_300_000, "application/octet-stream"),
            ("second.bin", b"\x02" * 1_300_000, "application/octet-stream"),
        )

        with encoded_attachment_counts() as counts:
            self.assertEqual(send(message, graph), 1)

        self.assertEqual(counts, [2, 0])
        self.assertEqual(graph.urls[0], MESSAGES_URL)

    def test_attachments_are_posted_to_the_draft(self):
        graph = FakeGraph()

        send(self.make_large_message(), graph)

        attachments = [
            payload_of(request)
            for request in graph.requests_to(f"{DRAFT_URL}/attachments")
        ]
        self.assertEqual(
            [attachment["name"] for attachment in attachments],
            ["first.bin", "second.bin"],
        )
        for attachment, byte in zip(attachments, (b"\x01", b"\x02"), strict=True):
            self.assertEqual(
                attachment["@odata.type"], "#microsoft.graph.fileAttachment"
            )
            self.assertEqual(attachment["contentType"], "application/octet-stream")
            self.assertEqual(
                base64.b64decode(attachment["contentBytes"]), byte * 2_000_000
            )

    def test_message_keeps_its_attachments(self):
        graph = FakeGraph()
        message = self.make_large_message()

        send(message, graph)

        self.assertEqual(
            [attachment.filename for attachment in message.attachments],
            ["first.bin", "second.bin"],
        )

    def test_inline_attachment_stays_inline(self):
        graph = FakeGraph()
        image = MIMEImage(b"\x89PNG" + b"\x00" * 1024, "png")
        image.add_header("Content-ID", "<logo>")
        image.add_header("Content-Disposition", "inline", filename="logo.png")
        message = EmailMultiAlternatives(
            subject="Subject",
            body="Body",
            from_email="sender@example.com",
            to=["recipient@example.com"],
            alternatives=[('<img src="cid:logo">', "text/html")],
        )
        message.attach(image)
        message.attach("large.bin", b"\x00" * 3_000_000, "application/octet-stream")

        self.assertEqual(send(message, graph), 1)

        attachments = [
            payload_of(request)
            for request in graph.requests_to(f"{DRAFT_URL}/attachments")
        ]
        self.assertEqual(attachments[0]["name"], "logo.png")
        self.assertEqual(attachments[0]["contentType"], "image/png")
        self.assertEqual(attachments[0]["contentId"], "logo")
        self.assertIs(attachments[0]["isInline"], True)
        self.assertNotIn("isInline", attachments[1])

    def test_draft_is_deleted_when_an_attachment_fails(self):
        graph = FakeGraph(fail_on="/attachments")

        with self.assertLogs("msgraphbackend", level="ERROR"):
            sent = send(self.make_large_message(), graph, fail_silently=True)

        self.assertEqual(sent, 0)
        self.assertEqual(
            graph.urls, [MESSAGES_URL, f"{DRAFT_URL}/attachments", DRAFT_URL]
        )
        self.assertEqual(graph.methods[-1], "DELETE")

    def test_draft_is_deleted_when_an_attachment_raises(self):
        graph = FakeGraph(fail_on="/attachments")

        with self.assertRaises(urllib.error.HTTPError):
            send(self.make_large_message(), graph)

        self.assertEqual(graph.urls[-1], DRAFT_URL)
        self.assertEqual(graph.methods[-1], "DELETE")

    def test_failing_cleanup_does_not_hide_the_original_error(self):
        graph = FakeGraph(fail_on=DRAFT_URL)

        # The logs are asserted around the error, so that the failing cleanup is
        # reported even though the error of the attachment is the one raised.
        with (
            self.assertLogs("msgraphbackend", level="ERROR"),
            self.assertRaises(urllib.error.HTTPError),
        ):
            send(self.make_large_message(), graph)

        self.assertEqual(graph.methods[-1], "DELETE")


class UploadSessionTests(unittest.TestCase):
    """The upload session that carries an attachment of its own."""

    def make_message_with_large_attachment(self) -> EmailMessage:
        return make_message(
            ("huge.bin", b"\x03" * 5_000_000, "application/octet-stream")
        )

    def test_attachment_is_uploaded_in_chunks(self):
        graph = FakeGraph()

        self.assertEqual(send(self.make_message_with_large_attachment(), graph), 1)

        self.assertEqual(
            graph.urls,
            [
                MESSAGES_URL,
                f"{DRAFT_URL}/attachments/createUploadSession",
                UPLOAD_URL,
                UPLOAD_URL,
                f"{DRAFT_URL}/send",
            ],
        )
        session = payload_of(graph.requests[1])
        self.assertEqual(
            session["AttachmentItem"],
            {
                "attachmentType": "file",
                "name": "huge.bin",
                "size": 5_000_000,
                "contentType": "application/octet-stream",
            },
        )

    def test_chunks_cover_the_attachment(self):
        graph = FakeGraph()

        send(self.make_message_with_large_attachment(), graph)

        chunks = graph.requests_to(UPLOAD_URL)
        self.assertEqual([chunk.get_method() for chunk in chunks], ["PUT", "PUT"])
        self.assertEqual(
            [header_of(chunk, "Content-Range") for chunk in chunks],
            ["bytes 0-4194303/5000000", "bytes 4194304-4999999/5000000"],
        )
        self.assertEqual(
            [header_of(chunk, "Content-Length") for chunk in chunks],
            ["4194304", "805696"],
        )
        self.assertEqual(b"".join(chunk.data for chunk in chunks), b"\x03" * 5_000_000)

    def test_upload_url_is_called_without_credentials(self):
        graph = FakeGraph()

        send(self.make_message_with_large_attachment(), graph)

        for chunk in graph.requests_to(UPLOAD_URL):
            self.assertIsNone(header_of(chunk, "Authorization"))

    def test_draft_is_deleted_when_a_chunk_fails(self):
        graph = FakeGraph(fail_on=UPLOAD_URL)

        with self.assertLogs("msgraphbackend", level="ERROR"):
            sent = send(
                self.make_message_with_large_attachment(), graph, fail_silently=True
            )

        self.assertEqual(sent, 0)
        self.assertEqual(graph.urls[-1], DRAFT_URL)
        self.assertEqual(graph.methods[-1], "DELETE")
        self.assertNotIn(f"{DRAFT_URL}/send", graph.urls)


if __name__ == "__main__":
    unittest.main()
