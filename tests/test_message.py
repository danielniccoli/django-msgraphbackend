"""Tests for the translation of a Django EmailMessage into a Graph message."""

from __future__ import annotations

import base64
import unittest
from email.mime.image import MIMEImage

from django.core.mail import EmailMessage, EmailMultiAlternatives

from msgraphbackend.message import graph_message, sendmail_payload


def make_message(**kwargs) -> EmailMessage:
    """Returns a message with sensible defaults that a test can override."""
    defaults = {
        "subject": "Subject",
        "body": "Body",
        "from_email": "sender@example.com",
        "to": ["recipient@example.com"],
    }
    return EmailMessage(**{**defaults, **kwargs})


def recipient(address: str, name: str | None = None) -> dict:
    """Returns the recipient resource that Graph expects."""
    email_address = {"address": address}
    if name:
        email_address["name"] = name
    return {"emailAddress": email_address}


class GraphMessageTests(unittest.TestCase):
    def test_minimal_message(self):
        self.assertEqual(
            graph_message(make_message()),
            {
                "subject": "Subject",
                "body": {"contentType": "text", "content": "Body"},
                "toRecipients": [recipient("recipient@example.com")],
                "from": recipient("sender@example.com"),
            },
        )

    def test_recipients_keep_their_display_names(self):
        message = make_message(
            from_email="Sender <sender@example.com>",
            to=["Récipient <recipient@example.com>", "other@example.com"],
            cc=["Copy <cc@example.com>"],
            bcc=["bcc@example.com"],
            reply_to=["Reply <reply@example.com>"],
        )

        resource = graph_message(message)

        self.assertEqual(resource["from"], recipient("sender@example.com", "Sender"))
        self.assertEqual(
            resource["toRecipients"],
            [
                recipient("recipient@example.com", "Récipient"),
                recipient("other@example.com"),
            ],
        )
        self.assertEqual(
            resource["ccRecipients"], [recipient("cc@example.com", "Copy")]
        )
        self.assertEqual(resource["bccRecipients"], [recipient("bcc@example.com")])
        self.assertEqual(resource["replyTo"], [recipient("reply@example.com", "Reply")])

    def test_empty_recipient_lists_are_left_out(self):
        resource = graph_message(make_message())

        for key in ("ccRecipients", "bccRecipients", "replyTo", "attachments"):
            self.assertNotIn(key, resource)

    def test_html_alternative_becomes_the_body(self):
        message = EmailMultiAlternatives(
            subject="Subject",
            body="Plain",
            from_email="sender@example.com",
            to=["recipient@example.com"],
            alternatives=[("<p>Rich</p>", "text/html")],
        )

        self.assertEqual(
            graph_message(message)["body"],
            {"contentType": "html", "content": "<p>Rich</p>"},
        )

    def test_html_content_subtype_becomes_an_html_body(self):
        message = make_message(body="<p>Rich</p>")
        message.content_subtype = "html"

        self.assertEqual(
            graph_message(message)["body"],
            {"contentType": "html", "content": "<p>Rich</p>"},
        )

    def test_other_alternatives_leave_the_text_body(self):
        message = EmailMultiAlternatives(
            subject="Subject",
            body="Plain",
            from_email="sender@example.com",
            to=["recipient@example.com"],
            alternatives=[("{}", "application/json")],
        )

        self.assertEqual(
            graph_message(message)["body"],
            {"contentType": "text", "content": "Plain"},
        )

    def test_from_and_reply_to_headers_override_the_attributes(self):
        message = make_message(
            reply_to=["ignored@example.com"],
            headers={
                "From": "Header <header@example.com>",
                "Reply-To": "one@example.com, Two <two@example.com>",
            },
        )

        resource = graph_message(message)

        self.assertEqual(resource["from"], recipient("header@example.com", "Header"))
        self.assertEqual(
            resource["replyTo"],
            [recipient("one@example.com"), recipient("two@example.com", "Two")],
        )
        self.assertNotIn("internetMessageHeaders", resource)

    def test_custom_headers_keep_their_name(self):
        message = make_message(headers={"X-Campaign": "spring", "x-id": "42"})

        self.assertEqual(
            graph_message(message)["internetMessageHeaders"],
            [
                {"name": "X-Campaign", "value": "spring"},
                {"name": "x-id", "value": "42"},
            ],
        )

    def test_other_headers_are_dropped_with_a_warning(self):
        message = make_message(
            headers={"List-Unsubscribe": "<mailto:stop@example.com>", "X-Kept": "1"}
        )

        with self.assertLogs("msgraphbackend.message", level="WARNING") as logs:
            resource = graph_message(message)

        self.assertEqual(
            resource["internetMessageHeaders"], [{"name": "X-Kept", "value": "1"}]
        )
        self.assertEqual(len(logs.records), 1)
        self.assertIn("List-Unsubscribe", logs.output[0])

    def test_attachments_are_file_attachments(self):
        message = make_message()
        message.attach("report.pdf", b"%PDF-1.7", "application/pdf")
        image = MIMEImage(b"\x89PNG", "png")
        image.add_header("Content-ID", "<logo>")
        message.attach(image)

        attachments = graph_message(message)["attachments"]

        self.assertEqual(
            attachments[0],
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": "report.pdf",
                "contentType": "application/pdf",
                "contentBytes": base64.b64encode(b"%PDF-1.7").decode("ascii"),
            },
        )
        self.assertEqual(attachments[1]["contentId"], "logo")
        self.assertIs(attachments[1]["isInline"], True)

    def test_attachments_can_be_left_out(self):
        message = make_message()
        message.attach("report.pdf", b"%PDF-1.7", "application/pdf")

        resource = graph_message(message, with_attachments=False)

        self.assertNotIn("attachments", resource)
        self.assertEqual(resource["subject"], "Subject")


class SendMailPayloadTests(unittest.TestCase):
    def test_message_is_wrapped(self):
        message = make_message()

        self.assertEqual(sendmail_payload(message), {"message": graph_message(message)})

    def test_save_to_sent_items_is_only_sent_when_false(self):
        message = make_message()

        payload = sendmail_payload(message, save_to_sent_items=False)

        self.assertIs(payload["saveToSentItems"], False)
        self.assertNotIn("saveToSentItems", sendmail_payload(message))


if __name__ == "__main__":
    unittest.main()
