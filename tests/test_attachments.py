"""Tests for the translation of Django attachments into Graph attachments."""

from __future__ import annotations

import base64
import unittest
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage

from django.core.mail import EmailMessage

from msgraphbackend.attachments import (
    GraphAttachment,
    attachments_size,
    iter_attachments,
    iter_chunks,
)


def message_with(*attachments) -> EmailMessage:
    """Returns a message with the given attachments."""
    message = EmailMessage(subject="Subject", body="Body", to=["to@example.com"])
    for attachment in attachments:
        if isinstance(attachment, MIMEBase):
            message.attach(attachment)
        else:
            message.attach(*attachment)
    return message


def attachments_of(*attachments) -> list[GraphAttachment]:
    """Returns the Graph attachments of a message with the given attachments."""
    return list(iter_attachments(message_with(*attachments)))


class IterAttachmentsTests(unittest.TestCase):
    def test_binary_attachment(self):
        (attachment,) = attachments_of(("report.pdf", b"%PDF-1.7", "application/pdf"))

        self.assertEqual(attachment.name, "report.pdf")
        self.assertEqual(attachment.content, b"%PDF-1.7")
        self.assertEqual(attachment.mimetype, "application/pdf")
        self.assertEqual(attachment.size, 8)
        self.assertFalse(attachment.is_inline)

    def test_text_attachment_is_encoded_with_a_declared_charset(self):
        # Django keeps the content of a text attachment as a string.
        (attachment,) = attachments_of(("notes.txt", "Grüße", "text/plain"))

        self.assertEqual(attachment.content, "Grüße".encode())
        self.assertEqual(attachment.mimetype, 'text/plain; charset="utf-8"')

    def test_attachment_without_a_name_gets_one(self):
        (attachment,) = attachments_of((None, b"content", "application/octet-stream"))

        self.assertEqual(attachment.name, "attachment")

    def test_mime_part(self):
        part = MIMEImage(b"\x89PNG", "png")
        part.add_header("Content-Disposition", "attachment", filename="chart.png")

        (attachment,) = attachments_of(part)

        self.assertEqual(attachment.name, "chart.png")
        self.assertEqual(attachment.content, b"\x89PNG")
        self.assertEqual(attachment.mimetype, "image/png")
        self.assertIsNone(attachment.content_id)
        self.assertFalse(attachment.is_inline)

    def test_mime_part_with_a_content_id_is_inline(self):
        part = MIMEImage(b"\x89PNG", "png")
        part.add_header("Content-ID", "<logo>")

        (attachment,) = attachments_of(part)

        self.assertEqual(attachment.content_id, "logo")
        self.assertTrue(attachment.is_inline)

    def test_several_attachments_keep_their_order(self):
        attachments = attachments_of(
            ("first.bin", b"1", "application/octet-stream"),
            ("second.bin", b"2", "application/octet-stream"),
        )

        self.assertEqual(
            [attachment.name for attachment in attachments], ["first.bin", "second.bin"]
        )


class GraphAttachmentTests(unittest.TestCase):
    def test_file_attachment_resource(self):
        attachment = GraphAttachment("report.pdf", b"%PDF-1.7", "application/pdf")

        self.assertEqual(
            attachment.as_file_attachment(),
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": "report.pdf",
                "contentType": "application/pdf",
                "contentBytes": base64.b64encode(b"%PDF-1.7").decode("ascii"),
            },
        )

    def test_attachment_item_resource(self):
        attachment = GraphAttachment("logo.png", b"\x89PNG", "image/png", "logo", True)

        self.assertEqual(
            attachment.as_attachment_item(),
            {
                "attachmentType": "file",
                "name": "logo.png",
                "size": 4,
                "contentType": "image/png",
                "contentId": "logo",
                "isInline": True,
            },
        )


class AttachmentsSizeTests(unittest.TestCase):
    def test_binary_attachment(self):
        message = message_with(("report.pdf", b"%PDF-1.7", "application/pdf"))

        self.assertEqual(attachments_size(message), 8)

    def test_text_attachment_counts_its_characters(self):
        message = message_with(("notes.txt", "Grüße", "text/plain"))

        self.assertEqual(attachments_size(message), 5)

    def test_mime_part_counts_its_encoded_payload(self):
        part = MIMEImage(b"\x89PNG", "png")

        self.assertEqual(attachments_size(message_with(part)), len(part.get_payload()))

    def test_attached_email_object_counts_as_zero(self):
        attached = EmailMessage(subject="Attached", body="Body", to=["to@example.com"])
        message = message_with(("attached.eml", attached, "message/rfc822"))

        self.assertEqual(attachments_size(message), 0)

    def test_several_attachments_are_summed(self):
        message = message_with(
            ("first.bin", b"12", "application/octet-stream"),
            ("second.bin", b"345", "application/octet-stream"),
        )

        self.assertEqual(attachments_size(message), 5)

    def test_message_without_attachments(self):
        self.assertEqual(attachments_size(message_with()), 0)


class IterChunksTests(unittest.TestCase):
    def test_content_smaller_than_a_chunk(self):
        self.assertEqual(list(iter_chunks(b"abc", 10)), [(0, 2, b"abc")])

    def test_content_that_fills_its_chunks_exactly(self):
        self.assertEqual(list(iter_chunks(b"abcd", 2)), [(0, 1, b"ab"), (2, 3, b"cd")])

    def test_content_with_a_remainder(self):
        self.assertEqual(
            list(iter_chunks(b"abcde", 2)),
            [(0, 1, b"ab"), (2, 3, b"cd"), (4, 4, b"e")],
        )

    def test_empty_content(self):
        self.assertEqual(list(iter_chunks(b"", 2)), [])

    def test_chunk_size_must_be_positive(self):
        with self.assertRaises(ValueError):
            list(iter_chunks(b"abc", 0))


if __name__ == "__main__":
    unittest.main()
