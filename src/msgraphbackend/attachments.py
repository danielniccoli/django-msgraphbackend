from __future__ import annotations

import base64
from dataclasses import dataclass
from email.mime.base import MIMEBase
from typing import TYPE_CHECKING

from django.conf import settings

if TYPE_CHECKING:
    from collections.abc import Iterator

    from django.core.mail.message import EmailMessage


# Graph requires every attachment to have a name and a content type. Django does
# not, so these stand in when the message does not provide them.
DEFAULT_ATTACHMENT_NAME = "attachment"
DEFAULT_ATTACHMENT_MIME_TYPE = "application/octet-stream"


@dataclass(frozen=True)
class GraphAttachment:
    """A single attachment, in the shape the Graph attachment APIs expect."""

    name: str
    content: bytes
    mimetype: str
    content_id: str | None = None
    is_inline: bool = False

    @property
    def size(self) -> int:
        """The size in bytes of the attached file itself, not of the request."""
        return len(self.content)

    def as_file_attachment(self) -> dict:
        """Returns the fileAttachment resource that is posted directly."""
        file_attachment = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": self.name,
            "contentType": self.mimetype,
            "contentBytes": base64.b64encode(self.content).decode("ascii"),
        }
        return file_attachment | self._inline_properties()

    def as_attachment_item(self) -> dict:
        """Returns the AttachmentItem resource that opens an upload session."""
        attachment_item = {
            "attachmentType": "file",
            "name": self.name,
            "size": self.size,
            "contentType": self.mimetype,
        }
        return attachment_item | self._inline_properties()

    def _inline_properties(self) -> dict:
        """Returns the properties that keep an inline attachment inline."""
        properties: dict = {}
        if self.content_id:
            properties["contentId"] = self.content_id
        if self.is_inline:
            properties["isInline"] = True
        return properties


def iter_attachments(email_message: EmailMessage) -> Iterator[GraphAttachment]:
    """
    Yields the attachments of an EmailMessage as GraphAttachment instances.

    Django stores an attachment either as a (filename, content, mimetype) tuple
    or, when it was attached as a MIME part, as a MIMEBase instance. Both are
    normalized here, so the callers only deal with a name and raw bytes.
    """
    encoding = email_message.encoding or settings.DEFAULT_CHARSET
    for attachment in email_message.attachments:
        if isinstance(attachment, MIMEBase):
            yield _from_mime_part(attachment)
        else:
            filename, content, mimetype = attachment
            yield _from_triple(filename, content, mimetype, encoding)


def attachments_size(email_message: EmailMessage) -> int:
    """
    Returns a lower bound of the bytes the attachments take up in the MIME message.

    Django transfer-encodes an attachment when it serializes the message, which
    never makes it smaller, so the size of what was attached is a bound that
    costs nothing to compute: nothing is encoded or decoded for it. An
    attachment whose size is unknown, such as an attached email object, counts
    as zero.
    """
    total = 0
    for attachment in email_message.attachments:
        if isinstance(attachment, MIMEBase):
            # The payload of a MIME part is already transfer-encoded.
            content = attachment.get_payload()
        else:
            _, content, _ = attachment
        if isinstance(content, (str, bytes, bytearray)):
            total += len(content)
    return total


def iter_chunks(content: bytes, chunk_size: int) -> Iterator[tuple[int, int, bytes]]:
    """
    Splits the content into the byte ranges of an upload session.

    Yields (first_byte, last_byte, chunk) triples, where the two indexes are the
    inclusive range that the Content-Range header of the chunk has to report.
    """
    if chunk_size <= 0:
        raise ValueError("The chunk size must be a positive number of bytes.")
    for first_byte in range(0, len(content), chunk_size):
        chunk = content[first_byte : first_byte + chunk_size]
        yield first_byte, first_byte + len(chunk) - 1, chunk


def _from_triple(
    filename: str | None, content, mimetype: str | None, encoding: str
) -> GraphAttachment:
    """Normalizes one of Django's (filename, content, mimetype) attachments."""
    mimetype = mimetype or DEFAULT_ATTACHMENT_MIME_TYPE
    if isinstance(content, str):
        # Django keeps the content of a text attachment as a string. Graph wants
        # the bytes of the file, so the charset has to be declared explicitly.
        content = content.encode(encoding)
        if mimetype.startswith("text/") and "charset=" not in mimetype:
            mimetype = f'{mimetype}; charset="{encoding}"'
    return GraphAttachment(
        name=filename or DEFAULT_ATTACHMENT_NAME,
        content=_as_bytes(content),
        mimetype=mimetype,
    )


def _from_mime_part(part: MIMEBase) -> GraphAttachment:
    """Normalizes an attachment that was attached as a MIME part."""
    content = part.get_payload(decode=True)
    if not isinstance(content, bytes):
        # A multipart part, such as an attached message/rfc822, has no decodable
        # payload. Attaching its serialization keeps the file intact.
        content = part.as_bytes()
    content_id = part.get("Content-ID")
    if content_id:
        content_id = str(content_id).strip().strip("<>")
    return GraphAttachment(
        name=part.get_filename() or DEFAULT_ATTACHMENT_NAME,
        content=content,
        mimetype=part.get_content_type(),
        content_id=content_id or None,
        # An inline part often only identifies itself by its Content-ID, which
        # the body references as cid:<content-id>.
        is_inline=part.get_content_disposition() == "inline" or bool(content_id),
    )


def _as_bytes(content) -> bytes:
    """Coerces an attachment payload to bytes."""
    if isinstance(content, (bytes, bytearray)):
        return bytes(content)
    # A message/rfc822 attachment may hold an email object instead of its bytes.
    if hasattr(content, "message"):
        content = content.message()
    if hasattr(content, "as_bytes"):
        return content.as_bytes()
    return bytes(content)
