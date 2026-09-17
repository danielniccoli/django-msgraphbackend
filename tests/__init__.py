"""Tests for the Microsoft Graph Backend for Django."""

from __future__ import annotations

import django
from django.conf import settings

if not settings.configured:
    settings.configure(
        DEFAULT_CHARSET="utf-8",
        EMAIL_USE_LOCALTIME=False,
        INSTALLED_APPS=["msgraphbackend"],
        USE_TZ=True,
    )
    django.setup()
