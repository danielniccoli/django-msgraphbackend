[![Python Build](https://github.com/danielniccoli/django-msgraphbackend/actions/workflows/python-publish.yml/badge.svg)](https://github.com/danielniccoli/django-msgraphbackend/actions/workflows/python-publish.yml)

# Microsoft Graph Backend for Django

An dependency-free email backend for Django that sends emails via Microsoft Graph.

The package is a drop-in replacement for any `BaseEmailBackend` such as the default SMTP email backend `django.core.mail.backends.smtp.EmailBackend`.


## Installation

### Django

To include the *Microsoft Graph Backend for Django* in your project, add `"msgraphbackend" ` to `INSTALLED_APPS` in your `settings.py`. Then set `EMAIL_BACKEND` to `"msgraphbackend.MSGraphBackend"`. The example below shows all required changes to your `settings.py`.

```python
INSTALLED_APPS = [
    "...",
    "msgraphbackend",
]

EMAIL_BACKEND = "msgraphbackend.MSGraphBackend"

MSGRAPH_TENANT_ID = "..."
MSGRAPH_CLIENT_ID = "..."
MSGRAPH_CLIENT_SECRET = "..."
MSGRAPH_USER_ID = "..."  # Optional
MSGRAPH_USE_JSON_API = False  # Optional
MSGRAPH_MAX_SENDMAIL_SIZE = 3_500_000  # Optional
MSGRAPH_ALWAYS_USE_SENDMAIL = False  # Optional
```

The `MSGRAPH_USER_ID` is optional and not needed if you follow the instructions in section [Microsoft Entra](#microsoft-entra). The remaining optional settings select the [message format](#message-format) and tune how [large emails](#large-emails) are sent.


### Microsoft Entra

> [!IMPORTANT]
> This step describes the most permissive setup. For a more restrictive setup, see section [Modes of Operation](#modes-of-operation).

To enable the backend to connect to the Microsoft Graph API, you first need to register a Microsoft Entra App, or use an existing one. The registered app should be granted the *application permissions* `User.Read.All` and `Mail.Send` and then admin consent must be given. Finally, an app secret needs to be created. This can all be done via the Microsoft Entra Portal. For your convenience, the PowerShell script below streamlines this task.


```PowerShell
Connect-MgGraph -Scopes Application.ReadWrite.All,AppRoleAssignment.ReadWrite.All -UseDeviceAuthentication

# Get Microsoft Graph properties
$mgEnterpriseApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
$mgUserReadAll = $mgEnterpriseApp.AppRoles | ? Value -eq "User.Read.All"
$mgMailSend = $mgEnterpriseApp.AppRoles | ? Value -eq "Mail.Send"

# Register a Microsoft Entra Application
$params = @{
    DisplayName            = "Microsoft Graph Backend for Django"
    Description            = "Client Application for Microsoft Graph Backend for Django."
    RequiredResourceAccess = @{
        ResourceAppId  = $mgEnterpriseApp.AppId # Microsoft Graph
        ResourceAccess = @(
            @{ Id = $mgUserReadAll.Id; Type = "Role" }
            @{ Id = $mgMailSend.Id; Type = "Role" }
        ) 
    }
    Tags = "HideApp"
}
$registeredApp = New-MgApplication @params

# Create a Microsoft Entra Enterprise App in your tenant
$enterpriseApp = New-MgServicePrincipal -AppId $registeredApp.AppId -AppRoleAssignmentRequired

# Grant Admin Consent for User.Read.All and Mail.Send
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $enterpriseApp.Id -PrincipalId $enterpriseApp.Id -AppRoleId $mgUserReadAll.Id -ResourceId $mgEnterpriseApp.Id | Out-Null
New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $enterpriseApp.Id -PrincipalId $enterpriseApp.Id -AppRoleId $mgMailSend.Id -ResourceId $mgEnterpriseApp.Id | Out-Null

# Create a Client Secret
$params = @{
    ApplicationId      = $registeredApp.Id
    PasswordCredential = @{
        displayName = "Django"
        endDateTime = (Get-Date).AddMonths(6)
    }
}
$secret = Add-MgApplicationPassword @params

# Output a summary
Write-Host "
Your Tenant ID:     $($enterpriseApp.AppOwnerOrganizationId)
Your Client ID:     $($enterpriseApp.AppId)
Your Client Secret: $($secret.SecretText)"

```

For detailed information, refer to the article [Register an application with the Microsoft identity platform](https://learn.microsoft.com/en-us/graph/auth-register-app-v2).


## Settings
The *Microsoft Graph Backend for Django* requires the following settings:

| Setting               | Required | Description |
|-----------------------|----------|-------------|
| MSGRAPH_TENANT_ID     | Yes      | Your [Microsoft Entra tenant ID](https://learn.microsoft.com/en-us/entra/fundamentals/how-to-find-tenant). |
| MSGRAPH_CLIENT_ID     | Yes      | The [application (client) ID](https://learn.microsoft.com/en-us/graph/auth-register-app-v2) of your Microsoft Entra app. |
| MSGRAPH_CLIENT_SECRET | Yes      | The secret to your Microsoft Entra application. |
| MSGRAPH_USER_ID       | No       | If you grant your application the `User.Read.All` application permission, this setting is not required. |
| MSGRAPH_USE_JSON_API  | No       | Send emails as the Microsoft Graph-typical JSON `message` resource instead of MIME content. Defaults to `False`. See [Message Format](#message-format). |
| MSGRAPH_MAX_SENDMAIL_SIZE | No   | The size in bytes up to which an encoded email is sent in a single `sendMail` request. Defaults to `3500000`. See [Large Emails](#large-emails). |
| MSGRAPH_ALWAYS_USE_SENDMAIL | No | Send every email in a single `sendMail` request, regardless of its size. Defaults to `False`. See [Large Emails](#large-emails). |



## Modes of Operation

> [!IMPORTANT]
> The Microsoft Graph API requires that all emails are sent from a particular mailbox. The sender's *from address* must match one of the email addresses assigned to that mailbox.

There are several modes of operation that you can choose from for how the *Microsoft Graph Backend for Django* sends emails. Some modes prioritize security, while others prioritize simplicity. 


### Most Permissive

This is the mode your configured if you followed section [Microsoft Entra](#microsoft-entra). 

In this mode, the *Microsoft Graph Backend for Django* is permitted to use any mailbox and *from address* available in your Microsoft 365 tenant, limited by the configuration of your Django project. For example by setting `DEFAULT_FROM_EMAIL` or `SERVER_EMAIL`.

When an email needs to be sent, the backend requests from the Graph API the user (read: mailbox) that has the email address assigned to it, that Django wants to use as the *from address*. This request requires the *application permission* `User.Read.All`. When a user is found, the mail is sent through his mailbox to the recipient. This request requires the *application permission* `Mail.Send`.

In this mode, setting `MSGRAPH_USER_ID` is not needed.

> [!WARNING]
> *Application permissions* are very permissive. A Microsoft Entra application with the `User.Read.All` permission is allowed to retrieve the complete user profile of any non-privileged user account in your Microsoft Entra ID tenant. A Microsoft Entra application with the `Mail.Send` permission is allowed to send emails from any mailbox and existing email address in your Microsoft 365 environment.


### More Restrictive

In this mode, the *Microsoft Graph Backend for Django* is restricted to sending only from selected mailboxes in your Microsoft 365 environment. This is achieved by implementing [RBAC for Apps in Exchange Online](https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac) and limiting the users (read: mailboxes) from which the Microsoft Entra Enterprise Application is allowed to send emails. This mode requires the *application permission* `User.Read.All` and setting `MSGRAPH_USER_ID` is not needed.

> [!WARNING]
> In this mode, do not grant your Microsoft Entra registered app the `Mail.Send` permission! If you already granted the permission, you should revoke it.

After ensuring that the `Mail.Send` is revoked, you need to create the Administrative Unit.

```PowerShell
$params = @{
    displayName = "Microsoft Graph Backend for Django allowed senders"
    description = "Manages senders that Microsoft Graph Backend for Django is allowed to impersonate (read: send from)."
}
$adminUnit = New-MgDirectoryAdministrativeUnit -BodyParameter $params
```

Assign any users (read: mailboxes) to this *administrative unit* that you want to send emails from using the *Microsoft Graph Backend for Django*.

Then connect to Exchange Online and create a new service principal and link it to your *Microsoft Entra Enterprise App* via client ID and object ID. The variables refer to the script in section [Microsoft Entra](#microsoft-entra).

```PowerShell
Connect-ExchangeOnline -Device
$sp = New-ServicePrincipal -DisplayName "Microsoft Graph Backend for Django" -AppId $enterpriseApp.AppId -ObjectId $enterpriseApp.Id
New-ManagementRoleAssignment -App $sp.ObjectId -Role "Application Mail.Send" -RecipientAdministrativeUnitScope $adminUnit.Id
```

At this point your Microsoft Entra Enterprise Application should have the permission to send as any user (read: mailbox) in your *administrative unit*.

```PowerShell
Test-ServicePrincipalAuthorization -Identity $enterpriseApp.Id -Resource "<user-id>"
```


### Most Restrictive

This mode is similar to the [more restrictive](#more-restrictive) mode, but it is limited to a single mailbox. Any *from address* that you wish to use needs to be assigned to the selected mailbox.

> [!IMPORTANT]
> This mode does not require any *application permission*. If you already granted `User.Read.All` or `Mail.Send`, you should revoke them.
> Ensure that your *administrative unit* is limited to a single mailbox. That mailbox can have multiple email addresses.

This mode requires setting `MSGRAPH_USER_ID` to the user id of your selected mailbox.


## Message Format

The Microsoft Graph API accepts an email either as base64 encoded MIME content or as its own JSON [message](https://learn.microsoft.com/en-us/graph/api/resources/message) resource. By default, the *Microsoft Graph Backend for Django* sends MIME content: `django.core.mail.message.EmailMessage` builds a MIME message anyway, so nothing is lost in translation, and every header, alternative and attachment Django can express arrives as Django serialized it.

Setting `MSGRAPH_USE_JSON_API = True` switches to the JSON format. The backend then maps the `EmailMessage` onto the `message` resource, both for the single `sendMail` request and for the draft of a [large email](#large-emails). This is the format Microsoft's own [Graph SDK for Python](https://github.com/microsoftgraph/msgraph-sdk-python) uses and the one most of the Graph documentation describes. Exchange builds the MIME message itself, which sidesteps the quirks of Exchange's MIME parsing, such as its handling of line endings. The mapping covers:

| EmailMessage | Graph message |
|--------------|---------------|
| `subject` | `subject` |
| `body` | `body` with `contentType` `text`, or `html` if `content_subtype` is `"html"` |
| `alternatives` of an `EmailMultiAlternatives` | The `text/html` alternative becomes the `html` body. Other alternatives are dropped. |
| `from_email`, or a `From` header | `from` |
| `to`, `cc`, `bcc` | `toRecipients`, `ccRecipients`, `bccRecipients` |
| `reply_to`, or a `Reply-To` header | `replyTo` |
| `headers` starting with `X-` | `internetMessageHeaders` |
| `attachments`, whether attached as `(filename, content, mimetype)` or as a MIME part | `attachments` as `fileAttachment` resources; a `Content-ID` keeps an inline attachment inline |

Display names in addresses (`Name <address>`) are preserved. Because Graph holds a single body, the plain text of an `EmailMultiAlternatives` is not sent when it also has an HTML alternative; Exchange derives the text version from the HTML instead.

> [!NOTE]
> Graph only accepts custom Internet message headers whose name starts with `x-`. Any other header in `EmailMessage.extra_headers`, such as `List-Unsubscribe` or an explicit `Message-ID`, is dropped with a warning in the `msgraphbackend.message` logger, so that the email is still sent. Use the MIME format if such headers matter.


## Large Emails

The Microsoft Graph API rejects a `sendMail` request whose body grows beyond a certain size with `413 Request Entity Too Large`. Historically that limit was about 4 MB. Today, the Graph documentation states the 4 MB limit only for S/MIME payloads, and plain `sendMail` requests above 4 MB have been observed to be accepted, but Microsoft does not document a guaranteed limit. The *Microsoft Graph Backend for Django* therefore stays conservative by default: every email whose encoded size is at most `MSGRAPH_MAX_SENDMAIL_SIZE` (3,500,000 bytes) is sent in one request, and only a larger email takes the longer route described below. An email whose attachments alone cannot fit into a single request is not even encoded for it, so a large file is only encoded once, for its upload.

For such an email, the backend creates the message in the mailbox first, without its attachments. Every attachment is then added to that message on its own, either in a single request, or, if the attachment is 3 MB or larger, through an upload session that transfers it in chunks. Once all attachments are in place, the message is sent. If an attachment cannot be added, the unsent message is deleted again, so that no leftover draft remains in the mailbox. This works the same in both [message formats](#message-format).

> [!IMPORTANT]
> Because the message exists in the mailbox while it is being assembled, this route requires the *application permission* `Mail.ReadWrite`, in addition to the `Mail.Send` permission that all other emails need. In the [more restrictive](#more-restrictive) and [most restrictive](#most-restrictive) modes, the Exchange Online role assignment needs to grant `Application Mail.ReadWrite` accordingly. Without that permission, emails below the limit are unaffected and keep being sent as before.

### Sending larger emails in a single request

If your tenant accepts larger `sendMail` requests, you can raise the threshold, or bypass it, and thereby avoid both the draft route and the `Mail.ReadWrite` permission it needs:

- `MSGRAPH_MAX_SENDMAIL_SIZE` sets the size in bytes, of the base64 or JSON encoded request body, up to which an email is sent in one `sendMail` request. Set it above the default, for example to `10_000_000`, to keep emails with attachments of a few megabytes in a single request. Set it below the default to move emails to the mailbox earlier than Microsoft requires.
- `MSGRAPH_ALWAYS_USE_SENDMAIL = True` sends every email in a single `sendMail` request, however large it is. Nothing is measured and no draft is ever created. An email that Graph does reject fails with the `HTTPError` (or, with `fail_silently`, is logged and counted as not sent), exactly like any other rejected request.

```python
# settings.py
MSGRAPH_MAX_SENDMAIL_SIZE = 10_000_000  # roughly 7 MB of attachments
# or
MSGRAPH_ALWAYS_USE_SENDMAIL = True
```

> [!WARNING]
> The behaviour above the documented 4 MB is not guaranteed by Microsoft and may differ between tenants or change over time. Verify it against your tenant before relying on it, and keep the default if emails must never fail on their size alone.

The remaining limits of the draft route are class attributes rather than settings, because they follow from what the Graph API documents. A subclass can still adjust them.

| Attribute                  | Default | Description |
|----------------------------|---------|-------------|
| MAX_SENDMAIL_SIZE          | 3500000 | The default of `MSGRAPH_MAX_SENDMAIL_SIZE`. |
| MAX_INLINE_ATTACHMENT_SIZE | 3145728 | The size in bytes from which an attachment is uploaded in chunks instead of in a single request. |
| UPLOAD_CHUNK_SIZE          | 4194304 | The size in bytes of a single chunk of an upload session. |

```python
from msgraphbackend import MSGraphBackend


class SmallerMailBackend(MSGraphBackend):
    MAX_SENDMAIL_SIZE = 1_000_000
```


## Notes

The *Microsoft Graph Backend for Django* sends email by default not in the Microsoft Graph-typical JSON, but in the MIME format. This is due to how `django.core.mail.message.EmailMessage` internally works. Its `message()` method returns the email in MIME format, so passing it on unchanged cannot lose anything. The Microsoft's own [Graph SDK for Python](https://github.com/microsoftgraph/msgraph-sdk-python) does not support sending emails in MIME format. The JSON format is available as an opt-in through `MSGRAPH_USE_JSON_API`, see [Message Format](#message-format).