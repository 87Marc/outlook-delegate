# Delegate Setup Guide

This guide walks through setting up your AI assistant as a delegate for the owner's Microsoft 365 mailbox and calendar.

## Overview

**Architecture:**
- Your AI assistant has its own M365 account: `assistant@yourdomain.com`
- The owner's account: `owner@yourdomain.com`
- The assistant authenticates as itself, but accesses the owner's resources

**What users see:**
- Emails sent by the assistant appear as: "Assistant on behalf of Owner"
- Calendar events show the owner as the organizer

## Step 1: Create the Assistant's M365 Account

If not already done, create a user account for the assistant in your Microsoft 365 admin center.

## Step 2: Grant Delegate Permissions

The owner (or an admin) must grant the assistant access to the mailbox and calendar.

### Option A: PowerShell (Recommended for full control)

Connect to Exchange Online:
```powershell
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline -UserPrincipalName admin@yourdomain.com
```

Grant permissions:
```powershell
# Full mailbox access (read, modify, organize)
Add-MailboxPermission -Identity "owner@yourdomain.com" `
    -User "assistant@yourdomain.com" `
    -AccessRights FullAccess `
    -InheritanceType All

# Send-on-behalf permission (emails show "on behalf of")
Set-Mailbox -Identity "owner@yourdomain.com" `
    -GrantSendOnBehalfTo "assistant@yourdomain.com"

# Calendar access with delegate flag
Add-MailboxFolderPermission -Identity "owner@yourdomain.com:\Calendar" `
    -User "assistant@yourdomain.com" `
    -AccessRights Editor `
    -SharingPermissionFlags Delegate
```

Verify:
```powershell
Get-MailboxPermission -Identity "owner@yourdomain.com" | Where-Object {$_.User -like "*assistant*"}
Get-Mailbox "owner@yourdomain.com" | Select-Object GrantSendOnBehalfTo
Get-MailboxFolderPermission -Identity "owner@yourdomain.com:\Calendar"
```

### Option B: Outlook Client (Owner does this)

1. Open Outlook → File → Account Settings → Delegate Access
2. Click "Add" and select the assistant account
3. Set permissions:
   - Calendar: Editor
   - Inbox: Editor (or Reviewer for read-only)
   - Check "Delegate can see my private items" if needed
4. Click OK

### Option C: Microsoft 365 Admin Center

1. Go to admin.microsoft.com
2. Users → Active Users → Owner → Mail → Manage mailbox delegation
3. Add the assistant to "Send on behalf" and "Read and manage"

## Step 3: Azure AD App Registration

Register an app that the assistant will use to authenticate.

### Create the App

1. Go to portal.azure.com → Azure Active Directory → App registrations
2. New registration:
   - Name: "AI Assistant Mail Access"
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: `http://localhost:8400/callback`
3. Note the **Application (client) ID**

### Configure Permissions

In your app → API permissions → Add a permission → Microsoft Graph → Delegated permissions:

- `Mail.ReadWrite.Shared` - Read/write shared mailboxes
- `Mail.Send.Shared` - Send on behalf
- `Calendars.ReadWrite.Shared` - Shared calendar access
- `User.Read` - Read own profile
- `offline_access` - Refresh tokens

Click "Grant admin consent" (requires admin).

### Create Client Secret

1. Certificates & secrets → New client secret
2. Description: "AI Assistant"
3. Expiration: Choose appropriate duration
4. Copy the **Value** immediately (shown only once)

## Step 4: Configure the Skill

Create the config directory and files:

```bash
mkdir -p ~/.outlook-mcp
```

Create `~/.outlook-mcp/config.json`:
```json
{
  "client_id": "YOUR-APP-CLIENT-ID",
  "client_secret": "YOUR-CLIENT-SECRET",
  "owner_email": "owner@yourdomain.com",
  "delegate_email": "assistant@yourdomain.com",
  "timezone": "UTC"
}
```

## Step 5: Authorize the Assistant

Run the OAuth flow to get tokens. The assistant signs in as itself:

```bash
# Generate auth URL
CLIENT_ID=$(jq -r '.client_id' ~/.outlook-mcp/config.json)
REDIRECT="http://localhost:8400/callback"
SCOPE="offline_access%20User.Read%20Mail.ReadWrite.Shared%20Mail.Send.Shared%20Calendars.ReadWrite.Shared"

echo "Open this URL in a browser and sign in as the assistant:"
echo "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=$CLIENT_ID&response_type=code&redirect_uri=$REDIRECT&scope=$SCOPE"
```

After signing in, you'll be redirected to localhost with a `code` parameter. Exchange it:

```bash
CODE="paste-the-code-here"
CLIENT_SECRET=$(jq -r '.client_secret' ~/.outlook-mcp/config.json)

curl -s -X POST "https://login.microsoftonline.com/common/oauth2/v2.0/token" \
  -d "client_id=$CLIENT_ID" \
  -d "client_secret=$CLIENT_SECRET" \
  -d "code=$CODE" \
  -d "redirect_uri=$REDIRECT" \
  -d "grant_type=authorization_code" \
  -d "scope=$SCOPE" > ~/.outlook-mcp/credentials.json

cat ~/.outlook-mcp/credentials.json | jq '{status: "authorized", expires_in}'
```

## Step 6: Test Access

```bash
./scripts/outlook-token.sh test
```

Expected output:
```
Testing delegate access...

1. Delegate identity (who is authenticated):
{
  "authenticated_as": "assistant@yourdomain.com",
  "display_name": "AI Assistant"
}

2. Owner mailbox access (owner@yourdomain.com):
{
  "status": "OK",
  "folder": "Inbox",
  "unread": 15,
  "total": 1234
}

3. Owner calendar access (owner@yourdomain.com):
{
  "status": "OK",
  "calendar": "Calendar",
  "canEdit": true
}
```

## Troubleshooting

### "Access is denied" or "ErrorAccessDenied"

The delegate permissions aren't set correctly.

1. Verify mailbox permission:
   ```powershell
   Get-MailboxPermission -Identity "owner@yourdomain.com" -User "assistant@yourdomain.com"
   ```

2. Check it shows `FullAccess` or at least `ReadPermission`

### "The specified object was not found"

The owner email is wrong, or the mailbox doesn't exist.

1. Verify the email in config.json
2. Check the account exists in M365 admin

### "Insufficient privileges to complete the operation"

The Azure AD app is missing `.Shared` permissions.

1. Go to App registrations → Your app → API permissions
2. Ensure you have `Mail.ReadWrite.Shared` (not just `Mail.ReadWrite`)
3. Click "Grant admin consent"

### Token Refresh Fails

The refresh token may have expired or been revoked.

1. Check if the assistant's password changed (invalidates tokens)
2. Re-run the authorization flow (Step 5)

### Emails Not Showing "On Behalf Of"

The SendOnBehalf permission is missing.

```powershell
Set-Mailbox -Identity "owner@yourdomain.com" -GrantSendOnBehalfTo "assistant@yourdomain.com"
```

## Security Considerations

1. **Audit Logging**: All actions are logged in the owner's mailbox audit
2. **Principle of Least Privilege**: Consider Reviewer instead of Editor if write access isn't needed
3. **Token Security**: Protect `~/.outlook-mcp/credentials.json`
4. **Regular Review**: Periodically review delegate access

## Permission Levels Reference

| Level | Mail | Calendar |
|-------|------|----------|
| Reviewer | Read only | Read only |
| Author | Read, create | Read, create |
| Editor | Read, create, modify, delete | Full control |
| FullAccess | Everything | - |

## Revoking Access

To remove the assistant's access:

```powershell
Remove-MailboxPermission -Identity "owner@yourdomain.com" -User "assistant@yourdomain.com" -AccessRights FullAccess
Set-Mailbox -Identity "owner@yourdomain.com" -GrantSendOnBehalfTo $null
Remove-MailboxFolderPermission -Identity "owner@yourdomain.com:\Calendar" -User "assistant@yourdomain.com"
```

Or via Outlook: File → Account Settings → Delegate Access → Remove
