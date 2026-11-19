# Office365Wrapper

A minimal Electron desktop shell for [Microsoft 365](https://www.microsoft365.com/)
and Outlook on the web.

## Requirements

- Node.js 22.12 or later and npm
- Windows, macOS, or Linux with a graphical desktop and Electron's platform dependencies
- An internet connection and a Microsoft account or eligible organizational account

## Run

```sh
npm install
npm start
```

On filesystems that do not support symlinks, use `npm install --no-bin-links`.
Keep Electron updated to receive Chromium security fixes.

The dashboard opens by default and provides shortcuts to Microsoft 365,
Outlook, Word, Excel, PowerPoint, OneNote, Copilot, and To Do on the web. These
web apps require internet access and may require sign-in or an eligible license.
Use **Navigate > Dashboard** to return to it.

Use the **Navigate** menu to switch between Microsoft 365, Outlook for work,
school, or personal accounts. Switching pages may discard unsaved changes.

To open an app directly instead:

```sh
# Microsoft 365
npm start -- --microsoft365

# Work or school account
npm start -- --outlook

# Personal Outlook.com, Hotmail, or Live account
npm start -- --outlook-personal
```

Mail and calendar are provided by the Outlook website; the shell does not
provide native mail synchronization.

## Graph search setup

The dashboard can search mail, calendar events, and accessible OneDrive and
SharePoint files using Microsoft Graph. This search API supports work or school
accounts in the global Microsoft cloud, not personal Microsoft accounts.

1. [Register an app in Microsoft Entra ID](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app) for organizational accounts.
2. Under **Authentication**, add the **Mobile and desktop applications**
   platform with redirect URI `http://localhost`, and enable public client flows.
3. Under **API permissions**, add delegated Microsoft Graph permissions
   `Mail.Read`, `Calendars.Read`, and `Files.Read.All`. Your organization may
   require administrator consent.
4. Set `OFFICE365_CLIENT_ID` to the app's **Application (client) ID** before
   starting the app (no client secret):

   ```sh
   export OFFICE365_CLIENT_ID="your-client-id"
   npm start
   ```

   In PowerShell, use `$env:OFFICE365_CLIENT_ID = "your-client-id"` instead of
   `export`.

Enter a query in the dashboard and sign in through the system browser when
prompted. Graph sign-in is separate from website sign-in. Search returns up to
25 results per category, with errors shown per category; some results may not
include a direct link. Microsoft Search indexes the user's own mailbox and
primary calendar, not delegated mailboxes or To Do tasks. The app keeps Graph
tokens in memory only. **Sign out of Graph** clears that in-memory account but
does not sign out the browser or Microsoft websites. If the sign-in browser
window is closed before completing authentication, restart the app to cancel
the pending sign-in.

## Sign in

Select **Sign in** on the Microsoft 365 or Outlook page and follow
Microsoft's prompts, including any required multifactor authentication. Sign
out or switch accounts using the website's account menu.

Electron stores its own persistent browser profile, including sign-in cookies,
in its user-data directory (for example, `%APPDATA%\office365-wrapper` on
Windows). It does not import the previous WebView2 profile, so sign in again
after upgrading. The shell does not collect passwords. Organizational policies
may restrict sign-in from embedded browsers.

## Notifications

Select **Notifications > Enable Microsoft website notifications for this
session**, then use **Navigate > Reload** and enable notifications in Outlook
if prompted. Notifications require OS notification permission and depend on
the website. Permission is limited to Microsoft 365 and Outlook origins and
resets when the app exits. Reloading may discard unsaved
changes. Background delivery while a different page is open is not guaranteed;
closing the window exits the app.

On Windows, when running from source, pin
`node_modules\electron\dist\electron.exe` to the Start menu for notifications.
Packaged apps need a Start menu shortcut with the appropriate AppUserModelID.

New-window requests and other website permissions are blocked; workflows
requiring popups, camera, microphone, or other permissions are not supported.
Closing the window exits the app. Electron bundles Chromium, so the WebView2
Runtime is not required.
