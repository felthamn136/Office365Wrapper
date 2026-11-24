# Office365Wrapper

![Office365Wrapper: Microsoft 365 on Linux](./assets/readme-banner.png)

A Linux desktop shell for [Microsoft 365](https://www.microsoft365.com/) and
Outlook on the web, built with Rust, Tauri, and WebKitGTK.

## Requirements

- Linux with a graphical desktop, Rust and Cargo, and WebKitGTK 4.1 development
  libraries.
- Internet access and a Microsoft account or eligible organizational account.

On Debian or Ubuntu, install the build dependencies:

```sh
sudo apt install build-essential libwebkit2gtk-4.1-dev \
  libayatana-appindicator3-dev librsvg2-dev libxdo-dev libssl-dev
```

## Build and run

WebKitGTK uses the desktop's native Wayland or X11 backend. Chromium options
such as `--ozone-platform` and `--disable-vulkan` do not apply.

```sh
cargo build --release --manifest-path src-tauri/Cargo.toml
./src-tauri/target/release/office-dashboard
```

This builds a native executable using the system WebKitGTK libraries. If the
project filesystem cannot support Cargo's build files, set `CARGO_TARGET_DIR`
to a directory on a local filesystem before building.

### AppImage

A prebuilt x86-64 [Office365Wrapper AppImage](./dist/Office365Wrapper_0.1.0_amd64.AppImage)
is available in `dist`. Re-integrate the AppImage if replacing an older version.

To build an AppImage, install `patchelf` and the GStreamer plugins needed for
media playback (for example, `gstreamer1.0-plugins-base` and
`gstreamer1.0-plugins-good` on Debian or Ubuntu). Then install the Tauri CLI and
build the bundle:

```sh
cargo install tauri-cli --version "^2" --locked
cargo tauri build --config packaging/linux/tauri.conf.json --bundles appimage
```

The bundle is created under
`src-tauri/target/release/bundle/appimage/` (or under `CARGO_TARGET_DIR`, if
set). For example:

```sh
chmod +x "src-tauri/target/release/bundle/appimage/Office365Wrapper_0.1.0_amd64.AppImage"
"src-tauri/target/release/bundle/appimage/Office365Wrapper_0.1.0_amd64.AppImage"
```

The bundle includes the application icon, WebKitGTK dependencies, and available
GStreamer media plugins. It may require FUSE 2; use
`--appimage-extract-and-run` if FUSE is unavailable. Desktop-menu integration
depends on the desktop environment or an AppImage integration tool.
New bundles use **Office365Wrapper** as the desktop application name.

When using Gear Lever, open a copy of the AppImage from outside its AppImages
folder before integrating it; Gear Lever may move the selected file. Build on
the oldest Linux distribution you intend to support to avoid glibc
compatibility issues, and rebuild periodically to pick up WebKitGTK security
updates.

## Using the dashboard

The Office Dashboard opens by default with shortcuts to Microsoft 365, Outlook,
OneDrive, Word, Excel, PowerPoint, OneNote, Copilot, and To Do. The websites may
require sign-in or an eligible license.

Use the **Navigate** menu to switch between the dashboard and Copilot.
Outlook for work, school, or personal accounts opens in a separate app window.
Outlook's WebKitGTK window identifies itself as Safari for compatibility;
Microsoft does not officially support WebKitGTK on Linux, so sign-in or some
features may still fail. If needed, use the **Navigate** menu's Outlook
**in browser** options to open Outlook in your default browser instead.
Switching embedded pages may discard unsaved changes.

You can also choose a starting destination from the command line:

```sh
# Open Microsoft 365 in the app
./src-tauri/target/release/office-dashboard --microsoft365

# Open Outlook in a separate app window
./src-tauri/target/release/office-dashboard --outlook
./src-tauri/target/release/office-dashboard --outlook-personal
```

Mail and calendar are provided by the Outlook website; the shell does not
provide native mail synchronization or dashboard-wide search. Use the search
features in the Microsoft websites.

## Sign-in and browser data

Select **Sign in** on the Microsoft 365 or Outlook page, then follow Microsoft's
prompts, including any required multifactor authentication. Sign out or
switch accounts using the website's account menu. The shell does not collect
passwords.

The app stores its WebKit profile and sign-in cookies for embedded Microsoft
pages in `~/.local/share/com.office365wrapper.dashboard/webkit` (or under
`$XDG_DATA_HOME`). Cookies are not imported from previous versions, so sign in
again after switching. Outlook opened in the default browser uses that
browser's separate profile and cookies.

## Notifications and links

To enable notifications, select **Notifications > Enable Microsoft website
notifications for this session**, then use **Navigate > Reload** and enable
notifications on the Microsoft 365 or Outlook page if prompted. Reload applies
to the focused app window. Notifications require OS permission and depend on
the website. Permission is enabled only on supported Microsoft 365 and Outlook
pages and resets when the app exits. WebKitGTK does not expose the requesting
frame's origin, so embedded frames on that page may also receive permission.

Microsoft pages can open HTTPS links (including app-launcher links) and
`about:blank` popups in separate WebKit windows. Other sites cannot open new
windows, and insecure navigation is blocked. Popups close with their opener.
Camera, microphone, and other website permissions remain blocked. Reloading
may discard unsaved changes, background notification delivery is not
guaranteed, and closing all windows exits the app.
