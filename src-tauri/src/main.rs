use std::{
    os::unix::fs::DirBuilderExt,
    process::Command,
    sync::atomic::{AtomicBool, Ordering},
};
use tauri::{
    menu::{CheckMenuItemBuilder, MenuBuilder, SubmenuBuilder},
    webview::{NewWindowFeatures, NewWindowResponse},
    AppHandle, Manager, WebviewUrl, WebviewWindow, WebviewWindowBuilder, WindowEvent,
};
use url::Url;
use webkit2gtk::{
    glib::prelude::*, NotificationPermissionRequest, PermissionRequestExt, WebViewExt,
};

struct AppState {
    notifications: AtomicBool,
}

const OUTLOOK_USER_AGENT: &str = concat!(
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ",
    "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Safari/605.1.15"
);

fn secure_url(url: &Url) -> bool {
    url.scheme() == "https" && url.username().is_empty() && url.password().is_none()
}

fn outlook_browser_url(url: &Url) -> Option<&'static str> {
    if !secure_url(url) || url.port().is_some() {
        return None;
    }
    match url.host_str()? {
        "outlook.office365.com" | "outlook.office.com" | "outlook.cloud.microsoft" => {
            Some("https://outlook.office.com/")
        }
        "outlook.live.com" => Some("https://outlook.live.com/mail/"),
        _ => None,
    }
}

fn open_outlook_in_browser(url: &Url) -> std::io::Result<()> {
    let browser_url = outlook_browser_url(url).expect("validated Outlook URL");
    Command::new("xdg-open").arg(browser_url).spawn()?;
    Ok(())
}

fn microsoft_url(url: &Url) -> bool {
    const DOMAINS: &[&str] = &[
        "microsoft.com",
        "microsoft365.com",
        "office.com",
        "office365.com",
        "cloud.microsoft",
        "live.com",
        "microsoftonline.com",
        "onedrive.com",
        "sharepoint.com",
    ];
    secure_url(url)
        && url.port().is_none()
        && url.host_str().is_some_and(|host| {
            DOMAINS
                .iter()
                .any(|domain| host == *domain || host.ends_with(&format!(".{domain}")))
        })
}

fn notification_origin(url: &Url) -> bool {
    secure_url(url)
        && url.port().is_none()
        && matches!(
            url.host_str(),
            Some(
                "www.microsoft365.com"
                    | "m365.cloud.microsoft"
                    | "outlook.office.com"
                    | "outlook.office365.com"
                    | "outlook.cloud.microsoft"
                    | "outlook.live.com"
            )
        )
}

fn dashboard_url() -> Url {
    Url::parse("tauri://localhost/dashboard.html").expect("valid dashboard URL")
}

fn allow_navigation(url: &Url) -> bool {
    secure_url(url) || *url == dashboard_url() || url.as_str() == "about:blank"
}

fn route_to_outlook_window(url: &Url, outlook_window: bool) -> bool {
    !outlook_window && outlook_browser_url(url).is_some()
}

fn open_outlook_window(app: &AppHandle, url: Url) -> tauri::Result<()> {
    if let Some(window) = app.get_webview_window("outlook") {
        window.navigate(url)?;
        window.show()?;
        window.set_focus()?;
    } else {
        create_window(app, "outlook", WebviewUrl::External(url), None, true)?;
    }
    Ok(())
}

fn install_permissions(window: &WebviewWindow) -> tauri::Result<()> {
    let app = window.app_handle().clone();
    window.with_webview(move |platform| {
        let display_app = app.clone();
        platform
            .inner()
            .connect_show_notification(move |view, _notification| {
                let enabled = display_app
                    .state::<AppState>()
                    .notifications
                    .load(Ordering::Relaxed);
                let trusted = view
                    .uri()
                    .and_then(|uri| Url::parse(&uri).ok())
                    .is_some_and(|url| notification_origin(&url));
                !(enabled && trusted)
            });
        platform
            .inner()
            .connect_permission_request(move |view, request| {
                let enabled = app
                    .state::<AppState>()
                    .notifications
                    .load(Ordering::Relaxed);
                let allowed = enabled
                    && request.is::<NotificationPermissionRequest>()
                    && view
                        .uri()
                        .and_then(|uri| Url::parse(&uri).ok())
                        .is_some_and(|url| notification_origin(&url));
                if allowed {
                    request.allow();
                } else {
                    request.deny();
                }
                true
            });
    })
}

fn close_window_with_webview(window: &WebviewWindow) -> tauri::Result<()> {
    let app = window.app_handle().clone();
    let label = window.label().to_owned();
    window.with_webview(move |platform| {
        // Wry destroys the WebKit widget on window.close(), not its native window.
        platform.inner().connect_local("destroy", false, move |_| {
            let app = app.clone();
            let label = label.clone();
            webkit2gtk::glib::idle_add_local_once(move || {
                if let Some(window) = app.get_webview_window(&label) {
                    if let Err(error) = window.close() {
                        eprintln!("Could not close popup window: {error}");
                    }
                }
            });
            None
        });
    })
}

fn create_window(
    app: &AppHandle,
    label: &str,
    initial: WebviewUrl,
    features: Option<NewWindowFeatures>,
    outlook_window: bool,
) -> tauri::Result<WebviewWindow> {
    let popup = features.is_some();
    let profile = app.path().app_data_dir()?.join("webkit");
    std::fs::DirBuilder::new()
        .recursive(true)
        .mode(0o700)
        .create(&profile)?;
    let handle = app.clone();
    let navigation_handle = app.clone();
    let navigation_label = label.to_owned();
    let parent_label = label.to_owned();
    let builder = WebviewWindowBuilder::new(app, label, initial)
        .title(if outlook_window {
            "Office365Wrapper - Outlook"
        } else {
            "Office365Wrapper"
        })
        .inner_size(1200.0, 800.0)
        .data_directory(profile)
        .on_navigation(move |url| {
            if route_to_outlook_window(url, outlook_window) {
                let app = navigation_handle.clone();
                let url = url.clone();
                let label = navigation_label.clone();
                if let Err(error) = navigation_handle.run_on_main_thread(move || {
                    if let Err(error) = open_outlook_window(&app, url) {
                        eprintln!("Could not open Outlook in the app: {error}");
                    } else if popup {
                        webkit2gtk::glib::idle_add_local_once(move || {
                            if let Some(window) = app.get_webview_window(&label) {
                                match window.url() {
                                    Ok(url) if url.as_str() == "about:blank" => {
                                        if let Err(error) = window.close() {
                                            eprintln!("Could not close blank popup: {error}");
                                        }
                                    }
                                    Ok(_) => {}
                                    Err(error) => eprintln!("Could not check popup URL: {error}"),
                                }
                            }
                        });
                    }
                }) {
                    eprintln!("Could not schedule Outlook navigation: {error}");
                }
                false
            } else {
                allow_navigation(url)
            }
        })
        .on_new_window(move |url, features| {
            let opener = features
                .opener()
                .webview
                .uri()
                .and_then(|uri| Url::parse(&uri).ok());
            if !opener.as_ref().is_some_and(microsoft_url)
                || !(secure_url(&url) || url.as_str() == "about:blank")
            {
                return NewWindowResponse::Deny;
            }
            let child_label = format!("popup-{}", uuid::Uuid::new_v4());
            match create_window(
                &handle,
                &child_label,
                WebviewUrl::External("about:blank".parse().expect("valid URL")),
                Some(features),
                outlook_window,
            ) {
                Ok(child) => {
                    if let Some(parent) = handle.get_webview_window(&parent_label) {
                        let app = handle.clone();
                        let label = child_label.clone();
                        parent.on_window_event(move |event| {
                            if matches!(event, WindowEvent::Destroyed) {
                                if let Some(popup) = app.get_webview_window(&label) {
                                    if let Err(error) = popup.close() {
                                        eprintln!("Could not close popup: {error}");
                                    }
                                }
                            }
                        });
                    }
                    NewWindowResponse::Create { window: child }
                }
                Err(error) => {
                    eprintln!("Could not open Microsoft app window: {error}");
                    NewWindowResponse::Deny
                }
            }
        });
    let builder = if outlook_window {
        builder.user_agent(OUTLOOK_USER_AGENT)
    } else {
        builder
    };
    let builder = if let Some(features) = features {
        builder.window_features(features)
    } else {
        builder
    };
    let window = builder.build()?;
    install_permissions(&window)?;
    if popup {
        close_window_with_webview(&window)?;
    }
    Ok(window)
}

fn initial_url() -> Result<WebviewUrl, String> {
    let args: Vec<_> = std::env::args().skip(1).collect();
    if args.len() > 1 {
        return Err("Only one destination can be specified.".into());
    }
    match args.first().map(String::as_str) {
        None => Ok(WebviewUrl::App("dashboard.html".into())),
        Some("--microsoft365") => Ok(WebviewUrl::External(
            "https://www.microsoft365.com/".parse().unwrap(),
        )),
        Some("--outlook") => Ok(WebviewUrl::External(
            "https://outlook.office365.com/".parse().unwrap(),
        )),
        Some("--outlook-personal") => Ok(WebviewUrl::External(
            "https://outlook.live.com/mail/".parse().unwrap(),
        )),
        Some(arg) => Err(format!("Unknown argument: {arg}")),
    }
}

fn run() -> Result<(), Box<dyn std::error::Error>> {
    let (initial, outlook) = match initial_url().map_err(std::io::Error::other)? {
        WebviewUrl::External(url) if outlook_browser_url(&url).is_some() => {
            (WebviewUrl::App("dashboard.html".into()), Some(url))
        }
        initial => (initial, None),
    };
    tauri::Builder::default()
        .manage(AppState {
            notifications: AtomicBool::new(false),
        })
        .setup(move |app| {
            let handle = app.handle();
            let navigate = SubmenuBuilder::new(handle, "Navigate")
                .text("dashboard", "Office Dashboard")
                .text("copilot", "Copilot")
                .text("outlook", "Outlook (work or school)")
                .text("personal", "Outlook (personal)")
                .separator()
                .text("outlook-browser", "Outlook (work or school) in browser")
                .text("personal-browser", "Outlook (personal) in browser")
                .separator()
                .text("reload", "Reload")
                .build()?;
            let toggle = CheckMenuItemBuilder::with_id(
                "notifications",
                "Enable Microsoft website notifications for this session",
            )
            .checked(false)
            .build(handle)?;
            let notifications = SubmenuBuilder::new(handle, "Notifications")
                .item(&toggle)
                .build()?;
            let edit = SubmenuBuilder::new(handle, "Edit")
                .cut()
                .copy()
                .paste()
                .select_all()
                .build()?;
            app.set_menu(
                MenuBuilder::new(handle)
                    .items(&[&navigate, &notifications, &edit])
                    .build()?,
            )?;
            create_window(handle, "main", initial, None, false)?;
            if let Some(url) = outlook {
                open_outlook_window(handle, url)?;
            }
            Ok(())
        })
        .on_menu_event(|app, event| {
            if event.id().as_ref() == "notifications" {
                let state = app.state::<AppState>();
                state.notifications.fetch_xor(true, Ordering::Relaxed);
                return;
            }
            let outlook_url = match event.id().as_ref() {
                "outlook" | "outlook-browser" => Some("https://outlook.office365.com/"),
                "personal" | "personal-browser" => Some("https://outlook.live.com/mail/"),
                _ => None,
            };
            if let Some(url) = outlook_url {
                let url = url.parse().expect("valid Outlook URL");
                if matches!(event.id().as_ref(), "outlook-browser" | "personal-browser") {
                    if let Err(error) = open_outlook_in_browser(&url) {
                        eprintln!("Could not open Outlook in the default browser: {error}");
                    }
                } else if let Err(error) = open_outlook_window(app, url) {
                    eprintln!("Could not open Outlook in the app: {error}");
                }
                return;
            }
            if event.id().as_ref() == "reload" {
                for window in app.webview_windows().into_values() {
                    match window.is_focused() {
                        Ok(true) => {
                            if let Err(error) = window.reload() {
                                eprintln!("Could not reload page: {error}");
                            }
                            break;
                        }
                        Ok(false) => {}
                        Err(error) => eprintln!("Could not determine focused window: {error}"),
                    }
                }
                return;
            }
            let url = match event.id().as_ref() {
                "dashboard" => Some(dashboard_url()),
                "copilot" => Some("https://copilot.cloud.microsoft/".parse().unwrap()),
                _ => None,
            };
            if let Some(url) = url {
                if let Some(window) = app.get_webview_window("main") {
                    if let Err(error) = window.navigate(url) {
                        eprintln!("Could not navigate: {error}");
                    } else if let Err(error) = window.set_focus() {
                        eprintln!("Could not focus dashboard window: {error}");
                    }
                } else {
                    let initial = if url == dashboard_url() {
                        WebviewUrl::App("dashboard.html".into())
                    } else {
                        WebviewUrl::External(url)
                    };
                    if let Err(error) = create_window(app, "main", initial, None, false) {
                        eprintln!("Could not reopen dashboard window: {error}");
                    }
                }
            }
        })
        .run(tauri::generate_context!())?;
    Ok(())
}

fn main() {
    if let Err(error) = run() {
        eprintln!("Office Dashboard error: {error}");
        std::process::exit(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn restricts_navigation_and_popup_origins() {
        assert!(allow_navigation(&dashboard_url()));
        assert!(allow_navigation(
            &Url::parse("https://tenant.sharepoint.com/").unwrap()
        ));
        assert!(!allow_navigation(
            &Url::parse("file:///etc/passwd").unwrap()
        ));
        assert!(microsoft_url(
            &Url::parse("https://m365.cloud.microsoft/").unwrap()
        ));
        assert!(!microsoft_url(
            &Url::parse("https://microsoft365.com.example.com/").unwrap()
        ));
        assert!(!microsoft_url(
            &Url::parse("https://user@www.microsoft365.com/").unwrap()
        ));
        assert!(notification_origin(
            &Url::parse("https://outlook.office.com/mail/").unwrap()
        ));
        assert!(!notification_origin(
            &Url::parse("https://untrusted.outlook.office.com/").unwrap()
        ));
    }

    #[test]
    fn opens_only_supported_outlook_hosts_in_the_default_browser() {
        assert_eq!(
            outlook_browser_url(&Url::parse("https://outlook.office365.com/").unwrap()),
            Some("https://outlook.office.com/")
        );
        assert_eq!(
            outlook_browser_url(&Url::parse("https://outlook.live.com/mail/").unwrap()),
            Some("https://outlook.live.com/mail/")
        );
        for value in [
            "http://outlook.office.com/",
            "https://outlook.office.com.example.com/",
            "https://outlook.office.com:444/",
            "https://user@outlook.office.com/",
            "https://www.microsoft365.com/",
        ] {
            assert_eq!(outlook_browser_url(&Url::parse(value).unwrap()), None);
        }
    }

    #[test]
    fn routes_outlook_links_only_from_other_windows() {
        for value in [
            "https://outlook.office365.com/mail/inbox",
            "https://outlook.office.com/mail/?path=/calendar",
            "https://outlook.cloud.microsoft/mail/",
            "https://outlook.live.com/mail/",
        ] {
            let url = Url::parse(value).unwrap();
            assert!(route_to_outlook_window(&url, false));
            assert!(!route_to_outlook_window(&url, true));
            assert!(allow_navigation(&url));
        }
        for value in [
            "https://login.microsoftonline.com/",
            "http://outlook.office.com/",
            "https://outlook.office.com.example.com/",
            "https://outlook.office.com:444/",
            "https://user@outlook.office.com/",
        ] {
            assert!(!route_to_outlook_window(&Url::parse(value).unwrap(), false));
        }
    }
}
