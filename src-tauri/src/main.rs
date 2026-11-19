mod graph;

use std::{
    os::unix::fs::DirBuilderExt,
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

fn secure_url(url: &Url) -> bool {
    url.scheme() == "https" && url.username().is_empty() && url.password().is_none()
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

fn create_window(
    app: &AppHandle,
    label: &str,
    initial: WebviewUrl,
    features: Option<NewWindowFeatures>,
) -> tauri::Result<WebviewWindow> {
    let profile = app.path().app_data_dir()?.join("webkit");
    std::fs::DirBuilder::new()
        .recursive(true)
        .mode(0o700)
        .create(&profile)?;
    let handle = app.clone();
    let parent_label = label.to_owned();
    let builder = WebviewWindowBuilder::new(app, label, initial)
        .title("Office Dashboard")
        .inner_size(1200.0, 800.0)
        .data_directory(profile)
        .on_navigation(allow_navigation)
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
            ) {
                Ok(child) => {
                    if let Some(parent) = handle.get_webview_window(&parent_label) {
                        let popup = child.clone();
                        parent.on_window_event(move |event| {
                            if matches!(event, WindowEvent::Destroyed) {
                                if let Err(error) = popup.close() {
                                    eprintln!("Could not close popup: {error}");
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
    let builder = if let Some(features) = features {
        builder.window_features(features)
    } else {
        builder
    };
    let window = builder.build()?;
    install_permissions(&window)?;
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
    let initial = initial_url().map_err(std::io::Error::other)?;
    tauri::Builder::default()
        .manage(AppState {
            notifications: AtomicBool::new(false),
        })
        .manage(graph::GraphState::new()?)
        .invoke_handler(tauri::generate_handler![
            graph::graph_search,
            graph::graph_sign_out
        ])
        .setup(move |app| {
            let handle = app.handle();
            let navigate = SubmenuBuilder::new(handle, "Navigate")
                .text("dashboard", "Office Dashboard")
                .text("microsoft365", "Microsoft 365")
                .text("outlook", "Outlook (work or school)")
                .text("personal", "Outlook (personal)")
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
            create_window(handle, "main", initial, None)?;
            Ok(())
        })
        .on_menu_event(|app, event| {
            if event.id().as_ref() == "notifications" {
                let state = app.state::<AppState>();
                state.notifications.fetch_xor(true, Ordering::Relaxed);
                return;
            }
            if let Some(window) = app.get_webview_window("main") {
                let url = match event.id().as_ref() {
                    "dashboard" => Some(dashboard_url()),
                    "microsoft365" => Some("https://www.microsoft365.com/".parse().unwrap()),
                    "outlook" => Some("https://outlook.office365.com/".parse().unwrap()),
                    "personal" => Some("https://outlook.live.com/mail/".parse().unwrap()),
                    "reload" => {
                        if let Err(error) = window.reload() {
                            eprintln!("Could not reload page: {error}");
                        }
                        None
                    }
                    _ => None,
                };
                if let Some(url) = url {
                    if let Err(error) = window.navigate(url) {
                        eprintln!("Could not navigate: {error}");
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
}
