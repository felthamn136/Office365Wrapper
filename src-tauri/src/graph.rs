use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine};
use rand::{rngs::OsRng, RngCore};
use serde::Serialize;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::time::{Duration, Instant};
use tauri::{State, WebviewWindow};
use tokio::{
    io::{AsyncReadExt, AsyncWriteExt},
    net::TcpListener,
    process::Command,
    sync::Mutex,
};
use url::Url;

const SCOPES: &str = "Mail.Read Calendars.Read Files.Read.All";
const AUTHORITY: &str = "https://login.microsoftonline.com/organizations/oauth2/v2.0";

struct Token {
    access_token: String,
    expires: Instant,
}

pub struct GraphState {
    token: Mutex<Option<Token>>,
    client: reqwest::Client,
}

impl GraphState {
    pub fn new() -> Result<Self, reqwest::Error> {
        Ok(Self {
            token: Mutex::new(None),
            client: reqwest::Client::builder()
                .timeout(Duration::from_secs(30))
                .build()?,
        })
    }
}

#[derive(Serialize)]
pub struct SearchItem {
    title: String,
    url: Option<String>,
}

#[derive(Serialize)]
pub struct SearchGroup {
    label: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    results: Option<Vec<SearchItem>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    more: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
}

fn require_dashboard(window: &WebviewWindow) -> Result<(), String> {
    let url = window.url().map_err(|e| e.to_string())?;
    if window.label() == "main"
        && url.scheme() == "tauri"
        && url.host_str() == Some("localhost")
        && url.path() == "/dashboard.html"
        && url.query().is_none()
    {
        Ok(())
    } else {
        Err("Graph access is available only from the Office Dashboard.".into())
    }
}

fn client_id() -> Result<String, String> {
    let id = std::env::var("OFFICE365_CLIENT_ID").unwrap_or_default();
    if uuid::Uuid::parse_str(&id).is_err() {
        return Err("Set OFFICE365_CLIENT_ID to the application (client) ID of your Microsoft Entra app registration.".into());
    }
    Ok(id)
}

fn random_string() -> String {
    let mut bytes = [0u8; 32];
    OsRng.fill_bytes(&mut bytes);
    URL_SAFE_NO_PAD.encode(bytes)
}

async fn callback(listener: &TcpListener, expected_state: &str) -> Result<String, String> {
    loop {
        let (mut stream, _) = listener
            .accept()
            .await
            .map_err(|e| format!("Sign-in callback failed: {e}"))?;
        let request = tokio::time::timeout(Duration::from_secs(5), async {
            let mut buffer = Vec::new();
            while buffer.len() < 8192 && !buffer.windows(4).any(|part| part == b"\r\n\r\n") {
                let mut chunk = [0; 1024];
                let size = stream.read(&mut chunk).await?;
                if size == 0 {
                    break;
                }
                buffer.extend_from_slice(&chunk[..size]);
            }
            Ok::<_, std::io::Error>(buffer)
        })
        .await;
        let buffer = match request {
            Ok(Ok(buffer))
                if buffer.len() <= 8192 && buffer.windows(4).any(|part| part == b"\r\n\r\n") =>
            {
                buffer
            }
            _ => continue,
        };
        let first = String::from_utf8_lossy(&buffer);
        let path = first.lines().next().and_then(|line| {
            let mut fields = line.split_whitespace();
            if fields.next()? == "GET" {
                fields.next()
            } else {
                None
            }
        });
        let parsed = path.and_then(|path| Url::parse(&format!("http://localhost{path}")).ok());
        let valid = parsed.as_ref().is_some_and(|url| {
            url.path() == "/"
                && url.query_pairs().filter(|(k, _)| k == "state").count() == 1
                && url
                    .query_pairs()
                    .any(|(k, v)| k == "state" && v == expected_state)
        });
        let (response, result) = if valid {
            let url = parsed.expect("validated URL");
            let codes: Vec<_> = url
                .query_pairs()
                .filter(|(k, _)| k == "code")
                .map(|(_, v)| v.into_owned())
                .collect();
            let code = if codes.len() == 1 && !codes[0].is_empty() {
                codes.into_iter().next()
            } else {
                None
            };
            if let Some(code) = code {
                ("HTTP/1.1 200 OK\r\nContent-Type: text/plain\r\nConnection: close\r\n\r\nSign-in complete. You can close this tab.", Some(Ok(code)))
            } else {
                ("HTTP/1.1 400 Bad Request\r\nContent-Type: text/plain\r\nConnection: close\r\n\r\nSign-in was not completed.", Some(Err("Microsoft sign-in was cancelled or failed.".into())))
            }
        } else {
            ("HTTP/1.1 400 Bad Request\r\nContent-Type: text/plain\r\nConnection: close\r\n\r\nInvalid sign-in callback.", None)
        };
        if let Err(error) = stream.write_all(response.as_bytes()).await {
            eprintln!("Could not send the sign-in callback response: {error}");
        }
        if let Some(result) = result {
            return result;
        }
    }
}

async fn authenticate(client: &reqwest::Client) -> Result<Token, String> {
    let client_id = client_id()?;
    let listener = TcpListener::bind("127.0.0.1:0")
        .await
        .map_err(|e| format!("Could not start the sign-in listener: {e}"))?;
    let port = listener.local_addr().map_err(|e| e.to_string())?.port();
    let redirect = format!("http://localhost:{port}/");
    let state = random_string();
    let verifier = random_string();
    let challenge = URL_SAFE_NO_PAD.encode(Sha256::digest(verifier.as_bytes()));
    let mut url = Url::parse(&format!("{AUTHORITY}/authorize")).map_err(|e| e.to_string())?;
    url.query_pairs_mut().extend_pairs([
        ("client_id", client_id.as_str()),
        ("response_type", "code"),
        ("redirect_uri", redirect.as_str()),
        ("response_mode", "query"),
        ("scope", SCOPES),
        ("state", state.as_str()),
        ("code_challenge", challenge.as_str()),
        ("code_challenge_method", "S256"),
    ]);
    let mut browser = Command::new("xdg-open")
        .arg(url.as_str())
        .spawn()
        .map_err(|e| format!("Could not open the sign-in browser: {e}"))?;
    let code = tokio::time::timeout(Duration::from_secs(300), async {
        tokio::select! {
            result = callback(&listener, &state) => result,
            status = browser.wait() => {
                if !status.map_err(|e| format!("Could not open the sign-in browser: {e}"))?.success() {
                    return Err("Could not open the sign-in browser.".into());
                }
                callback(&listener, &state).await
            }
        }
    }).await.map_err(|_| "Microsoft sign-in timed out.".to_string())??;
    let response = client
        .post(format!("{AUTHORITY}/token"))
        .form(&[
            ("client_id", client_id.as_str()),
            ("grant_type", "authorization_code"),
            ("code", code.as_str()),
            ("redirect_uri", redirect.as_str()),
            ("code_verifier", verifier.as_str()),
            ("scope", SCOPES),
        ])
        .send()
        .await
        .map_err(|e| format!("Token request failed: {e}"))?;
    if !response.status().is_success() {
        return Err(format!(
            "Microsoft sign-in failed (HTTP {}).",
            response.status().as_u16()
        ));
    }
    let body: Value = response
        .json()
        .await
        .map_err(|e| format!("Invalid token response: {e}"))?;
    let access_token = body["access_token"]
        .as_str()
        .filter(|s| !s.is_empty())
        .ok_or("Microsoft sign-in did not return an access token.")?
        .to_owned();
    let lifetime = body["expires_in"]
        .as_u64()
        .ok_or("Microsoft sign-in did not return a token expiry.")?;
    Ok(Token {
        access_token,
        expires: Instant::now() + Duration::from_secs(lifetime.saturating_sub(60)),
    })
}

fn result_url(value: &Value) -> Option<String> {
    let url = Url::parse(value.as_str()?).ok()?;
    if url.scheme() == "https" && url.username().is_empty() && url.password().is_none() {
        Some(url.into())
    } else {
        None
    }
}

fn parse_results(data: &Value, entity: &str) -> Result<(Vec<SearchItem>, bool), String> {
    let containers = data["value"]
        .as_array()
        .and_then(|v| v.first())
        .and_then(|v| v["hitsContainers"].as_array())
        .ok_or("Microsoft Graph returned an unexpected search response.")?;
    let mut items = Vec::new();
    let mut more = false;
    for container in containers {
        more |= container["moreResultsAvailable"].as_bool() == Some(true);
        for hit in container["hits"]
            .as_array()
            .ok_or("Microsoft Graph returned an unexpected search results container.")?
        {
            let resource = &hit["resource"];
            let title = resource[if entity == "driveItem" {
                "name"
            } else {
                "subject"
            }]
            .as_str()
            .filter(|s| !s.trim().is_empty())
            .unwrap_or("(Untitled)")
            .to_owned();
            let url = result_url(&resource["webUrl"]).or_else(|| result_url(&resource["webLink"]));
            items.push(SearchItem { title, url });
        }
    }
    Ok((items, more))
}

async fn search_category(
    client: &reqwest::Client,
    token: &str,
    query: &str,
    entity: &str,
    label: &'static str,
) -> SearchGroup {
    let result = async {
        let response = client.post("https://graph.microsoft.com/v1.0/search/query")
            .bearer_auth(token).json(&json!({"requests": [{"entityTypes": [entity], "query": {"queryString": query}, "from": 0, "size": 25}]}))
            .send().await.map_err(|e| format!("Search request failed: {e}"))?;
        if !response.status().is_success() { return Err(format!("Microsoft Graph returned HTTP {}.", response.status().as_u16())); }
        let data = response.json::<Value>().await.map_err(|e| format!("Invalid Graph response: {e}"))?;
        parse_results(&data, entity)
    }.await;
    match result {
        Ok((results, more)) => SearchGroup {
            label,
            results: Some(results),
            more: Some(more),
            error: None,
        },
        Err(error) => SearchGroup {
            label,
            results: None,
            more: None,
            error: Some(error),
        },
    }
}

#[tauri::command]
pub async fn graph_search(
    window: WebviewWindow,
    state: State<'_, GraphState>,
    query: String,
) -> Result<Vec<SearchGroup>, String> {
    require_dashboard(&window)?;
    let query = query.trim();
    if query.is_empty() || query.encode_utf16().count() > 200 {
        return Err("Enter a search term between 1 and 200 characters.".into());
    }
    let mut token = state.token.lock().await;
    if token
        .as_ref()
        .is_none_or(|token| token.expires <= Instant::now())
    {
        *token = Some(authenticate(&state.client).await?);
    }
    require_dashboard(&window)?;
    let access_token = &token.as_ref().expect("authenticated token").access_token;
    let (mail, calendar, files) = tokio::join!(
        search_category(&state.client, access_token, query, "message", "Mail"),
        search_category(&state.client, access_token, query, "event", "Calendar"),
        search_category(&state.client, access_token, query, "driveItem", "Files"),
    );
    require_dashboard(&window)?;
    Ok(vec![mail, calendar, files])
}

#[tauri::command]
pub async fn graph_sign_out(
    window: WebviewWindow,
    state: State<'_, GraphState>,
) -> Result<(), String> {
    require_dashboard(&window)?;
    *state.token.lock().await = None;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_graph_results_and_rejects_unsafe_urls() {
        let input = json!({"value": [{"hitsContainers": [{"moreResultsAvailable": true, "hits": [
            {"resource": {"name": "Report", "webUrl": "https://example.com/report"}},
            {"resource": {"webUrl": "file:///etc/passwd"}},
            {"resource": {"name": "Unsafe", "webUrl": "https://user@example.com/"}}
        ]}]}]});
        let (items, more) = parse_results(&input, "driveItem").unwrap();
        assert!(more);
        assert_eq!(items[0].title, "Report");
        assert_eq!(items[0].url.as_deref(), Some("https://example.com/report"));
        assert_eq!(items[1].title, "(Untitled)");
        assert!(items[1].url.is_none());
        assert!(items[2].url.is_none());
        assert!(parse_results(&json!({"value": []}), "message").is_err());
    }

    #[tokio::test(flavor = "current_thread")]
    async fn callback_rejects_wrong_state_and_accepts_fragmented_requests() {
        let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
        let addr = listener.local_addr().unwrap();
        let sender = tokio::spawn(async move {
            let mut bad = tokio::net::TcpStream::connect(addr).await.unwrap();
            bad.write_all(b"GET /?state=bad&code=x HTTP/1.1\r\nHost: localhost\r\n\r\n")
                .await
                .unwrap();
            let mut response = [0; 256];
            let size = bad.read(&mut response).await.unwrap();
            assert!(String::from_utf8_lossy(&response[..size]).contains("400 Bad Request"));
            let mut good = tokio::net::TcpStream::connect(addr).await.unwrap();
            good.write_all(b"GET /?state=secret&").await.unwrap();
            tokio::task::yield_now().await;
            good.write_all(b"code=correct HTTP/1.1\r\nHost: localhost\r\n\r\n")
                .await
                .unwrap();
            let size = good.read(&mut response).await.unwrap();
            assert!(String::from_utf8_lossy(&response[..size]).contains("200 OK"));
        });
        assert_eq!(callback(&listener, "secret").await.unwrap(), "correct");
        sender.await.unwrap();
    }
}
