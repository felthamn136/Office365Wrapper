fn main() {
    tauri_build::try_build(tauri_build::Attributes::new().app_manifest(
        tauri_build::AppManifest::new().commands(&["graph_search", "graph_sign_out"]),
    ))
    .expect("failed to build Tauri application metadata")
}
