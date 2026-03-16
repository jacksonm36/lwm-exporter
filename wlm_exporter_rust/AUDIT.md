# Code audit: quality, bugs, security

Audit date: 2025. Summary of findings and fixes applied.

---

## Security

| Finding | Risk | Fix |
|--------|------|-----|
| **CSV formula injection** | If `export_results.csv` is opened in Excel, fields starting with `=`, `+`, `-`, `@` can be interpreted as formulas and execute code. | Added `csv_sanitize()`; any user-controlled or error-derived field written to CSV is sanitized (leading single quote) so Excel treats it as text. |
| **Symlink traversal** | Following symlinks under the WLM folder could read files outside the chosen directory. | `WalkDir` now explicitly uses `.follow_links(false)` so symlinks are not followed. |
| **OOM from huge .eml** | Reading an enormous .eml into memory could exhaust memory. | Added `MAX_EML_FILE_SIZE_BYTES` (50 MiB). Fallback parse checks file size before `fs::read` and returns a clear error if exceeded. |
| **Path traversal in folder names** | Components like `.` or `..` in relative paths could be used in folder cache keys. | When building PST folder keys from path components, we now filter out empty, `.`, and `..`. |

---

## Bugs / correctness

| Finding | Fix |
|--------|-----|
| **Cross-platform build** | `get_short_path()` uses Windows API; would not compile on non-Windows. | Added `#[cfg(windows)]` for the Windows implementation and `#[cfg(not(windows))]` fallback that returns the path as-is. |
| **Folder key from path parts** | Literal folder names `.` or `..` could create confusing cache keys. | Path components are filtered so only safe folder names are used in the cache. |

---

## Code quality

| Finding | Fix |
|--------|-----|
| **Symlinks** | Default behavior was correct but implicit. | Explicit `.follow_links(false)` for clarity and future-proofing. |
| **Constants** | Limits were already centralized. | Added `MAX_EML_FILE_SIZE_BYTES` for the new EML size check. |

---

## Not changed (by design)

- **Log file in cwd**: Log is written to `export_log.txt` in the current working directory. If cwd is not writable, the app fails at startup. Acceptable for a desktop tool; no change.
- **Export on UI thread**: Export runs on the GUI thread and blocks the UI. Improving this would require moving to an async/threaded design with progress messages; left as-is for this audit.
- **Dependencies**: `encoding_rs`, `chrono`, `thiserror` in `Cargo.toml` may be transitive (e.g. simplelog, mailparse, iced). Left as-is.

---

## Tauri v2 frontend (lwm-exporter-v2)

| Finding | Fix |
|--------|-----|
| **Error display** | Frontend `catch (e)` used `'Error: ' + e`; when Tauri returns an error object this shows `[object Object]`. | Added `errStr(e)` helper: prefer `e.message`, then `e.toString()`, else `String(e)`. |
| **Dialog return value** | `open()`/`save()` may return string, `{ path }`, or array depending on Tauri/dialog version. | Added `pathFromSelected(selected)` to normalize to a single path string (handles string, array[0], or `.path`). |
| **Input validation** | Empty trimmed paths could be passed to Rust. | In `run_export_command`, trim paths and return a clear error if either is empty. |
| **XSS** | User/error content shown in UI. | All dynamic content is set via `textContent` (no `innerHTML`), so no HTML/script injection. |

---

## Recommendations

1. When implementing real Outlook COM: validate and sanitize any strings passed to Outlook (e.g. folder names) to avoid injection or invalid names.
2. Consider making `MAX_EML_FILES` and `MAX_EML_FILE_SIZE_BYTES` configurable (e.g. via a config file or UI) for power users.
3. Run `cargo clippy` and `cargo test` (if tests are added) in CI.
