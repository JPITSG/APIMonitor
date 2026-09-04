# APIMonitor

A Windows system tray application that monitors an API endpoint and displays its status as a tray icon. The icon changes color based on the API response, giving an at-a-glance view of service health.

## Features

- System tray icon that reflects API status (success, fail, error)
- Configurable API URL with live validation, healthy/down check intervals, logging toggle, and history limit
- Modern WebView2-based configuration and history dialogs (React + Tailwind CSS)
- Crisp Per-Monitor V2 rendering on displays with Windows scaling enabled
- Status change history with timestamps, copy-to-clipboard, and clear
- Configuration stored in the Windows registry (`HKCU\SOFTWARE\JPIT\APIMonitor`)
- First-launch configuration dialog
- Automatic retry on network errors (3 attempts, 2s delay)
- Independently configurable polling intervals for healthy and non-success states
- Completion-driven polling that never runs more than one API check at a time
- Log file at `ProgramData\APIMonitor\APIMonitor.log` (auto-truncated at 10MB)
- Single-instance enforcement
- Display/DPI change detection for RDP reconnects
- Self-update flow with embedded version checks, cancellable downloads, UAC replacement, rollback, and automatic hourly checks

## Requirements

- **Windows 7+**
- **Microsoft Edge WebView2 Runtime** — required for the configuration and history dialogs. Usually pre-installed on Windows 10/11; can be downloaded from [Microsoft](https://developer.microsoft.com/en-us/microsoft-edge/webview2/).

## API Response Format

The monitored endpoint must return XML containing a `<result>` (or short `<r>`) tag with a value of `success` or `fail`. An optional `<message>` tag provides detail shown in the tooltip.

### Success Response

```xml
<result>success</result>
<message>All systems operational</message>
```

Or using the short tag:

```xml
<r>success</r>
<message>All systems operational</message>
```

### Fail Response

```xml
<result>fail</result>
<message>Database connection timeout</message>
```

### Tray Icon States

| State | Icon | Refresh Interval |
|-------|------|------------------|
| Success | Green | Configured healthy interval (default 60s) |
| Fail | Red | Configured down interval (default 10s) |
| Error (network/HTTP) | Empty | Configured down interval (default 10s) |
| Invalid (bad XML) | Empty | Configured down interval (default 10s) |

Each interval starts after the preceding API check, including its retry cycle,
has fully completed. A new scheduled or manual check cannot overlap one already
in progress.

## Building

Requires MinGW-w64 cross-compiler and Node.js (for the frontend build).

```sh
make
```

This builds the React frontend (`assets/dist/index.html`), compiles resources, and outputs `release/APIMonitor.exe`.

To regenerate `.ico` files from `.svg` sources (requires ImageMagick):

```sh
make icons
```

To clean all build artifacts (including `assets/dist` and `assets/node_modules`):

```sh
make clean
```

## Configuration

On first launch a configuration dialog is shown. It can also be opened from the tray icon right-click menu under **Configure**.

| Setting | Registry Value | Type | Default |
|---------|---------------|------|---------|
| API URL | `ApiUrl` | REG_SZ | `http://example.com/api/status` |
| Everything OK Interval | `RefreshInterval` | REG_DWORD | `60` (seconds) |
| Something Down Interval | `DownRefreshInterval` | REG_DWORD | `10` (seconds) |
| Enable Logging | `LoggingEnabled` | REG_DWORD | `1` |
| History Limit | `HistoryLimit` | REG_DWORD | `100` (10–10,000) |
| Automatically Check for Updates | `AutoCheckForUpdates` | REG_DWORD | `1` |

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\JPIT\APIMonitor`.

If a `config.ini` file exists from a previous version, settings are migrated to the registry on first launch.

## Updates

When automatic checks are enabled, APIMonitor checks at startup, whenever the
configuration dialog opens, and every 60 minutes. A newer build opens Configure
and displays its update prompt. **Ignore this version** suppresses that version
during later automatic checks; the manual **Update** button still displays every
result and can reinstall the current version.

Checks download [`release/APIMonitor.exe`](release/APIMonitor.exe) to the user's
temporary directory and compare its embedded Windows file version with the
running executable. The download is size-limited, reports transfer speed, and
can be cancelled. An older repository build is never installable.

Installing uses a short-lived elevated helper to replace the executable and
restart APIMonitor in the user's normal session. If replacement or restart
fails, the previous executable is restored. Temporary files are removed after
the restarted application confirms a successful handoff.

## Project Structure

```
├── main.c              # Application source (tray icon, API polling, WebView2 integration)
├── resource.h          # Resource IDs
├── resources.rc        # Resource definitions (icons, HTML, DLL)
├── APIMonitor.manifest # Windows compatibility and Per-Monitor V2 DPI awareness
├── version.h           # Application and Windows resource version
├── Makefile            # Cross-compilation build system
├── assets/
│   ├── src/
│   │   ├── App.tsx           # Root component (view router, resize reporting)
│   │   ├── ConfigView.tsx    # Configuration form with URL validation
│   │   ├── HistoryView.tsx   # Status change history table
│   │   ├── lib/
│   │   │   ├── bridge.ts     # C <-> JS communication bridge
│   │   │   └── utils.ts      # Tailwind merge utility
│   │   └── components/ui/    # Reusable UI components (button, input, switch, etc.)
│   ├── *.ico, *.svg          # Tray icons and source SVGs
│   ├── WebView2Loader.dll    # Embedded WebView2 loader
│   ├── package.json          # Frontend dependencies
│   ├── vite.config.ts        # Vite + single-file plugin config
│   └── tailwind.config.ts    # Tailwind CSS config
└── release/
    └── APIMonitor.exe        # Built executable
```

## License

[MIT](LICENSE)
