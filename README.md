# APIMonitor

A Windows system tray application that monitors an API endpoint and displays its status as a tray icon. The icon changes color based on the API response, giving an at-a-glance view of service health.

## Features

- System tray icon that reflects API status (success, fail, error)
- Configurable API URL with live validation, check interval, logging toggle, and history limit
- Modern WebView2-based configuration and history dialogs (React + Tailwind CSS)
- Status change history with timestamps, copy-to-clipboard, and clear
- Configuration stored in the Windows registry (`HKCU\SOFTWARE\JPIT\APIMonitor`)
- First-launch configuration dialog
- Automatic retry on network errors (3 attempts, 2s delay)
- Accelerated polling (every 10s) when the API is in a non-success state
- Log file at `ProgramData\APIMonitor\APIMonitor.log` (auto-truncated at 10MB)
- Single-instance enforcement
- Display/DPI change detection for RDP reconnects

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
| Success | Green | Configured interval (default 60s) |
| Fail | Red | 10 seconds |
| Error (network/HTTP) | Empty | 10 seconds |
| Invalid (bad XML) | Empty | 10 seconds |

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
| Check Interval | `RefreshInterval` | REG_DWORD | `60` (seconds) |
| Enable Logging | `LoggingEnabled` | REG_DWORD | `1` |
| History Limit | `HistoryLimit` | REG_DWORD | `100` (10–10,000) |

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\JPIT\APIMonitor`.

If a `config.ini` file exists from a previous version, settings are migrated to the registry on first launch.

## Project Structure

```
├── main.c              # Application source (tray icon, API polling, WebView2 integration)
├── resource.h          # Resource IDs
├── resources.rc        # Resource definitions (icons, HTML, DLL)
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
