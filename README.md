# APIMonitor

A Windows system tray application that monitors an API endpoint and displays its status as a tray icon. The icon changes color based on the API response, giving an at-a-glance view of service health.

## Features

- System tray icon that reflects API status (success, fail, error)
- Configurable API URL, check interval, and logging
- Configuration stored in the Windows registry (`HKCU\SOFTWARE\JPIT\APIMonitor`)
- First-launch configuration dialog
- Automatic retry on network errors (3 attempts, 2s delay)
- Accelerated polling (every 10s) when the API is in a non-success state
- Log file at `ProgramData\APIMonitor\APIMonitor.log` (auto-truncated at 10MB)
- Single-instance enforcement
- Display/DPI change detection for RDP reconnects

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

Requires MinGW-w64 cross-compiler.

```sh
make
```

The executable is output to `release/APIMonitor.exe`.

To regenerate `.ico` files from `.svg` sources (requires ImageMagick):

```sh
make icons
```

To clean build artifacts:

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

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\JPIT\APIMonitor`.

If a `config.ini` file exists from a previous version, settings are migrated to the registry on first launch.
