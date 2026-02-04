// main.c
#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <winhttp.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <commctrl.h>
#include "resource.h"

#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "advapi32.lib")

#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define ID_TRAY_REFRESH 1002
#define ID_TRAY_CONFIGURE 1004
#define ID_TRAY_HISTORY 1005

// Registry settings
#define REG_KEY_PATH        "SOFTWARE\\JPIT\\APIMonitor"
#define REG_VALUE_URL       "ApiUrl"
#define REG_VALUE_INTERVAL  "RefreshInterval"
#define REG_VALUE_LOGGING   "LoggingEnabled"
#define REG_VALUE_CONFIGURED "Configured"
#define REG_VALUE_HISTORY_LIMIT "HistoryLimit"
#define REG_VALUE_HISTORY_COUNT "HistoryCount"
#define REG_VALUE_HISTORY_DATA  "HistoryData"

// Dialog control IDs
#define IDC_EDIT_URL        2101
#define IDC_STATIC_URL      2102
#define IDC_COMBO_INTERVAL  2103
#define IDC_STATIC_INTERVAL 2104
#define IDC_CHECK_LOGGING   2105
#define IDC_STATIC_VALID_LABEL  2106
#define IDC_STATIC_VALID_STATUS 2107
#define IDT_VALIDATE_DEBOUNCE   3001
#define WM_VALIDATE_RESULT      (WM_APP + 1)
#define IDC_EDIT_HISTORY_LIMIT  2108
#define IDC_STATIC_HISTORY      2109
#define IDC_HISTORY_LIST        2110
#define IDC_HISTORY_CLOSE       2111

typedef enum {
    RESULT_ERROR,        // Connection/network error
    RESULT_INVALID,      // Connected but invalid response
    RESULT_SUCCESS,
    RESULT_FAIL
} ApiResult;

typedef struct {
    ApiResult result;
    char message[256];
} ApiResponse;

typedef struct {
    int attempt;
    int maxAttempts;
} ThreadParams;

typedef struct {
    SYSTEMTIME timestamp;
    ApiResult oldResult;
    ApiResult newResult;
    char oldMessage[256];
    char newMessage[256];
} HistoryEntry;

// Global variables
static NOTIFYICONDATA nid = {0};
static HMENU hMenu = NULL;
static char configApiUrl[512] = "http://example.com/api/status";
static int configRefreshInterval = 60;
static char logFilePath[MAX_PATH];
static HICON hIconEmpty = NULL;
static HICON hIconSuccess = NULL;
static HICON hIconFail = NULL;
static HICON hIconBlank = NULL;
static UINT_PTR timerRefresh = 0;
static UINT_PTR timerTooltip = 0;
static BOOL iconVisible = TRUE;
static HICON currentIcon = NULL;
static char currentMessage[256] = "";
static ApiResult currentResult = RESULT_ERROR;
static SYSTEMTIME lastUpdateTime = {0};
static HWND g_hwnd = NULL;
static HINSTANCE g_hInstance = NULL;
static HANDLE g_hMutex = NULL;  // Mutex for single instance check
static CRITICAL_SECTION logCriticalSection;  // For thread-safe logging
static BOOL configLoggingEnabled = TRUE; // Global variable for logging toggle (default true)
static int configHistoryLimit = 100;
static HistoryEntry* historyBuffer = NULL;
static int historyCapacity = 0;
static int historyCount = 0;
static int historyHead = 0;

// URL validation thread params
typedef struct {
    char url[512];
    HWND hDlg;
    LONG generation;
} ValidateParams;

static volatile LONG g_validateGeneration = 0;

// Display settings tracking (for RDP reconnect icon refresh)
static int lastScreenWidth = 0;
static int lastScreenHeight = 0;
static int lastDpiX = 0;
static int lastDpiY = 0;

// Function prototypes
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void InitTrayIcon(HWND hwnd);
void CreateContextMenu();
BOOL LoadConfigFromRegistry();
void SaveConfigToRegistry();
BOOL IsFirstLaunch();
void MarkAsConfigured();
void LoadConfigFromIni(const char* iniPath);
void ApplyConfiguration();
void ShowConfigDialog(HWND hwndParent);
void UpdateStatus(ApiResult result, const char* message);
void RefreshStatus();
DWORD WINAPI RefreshThread(LPVOID param);
void SetIcon(HICON icon);
void CALLBACK TooltipTimer(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
void CALLBACK RefreshTimer(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime);
void ParseXmlResponse(const char* xml, ApiResponse* response);
void ExitApplication(HWND hwnd);
void UpdateTooltip();
void SetRefreshInterval(int seconds, BOOL isUserSetting);
void CaptureCurrentDisplaySettings();
BOOL HasDisplaySettingsChanged();
void RefreshTrayIconForNewResolution();
void LogMessage(const char* format, ...);
void CheckLogFileSize();
DWORD WINAPI ValidateUrlThread(LPVOID param);
const char* ApiResultToString(ApiResult r);
void InitHistoryBuffer(int capacity);
void AddHistoryEntry(ApiResult oldResult, const char* oldMsg, ApiResult newResult, const char* newMsg);
HistoryEntry* GetHistoryEntry(int displayIndex);
void FreeHistoryBuffer(void);
void ShowHistoryDialog(HWND hwndParent);
void SaveHistoryToRegistry(void);
void LoadHistoryFromRegistry(void);

// Logging function: writes to ProgramData\APIMonitor.log with timestamp and thread ID
void LogMessage(const char* format, ...) {
	// Skip logging if disabled
    if (!configLoggingEnabled) {
        return;
    }

	EnterCriticalSection(&logCriticalSection);

    // Check log file size before writing (limit to ~10MB)
    CheckLogFileSize();

    FILE* logFile = fopen(logFilePath, "a");
    if (!logFile) {
        LeaveCriticalSection(&logCriticalSection);
        return;
    }

    // Timestamp
    SYSTEMTIME st;
    GetLocalTime(&st);
    fprintf(logFile, "[%04d-%02d-%02d %02d:%02d:%02d] [TID:%lu] ",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond,
            GetCurrentThreadId());

    // Message
    va_list args;
    va_start(args, format);
    vfprintf(logFile, format, args);
    va_end(args);

    fprintf(logFile, "\n");
    fclose(logFile);
    LeaveCriticalSection(&logCriticalSection);
}

void CheckLogFileSize() {
    FILE* f = fopen(logFilePath, "r");
    if (!f) return;

    fseek(f, 0, SEEK_END);
    long size = ftell(f);
    fclose(f);

    // If larger than 10MB, truncate
    if (size > 10 * 1024 * 1024) {
        fclose(fopen(logFilePath, "w"));
        LogMessage("Log file exceeded 10MB, restarted.");
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    // Initialize logging system
    InitializeCriticalSection(&logCriticalSection);

    // Get ProgramData folder for logging
    char programDataPath[MAX_PATH];
    if (SHGetFolderPathA(NULL, CSIDL_COMMON_APPDATA, NULL, 0, programDataPath) == S_OK) {
        strcpy(logFilePath, programDataPath);
        strcat(logFilePath, "\\APIMonitor");
        CreateDirectoryA(logFilePath, NULL);
        strcat(logFilePath, "\\APIMonitor.log");
    } else {
        // Fallback to executable directory
        char exePath[MAX_PATH];
        GetModuleFileNameA(NULL, exePath, MAX_PATH);
        char* lastSlash = strrchr(exePath, '\\');
        if (lastSlash) *lastSlash = '\0';
        strcpy(logFilePath, exePath);
        strcat(logFilePath, "\\APIMonitor.log");
    }

    LogMessage("=== Application starting (Version: APIMonitor/1.0) ===");

    // Single instance check
    g_hMutex = CreateMutexA(NULL, TRUE, "Global\\APIMonitor_SingleInstance_Mutex");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        LogMessage("ERROR: Another instance is already running. Exiting.");
        MessageBoxA(NULL,
                   "Only one copy of the API monitor can be running at any given time.",
                   "API Monitor Already Running",
                   MB_OK | MB_ICONINFORMATION);
        if (g_hMutex) CloseHandle(g_hMutex);
        DeleteCriticalSection(&logCriticalSection);
        return 0;
    }
    LogMessage("Single instance check passed.");

    g_hInstance = hInstance;

    // Check if this is a first launch (no Configured flag in registry)
    BOOL firstLaunch = IsFirstLaunch();

    // Try loading configuration from registry
    BOOL loadedFromRegistry = LoadConfigFromRegistry();

    // If registry was empty, try INI migration
    if (!loadedFromRegistry) {
        char iniPath[MAX_PATH];
        GetModuleFileNameA(NULL, iniPath, MAX_PATH);
        char* lastSlash = strrchr(iniPath, '\\');
        if (lastSlash) *lastSlash = '\0';
        strcat(iniPath, "\\config.ini");
        LoadConfigFromIni(iniPath);
        SaveConfigToRegistry();
        LogMessage("Migrated configuration from INI to registry.");
    }

    LogMessage("Configuration loaded: URL=%s, Interval=%d, Logging=%s, HistoryLimit=%d",
               configApiUrl, configRefreshInterval, configLoggingEnabled ? "enabled" : "disabled", configHistoryLimit);

    // Initialize history buffer
    InitHistoryBuffer(configHistoryLimit);
    LoadHistoryFromRegistry();

    // On first launch, show configuration dialog
    if (firstLaunch) {
        LogMessage("First launch detected, showing configuration dialog.");
        ShowConfigDialog(NULL);
        MarkAsConfigured();
    }

    // Load icons
    hIconEmpty = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(IDI_EMPTY),
                                   IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
    hIconSuccess = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(IDI_SUCCESS),
                                     IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
    hIconFail = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(IDI_FAIL),
                                  IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
    hIconBlank = (HICON)LoadImage(hInstance, MAKEINTRESOURCE(IDI_BLANK),
                                   IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);

    if (!hIconEmpty || !hIconSuccess || !hIconFail || !hIconBlank) {
        char errMsg[256];
        sprintf(errMsg, "Failed to load embedded icons. Error: %lu", GetLastError());
        LogMessage("ERROR: %s", errMsg);
        MessageBoxA(NULL, errMsg, "Icon Loading Error", MB_OK | MB_ICONERROR);
        DeleteCriticalSection(&logCriticalSection);
        return 1;
    }
    LogMessage("Icons loaded successfully.");

    // Capture initial display settings
    CaptureCurrentDisplaySettings();

    // Create window class
    WNDCLASSEXA wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXA);
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = "APIMonitorClass";

    if (!RegisterClassExA(&wc)) {
        LogMessage("ERROR: Failed to register window class. Error: %lu", GetLastError());
        MessageBoxA(NULL, "Failed to register window class", "Error", MB_OK | MB_ICONERROR);
        DeleteCriticalSection(&logCriticalSection);
        return 1;
    }

    // Create hidden message window
    HWND hwnd = CreateWindowExA(0, "APIMonitorClass", "APIMonitor", 0,
                              0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);
    if (!hwnd) {
        LogMessage("ERROR: Failed to create window. Error: %lu", GetLastError());
        MessageBoxA(NULL, "Failed to create window", "Error", MB_OK | MB_ICONERROR);
        DeleteCriticalSection(&logCriticalSection);
        return 1;
    }

    g_hwnd = hwnd;
    LogMessage("Message window created.");

    // Initialize tray icon
    InitTrayIcon(hwnd);
    LogMessage("Tray icon initialized.");

    // Create context menu
    CreateContextMenu();

    // Set up timers
    timerTooltip = SetTimer(hwnd, 2, 1000, TooltipTimer);

    // Initial check
    LogMessage("Performing initial API check.");
    RefreshStatus();

    // Start refresh timer
    timerRefresh = SetTimer(hwnd, 1, configRefreshInterval * 1000, RefreshTimer);
    LogMessage("Refresh timer started with %d second interval.", configRefreshInterval);

    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    ExitApplication(hwnd);
    DeleteCriticalSection(&logCriticalSection);
    return 0;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_TRAYICON:
            if (lParam == WM_RBUTTONUP) {
                POINT pt;
                GetCursorPos(&pt);
                SetForegroundWindow(hwnd);
                TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
                LogMessage("Context menu opened at position (%ld, %ld).", pt.x, pt.y);
            } else if (lParam == WM_LBUTTONDBLCLK) {
                LogMessage("Tray icon double-clicked. Triggering manual refresh.");
                RefreshStatus();
            }
            break;

        case WM_DISPLAYCHANGE:
            Sleep(1000);
            if (HasDisplaySettingsChanged()) {
                LogMessage("Display settings changed. Screen: %dx%d, DPI: %dx%d -> %dx%d, DPI: %dx%d",
                          lastScreenWidth, lastScreenHeight, lastDpiX, lastDpiY,
                          GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN),
                          GetDeviceCaps(GetDC(NULL), LOGPIXELSX), GetDeviceCaps(GetDC(NULL), LOGPIXELSY));
                RefreshTrayIconForNewResolution();
                CaptureCurrentDisplaySettings();
            }
            break;

        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case ID_TRAY_EXIT:
                    LogMessage("User selected Exit from context menu.");
                    ExitApplication(hwnd);
                    break;
                case ID_TRAY_REFRESH:
                    LogMessage("User selected Refresh from context menu.");
                    RefreshStatus();
                    break;
                case ID_TRAY_CONFIGURE:
                    LogMessage("User selected Configure from context menu.");
                    ShowConfigDialog(g_hwnd);
                    break;
                case ID_TRAY_HISTORY:
                    LogMessage("User selected History from context menu.");
                    ShowHistoryDialog(g_hwnd);
                    break;
            }
            break;

        case WM_TIMER:
            if (wParam == 1) RefreshTimer(hwnd, uMsg, wParam, 0);
            else if (wParam == 2) TooltipTimer(hwnd, uMsg, wParam, 0);
            break;

        case WM_DESTROY:
            LogMessage("Window destroyed.");
            PostQuitMessage(0);
            break;

        default:
            return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

void InitTrayIcon(HWND hwnd) {
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = hwnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = hIconEmpty;
    strcpy(nid.szTip, "API Monitor - Initializing...");
    Shell_NotifyIconA(NIM_ADD, &nid);
    LogMessage("Tray icon added to system tray.");
}

void CreateContextMenu() {
    hMenu = CreatePopupMenu();
    AppendMenuA(hMenu, MF_STRING, ID_TRAY_REFRESH, "Refresh");
    AppendMenuA(hMenu, MF_STRING, ID_TRAY_HISTORY, "History");
    AppendMenuA(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuA(hMenu, MF_STRING, ID_TRAY_CONFIGURE, "Configure");
    AppendMenuA(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuA(hMenu, MF_STRING, ID_TRAY_EXIT, "Exit");
    LogMessage("Context menu created.");
}

// --- Registry-based configuration ---

BOOL LoadConfigFromRegistry() {
    HKEY hKey;
    LONG result = RegOpenKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) {
        return FALSE;
    }

    DWORD type, size;

    // Read ApiUrl (REG_SZ)
    size = sizeof(configApiUrl);
    if (RegQueryValueExA(hKey, REG_VALUE_URL, NULL, &type, (LPBYTE)configApiUrl, &size) != ERROR_SUCCESS
        || type != REG_SZ) {
        strcpy(configApiUrl, "http://example.com/api/status");
    }

    // Read RefreshInterval (REG_DWORD)
    DWORD dwInterval = 60;
    size = sizeof(dwInterval);
    if (RegQueryValueExA(hKey, REG_VALUE_INTERVAL, NULL, &type, (LPBYTE)&dwInterval, &size) == ERROR_SUCCESS
        && type == REG_DWORD) {
        configRefreshInterval = (int)dwInterval;
    }

    // Read LoggingEnabled (REG_DWORD)
    DWORD dwLogging = 1;
    size = sizeof(dwLogging);
    if (RegQueryValueExA(hKey, REG_VALUE_LOGGING, NULL, &type, (LPBYTE)&dwLogging, &size) == ERROR_SUCCESS
        && type == REG_DWORD) {
        configLoggingEnabled = (BOOL)dwLogging;
    }

    // Read HistoryLimit (REG_DWORD)
    DWORD dwHistoryLimit = 100;
    size = sizeof(dwHistoryLimit);
    if (RegQueryValueExA(hKey, REG_VALUE_HISTORY_LIMIT, NULL, &type, (LPBYTE)&dwHistoryLimit, &size) == ERROR_SUCCESS
        && type == REG_DWORD) {
        configHistoryLimit = (int)dwHistoryLimit;
        if (configHistoryLimit < 10) configHistoryLimit = 10;
        if (configHistoryLimit > 10000) configHistoryLimit = 10000;
    }

    RegCloseKey(hKey);
    return TRUE;
}

void SaveConfigToRegistry() {
    HKEY hKey;
    DWORD disposition;
    LONG result = RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, &disposition);
    if (result != ERROR_SUCCESS) {
        LogMessage("ERROR: Failed to create/open registry key. Error: %lu", result);
        return;
    }

    // Write ApiUrl (REG_SZ)
    RegSetValueExA(hKey, REG_VALUE_URL, 0, REG_SZ,
                   (const BYTE*)configApiUrl, (DWORD)(strlen(configApiUrl) + 1));

    // Write RefreshInterval (REG_DWORD)
    DWORD dwInterval = (DWORD)configRefreshInterval;
    RegSetValueExA(hKey, REG_VALUE_INTERVAL, 0, REG_DWORD,
                   (const BYTE*)&dwInterval, sizeof(dwInterval));

    // Write LoggingEnabled (REG_DWORD)
    DWORD dwLogging = (DWORD)configLoggingEnabled;
    RegSetValueExA(hKey, REG_VALUE_LOGGING, 0, REG_DWORD,
                   (const BYTE*)&dwLogging, sizeof(dwLogging));

    // Write HistoryLimit (REG_DWORD)
    DWORD dwHistoryLimit = (DWORD)configHistoryLimit;
    RegSetValueExA(hKey, REG_VALUE_HISTORY_LIMIT, 0, REG_DWORD,
                   (const BYTE*)&dwHistoryLimit, sizeof(dwHistoryLimit));

    RegCloseKey(hKey);
    LogMessage("Configuration saved to registry: URL=%s, Interval=%d, Logging=%s, HistoryLimit=%d",
               configApiUrl, configRefreshInterval, configLoggingEnabled ? "enabled" : "disabled", configHistoryLimit);
}

BOOL IsFirstLaunch() {
    HKEY hKey;
    LONG result = RegOpenKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) {
        return TRUE;
    }

    DWORD type, size;
    DWORD dwConfigured = 0;
    size = sizeof(dwConfigured);
    result = RegQueryValueExA(hKey, REG_VALUE_CONFIGURED, NULL, &type, (LPBYTE)&dwConfigured, &size);
    RegCloseKey(hKey);

    if (result != ERROR_SUCCESS || type != REG_DWORD || dwConfigured == 0) {
        return TRUE;
    }
    return FALSE;
}

void MarkAsConfigured() {
    HKEY hKey;
    DWORD disposition;
    LONG result = RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, &disposition);
    if (result != ERROR_SUCCESS) {
        LogMessage("ERROR: Failed to create/open registry key for MarkAsConfigured. Error: %lu", result);
        return;
    }

    DWORD dwConfigured = 1;
    RegSetValueExA(hKey, REG_VALUE_CONFIGURED, 0, REG_DWORD,
                   (const BYTE*)&dwConfigured, sizeof(dwConfigured));
    RegCloseKey(hKey);
}

// INI fallback for one-time migration from old config.ini
void LoadConfigFromIni(const char* iniPath) {
    GetPrivateProfileStringA("General", "ApiUrl", "http://example.com/api/status",
                            configApiUrl, sizeof(configApiUrl), iniPath);
    configRefreshInterval = GetPrivateProfileIntA("General", "RefreshInterval", 60, iniPath);
    configLoggingEnabled = (BOOL)GetPrivateProfileIntA("General", "LoggingEnabled", 1, iniPath);
}

// --- History ring buffer ---

const char* ApiResultToString(ApiResult r) {
    switch (r) {
        case RESULT_SUCCESS: return "Success";
        case RESULT_FAIL:    return "Fail";
        case RESULT_ERROR:   return "Error";
        case RESULT_INVALID: return "Invalid";
        default:             return "Unknown";
    }
}

void InitHistoryBuffer(int capacity) {
    if (capacity < 10) capacity = 10;
    if (capacity > 10000) capacity = 10000;

    if (historyBuffer && capacity == historyCapacity) return;

    HistoryEntry* newBuf = (HistoryEntry*)calloc(capacity, sizeof(HistoryEntry));
    if (!newBuf) return;

    // Preserve most recent entries on resize
    if (historyBuffer && historyCount > 0) {
        int toCopy = historyCount < capacity ? historyCount : capacity;
        for (int i = 0; i < toCopy; i++) {
            // GetHistoryEntry(0) = most recent, so copy in reverse display order
            HistoryEntry* src = GetHistoryEntry(i);
            if (src) {
                newBuf[(toCopy - 1 - i) % capacity] = *src;
            }
        }
        historyHead = toCopy % capacity;
        historyCount = toCopy;
    } else {
        historyHead = 0;
        historyCount = 0;
    }

    free(historyBuffer);
    historyBuffer = newBuf;
    historyCapacity = capacity;
}

void AddHistoryEntry(ApiResult oldResult, const char* oldMsg, ApiResult newResult, const char* newMsg) {
    if (!historyBuffer || historyCapacity <= 0) return;

    HistoryEntry* entry = &historyBuffer[historyHead];
    GetLocalTime(&entry->timestamp);
    entry->oldResult = oldResult;
    entry->newResult = newResult;
    strncpy(entry->oldMessage, oldMsg ? oldMsg : "", sizeof(entry->oldMessage) - 1);
    entry->oldMessage[sizeof(entry->oldMessage) - 1] = '\0';
    strncpy(entry->newMessage, newMsg ? newMsg : "", sizeof(entry->newMessage) - 1);
    entry->newMessage[sizeof(entry->newMessage) - 1] = '\0';

    historyHead = (historyHead + 1) % historyCapacity;
    if (historyCount < historyCapacity) historyCount++;
}

HistoryEntry* GetHistoryEntry(int displayIndex) {
    if (!historyBuffer || displayIndex < 0 || displayIndex >= historyCount) return NULL;
    // displayIndex 0 = most recent
    int bufIdx = (historyHead - 1 - displayIndex + historyCapacity) % historyCapacity;
    return &historyBuffer[bufIdx];
}

void FreeHistoryBuffer(void) {
    free(historyBuffer);
    historyBuffer = NULL;
    historyCapacity = 0;
    historyCount = 0;
    historyHead = 0;
}

void SaveHistoryToRegistry(void) {
    HKEY hKey;
    DWORD disposition;
    LONG result = RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, &disposition);
    if (result != ERROR_SUCCESS) {
        LogMessage("ERROR: Failed to open registry for history save. Error: %lu", result);
        return;
    }

    if (historyCount == 0 || !historyBuffer) {
        DWORD zero = 0;
        RegSetValueExA(hKey, REG_VALUE_HISTORY_COUNT, 0, REG_DWORD,
                       (const BYTE*)&zero, sizeof(zero));
        RegDeleteValueA(hKey, REG_VALUE_HISTORY_DATA);
        RegCloseKey(hKey);
        LogMessage("History saved to registry: 0 entries.");
        return;
    }

    // Serialize most-recent-first
    DWORD dataSize = (DWORD)(historyCount * sizeof(HistoryEntry));
    BYTE* data = (BYTE*)malloc(dataSize);
    if (!data) {
        RegCloseKey(hKey);
        return;
    }

    for (int i = 0; i < historyCount; i++) {
        HistoryEntry* entry = GetHistoryEntry(i); // 0 = most recent
        if (entry) {
            memcpy(data + i * sizeof(HistoryEntry), entry, sizeof(HistoryEntry));
        }
    }

    DWORD dwCount = (DWORD)historyCount;
    RegSetValueExA(hKey, REG_VALUE_HISTORY_COUNT, 0, REG_DWORD,
                   (const BYTE*)&dwCount, sizeof(dwCount));
    RegSetValueExA(hKey, REG_VALUE_HISTORY_DATA, 0, REG_BINARY,
                   data, dataSize);

    free(data);
    RegCloseKey(hKey);
    LogMessage("History saved to registry: %d entries.", historyCount);
}

void LoadHistoryFromRegistry(void) {
    HKEY hKey;
    LONG result = RegOpenKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) return;

    DWORD type, size;

    // Read count
    DWORD dwCount = 0;
    size = sizeof(dwCount);
    if (RegQueryValueExA(hKey, REG_VALUE_HISTORY_COUNT, NULL, &type, (LPBYTE)&dwCount, &size) != ERROR_SUCCESS
        || type != REG_DWORD || dwCount == 0) {
        RegCloseKey(hKey);
        return;
    }

    // Read data size first
    size = 0;
    if (RegQueryValueExA(hKey, REG_VALUE_HISTORY_DATA, NULL, &type, NULL, &size) != ERROR_SUCCESS
        || type != REG_BINARY) {
        RegCloseKey(hKey);
        return;
    }

    // Validate size matches count
    if (size != dwCount * sizeof(HistoryEntry)) {
        LogMessage("WARNING: History data size mismatch (expected %lu, got %lu). Discarding.",
                   (unsigned long)(dwCount * sizeof(HistoryEntry)), (unsigned long)size);
        RegCloseKey(hKey);
        return;
    }

    BYTE* data = (BYTE*)malloc(size);
    if (!data) {
        RegCloseKey(hKey);
        return;
    }

    if (RegQueryValueExA(hKey, REG_VALUE_HISTORY_DATA, NULL, &type, data, &size) != ERROR_SUCCESS) {
        free(data);
        RegCloseKey(hKey);
        return;
    }

    RegCloseKey(hKey);

    // Cap to current buffer capacity
    int toLoad = (int)dwCount;
    if (toLoad > historyCapacity) toLoad = historyCapacity;

    // Data is stored most-recent-first; insert oldest-first so ring buffer order is correct
    for (int i = toLoad - 1; i >= 0; i--) {
        HistoryEntry* src = (HistoryEntry*)(data + i * sizeof(HistoryEntry));
        HistoryEntry* dst = &historyBuffer[historyHead];
        *dst = *src;
        historyHead = (historyHead + 1) % historyCapacity;
        if (historyCount < historyCapacity) historyCount++;
    }

    free(data);
    LogMessage("History loaded from registry: %d entries.", toLoad);
}

void ApplyConfiguration() {
    if (g_hwnd) {
        if (timerRefresh) KillTimer(g_hwnd, 1);
        timerRefresh = SetTimer(g_hwnd, 1, configRefreshInterval * 1000, RefreshTimer);
    }
    LogMessage("Configuration applied: URL=%s, Interval=%d, Logging=%s",
               configApiUrl, configRefreshInterval, configLoggingEnabled ? "enabled" : "disabled");
}

// --- Configuration dialog ---

// Helper to add a control to dialog template
static BYTE* AddDialogControl(BYTE* ptr, WORD ctrlId, WORD classAtom, DWORD style,
                               short posX, short posY, short width, short height,
                               const wchar_t* text) {
    // Align to DWORD
    ptr = (BYTE*)(((ULONG_PTR)ptr + 3) & ~3);

    DLGITEMTEMPLATE* item = (DLGITEMTEMPLATE*)ptr;
    item->style = style;
    item->dwExtendedStyle = 0;
    item->x = posX;
    item->y = posY;
    item->cx = width;
    item->cy = height;
    item->id = ctrlId;
    ptr += sizeof(DLGITEMTEMPLATE);

    // Class (atom)
    *(WORD*)ptr = 0xFFFF;
    ptr += sizeof(WORD);
    *(WORD*)ptr = classAtom;
    ptr += sizeof(WORD);

    // Text
    size_t textLen = wcslen(text) + 1;
    memcpy(ptr, text, textLen * sizeof(wchar_t));
    ptr += textLen * sizeof(wchar_t);

    // Creation data (none)
    *(WORD*)ptr = 0;
    ptr += sizeof(WORD);

    return ptr;
}

// Helper: launch a validation thread for the given URL
static void StartValidation(HWND hDlg, const char* url) {
    LONG gen = InterlockedIncrement(&g_validateGeneration);
    ValidateParams* vp = (ValidateParams*)malloc(sizeof(ValidateParams));
    if (!vp) return;
    strncpy(vp->url, url, sizeof(vp->url) - 1);
    vp->url[sizeof(vp->url) - 1] = '\0';
    vp->hDlg = hDlg;
    vp->generation = gen;
    HANDLE hThread = CreateThread(NULL, 0, ValidateUrlThread, vp, 0, NULL);
    if (hThread) CloseHandle(hThread);
    else free(vp);
}

// Configuration dialog procedure
static INT_PTR CALLBACK ConfigDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    // 0=none, 1=checking, 2=valid, 3=invalid
    static int validationState = 0;

    switch (message) {
        case WM_INITDIALOG: {
            validationState = 0;

            // Set URL
            SetDlgItemTextA(hDlg, IDC_EDIT_URL, configApiUrl);

            // Populate interval combo
            HWND hCombo = GetDlgItem(hDlg, IDC_COMBO_INTERVAL);
            SendMessageA(hCombo, CB_ADDSTRING, 0, (LPARAM)"Every 1 minute");
            SendMessageA(hCombo, CB_ADDSTRING, 0, (LPARAM)"Every 2 minutes");
            SendMessageA(hCombo, CB_ADDSTRING, 0, (LPARAM)"Every 5 minutes");

            // Select current interval
            int sel = 0;
            if (configRefreshInterval == 120) sel = 1;
            else if (configRefreshInterval == 300) sel = 2;
            SendMessageA(hCombo, CB_SETCURSEL, (WPARAM)sel, 0);

            // Set logging checkbox
            CheckDlgButton(hDlg, IDC_CHECK_LOGGING, configLoggingEnabled ? BST_CHECKED : BST_UNCHECKED);

            // Set history limit
            SetDlgItemInt(hDlg, IDC_EDIT_HISTORY_LIMIT, (UINT)configHistoryLimit, FALSE);

            // Center dialog on screen
            RECT rcDlg, rcScreen;
            GetWindowRect(hDlg, &rcDlg);
            SystemParametersInfoW(SPI_GETWORKAREA, 0, &rcScreen, 0);
            int x = rcScreen.left + ((rcScreen.right - rcScreen.left) - (rcDlg.right - rcDlg.left)) / 2;
            int y = rcScreen.top + ((rcScreen.bottom - rcScreen.top) - (rcDlg.bottom - rcDlg.top)) / 2;
            SetWindowPos(hDlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

            // If URL is non-empty, trigger immediate validation
            if (configApiUrl[0] != '\0') {
                SetDlgItemTextA(hDlg, IDC_STATIC_VALID_STATUS, "Checking...");
                validationState = 1;
                InvalidateRect(GetDlgItem(hDlg, IDC_STATIC_VALID_STATUS), NULL, TRUE);
                StartValidation(hDlg, configApiUrl);
            }

            return TRUE;
        }

        case WM_COMMAND:
            if (HIWORD(wParam) == EN_CHANGE && LOWORD(wParam) == IDC_EDIT_URL) {
                char url[512];
                GetDlgItemTextA(hDlg, IDC_EDIT_URL, url, sizeof(url));

                // Trim whitespace for check
                char* p = url;
                while (*p && isspace((unsigned char)*p)) p++;

                if (*p == '\0') {
                    // Empty URL
                    SetDlgItemTextA(hDlg, IDC_STATIC_VALID_STATUS, "");
                    validationState = 0;
                    KillTimer(hDlg, IDT_VALIDATE_DEBOUNCE);
                    InterlockedIncrement(&g_validateGeneration); // cancel pending
                } else {
                    SetDlgItemTextA(hDlg, IDC_STATIC_VALID_STATUS, "Checking...");
                    validationState = 1;
                    KillTimer(hDlg, IDT_VALIDATE_DEBOUNCE);
                    SetTimer(hDlg, IDT_VALIDATE_DEBOUNCE, 500, NULL);
                }
                InvalidateRect(GetDlgItem(hDlg, IDC_STATIC_VALID_STATUS), NULL, TRUE);
                return TRUE;
            }
            switch (LOWORD(wParam)) {
                case IDOK: {
                    KillTimer(hDlg, IDT_VALIDATE_DEBOUNCE);

                    // Get URL
                    char url[512];
                    GetDlgItemTextA(hDlg, IDC_EDIT_URL, url, sizeof(url));

                    // Trim whitespace
                    char* p = url;
                    while (*p && isspace((unsigned char)*p)) p++;
                    if (p != url) memmove(url, p, strlen(p) + 1);
                    size_t len = strlen(url);
                    while (len > 0 && isspace((unsigned char)url[len - 1])) url[--len] = '\0';

                    // Validate URL is not empty
                    if (url[0] == '\0') {
                        MessageBoxA(hDlg, "API URL cannot be empty.", "Validation Error", MB_OK | MB_ICONWARNING);
                        SetFocus(GetDlgItem(hDlg, IDC_EDIT_URL));
                        return TRUE;
                    }
                    strncpy(configApiUrl, url, sizeof(configApiUrl) - 1);
                    configApiUrl[sizeof(configApiUrl) - 1] = '\0';

                    // Get interval
                    HWND hCombo = GetDlgItem(hDlg, IDC_COMBO_INTERVAL);
                    int sel = (int)SendMessageA(hCombo, CB_GETCURSEL, 0, 0);
                    int intervals[] = {60, 120, 300};
                    configRefreshInterval = intervals[sel >= 0 && sel < 3 ? sel : 0];

                    // Get logging
                    configLoggingEnabled = (IsDlgButtonChecked(hDlg, IDC_CHECK_LOGGING) == BST_CHECKED);

                    // Get history limit
                    {
                        BOOL translated = FALSE;
                        UINT hlVal = GetDlgItemInt(hDlg, IDC_EDIT_HISTORY_LIMIT, &translated, FALSE);
                        if (translated) {
                            if (hlVal < 10) hlVal = 10;
                            if (hlVal > 10000) hlVal = 10000;
                            configHistoryLimit = (int)hlVal;
                            InitHistoryBuffer(configHistoryLimit);
                        }
                    }

                    SaveConfigToRegistry();
                    MarkAsConfigured();
                    ApplyConfiguration();

                    LogMessage("Configuration updated via dialog: URL=%s, Interval=%d, Logging=%s",
                               configApiUrl, configRefreshInterval, configLoggingEnabled ? "enabled" : "disabled");

                    EndDialog(hDlg, IDOK);
                    return TRUE;
                }

                case IDCANCEL:
                    KillTimer(hDlg, IDT_VALIDATE_DEBOUNCE);
                    EndDialog(hDlg, IDCANCEL);
                    return TRUE;
            }
            break;

        case WM_TIMER:
            if (wParam == IDT_VALIDATE_DEBOUNCE) {
                KillTimer(hDlg, IDT_VALIDATE_DEBOUNCE);

                char url[512];
                GetDlgItemTextA(hDlg, IDC_EDIT_URL, url, sizeof(url));

                // Trim whitespace
                char* p = url;
                while (*p && isspace((unsigned char)*p)) p++;
                char* end = p + strlen(p) - 1;
                while (end > p && isspace((unsigned char)*end)) *end-- = '\0';

                if (*p != '\0') {
                    StartValidation(hDlg, p);
                }
                return TRUE;
            }
            break;

        case WM_VALIDATE_RESULT:
            if ((LONG)wParam == g_validateGeneration) {
                if (lParam) {
                    SetDlgItemTextA(hDlg, IDC_STATIC_VALID_STATUS, "Valid");
                    validationState = 2;
                } else {
                    SetDlgItemTextA(hDlg, IDC_STATIC_VALID_STATUS, "Invalid");
                    validationState = 3;
                }
                InvalidateRect(GetDlgItem(hDlg, IDC_STATIC_VALID_STATUS), NULL, TRUE);
            }
            return TRUE;

        case WM_CTLCOLORSTATIC: {
            HWND hCtrl = (HWND)lParam;
            if (hCtrl == GetDlgItem(hDlg, IDC_STATIC_VALID_STATUS)) {
                HDC hdc = (HDC)wParam;
                if (validationState == 2)
                    SetTextColor(hdc, RGB(0, 128, 0));
                else if (validationState == 3)
                    SetTextColor(hdc, RGB(192, 0, 0));
                else if (validationState == 1)
                    SetTextColor(hdc, RGB(128, 128, 128));
                SetBkColor(hdc, GetSysColor(COLOR_BTNFACE));
                return (INT_PTR)GetSysColorBrush(COLOR_BTNFACE);
            }
            break;
        }

        case WM_CLOSE:
            KillTimer(hDlg, IDT_VALIDATE_DEBOUNCE);
            EndDialog(hDlg, IDCANCEL);
            return TRUE;
    }
    return FALSE;
}

// Validation thread: makes HTTP GET and posts result back to dialog
DWORD WINAPI ValidateUrlThread(LPVOID param) {
    ValidateParams* vp = (ValidateParams*)param;
    HWND hDlg = vp->hDlg;
    LONG myGen = vp->generation;
    LRESULT valid = 0;

    // Parse URL
    char urlCopy[512];
    strncpy(urlCopy, vp->url, sizeof(urlCopy) - 1);
    urlCopy[sizeof(urlCopy) - 1] = '\0';

    char* host = NULL;
    char* path = NULL;
    int port = 80;
    BOOL isHttps = FALSE;

    if (strncmp(urlCopy, "http://", 7) == 0) {
        host = urlCopy + 7;
    } else if (strncmp(urlCopy, "https://", 8) == 0) {
        host = urlCopy + 8;
        port = 443;
        isHttps = TRUE;
    } else {
        host = urlCopy;
    }

    char* slash = strchr(host, '/');
    if (slash) {
        *slash = '\0';
        path = slash + 1;
    } else {
        path = "";
    }

    char* colon = strchr(host, ':');
    if (colon) {
        *colon = '\0';
        port = atoi(colon + 1);
    }

    wchar_t wHost[256], wPath[512];
    MultiByteToWideChar(CP_UTF8, 0, host, -1, wHost, 256);
    char fullPath[512] = "/";
    if (strlen(path) > 0) {
        sprintf(fullPath, "/%s", path);
    }
    MultiByteToWideChar(CP_UTF8, 0, fullPath, -1, wPath, 512);

    HINTERNET hSession = NULL, hConnect = NULL, hRequest = NULL;

    hSession = WinHttpOpen(L"APIMonitor/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                          WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) goto cleanup;

    hConnect = WinHttpConnect(hSession, wHost, (INTERNET_PORT)port, 0);
    if (!hConnect) goto cleanup;

    {
        DWORD flags = isHttps ? WINHTTP_FLAG_SECURE : 0;
        hRequest = WinHttpOpenRequest(hConnect, L"GET", wPath,
                                     NULL, WINHTTP_NO_REFERER,
                                     WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    }
    if (!hRequest) goto cleanup;

    // 5 second timeouts
    {
        int timeout = 5000;
        WinHttpSetOption(hRequest, WINHTTP_OPTION_CONNECT_TIMEOUT, &timeout, sizeof(timeout));
        WinHttpSetOption(hRequest, WINHTTP_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));
        WinHttpSetOption(hRequest, WINHTTP_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
    }

    if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                           WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
        goto cleanup;

    if (!WinHttpReceiveResponse(hRequest, NULL))
        goto cleanup;

    {
        DWORD statusCode = 0, size = sizeof(statusCode);
        if (!WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                NULL, &statusCode, &size, NULL))
            goto cleanup;
        if (statusCode != 200)
            goto cleanup;
    }

    // Read response body
    {
        char response[4096] = {0};
        DWORD totalSize = 0, downloaded = 0;

        do {
            DWORD sizeAvailable = 0;
            if (!WinHttpQueryDataAvailable(hRequest, &sizeAvailable)) break;
            if (sizeAvailable == 0) break;

            char* buffer = malloc(sizeAvailable + 1);
            if (!buffer) break;
            if (!WinHttpReadData(hRequest, buffer, sizeAvailable, &downloaded)) {
                free(buffer);
                break;
            }
            buffer[downloaded] = '\0';

            if (totalSize + downloaded < sizeof(response) - 1) {
                strcat(response, buffer);
                totalSize += downloaded;
            }
            free(buffer);
        } while (downloaded > 0);

        // Check for expected XML tags
        if (strstr(response, "<result") || strstr(response, "<r>")) {
            valid = 1;
        }
    }

cleanup:
    if (hRequest) WinHttpCloseHandle(hRequest);
    if (hConnect) WinHttpCloseHandle(hConnect);
    if (hSession) WinHttpCloseHandle(hSession);

    // Only post result if this generation is still current
    if (InterlockedCompareExchange(&g_validateGeneration, myGen, myGen) == myGen) {
        PostMessage(hDlg, WM_VALIDATE_RESULT, (WPARAM)myGen, valid);
    }

    free(vp);
    return 0;
}

// Create and show configuration dialog
void ShowConfigDialog(HWND hwndParent) {
    // Dialog dimensions (in dialog units)
    const short DLG_WIDTH = 320;
    const short MARGIN_X = 8;
    const short MARGIN_Y = 8;
    const short LABEL_H = 10;
    const short LABEL_GAP = 2;
    const short EDIT_H = 14;
    const short SPACING = 6;
    const short BTN_W = 50;
    const short BTN_H = 14;
    const short COMBO_DROPDOWN = 60;

    // Calculate dialog height based on contents
    // 1 single-line field + validation row + 1 combo field + 1 checkbox + 1 history limit field + buttons
    const short DLG_HEIGHT = MARGIN_Y
        + (LABEL_H + LABEL_GAP + EDIT_H + SPACING)      // API URL
        + (LABEL_H + SPACING)                            // API Valid status row
        + (LABEL_H + LABEL_GAP + EDIT_H + SPACING)      // Interval (combo visible height = EDIT_H)
        + (EDIT_H + SPACING)                             // Checkbox
        + (LABEL_H + LABEL_GAP + EDIT_H + SPACING)      // History Limit
        + BTN_H + MARGIN_Y;                              // Buttons + bottom margin

    short editW = DLG_WIDTH - (2 * MARGIN_X);
    short yPos = MARGIN_Y;

    // Allocate buffer for dialog template
    size_t templateSize = 4096;
    BYTE* templateBuffer = (BYTE*)calloc(1, templateSize);
    if (!templateBuffer) return;

    BYTE* ptr = templateBuffer;

    // DLGTEMPLATE
    DLGTEMPLATE* dlgTemplate = (DLGTEMPLATE*)ptr;
    dlgTemplate->style = DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_SETFONT;
    dlgTemplate->dwExtendedStyle = 0;
    dlgTemplate->cdit = 11;  // 2 labels + 1 edit + 2 valid statics + 1 combo + 1 checkbox + 1 history label + 1 history edit + 2 buttons
    dlgTemplate->x = 0;
    dlgTemplate->y = 0;
    dlgTemplate->cx = DLG_WIDTH;
    dlgTemplate->cy = DLG_HEIGHT;
    ptr += sizeof(DLGTEMPLATE);

    // Menu (none)
    *(WORD*)ptr = 0;
    ptr += sizeof(WORD);

    // Class (default)
    *(WORD*)ptr = 0;
    ptr += sizeof(WORD);

    // Title
    const wchar_t* title = L"Configuration";
    size_t titleLen = wcslen(title) + 1;
    memcpy(ptr, title, titleLen * sizeof(wchar_t));
    ptr += titleLen * sizeof(wchar_t);

    // Font size
    *(WORD*)ptr = 8;
    ptr += sizeof(WORD);

    // Font name
    const wchar_t* fontName = L"Segoe UI";
    size_t fontLen = wcslen(fontName) + 1;
    memcpy(ptr, fontName, fontLen * sizeof(wchar_t));
    ptr += fontLen * sizeof(wchar_t);

    // API URL label
    ptr = AddDialogControl(ptr, IDC_STATIC_URL, 0x0082, WS_CHILD | WS_VISIBLE | SS_LEFT,
                           MARGIN_X, yPos, editW, LABEL_H, L"API URL:");
    yPos += LABEL_H + LABEL_GAP;

    // API URL edit
    ptr = AddDialogControl(ptr, IDC_EDIT_URL, 0x0081,
                           WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL,
                           MARGIN_X, yPos, editW, EDIT_H, L"");
    yPos += EDIT_H + SPACING;

    // API Valid label
    ptr = AddDialogControl(ptr, IDC_STATIC_VALID_LABEL, 0x0082, WS_CHILD | WS_VISIBLE | SS_LEFT,
                           MARGIN_X, yPos, 38, LABEL_H, L"API Valid:");

    // API Valid status
    ptr = AddDialogControl(ptr, IDC_STATIC_VALID_STATUS, 0x0082, WS_CHILD | WS_VISIBLE | SS_LEFT,
                           MARGIN_X + 40, yPos, 60, LABEL_H, L"");
    yPos += LABEL_H + SPACING;

    // Interval label
    ptr = AddDialogControl(ptr, IDC_STATIC_INTERVAL, 0x0082, WS_CHILD | WS_VISIBLE | SS_LEFT,
                           MARGIN_X, yPos, editW, LABEL_H, L"Check Interval:");
    yPos += LABEL_H + LABEL_GAP;

    // Interval combo (cy includes dropdown height)
    ptr = AddDialogControl(ptr, IDC_COMBO_INTERVAL, 0x0085,
                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                           MARGIN_X, yPos, 130, COMBO_DROPDOWN, L"");
    yPos += EDIT_H + SPACING;

    // Enable Logging checkbox
    ptr = AddDialogControl(ptr, IDC_CHECK_LOGGING, 0x0080,
                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                           MARGIN_X, yPos, editW, EDIT_H, L"Enable Logging");
    yPos += EDIT_H + SPACING;

    // History Limit label
    ptr = AddDialogControl(ptr, IDC_STATIC_HISTORY, 0x0082, WS_CHILD | WS_VISIBLE | SS_LEFT,
                           MARGIN_X, yPos, editW, LABEL_H, L"History Limit (10-10000):");
    yPos += LABEL_H + LABEL_GAP;

    // History Limit edit
    ptr = AddDialogControl(ptr, IDC_EDIT_HISTORY_LIMIT, 0x0081,
                           WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_NUMBER,
                           MARGIN_X, yPos, 80, EDIT_H, L"");
    yPos += EDIT_H + SPACING;

    // OK button
    short okX = DLG_WIDTH - MARGIN_X - BTN_W - 4 - BTN_W - 4;
    ptr = AddDialogControl(ptr, IDOK, 0x0080,
                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                           okX, yPos, BTN_W, BTN_H, L"OK");

    // Cancel button
    short cancelX = DLG_WIDTH - MARGIN_X - BTN_W - 4;
    ptr = AddDialogControl(ptr, IDCANCEL, 0x0080,
                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                           cancelX, yPos, BTN_W, BTN_H, L"Cancel");

    DialogBoxIndirectParamW(g_hInstance, (DLGTEMPLATE*)templateBuffer,
                            hwndParent, ConfigDialogProc, 0);
    free(templateBuffer);
}

// --- Timer and refresh ---

void SetRefreshInterval(int seconds, BOOL isUserSetting) {
    if (isUserSetting) {
        configRefreshInterval = seconds;
        SaveConfigToRegistry();

        if (currentResult != RESULT_SUCCESS) {
            LogMessage("Interval change to %d seconds requested, but not applied due to non-success state.", seconds);
            return;
        }
    }

    if (timerRefresh) KillTimer(g_hwnd, 1);
    timerRefresh = SetTimer(g_hwnd, 1, seconds * 1000, RefreshTimer);
    LogMessage("Refresh timer updated: %d seconds (userSetting=%s)", seconds, isUserSetting ? "true" : "false");
}

void RefreshStatus() {
    LogMessage("RefreshStatus() called.");
    DWORD threadId;

    // Allocate thread parameters
    ThreadParams* params = (ThreadParams*)malloc(sizeof(ThreadParams));
    if (!params) {
        LogMessage("ERROR: Failed to allocate memory for thread parameters");
        return;
    }
    params->attempt = 1;
    params->maxAttempts = 3;

    // Create thread with parameters
    CloseHandle(CreateThread(NULL, 0, RefreshThread, params, 0, &threadId));
}

void ParseXmlResponse(const char* xml, ApiResponse* response) {
    response->result = RESULT_INVALID;
    strncpy(response->message, "Invalid XML", sizeof(response->message) - 1);

    if (!xml || strlen(xml) == 0) {
        LogMessage("ERROR: Empty XML response received.");
        strncpy(response->message, "Empty response", sizeof(response->message) - 1);
        return;
    }

    const char* resultTag = strstr(xml, "<r>");
    const char* resultEnd = NULL;
    const char* valueStart = NULL;
    BOOL isShortTag = FALSE;

    if (resultTag) {
        isShortTag = TRUE;
        resultEnd = strstr(resultTag, "</r>");
        if (!resultEnd) {
            LogMessage("ERROR: Invalid XML - unclosed <r> tag. Raw: %.100s", xml);
            strncpy(response->message, "Unclosed <r> tag", sizeof(response->message) - 1);
            return;
        }
        valueStart = resultTag + 3;
    } else {
        resultTag = strstr(xml, "<result");
        if (!resultTag) {
            LogMessage("ERROR: Invalid XML - no <result> tag found. Raw: %.100s", xml);
            strncpy(response->message, "No result tag", sizeof(response->message) - 1);
            return;
        }

        resultEnd = strstr(resultTag, "</result>");
        if (!resultEnd) {
            LogMessage("ERROR: Invalid XML - unclosed <result> tag. Raw: %.100s", xml);
            strncpy(response->message, "Unclosed <result> tag", sizeof(response->message) - 1);
            return;
        }

        const char* closeBracket = strchr(resultTag, '>');
        if (!closeBracket || closeBracket >= resultEnd) {
            LogMessage("ERROR: Invalid XML - malformed <result> tag. Raw: %.100s", xml);
            strncpy(response->message, "Malformed <result> tag", sizeof(response->message) - 1);
            return;
        }
        valueStart = closeBracket + 1;
    }

    char resultValue[32] = {0};
    size_t valueLen = resultEnd - valueStart;
    if (valueLen >= sizeof(resultValue)) valueLen = sizeof(resultValue) - 1;
    strncpy(resultValue, valueStart, valueLen);
    resultValue[valueLen] = '\0';

    const char* msgTag = strstr(xml, "<message>");
    if (msgTag) {
        const char* msgEnd = strstr(msgTag, "</message>");
        if (msgEnd) {
            msgTag += 9;
            size_t msgLen = msgEnd - msgTag;
            if (msgLen >= sizeof(response->message)) msgLen = sizeof(response->message) - 1;
            strncpy(response->message, msgTag, msgLen);
            response->message[msgLen] = '\0';

            char* p = response->message;
            while (*p && isspace((unsigned char)*p)) p++;
            char* end = p + strlen(p) - 1;
            while (end > p && isspace((unsigned char)*end)) *end-- = '\0';
        }
    }

    char* p = resultValue;
    while (*p && isspace((unsigned char)*p)) p++;
    char* end = p + strlen(p) - 1;
    while (end > p && isspace((unsigned char)*end)) *end-- = '\0';

    for (char* c = p; *c; c++) {
        *c = tolower((unsigned char)*c);
    }

    if (strcmp(p, "success") == 0) {
        response->result = RESULT_SUCCESS;
        LogMessage("XML parsed successfully: result=success, message=%s", response->message);
    } else if (strcmp(p, "fail") == 0) {
        response->result = RESULT_FAIL;
        LogMessage("XML parsed: result=fail, message=%s", response->message);
    } else {
        LogMessage("ERROR: Invalid XML - unknown result value '%s'. Raw: %.100s", p, xml);
    }
}

DWORD WINAPI RefreshThread(LPVOID param) {
    ThreadParams* params = (ThreadParams*)param;
    const int maxAttempts = params->maxAttempts;
    ApiResult finalResult = RESULT_ERROR;
    char finalMessage[256] = "Unknown error";
    char response[4096] = {0};

    // Validate API URL
    if (strlen(configApiUrl) == 0) {
        LogMessage("ERROR: API URL is not configured.");
        strncpy(finalMessage, "API URL not configured", sizeof(finalMessage) - 1);
        UpdateStatus(RESULT_ERROR, finalMessage);
        free(params);
        return 0;
    }

    // Retry loop
    for (int attempt = params->attempt; attempt <= maxAttempts; attempt++) {
        // Update tooltip with attempt count
        char tip[128];
        snprintf(tip, sizeof(tip), "Updating API contents [%d/%d]...", attempt, maxAttempts);
        strcpy(nid.szTip, tip);
        nid.uFlags = NIF_TIP;
        Shell_NotifyIconA(NIM_MODIFY, &nid);

        LogMessage("API refresh attempt %d/%d started.", attempt, maxAttempts);

        // Reset response buffer
        response[0] = '\0';

        // Parse URL (existing logic)
        char urlCopy[512];
        strncpy(urlCopy, configApiUrl, sizeof(urlCopy) - 1);
        urlCopy[sizeof(urlCopy) - 1] = '\0';

        char* protocol = NULL;
        char* host = NULL;
        char* path = NULL;
        int port = 80;
        BOOL isHttps = FALSE;

        if (strncmp(urlCopy, "http://", 7) == 0) {
            protocol = urlCopy;
            host = urlCopy + 7;
        } else if (strncmp(urlCopy, "https://", 8) == 0) {
            protocol = urlCopy;
            host = urlCopy + 8;
            port = 443;
            isHttps = TRUE;
        } else {
            host = urlCopy;
        }

        char* slash = strchr(host, '/');
        if (slash) {
            *slash = '\0';
            path = slash + 1;
        } else {
            path = "";
        }

        char* colon = strchr(host, ':');
        if (colon) {
            *colon = '\0';
            port = atoi(colon + 1);
        }

        wchar_t wHost[256], wPath[512];
        MultiByteToWideChar(CP_UTF8, 0, host, -1, wHost, 256);
        char fullPath[512] = "/";
        if (strlen(path) > 0) {
            sprintf(fullPath, "/%s", path);
        }
        MultiByteToWideChar(CP_UTF8, 0, fullPath, -1, wPath, 512);

        // HTTP Request with error handling
        HINTERNET hSession = NULL;
        HINTERNET hConnect = NULL;
        HINTERNET hRequest = NULL;
        BOOL networkError = FALSE;
        char errorMsg[128] = "";

        // Create session
        hSession = WinHttpOpen(L"APIMonitor/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                              WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
        if (!hSession) {
            sprintf(errorMsg, "HTTP init failed: %lu", GetLastError());
            LogMessage("ERROR: %s (attempt %d/%d)", errorMsg, attempt, maxAttempts);
            networkError = TRUE;
        }

        // Connect
        if (!networkError) {
            hConnect = WinHttpConnect(hSession, wHost, (INTERNET_PORT)port, 0);
            if (!hConnect) {
                sprintf(errorMsg, "Connection failed: %lu", GetLastError());
                LogMessage("ERROR: %s (attempt %d/%d)", errorMsg, attempt, maxAttempts);
                networkError = TRUE;
            }
        }

        // Create request
        if (!networkError) {
            DWORD flags = isHttps ? WINHTTP_FLAG_SECURE : 0;
            hRequest = WinHttpOpenRequest(hConnect, L"GET", wPath,
                                         NULL, WINHTTP_NO_REFERER,
                                         WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
            if (!hRequest) {
                sprintf(errorMsg, "Request creation failed: %lu", GetLastError());
                LogMessage("ERROR: %s (attempt %d/%d)", errorMsg, attempt, maxAttempts);
                networkError = TRUE;
            }
        }

        // Set timeouts
        if (!networkError) {
            int timeout = 10000;
            WinHttpSetOption(hRequest, WINHTTP_OPTION_RECEIVE_TIMEOUT, &timeout, sizeof(timeout));
            WinHttpSetOption(hRequest, WINHTTP_OPTION_SEND_TIMEOUT, &timeout, sizeof(timeout));
        }

        // Send request
        if (!networkError) {
            if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                                   WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) {
                sprintf(errorMsg, "Request failed: %lu", GetLastError());
                LogMessage("ERROR: %s (attempt %d/%d)", errorMsg, attempt, maxAttempts);
                networkError = TRUE;
            }
        }

        // Receive response
        if (!networkError) {
            if (!WinHttpReceiveResponse(hRequest, NULL)) {
                sprintf(errorMsg, "No response: %lu", GetLastError());
                LogMessage("ERROR: %s (attempt %d/%d)", errorMsg, attempt, maxAttempts);
                networkError = TRUE;
            }
        }

        // Check status code (only if we got a response)
        DWORD statusCode = 0;
        if (!networkError) {
            DWORD size = sizeof(statusCode);
            if (WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                   NULL, &statusCode, &size, NULL)) {
                if (statusCode != 200) {
                    sprintf(errorMsg, "HTTP %lu", statusCode);
                    LogMessage("ERROR: Received %s (attempt %d/%d), not retrying.", errorMsg, attempt, maxAttempts);
                    finalResult = RESULT_ERROR;
                    strncpy(finalMessage, errorMsg, sizeof(finalMessage) - 1);
                    // Clean up handles
                    WinHttpCloseHandle(hRequest);
                    WinHttpCloseHandle(hConnect);
                    WinHttpCloseHandle(hSession);
                    break; // Don't retry on HTTP errors
                }
            }
        }

        // Read response data if no network error and status is 200
        if (!networkError && statusCode == 200) {
            DWORD totalSize = 0;
            DWORD downloaded = 0;

            do {
                DWORD sizeAvailable = 0;
                if (!WinHttpQueryDataAvailable(hRequest, &sizeAvailable)) break;
                if (sizeAvailable == 0) break;

                char* buffer = malloc(sizeAvailable + 1);
                if (!WinHttpReadData(hRequest, buffer, sizeAvailable, &downloaded)) {
                    free(buffer);
                    break;
                }
                buffer[downloaded] = '\0';

                if (totalSize + downloaded < sizeof(response) - 1) {
                    strcat(response, buffer);
                    totalSize += downloaded;
                }
                free(buffer);
            } while (downloaded > 0);

            LogMessage("API response received (attempt %d/%d): %.500s", attempt, maxAttempts, response);
        }

        // Clean up handles
        if (hRequest) WinHttpCloseHandle(hRequest);
        if (hConnect) WinHttpCloseHandle(hConnect);
        if (hSession) WinHttpCloseHandle(hSession);

        // Check if we should retry
        if (networkError) {
            if (attempt < maxAttempts) {
                LogMessage("Network error on attempt %d/%d - retrying in 2 seconds...", attempt, maxAttempts);
                Sleep(2000); // Wait before retry
                continue; // Retry loop
            } else {
                // All attempts exhausted
                strncpy(finalMessage, errorMsg, sizeof(finalMessage) - 1);
                finalResult = RESULT_ERROR;
                break;
            }
        }

        // Parse the XML response
        ApiResponse apiResponse = {0};
        ParseXmlResponse(response, &apiResponse);

        // If API returns "fail", don't retry further
        if (apiResponse.result == RESULT_FAIL) {
            LogMessage("API returned 'fail' on attempt %d/%d - no further retries.", attempt, maxAttempts);
            finalResult = RESULT_FAIL;
            strncpy(finalMessage, apiResponse.message, sizeof(finalMessage) - 1);
            break; // Exit retry loop
        }

        // For success or invalid XML, use the result and exit loop
        finalResult = apiResponse.result;
        strncpy(finalMessage, apiResponse.message, sizeof(finalMessage) - 1);
        break; // Success or non-retryable error
    }

    // Update the UI with final result
    UpdateStatus(finalResult, finalMessage);

    // Clean up parameters
    free(params);
    LogMessage("API refresh thread completed with result: %d", finalResult);
    return 0;
}

void UpdateStatus(ApiResult result, const char* message) {
    // Detect status changes and record in history
    BOOL resultChanged = (result != currentResult);
    BOOL messageChanged = (message && strcmp(currentMessage, message) != 0);
    if (resultChanged || (messageChanged && result != RESULT_SUCCESS)) {
        AddHistoryEntry(currentResult, currentMessage, result, message ? message : "");
    }

    currentResult = result;
    if (message) {
        strncpy(currentMessage, message, sizeof(currentMessage) - 1);
        currentMessage[sizeof(currentMessage) - 1] = '\0';
    }
    GetSystemTime(&lastUpdateTime);
    UpdateTooltip();

    switch (result) {
        case RESULT_SUCCESS:
            LogMessage("Status update: SUCCESS - %s", message ? message : "No message");
            SetIcon(hIconSuccess);
            SetRefreshInterval(configRefreshInterval, FALSE);
            break;
        case RESULT_FAIL:
            LogMessage("Status update: FAIL - %s", message ? message : "No message");
            SetIcon(hIconFail);
            SetRefreshInterval(10, FALSE);
            break;
        case RESULT_ERROR:
            LogMessage("Status update: ERROR - %s", message ? message : "No message");
            SetIcon(hIconEmpty);
            SetRefreshInterval(10, FALSE);
            break;
        case RESULT_INVALID:
            LogMessage("Status update: INVALID - %s", message ? message : "No message");
            SetIcon(hIconEmpty);
            SetRefreshInterval(10, FALSE);
            break;
    }
}

void UpdateTooltip() {
    char tooltip[128];
    if (currentResult == RESULT_ERROR) {
        strcpy(tooltip, "Unable to connect to API!");
    } else if (currentResult == RESULT_INVALID) {
        strcpy(tooltip, "API response incorrect!");
    } else {
        SYSTEMTIME now;
        GetSystemTime(&now);

        FILETIME ftLast, ftNow;
        SystemTimeToFileTime(&lastUpdateTime, &ftLast);
        SystemTimeToFileTime(&now, &ftNow);

        ULARGE_INTEGER last, current;
        last.LowPart = ftLast.dwLowDateTime;
        last.HighPart = ftLast.dwHighDateTime;
        current.LowPart = ftNow.dwLowDateTime;
        current.HighPart = ftNow.dwHighDateTime;

        ULONGLONG diff = (current.QuadPart - last.QuadPart) / 10000000;
        if (diff == 1) {
            snprintf(tooltip, sizeof(tooltip), "Updated %llu second ago", diff);
        } else {
            snprintf(tooltip, sizeof(tooltip), "Updated %llu seconds ago", diff);
        }

        if (strlen(currentMessage) > 0) {
            size_t remaining = sizeof(tooltip) - strlen(tooltip) - 1;
            strncat(tooltip, "\n", remaining);
            remaining = sizeof(tooltip) - strlen(tooltip) - 1;
            strncat(tooltip, currentMessage, remaining);
        }
    }

    if (strlen(tooltip) > 63) {
        tooltip[60] = '.';
        tooltip[61] = '.';
        tooltip[62] = '.';
        tooltip[63] = '\0';
    }

    strcpy(nid.szTip, tooltip);
    nid.uFlags = NIF_TIP;
    Shell_NotifyIconA(NIM_MODIFY, &nid);
}

void SetIcon(HICON icon) {
    const char* iconName = "Unknown";
    if (icon == hIconSuccess) iconName = "Success";
    else if (icon == hIconFail) iconName = "Fail";
    else if (icon == hIconEmpty) iconName = "Empty";
    else if (icon == hIconBlank) iconName = "Blank";

    LogMessage("SetIcon called: icon=%s", iconName);

    currentIcon = icon;
    iconVisible = TRUE;
    nid.hIcon = icon;
    nid.uFlags = NIF_ICON;
    Shell_NotifyIconA(NIM_MODIFY, &nid);
}

void CALLBACK TooltipTimer(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime) {
    UNREFERENCED_PARAMETER(hwnd);
    UNREFERENCED_PARAMETER(uMsg);
    UNREFERENCED_PARAMETER(idEvent);
    UNREFERENCED_PARAMETER(dwTime);

    if (currentResult != RESULT_ERROR && currentResult != RESULT_INVALID) {
        UpdateTooltip();
    }
}

void CALLBACK RefreshTimer(HWND hwnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime) {
    UNREFERENCED_PARAMETER(hwnd);
    UNREFERENCED_PARAMETER(uMsg);
    UNREFERENCED_PARAMETER(idEvent);
    UNREFERENCED_PARAMETER(dwTime);

    LogMessage("Scheduled refresh timer fired.");
    RefreshStatus();
}

void CaptureCurrentDisplaySettings() {
    lastScreenWidth = GetSystemMetrics(SM_CXSCREEN);
    lastScreenHeight = GetSystemMetrics(SM_CYSCREEN);

    HDC hdc = GetDC(NULL);
    if (hdc) {
        lastDpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        lastDpiY = GetDeviceCaps(hdc, LOGPIXELSY);
        ReleaseDC(NULL, hdc);
    }
    LogMessage("Display settings captured: %dx%d, DPI: %dx%d",
               lastScreenWidth, lastScreenHeight, lastDpiX, lastDpiY);
}

BOOL HasDisplaySettingsChanged() {
    int currentWidth = GetSystemMetrics(SM_CXSCREEN);
    int currentHeight = GetSystemMetrics(SM_CYSCREEN);

    int currentDpiX = 0;
    int currentDpiY = 0;
    HDC hdc = GetDC(NULL);
    if (hdc) {
        currentDpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        currentDpiY = GetDeviceCaps(hdc, LOGPIXELSY);
        ReleaseDC(NULL, hdc);
    }

    BOOL changed = (currentWidth != lastScreenWidth ||
                    currentHeight != lastScreenHeight ||
                    currentDpiX != lastDpiX ||
                    currentDpiY != lastDpiY);

    if (changed) {
        LogMessage("Display settings change detected.");
    }

    return changed;
}

void RefreshTrayIconForNewResolution() {
    LogMessage("Refreshing tray icon for new resolution/DPI.");

    HICON savedIcon = currentIcon;

    Shell_NotifyIconA(NIM_DELETE, &nid);
    Sleep(10);

    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = g_hwnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = savedIcon;
    Shell_NotifyIconA(NIM_ADD, &nid);

    UpdateTooltip();
}

// --- History dialog ---

static INT_PTR CALLBACK HistoryDialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    UNREFERENCED_PARAMETER(lParam);

    switch (message) {
        case WM_INITDIALOG: {
            // Center dialog on screen
            RECT rcDlg, rcScreen;
            GetWindowRect(hDlg, &rcDlg);
            SystemParametersInfoW(SPI_GETWORKAREA, 0, &rcScreen, 0);
            int x = rcScreen.left + ((rcScreen.right - rcScreen.left) - (rcDlg.right - rcDlg.left)) / 2;
            int y = rcScreen.top + ((rcScreen.bottom - rcScreen.top) - (rcDlg.bottom - rcDlg.top)) / 2;
            SetWindowPos(hDlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);

            // Get the placeholder static control position and replace with ListView
            HWND hPlaceholder = GetDlgItem(hDlg, IDC_HISTORY_LIST);
            RECT rcList;
            GetWindowRect(hPlaceholder, &rcList);
            MapWindowPoints(HWND_DESKTOP, hDlg, (LPPOINT)&rcList, 2);
            ShowWindow(hPlaceholder, SW_HIDE);

            // Initialize common controls
            INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_LISTVIEW_CLASSES };
            InitCommonControlsEx(&icc);

            // Create ListView
            HWND hList = CreateWindowExA(0, WC_LISTVIEWA, "",
                WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_NOSORTHEADER,
                rcList.left, rcList.top,
                rcList.right - rcList.left, rcList.bottom - rcList.top,
                hDlg, (HMENU)(UINT_PTR)IDC_HISTORY_LIST, g_hInstance, NULL);

            SendMessageA(hList, LVM_SETEXTENDEDLISTVIEWSTYLE, 0,
                         LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

            // Add columns
            LVCOLUMNA col = {0};
            col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
            col.fmt = LVCFMT_LEFT;

            col.pszText = "Time";
            col.cx = 130;
            SendMessageA(hList, LVM_INSERTCOLUMNA, 0, (LPARAM)&col);

            col.pszText = "From";
            col.cx = 70;
            SendMessageA(hList, LVM_INSERTCOLUMNA, 1, (LPARAM)&col);

            col.pszText = "To";
            col.cx = 70;
            SendMessageA(hList, LVM_INSERTCOLUMNA, 2, (LPARAM)&col);

            col.pszText = "Message";
            col.cx = 200;
            SendMessageA(hList, LVM_INSERTCOLUMNA, 3, (LPARAM)&col);

            // Populate from ring buffer, most recent first
            if (historyCount == 0) {
                LVITEMA item = {0};
                item.mask = LVIF_TEXT;
                item.iItem = 0;
                item.iSubItem = 0;
                item.pszText = "No status changes recorded yet.";
                SendMessageA(hList, LVM_INSERTITEMA, 0, (LPARAM)&item);
            } else {
                for (int i = 0; i < historyCount; i++) {
                    HistoryEntry* entry = GetHistoryEntry(i);
                    if (!entry) continue;

                    char timeBuf[64];
                    sprintf(timeBuf, "%04d-%02d-%02d %02d:%02d:%02d",
                            entry->timestamp.wYear, entry->timestamp.wMonth, entry->timestamp.wDay,
                            entry->timestamp.wHour, entry->timestamp.wMinute, entry->timestamp.wSecond);

                    LVITEMA item = {0};
                    item.mask = LVIF_TEXT;
                    item.iItem = i;
                    item.iSubItem = 0;
                    item.pszText = timeBuf;
                    SendMessageA(hList, LVM_INSERTITEMA, 0, (LPARAM)&item);

                    item.iSubItem = 1;
                    item.pszText = (char*)ApiResultToString(entry->oldResult);
                    SendMessageA(hList, LVM_SETITEMA, 0, (LPARAM)&item);

                    item.iSubItem = 2;
                    item.pszText = (char*)ApiResultToString(entry->newResult);
                    SendMessageA(hList, LVM_SETITEMA, 0, (LPARAM)&item);

                    item.iSubItem = 3;
                    item.pszText = entry->newMessage;
                    SendMessageA(hList, LVM_SETITEMA, 0, (LPARAM)&item);
                }
            }

            // Auto-size Message column to fill remaining space
            SendMessageA(hList, LVM_SETCOLUMNWIDTH, 3, LVSCW_AUTOSIZE_USEHEADER);

            return TRUE;
        }

        case WM_COMMAND:
            if (LOWORD(wParam) == IDC_HISTORY_CLOSE || LOWORD(wParam) == IDCANCEL) {
                EndDialog(hDlg, 0);
                return TRUE;
            }
            break;

        case WM_CLOSE:
            EndDialog(hDlg, 0);
            return TRUE;
    }
    return FALSE;
}

void ShowHistoryDialog(HWND hwndParent) {
    const short DLG_WIDTH = 520;
    const short DLG_HEIGHT = 320;
    const short MARGIN_X = 8;
    const short MARGIN_Y = 8;
    const short BTN_W = 50;
    const short BTN_H = 14;

    size_t templateSize = 4096;
    BYTE* templateBuffer = (BYTE*)calloc(1, templateSize);
    if (!templateBuffer) return;

    BYTE* ptr = templateBuffer;

    // DLGTEMPLATE
    DLGTEMPLATE* dlgTemplate = (DLGTEMPLATE*)ptr;
    dlgTemplate->style = DS_MODALFRAME | DS_CENTER | WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_SETFONT;
    dlgTemplate->dwExtendedStyle = 0;
    dlgTemplate->cdit = 2;  // placeholder static + close button
    dlgTemplate->x = 0;
    dlgTemplate->y = 0;
    dlgTemplate->cx = DLG_WIDTH;
    dlgTemplate->cy = DLG_HEIGHT;
    ptr += sizeof(DLGTEMPLATE);

    // Menu (none)
    *(WORD*)ptr = 0;
    ptr += sizeof(WORD);

    // Class (default)
    *(WORD*)ptr = 0;
    ptr += sizeof(WORD);

    // Title
    const wchar_t* title = L"Status Change History";
    size_t titleLen = wcslen(title) + 1;
    memcpy(ptr, title, titleLen * sizeof(wchar_t));
    ptr += titleLen * sizeof(wchar_t);

    // Font size
    *(WORD*)ptr = 8;
    ptr += sizeof(WORD);

    // Font name
    const wchar_t* fontName = L"Segoe UI";
    size_t fontLen = wcslen(fontName) + 1;
    memcpy(ptr, fontName, fontLen * sizeof(wchar_t));
    ptr += fontLen * sizeof(wchar_t);

    // Placeholder static for ListView (will be replaced in WM_INITDIALOG)
    short listH = DLG_HEIGHT - MARGIN_Y - MARGIN_Y - BTN_H - 6;
    short listW = DLG_WIDTH - (2 * MARGIN_X);
    ptr = AddDialogControl(ptr, IDC_HISTORY_LIST, 0x0082, WS_CHILD | WS_VISIBLE,
                           MARGIN_X, MARGIN_Y, listW, listH, L"");

    // Close button
    short btnY = MARGIN_Y + listH + 4;
    short btnX = (DLG_WIDTH - BTN_W) / 2;
    ptr = AddDialogControl(ptr, IDC_HISTORY_CLOSE, 0x0080,
                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                           btnX, btnY, BTN_W, BTN_H, L"Close");

    DialogBoxIndirectParamW(g_hInstance, (DLGTEMPLATE*)templateBuffer,
                            hwndParent, HistoryDialogProc, 0);
    free(templateBuffer);
}

void ExitApplication(HWND hwnd) {
    LogMessage("=== Application shutting down ===");

    if (timerRefresh) KillTimer(hwnd, 1);
    if (timerTooltip) KillTimer(hwnd, 2);

    SaveHistoryToRegistry();
    FreeHistoryBuffer();

    Shell_NotifyIconA(NIM_DELETE, &nid);

    if (hIconEmpty) DestroyIcon(hIconEmpty);
    if (hIconSuccess) DestroyIcon(hIconSuccess);
    if (hIconFail) DestroyIcon(hIconFail);
    if (hIconBlank) DestroyIcon(hIconBlank);
    if (hMenu) DestroyMenu(hMenu);

    if (g_hMutex) {
        ReleaseMutex(g_hMutex);
        CloseHandle(g_hMutex);
        g_hMutex = NULL;
    }

    DestroyWindow(hwnd);
}
