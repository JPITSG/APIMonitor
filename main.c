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
#include <objbase.h>
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

#define WM_VALIDATE_RESULT      (WM_APP + 1)
#define WM_SHOW_FIRST_CONFIG    (WM_USER + 2)

typedef enum {
    RESULT_NONE,         // Initial state - no result yet
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
static ApiResult currentResult = RESULT_NONE;
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

// ============================================================================
// WebView2 COM interface definitions (minimal vtable approach)
// ============================================================================

// GUIDs
DEFINE_GUID(IID_ICoreWebView2Environment, 0xb96d755e,0x0319,0x4e92,0xa2,0x96,0x23,0x43,0x6f,0x46,0xa1,0xfc);
DEFINE_GUID(IID_ICoreWebView2Controller, 0x4d00c0d1,0x9583,0x4f38,0x8e,0x50,0xa9,0xa6,0xb3,0x44,0x78,0xcd);
DEFINE_GUID(IID_ICoreWebView2, 0x76eceacb,0x0462,0x4d94,0xac,0x83,0x42,0x3a,0x67,0x93,0x77,0x5e);
DEFINE_GUID(IID_ICoreWebView2Settings, 0xe562e4f0,0xd7fa,0x43ac,0x8d,0x71,0xc0,0x51,0x50,0x49,0x9f,0x00);

typedef struct EventRegistrationToken { __int64 value; } EventRegistrationToken;

// Forward declarations of COM interfaces
typedef struct ICoreWebView2Environment ICoreWebView2Environment;
typedef struct ICoreWebView2Controller ICoreWebView2Controller;
typedef struct ICoreWebView2 ICoreWebView2;
typedef struct ICoreWebView2Settings ICoreWebView2Settings;
typedef struct ICoreWebView2WebMessageReceivedEventArgs ICoreWebView2WebMessageReceivedEventArgs;
typedef struct ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler;
typedef struct ICoreWebView2CreateCoreWebView2ControllerCompletedHandler ICoreWebView2CreateCoreWebView2ControllerCompletedHandler;
typedef struct ICoreWebView2WebMessageReceivedEventHandler ICoreWebView2WebMessageReceivedEventHandler;

// ICoreWebView2Environment vtable
typedef struct ICoreWebView2EnvironmentVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2Environment*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2Environment*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2Environment*);
    HRESULT (STDMETHODCALLTYPE *CreateCoreWebView2Controller)(ICoreWebView2Environment*, HWND, ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*);
    HRESULT (STDMETHODCALLTYPE *CreateWebResourceResponse)(ICoreWebView2Environment*, void*, int, LPCWSTR, LPCWSTR, void**);
    HRESULT (STDMETHODCALLTYPE *get_BrowserVersionString)(ICoreWebView2Environment*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *add_NewBrowserVersionAvailable)(ICoreWebView2Environment*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NewBrowserVersionAvailable)(ICoreWebView2Environment*, EventRegistrationToken);
} ICoreWebView2EnvironmentVtbl;

struct ICoreWebView2Environment { const ICoreWebView2EnvironmentVtbl *lpVtbl; };

// ICoreWebView2Controller vtable
typedef struct ICoreWebView2ControllerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2Controller*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2Controller*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2Controller*);
    HRESULT (STDMETHODCALLTYPE *get_IsVisible)(ICoreWebView2Controller*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsVisible)(ICoreWebView2Controller*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_Bounds)(ICoreWebView2Controller*, RECT*);
    HRESULT (STDMETHODCALLTYPE *put_Bounds)(ICoreWebView2Controller*, RECT);
    HRESULT (STDMETHODCALLTYPE *get_ZoomFactor)(ICoreWebView2Controller*, double*);
    HRESULT (STDMETHODCALLTYPE *put_ZoomFactor)(ICoreWebView2Controller*, double);
    HRESULT (STDMETHODCALLTYPE *add_ZoomFactorChanged)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ZoomFactorChanged)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *SetBoundsAndZoomFactor)(ICoreWebView2Controller*, RECT, double);
    HRESULT (STDMETHODCALLTYPE *MoveFocus)(ICoreWebView2Controller*, int);
    HRESULT (STDMETHODCALLTYPE *add_MoveFocusRequested)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_MoveFocusRequested)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_GotFocus)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_GotFocus)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_LostFocus)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_LostFocus)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_AcceleratorKeyPressed)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_AcceleratorKeyPressed)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *get_ParentWindow)(ICoreWebView2Controller*, HWND*);
    HRESULT (STDMETHODCALLTYPE *put_ParentWindow)(ICoreWebView2Controller*, HWND);
    HRESULT (STDMETHODCALLTYPE *NotifyParentWindowPositionChanged)(ICoreWebView2Controller*);
    HRESULT (STDMETHODCALLTYPE *Close)(ICoreWebView2Controller*);
    HRESULT (STDMETHODCALLTYPE *get_CoreWebView2)(ICoreWebView2Controller*, ICoreWebView2**);
} ICoreWebView2ControllerVtbl;

struct ICoreWebView2Controller { const ICoreWebView2ControllerVtbl *lpVtbl; };

// ICoreWebView2 vtable (full table required)
typedef struct ICoreWebView2Vtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *get_Settings)(ICoreWebView2*, ICoreWebView2Settings**);
    HRESULT (STDMETHODCALLTYPE *get_Source)(ICoreWebView2*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *Navigate)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *NavigateToString)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *add_NavigationStarting)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NavigationStarting)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_ContentLoading)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ContentLoading)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_SourceChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_SourceChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_HistoryChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_HistoryChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_NavigationCompleted)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NavigationCompleted)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_FrameNavigationStarting)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_FrameNavigationStarting)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_FrameNavigationCompleted)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_FrameNavigationCompleted)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_ScriptDialogOpening)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ScriptDialogOpening)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_PermissionRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_PermissionRequested)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_ProcessFailed)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ProcessFailed)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *AddScriptToExecuteOnDocumentCreated)(ICoreWebView2*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *RemoveScriptToExecuteOnDocumentCreated)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *ExecuteScript)(ICoreWebView2*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *CapturePreview)(ICoreWebView2*, int, void*, void*);
    HRESULT (STDMETHODCALLTYPE *Reload)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *PostWebMessageAsJson)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *PostWebMessageAsString)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *add_WebMessageReceived)(ICoreWebView2*, ICoreWebView2WebMessageReceivedEventHandler*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_WebMessageReceived)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *CallDevToolsProtocolMethod)(ICoreWebView2*, LPCWSTR, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *get_BrowserProcessId)(ICoreWebView2*, UINT32*);
    HRESULT (STDMETHODCALLTYPE *get_CanGoBack)(ICoreWebView2*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *get_CanGoForward)(ICoreWebView2*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *GoBack)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *GoForward)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *GetDevToolsProtocolEventReceiver)(ICoreWebView2*, LPCWSTR, void**);
    HRESULT (STDMETHODCALLTYPE *Stop)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *add_NewWindowRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NewWindowRequested)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_DocumentTitleChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_DocumentTitleChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *get_DocumentTitle)(ICoreWebView2*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *AddHostObjectToScript)(ICoreWebView2*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *RemoveHostObjectFromScript)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *OpenDevToolsWindow)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *add_ContainsFullScreenElementChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ContainsFullScreenElementChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *get_ContainsFullScreenElement)(ICoreWebView2*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *add_WebResourceRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_WebResourceRequested)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *AddWebResourceRequestedFilter)(ICoreWebView2*, LPCWSTR, int);
    HRESULT (STDMETHODCALLTYPE *RemoveWebResourceRequestedFilter)(ICoreWebView2*, LPCWSTR, int);
    HRESULT (STDMETHODCALLTYPE *add_WindowCloseRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_WindowCloseRequested)(ICoreWebView2*, EventRegistrationToken);
} ICoreWebView2Vtbl;

struct ICoreWebView2 { const ICoreWebView2Vtbl *lpVtbl; };

// ICoreWebView2Settings vtable
typedef struct ICoreWebView2SettingsVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2Settings*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2Settings*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2Settings*);
    HRESULT (STDMETHODCALLTYPE *get_IsScriptEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsScriptEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsWebMessageEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsWebMessageEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreDefaultScriptDialogsEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreDefaultScriptDialogsEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsStatusBarEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsStatusBarEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreDevToolsEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreDevToolsEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreDefaultContextMenusEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreDefaultContextMenusEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreHostObjectsAllowed)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreHostObjectsAllowed)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsZoomControlEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsZoomControlEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsBuiltInErrorPageEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsBuiltInErrorPageEnabled)(ICoreWebView2Settings*, BOOL);
} ICoreWebView2SettingsVtbl;

struct ICoreWebView2Settings { const ICoreWebView2SettingsVtbl *lpVtbl; };

// ICoreWebView2WebMessageReceivedEventArgs vtable
typedef struct ICoreWebView2WebMessageReceivedEventArgsVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2WebMessageReceivedEventArgs*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2WebMessageReceivedEventArgs*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2WebMessageReceivedEventArgs*);
    HRESULT (STDMETHODCALLTYPE *get_Source)(ICoreWebView2WebMessageReceivedEventArgs*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *get_WebMessageAsJson)(ICoreWebView2WebMessageReceivedEventArgs*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *TryGetWebMessageAsString)(ICoreWebView2WebMessageReceivedEventArgs*, LPWSTR*);
} ICoreWebView2WebMessageReceivedEventArgsVtbl;

struct ICoreWebView2WebMessageReceivedEventArgs { const ICoreWebView2WebMessageReceivedEventArgsVtbl *lpVtbl; };

// ============================================================================
// COM callback handler types
// ============================================================================

typedef struct EnvironmentCompletedHandlerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*, HRESULT, ICoreWebView2Environment*);
} EnvironmentCompletedHandlerVtbl;

struct ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler {
    const EnvironmentCompletedHandlerVtbl *lpVtbl;
    ULONG refCount;
};

typedef struct ControllerCompletedHandlerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, HRESULT, ICoreWebView2Controller*);
} ControllerCompletedHandlerVtbl;

struct ICoreWebView2CreateCoreWebView2ControllerCompletedHandler {
    const ControllerCompletedHandlerVtbl *lpVtbl;
    ULONG refCount;
};

typedef struct WebMessageReceivedHandlerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2WebMessageReceivedEventHandler*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2WebMessageReceivedEventHandler*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2WebMessageReceivedEventHandler*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(ICoreWebView2WebMessageReceivedEventHandler*, ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs*);
} WebMessageReceivedHandlerVtbl;

struct ICoreWebView2WebMessageReceivedEventHandler {
    const WebMessageReceivedHandlerVtbl *lpVtbl;
    ULONG refCount;
};

// ============================================================================
// WebView2 globals
// ============================================================================

static HWND g_webviewHwnd = NULL;
static ICoreWebView2Environment *g_webviewEnv = NULL;
static ICoreWebView2Controller *g_webviewController = NULL;
static ICoreWebView2 *g_webviewView = NULL;
static char g_pendingView[16] = "";

typedef HRESULT (STDAPICALLTYPE *PFN_CreateCoreWebView2EnvironmentWithOptions)(
    LPCWSTR browserExecutableFolder, LPCWSTR userDataFolder, void* options,
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* handler);

static PFN_CreateCoreWebView2EnvironmentWithOptions fnCreateEnvironment = NULL;
static WCHAR g_extractedDllPath[MAX_PATH] = {0};

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
void ShowHistoryDialog(HWND hwndParent);
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
void SaveHistoryToRegistry(void);
void LoadHistoryFromRegistry(void);
static void ShowWebViewDialog(const char* view, int width, int height);

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

    // On first launch, post message to show config dialog after message loop starts
    if (firstLaunch) {
        LogMessage("First launch detected, will show configuration dialog.");
        MarkAsConfigured();
        PostMessage(hwnd, WM_SHOW_FIRST_CONFIG, 0, 0);
    }

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
        case WM_SHOW_FIRST_CONFIG:
            ShowConfigDialog(g_hwnd);
            break;

        case WM_TRAYICON:
            if (lParam == WM_RBUTTONUP) {
                POINT pt;
                GetCursorPos(&pt);
                SetForegroundWindow(hwnd);
                EnableMenuItem(hMenu, ID_TRAY_CONFIGURE, g_webviewHwnd ? MF_GRAYED : MF_ENABLED);
                EnableMenuItem(hMenu, ID_TRAY_HISTORY, g_webviewHwnd ? MF_GRAYED : MF_ENABLED);
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
        case RESULT_NONE:    return "-";
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

void ShowConfigDialog(HWND hwndParent) {
    (void)hwndParent;
    ShowWebViewDialog("config", 480, 380);
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
    // Detect status changes and record in history (skip if this is the first result)
    BOOL resultChanged = (result != currentResult);
    BOOL messageChanged = (message && strcmp(currentMessage, message) != 0);
    if (currentResult != RESULT_NONE && (resultChanged || (messageChanged && result != RESULT_SUCCESS))) {
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

void ShowHistoryDialog(HWND hwndParent) {
    (void)hwndParent;
    ShowWebViewDialog("history", 700, 500);
}

// ============================================================================
// WebView2 helper functions
// ============================================================================

static BOOL load_webview2_loader(void) {
    HRSRC hRes = FindResource(NULL, MAKEINTRESOURCE(IDR_WEBVIEW2_DLL), RT_RCDATA);
    if (hRes) {
        HGLOBAL hData = LoadResource(NULL, hRes);
        DWORD dllSize = SizeofResource(NULL, hRes);
        const void *dllBytes = LockResource(hData);
        if (dllBytes && dllSize > 0) {
            WCHAR tempDir[MAX_PATH];
            DWORD tempLen = GetTempPathW(MAX_PATH, tempDir);
            if (tempLen > 0 && tempLen < MAX_PATH - 30) {
                swprintf(g_extractedDllPath, MAX_PATH, L"%sWebView2Loader.dll", tempDir);
                HANDLE hFile = CreateFileW(g_extractedDllPath, GENERIC_WRITE, 0, NULL,
                    CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
                if (hFile != INVALID_HANDLE_VALUE) {
                    DWORD written = 0;
                    WriteFile(hFile, dllBytes, dllSize, &written, NULL);
                    CloseHandle(hFile);
                    if (written == dllSize) {
                        HMODULE hMod = LoadLibraryW(g_extractedDllPath);
                        if (hMod) {
                            fnCreateEnvironment = (PFN_CreateCoreWebView2EnvironmentWithOptions)
                                GetProcAddress(hMod, "CreateCoreWebView2EnvironmentWithOptions");
                            if (fnCreateEnvironment) return TRUE;
                        }
                    }
                }
            }
        }
    }
    return FALSE;
}

static void webview_execute_script(const wchar_t* script) {
    if (g_webviewView) {
        g_webviewView->lpVtbl->ExecuteScript(g_webviewView, script, NULL);
    }
}

// Minimal JSON parser helpers
static BOOL json_get_string(const char *json, const char *key, char *out, size_t outLen) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return FALSE;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    if (*p != '"') return FALSE;
    p++;
    size_t i = 0;
    while (*p && *p != '"' && i < outLen - 1) {
        out[i++] = *p++;
    }
    out[i] = '\0';
    return TRUE;
}

static BOOL json_get_int(const char *json, const char *key, int *out) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return FALSE;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    *out = atoi(p);
    return TRUE;
}

static BOOL json_get_bool(const char *json, const char *key, BOOL *out) {
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return FALSE;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    *out = (strncmp(p, "true", 4) == 0) ? TRUE : FALSE;
    return TRUE;
}

// Escape a string for safe JSON embedding
static void json_escape_string(const char *in, wchar_t *out, size_t outLen) {
    size_t j = 0;
    for (size_t i = 0; in[i] && j < outLen - 2; i++) {
        char c = in[i];
        if (c == '"' || c == '\\') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = (wchar_t)c;
        } else if (c == '\n') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'n';
        } else if (c == '\r') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'r';
        } else {
            out[j++] = (wchar_t)(unsigned char)c;
        }
    }
    out[j] = L'\0';
}

// Wide-char version of ApiResultToString
static const wchar_t* ApiResultToStringW(ApiResult r) {
    switch (r) {
        case RESULT_NONE:    return L"-";
        case RESULT_SUCCESS: return L"Success";
        case RESULT_FAIL:    return L"Fail";
        case RESULT_ERROR:   return L"Error";
        case RESULT_INVALID: return L"Invalid";
        default:             return L"Unknown";
    }
}

// ============================================================================
// Push functions (C -> JS)
// ============================================================================

static void webview_push_init_config(void) {
    wchar_t wUrl[1024];
    json_escape_string(configApiUrl, wUrl, 1024);
    wchar_t script[2048];
    swprintf(script, 2048,
        L"window.onInit({\"view\":\"config\",\"config\":{\"url\":\"%s\",\"interval\":%d,\"loggingEnabled\":%s,\"historyLimit\":%d}})",
        wUrl, configRefreshInterval,
        configLoggingEnabled ? L"true" : L"false",
        configHistoryLimit);
    webview_execute_script(script);
}

static void webview_push_history_json(wchar_t *buf, size_t bufLen) {
    size_t pos = 0;
    pos += swprintf(buf + pos, bufLen - pos, L"[");
    for (int i = 0; i < historyCount && pos < bufLen - 200; i++) {
        HistoryEntry* entry = GetHistoryEntry(i);
        if (!entry) continue;
        if (i > 0) pos += swprintf(buf + pos, bufLen - pos, L",");

        wchar_t wMsg[512];
        json_escape_string(entry->newMessage, wMsg, 512);

        pos += swprintf(buf + pos, bufLen - pos,
            L"{\"time\":\"%04d-%02d-%02d %02d:%02d:%02d\",\"from\":\"%s\",\"to\":\"%s\",\"message\":\"%s\"}",
            entry->timestamp.wYear, entry->timestamp.wMonth, entry->timestamp.wDay,
            entry->timestamp.wHour, entry->timestamp.wMinute, entry->timestamp.wSecond,
            ApiResultToStringW(entry->oldResult),
            ApiResultToStringW(entry->newResult),
            wMsg);
    }
    if (pos < bufLen - 1) pos += swprintf(buf + pos, bufLen - pos, L"]");
}

static void webview_push_init_history(void) {
    // Allocate buffer large enough for history entries
    size_t bufLen = (size_t)historyCount * 256 + 256;
    if (bufLen < 1024) bufLen = 1024;
    wchar_t *histJson = (wchar_t*)malloc(bufLen * sizeof(wchar_t));
    if (!histJson) return;
    webview_push_history_json(histJson, bufLen);

    size_t scriptLen = bufLen + 256;
    wchar_t *script = (wchar_t*)malloc(scriptLen * sizeof(wchar_t));
    if (!script) { free(histJson); return; }
    swprintf(script, scriptLen, L"window.onInit({\"view\":\"history\",\"history\":%s})", histJson);
    webview_execute_script(script);
    free(script);
    free(histJson);
}

static void webview_push_validation_result(BOOL valid) {
    wchar_t script[128];
    swprintf(script, 128, L"window.onValidationResult({\"valid\":%s})", valid ? L"true" : L"false");
    webview_execute_script(script);
}

static void webview_push_history_update(void) {
    size_t bufLen = (size_t)historyCount * 256 + 256;
    if (bufLen < 1024) bufLen = 1024;
    wchar_t *histJson = (wchar_t*)malloc(bufLen * sizeof(wchar_t));
    if (!histJson) return;
    webview_push_history_json(histJson, bufLen);

    size_t scriptLen = bufLen + 128;
    wchar_t *script = (wchar_t*)malloc(scriptLen * sizeof(wchar_t));
    if (!script) { free(histJson); return; }
    swprintf(script, scriptLen, L"window.onHistoryUpdate(%s)", histJson);
    webview_execute_script(script);
    free(script);
    free(histJson);
}

// ============================================================================
// COM callback handler implementations
// ============================================================================

static HRESULT STDMETHODCALLTYPE EnvCompleted_Invoke(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*, HRESULT, ICoreWebView2Environment*);
static HRESULT STDMETHODCALLTYPE CtrlCompleted_Invoke(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, HRESULT, ICoreWebView2Controller*);
static HRESULT STDMETHODCALLTYPE MsgReceived_Invoke(ICoreWebView2WebMessageReceivedEventHandler*, ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs*);

static HRESULT STDMETHODCALLTYPE EnvCompleted_QueryInterface(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This, REFIID riid, void **ppv) {
    (void)riid;
    *ppv = This;
    This->lpVtbl->AddRef(This);
    return S_OK;
}
static ULONG STDMETHODCALLTYPE EnvCompleted_AddRef(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This) {
    return ++This->refCount;
}
static ULONG STDMETHODCALLTYPE EnvCompleted_Release(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This) {
    ULONG rc = --This->refCount;
    if (rc == 0) free(This);
    return rc;
}

static HRESULT STDMETHODCALLTYPE EnvCompleted_Invoke(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This, HRESULT result, ICoreWebView2Environment *env) {
    (void)This;
    if (FAILED(result) || !env) return result;
    g_webviewEnv = env;
    env->lpVtbl->AddRef(env);

    static ControllerCompletedHandlerVtbl ctrlVtbl = {0};
    static BOOL ctrlVtblInit = FALSE;
    if (!ctrlVtblInit) {
        ctrlVtbl.QueryInterface = (HRESULT (STDMETHODCALLTYPE *)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, REFIID, void**))EnvCompleted_QueryInterface;
        ctrlVtbl.AddRef = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*))EnvCompleted_AddRef;
        ctrlVtbl.Release = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*))EnvCompleted_Release;
        ctrlVtbl.Invoke = CtrlCompleted_Invoke;
        ctrlVtblInit = TRUE;
    }

    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler *handler = malloc(sizeof(*handler));
    handler->lpVtbl = &ctrlVtbl;
    handler->refCount = 1;

    env->lpVtbl->CreateCoreWebView2Controller(env, g_webviewHwnd, handler);
    handler->lpVtbl->Release(handler);
    return S_OK;
}

static EnvironmentCompletedHandlerVtbl g_envCompletedVtbl = {
    EnvCompleted_QueryInterface,
    EnvCompleted_AddRef,
    EnvCompleted_Release,
    EnvCompleted_Invoke
};

static HRESULT STDMETHODCALLTYPE CtrlCompleted_Invoke(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler *This, HRESULT result, ICoreWebView2Controller *controller) {
    (void)This;
    if (FAILED(result) || !controller) return result;

    g_webviewController = controller;
    controller->lpVtbl->AddRef(controller);

    RECT bounds;
    GetClientRect(g_webviewHwnd, &bounds);
    controller->lpVtbl->put_Bounds(controller, bounds);

    ICoreWebView2 *webview = NULL;
    controller->lpVtbl->get_CoreWebView2(controller, &webview);
    if (!webview) return E_FAIL;
    g_webviewView = webview;

    ICoreWebView2Settings *settings = NULL;
    webview->lpVtbl->get_Settings(webview, &settings);
    if (settings) {
        settings->lpVtbl->put_AreDefaultContextMenusEnabled(settings, FALSE);
        settings->lpVtbl->put_AreDevToolsEnabled(settings, FALSE);
        settings->lpVtbl->put_IsStatusBarEnabled(settings, FALSE);
        settings->lpVtbl->put_IsZoomControlEnabled(settings, FALSE);
        settings->lpVtbl->Release(settings);
    }

    static WebMessageReceivedHandlerVtbl msgVtbl = {0};
    static BOOL msgVtblInit = FALSE;
    if (!msgVtblInit) {
        msgVtbl.QueryInterface = (HRESULT (STDMETHODCALLTYPE *)(ICoreWebView2WebMessageReceivedEventHandler*, REFIID, void**))EnvCompleted_QueryInterface;
        msgVtbl.AddRef = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2WebMessageReceivedEventHandler*))EnvCompleted_AddRef;
        msgVtbl.Release = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2WebMessageReceivedEventHandler*))EnvCompleted_Release;
        msgVtbl.Invoke = MsgReceived_Invoke;
        msgVtblInit = TRUE;
    }

    ICoreWebView2WebMessageReceivedEventHandler *msgHandler = malloc(sizeof(*msgHandler));
    msgHandler->lpVtbl = &msgVtbl;
    msgHandler->refCount = 1;

    EventRegistrationToken token;
    webview->lpVtbl->add_WebMessageReceived(webview, msgHandler, &token);
    msgHandler->lpVtbl->Release(msgHandler);

    // Load embedded HTML from resources
    HRSRC hRes = FindResource(NULL, MAKEINTRESOURCE(IDR_HTML_UI), RT_RCDATA);
    if (hRes) {
        HGLOBAL hData = LoadResource(NULL, hRes);
        if (hData) {
            DWORD htmlSize = SizeofResource(NULL, hRes);
            const char *htmlUtf8 = (const char *)LockResource(hData);
            if (htmlUtf8 && htmlSize > 0) {
                int wLen = MultiByteToWideChar(CP_UTF8, 0, htmlUtf8, (int)htmlSize, NULL, 0);
                wchar_t *wHtml = malloc((wLen + 1) * sizeof(wchar_t));
                MultiByteToWideChar(CP_UTF8, 0, htmlUtf8, (int)htmlSize, wHtml, wLen);
                wHtml[wLen] = L'\0';
                webview->lpVtbl->NavigateToString(webview, wHtml);
                free(wHtml);
            }
        }
    }

    return S_OK;
}

// --- WebMessageReceivedHandler ---

static HRESULT STDMETHODCALLTYPE MsgReceived_Invoke(ICoreWebView2WebMessageReceivedEventHandler *This, ICoreWebView2 *sender, ICoreWebView2WebMessageReceivedEventArgs *args) {
    (void)This; (void)sender;

    LPWSTR wMsg = NULL;
    args->lpVtbl->TryGetWebMessageAsString(args, &wMsg);
    if (!wMsg) return S_OK;

    int len = WideCharToMultiByte(CP_UTF8, 0, wMsg, -1, NULL, 0, NULL, NULL);
    char *msg = malloc(len);
    WideCharToMultiByte(CP_UTF8, 0, wMsg, -1, msg, len, NULL, NULL);
    CoTaskMemFree(wMsg);

    char action[64] = {0};
    json_get_string(msg, "action", action, sizeof(action));

    if (strcmp(action, "getInit") == 0) {
        if (strcmp(g_pendingView, "config") == 0) {
            webview_push_init_config();
        } else if (strcmp(g_pendingView, "history") == 0) {
            webview_push_init_history();
        }
    } else if (strcmp(action, "validateUrl") == 0) {
        char url[512] = {0};
        json_get_string(msg, "url", url, sizeof(url));
        if (url[0]) {
            StartValidation(g_webviewHwnd, url);
        }
    } else if (strcmp(action, "saveSettings") == 0) {
        char url[512] = {0};
        int interval = 60;
        BOOL logging = TRUE;
        int histLimit = 100;
        json_get_string(msg, "url", url, sizeof(url));
        json_get_int(msg, "interval", &interval);
        json_get_bool(msg, "loggingEnabled", &logging);
        json_get_int(msg, "historyLimit", &histLimit);

        if (url[0]) {
            strncpy(configApiUrl, url, sizeof(configApiUrl) - 1);
            configApiUrl[sizeof(configApiUrl) - 1] = '\0';
        }
        if (interval == 60 || interval == 120 || interval == 300) {
            configRefreshInterval = interval;
        }
        configLoggingEnabled = logging;
        if (histLimit >= 10 && histLimit <= 10000) {
            configHistoryLimit = histLimit;
            InitHistoryBuffer(configHistoryLimit);
        }

        SaveConfigToRegistry();
        MarkAsConfigured();
        ApplyConfiguration();
        LogMessage("Configuration updated via WebView dialog: URL=%s, Interval=%d, Logging=%s, HistoryLimit=%d",
                   configApiUrl, configRefreshInterval, configLoggingEnabled ? "enabled" : "disabled", configHistoryLimit);
        PostMessage(g_webviewHwnd, WM_CLOSE, 0, 0);
    } else if (strcmp(action, "close") == 0) {
        PostMessage(g_webviewHwnd, WM_CLOSE, 0, 0);
    } else if (strcmp(action, "clearHistory") == 0) {
        historyCount = 0;
        historyHead = 0;
        webview_push_history_update();
    } else if (strcmp(action, "resize") == 0) {
        int contentHeight = 0;
        json_get_int(msg, "height", &contentHeight);
        if (contentHeight > 0 && g_webviewHwnd) {
            RECT clientRect = {0}, windowRect = {0};
            GetClientRect(g_webviewHwnd, &clientRect);
            GetWindowRect(g_webviewHwnd, &windowRect);
            int chromeH = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top);
            int newWindowH = contentHeight + chromeH;
            int windowW = windowRect.right - windowRect.left;
            SetWindowPos(g_webviewHwnd, NULL, 0, 0, windowW, newWindowH,
                SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
    }

    free(msg);
    return S_OK;
}

// ============================================================================
// WebView2 window
// ============================================================================

static LRESULT CALLBACK WebViewWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_SIZE:
            if (g_webviewController) {
                RECT bounds;
                GetClientRect(hwnd, &bounds);
                g_webviewController->lpVtbl->put_Bounds(g_webviewController, bounds);
            }
            return 0;

        case WM_VALIDATE_RESULT:
            if ((LONG)wParam == g_validateGeneration) {
                webview_push_validation_result(lParam ? TRUE : FALSE);
            }
            return 0;

        case WM_CLOSE:
            if (g_webviewController) {
                g_webviewController->lpVtbl->Close(g_webviewController);
                g_webviewController->lpVtbl->Release(g_webviewController);
                g_webviewController = NULL;
            }
            if (g_webviewView) {
                g_webviewView->lpVtbl->Release(g_webviewView);
                g_webviewView = NULL;
            }
            if (g_webviewEnv) {
                g_webviewEnv->lpVtbl->Release(g_webviewEnv);
                g_webviewEnv = NULL;
            }
            DestroyWindow(hwnd);
            return 0;

        case WM_DESTROY:
            g_webviewHwnd = NULL;
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowWebViewDialog(const char* view, int width, int height) {
    // If already open, bring to front
    if (g_webviewHwnd != NULL) {
        SetForegroundWindow(g_webviewHwnd);
        return;
    }

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (!fnCreateEnvironment && !load_webview2_loader()) {
        MessageBoxW(NULL,
            L"Failed to load WebView2.\n\n"
            L"Please ensure the Microsoft Edge WebView2 Runtime is installed.\n"
            L"Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            L"API Monitor", MB_ICONERROR | MB_OK);
        return;
    }

    strncpy(g_pendingView, view, sizeof(g_pendingView) - 1);
    g_pendingView[sizeof(g_pendingView) - 1] = '\0';

    // Register window class (once)
    static BOOL classRegistered = FALSE;
    if (!classRegistered) {
        WNDCLASSEXW wc = {0};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = WebViewWndProc;
        wc.hInstance = g_hInstance;
        wc.hIcon = LoadIconW(g_hInstance, MAKEINTRESOURCEW(IDI_APPICON));
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"APIMonitorWebViewWnd";
        wc.hIconSm = LoadIconW(g_hInstance, MAKEINTRESOURCEW(IDI_APPICON));
        RegisterClassExW(&wc);
        classRegistered = TRUE;
    }

    // Window title based on view
    const wchar_t *title = L"Configuration";
    if (strcmp(view, "history") == 0) title = L"Status Change History";

    // Center on screen
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int posX = (screenW - width) / 2;
    int posY = (screenH - height) / 2;

    g_webviewHwnd = CreateWindowExW(0, L"APIMonitorWebViewWnd", title,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        posX, posY, width, height,
        NULL, NULL, g_hInstance, NULL);

    if (!g_webviewHwnd) {
        LogMessage("ERROR: Failed to create WebView2 window.");
        return;
    }

    ShowWindow(g_webviewHwnd, SW_SHOW);
    UpdateWindow(g_webviewHwnd);

    // Build user data folder path
    WCHAR userDataFolder[MAX_PATH];
    DWORD tempLen = GetTempPathW(MAX_PATH, userDataFolder);
    if (tempLen > 0 && tempLen < MAX_PATH - 30) {
        wcscat(userDataFolder, L"APIMonitor.WebView2");
    } else {
        wcscpy(userDataFolder, L"");
    }

    // Create environment
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *envHandler = malloc(sizeof(*envHandler));
    envHandler->lpVtbl = &g_envCompletedVtbl;
    envHandler->refCount = 1;

    HRESULT hr = fnCreateEnvironment(NULL, userDataFolder[0] ? userDataFolder : NULL, NULL, envHandler);
    envHandler->lpVtbl->Release(envHandler);

    if (FAILED(hr)) {
        LogMessage("ERROR: Failed to initialize WebView2 environment. HRESULT=0x%08lx", hr);
        MessageBoxW(NULL,
            L"Failed to initialize WebView2.\n\n"
            L"Please ensure the Microsoft Edge WebView2 Runtime is installed.\n"
            L"Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            L"API Monitor", MB_ICONERROR | MB_OK);
        DestroyWindow(g_webviewHwnd);
        g_webviewHwnd = NULL;
    }
}

void ExitApplication(HWND hwnd) {
    static BOOL alreadyExiting = FALSE;
    if (alreadyExiting) return;
    alreadyExiting = TRUE;

    LogMessage("=== Application shutting down ===");

    // Close WebView2 dialog if open
    if (g_webviewHwnd) SendMessage(g_webviewHwnd, WM_CLOSE, 0, 0);

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
