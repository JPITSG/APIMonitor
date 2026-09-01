// main.c
#define _WIN32_WINNT 0x0600
#include <windows.h>
#include <winhttp.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <wchar.h>
#include <stdarg.h>
#include <commctrl.h>
#include <objbase.h>
#include <shellapi.h>
#include <sddl.h>
#include <shlwapi.h>
#include <bcrypt.h>
#include <userenv.h>
#include <winver.h>
#include "resource.h"
#include "version.h"

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
#define REG_VALUE_AUTO_UPDATE "AutoCheckForUpdates"
#define REG_VALUE_IGNORED_UPDATE_VERSION_W L"IgnoredUpdateVersion"
#define REG_KEY_PATH_W L"SOFTWARE\\JPIT\\APIMonitor"

#define WM_VALIDATE_RESULT      (WM_APP + 1)
#define WM_APP_UPDATE_RESULT    (WM_APP + 2)
#define WM_APP_UPDATE_PROGRESS  (WM_APP + 3)
#define WM_SHOW_FIRST_CONFIG    (WM_USER + 2)
#define ID_TIMER_WEBVIEW_SHOW_FALLBACK 1006
#define WEBVIEW_SHOW_FALLBACK_DELAY_MS 350
#define ID_TIMER_AUTO_UPDATE 1007
#define AUTO_UPDATE_INTERVAL_MS (60u * 60u * 1000u)

#define APP_NAME L"APIMonitor"
#define UPDATE_URL L"https://github.com/JPITSG/APIMonitor/raw/refs/heads/main/release/APIMonitor.exe"
#define UPDATE_MAX_BYTES (100ULL * 1024ULL * 1024ULL)
#define UPDATE_PROGRESS_INTERVAL_MS 250
#define UPDATE_HELPER_READY_MS 10000
#define UPDATE_HELPER_WAIT_MS 120000

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
static BOOL configAutoCheckForUpdates = TRUE;
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
static BOOL g_webviewWindowShown = FALSE;
static BOOL g_configViewReady = FALSE;
static BOOL g_updateConfirmationPending = FALSE;
static wchar_t g_webView2Version[128] = L"Unknown";

typedef struct {
    WORD major;
    WORD minor;
    WORD patch;
    WORD build;
} ExecutableVersion;

typedef enum {
    UPDATE_CHECK_SAME = 1,
    UPDATE_CHECK_NEWER,
    UPDATE_CHECK_OLDER,
    UPDATE_CHECK_CANCELLED,
    UPDATE_CHECK_ERROR
} UpdateCheckKind;

typedef struct {
    HWND targetWindow;
    BOOL automatic;
    UpdateCheckKind kind;
    ULONGLONG cacheBuster;
    ExecutableVersion runningVersion;
    ExecutableVersion availableVersion;
    wchar_t message[512];
    wchar_t targetPath[MAX_PATH];
    wchar_t stagedPath[MAX_PATH];
} UpdateCheckTask;

static volatile LONG g_updateCheckPending = FALSE;
static volatile LONG g_updateCheckAutomatic = FALSE;
static BOOL g_updateInstallReady = FALSE;
static volatile LONG g_updateRequestSequence = 0;
static HANDLE g_updateCancelEvent = NULL;
static volatile LONG g_updateSpeedKbps = 0;
static volatile LONG g_updateProgressPosted = FALSE;
static UpdateCheckTask* volatile g_updatePostedResult = NULL;
static UpdateCheckTask* g_updateNoticeTask = NULL;
static UpdateCheckTask* g_updateReadyTask = NULL;
static wchar_t g_ignoredUpdateVersion[32] = L"";

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
static void StartUpdateCheck(BOOL automatic);
static void CancelUpdateCheck(void);
static void InstallPreparedUpdate(void);
static void DiscardPreparedUpdate(void);
static void DiscardPendingUpdateNotice(void);
static void DiscardUpdateTask(UpdateCheckTask* task);
static void HandleCompletedUpdateCheck(UpdateCheckTask* task);
static int HandleUpdateCommandLine(BOOL* handled, BOOL* updateCompleted);
static void webview_execute_script(const wchar_t* script);

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
    BOOL updateHelperHandled = FALSE;
    BOOL updateCompleted = FALSE;
    int updateHelperResult = HandleUpdateCommandLine(
        &updateHelperHandled, &updateCompleted);
    if (updateHelperHandled) return updateHelperResult;

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

    LogMessage("=== Application starting (Version: APIMonitor/%s) ===", APP_VERSION_STRING);

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

    SetTimer(hwnd, ID_TIMER_AUTO_UPDATE, AUTO_UPDATE_INTERVAL_MS, NULL);
    if (configAutoCheckForUpdates) StartUpdateCheck(TRUE);

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

    if (updateCompleted) {
        g_updateConfirmationPending = TRUE;
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
        case WM_APP_UPDATE_PROGRESS:
            InterlockedExchange(&g_updateProgressPosted, FALSE);
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                g_configViewReady) {
                DWORD speedKbps = (DWORD)InterlockedCompareExchange(
                    &g_updateSpeedKbps, 0, 0);
                wchar_t script[160];
                swprintf(script, sizeof(script) / sizeof(wchar_t),
                    L"window.onUpdateProgress({\"kilobytesPerSecond\":%lu})",
                    (unsigned long)speedKbps);
                webview_execute_script(script);
            }
            return 0;

        case WM_APP_UPDATE_RESULT: {
            UpdateCheckTask* task = (UpdateCheckTask*)InterlockedExchangePointer(
                (PVOID volatile*)&g_updatePostedResult, NULL);
            HandleCompletedUpdateCheck(task);
            return 0;
        }

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
            else if (wParam == ID_TIMER_AUTO_UPDATE && configAutoCheckForUpdates) {
                StartUpdateCheck(TRUE);
            }
            break;

        case WM_DESTROY:
            KillTimer(hwnd, ID_TIMER_AUTO_UPDATE);
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

    // Read automatic update checks (enabled by default).
    DWORD dwAutoUpdate = 1;
    size = sizeof(dwAutoUpdate);
    if (RegQueryValueExA(hKey, REG_VALUE_AUTO_UPDATE, NULL, &type,
                         (LPBYTE)&dwAutoUpdate, &size) == ERROR_SUCCESS &&
        type == REG_DWORD) {
        configAutoCheckForUpdates = dwAutoUpdate != 0;
    } else {
        configAutoCheckForUpdates = TRUE;
    }

    size = sizeof(g_ignoredUpdateVersion);
    if (RegQueryValueExW(hKey, REG_VALUE_IGNORED_UPDATE_VERSION_W, NULL,
                         &type, (LPBYTE)g_ignoredUpdateVersion,
                         &size) != ERROR_SUCCESS || type != REG_SZ) {
        g_ignoredUpdateVersion[0] = L'\0';
    }
    g_ignoredUpdateVersion[
        (sizeof(g_ignoredUpdateVersion) / sizeof(wchar_t)) - 1] = L'\0';

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

    DWORD dwAutoUpdate = configAutoCheckForUpdates ? 1 : 0;
    RegSetValueExA(hKey, REG_VALUE_AUTO_UPDATE, 0, REG_DWORD,
                   (const BYTE*)&dwAutoUpdate, sizeof(dwAutoUpdate));

    RegCloseKey(hKey);
    LogMessage("Configuration saved to registry: URL=%s, Interval=%d, Logging=%s, HistoryLimit=%d, AutoUpdate=%s",
               configApiUrl, configRefreshInterval,
               configLoggingEnabled ? "enabled" : "disabled", configHistoryLimit,
               configAutoCheckForUpdates ? "enabled" : "disabled");
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

    hSession = WinHttpOpen(L"APIMonitor/" APP_VERSION_WSTRING,
                           WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
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
        hSession = WinHttpOpen(L"APIMonitor/" APP_VERSION_WSTRING,
                               WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
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
    if (!hRes) {
        MessageBoxW(NULL, L"Failed to find WebView2Loader.dll in embedded resources.\n"
            L"The executable may need to be rebuilt.", L"API Monitor", MB_ICONERROR);
        return FALSE;
    }
    HGLOBAL hData = LoadResource(NULL, hRes);
    DWORD dllSize = SizeofResource(NULL, hRes);
    const void *dllBytes = LockResource(hData);
    if (!dllBytes || dllSize == 0) {
        MessageBoxW(NULL, L"Failed to load WebView2Loader.dll from embedded resources.",
            L"API Monitor", MB_ICONERROR);
        return FALSE;
    }
    WCHAR tempDir[MAX_PATH];
    DWORD tempLen = GetTempPathW(MAX_PATH, tempDir);
    if (tempLen == 0 || tempLen >= MAX_PATH - 50) {
        MessageBoxW(NULL, L"Failed to get temp directory path.", L"API Monitor", MB_ICONERROR);
        return FALSE;
    }
    // Use an APIMonitor-specific subdirectory to avoid conflicts
    swprintf(g_extractedDllPath, MAX_PATH, L"%sAPIMonitor", tempDir);
    CreateDirectoryW(g_extractedDllPath, NULL);
    swprintf(g_extractedDllPath, MAX_PATH, L"%sAPIMonitor\\WebView2Loader.dll", tempDir);

    // Try to load existing copy first (may already be extracted from a previous run)
    HMODULE hMod = LoadLibraryW(g_extractedDllPath);
    if (!hMod) {
        // Extract fresh copy
        HANDLE hFile = CreateFileW(g_extractedDllPath, GENERIC_WRITE, 0, NULL,
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile == INVALID_HANDLE_VALUE) {
            WCHAR msg[512];
            swprintf(msg, 512, L"Failed to write WebView2Loader.dll to temp directory.\n\n"
                L"Path: %s\nError: %lu", g_extractedDllPath, GetLastError());
            MessageBoxW(NULL, msg, L"API Monitor", MB_ICONERROR);
            return FALSE;
        }
        DWORD written = 0;
        WriteFile(hFile, dllBytes, dllSize, &written, NULL);
        CloseHandle(hFile);
        if (written != dllSize) {
            MessageBoxW(NULL, L"Failed to write complete WebView2Loader.dll to temp directory.",
                L"API Monitor", MB_ICONERROR);
            return FALSE;
        }
        hMod = LoadLibraryW(g_extractedDllPath);
    }
    if (!hMod) {
        WCHAR msg[512];
        swprintf(msg, 512, L"Failed to load WebView2Loader.dll.\n\n"
            L"Path: %s\nError: %lu", g_extractedDllPath, GetLastError());
        MessageBoxW(NULL, msg, L"API Monitor", MB_ICONERROR);
        return FALSE;
    }
    fnCreateEnvironment = (PFN_CreateCoreWebView2EnvironmentWithOptions)
        GetProcAddress(hMod, "CreateCoreWebView2EnvironmentWithOptions");
    if (!fnCreateEnvironment) {
        MessageBoxW(NULL, L"WebView2Loader.dll loaded but CreateCoreWebView2EnvironmentWithOptions not found.\n\n"
            L"The DLL may be corrupted or the wrong version.", L"API Monitor", MB_ICONERROR);
        return FALSE;
    }
    return TRUE;
}

static void webview_execute_script(const wchar_t* script) {
    if (g_webviewView) {
        g_webviewView->lpVtbl->ExecuteScript(g_webviewView, script, NULL);
    }
}

static void webview_sync_controller_bounds(void) {
    if (!g_webviewController || !g_webviewHwnd) return;
    RECT bounds;
    GetClientRect(g_webviewHwnd, &bounds);
    g_webviewController->lpVtbl->put_Bounds(g_webviewController, bounds);
    g_webviewController->lpVtbl->put_IsVisible(g_webviewController, TRUE);
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

static void LogUpdateMessage(const wchar_t* format, ...) {
    wchar_t wideMessage[2048];
    va_list args;
    va_start(args, format);
    vswprintf(wideMessage,
              sizeof(wideMessage) / sizeof(wchar_t), format, args);
    va_end(args);

    int length = WideCharToMultiByte(CP_UTF8, 0, wideMessage, -1,
                                     NULL, 0, NULL, NULL);
    if (length <= 0) return;
    char* utf8Message = (char*)malloc((size_t)length);
    if (!utf8Message) return;
    WideCharToMultiByte(CP_UTF8, 0, wideMessage, -1,
                        utf8Message, length, NULL, NULL);
    LogMessage("%s", utf8Message);
    free(utf8Message);
}

static void json_escape_wstring(const wchar_t *in, wchar_t *out,
                                size_t outLen) {
    size_t j = 0;
    for (size_t i = 0; in[i] && j < outLen - 2; i++) {
        wchar_t c = in[i];
        if (c == L'"' || c == L'\\') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = c;
        } else if (c == L'\n') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'n';
        } else if (c == L'\r') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'r';
        } else if (c == L'\t') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L't';
        } else {
            out[j++] = c;
        }
    }
    out[j] = L'\0';
}

// --- Self update -----------------------------------------------------------

typedef struct {
    HINTERNET session;
    HINTERNET connection;
    HINTERNET request;
} UpdateHttpRequest;

static void CloseUpdateHttpRequest(UpdateHttpRequest* http) {
    if (!http) return;
    if (http->request) WinHttpCloseHandle(http->request);
    if (http->connection) WinHttpCloseHandle(http->connection);
    if (http->session) WinHttpCloseHandle(http->session);
    ZeroMemory(http, sizeof(*http));
}

static void SetUpdateTaskError(UpdateCheckTask* task, LPCWSTR message,
                               DWORD errorCode) {
    if (!task) return;
    task->kind = UPDATE_CHECK_ERROR;
    if (errorCode) {
        swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                   L"%s (Windows error %lu).", message, (unsigned long)errorCode);
    } else {
        wcscpy_s(task->message, sizeof(task->message) / sizeof(wchar_t), message);
    }
}

static BOOL CancelUpdateTaskIfRequested(UpdateCheckTask* task) {
    if (!task || !g_updateCancelEvent ||
        WaitForSingleObject(g_updateCancelEvent, 0) != WAIT_OBJECT_0) {
        return FALSE;
    }
    task->kind = UPDATE_CHECK_CANCELLED;
    task->message[0] = L'\0';
    return TRUE;
}

static void PublishUpdateProgress(UpdateCheckTask* task, DWORD speedKbps) {
    if (!task || CancelUpdateTaskIfRequested(task) ||
        !IsWindow(task->targetWindow)) {
        return;
    }

    InterlockedExchange(&g_updateSpeedKbps, (LONG)speedKbps);
    if (InterlockedCompareExchange(&g_updateProgressPosted, TRUE, FALSE) == FALSE &&
        !PostMessageW(task->targetWindow, WM_APP_UPDATE_PROGRESS, 0, 0)) {
        InterlockedExchange(&g_updateProgressPosted, FALSE);
    }
}

static BOOL OpenUpdateHttpRequest(LPCWSTR verb, ULONGLONG cacheBuster,
                                  UpdateHttpRequest* http, DWORD* statusCode) {
    if (!verb || !http) return FALSE;
    ZeroMemory(http, sizeof(*http));
    if (statusCode) *statusCode = 0;

    wchar_t hostName[256] = L"";
    wchar_t urlPath[2048] = L"";
    wchar_t extraInfo[512] = L"";
    URL_COMPONENTS components = {0};
    components.dwStructSize = sizeof(components);
    components.lpszHostName = hostName;
    components.dwHostNameLength = sizeof(hostName) / sizeof(wchar_t);
    components.lpszUrlPath = urlPath;
    components.dwUrlPathLength = sizeof(urlPath) / sizeof(wchar_t);
    components.lpszExtraInfo = extraInfo;
    components.dwExtraInfoLength = sizeof(extraInfo) / sizeof(wchar_t);
    if (!WinHttpCrackUrl(UPDATE_URL, 0, 0, &components)) return FALSE;

    wchar_t objectName[2560];
    if (wcscpy_s(objectName, sizeof(objectName) / sizeof(wchar_t), urlPath) != 0 ||
        wcscat_s(objectName, sizeof(objectName) / sizeof(wchar_t), extraInfo) != 0) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    // A unique query value makes each button click reach the current branch
    // artifact even when an HTTP proxy or GitHub edge cache retains the
    // previous response. HEAD and GET share the same value within one check.
    wchar_t cacheSuffix[64];
    wchar_t separator = wcschr(objectName, L'?') ? L'&' : L'?';
    int cacheSuffixLength = swprintf_s(cacheSuffix,
        sizeof(cacheSuffix) / sizeof(wchar_t), L"%lcslUpdate=%016llx",
        separator, (unsigned long long)cacheBuster);
    if (cacheSuffixLength <= 0 ||
        wcscat_s(objectName, sizeof(objectName) / sizeof(wchar_t), cacheSuffix) != 0) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    http->session = WinHttpOpen(L"APIMonitor Update",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!http->session) goto fail;
    WinHttpSetTimeouts(http->session, 10000, 10000, 15000, 30000);

    http->connection = WinHttpConnect(http->session, hostName,
                                      components.nPort, 0);
    if (!http->connection) goto fail;

    DWORD flags = WINHTTP_FLAG_REFRESH;
    if (components.nScheme == INTERNET_SCHEME_HTTPS) flags |= WINHTTP_FLAG_SECURE;
    http->request = WinHttpOpenRequest(http->connection, verb, objectName,
        NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!http->request) goto fail;

    static const wchar_t noCacheHeaders[] =
        L"Cache-Control: no-cache, no-store, max-age=0\r\nPragma: no-cache\r\n";
    WinHttpAddRequestHeaders(http->request, noCacheHeaders, (DWORD)-1L,
                            WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);
    if (!WinHttpSendRequest(http->request, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(http->request, NULL)) {
        goto fail;
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    if (!WinHttpQueryHeaders(http->request,
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize,
            WINHTTP_NO_HEADER_INDEX)) {
        goto fail;
    }
    if (statusCode) *statusCode = status;
    if (status != 200) {
        CloseUpdateHttpRequest(http);
        SetLastError(ERROR_WINHTTP_INVALID_SERVER_RESPONSE);
        return FALSE;
    }
    return TRUE;

fail: {
        DWORD errorCode = GetLastError();
        CloseUpdateHttpRequest(http);
        SetLastError(errorCode);
        return FALSE;
    }
}

static BOOL QueryUpdateContentLength(HINTERNET request, ULONGLONG* size) {
    if (!request || !size) return FALSE;
    wchar_t lengthText[64] = L"";
    DWORD lengthBytes = sizeof(lengthText);
    if (!WinHttpQueryHeaders(request, WINHTTP_QUERY_CONTENT_LENGTH,
            WINHTTP_HEADER_NAME_BY_INDEX, lengthText, &lengthBytes,
            WINHTTP_NO_HEADER_INDEX)) {
        return FALSE;
    }
    lengthText[(sizeof(lengthText) / sizeof(wchar_t)) - 1] = L'\0';

    wchar_t* end = NULL;
    unsigned long long parsed = _wcstoui64(lengthText, &end, 10);
    if (end == lengthText || !end || *end != L'\0' || parsed == 0) {
        SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }
    *size = (ULONGLONG)parsed;
    return TRUE;
}

static BOOL GetExecutableVersion(LPCWSTR path, ExecutableVersion* version) {
    if (!path || !*path || !version) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD ignored = 0;
    DWORD infoSize = GetFileVersionInfoSizeW(path, &ignored);
    if (infoSize == 0) {
        DWORD errorCode = GetLastError();
        SetLastError(errorCode ? errorCode : ERROR_RESOURCE_DATA_NOT_FOUND);
        return FALSE;
    }

    BYTE* infoData = (BYTE*)malloc(infoSize);
    if (!infoData) {
        SetLastError(ERROR_NOT_ENOUGH_MEMORY);
        return FALSE;
    }

    if (!GetFileVersionInfoW(path, 0, infoSize, infoData)) {
        DWORD errorCode = GetLastError();
        free(infoData);
        SetLastError(errorCode ? errorCode : ERROR_INVALID_DATA);
        return FALSE;
    }

    VS_FIXEDFILEINFO* fixedInfo = NULL;
    UINT fixedInfoSize = 0;
    if (!VerQueryValueW(infoData, L"\\", (LPVOID*)&fixedInfo, &fixedInfoSize) ||
        !fixedInfo || fixedInfoSize < sizeof(*fixedInfo) ||
        fixedInfo->dwSignature != VS_FFI_SIGNATURE) {
        free(infoData);
        SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }

    version->major = HIWORD(fixedInfo->dwFileVersionMS);
    version->minor = LOWORD(fixedInfo->dwFileVersionMS);
    version->patch = HIWORD(fixedInfo->dwFileVersionLS);
    version->build = LOWORD(fixedInfo->dwFileVersionLS);
    free(infoData);
    return TRUE;
}

static int CompareExecutableVersions(const ExecutableVersion* left,
                                     const ExecutableVersion* right) {
    const WORD leftParts[] = {
        left->major, left->minor, left->patch, left->build
    };
    const WORD rightParts[] = {
        right->major, right->minor, right->patch, right->build
    };
    for (size_t index = 0; index < sizeof(leftParts) / sizeof(leftParts[0]); index++) {
        if (leftParts[index] < rightParts[index]) return -1;
        if (leftParts[index] > rightParts[index]) return 1;
    }
    return 0;
}

static void FormatExecutableVersion(const ExecutableVersion* version,
                                    wchar_t* text, size_t textCch) {
    if (!version || !text || textCch == 0) return;
    if (swprintf_s(text, textCch, L"%u.%u.%u.%u",
                   (unsigned int)version->major,
                   (unsigned int)version->minor,
                   (unsigned int)version->patch,
                   (unsigned int)version->build) <= 0) {
        text[0] = L'\0';
    }
}

static void FormatExecutableVersionForDisplay(const ExecutableVersion* version,
                                              wchar_t* text, size_t textCch) {
    if (!version || !text || textCch == 0) return;
    if (version->build == 0) {
        if (swprintf_s(text, textCch, L"%u.%u.%u",
                       (unsigned int)version->major,
                       (unsigned int)version->minor,
                       (unsigned int)version->patch) <= 0) {
            text[0] = L'\0';
        }
        return;
    }
    FormatExecutableVersion(version, text, textCch);
}

static BOOL BuildUpdateTempPath(wchar_t path[MAX_PATH], LPCWSTR role,
                                DWORD processId) {
    if (!path || !role || !*role || processId == 0) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    wchar_t tempDirectory[MAX_PATH];
    DWORD tempLength = GetTempPathW(MAX_PATH, tempDirectory);
    if (tempLength == 0) return FALSE;
    if (tempLength >= MAX_PATH) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    int length = swprintf_s(path, MAX_PATH, L"%s%s-%s-%lu.exe",
                            tempDirectory, APP_NAME, role,
                            (unsigned long)processId);
    if (length <= 0 || length >= MAX_PATH) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }
    return TRUE;
}

static BOOL DeleteUpdateTempFile(LPCWSTR path) {
    if (!path || !*path) return FALSE;
    SetFileAttributesW(path, FILE_ATTRIBUTE_NORMAL);
    if (DeleteFileW(path)) return TRUE;
    DWORD errorCode = GetLastError();
    return errorCode == ERROR_FILE_NOT_FOUND || errorCode == ERROR_PATH_NOT_FOUND;
}

static BOOL QueryRemoteUpdateSize(UpdateCheckTask* task, ULONGLONG* size) {
    if (CancelUpdateTaskIfRequested(task)) return FALSE;

    UpdateHttpRequest http;
    DWORD status = 0;
    if (!OpenUpdateHttpRequest(L"HEAD", task->cacheBuster, &http, &status)) {
        DWORD errorCode = GetLastError();
        if (CancelUpdateTaskIfRequested(task)) return FALSE;
        if (status) {
            swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                       L"The update server returned HTTP status %lu.",
                       (unsigned long)status);
            task->kind = UPDATE_CHECK_ERROR;
        } else {
            SetUpdateTaskError(task, L"Could not contact the update server", errorCode);
        }
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    BOOL ok = QueryUpdateContentLength(http.request, size);
    DWORD errorCode = ok ? ERROR_SUCCESS : GetLastError();
    CloseUpdateHttpRequest(&http);
    if (CancelUpdateTaskIfRequested(task)) return FALSE;
    if (!ok) {
        SetUpdateTaskError(task, L"The update server did not report a valid file size",
                           errorCode);
        return FALSE;
    }
    if (*size > UPDATE_MAX_BYTES) {
        SetUpdateTaskError(task, L"The available update is unexpectedly large", 0);
        return FALSE;
    }
    return TRUE;
}

static BOOL DownloadUpdateFile(UpdateCheckTask* task, ULONGLONG expectedSize) {
    if (CancelUpdateTaskIfRequested(task)) return FALSE;

    UpdateHttpRequest http;
    DWORD status = 0;
    if (!OpenUpdateHttpRequest(L"GET", task->cacheBuster, &http, &status)) {
        DWORD errorCode = GetLastError();
        if (CancelUpdateTaskIfRequested(task)) return FALSE;
        if (status) {
            swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                       L"The update download returned HTTP status %lu.",
                       (unsigned long)status);
            task->kind = UPDATE_CHECK_ERROR;
        } else {
            SetUpdateTaskError(task, L"Could not download the update", errorCode);
        }
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    ULONGLONG downloadSize = 0;
    if (QueryUpdateContentLength(http.request, &downloadSize) &&
        downloadSize != expectedSize) {
        CloseUpdateHttpRequest(&http);
        SetUpdateTaskError(task,
            L"The available update changed while it was being downloaded. Try again", 0);
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    DeleteUpdateTempFile(task->stagedPath);
    HANDLE file = CreateFileW(task->stagedPath, GENERIC_WRITE, 0, NULL,
                              CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        DWORD errorCode = GetLastError();
        CloseUpdateHttpRequest(&http);
        SetUpdateTaskError(task, L"Could not create the staged update file", errorCode);
        return FALSE;
    }

    BOOL ok = TRUE;
    ULONGLONG totalWritten = 0;
    ULONGLONG speedWindowBytes = 0;
    ULONGLONG speedWindowStarted = GetTickCount64();
    BYTE buffer[64 * 1024];
    while (ok) {
        if (CancelUpdateTaskIfRequested(task)) {
            ok = FALSE;
            break;
        }

        DWORD bytesRead = 0;
        if (!WinHttpReadData(http.request, buffer, sizeof(buffer), &bytesRead)) {
            DWORD errorCode = GetLastError();
            if (!CancelUpdateTaskIfRequested(task)) {
                SetUpdateTaskError(task, L"The update download was interrupted", errorCode);
            }
            ok = FALSE;
            break;
        }
        if (CancelUpdateTaskIfRequested(task)) {
            ok = FALSE;
            break;
        }
        if (bytesRead == 0) break;
        if (totalWritten + bytesRead > expectedSize) {
            SetUpdateTaskError(task, L"The downloaded update has an invalid size", 0);
            ok = FALSE;
            break;
        }

        DWORD bytesWritten = 0;
        if (!WriteFile(file, buffer, bytesRead, &bytesWritten, NULL)) {
            SetUpdateTaskError(task, L"Could not write the staged update", GetLastError());
            ok = FALSE;
            break;
        }
        if (bytesWritten != bytesRead) {
            SetUpdateTaskError(task, L"Could not write the staged update",
                               ERROR_WRITE_FAULT);
            ok = FALSE;
            break;
        }
        totalWritten += bytesWritten;
        speedWindowBytes += bytesWritten;

        ULONGLONG now = GetTickCount64();
        ULONGLONG elapsed = now - speedWindowStarted;
        if (elapsed >= UPDATE_PROGRESS_INTERVAL_MS) {
            ULONGLONG speed = (speedWindowBytes * 1000ULL) /
                              (elapsed * 1024ULL);
            if (speed == 0 && speedWindowBytes > 0) speed = 1;
            if (speed > MAXLONG) speed = MAXLONG;
            PublishUpdateProgress(task, (DWORD)speed);
            speedWindowBytes = 0;
            speedWindowStarted = now;
        }
    }

    if (ok && CancelUpdateTaskIfRequested(task)) ok = FALSE;
    if (ok && totalWritten != expectedSize) {
        SetUpdateTaskError(task, L"The downloaded update is incomplete", 0);
        ok = FALSE;
    }
    if (ok && !FlushFileBuffers(file)) {
        SetUpdateTaskError(task, L"Could not finish writing the staged update", GetLastError());
        ok = FALSE;
    }
    CloseHandle(file);
    CloseUpdateHttpRequest(&http);

    if (ok && CancelUpdateTaskIfRequested(task)) ok = FALSE;
    DWORD binaryType = 0;
    if (ok && (!GetBinaryTypeW(task->stagedPath, &binaryType) ||
               binaryType != SCS_64BIT_BINARY)) {
        SetUpdateTaskError(task, L"The downloaded file is not a valid 64-bit application", 0);
        ok = FALSE;
    }
    if (!ok) DeleteUpdateTempFile(task->stagedPath);
    return ok;
}

static void DiscardUpdateTask(UpdateCheckTask* task) {
    if (!task) return;
    if (task->stagedPath[0]) {
        DeleteUpdateTempFile(task->stagedPath);
    }
    free(task);
}

static void PublishUpdateTask(UpdateCheckTask* task) {
    CancelUpdateTaskIfRequested(task);
    InterlockedExchange(&g_updateCheckPending, FALSE);
    InterlockedExchange(&g_updateCheckAutomatic, FALSE);
    if (!task || !IsWindow(task->targetWindow)) {
        DiscardUpdateTask(task);
        return;
    }

    UpdateCheckTask* previous = (UpdateCheckTask*)InterlockedExchangePointer(
        (PVOID volatile*)&g_updatePostedResult, task);
    DiscardUpdateTask(previous);
    if (!PostMessageW(task->targetWindow, WM_APP_UPDATE_RESULT, 0, 0)) {
        UpdateCheckTask* unclaimed = (UpdateCheckTask*)InterlockedExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, NULL);
        DiscardUpdateTask(unclaimed);
    }
}

static DWORD WINAPI UpdateCheckThread(LPVOID parameter) {
    UpdateCheckTask* task = (UpdateCheckTask*)parameter;
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    DWORD pathLength = GetModuleFileNameW(NULL, task->targetPath,
                                         sizeof(task->targetPath) / sizeof(wchar_t));
    if (pathLength == 0 || pathLength >= sizeof(task->targetPath) / sizeof(wchar_t)) {
        SetUpdateTaskError(task, L"Could not determine the running executable path",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }

    if (!GetExecutableVersion(task->targetPath, &task->runningVersion)) {
        SetUpdateTaskError(task, L"Could not read the running application version",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    ULONGLONG remoteSize = 0;
    if (!QueryRemoteUpdateSize(task, &remoteSize)) {
        PublishUpdateTask(task);
        return 0;
    }
    if (!BuildUpdateTempPath(task->stagedPath, L"download",
                             GetCurrentProcessId())) {
        SetUpdateTaskError(task, L"Could not create the temporary update path",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    if (!DownloadUpdateFile(task, remoteSize)) {
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    if (!GetExecutableVersion(task->stagedPath, &task->availableVersion)) {
        SetUpdateTaskError(task,
            L"The downloaded application does not contain valid version information",
            GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    LogUpdateMessage(L"[INFO] Update versions: running %u.%u.%u.%u, available %u.%u.%u.%u\n",
               (unsigned int)task->runningVersion.major,
               (unsigned int)task->runningVersion.minor,
               (unsigned int)task->runningVersion.patch,
               (unsigned int)task->runningVersion.build,
               (unsigned int)task->availableVersion.major,
               (unsigned int)task->availableVersion.minor,
               (unsigned int)task->availableVersion.patch,
               (unsigned int)task->availableVersion.build);
    int comparison = CompareExecutableVersions(&task->availableVersion,
                                               &task->runningVersion);
    task->kind = comparison > 0 ? UPDATE_CHECK_NEWER
               : comparison < 0 ? UPDATE_CHECK_OLDER
                                : UPDATE_CHECK_SAME;
    PublishUpdateTask(task);
    return 0;
}

static BOOL ParseUpdateProcessId(LPCWSTR text, DWORD* processId) {
    if (!text || !processId || !*text) return FALSE;
    wchar_t* end = NULL;
    unsigned long value = wcstoul(text, &end, 10);
    if (!end || *end != L'\0' || value == 0) return FALSE;
    *processId = (DWORD)value;
    return TRUE;
}

static BOOL ValidateUpdateTempFilePair(LPCWSTR helperPath,
                                       LPCWSTR stagedPath,
                                       DWORD processId) {
    if (!helperPath || !stagedPath || !*helperPath || !*stagedPath) return FALSE;

    wchar_t expectedHelperName[96], expectedStagedName[96];
    int helperNameLength = swprintf_s(expectedHelperName,
        sizeof(expectedHelperName) / sizeof(wchar_t),
        APP_NAME L"-updater-%lu.exe", (unsigned long)processId);
    int stagedNameLength = swprintf_s(expectedStagedName,
        sizeof(expectedStagedName) / sizeof(wchar_t),
        APP_NAME L"-download-%lu.exe", (unsigned long)processId);
    if (helperNameLength <= 0 || stagedNameLength <= 0 ||
        _wcsicmp(PathFindFileNameW(helperPath), expectedHelperName) != 0 ||
        _wcsicmp(PathFindFileNameW(stagedPath), expectedStagedName) != 0) {
        return FALSE;
    }

    wchar_t helperDirectory[MAX_PATH], stagedDirectory[MAX_PATH];
    if (wcscpy_s(helperDirectory, MAX_PATH, helperPath) != 0 ||
        wcscpy_s(stagedDirectory, MAX_PATH, stagedPath) != 0 ||
        !PathRemoveFileSpecW(helperDirectory) ||
        !PathRemoveFileSpecW(stagedDirectory)) {
        return FALSE;
    }
    return _wcsicmp(helperDirectory, stagedDirectory) == 0;
}

static HANDLE DuplicateUpdateLaunchToken(HANDLE process) {
    HANDLE processToken = NULL;
    HANDLE launchToken = NULL;
    if (!process ||
        !OpenProcessToken(process, TOKEN_QUERY | TOKEN_DUPLICATE,
                          &processToken)) {
        return NULL;
    }
    DuplicateTokenEx(processToken, MAXIMUM_ALLOWED, NULL,
                     SecurityImpersonation, TokenPrimary, &launchToken);
    CloseHandle(processToken);
    return launchToken;
}

static BOOL LaunchUpdateTarget(LPCWSTR targetPath, LPCWSTR stagedPath,
                               LPCWSTR helperPath, DWORD helperProcessId,
                               DWORD oldProcessId, HANDLE launchToken,
                               BOOL successfulUpdate) {
    wchar_t commandLine[MAX_PATH * 3 + 256];
    LPCWSTR finishAction = successfulUpdate
        ? L"--finish-update"
        : L"--finish-update-cleanup";
    int commandLength = swprintf_s(commandLine,
        sizeof(commandLine) / sizeof(wchar_t),
        L"\"%s\" %s %lu %lu \"%s\" \"%s\"", targetPath, finishAction,
        (unsigned long)helperProcessId, (unsigned long)oldProcessId,
        stagedPath, helperPath);
    if (commandLength <= 0 ||
        commandLength >= (int)(sizeof(commandLine) / sizeof(wchar_t))) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    STARTUPINFOW startupInfo = {0};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo = {0};
    BOOL launched = FALSE;
    if (launchToken) {
        wchar_t tokenCommandLine[MAX_PATH * 3 + 256];
        wcscpy_s(tokenCommandLine,
                 sizeof(tokenCommandLine) / sizeof(wchar_t), commandLine);
        LPVOID environment = NULL;
        BOOL hasEnvironment = CreateEnvironmentBlock(&environment, launchToken,
                                                     FALSE);
        launched = CreateProcessWithTokenW(launchToken, 0, targetPath,
                                           tokenCommandLine,
                                           hasEnvironment
                                               ? CREATE_UNICODE_ENVIRONMENT : 0,
                                           environment, NULL,
                                           &startupInfo, &processInfo);
        if (environment) DestroyEnvironmentBlock(environment);
    }
    if (!launched) {
        ZeroMemory(&processInfo, sizeof(processInfo));
        launched = CreateProcessW(targetPath, commandLine, NULL, NULL, FALSE,
                                  0, NULL, NULL, &startupInfo, &processInfo);
    }
    if (launched) {
        CloseHandle(processInfo.hProcess);
        CloseHandle(processInfo.hThread);
    }
    return launched;
}

static int RestartAfterUpdateFailure(LPCWSTR targetPath, LPCWSTR stagedPath,
                                     LPCWSTR helperPath, DWORD oldProcessId,
                                     HANDLE launchToken, LPCWSTR message) {
    MessageBoxW(NULL, message, APP_NAME L" Update", MB_OK | MB_ICONERROR);
    LaunchUpdateTarget(targetPath, stagedPath, helperPath,
                       GetCurrentProcessId(), oldProcessId, launchToken, FALSE);
    if (launchToken) CloseHandle(launchToken);
    DeleteUpdateTempFile(stagedPath);
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
    MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    return 1;
}

static int RunUpdateApplyHelper(DWORD oldProcessId, LPCWSTR readyEventName,
                                LPCWSTR targetPath, LPCWSTR stagedPath) {
    wchar_t expectedEventPrefix[96];
    int prefixLength = swprintf_s(expectedEventPrefix,
        sizeof(expectedEventPrefix) / sizeof(wchar_t),
        L"Local\\APIMonitor_UpdateReady_%lu_",
        (unsigned long)oldProcessId);
    if (prefixLength <= 0 || !readyEventName ||
        _wcsnicmp(readyEventName, expectedEventPrefix,
                  (size_t)prefixLength) != 0) {
        return ERROR_INVALID_DATA;
    }

    HANDLE readyEvent = OpenEventW(EVENT_MODIFY_STATE, FALSE, readyEventName);
    if (!readyEvent) return (int)GetLastError();

    HANDLE oldProcess = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                    FALSE, oldProcessId);
    if (!oldProcess) {
        DWORD errorCode = GetLastError();
        CloseHandle(readyEvent);
        return (int)errorCode;
    }

    wchar_t oldProcessPath[MAX_PATH];
    DWORD oldProcessPathLength = sizeof(oldProcessPath) / sizeof(wchar_t);
    if (!QueryFullProcessImageNameW(oldProcess, 0, oldProcessPath,
                                    &oldProcessPathLength)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return (int)errorCode;
    }
    if (_wcsicmp(oldProcessPath, targetPath) != 0) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_INVALID_DATA;
    }

    wchar_t helperPath[MAX_PATH];
    DWORD helperPathLength = GetModuleFileNameW(NULL, helperPath,
                                               sizeof(helperPath) / sizeof(wchar_t));
    DWORD binaryType = 0;
    if (helperPathLength == 0 || helperPathLength >= MAX_PATH ||
        !ValidateUpdateTempFilePair(helperPath, stagedPath, oldProcessId)) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_INVALID_DATA;
    }
    if (!GetBinaryTypeW(stagedPath, &binaryType)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return (int)errorCode;
    }
    if (binaryType != SCS_64BIT_BINARY) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_BAD_EXE_FORMAT;
    }

    // The helper is elevated only for file replacement. Preserve a primary
    // token from the original process so the restarted launcher normally
    // returns to the user's non-elevated session.
    HANDLE launchToken = DuplicateUpdateLaunchToken(oldProcess);

    // Only let the parent exit once this helper has verified every path and
    // owns the process handle it must wait on.
    if (!SetEvent(readyEvent)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        if (launchToken) CloseHandle(launchToken);
        return (int)errorCode;
    }
    CloseHandle(readyEvent);

    DWORD waitResult = WaitForSingleObject(oldProcess, UPDATE_HELPER_WAIT_MS);
    CloseHandle(oldProcess);
    if (waitResult != WAIT_OBJECT_0) {
        if (launchToken) CloseHandle(launchToken);
        DeleteUpdateTempFile(stagedPath);
        SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
        MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
        MessageBoxW(NULL, L"The running application did not close in time.",
                    APP_NAME L" Update", MB_OK | MB_ICONERROR);
        return ERROR_TIMEOUT;
    }

    wchar_t replacementPath[MAX_PATH], backupPath[MAX_PATH];
    DWORD helperProcessId = GetCurrentProcessId();
    int replacementLength = swprintf_s(replacementPath,
        sizeof(replacementPath) / sizeof(wchar_t), L"%s.new.%lu.exe",
        targetPath, (unsigned long)helperProcessId);
    int backupLength = swprintf_s(backupPath,
        sizeof(backupPath) / sizeof(wchar_t), L"%s.backup.%lu.exe",
        targetPath, (unsigned long)helperProcessId);
    if (replacementLength <= 0 || replacementLength >= MAX_PATH ||
        backupLength <= 0 || backupLength >= MAX_PATH) {
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The update paths were too long. The previous version will restart.");
    }

    SetFileAttributesW(replacementPath, FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(replacementPath);
    SetFileAttributesW(backupPath, FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(backupPath);
    if (!CopyFileW(stagedPath, replacementPath, FALSE)) {
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The update could not be prepared. The previous version will restart.");
    }

    DWORD targetAttributes = GetFileAttributesW(targetPath);
    BOOL clearedReadOnly = FALSE;
    if (targetAttributes != INVALID_FILE_ATTRIBUTES &&
        (targetAttributes & FILE_ATTRIBUTE_READONLY)) {
        clearedReadOnly = SetFileAttributesW(
            targetPath, targetAttributes & ~FILE_ATTRIBUTE_READONLY);
    }
    if (!ReplaceFileW(targetPath, replacementPath, backupPath,
                      REPLACEFILE_WRITE_THROUGH, NULL, NULL)) {
        if (clearedReadOnly) SetFileAttributesW(targetPath, targetAttributes);
        DeleteFileW(replacementPath);
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The executable could not be replaced. The previous version will restart.");
    }

    if (!LaunchUpdateTarget(targetPath, stagedPath, helperPath,
                            helperProcessId, oldProcessId, launchToken, TRUE)) {
        DeleteFileW(targetPath);
        if (!MoveFileExW(backupPath, targetPath,
                         MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
            if (launchToken) CloseHandle(launchToken);
            DeleteUpdateTempFile(stagedPath);
            SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
            MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
            MessageBoxW(NULL,
                L"The updated application could not start and the previous executable "
                L"could not be restored. A backup remains beside the application.",
                APP_NAME L" Update", MB_OK | MB_ICONERROR);
            return 1;
        }
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The updated application could not start. The previous version was restored.");
    }
    if (launchToken) CloseHandle(launchToken);

    if (!DeleteFileW(backupPath)) {
        MoveFileExW(backupPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    }
    DeleteUpdateTempFile(stagedPath);
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
    MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    return 0;
}

static BOOL FinishUpdateCleanup(DWORD helperProcessId, DWORD oldProcessId,
                                LPCWSTR stagedPath, LPCWSTR helperPath) {
    wchar_t targetPath[MAX_PATH], expectedStagedPath[MAX_PATH];
    wchar_t expectedHelperPath[MAX_PATH];
    DWORD targetLength = GetModuleFileNameW(NULL, targetPath,
                                           sizeof(targetPath) / sizeof(wchar_t));
    if (targetLength == 0 || targetLength >= MAX_PATH ||
        !BuildUpdateTempPath(expectedStagedPath, L"download", oldProcessId) ||
        !BuildUpdateTempPath(expectedHelperPath, L"updater", oldProcessId) ||
        _wcsicmp(stagedPath, expectedStagedPath) != 0 ||
        _wcsicmp(helperPath, expectedHelperPath) != 0) {
        return FALSE;
    }

    HANDLE helperProcess = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                       FALSE, helperProcessId);
    if (helperProcess) {
        wchar_t runningHelperPath[MAX_PATH];
        DWORD runningHelperPathLength = MAX_PATH;
        if (QueryFullProcessImageNameW(helperProcess, 0, runningHelperPath,
                                      &runningHelperPathLength) &&
            _wcsicmp(runningHelperPath, helperPath) == 0) {
            WaitForSingleObject(helperProcess, UPDATE_HELPER_WAIT_MS);
        }
        CloseHandle(helperProcess);
    }
    for (int attempt = 0;
         attempt < 20 && !DeleteUpdateTempFile(stagedPath);
         ++attempt) {
        Sleep(100);
    }
    for (int attempt = 0;
         attempt < 20 && !DeleteUpdateTempFile(helperPath);
         ++attempt) {
        Sleep(100);
    }
    return TRUE;
}

// Returns an exit code and sets handled for the temporary updater process.
// Both finish modes perform cleanup and continue normal application startup;
// updateCompleted identifies only a successful executable replacement.
static int HandleUpdateCommandLine(BOOL* handled, BOOL* updateCompleted) {
    if (handled) *handled = FALSE;
    if (updateCompleted) *updateCompleted = FALSE;
    int argumentCount = 0;
    LPWSTR* arguments = CommandLineToArgvW(GetCommandLineW(), &argumentCount);
    if (!arguments) return 0;

    int result = 0;
    if (argumentCount == 6 && wcscmp(arguments[1], L"--apply-update") == 0) {
        DWORD oldProcessId = 0;
        if (handled) *handled = TRUE;
        if (!ParseUpdateProcessId(arguments[2], &oldProcessId)) {
            result = ERROR_INVALID_PARAMETER;
        } else {
            result = RunUpdateApplyHelper(oldProcessId, arguments[3],
                                          arguments[4], arguments[5]);
        }
    } else if (argumentCount == 6 &&
               (wcscmp(arguments[1], L"--finish-update") == 0 ||
                wcscmp(arguments[1], L"--finish-update-cleanup") == 0)) {
        DWORD helperProcessId = 0, oldProcessId = 0;
        if (ParseUpdateProcessId(arguments[2], &helperProcessId) &&
            ParseUpdateProcessId(arguments[3], &oldProcessId)) {
            BOOL recognizedHandoff = FinishUpdateCleanup(
                helperProcessId, oldProcessId, arguments[4], arguments[5]);
            if (recognizedHandoff && updateCompleted &&
                wcscmp(arguments[1], L"--finish-update") == 0) {
                *updateCompleted = TRUE;
            }
        }
    }
    LocalFree(arguments);
    return result;
}

static void CfgSendUpdateResultWithVersions(LPCWSTR status, LPCWSTR title,
                                            LPCWSTR message,
                                            LPCWSTR currentVersion,
                                            LPCWSTR remoteVersion,
                                            BOOL automatic) {
    if (!g_webviewView || !status || !title || !message ||
        !currentVersion || !remoteVersion) return;
    wchar_t escapedStatus[64], escapedTitle[256], escapedMessage[1024];
    wchar_t escapedCurrentVersion[64], escapedRemoteVersion[64];
    json_escape_wstring(status, escapedStatus,
                        sizeof(escapedStatus) / sizeof(wchar_t));
    json_escape_wstring(title, escapedTitle,
                        sizeof(escapedTitle) / sizeof(wchar_t));
    json_escape_wstring(message, escapedMessage,
                        sizeof(escapedMessage) / sizeof(wchar_t));
    json_escape_wstring(currentVersion, escapedCurrentVersion,
                        sizeof(escapedCurrentVersion) / sizeof(wchar_t));
    json_escape_wstring(remoteVersion, escapedRemoteVersion,
                        sizeof(escapedRemoteVersion) / sizeof(wchar_t));

    wchar_t script[1792];
    int written = swprintf_s(script, sizeof(script) / sizeof(wchar_t),
        L"window.onUpdateResult({\"status\":\"%s\",\"title\":\"%s\","
        L"\"message\":\"%s\",\"currentVersion\":\"%s\","
        L"\"remoteVersion\":\"%s\",\"automatic\":%s})",
        escapedStatus, escapedTitle, escapedMessage,
        escapedCurrentVersion, escapedRemoteVersion,
        automatic ? L"true" : L"false");
    if (written > 0) webview_execute_script(script);
}

static void CfgSendUpdateResult(LPCWSTR status, LPCWSTR title, LPCWSTR message) {
    CfgSendUpdateResultWithVersions(status, title, message, L"", L"", FALSE);
}

static void CfgSendUpdateProgress(DWORD speedKbps) {
    wchar_t script[160];
    int written = swprintf_s(script, sizeof(script) / sizeof(wchar_t),
        L"window.onUpdateProgress({\"kilobytesPerSecond\":%lu})",
        (unsigned long)speedKbps);
    if (written > 0) webview_execute_script(script);
}

static void DiscardPendingUpdateNotice(void) {
    UpdateCheckTask* task = g_updateNoticeTask;
    g_updateNoticeTask = NULL;
    DiscardUpdateTask(task);
}

static void StartUpdateCheck(BOOL automatic) {
    HWND targetWindow = g_hwnd ? g_hwnd : g_webviewHwnd;
    if (!targetWindow) return;
    if (automatic && (g_updateNoticeTask || g_updateReadyTask)) return;
    if (InterlockedCompareExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL) {
        if (!automatic) {
            CfgSendUpdateResult(L"error", L"Update check in progress",
                L"Another update check is still finishing. Try again shortly.");
        }
        return;
    }
    if (InterlockedCompareExchange(&g_updateCheckPending, TRUE, FALSE) != FALSE) {
        if (!automatic) {
            CfgSendUpdateResult(L"error", L"Update check in progress",
                L"Another update check is still finishing. Try again shortly.");
        }
        return;
    }
    InterlockedExchange(&g_updateCheckAutomatic, automatic ? TRUE : FALSE);

    // Every accepted request starts from scratch. Automatic requests are
    // skipped above while a result is awaiting user action, avoiding an
    // hourly re-download of the same prepared executable.
    DiscardPendingUpdateNotice();
    DiscardPreparedUpdate();

    if (!g_updateCancelEvent) {
        g_updateCancelEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
        if (!g_updateCancelEvent) {
            DWORD errorCode = GetLastError();
            InterlockedExchange(&g_updateCheckPending, FALSE);
            InterlockedExchange(&g_updateCheckAutomatic, FALSE);
            wchar_t message[256];
            swprintf_s(message, sizeof(message) / sizeof(wchar_t),
                L"Could not initialize update cancellation (Windows error %lu).",
                (unsigned long)errorCode);
            CfgSendUpdateResultWithVersions(L"error", L"Update failed", message,
                                            L"", L"", automatic);
            return;
        }
    }
    ResetEvent(g_updateCancelEvent);
    InterlockedExchange(&g_updateSpeedKbps, 0);
    InterlockedExchange(&g_updateProgressPosted, FALSE);

    UpdateCheckTask* task = (UpdateCheckTask*)calloc(1, sizeof(UpdateCheckTask));
    if (!task) {
        InterlockedExchange(&g_updateCheckPending, FALSE);
        InterlockedExchange(&g_updateCheckAutomatic, FALSE);
        CfgSendUpdateResultWithVersions(L"error", L"Update failed",
            L"There was not enough memory to check for updates.",
            L"", L"", automatic);
        return;
    }
    task->targetWindow = targetWindow;
    task->automatic = automatic;
    LONG sequence = InterlockedIncrement(&g_updateRequestSequence);
    task->cacheBuster =
        ((GetTickCount64() ^ GetCurrentProcessId()) << 32) | (DWORD)sequence;
    if (task->cacheBuster == 0) task->cacheBuster = 1;

    HANDLE thread = CreateThread(NULL, 0, UpdateCheckThread, task, 0, NULL);
    if (!thread) {
        DWORD errorCode = GetLastError();
        free(task);
        InterlockedExchange(&g_updateCheckPending, FALSE);
        InterlockedExchange(&g_updateCheckAutomatic, FALSE);
        wchar_t message[256];
        swprintf_s(message, sizeof(message) / sizeof(wchar_t),
                   L"Could not start the update check (Windows error %lu).",
                   (unsigned long)errorCode);
        CfgSendUpdateResultWithVersions(L"error", L"Update failed", message,
                                        L"", L"", automatic);
        return;
    }
    CloseHandle(thread);
}

static BOOL IsIgnoredUpdateVersion(const ExecutableVersion* version) {
    wchar_t formatted[32];
    if (!version || !g_ignoredUpdateVersion[0]) return FALSE;
    FormatExecutableVersion(version, formatted,
                            sizeof(formatted) / sizeof(wchar_t));
    return wcscmp(formatted, g_ignoredUpdateVersion) == 0;
}

static void SaveIgnoredUpdateVersion(LPCWSTR version) {
    HKEY key;
    DWORD disposition;
    if (!version) return;
    wcsncpy_s(g_ignoredUpdateVersion,
              sizeof(g_ignoredUpdateVersion) / sizeof(wchar_t), version,
              _TRUNCATE);
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_KEY_PATH_W, 0, NULL,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &key,
                        &disposition) == ERROR_SUCCESS) {
        RegSetValueExW(key, REG_VALUE_IGNORED_UPDATE_VERSION_W, 0, REG_SZ,
                       (const BYTE*)g_ignoredUpdateVersion,
                       (DWORD)((wcslen(g_ignoredUpdateVersion) + 1) *
                               sizeof(wchar_t)));
        RegCloseKey(key);
    }
}

static void PresentPendingUpdateNotice(void) {
    if (!g_configViewReady || !g_webviewView || !g_updateNoticeTask) return;

    UpdateCheckTask* task = g_updateNoticeTask;
    g_updateNoticeTask = NULL;
    LPCWSTR status = NULL;
    LPCWSTR title = NULL;
    LPCWSTR message = NULL;
    wchar_t currentVersion[32] = L"";
    wchar_t remoteVersion[32] = L"";
    BOOL installable = FALSE;

    if (task->kind == UPDATE_CHECK_CANCELLED) {
        status = L"cancelled";
        title = L"";
        message = L"";
    } else if (task->kind == UPDATE_CHECK_ERROR) {
        status = L"error";
        title = L"Update failed";
        message = task->message;
    } else {
        FormatExecutableVersionForDisplay(
            &task->runningVersion, currentVersion,
            sizeof(currentVersion) / sizeof(wchar_t));
        FormatExecutableVersionForDisplay(
            &task->availableVersion, remoteVersion,
            sizeof(remoteVersion) / sizeof(wchar_t));
        if (task->kind == UPDATE_CHECK_NEWER) {
            status = L"newer";
            title = L"Update available";
            message = L"A newer version is ready to install.";
            installable = TRUE;
        } else if (task->kind == UPDATE_CHECK_SAME) {
            status = L"same";
            title = L"You're up to date";
            message = L"The remote build matches your current version. "
                      L"You can force a reinstall if needed.";
            installable = !task->automatic;
        } else if (task->kind == UPDATE_CHECK_OLDER) {
            status = L"older";
            title = L"No update available";
            message = L"The remote build is older than your current version.";
        }
    }

    if (!status) {
        DiscardUpdateTask(task);
        return;
    }
    if (installable) {
        DiscardPreparedUpdate();
        g_updateReadyTask = task;
    }
    CfgSendUpdateResultWithVersions(status, title, message,
                                    currentVersion, remoteVersion,
                                    task->automatic);
    if (!installable) DiscardUpdateTask(task);
}

static void QueueUpdateNotice(UpdateCheckTask* task) {
    DiscardPendingUpdateNotice();
    g_updateNoticeTask = task;
    PresentPendingUpdateNotice();
}

static void HandleCompletedUpdateCheck(UpdateCheckTask* task) {
    if (!task) return;
    InterlockedExchange(&g_updateProgressPosted, FALSE);
    InterlockedExchange(&g_updateSpeedKbps, 0);

    if (task->kind == UPDATE_CHECK_CANCELLED) {
        LogUpdateMessage(L"[INFO] Update check cancelled\n");
    } else if (task->kind == UPDATE_CHECK_ERROR) {
        LogUpdateMessage(L"[WARNING] Update check failed: %s\n", task->message);
    }

    if (task->automatic && !configAutoCheckForUpdates) {
        DiscardUpdateTask(task);
        return;
    }

    if (task->automatic && task->kind == UPDATE_CHECK_NEWER &&
        IsIgnoredUpdateVersion(&task->availableVersion)) {
        LogUpdateMessage(L"[INFO] Automatic update prompt suppressed for ignored version\n");
        task->kind = UPDATE_CHECK_CANCELLED;
        if (g_webviewHwnd) {
            QueueUpdateNotice(task);
        } else {
            DiscardUpdateTask(task);
        }
        return;
    }

    if (task->automatic && task->kind == UPDATE_CHECK_NEWER) {
        if (g_webviewHwnd && strcmp(g_pendingView, "config") != 0) {
            SendMessageW(g_webviewHwnd, WM_CLOSE, 0, 0);
        }
        QueueUpdateNotice(task);
        ShowConfigDialog(g_hwnd);
        if (!g_webviewHwnd) DiscardPendingUpdateNotice();
        return;
    }

    if (g_webviewHwnd) {
        QueueUpdateNotice(task);
    } else {
        DiscardUpdateTask(task);
    }
}

static void IgnorePreparedUpdateVersion(const char* requestedVersion) {
    UpdateCheckTask* task = g_updateReadyTask;
    wchar_t preparedVersion[32];
    wchar_t preparedDisplayVersion[32];
    wchar_t requestedVersionW[32] = L"";
    if (!task || !task->automatic || task->kind != UPDATE_CHECK_NEWER ||
        !requestedVersion ||
        !MultiByteToWideChar(CP_UTF8, 0, requestedVersion, -1,
                             requestedVersionW,
                             sizeof(requestedVersionW) / sizeof(wchar_t))) {
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The update version could not be ignored. Check for updates again.");
        return;
    }
    FormatExecutableVersion(&task->availableVersion, preparedVersion,
                            sizeof(preparedVersion) / sizeof(wchar_t));
    FormatExecutableVersionForDisplay(
        &task->availableVersion, preparedDisplayVersion,
        sizeof(preparedDisplayVersion) / sizeof(wchar_t));
    if (wcscmp(preparedDisplayVersion, requestedVersionW) != 0) {
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The update version changed. Check for updates again.");
        return;
    }
    SaveIgnoredUpdateVersion(preparedVersion);
    LogUpdateMessage(L"[INFO] Automatic update version added to the ignore list\n");
    DiscardPreparedUpdate();
}

static void CancelUpdateCheck(void) {
    if (InterlockedCompareExchange(&g_updateCheckPending, FALSE, FALSE) == TRUE &&
        g_updateCancelEvent) {
        LogUpdateMessage(L"[INFO] Update check cancellation requested\n");
        SetEvent(g_updateCancelEvent);
    }
}

static HANDLE CreateUpdateReadyEvent(DWORD processId, wchar_t* eventName,
                                     size_t eventNameCch) {
    if (!processId || !eventName || eventNameCch < 96) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return NULL;
    }

    ULONGLONG nonce = GetTickCount64() ^
                      ((ULONGLONG)GetCurrentThreadId() << 32);
    BCryptGenRandom(NULL, (PUCHAR)&nonce, sizeof(nonce),
                    BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    int nameLength = swprintf_s(eventName, eventNameCch,
        L"Local\\APIMonitor_UpdateReady_%lu_%016llx",
        (unsigned long)processId, (unsigned long long)nonce);
    if (nameLength <= 0 || nameLength >= (int)eventNameCch) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return NULL;
    }

    // The elevated helper can run with a split administrator token (or with
    // alternate administrator credentials). Grant interactive users access
    // to this random, session-local event so either UAC path can acknowledge
    // readiness without exposing any file or process permissions.
    PSECURITY_DESCRIPTOR descriptor = NULL;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
            L"D:P(A;;GA;;;IU)(A;;GA;;;BA)(A;;GA;;;SY)",
            SDDL_REVISION_1, &descriptor, NULL)) {
        return NULL;
    }
    SECURITY_ATTRIBUTES securityAttributes = {0};
    securityAttributes.nLength = sizeof(securityAttributes);
    securityAttributes.lpSecurityDescriptor = descriptor;

    HANDLE readyEvent = CreateEventW(&securityAttributes, TRUE, FALSE,
                                     eventName);
    DWORD errorCode = readyEvent ? ERROR_SUCCESS : GetLastError();
    LocalFree(descriptor);
    if (!readyEvent) SetLastError(errorCode);
    return readyEvent;
}

static BOOL LaunchStagedUpdate(LPCWSTR stagedPath, LPCWSTR targetPath) {
    if (!stagedPath || !targetPath || !*stagedPath || !*targetPath) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD oldProcessId = GetCurrentProcessId();
    wchar_t helperPath[MAX_PATH];
    if (!BuildUpdateTempPath(helperPath, L"updater", oldProcessId)) {
        return FALSE;
    }
    DeleteUpdateTempFile(helperPath);
    if (!CopyFileW(targetPath, helperPath, TRUE)) return FALSE;
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);

    // CopyFile preserves alternate data streams. The source is already the
    // running, user-approved executable, so do not carry its download-zone
    // marker onto the short-lived updater copy and trigger a second warning.
    wchar_t zonePath[MAX_PATH + 32];
    if (swprintf_s(zonePath, sizeof(zonePath) / sizeof(wchar_t),
                   L"%s:Zone.Identifier", helperPath) > 0) {
        DeleteFileW(zonePath);
    }

    wchar_t readyEventName[160];
    HANDLE readyEvent = CreateUpdateReadyEvent(oldProcessId, readyEventName,
        sizeof(readyEventName) / sizeof(wchar_t));
    if (!readyEvent) {
        DWORD errorCode = GetLastError();
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }

    wchar_t parameters[MAX_PATH * 2 + 512];
    int parameterLength = swprintf_s(parameters,
        sizeof(parameters) / sizeof(wchar_t),
        L"--apply-update %lu \"%s\" \"%s\" \"%s\"",
        (unsigned long)oldProcessId, readyEventName, targetPath, stagedPath);
    if (parameterLength <= 0 ||
        parameterLength >= (int)(sizeof(parameters) / sizeof(wchar_t))) {
        CloseHandle(readyEvent);
        DeleteUpdateTempFile(helperPath);
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    SHELLEXECUTEINFOW executeInfo = {0};
    executeInfo.cbSize = sizeof(executeInfo);
    executeInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
    executeInfo.hwnd = g_webviewHwnd;
    executeInfo.lpVerb = L"runas";
    executeInfo.lpFile = helperPath;
    executeInfo.lpParameters = parameters;
    executeInfo.nShow = SW_HIDE;
    BOOL elevated = ShellExecuteExW(&executeInfo);
    if (!elevated || !executeInfo.hProcess) {
        DWORD errorCode = elevated ? ERROR_INVALID_HANDLE : GetLastError();
        if (!errorCode) errorCode = ERROR_ACCESS_DENIED;
        CloseHandle(readyEvent);
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }

    HANDLE waitHandles[2] = { readyEvent, executeInfo.hProcess };
    DWORD waitResult = WaitForMultipleObjects(2, waitHandles, FALSE,
                                              UPDATE_HELPER_READY_MS);
    DWORD errorCode = ERROR_SUCCESS;
    if (waitResult != WAIT_OBJECT_0) {
        if (waitResult == WAIT_OBJECT_0 + 1) {
            DWORD exitCode = ERROR_INSTALL_FAILURE;
            if (!GetExitCodeProcess(executeInfo.hProcess, &exitCode) ||
                exitCode == ERROR_SUCCESS || exitCode == STILL_ACTIVE) {
                exitCode = ERROR_INSTALL_FAILURE;
            }
            errorCode = exitCode;
        } else {
            errorCode = waitResult == WAIT_TIMEOUT ? ERROR_TIMEOUT
                                                   : GetLastError();
            if (!errorCode) errorCode = ERROR_INSTALL_FAILURE;
        }
    }
    CloseHandle(executeInfo.hProcess);
    CloseHandle(readyEvent);

    if (waitResult != WAIT_OBJECT_0) {
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }
    return TRUE;
}

static void DiscardPreparedUpdate(void) {
    UpdateCheckTask* task = g_updateReadyTask;
    g_updateReadyTask = NULL;
    DiscardUpdateTask(task);
}

static void InstallPreparedUpdate(void) {
    UpdateCheckTask* task = g_updateReadyTask;
    g_updateReadyTask = NULL;
    if (!task || (task->kind != UPDATE_CHECK_NEWER &&
                  task->kind != UPDATE_CHECK_SAME)) {
        DiscardUpdateTask(task);
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The prepared update is no longer available. Check for updates again.");
        return;
    }

    if (LaunchStagedUpdate(task->stagedPath, task->targetPath)) {
        LogUpdateMessage(L"[INFO] Update accepted; exiting for replacement\n");
        g_updateInstallReady = TRUE;
        free(task);  // The updater process now owns the staged file.
        if (g_webviewHwnd) PostMessageW(g_webviewHwnd, WM_CLOSE, 0, 0);
        return;
    }

    DWORD errorCode = GetLastError();
    wchar_t message[384];
    LPCWSTR title = L"Update failed";
    if (errorCode == ERROR_CANCELLED) {
        title = L"Update cancelled";
        wcscpy_s(message, sizeof(message) / sizeof(wchar_t),
            L"Administrator approval was cancelled. Your current version is still running.");
    } else {
        swprintf_s(message, sizeof(message) / sizeof(wchar_t),
            L"The elevated update process could not be started (Windows error %lu).",
            (unsigned long)errorCode);
    }
    LogUpdateMessage(L"[WARNING] %s\n", message);
    CfgSendUpdateResult(L"error", title, message);
    DiscardUpdateTask(task);
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
    wchar_t wLogPath[1024];
    json_escape_string(logFilePath, wLogPath, 1024);
    wchar_t wWebView2Version[256];
    wchar_t wUpdateCompletedVersion[64];
    json_escape_wstring(g_webView2Version, wWebView2Version,
                        sizeof(wWebView2Version) / sizeof(wchar_t));
    json_escape_wstring(g_updateConfirmationPending ? APP_VERSION_WSTRING : L"",
                        wUpdateCompletedVersion,
                        sizeof(wUpdateCompletedVersion) / sizeof(wchar_t));
    BOOL updatePending =
        InterlockedCompareExchange(&g_updateCheckPending, FALSE, FALSE) == TRUE ||
        InterlockedCompareExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL;

    wchar_t script[4608];
    swprintf(script, sizeof(script) / sizeof(wchar_t),
        L"window.onInit({\"view\":\"config\",\"config\":{\"url\":\"%s\",\"interval\":%d,\"loggingEnabled\":%s,\"historyLimit\":%d,\"logPath\":\"%s\",\"autoCheckForUpdates\":%s,\"updateCheckPending\":%s,\"updatePromptPending\":%s},\"webView2Version\":\"%s\",\"updateCompletedVersion\":\"%s\"})",
        wUrl, configRefreshInterval,
        configLoggingEnabled ? L"true" : L"false",
        configHistoryLimit, wLogPath,
        configAutoCheckForUpdates ? L"true" : L"false",
        updatePending ? L"true" : L"false",
        g_updateNoticeTask ? L"true" : L"false",
        wWebView2Version, wUpdateCompletedVersion);
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

    LPWSTR browserVersion = NULL;
    if (SUCCEEDED(env->lpVtbl->get_BrowserVersionString(env, &browserVersion)) &&
        browserVersion && browserVersion[0]) {
        wcsncpy_s(g_webView2Version,
                  sizeof(g_webView2Version) / sizeof(wchar_t),
                  browserVersion, _TRUNCATE);
    }
    CoTaskMemFree(browserVersion);

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
    controller->lpVtbl->put_IsVisible(controller, TRUE);

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
    } else if (strcmp(action, "configReady") == 0) {
        if (g_webviewHwnd && strcmp(g_pendingView, "config") == 0) {
            BOOL updateWorkAlreadyActive =
                InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE ||
                InterlockedCompareExchangePointer(
                    (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL ||
                g_updateNoticeTask || g_updateReadyTask;
            BOOL checkAutomatically = FALSE;
            json_get_bool(msg, "checkAutomatically", &checkAutomatically);
            g_configViewReady = TRUE;
            PresentPendingUpdateNotice();
            if (checkAutomatically && configAutoCheckForUpdates &&
                !g_updateConfirmationPending && !updateWorkAlreadyActive) {
                StartUpdateCheck(TRUE);
            }
        }
    } else if (strcmp(action, "checkUpdate") == 0) {
        BOOL automatic = FALSE;
        json_get_bool(msg, "automatic", &automatic);
        StartUpdateCheck(automatic);
    } else if (strcmp(action, "cancelUpdateCheck") == 0) {
        CancelUpdateCheck();
    } else if (strcmp(action, "installUpdate") == 0) {
        InstallPreparedUpdate();
    } else if (strcmp(action, "dismissUpdate") == 0) {
        DiscardPreparedUpdate();
    } else if (strcmp(action, "ignoreUpdateVersion") == 0) {
        char version[32] = {0};
        json_get_string(msg, "version", version, sizeof(version));
        IgnorePreparedUpdateVersion(version);
    } else if (strcmp(action, "dismissUpdateConfirmation") == 0) {
        g_updateConfirmationPending = FALSE;
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
        BOOL autoUpdate = configAutoCheckForUpdates;
        int histLimit = 100;
        json_get_string(msg, "url", url, sizeof(url));
        json_get_int(msg, "interval", &interval);
        json_get_bool(msg, "loggingEnabled", &logging);
        json_get_bool(msg, "autoCheckForUpdates", &autoUpdate);
        json_get_int(msg, "historyLimit", &histLimit);

        if (url[0]) {
            strncpy(configApiUrl, url, sizeof(configApiUrl) - 1);
            configApiUrl[sizeof(configApiUrl) - 1] = '\0';
        }
        if (interval == 60 || interval == 120 || interval == 300) {
            configRefreshInterval = interval;
        }
        configLoggingEnabled = logging;
        configAutoCheckForUpdates = autoUpdate;
        if (histLimit >= 10 && histLimit <= 10000) {
            configHistoryLimit = histLimit;
            InitHistoryBuffer(configHistoryLimit);
        }

        SaveConfigToRegistry();
        MarkAsConfigured();
        ApplyConfiguration();
        LogMessage("Configuration updated via WebView dialog: URL=%s, Interval=%d, Logging=%s, HistoryLimit=%d, AutoUpdate=%s",
                   configApiUrl, configRefreshInterval,
                   configLoggingEnabled ? "enabled" : "disabled", configHistoryLimit,
                   configAutoCheckForUpdates ? "enabled" : "disabled");
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
            UINT flags = SWP_NOMOVE | SWP_NOZORDER;
            if (g_webviewWindowShown) {
                flags |= SWP_NOACTIVATE;
            } else {
                flags |= SWP_SHOWWINDOW;
                KillTimer(g_webviewHwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            }
            SetWindowPos(g_webviewHwnd, NULL, 0, 0, windowW, newWindowH, flags);
            g_webviewWindowShown = TRUE;
            webview_sync_controller_bounds();
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
        case WM_APP_UPDATE_PROGRESS:
            InterlockedExchange(&g_updateProgressPosted, FALSE);
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                g_configViewReady) {
                DWORD speedKbps = (DWORD)InterlockedCompareExchange(
                    &g_updateSpeedKbps, 0, 0);
                CfgSendUpdateProgress(speedKbps);
            }
            return 0;

        case WM_APP_UPDATE_RESULT: {
            UpdateCheckTask* task = (UpdateCheckTask*)InterlockedExchangePointer(
                (PVOID volatile*)&g_updatePostedResult, NULL);
            HandleCompletedUpdateCheck(task);
            return 0;
        }

        case WM_SIZE:
            webview_sync_controller_bounds();
            return 0;

        case WM_VALIDATE_RESULT:
            if ((LONG)wParam == g_validateGeneration) {
                webview_push_validation_result(lParam ? TRUE : FALSE);
            }
            return 0;

        case WM_TIMER:
            if (wParam == ID_TIMER_WEBVIEW_SHOW_FALLBACK) {
                KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
                if (!g_webviewWindowShown) {
                    ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                    UpdateWindow(hwnd);
                    g_webviewWindowShown = TRUE;
                    webview_sync_controller_bounds();
                }
                return 0;
            }
            break;

        case WM_CLOSE:
            g_webviewWindowShown = FALSE;
            KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
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
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                InterlockedCompareExchange(&g_updateCheckAutomatic,
                                           FALSE, FALSE) == FALSE &&
                g_updateCancelEvent) {
                SetEvent(g_updateCancelEvent);
            }
            DiscardPendingUpdateNotice();
            DiscardPreparedUpdate();
            g_webviewHwnd = NULL;
            g_webviewWindowShown = FALSE;
            g_configViewReady = FALSE;
            KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            if (g_updateInstallReady) PostQuitMessage(0);
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
        return;
    }

    strncpy(g_pendingView, view, sizeof(g_pendingView) - 1);
    g_pendingView[sizeof(g_pendingView) - 1] = '\0';
    g_configViewReady = FALSE;

    // Register window class (once)
    static BOOL classRegistered = FALSE;
    if (!classRegistered) {
        WNDCLASSEXW wc = {0};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = WebViewWndProc;
        wc.hInstance = g_hInstance;
        wc.hIcon = (HICON)LoadImageW(g_hInstance, MAKEINTRESOURCEW(IDI_APPICON),
            IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"APIMonitorWebViewWnd";
        wc.hIconSm = (HICON)LoadImageW(g_hInstance, MAKEINTRESOURCEW(IDI_APPICON),
            IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
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
    g_webviewWindowShown = FALSE;
    SetTimer(g_webviewHwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK, WEBVIEW_SHOW_FALLBACK_DELAY_MS, NULL);

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

    KillTimer(hwnd, ID_TIMER_AUTO_UPDATE);
    if (g_updateCancelEvent) SetEvent(g_updateCancelEvent);
    DiscardUpdateTask((UpdateCheckTask*)InterlockedExchangePointer(
        (PVOID volatile*)&g_updatePostedResult, NULL));
    DiscardPendingUpdateNotice();
    DiscardPreparedUpdate();

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
