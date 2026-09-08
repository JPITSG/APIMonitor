// Linux harness: execute production timer/parser with deterministic OS stubs.
#include <assert.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <wchar.h>
typedef int BOOL;
typedef uint32_t DWORD, UINT;
typedef uintptr_t UINT_PTR, HWND;
typedef int64_t LONG64;
typedef uint64_t ULONGLONG;
typedef const wchar_t *LPCWSTR;
typedef wchar_t *LPWSTR;
#define FALSE 0
#define TRUE 1
#define CALLBACK
#define MAXLONG INT32_MAX
#define ERROR_INVALID_PARAMETER 87
#define WAIT_OBJECT_0 0
#define UPDATE_PROGRESS_INTERVAL_MS 250
#define UNREFERENCED_PARAMETER(x) (void)(x)
static int argcMock, cleanupValid, applyCalls, applyReopen, cleanupCalls;
static LPWSTR *argvMock;
static LPWSTR GetCommandLineW(void) { return NULL; }
static LPWSTR *CommandLineToArgvW(LPWSTR text, int *count) {
    *count = argcMock;
    return argvMock;
}
static void LocalFree(void *pointer) {}
static int RunUpdateApplyHelper(DWORD pid, LPCWSTR event, LPCWSTR target,
                                LPCWSTR staged, BOOL reopen) {
    applyCalls++;
    applyReopen = reopen;
    return 42;
}
static BOOL FinishUpdateCleanup(DWORD helper, DWORD old, LPCWSTR staged,
                               LPCWSTR path) {
    cleanupCalls++;
    return cleanupValid;
}
static int g_updateCheckPending, g_updateCancelEvent, g_configViewReady;
static volatile LONG64 g_updateTransferStarted, g_updateReceivedBytes;
static ULONGLONG g_updateSampleTime, g_updateSampleBytes, nowMock;
static DWORD lastSpeed;
static int sent, killed;
static int InterlockedCompareExchange(int *value, int swap, int compare) {
    return *value;
}
static LONG64 InterlockedCompareExchange64(volatile LONG64 *value,
                                          LONG64 swap, LONG64 compare) {
    return *value;
}
static DWORD WaitForSingleObject(int event, DWORD timeout) { return 0; }
static void KillTimer(HWND hwnd, UINT_PTR id) { killed++; }
static ULONGLONG GetTickCount64(void) { return nowMock; }
static void CfgSendUpdateProgress(DWORD speed) { lastSpeed = speed; sent++; }

/* PRODUCTION FUNCTIONS */

static void command(LPCWSTR action, LPCWSTR flag, int count, BOOL recognized,
                    BOOL expectedHandled, BOOL expectedComplete, BOOL expectedReopen) {
    LPWSTR args[] = {L"app.exe", (LPWSTR)action, L"123", L"456",
                     L"C:\\Temp\\download.exe", L"C:\\Temp\\helper.exe", (LPWSTR)flag};
    argcMock = count;
    argvMock = args;
    cleanupValid = recognized;
    applyCalls = cleanupCalls = 0;
    BOOL handled = TRUE, complete = TRUE, reopen = TRUE;
    HandleUpdateCommandLine(&handled, &complete, &reopen);
    assert(handled == expectedHandled);
    assert(complete == expectedComplete);
    assert(reopen == expectedReopen);
}

int main(void) {
    command(L"--finish-update", L"--reopen-settings", 7, TRUE, FALSE, TRUE, TRUE);
    command(L"--finish-update", L"", 6, TRUE, FALSE, TRUE, FALSE);
    command(L"--finish-update", L"--reopen-settings", 7, FALSE, FALSE, FALSE, FALSE);
    command(L"--finish-update-cleanup", L"--reopen-settings", 7, TRUE, FALSE, FALSE, FALSE);
    command(L"--finish-update-cleanup", L"", 6, TRUE, FALSE, FALSE, FALSE);
    command(L"--apply-update", L"--reopen-settings", 7, TRUE, TRUE, FALSE, FALSE);
    assert(applyCalls == 1 && applyReopen);
    command(L"--apply-update", L"", 6, TRUE, TRUE, FALSE, FALSE);
    assert(applyCalls == 1 && !applyReopen);
    command(L"--finish-update", L"--unknown", 7, TRUE, FALSE, FALSE, FALSE);
    assert(!cleanupCalls);
    command(L"--reopen-settings", L"", 2, TRUE, FALSE, FALSE, FALSE);
    command(L"", L"", 1, TRUE, FALSE, FALSE, FALSE);
    command(L"--finish-update", L"--reopen-settings", 5, TRUE, FALSE, FALSE, FALSE);
    LPWSTR invalid[] = {L"app", L"--finish-update", L"bad", L"456", L"x", L"y", L"--reopen-settings"};
    argvMock = invalid; argcMock = 7;
    BOOL handled, complete, reopen;
    HandleUpdateCommandLine(&handled, &complete, &reopen);
    assert(!handled && !complete && !reopen);

    g_updateCheckPending = g_configViewReady = TRUE;
    nowMock = 1000;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(!sent); // HEAD phase.
    g_updateTransferStarted = 1000;
    g_updateReceivedBytes = 25600;
    nowMock = 1249;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(!sent); // Throttle.
    nowMock = 1250;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 1 && lastSpeed == 100);
    nowMock = 1500; g_updateReceivedBytes += 128;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(lastSpeed == 1); // 0.5 rounds up.
    nowMock = 1750; g_updateReceivedBytes += 127;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(lastSpeed == 0); // Below 0.5 rounds down.
    nowMock = 2000;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(lastSpeed == 0); // Blocked read/stall.
    nowMock = 3000; g_updateReceivedBytes += 102400;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(lastSpeed == 100); // Delayed timer uses actual elapsed time.
    int before = sent;
    g_updateCancelEvent = TRUE; nowMock += 250;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == before);
    g_updateCancelEvent = FALSE; g_updateTransferStarted = 0;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == before); // Validation/finished transfer.
    g_updateCheckPending = FALSE;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(killed == 1 && sent == before);
    puts("Native speed and updater handoff tests passed");
    return 0;
}
