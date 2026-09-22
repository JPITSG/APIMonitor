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
#define ERROR_INVALID_PARAMETER 87
#define WAIT_OBJECT_0 0
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
static volatile LONG64 g_updateExpectedBytes, g_updateReceivedBytes;
static DWORD lastPercent;
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
static void CfgSendUpdateProgress(DWORD percent) { lastPercent = percent; sent++; }

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

    assert(CalculateUpdateProgressPercent(0, 1000) == 0);
    assert(CalculateUpdateProgressPercent(9, 1000) == 0);
    assert(CalculateUpdateProgressPercent(10, 1000) == 1);
    assert(CalculateUpdateProgressPercent(1, 3) == 33);
    assert(CalculateUpdateProgressPercent(2, 3) == 66); // Rounds down.
    assert(CalculateUpdateProgressPercent(999, 1000) == 99); // Never an early 100.
    assert(CalculateUpdateProgressPercent(1000, 1000) == 100);
    assert(CalculateUpdateProgressPercent(2000, 1000) == 100);
    assert(CalculateUpdateProgressPercent(5, 0) == 0);
    assert(CalculateUpdateProgressPercent(100ULL * 1024 * 1024 - 1,
                                          100ULL * 1024 * 1024) == 99);
    assert(CalculateUpdateProgressPercent(100ULL * 1024 * 1024,
                                          100ULL * 1024 * 1024) == 100);

    g_updateCheckPending = g_configViewReady = TRUE;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(!sent); // HEAD phase.
    g_updateExpectedBytes = 1000;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 1 && lastPercent == 0); // Body transfer started.
    g_updateReceivedBytes = 429;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 2 && lastPercent == 42);
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 3 && lastPercent == 42); // Stall resends for a late dialog.
    g_configViewReady = FALSE;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 3); // No settings dialog.
    g_configViewReady = TRUE;
    g_updateReceivedBytes = 1000;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 4 && lastPercent == 100); // Kept while validating.
    g_updateCancelEvent = TRUE;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(sent == 4); // Stopping.
    g_updateCancelEvent = FALSE;
    g_updateCheckPending = FALSE;
    UpdateProgressTimer(0, 0, 0, 0);
    assert(killed == 1 && sent == 4);
    puts("Native progress and updater handoff tests passed");
    return 0;
}
