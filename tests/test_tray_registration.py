"""Run the production tray lifecycle against a controllable shell on Linux."""
from pathlib import Path
import re
import struct
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / "main.c").read_text()
PROC = "WindowProc"


def function(name):
    match = re.search(r"^(?:static )?(?:void|BOOL|UINT|LRESULT CALLBACK) " + name +
                      r"\([^;{]*\{.*?^}", SOURCE, re.M | re.S)
    if not match:
        raise AssertionError("Missing production function: " + name)
    return match.group()


class TrayRegistrationTests(unittest.TestCase):
    def test_shell_lifecycle(self):
        # Use the real message interception too: a killed timer's queued tick
        # and an Explorer broadcast must follow the same guards as production.
        messages = function(PROC).split("{", 1)[1].split("    switch (", 1)[0]
        messages = re.sub(r"\buMsg\b", "msg", messages)
        stubs = r'''
#include <assert.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
typedef unsigned UINT;
typedef unsigned DWORD;
typedef uintptr_t WPARAM;
typedef long LRESULT;
typedef void* HWND;
#define TRUE 1
#define FALSE 0
#define NIM_ADD 0
#define NIM_MODIFY 1
#define NIM_DELETE 2
#define NIF_ICON 2
#define NIF_MESSAGE 1
#define NIF_TIP 4
#define WM_TIMER 0x113
#define LogMessage(...) ((void)0)
#define DebugPrint(...) ((void)0)
#define OutputDebugStringW(...) ((void)0)
typedef struct {
    UINT cbSize, uID, uFlags, uCallbackMessage;
    HWND hWnd;
    void* hIcon;
    wchar_t szTip[128];
} NOTIFYICONDATAW;
typedef NOTIFYICONDATAW NOTIFYICONDATAA;
static NOTIFYICONDATAW cache, published;
#define g_nid cache
#define nid cache
static BOOL g_trayActive, g_trayRegistered, g_trayRetryPending;
typedef unsigned long long ULONGLONG;
static BOOL g_trayBlinkHidden;
static UINT g_trayBlinkInterval;
static ULONGLONG g_trayBlinkMs;
static int configFastBlinkMinutes = FAST_BLINK_DEFAULT_MINUTES;
static void* hIconEmpty = (void*)1001;
static void* hIconFail = (void*)1002;
static void* hIconSuccess = (void*)1003;
static void* hIconBlank = (void*)1004;
static UINT g_WM_TASKBARCREATED = 0xc001;
static int ready, held, timer, failTimer, arms, calls, deletes, blinkTimer, blinkArms;
static UINT blinkDelay; /* Interval of the running blink timer. */
static int iconOnly; /* Expect (and count) blink ticks, which never resend the tooltip. */
static DWORD commands[4096];
static BOOL Notify(DWORD command, NOTIFYICONDATAW* data) {
    assert(calls < 4096);
    commands[calls++] = command;
    assert(data->hWnd == (HWND)1 && data->uID == 17);
    if (command == NIM_DELETE) { held = 0; deletes++; return TRUE; }
    if (data->uFlags == NIF_ICON) {
        assert(command == NIM_MODIFY && g_trayBlinkInterval); /* Never a partial ADD. */
        iconOnly++;
        if (!ready || !held) return FALSE;
        published.hIcon = data->hIcon;
        return TRUE;
    }
    assert(data->uFlags == (NIF_ICON | NIF_MESSAGE | NIF_TIP));
    assert(data->uCallbackMessage == 0x8001);
    if (!ready || (command == NIM_ADD && held) || (command == NIM_MODIFY && !held)) return FALSE;
    held = 1;
    published = *data;
    return TRUE;
}
#define Shell_NotifyIconW Notify
#define Shell_NotifyIconA Notify
static UINT SetTimer(HWND window, UINT id, UINT delay, void* callback) {
    assert(window == (HWND)1 && callback == NULL);
    if (id == ID_TIMER_TRAY_BLINK) {
        assert(delay == TRAY_BLINK_MS || delay == TRAY_BLINK_FAST_MS);
        if (failTimer) return 0;
        blinkArms++;
        blinkTimer = 1;
        blinkDelay = delay;
        return id;
    }
    assert(id == ID_TIMER_TRAY_RETRY && delay == 2000);
    arms++;
    if (failTimer) return 0;
    timer = 1;
    return id;
}
static void KillTimer(HWND window, UINT id) {
    assert(window == (HWND)1);
    if (id == ID_TIMER_TRAY_BLINK) { blinkTimer = 0; blinkDelay = 0; return; }
    assert(id == ID_TIMER_TRAY_RETRY);
    timer = 0;
}
/* A check reporting its icon, as ApplyTrayUpdate does on the window thread. */
static void Report(void* icon);
#define TICK() TrayMessages(WM_TIMER, ID_TIMER_TRAY_BLINK)
#define TICKS(minutes) ((minutes) * 60 * 1000 / TRAY_BLINK_MS)
/* Save the dialog with another choice, as ApplyConfiguration does. */
#define CHOOSE(minutes) (configFastBlinkMinutes = (minutes), UpdateTrayBlinkRate())
/* Blink count times at the current rate, checking that the rate never changes early. */
#define BLINK(count, rate) do { \
    for (int n = 0; n < (int)(count); n++) { \
        assert(blinkTimer && blinkDelay == (rate) && g_trayBlinkInterval == (rate)); \
        BOOL hidden = g_trayBlinkHidden; \
        calls = 0; \
        TICK(); \
        assert(published.hIcon == (hidden ? cache.hIcon : hIconBlank)); \
    } \
} while (0)
'''
        assertions = r'''
int main(void) {
    cache.hWnd = (HWND)1;
    cache.uID = 17;
    cache.uCallbackMessage = 0x8001;
    cache.hIcon = (void*)100;
    cache.uFlags = NIF_TIP; /* Last ordinary update was tooltip-only. */
    wcscpy(cache.szTip, L"initial");
    TrayMessages(g_WM_TASKBARCREATED, 0);
    assert(calls == 0); /* Broadcast before initialization. */
    g_trayActive = TRUE;
    PublishTrayIcon(); /* Explorer is not ready at logon. */
    assert(!g_trayRegistered && g_trayRetryPending && timer && arms == 1);
    assert(commands[0] == NIM_ADD && commands[1] == NIM_MODIFY);
    for (int i = 0; i < 120; i++) TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(arms == 1 && timer); /* No retry limit, no timer reset/starvation. */
    cache.hIcon = (void*)200;
    wcscpy(cache.szTip, L"latest status");
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(g_trayRegistered && !g_trayRetryPending && !timer);
    assert(published.hIcon == (void*)200 && wcscmp(published.szTip, L"latest status") == 0);
    int before = calls;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY); /* Queued stale tick. */
    assert(calls == before);
    PublishTrayIcon(); /* Ordinary update never deletes/re-adds a healthy icon. */
    assert(calls == before + 1 && commands[before] == NIM_MODIFY && deletes == 0);
    held = 0;
    TrayMessages(g_WM_TASKBARCREATED, 0); /* Real Explorer restart. */
    assert(held && g_trayRegistered && !timer && deletes == 0);
    before = calls;
    TrayMessages(g_WM_TASKBARCREATED, 0); /* Broadcast with registration retained. */
    assert(calls == before + 2 && commands[before] == NIM_ADD && commands[before+1] == NIM_MODIFY);
    held = 0;
    before = calls;
    PublishTrayIcon(); /* Shell silently lost the icon. */
    assert(calls == before + 2 && commands[before] == NIM_MODIFY && commands[before+1] == NIM_ADD);
    ready = 0; held = 0;
    TrayMessages(g_WM_TASKBARCREATED, 0); /* Broadcast can precede readiness. */
    assert(timer && !g_trayRegistered && g_trayRetryPending);
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(held && !timer && g_trayRegistered);
    ready = 0; held = 0;
    PublishTrayIcon();
    assert(timer);
    StopTrayRegistration();
    before = calls;
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    TrayMessages(g_WM_TASKBARCREATED, 0);
    PublishTrayIcon();
    StopTrayRegistration();
    assert(calls == before && !timer && !g_trayActive && !g_trayRegistered);
    g_trayActive = TRUE; ready = 0; failTimer = 1;
    PublishTrayIcon();
    assert(!timer && !g_trayRetryPending && !g_trayRegistered);
    failTimer = 0;
    TrayMessages(g_WM_TASKBARCREATED, 0);
    assert(timer && g_trayRetryPending); /* Later event can re-arm after timer failure. */
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(g_trayRegistered && !timer);
    StopTrayRegistration();

    /* --- Blinking --- */
    calls = 0; held = 0; deletes = 0; failTimer = 0; ready = 1;
    cache.hIcon = hIconSuccess;
    g_trayActive = TRUE;
    RestartTrayBlink();
    PublishTrayIcon();
    assert(!g_trayBlinkInterval && !blinkTimer && published.hIcon == hIconSuccess);
    before = calls;
    for (int i = 0; i < 10; i++) TICK();
    assert(calls == before); /* Healthy: solid, no extra shell traffic. */

    void* blinkers[] = { hIconEmpty, hIconFail };
    for (int i = 0; i < 2; i++) {
        void* icon = blinkers[i];
        Report(icon);
        assert(g_trayBlinkInterval && blinkTimer && published.hIcon == icon); /* Visible at once. */
        int armed = blinkArms;
        iconOnly = 0;
        for (int n = 0; n < 20; n++) {
            TICK();
            assert(published.hIcon == (n % 2 ? icon : hIconBlank));
            assert(cache.hIcon == icon); /* The status itself is never lost. */
        }
        assert(iconOnly == 20 && deletes == 0);
        TICK();
        assert(published.hIcon == hIconBlank);
        Report(icon); /* The next failing check must not disturb the rhythm... */
        assert(blinkArms == armed && published.hIcon == hIconBlank);
        PublishTrayIcon(); /* ...and neither must a tooltip update. */
        assert(published.hIcon == hIconBlank);
        TICK();
        assert(published.hIcon == icon);
        TICK();
        assert(published.hIcon == hIconBlank); /* Leave it hidden for the next change. */
    }

    /* Straight from one blinking status to the other: shown at once, timer restarted. */
    before = blinkArms;
    Report(hIconEmpty);
    assert(published.hIcon == hIconEmpty && blinkArms == before + 1 && blinkTimer);

    /* Recovery during the hidden half stops on a solid icon; a queued tick is ignored. */
    TICK();
    assert(published.hIcon == hIconBlank);
    Report(hIconSuccess);
    assert(!g_trayBlinkInterval && !blinkTimer && published.hIcon == hIconSuccess);
    before = calls;
    TICK();
    assert(calls == before && published.hIcon == hIconSuccess);

    /* While Explorer is away only the retry timer calls the shell, and the icon is
     * visible when it comes back. */
    Report(hIconFail);
    ready = 0; held = 0;
    TrayMessages(g_WM_TASKBARCREATED, 0);
    assert(timer && !g_trayRegistered);
    before = calls;
    for (int i = 0; i < 7; i++) TICK();
    assert(calls == before && !g_trayBlinkHidden);
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(held && g_trayRegistered && published.hIcon == hIconFail);
    TICK();
    assert(published.hIcon == hIconBlank);

    /* A blink tick that finds the icon gone restores the full registration. */
    held = 0;
    memset(&published, 0, sizeof(published));
    TICK();
    assert(held && g_trayRegistered && !timer && published.hIcon == hIconFail);
    assert(published.uFlags == (NIF_ICON | NIF_MESSAGE | NIF_TIP));
    assert(wcscmp(published.szTip, L"latest status") == 0);
    ready = 0; held = 0;
    TICK();
    assert(timer && !g_trayRegistered && g_trayRetryPending); /* Or hands over to the retry. */
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(held && !timer);

    /* Without a timer the icon simply stays solid. */
    Report(hIconSuccess);
    failTimer = 1;
    Report(hIconEmpty);
    assert(!g_trayBlinkInterval && published.hIcon == hIconEmpty);
    TICK();
    assert(published.hIcon == hIconEmpty);
    failTimer = 0;

    /* --- Faster after the configured time --- */
    assert(configFastBlinkMinutes == 10); /* The default. */
    assert(TRAY_BLINK_FAST_MS < TRAY_BLINK_MS);
    int choices[] = { 10, 1, 60 };
    for (int c = 0; c < 3; c++) {
        CHOOSE(choices[c]);
        for (int i = 0; i < 2; i++) {
            Report(hIconSuccess);
            assert(!blinkTimer && g_trayBlinkMs == 0);
            Report(blinkers[i]);
            BLINK(TICKS(choices[c]), TRAY_BLINK_MS); /* Not a moment early... */
            assert(blinkDelay == TRAY_BLINK_FAST_MS && cache.hIcon == blinkers[i]);
            int armed = blinkArms;
            BLINK(5000, TRAY_BLINK_FAST_MS); /* ...then for good, armed once. */
            Report(blinkers[i]); /* Further failing checks change nothing. */
            BLINK(5, TRAY_BLINK_FAST_MS);
            assert(blinkArms == armed);
            Report(blinkers[!i]); /* Still the same outage: fast from the first blink. */
            assert(published.hIcon == blinkers[!i]);
            BLINK(5, TRAY_BLINK_FAST_MS);
        }
    }
    CHOOSE(10);

    /* Changing icons part-way neither restarts nor loses the time so far. */
    Report(hIconSuccess);
    Report(hIconEmpty);
    BLINK(TICKS(4), TRAY_BLINK_MS);
    Report(hIconFail);
    BLINK(TICKS(6), TRAY_BLINK_MS);
    assert(blinkDelay == TRAY_BLINK_FAST_MS);

    /* Returning to green starts over, even one blink before the limit. */
    Report(hIconSuccess);
    assert(!blinkTimer && published.hIcon == hIconSuccess);
    Report(hIconFail);
    BLINK(TICKS(10) - 1, TRAY_BLINK_MS);
    Report(hIconSuccess);
    Report(hIconFail);
    BLINK(TICKS(10) - 1, TRAY_BLINK_MS);
    assert(blinkDelay == TRAY_BLINK_MS);

    /* Nothing blinks while Explorer is away, so that time is not counted. */
    ready = 0; held = 0;
    calls = 0;
    TrayMessages(g_WM_TASKBARCREATED, 0);
    for (int i = 0; i < 2000; i++) TICK();
    assert(blinkDelay == TRAY_BLINK_MS);
    ready = 1;
    TrayMessages(WM_TIMER, ID_TIMER_TRAY_RETRY);
    assert(held && blinkDelay == TRAY_BLINK_MS);

    /* If the faster timer cannot be set, the running one continues until it can. */
    failTimer = 1;
    BLINK(1, TRAY_BLINK_MS);
    BLINK(20, TRAY_BLINK_MS);
    failTimer = 0;
    BLINK(1, TRAY_BLINK_MS);
    BLINK(5, TRAY_BLINK_FAST_MS);

    /* Never means never, and the time still counts if that is changed later. */
    Report(hIconSuccess);
    CHOOSE(0);
    Report(hIconEmpty);
    BLINK(TICKS(90), TRAY_BLINK_MS);
    Report(hIconFail);
    BLINK(5, TRAY_BLINK_MS);
    BOOL hidden = g_trayBlinkHidden;
    CHOOSE(60); /* Saved during an outage: applies at once, mid-rhythm. */
    assert(blinkDelay == TRAY_BLINK_FAST_MS && g_trayBlinkHidden == hidden);
    BLINK(5, TRAY_BLINK_FAST_MS);
    before = blinkArms;
    CHOOSE(10); /* Already past this one too: nothing to do. */
    assert(blinkArms == before);
    CHOOSE(0);
    assert(blinkDelay == TRAY_BLINK_MS);
    BLINK(5, TRAY_BLINK_MS);

    /* Choosing a longer time goes back to normal until it is reached. */
    Report(hIconSuccess);
    CHOOSE(1);
    Report(hIconFail);
    BLINK(TICKS(1), TRAY_BLINK_MS);
    assert(blinkDelay == TRAY_BLINK_FAST_MS);
    CHOOSE(2);
    BLINK(TICKS(1), TRAY_BLINK_MS);
    assert(blinkDelay == TRAY_BLINK_FAST_MS);

    /* Saving settings never starts a blink on a healthy icon. */
    Report(hIconSuccess);
    before = blinkArms;
    CHOOSE(10); CHOOSE(0); CHOOSE(10);
    assert(blinkArms == before && !blinkTimer && !g_trayBlinkInterval);

    /* Shutdown: no tick may use the icon handles once they are about to be destroyed. */
    Report(hIconFail);
    TICK();
    StopTrayRegistration();
    before = calls;
    TICK();
    BlinkTrayIcon();
    CHOOSE(1);
    assert(calls == before && !blinkTimer && !g_trayBlinkInterval && !g_trayBlinkHidden && !held);
    assert(g_trayBlinkMs == 0);
    puts("tray lifecycle passed");
    return 0;
}
'''
        defines = "\n".join(re.findall(r"^#define (?:ID_TIMER_TRAY_RETRY|TRAY_RETRY_MS|ID_TIMER_TRAY_BLINK|TRAY_BLINK_\w+|FAST_BLINK_\w+) .+$", SOURCE, re.M))
        harness = defines + "\n" + stubs + "\n".join(map(function, (
            "PublishTrayIcon", "TrayBlinkIntervalDue", "StartTrayBlinkTimer",
            "UpdateTrayBlinkRate", "RestartTrayBlink", "BlinkTrayIcon",
            "StopTrayRegistration")))
        harness += "\nstatic LRESULT TrayMessages(UINT msg, WPARAM wParam) {" + messages + "return -1;\n}\n"
        apply = re.sub(r"memcpy\(nid.szTip[^;]*;", "", function("ApplyTrayUpdate"))
        harness += "typedef struct { UINT flags; void* icon; } TrayUpdate;\n" + apply
        harness += """
static void Report(void* icon) {
    TrayUpdate update = { NIF_ICON, icon };
    ApplyTrayUpdate(&update);
}
"""
        harness += assertions
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "tray.c"
            exe = Path(directory) / "tray"
            source.write_text(harness)
            subprocess.run(["gcc", "-std=c99", "-Wall", "-Wextra", "-Werror", str(source), "-o", str(exe)], check=True)
            subprocess.run([str(exe)], check=True)

    def test_worker_updates_are_owned_snapshots_on_the_window_thread(self):
        harness = r'''
#include <assert.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
typedef unsigned UINT;
typedef void* HICON;
typedef intptr_t LPARAM;
#define NIF_ICON 2
#define NIF_TIP 4
#define WM_APP_TRAY_UPDATE 0x8005
typedef struct { UINT flags; HICON icon; char tip[128]; } TrayUpdate;
static struct { HICON hIcon; char szTip[128]; } nid;
static int g_trayActive = 1, publishes, thread = 2, postSucceeds = 1, outstanding;
static void* g_hwnd = (void*)1;
static TrayUpdate* queued;
static int restarts;
static void PublishTrayIcon(void) { assert(thread == 1); publishes++; }
static void RestartTrayBlink(void) { assert(thread == 1); restarts++; }
static UINT GetCurrentThreadId(void) { return thread; }
static UINT GetWindowThreadProcessId(void* hwnd, void* pid) { assert(hwnd == g_hwnd && !pid); return 1; }
static int PostMessage(void* hwnd, UINT msg, uintptr_t wp, LPARAM lp) {
    assert(hwnd == g_hwnd && msg == WM_APP_TRAY_UPDATE && !wp && !queued);
    if (!postSucceeds) return 0;
    queued = (TrayUpdate*)lp;
    return 1;
}
static void* Allocate(size_t size) { outstanding++; return malloc(size); }
static void Release(void* ptr) { outstanding--; free(ptr); }
#define malloc Allocate
#define free Release
'''
        harness += function("ApplyTrayUpdate") + "\n" + function("QueueTrayUpdate")
        harness += r'''
int main(void) {
    char text[128] = "worker status";
    QueueTrayUpdate(NIF_TIP, NULL, text);
    strcpy(text, "changed after posting");
    assert(queued && publishes == 0 && !nid.szTip[0]);
    thread = 1;
    ApplyTrayUpdate(queued); free(queued); queued = NULL;
    assert(publishes == 1 && strcmp(nid.szTip, "worker status") == 0 && outstanding == 0);
    QueueTrayUpdate(NIF_ICON, (HICON)23, NULL);
    assert(!queued && publishes == 2 && nid.hIcon == (HICON)23 && restarts == 1);
    QueueTrayUpdate(NIF_ICON, (HICON)23, NULL);
    assert(publishes == 3 && restarts == 1); /* Same icon again: no new rhythm. */
    publishes = 2;
    assert(strcmp(nid.szTip, "worker status") == 0);
    thread = 2;
    QueueTrayUpdate(NIF_ICON | NIF_TIP, (HICON)24, "newer");
    assert(publishes == 2 && nid.hIcon == (HICON)23);
    g_trayActive = 0; thread = 1;
    ApplyTrayUpdate(queued); free(queued); queued = NULL;
    assert(publishes == 2 && nid.hIcon == (HICON)23 && outstanding == 0);
    thread = 2; postSucceeds = 0;
    QueueTrayUpdate(NIF_TIP, NULL, "closed window");
    assert(!queued && outstanding == 0 && publishes == 2);
    return 0;
}
'''
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "updates.c"
            exe = Path(directory) / "updates"
            source.write_text(harness)
            subprocess.run(["gcc", "-std=c99", "-Wall", "-Wextra", "-Werror", str(source), "-o", str(exe)], check=True)
            subprocess.run([str(exe)], check=True)
        worker = re.search(r"^DWORD WINAPI RefreshThread\([^;{]*\{.*?^}", SOURCE, re.M | re.S).group()
        self.assertNotIn("nid.", worker)
        self.assertNotIn("Shell_NotifyIcon", worker)
        self.assertRegex(function(PROC), r"case WM_APP_TRAY_UPDATE:.*?ApplyTrayUpdate\(update\);\s+free\(update\);".replace(".*?", "[\\s\\S]*?"))

    def test_broadcast_and_shutdown_wiring(self):
        self.assertNotIn("HWND_MESSAGE", SOURCE)
        register = function("RegisterTaskbarMessage")
        self.assertIn('RegisterWindowMessageW(L"TaskbarCreated")', register)
        self.assertIn('"ChangeWindowMessageFilterEx"', register)
        self.assertIn('allow(hwnd, g_WM_TASKBARCREATED, 1 /* MSGFLT_ALLOW */', register)
        main = SOURCE[SOURCE.index("int WINAPI WinMain("):]
        self.assertLess(main.index("RegisterTaskbarMessage("), main.index("InitTrayIcon("))
        create = function("InitTrayIcon")
        self.assertIn("g_trayActive = TRUE;", create)
        self.assertIn("PublishTrayIcon();", create)
        self.assertLess(create.index("nid.hIcon = "), create.index("RestartTrayBlink();"))
        self.assertLess(create.index("RestartTrayBlink();"), create.index("PublishTrayIcon();"))
        proc = function(PROC)
        self.assertRegex(proc, r"case WM_DESTROY:\s+StopTrayRegistration\(\);")
        ids = re.findall(r"^#define ID_TIMER_\w+\s+(\d+)\s*$", SOURCE, re.M)
        self.assertEqual(len(ids), len(set(ids)), "Retry timer must not collide with another timer")

    def test_fast_blink_setting(self):
        dialog = (ROOT / "assets/src/ConfigView.tsx").read_text()
        select = dialog[dialog.index('id="fast-blink"'):]
        select = select[:select.index("</select>")]
        offered = dict(re.findall(r"<option value=\{(\d+)\}>([^<]+)</option>", select))
        allowed = re.findall(r"case (\d+):", function("IsFastBlinkChoice"))
        self.assertEqual(sorted(offered), sorted(allowed), "Every offered choice must be accepted")
        self.assertEqual(offered["0"], "Never")
        default = re.search(r"^#define FAST_BLINK_DEFAULT_MINUTES (\d+)$", SOURCE, re.M).group(1)
        self.assertEqual(offered[default], "After 10 minutes")
        self.assertEqual(dialog.count("config.fastBlinkMinutes ?? " + default), 2)
        # Loaded, saved, sent to the dialog, taken back, and applied straight away.
        load, save = (function(name) for name in ("LoadConfigFromRegistry", "SaveConfigToRegistry"))
        self.assertRegex(load, r"REG_VALUE_FAST_BLINK[\s\S]*?type == REG_DWORD && IsFastBlinkChoice")
        self.assertIn("REG_VALUE_FAST_BLINK, 0, REG_DWORD", save)
        self.assertIn('\\"fastBlinkMinutes\\":%d', SOURCE)
        handler = SOURCE[SOURCE.index('strcmp(action, "saveSettings")'):]
        handler = handler[:handler.index('strcmp(action, "close")')]
        self.assertIn("int fastBlinkMinutes = configFastBlinkMinutes;", handler)
        self.assertIn('json_get_int(msg, "fastBlinkMinutes", &fastBlinkMinutes);', handler)
        self.assertRegex(handler, r"if \(IsFastBlinkChoice\(fastBlinkMinutes\)\) \{\s+"
                                  r"configFastBlinkMinutes = fastBlinkMinutes;")
        self.assertLess(handler.index("configFastBlinkMinutes = fast"),
                        handler.index("SaveConfigToRegistry();"))
        self.assertLess(handler.index("SaveConfigToRegistry();"), handler.index("ApplyConfiguration();"))
        self.assertIn("UpdateTrayBlinkRate();", function("ApplyConfiguration"))

    def test_icon_artwork(self):
        icons = {name: (ROOT / "assets" / (name + ".ico")).read_bytes()
                 for name in ("empty", "fail", "blank")}
        for name, data in icons.items():
            reserved, kind, count = struct.unpack("<HHH", data[:6])
            width, height, _, _, _, bpp, size, offset = struct.unpack("<BBBBHHII", data[6:22])
            self.assertEqual((reserved, kind, count, width, height, bpp, size + offset),
                             (0, 1, 1, 16, 16, 32, len(data)), name)
        # Hidden means transparent however it is drawn: by alpha, or by mask when a
        # reader treats the near-empty alpha channel as absent.
        blank = icons["blank"][22:]
        self.assertEqual(struct.unpack("<IiiHHI", blank[:20]), (40, 16, 32, 1, 32, 0))
        pixels, mask = blank[40:40 + 16 * 16 * 4], blank[40 + 16 * 16 * 4:]
        self.assertEqual([p for p in pixels if p], [1])
        self.assertEqual(pixels.index(1) % 4, 3)
        self.assertEqual(mask, b"\xff\xff\x00\x00" * 16)


if __name__ == "__main__":
    unittest.main()
