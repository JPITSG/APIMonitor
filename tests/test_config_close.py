"""Exercise the native close gate; the browser suite checks edit decisions."""
from pathlib import Path
import re
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'main.c').read_text()


def function(name):
    match = re.search(r'^static [^;{]*\b' + name + r'\([^;{]*\) \{', SOURCE, re.M)
    assert match, name
    return SOURCE[match.start():SOURCE.index('\n}', match.end()) + 2]


HARNESS = r'''
#include <assert.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
#define TRUE 1
#define FALSE 0
static char g_pendingView[16] = "config";
static BOOL g_configViewReady, g_configCloseApproved, g_updateInstallReady;
static void *g_webviewView;
static int requests;
static void webview_execute_script(const wchar_t *script) {
    assert(wcscmp(script, L"window.onCloseRequested()") == 0);
    requests++;
}
/* PRODUCTION */
int main(void) {
    assert(!RequestConfigClose());  // Loading or failed WebView can still close.
    g_webviewView = (void *)1;
    assert(!RequestConfigClose());
    g_configViewReady = TRUE;
    assert(RequestConfigClose());
    assert(RequestConfigClose());  // Repeated X does not bypass the prompt.
    assert(requests == 2);
    g_configCloseApproved = TRUE;
    assert(!RequestConfigClose());  // Save, discard, or unchanged configuration.
    g_configCloseApproved = FALSE;
    g_updateInstallReady = TRUE;
    assert(!RequestConfigClose());  // Updater must be able to finish its handoff.
    g_updateInstallReady = FALSE;
    strcpy(g_pendingView, "history");
    assert(!RequestConfigClose());
    assert(requests == 2);
    return 0;
}
'''

with tempfile.TemporaryDirectory(prefix='apimonitor-close-') as directory:
    path = Path(directory)
    (path / 'test.c').write_text(HARNESS.replace('/* PRODUCTION */', function('RequestConfigClose')))
    subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                    str(path / 'test.c'), '-o', str(path / 'test')], check=True)
    subprocess.run([str(path / 'test')], check=True, timeout=10)

proc = function('WebViewWndProc')
close = proc.split('case WM_CLOSE:', 1)[1].split('case WM_DESTROY:', 1)[0]
assert close.index('if (RequestConfigClose()) return 0;') < close.index('->Close(')
assert close.index('if (RequestConfigClose()) return 0;') < close.index('DestroyWindow(')
assert 'g_configCloseApproved = FALSE;' in proc.split('case WM_DESTROY:', 1)[1]
assert 'g_configCloseApproved = FALSE;' in function('ShowWebViewDialog')
handler = function('MsgReceived_Invoke')
save = handler.split('strcmp(action, "saveSettings")', 1)[1].split('strcmp(action, "close")', 1)[0]
assert save.index('SaveConfigToRegistry();') < save.index('g_configCloseApproved = TRUE;')
assert save.index('g_configCloseApproved = TRUE;') < save.index('PostMessage(g_webviewHwnd, WM_CLOSE')
close = handler.split('strcmp(action, "close")', 1)[1].split('strcmp(action, "clearHistory")', 1)[0]
assert close.index('g_configCloseApproved = TRUE;') < close.index('PostMessage(g_webviewHwnd, WM_CLOSE')
print('Native configuration close gate and save/discard wiring checks passed')
