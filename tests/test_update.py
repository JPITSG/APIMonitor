"""Run native update logic with mocked OS calls; never launch the application."""
from pathlib import Path
import re
import subprocess
import tempfile

root = Path(__file__).resolve().parents[1]
source = (root / 'main.c').read_text()


def function(name):
    match = re.search(r'^static [^;{]*\b' + name + r'\([^;{]*\) \{', source, re.M)
    assert match, name
    end = source.index('\n}', match.end()) + 2
    return source[match.start():end]


harness = (root / 'tests/update_harness.c').read_text()
harness = harness.replace('/* PRODUCTION FUNCTIONS */', '\n\n'.join(
    function(name) for name in (
        'ParseUpdateProcessId', 'HandleUpdateCommandLine', 'UpdateProgressTimer')))
with tempfile.TemporaryDirectory(prefix='apimonitor-tests-') as directory:
    path = Path(directory)
    (path / 'test.c').write_text(harness)
    subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                    '-Wno-unused-parameter', str(path / 'test.c'),
                    '-o', str(path / 'test')], check=True)
    subprocess.run([str(path / 'test')], check=True)

# Verify the two launch hops and rollback remain wired to the tested parser.
assert 'installUpdate(reopenSettingsAfterUpdate)' in (root / 'assets/src/ConfigView.tsx').read_text()
assert 'LaunchStagedUpdate(task->stagedPath, task->targetPath, reopenSettings)' in source
assert 'reopenSettings ? L" --reopen-settings" : L""' in function('LaunchStagedUpdate')
assert 'successfulUpdate && reopenSettings ? L" --reopen-settings" : L""' in function('LaunchUpdateTarget')
assert 'launchToken, TRUE,\n                            reopenSettings)' in function('RunUpdateApplyHelper')
assert 'launchToken, FALSE, FALSE)' in function('RestartAfterUpdateFailure')
assert 'if (updateCompleted && reopenSettings)' in source
assert 'InterlockedAdd64(&g_updateReceivedBytes, bytesRead)' in function('DownloadUpdateFile')
print('Updater launch/startup wiring checks passed')
