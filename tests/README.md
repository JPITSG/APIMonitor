# Update regression checks

Run `python3 tests/test_update.py` with Python 3 and a host C compiler. It compiles
production timer and command-line parsing functions against deterministic OS
stubs, tests rounding, throttling, stalls, cancellation, and successful/failed
handoffs, then checks launch/startup wiring. It never runs the application.

After `make`, run `node tests/ui_update.cjs` with Puppeteer available to Node.
Set `CHROMIUM_PATH` to a Chromium executable if needed. This loads the built UI
with a mocked WebView bridge and checks the exact speed label, red styling,
cancellation, confirmation checkbox defaults/resets, and install payloads.
The browser screenshot is written to `/tmp/apimonitor-update-modal.png`.

These checks do not verify Windows WebView2, UAC, executable replacement, or
process restart behavior on Windows.
