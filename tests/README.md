# Update regression checks

Run `python3 tests/test_update.py` with Python 3 and a host C compiler. It compiles
production timer and command-line parsing functions against deterministic OS
stubs, tests percentage rounding, the pre-transfer state, stalls, cancellation,
and successful/failed handoffs, then checks launch/startup wiring. It never
runs the application.

After `make`, run `node tests/ui_update.cjs` with Puppeteer available to Node.
Set `CHROMIUM_PATH` to a Chromium executable if needed. This loads the built UI
with a mocked WebView bridge and checks the exact percentage label, red styling,
cancellation, confirmation checkbox defaults/resets, and install payloads.
The browser screenshot is written to `/tmp/apimonitor-update-modal.png`.

Run `python3 tests/test_fixed_frame.py` for the fixed dialog size. It compiles
the production frame handling against stubbed window calls: the Close-only
title bar and system menu, edge and corner drags, the Size and Maximize
commands, and the track size pinned to the size the app chose (which also stops
Snap). It then checks that the dialog is created with that frame and that every
place the app sizes it pins the size first.

These checks do not verify Windows WebView2, UAC, executable replacement, or
process restart behavior on Windows.
