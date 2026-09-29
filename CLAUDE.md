# Notes for Claude

Chrome extension (Manifest V3): a local spreadsheet / text / design editor, published on the Chrome Web Store. No build step — the repo is the extension (`manifest.json`, `editor.html`, `editor.js`, `design.js`, `styles.css`, `lib/`).

## Workflow (standing instruction from the owner)

- **Always merge and release.** Once a change is done and tested, open the PR, squash-merge it yourself, and confirm the release published. Don't ask first.
- Every change that ships to users bumps `"version"` in `manifest.json` (patch bump, e.g. 1.0.5 → 1.0.6). Merging a version bump to `main` makes `.github/workflows/release.yml` publish `v<version>` with `minimal-editor-v<version>.zip` attached; check it with the latest release afterwards and give the owner the zip link. Repo-only changes (docs, CI, this file) don't need a bump.
- The owner uploads the zip to the Chrome Web Store themselves.
- Test in Chromium with the unpacked extension loaded (Playwright, `--headless=new` plus `--load-extension`) before merging; check the console for errors.

## Things to keep in mind

- Chrome Web Store review: no obfuscated code, no `eval`/`new Function`/string-built code, no remote code. `lib/hyperformula.full.js` is the unminified HyperFormula build with documented `[Extension patch]` edits (see `lib/README.txt`) — re-apply them when upgrading it.
- Store listing: avoid keyword-style lists (the short description was once rejected for "XLSX/CSV/TSV/…").
- `scripts/package.sh` builds the store zip; repo-only files (docs, CI, scripts, this file) are excluded there.
