![enter image description here](https://files.catbox.moe/wu3657.png)
![enter image description here](https://files.catbox.moe/rgeipq.png)
![enter image description here](https://files.catbox.moe/nah6o8.png)
![enter image description here](https://files.catbox.moe/evk5jl.png)
# Minimal Spreadsheet, Text & Design Editor: Runs Local, Works Fully Offline

A simple, offline Chrome extension for editing spreadsheets, text/code files, and quick image designs — right in your browser, with nothing uploaded anywhere.

> ## ⚠️ Please read: this is a very basic editor
>
> 🚫 **Not a replacement.** This is not an equivalent or a replacement for Microsoft Excel, Apple Numbers, or OpenDocument spreadsheets (LibreOffice Calc and similar).
>
> 🧮 **Made for casual number crunching.** Formula support is limited: not every spreadsheet function is available, and some formulas may not calculate.
>
> 🎨 **Formatting is not preserved.** Styles, formatting, charts and graphs, and other sophisticated spreadsheet features (pivot tables, conditional formatting, images, macros, and so on) are not preserved when you save.
>
> 🛡️ **Your original file is always safe.** The editor never overwrites the file you opened. Saving downloads a new copy (for example `report_edited.xlsx`), so your original stays untouched.

## What it does

- **Spreadsheets** — open `.xlsx`, `.csv`, `.tsv`, `.ods`, or Apple Numbers files, edit cells, add/delete rows and columns, use formulas (`=SUM(A1:A3)`), and save back. Password-protected `.xlsx` files are supported (decrypted locally, in your browser).
- **Text & code files** — open any plain-text or code file in a lightweight editor with syntax highlighting.
- **Design templates** — a simple canvas for combining images, text, and shapes, exportable as PNG/JPG.

## Privacy

Everything happens locally in your browser. No files, formulas, or images are ever sent to a server — there's nothing to upload to in the first place. The extension only needs access to `docs.google.com` (to let you import a Google Sheet you have open).

## Install

![enter image description here](https://fonts.gstatic.com/s/i/productlogos/chrome_store/v7/192px.svg)

INSTALL FROM CHROME WEBSTORE

**From source (for developers)**
1. Download or clone this repo.
2. Go to `chrome://extensions` in Chrome.
3. Turn on **Developer mode** (top-right toggle).
4. Click **Load unpacked** and select this folder.
5. Click the extension's icon in your toolbar — it opens as a full browser tab.

## Basic usage

- **Open File**, or drag a file onto the home screen, to start editing. You can also **Create blank text**, **Create blank spreadsheet**, or **Import from Google Sheet**.
- Click a cell to edit it. Formulas must start with `=`.
- Select rows/columns by clicking their numbers/letters, then use the toolbar **Delete** button.
- **Save / Export** downloads the edited file.
- Switch to the **design** tab to lay out images, text, and shapes, then export as an image.

## License

This project's own code is licensed under the **GNU General Public License v3.0 (GPLv3)** — see [`LICENSE`](./LICENSE) for the full text. In short: you're free to use, study, share, and modify it, but any copies or modified versions you distribute must also be GPLv3 (source code included).

### Why GPLv3?

This extension bundles [HyperFormula](https://hyperformula.handsontable.com/) (the formula engine behind spreadsheet calculations) under its **free, open-source license**, which requires the whole project to be GPLv3-compatible. If you fork this project for commercial/closed-source use, you'd need to either buy a [proprietary HyperFormula license](https://hyperformula.handsontable.com/guide/licensing.html) or swap it out for a different formula engine.

### Open-source libraries used

| Library | Version | License | Used for |
|---|---|---|---|
| [HyperFormula](https://hyperformula.handsontable.com/) | 3.4.0 | GPLv3 (free tier) | Spreadsheet formula calculations |
| [SheetJS (xlsx)](https://git.sheetjs.com/sheetjs/sheetjs) | 0.20.3 | Apache-2.0 | Reading/writing `.xlsx` / `.csv` / `.ods` files |
| [Konva](https://konvajs.org/) | 10.6.0 | MIT | Canvas engine behind the design/template editor |
| [highlight.js](https://highlightjs.org/) | 11.12.0 | BSD-3-Clause | Syntax highlighting in the text/code editor |

Each library's own license text is kept alongside it in this repo (see `lib/hljs/LICENSE.txt`, and `lib/README.txt` for the rest) and is unmodified from what the project publishes.

The password-decryption code (`office-crypto.js`) is original code written for this project, using only the browser's built-in Web Crypto API — no third-party library involved.
