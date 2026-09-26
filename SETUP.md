
## Load the extension in Chrome

1. Go to `chrome://extensions`
2. Turn on **Developer mode** (top-right toggle)
3. Click **Load unpacked**
4. Select this folder (`sheet-editor-extension`)
5. Click the extension's icon in your toolbar — it opens the editor
   as a full browser tab

## Using it

- **Open File** — or drag a file anywhere onto the dashed area of the
  home screen — to open a spreadsheet (`.xlsx`, `.csv`, `.tsv`, `.ods`,
  `.numbers`, …) or any plain-text file (see below). The home screen also
  has **Create blank text**, **Create blank spreadsheet** and **Import
  from Google Sheet**
- **Password-protected `.xlsx` files** are supported: if the file is
  encrypted, you'll be prompted for its password. Decryption happens
  entirely locally in your browser (via the Web Crypto API) — the
  file and password never leave your machine. Both of Excel's
  encryption schemes are supported (the modern "Agile" scheme used
  by Excel 2013+, and the older "Standard" AES scheme from Excel
  2007–2010).
- Click any cell to edit it directly. Type `=SUM(A1:A3)` style text
  to enter a formula (must start with `=`)
- Click (or drag across) row numbers / column letters to select whole
  rows/columns, then click the toolbar's **Delete** button to remove
  them. With plain cells selected, **Delete** just clears their contents
- At the far left of the sheet-tab bar:
  - **✎** shows the rename button on every tab and also enables
    drag-to-reorder for tabs (off = no accidental moves)
  - **✕** shows a delete checkbox on every tab. Tick one or more, then
    click the toolbar's **Delete**; nothing is deleted until you do (if
    rows/columns are also selected, you're asked which to delete)
  - **+** adds a new sheet right after the one you're on
  - **≡** lists all sheets; click a name to jump to it
- **+ Row** / **+ Column** insert after the selected cell
- **Save / Export** downloads the edited file
  (`yourfile_edited.xlsx` or `.csv`)

## What's preserved, and what isn't

- **Formulas** are read and written back as real formulas (not just
  their last calculated value).
- **Sheet structure** (multiple sheets, merged cells, column widths)
  is kept intact for any row/column you don't delete.
- **Fonts, fill colors, borders, and other visual styling**: SheetJS's
  free/community edition (used here) has only limited support for
  *writing* style information back into a file. Values and formulas
  will be correct, but heavy visual formatting may not survive a
  round trip perfectly. If pixel-perfect styling matters, treat this
  as a data/formula editor rather than a full Excel replacement.

## Text & code files

Anything that isn't a spreadsheet opens in a plain-text editor instead
of the grid. It decides from the file's *content*, not its extension, so
`notes.txt`, `Dockerfile`, `.env`, `server.log`, a script with no
extension, or a file with an extension nobody's heard of all open. Files
that are genuinely binary (images, PDFs, archives, Office documents…)
are turned away with a message. Files up to 20 MB are accepted; past
about 1 MB the editor drops to a bare, uncoloured textarea so typing
stays responsive.

- Syntax colours are picked from the file name (or, for extensionless
  files, a `#!` line / `<?xml` / JSON content) and can be changed from
  the language menu in the bottom bar. The colouring engine
  (highlight.js 11, in `lib/hljs/`) is fetched on demand the first time
  a text file opens, plus one small file for any language it doesn't
  already contain; if a language has no grammar the file simply shows as
  plain text.
- **Tab / Shift+Tab** indent / outdent the selected lines (Esc, then
  Tab, moves focus out of the editor). **Ctrl/Cmd+S** saves.
- **Clear** (in the toolbar) empties the file; **Ctrl/Cmd+Z** brings it
  back. **Jump to row** (bottom-left, or **Ctrl/Cmd+G**, or click the
  “Ln, Col” text) moves the caret to a row number and scrolls it into
  view, even in wrapped or very large files.
- The bottom bar shows the caret position, line/character counts, the
  encoding, and toggles for **line endings** (LF/CRLF) and **Wrap**.
- Encoding (UTF-8, UTF-8 with BOM, UTF-16, or Windows-1252 for old
  single-byte files) and line endings are kept exactly as they were, so
  an unedited file saves back byte-for-byte identical. A file with mixed
  line endings is saved with the dominant one.
- **Save / Export** downloads `yourfile_edited.ext`, same as for
  spreadsheets. The ✎ button next to the file name renames it (and can
  change the language colouring by changing the extension).

## Template design (images, text, shapes)

The home screen's **Design** section starts a small image editor. Use
**Create a blank template** for an empty 1280 × 800 canvas, or open /
drop a **PNG, JPG or WebP** file on the home screen to start a design
whose canvas is exactly that picture's size. (`.svg` files still open in
the text editor on the home screen; inside a design they drop onto the
canvas like any image.)

- **Layers** (left): the top of the list is the top of the stack. Drag a
  row to reorder; the eye hides, the padlock locks (a locked layer can't
  be moved on the canvas) and the bin deletes. Double-click a name to
  rename it. Clicking an element on the canvas highlights its row and
  vice-versa. The **Canvas** is the last row: select it and drag its
  right / bottom handles (or use W, H and **Apply**) to resize it.
- **Adding things**: **Image** / **Text** / **Shapes** in the tool bar, or
  simply drag image and SVG files onto the canvas (or paste an image).
  Images keep their proportions when resized from the corners. Double-click
  text to edit it in place; **Enter** adds a line, **Ctrl/Cmd+Enter** or
  clicking away finishes, **Esc** cancels.
- **Tool bar**: undo / redo, **Crop** (select an image; drag the handles;
  **Enter** applies, **Esc** cancels; cropping is non-destructive, so you
  can re-crop to bring cut-off parts back), centre horizontally /
  vertically, canvas colour, and — for selected text — font, size, bold,
  italic and colour. To use your own font, drop a `.ttf`, `.otf`, `.woff`
  or `.woff2` file onto the canvas (or use **Add a font file…** in the
  font menu). Fonts used in a saved Template are embedded in the file, so a
  reopened design looks right even in a browser that doesn't have the font.
- **Keyboard**: arrows nudge (Shift = 10 px), **Delete** removes,
  **Ctrl/Cmd+D** duplicates, **Ctrl/Cmd+Z / Shift+Z** undo / redo,
  **Ctrl/Cmd+S** exports. Ctrl/Cmd + wheel zooms, the wheel or dragging the
  grey background pans, **Fit** re-fits the canvas.
- **Image export**: pick **PNG** or **JPG**; both default to full quality
  (PNG is lossless; JPG is 100 %). To hit a file size, enter a number in
  **Target size** and choose **KB** or **MB** — this only works for JPG (PNG
  can't trade quality for size), so typing a target switches the format to
  JPG automatically. The result, and any quality/size reduction, are shown in
  the message after saving. Anything hanging off the canvas is cropped, and
  hidden layers are not exported.
- **Save Export / PNG | JPG | Template**: choose what Save / Export produces.
  **Template** saves the whole design — every layer, picture and custom font —
  as a `.design.json` file you can reopen with **Open File** or by dropping it
  back in (also works if it's renamed, as long as it keeps the `.json`
  extension). PNG/JPG export doesn't count as saving the design, so the
  unsaved-changes prompt still applies until you save a Template.
- Dragging an element snaps to the canvas edges, centre lines and other
  elements' edges once it's close (about 7 screen px); a magenta guide line
  shows the snap. It's a sticky point, not a wall — keep dragging and it lets
  go, including past the canvas edge. Hold **Alt / Option** to drag without
  snapping.
- **Clear** (next to Undo/Redo) removes every layer but keeps the canvas
  size and colour; **Ctrl/Cmd+Z** brings it back.

## Performance notes

- **Nothing heavy loads at startup.** The spreadsheet libraries (SheetJS
  and HyperFormula, ~3.1 MB) are only fetched the first time you open a
  spreadsheet or create a blank one, and stay loaded for the rest of the
  session. The design editor (Konva, `lib/konva.min.js`, plus `design.js`)
  is likewise fetched only the first time you start a design. The home
  screen and the text editor never load either.
- **Loading bar.** While a file opens, a card with a progress bar shows
  the current step (reading the file, loading the spreadsheet engine,
  parsing, building the sheet). Small files that open instantly don't
  flash it.
- **Flat styling.** The interface uses no drop shadows, colour ramps or
  blur effects, which keeps repainting cheap (in particular when you
  switch back to the tab). Panels are separated by plain borders. The
  coloured boxes shown around cell ranges while you type a formula are
  drawn as one overlay box per range.
