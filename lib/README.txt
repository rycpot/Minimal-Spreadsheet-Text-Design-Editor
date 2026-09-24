This folder needs ONE file that isn't included here because it's a
large third-party library file (SheetJS), not something I can safely
embed as text.

Step 1: Download this exact file:
  https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js

Step 2: Save it into this "lib" folder, next to this README, with the
name exactly as:
  xlsx.full.min.js

That's it — the extension will pick it up automatically.

(SheetJS is the standard open-source library for reading/writing
Excel files in the browser; using it directly, unmodified, from its
official CDN keeps the extension small and lets it preserve formulas
correctly.)

------------------------------------------------------------------
Also in this folder
------------------------------------------------------------------
konva.min.js   Konva 10.6.0 (MIT), the canvas library behind the template
               design editor. Bundled as supplied; loaded only when you
               start a design.
