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

hyperformula.full.js
               HyperFormula 3.4.0 (GPLv3), the formula engine. This is the
               official, unminified dist/hyperformula.full.js from the npm
               package hyperformula@3.4.0, with these local changes (each is
               marked in the file with an "[Extension patch]" comment):

               1. In the bundled core-js `Object.create` polyfill (webpack
               module 87), the old-IE fallback that built a null-prototype
               object via an ActiveX document or a hidden iframe with a
               "javascript:" src was removed, and replaced by a plain
               `Object.create(null)`. That path could never run in Chrome
               (which has a native Object.create), and it was flagged by
               Chrome Web Store review.

               2. No code is built from strings. The two global-object
               lookups that fell back to a Function-constructor call now use
               `globalThis`; an unreachable eval call placed after a
               `return` (a V8 optimisation hint in Chevrotain) was deleted;
               and Chevrotain's optional parser code generator
               (generateParserFactory, never used by HyperFormula) now throws
               instead of compiling code. The extension's CSP would block
               all of these anyway.

               To update HyperFormula, take the new
               dist/hyperformula.full.js and re-apply the same edits.
