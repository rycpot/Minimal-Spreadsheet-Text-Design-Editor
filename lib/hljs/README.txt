Syntax colouring for the plain-text / code editor.

  hljs.min.js         highlight.js 11.12.0 "common" build (~35 popular languages built in),
                      loaded the first time a text file opens
  lang/<name>.min.js  one small grammar per additional language, loaded only when a file
                      needs it (built unmodified from the 11.12.0 source)
  LICENSE.txt         highlight.js's BSD-3-Clause licence

A grammar the core already contains is never fetched a second time. The
mapping from file extension / name to grammar lives in TEXT_LANGUAGES in
editor.js. To add a language: drop its grammar in lang/ as <name>.min.js
(a file that calls hljs.registerLanguage("<name>", ...), like the ones
highlight.js publishes) and add a row to TEXT_LANGUAGES. If anything in
here is missing, files still open — they just show without colours.
