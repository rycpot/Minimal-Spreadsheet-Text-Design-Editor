'use strict';
/* ------------------------------------------------------------------ *
 * Template design editor (loaded on demand, together with lib/konva.min.js)
 * ------------------------------------------------------------------ *
 * A small Canva/Polotno-style editor: a canvas ("artboard") holding layers —
 * images, SVGs, text and simple shapes — with a Layers panel on the left and
 * a tool bar on top. Everything is kept in a plain data model (M below); the
 * Konva nodes are just a view of it, which is what makes undo/redo, the
 * Layers list and the export (a second, off-screen render of the same model)
 * all agree with each other.
 *
 * It borrows a few things from editor.js (state, el, showToast, ...), which
 * is always loaded first. Public API: window.Design.
 * ------------------------------------------------------------------ */
(function () {
  const DEFAULT_W = 1280;
  const DEFAULT_H = 800;
  const MIN_DIM = 16;
  const MAX_DIM = 8192;
  const HISTORY_LIMIT = 100;
  const VIEW_PAD = 56;
  const PREVIEW_MAX = 2560; // photos bigger than this are shown from a smaller copy on screen (export uses the original)
  const ACCENT = '#3583eb';
  const SNAP_PX = 7;               // how close (on screen) a dragged element must get to a canvas edge/centre to stick to it
  const GUIDE_COLOR = '#ff2d78';
  const TEMPLATE_TYPE = 'sheet-editor-design';
  const TEMPLATE_VERSION = 1;

  const BUILTIN_FONTS = [
    'Arial', 'Helvetica', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Times New Roman', 'Georgia',
    'Palatino', 'Garamond', 'Courier New', 'Impact', 'Comic Sans MS', 'Brush Script MT',
    'sans-serif', 'serif', 'monospace',
  ];

  const SHAPES = [
    { id: 'rect', label: 'Rectangle' },
    { id: 'ellipse', label: 'Circle' },
    { id: 'triangle', label: 'Triangle' },
    { id: 'star', label: 'Star' },
    { id: 'line', label: 'Line' },
    { id: 'arrow', label: 'Arrow' },
  ];

  const ICON = {
    eye: '<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="2.8"/>',
    eyeOff: '<path d="M4 4l16 16"/><path d="M9.9 5.8A9.7 9.7 0 0 1 12 5.5c6 0 9.5 6.5 9.5 6.5a16 16 0 0 1-3 3.7M6.3 7.6A15.6 15.6 0 0 0 2.5 12S6 18.5 12 18.5c1.5 0 2.8-.4 4-.9"/><path d="M9.9 9.9a2.8 2.8 0 0 0 4 4"/>',
    lock: '<rect x="5.5" y="10.5" width="13" height="9" rx="1.6"/><path d="M8.5 10.5V8a3.5 3.5 0 0 1 7 0v2.5"/>',
    unlock: '<rect x="5.5" y="10.5" width="13" height="9" rx="1.6"/><path d="M8.5 10.5V8a3.5 3.5 0 0 1 6.6-1.6"/>',
    trash: '<path d="M5 7h14M9 7V4.5h6V7M7 7l1 13h8l1-13M10 10.5v6M14 10.5v6"/>',
  };
  const svgIcon = (name) => `<svg class="ui-icon" viewBox="0 0 24 24" aria-hidden="true">${ICON[name]}</svg>`;

  /* ---------------------------------------------------------------- state */
  const S = {
    active: false,
    stage: null, bgLayer: null, contentLayer: null, uiLayer: null,
    artboard: null, content: null, transformer: null, cropUI: null,
    nodes: new Map(),          // layer id -> Konva node
    assets: new Map(),         // asset id -> { kind, full, view, w, h, url }
    customFonts: [],           // family names dragged in by the user
    selectedId: null,          // 'canvas' | layer id | null
    clipboard: null,           // layer data copied with Ctrl/Cmd+C, pasted (offset) with Ctrl/Cmd+V
    history: [], hIndex: -1,
    boundToScratch: false,     // true when the current session was opened via the design scratchpad (see enter/leave)
    scratchHistory: null, scratchHIndex: -1, // undo/redo stashed here across leave()/enter() so a scratchpad Clear + reopen can still be undone
    scratchAssets: null,       // asset blobs (images/SVGs) stashed alongside scratchHistory so a resumed history's layers still have something to render
    view: { scale: 1, auto: true },
    editing: null,             // in-place text editor { id, ta, ... }
    cropping: null,
    resizeObserver: null,
    idSeq: 1,
    lastAnchor: null,
    dom: null,
    export: { busy: false },
    fontBlobs: new Map(),      // custom font family -> the file it came from (so a saved template can carry it)
    guides: null,              // the two snap guide lines
  };
  const M = { canvas: { w: DEFAULT_W, h: DEFAULT_H, fill: '#ffffff' }, layers: [] };

  const uid = (p) => `${p}${Date.now().toString(36)}${(S.idSeq++).toString(36)}`;
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
  const getLayer = (id) => M.layers.find((l) => l.id === id) || null;
  const $ = (id) => document.getElementById(id);

  /* ---------------------------------------------------------------- assets */
  function decodeImageUrl(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('Could not decode image'));
      img.src = url;
    });
  }

  function downscaleToCanvas(img, w, h, maxSide) {
    const k = Math.min(1, maxSide / Math.max(w, h));
    if (k >= 1) return img;
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(w * k));
    c.height = Math.max(1, Math.round(h * k));
    const ctx = c.getContext('2d');
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(img, 0, 0, c.width, c.height);
    return c;
  }

  // decode*Asset only builds the asset (so opening a file can fail without touching the
  // design on screen); create*Asset also registers it. `blob` keeps the original bytes,
  // which is what a saved template embeds.
  async function decodeRasterAsset(file, id) {
    const url = URL.createObjectURL(file);
    const img = await decodeImageUrl(url);
    const w = img.naturalWidth;
    const h = img.naturalHeight;
    return { id: id || uid('a'), kind: 'raster', full: img, view: downscaleToCanvas(img, w, h, PREVIEW_MAX), w, h, url, blob: file, mime: file.type || 'image/png' };
  }

  async function createRasterAsset(file) {
    const asset = await decodeRasterAsset(file);
    S.assets.set(asset.id, asset);
    return asset;
  }

  // Works out an SVG's natural size (width/height attributes, else its viewBox)
  // and rewrites the markup so it always has one — an <img> needs it.
  function prepareSvg(text) {
    const doc = new DOMParser().parseFromString(text, 'image/svg+xml');
    const svg = doc.documentElement;
    if (!svg || svg.nodeName.toLowerCase() !== 'svg' || doc.querySelector('parsererror')) throw new Error('Not a valid SVG');
    const num = (v) => { const n = parseFloat(v); return v && !/%$/.test(v) && isFinite(n) && n > 0 ? n : 0; };
    let w = num(svg.getAttribute('width'));
    let h = num(svg.getAttribute('height'));
    const vb = (svg.getAttribute('viewBox') || '').trim().split(/[\s,]+/).map(Number);
    if ((!w || !h) && vb.length === 4 && vb[2] > 0 && vb[3] > 0) {
      if (!w && !h) { w = vb[2]; h = vb[3]; } else if (!w) { w = h * vb[2] / vb[3]; } else { h = w * vb[3] / vb[2]; }
    }
    if (!w || !h) { w = 300; h = 150; }
    if (!svg.getAttribute('viewBox')) svg.setAttribute('viewBox', `0 0 ${w} ${h}`);
    // Keep the rasterised copy crisp when it's scaled up.
    const k = Math.max(1, 2048 / Math.max(w, h));
    svg.setAttribute('width', String(Math.round(w * k)));
    svg.setAttribute('height', String(Math.round(h * k)));
    if (!svg.getAttribute('xmlns')) svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    return { markup: new XMLSerializer().serializeToString(svg), w, h };
  }

  async function decodeSvgAsset(file, id) {
    const text = await file.text();
    const { markup, w, h } = prepareSvg(text);
    const blob = new Blob([markup], { type: 'image/svg+xml' });
    const url = URL.createObjectURL(blob);
    const img = await decodeImageUrl(url);
    return { id: id || uid('a'), kind: 'svg', full: img, view: img, w, h, url, blob, mime: 'image/svg+xml' };
  }

  async function createSvgAsset(file) {
    const asset = await decodeSvgAsset(file);
    S.assets.set(asset.id, asset);
    return asset;
  }

  /* ---------------------------------------------------------------- nodes */
  function shapeSceneFunc(kind) {
    return function (ctx, shape) {
      const w = shape.width();
      const h = shape.height();
      ctx.beginPath();
      if (kind === 'ellipse') {
        ctx.ellipse(w / 2, h / 2, Math.max(w / 2, 0.01), Math.max(h / 2, 0.01), 0, 0, Math.PI * 2);
        ctx.closePath();
        ctx.fillStrokeShape(shape);
      } else if (kind === 'triangle') {
        ctx.moveTo(w / 2, 0); ctx.lineTo(w, h); ctx.lineTo(0, h); ctx.closePath();
        ctx.fillStrokeShape(shape);
      } else if (kind === 'star') {
        const cx = w / 2; const cy = h / 2;
        for (let i = 0; i < 10; i++) {
          const r = i % 2 === 0 ? 1 : 0.4;
          const a = -Math.PI / 2 + (i * Math.PI) / 5;
          const px = cx + Math.cos(a) * (w / 2) * r;
          const py = cy + Math.sin(a) * (h / 2) * r;
          if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
        }
        ctx.closePath();
        ctx.fillStrokeShape(shape);
      } else if (kind === 'line') {
        ctx.moveTo(0, h / 2); ctx.lineTo(w, h / 2);
        ctx.strokeShape(shape);
      } else if (kind === 'arrow') {
        const head = Math.min(h, w * 0.5, 60);
        ctx.moveTo(0, h / 2); ctx.lineTo(w, h / 2);
        ctx.moveTo(w - head, h / 2 - head / 2); ctx.lineTo(w, h / 2); ctx.lineTo(w - head, h / 2 + head / 2);
        ctx.strokeShape(shape);
      }
    };
  }

  function makeNode(layer, forExport) {
    let node;
    if (layer.type === 'image' || layer.type === 'svg') {
      const a = S.assets.get(layer.asset);
      const src = forExport ? a.full : a.view;
      const sw = src.naturalWidth || src.width;
      const sh = src.naturalHeight || src.height;
      const c = layer.crop || { x: 0, y: 0, w: 1, h: 1 };
      node = new Konva.Image({
        image: src, width: layer.width, height: layer.height,
        crop: { x: c.x * sw, y: c.y * sh, width: c.w * sw, height: c.h * sh },
      });
    } else if (layer.type === 'text') {
      node = new Konva.Text({ lineHeight: 1.2, wrap: 'word', padding: 0 });
    } else {
      const strokeOnly = layer.shape === 'line' || layer.shape === 'arrow';
      if (layer.shape === 'rect') {
        node = new Konva.Rect({ cornerRadius: layer.radius || 0 });
      } else {
        node = new Konva.Shape({ sceneFunc: shapeSceneFunc(layer.shape), hitStrokeWidth: strokeOnly ? 20 : 0 });
      }
    }
    node.id(layer.id);
    applyLayerToNode(node, layer);
    if (!forExport) {
      node.perfectDrawEnabled(false);
    }
    return node;
  }

  function applyLayerToNode(node, layer) {
    node.setAttrs({
      x: layer.x, y: layer.y, rotation: layer.rotation || 0, scaleX: 1, scaleY: 1,
      visible: layer.visible !== false,
    });
    if (layer.type === 'image' || layer.type === 'svg') {
      node.setAttrs({ width: layer.width, height: layer.height });
    } else if (layer.type === 'text') {
      node.setAttrs({
        text: layer.text, fontFamily: layer.fontFamily, fontSize: layer.fontSize,
        fontStyle: layer.fontStyle || 'normal', fill: layer.fill,
        width: layer.autoWidth ? 'auto' : layer.width,
      });
    } else {
      const strokeOnly = layer.shape === 'line' || layer.shape === 'arrow';
      node.setAttrs({
        width: layer.width, height: layer.height,
        fill: strokeOnly ? undefined : layer.fill,
        stroke: strokeOnly ? layer.fill : (layer.stroke || undefined),
        strokeWidth: strokeOnly ? (layer.strokeWidth || 6) : (layer.stroke ? (layer.strokeWidth || 2) : 0),
        lineCap: 'round', lineJoin: 'round',
      });
    }
  }

  // Snapping, like Canva: while an element is dragged, its left / centre / right
  // (and top / middle / bottom) stick to the canvas edges and centre lines once
  // they come within a few pixels, and a guide line shows which one. It's a
  // sticky spot, not a wall: keep dragging and the element pulls free (and can
  // go past the border). Hold Alt / Option to drag without snapping.
  function snapDrag(node, evt) {
    const alt = !!(evt && evt.evt && evt.evt.altKey);
    let gx = null;
    let gy = null;
    if (!alt) {
      const th = SNAP_PX / S.view.scale;
      const W = M.canvas.w;
      const H = M.canvas.h;
      const box = node.getClientRect({ skipStroke: true, skipShadow: true, relativeTo: S.content });
      const pick = (edges, targets) => {
        let best = null;
        for (const e of edges) for (const t of targets) {
          const d = t - e;
          if (Math.abs(d) <= th && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, at: t };
        }
        return best;
      };
      const bx = pick([box.x, box.x + box.width / 2, box.x + box.width], [0, W / 2, W]);
      const by = pick([box.y, box.y + box.height / 2, box.y + box.height], [0, H / 2, H]);
      if (bx) { node.x(node.x() + bx.d); gx = bx.at; }
      if (by) { node.y(node.y() + by.d); gy = by.at; }
    }
    showGuides(gx, gy);
  }

  function showGuides(x, y) {
    const g = S.guides;
    if (!g) return;
    g.v.visible(x !== null);
    g.h.visible(y !== null);
    if (x !== null) g.v.points([x, 0, x, M.canvas.h]);
    if (y !== null) g.h.points([0, y, M.canvas.w, y]);
    S.uiLayer.batchDraw();
  }

  function hideGuides() { showGuides(null, null); }

  // boundBoxFunc doesn't get handed the pointer event the way dragmove does,
  // so Alt-to-bypass-snapping (see snapDrag above) needs its own tiny bit of
  // state tracked here instead of read off the event directly.
  let altHeld = false;
  window.addEventListener('keydown', (e) => { if (e.key === 'Alt') altHeld = true; });
  window.addEventListener('keyup', (e) => { if (e.key === 'Alt') altHeld = false; });
  window.addEventListener('blur', () => { altHeld = false; }); // don't get stuck "on" if Alt was released while unfocused

  // Same sticky-edge behavior as snapDrag, but for resizing: whichever
  // edge(s) the active handle is actually moving stick to the canvas
  // edges/centre lines once within a few screen pixels. A corner handle
  // moves two edges (e.g. top-left moves the top and left edges, leaving
  // bottom/right fixed); a side handle moves just one. Rotated shapes are
  // left alone — snapping a single "edge" of a rotated box isn't a
  // well-defined idea the way it is for an axis-aligned one.
  function snapResize(oldBox, newBox) {
    if (Math.abs(newBox.width) < 8 || Math.abs(newBox.height) < 8) return oldBox;

    const anchor = S.transformer.getActiveAnchor();
    if (altHeld || !anchor || anchor === 'rotater' || Math.abs(newBox.rotation) > 0.01) {
      hideGuides();
      return newBox;
    }

    const isCorner = anchor === 'top-left' || anchor === 'top-right' || anchor === 'bottom-left' || anchor === 'bottom-right';
    if (S.transformer.keepRatio() && isCorner) return snapResizeRatio(newBox, anchor);

    // newBox/oldBox are in absolute (stage) pixel space, already accounting
    // for the current zoom/pan — convert the canvas's own edges/centre into
    // that same space to compare against, and use the on-screen threshold
    // directly (no /scale here, unlike snapDrag's content-local version).
    const { xTargets, yTargets, closest, scale, originX, originY } = snapTargets();

    const box = { x: newBox.x, y: newBox.y, width: newBox.width, height: newBox.height };
    let gx = null;
    let gy = null;

    if (anchor.includes('left')) {
      const hit = closest(box.x, xTargets);
      if (hit) { const oldRight = box.x + box.width; box.x = hit.at; box.width = oldRight - hit.at; gx = hit.at; }
    } else if (anchor.includes('right')) {
      const hit = closest(box.x + box.width, xTargets);
      if (hit) { box.width = hit.at - box.x; gx = hit.at; }
    }

    if (anchor.includes('top')) {
      const hit = closest(box.y, yTargets);
      if (hit) { const oldBottom = box.y + box.height; box.y = hit.at; box.height = oldBottom - hit.at; gy = hit.at; }
    } else if (anchor.includes('bottom')) {
      const hit = closest(box.y + box.height, yTargets);
      if (hit) { box.height = hit.at - box.y; gy = hit.at; }
    }

    if (box.width < 8 || box.height < 8) { hideGuides(); return newBox; }

    showGuides(gx === null ? null : (gx - originX) / scale, gy === null ? null : (gy - originY) / scale);
    return { ...newBox, x: box.x, y: box.y, width: box.width, height: box.height };
  }

  // Shared by both snapResize paths: the canvas's own edges/centre lines,
  // converted from content-local (0..w / 0..h) into absolute stage pixels,
  // plus a closest-within-threshold helper.
  function snapTargets() {
    const scale = S.stage.scaleX() || 1;
    const originX = S.stage.x();
    const originY = S.stage.y();
    const toAbsX = (localX) => originX + localX * scale;
    const toAbsY = (localY) => originY + localY * scale;
    const th = SNAP_PX;
    const xTargets = [toAbsX(0), toAbsX(M.canvas.w / 2), toAbsX(M.canvas.w)];
    const yTargets = [toAbsY(0), toAbsY(M.canvas.h / 2), toAbsY(M.canvas.h)];
    const closest = (value, targets) => {
      let best = null;
      for (const t of targets) {
        const d = Math.abs(value - t);
        if (d <= th && (!best || d < best.d)) best = { d, at: t };
      }
      return best;
    };
    return { xTargets, yTargets, closest, scale, originX, originY };
  }

  // Ratio-locked version of snapResize, for a corner handle while keepRatio
  // is on (images/SVGs always; text too, via its corner anchors). Only one
  // dimension can "drive" a ratio-locked resize at a time — Konva's own
  // corner-drag-with-keepRatio picks a dominant axis the same way — so this
  // finds whichever single live edge (the one or two edges the active
  // corner actually moves) is closest to a snap target, snaps just that
  // one, and derives the other dimension from newBox's own aspect ratio
  // (already correct, since Konva computed newBox before boundBoxFunc runs)
  // rather than snapping width and height independently, which is what was
  // fighting Konva's own ratio-lock and produced the squash/stretch right
  // at the canvas edge.
  function snapResizeRatio(newBox, anchor) {
    const ratio = newBox.width / newBox.height;
    const { xTargets, yTargets, closest, scale, originX, originY } = snapTargets();

    const left = anchor.includes('left');
    const top = anchor.includes('top');

    // The opposite corner from the active handle doesn't move — compute it
    // from newBox before any snapping, which is still correct since it's
    // untouched so far.
    const fixedX = left ? newBox.x + newBox.width : newBox.x;
    const fixedY = top ? newBox.y + newBox.height : newBox.y;
    const liveX = left ? newBox.x : newBox.x + newBox.width;
    const liveY = top ? newBox.y : newBox.y + newBox.height;

    const hitX = closest(liveX, xTargets);
    const hitY = closest(liveY, yTargets);
    let hit = null;
    let axis = null;
    if (hitX && (!hitY || hitX.d <= hitY.d)) { hit = hitX; axis = 'x'; }
    else if (hitY) { hit = hitY; axis = 'y'; }

    if (!hit) { hideGuides(); return newBox; }

    let width;
    let height;
    if (axis === 'x') { width = Math.abs(fixedX - hit.at); height = width / ratio; }
    else { height = Math.abs(fixedY - hit.at); width = height * ratio; }
    if (width < 8 || height < 8) { hideGuides(); return newBox; }

    const x = left ? fixedX - width : fixedX;
    const y = top ? fixedY - height : fixedY;

    showGuides(
      axis === 'x' ? (hit.at - originX) / scale : null,
      axis === 'y' ? (hit.at - originY) / scale : null
    );
    return { ...newBox, x, y, width, height };
  }

  function attachNodeEvents(node, layer) {
    node.on('mousedown touchstart', () => { if (!S.cropping && !S.editing) select(layer.id); });
    node.on('dragstart', () => { if (S.selectedId !== layer.id) select(layer.id); });
    node.on('dragmove', (e) => snapDrag(node, e));
    node.on('dragend', () => {
      hideGuides();
      const l = getLayer(layer.id);
      if (!l) return;
      l.x = node.x(); l.y = node.y();
      commit();
    });
    node.on('transformstart', () => { S.lastAnchor = null; });
    node.on('transform', () => {
      const l = getLayer(layer.id);
      if (!l) return;
      S.lastAnchor = S.transformer.getActiveAnchor() || S.lastAnchor;
      if (l.type === 'text' && (S.lastAnchor === 'middle-left' || S.lastAnchor === 'middle-right')) {
        // Side handles change the wrapping width, not the type size.
        node.setAttrs({ width: Math.max(24, node.width() * node.scaleX()), scaleX: 1 });
      }
    });
    node.on('transformend', () => { hideGuides(); bakeTransform(node, layer.id); });
    node.on('mouseenter', () => { if (!S.cropping && node.draggable()) S.stage.container().style.cursor = 'move'; });
    node.on('mouseleave', () => { S.stage.container().style.cursor = ''; });
    if (layer.type === 'text') node.on('dblclick dbltap', () => startTextEdit(layer.id));
  }

  // Turns the live scale a Transformer leaves on a node into real sizes, so the
  // model (and the export) never depends on a leftover scale.
  function bakeTransform(node, id) {
    const l = getLayer(id);
    if (!l) return;
    const sx = node.scaleX();
    const sy = node.scaleY();
    if (l.type === 'text') {
      if (Math.abs(sy - 1) > 1e-6 || Math.abs(sx - 1) > 1e-6) {
        const k = Math.abs(sy - 1) > 1e-6 ? sy : sx;
        l.fontSize = Math.max(4, Math.round(l.fontSize * k * 10) / 10);
        if (!l.autoWidth) l.width = Math.max(24, node.width() * sx);
      } else if (S.lastAnchor === 'middle-left' || S.lastAnchor === 'middle-right') {
        l.autoWidth = false;
        l.width = Math.max(24, node.width());
      }
    } else {
      l.width = Math.max(2, node.width() * sx);
      l.height = Math.max(2, node.height() * sy);
    }
    l.x = node.x(); l.y = node.y(); l.rotation = node.rotation();
    applyLayerToNode(node, l);
    S.transformer.forceUpdate();
    syncToolbar();
    commit();
  }

  function addNodeFor(layer) {
    const node = makeNode(layer, false);
    attachNodeEvents(node, layer);
    S.content.add(node);
    S.nodes.set(layer.id, node);
    setInteractive(layer);
    return node;
  }

  function setInteractive(layer) {
    const node = S.nodes.get(layer.id);
    if (!node) return;
    const on = !layer.locked && layer.visible !== false;
    node.draggable(on);
    node.listening(on);
  }

  function rebuildAll() {
    S.transformer.nodes([]); // detach first: a Transformer left holding destroyed nodes keeps drawing its old box
    S.content.destroyChildren();
    S.nodes.clear();
    for (const l of M.layers) addNodeFor(l);
    applyCanvasToStage();
    updateTransformer();
    renderLayers();
    syncToolbar();
    S.contentLayer.batchDraw();
  }

  function refreshTextNodes() {
    for (const l of M.layers) {
      if (l.type !== 'text') continue;
      const n = S.nodes.get(l.id);
      if (n) { applyLayerToNode(n, l); }
    }
    S.contentLayer.batchDraw();
    S.transformer && S.transformer.forceUpdate();
  }

  function reorderNodes() {
    M.layers.forEach((l, i) => { const n = S.nodes.get(l.id); if (n) n.zIndex(i); });
  }

  /* ---------------------------------------------------------------- canvas / view */
  function applyCanvasToStage() {
    S.artboard.setAttrs({ width: M.canvas.w, height: M.canvas.h, fill: M.canvas.fill, scaleX: 1, scaleY: 1 });
    const d = S.dom;
    d.w.value = Math.round(M.canvas.w);
    d.h.value = Math.round(M.canvas.h);
    d.canvasColor.value = M.canvas.fill;
    S.bgLayer.batchDraw();
    S.contentLayer.batchDraw();
    if (S.view.auto) fitView(); else applyView();
  }

  function viewportSize() {
    const vp = S.dom.viewport;
    return { w: Math.max(50, vp.clientWidth), h: Math.max(50, vp.clientHeight) };
  }

  function applyView(centerOn) {
    const { w: vw, h: vh } = viewportSize();
    const z = S.view.scale;
    S.stage.size({ width: vw, height: vh });
    S.stage.scale({ x: z, y: z });
    if (centerOn === undefined || centerOn === 'center') {
      S.stage.position({ x: (vw - M.canvas.w * z) / 2, y: (vh - M.canvas.h * z) / 2 });
    }
    S.stage.batchDraw();
    S.dom.zoomLabel.textContent = `${Math.round(z * 100)}%`;
  }

  function fitView() {
    const { w: vw, h: vh } = viewportSize();
    const z = clamp(Math.min((vw - VIEW_PAD * 2) / M.canvas.w, (vh - VIEW_PAD * 2) / M.canvas.h), 0.02, 2);
    S.view.scale = z;
    S.view.auto = true;
    applyView('center');
    positionTextEditor();
  }

  function zoomTo(newScale, anchor) {
    const stage = S.stage;
    const old = S.view.scale;
    const z = clamp(newScale, 0.05, 8);
    const p = anchor || { x: stage.width() / 2, y: stage.height() / 2 };
    const wx = (p.x - stage.x()) / old;
    const wy = (p.y - stage.y()) / old;
    S.view.scale = z;
    S.view.auto = false;
    stage.scale({ x: z, y: z });
    stage.position({ x: p.x - wx * z, y: p.y - wy * z });
    stage.batchDraw();
    S.dom.zoomLabel.textContent = `${Math.round(z * 100)}%`;
    positionTextEditor();
  }

  // (Konva's Transformer draws its handles at a fixed on-screen size whatever the
  // stage zoom is, so they are configured once at creation and need no zoom handling.)

  /* ---------------------------------------------------------------- selection */
  function select(id) {
    if (S.editing) finishTextEdit(true);
    S.selectedId = id;
    updateTransformer();
    renderLayers();
    syncToolbar();
  }

  function updateTransformer() {
    const tr = S.transformer;
    if (!tr) return;
    const id = S.selectedId;
    if (S.cropping) { tr.nodes([]); tr.visible(false); S.uiLayer.batchDraw(); return; }
    if (id === 'canvas') {
      const locked = false;
      tr.setAttrs({ enabledAnchors: ['middle-right', 'bottom-center', 'bottom-right'], keepRatio: false, rotateEnabled: false, flipEnabled: false, shiftBehavior: 'none' });
      tr.nodes(locked ? [] : [S.artboard]);
    } else {
      const l = getLayer(id);
      const node = l && S.nodes.get(id);
      if (!l || !node || l.locked || l.visible === false) {
        tr.nodes([]);
      } else {
        if (l.type === 'image' || l.type === 'svg') {
          tr.setAttrs({ enabledAnchors: ['top-left', 'top-right', 'bottom-left', 'bottom-right'], keepRatio: true, rotateEnabled: true, flipEnabled: false, shiftBehavior: 'none' });
        } else if (l.type === 'text') {
          tr.setAttrs({ enabledAnchors: ['top-left', 'top-right', 'bottom-left', 'bottom-right', 'middle-left', 'middle-right'], keepRatio: true, rotateEnabled: true, flipEnabled: false, shiftBehavior: 'none' });
        } else {
          tr.setAttrs({ enabledAnchors: ['top-left', 'top-center', 'top-right', 'middle-left', 'middle-right', 'bottom-left', 'bottom-center', 'bottom-right'], keepRatio: false, rotateEnabled: true, flipEnabled: false, shiftBehavior: 'default' });
        }
        tr.nodes([node]);
      }
    }
    tr.visible(tr.nodes().length > 0);
    S.uiLayer.batchDraw();
  }

  /* ---------------------------------------------------------------- history */
  function snapshot() { return JSON.stringify({ canvas: M.canvas, layers: M.layers }); }

  function commit() {
    const snap = snapshot();
    if (S.history[S.hIndex] === snap) { syncHistoryButtons(); return; }
    S.history.splice(S.hIndex + 1);
    S.history.push(snap);
    if (S.history.length > HISTORY_LIMIT) S.history.shift();
    S.hIndex = S.history.length - 1;
    if (S.hIndex > 0) state.dirty = true;
    syncHistoryButtons();
    renderLayers();
    scheduleScratchAutosave(); // no-op unless this design is the scratchpad (see editor.js)
  }

  function restore(index) {
    const d = JSON.parse(S.history[index]);
    M.canvas = d.canvas;
    M.layers = d.layers;
    S.hIndex = index;
    if (S.selectedId && S.selectedId !== 'canvas' && !getLayer(S.selectedId)) S.selectedId = null;
    rebuildAll();
    state.dirty = index > 0 || state.dirty;
    syncHistoryButtons();
    scheduleScratchAutosave();
  }

  function undo() {
    if (S.editing) return;
    if (S.cropping) { endCrop(false); return; }
    if (S.hIndex > 0) { restore(S.hIndex - 1); showToast('Undo', 900); }
  }

  function redo() {
    if (S.editing || S.cropping) return;
    if (S.hIndex < S.history.length - 1) { restore(S.hIndex + 1); showToast('Redo', 900); }
  }

  function syncHistoryButtons() {
    const d = S.dom;
    if (!d) return;
    d.undo.disabled = S.hIndex <= 0;
    d.redo.disabled = S.hIndex >= S.history.length - 1;
  }

  /* ---------------------------------------------------------------- adding things */
  function centerPos(w, h, at) {
    const cx = at ? at.x : M.canvas.w / 2;
    const cy = at ? at.y : M.canvas.h / 2;
    return { x: Math.round(cx - w / 2), y: Math.round(cy - h / 2) };
  }

  function stageToDesign(clientX, clientY) {
    const box = S.dom.viewport.getBoundingClientRect();
    const z = S.view.scale;
    return { x: (clientX - box.left - S.stage.x()) / z, y: (clientY - box.top - S.stage.y()) / z };
  }

  function addLayer(layer, { select: doSelect = true } = {}) {
    M.layers.push(layer);
    addNodeFor(layer);
    reorderNodes();
    S.contentLayer.batchDraw();
    commit();
    if (doSelect) select(layer.id);
    return layer;
  }

  async function addImageFile(file, at, offset = 0) {
    const isSvg = /\.svg$/i.test(file.name) || file.type === 'image/svg+xml';
    const asset = isSvg ? await createSvgAsset(file) : await createRasterAsset(file);
    // Fit inside the canvas (never bigger than 80% of it), keeping the picture's proportions.
    const k = Math.min(1, (M.canvas.w * 0.8) / asset.w, (M.canvas.h * 0.8) / asset.h);
    const w = Math.max(8, Math.round(asset.w * k));
    const h = Math.max(8, Math.round(asset.h * k));
    const p = centerPos(w, h, at ? { x: at.x + offset, y: at.y + offset } : (offset ? { x: M.canvas.w / 2 + offset, y: M.canvas.h / 2 + offset } : null));
    return addLayer({
      id: uid('l'), type: isSvg ? 'svg' : 'image', name: file.name.replace(/\.[^.]+$/, '') || (isSvg ? 'SVG' : 'Image'),
      asset: asset.id, crop: { x: 0, y: 0, w: 1, h: 1 },
      x: p.x, y: p.y, width: w, height: h, rotation: 0, visible: true, locked: false,
    });
  }

  function addText(text = 'Text') {
    const size = Math.max(16, Math.round(Math.min(M.canvas.w, M.canvas.h) * 0.06));
    const d = S.dom;
    const layer = {
      id: uid('l'), type: 'text', name: null, text, fontFamily: d.fontBtn.dataset.family || 'Arial', fontSize: size,
      fontStyle: 'normal', fill: '#202124', autoWidth: true, width: 200,
      x: 0, y: 0, rotation: 0, visible: true, locked: false,
    };
    const tmp = new Konva.Text({ text, fontFamily: layer.fontFamily, fontSize: size });
    layer.x = Math.round(M.canvas.w / 2 - tmp.width() / 2);
    layer.y = Math.round(M.canvas.h / 2 - tmp.height() / 2);
    tmp.destroy();
    addLayer(layer);
    startTextEdit(layer.id, true);
    return layer;
  }

  function addShape(kind) {
    const base = Math.round(Math.min(M.canvas.w, M.canvas.h) * 0.3);
    const linear = kind === 'line' || kind === 'arrow';
    const w = linear ? Math.round(base * 1.4) : base;
    const h = linear ? Math.max(24, Math.round(base * 0.2)) : base;
    const p = centerPos(w, h);
    const label = (SHAPES.find((s) => s.id === kind) || {}).label || 'Shape';
    return addLayer({
      id: uid('l'), type: 'shape', shape: kind, name: label,
      fill: linear ? '#202124' : ACCENT, stroke: '', strokeWidth: linear ? 8 : 0,
      x: p.x, y: p.y, width: w, height: h, rotation: 0, visible: true, locked: false,
    });
  }

  function removeLayer(id) {
    const i = M.layers.findIndex((l) => l.id === id);
    if (i < 0) return;
    const node = S.nodes.get(id);
    if (S.selectedId === id) S.transformer.nodes([]);
    if (node) node.destroy();
    S.nodes.delete(id);
    M.layers.splice(i, 1);
    if (S.selectedId === id) S.selectedId = null;
    updateTransformer();
    S.contentLayer.batchDraw();
    commit();
    syncToolbar();
  }

  function duplicateLayer(id) {
    const src = getLayer(id);
    if (!src) return;
    const copy = JSON.parse(JSON.stringify(src));
    copy.id = uid('l');
    copy.x += 24; copy.y += 24;
    if (copy.name) copy.name = `${copy.name} copy`;
    const i = M.layers.indexOf(src);
    M.layers.splice(i + 1, 0, copy);
    addNodeFor(copy);
    reorderNodes();
    S.contentLayer.batchDraw();
    commit();
    select(copy.id);
  }

  // Ctrl/Cmd+C: remembers the selected layer's data (not a live reference, so later edits
  // to the layer don't change what's on the "clipboard").
  function copySelected(id) {
    const src = getLayer(id);
    if (!src) return;
    S.clipboard = JSON.parse(JSON.stringify(src));
    showToast('Copied', 1200);
  }

  // Ctrl/Cmd+V: pastes the copied layer as a new layer, offset like duplicateLayer so it
  // doesn't land exactly on top of the original. Each paste offsets a bit further than the
  // last (cascading), the way most design tools handle repeated pastes.
  function pasteClipboard() {
    if (!S.clipboard) return;
    const copy = JSON.parse(JSON.stringify(S.clipboard));
    copy.id = uid('l');
    copy.x += 24; copy.y += 24;
    M.layers.push(copy);
    addNodeFor(copy);
    reorderNodes();
    S.contentLayer.batchDraw();
    commit();
    select(copy.id);
    S.clipboard.x = copy.x; S.clipboard.y = copy.y; // so the next paste cascades further still
  }

  function moveLayer(id, toIndex) {
    const from = M.layers.findIndex((l) => l.id === id);
    if (from < 0) return;
    const [l] = M.layers.splice(from, 1);
    M.layers.splice(clamp(toIndex, 0, M.layers.length), 0, l);
    reorderNodes();
    S.contentLayer.batchDraw();
    commit();
  }

  /* ---------------------------------------------------------------- layers panel */
  const TYPE_LABEL = { image: 'Image', svg: 'SVG', text: 'Text', shape: 'Shape' };

  function layerDisplayName(l) {
    if (l.name) return l.name;
    if (l.type === 'text') return (l.text || '').replace(/\s+/g, ' ').trim().slice(0, 40) || 'Text';
    return TYPE_LABEL[l.type];
  }

  function renderLayers() {
    const list = S.dom.layers;
    if (!list) return;
    const scroll = list.scrollTop;
    list.innerHTML = '';
    // Top of the list = top of the stack, like every design tool.
    for (let i = M.layers.length - 1; i >= 0; i--) {
      const l = M.layers[i];
      const li = document.createElement('li');
      li.className = 'd-layer' + (l.id === S.selectedId ? ' selected' : '') + (l.visible === false ? ' is-hidden' : '') + (l.locked ? ' is-locked' : '');
      li.dataset.id = l.id;
      li.draggable = true;
      li.innerHTML =
        '<span class="d-grip" title="Drag to reorder" aria-hidden="true"><svg viewBox="0 0 24 24" class="ui-icon"><path d="M9 6h.01M15 6h.01M9 12h.01M15 12h.01M9 18h.01M15 18h.01" stroke-width="3.2"/></svg></span>' +
        `<span class="d-type">${TYPE_LABEL[l.type]}</span>` +
        '<span class="d-name"></span>' +
        `<button class="d-ic d-eye" type="button" title="${l.visible === false ? 'Show' : 'Hide'}" aria-label="${l.visible === false ? 'Show' : 'Hide'} layer" aria-pressed="${l.visible === false}">${svgIcon(l.visible === false ? 'eyeOff' : 'eye')}</button>` +
        `<button class="d-ic d-lock" type="button" title="${l.locked ? 'Unlock' : 'Lock'}" aria-label="${l.locked ? 'Unlock' : 'Lock'} layer" aria-pressed="${!!l.locked}">${svgIcon(l.locked ? 'lock' : 'unlock')}</button>` +
        `<button class="d-ic d-del" type="button" title="Delete" aria-label="Delete layer">${svgIcon('trash')}</button>`;
      li.querySelector('.d-name').textContent = layerDisplayName(l);
      list.appendChild(li);
    }
    // The canvas itself: always last, can't be moved or deleted, but selecting it lets you resize it.
    const cv = document.createElement('li');
    cv.className = 'd-layer d-layer-canvas' + (S.selectedId === 'canvas' ? ' selected' : '');
    cv.dataset.id = 'canvas';
    cv.innerHTML = '<span class="d-grip d-grip-off" aria-hidden="true"></span><span class="d-type">Canvas</span><span class="d-name"></span>';
    cv.querySelector('.d-name').textContent = `${Math.round(M.canvas.w)} × ${Math.round(M.canvas.h)}`;
    list.appendChild(cv);
    list.scrollTop = scroll;
    S.dom.layersEmpty.hidden = M.layers.length > 0;
  }

  function scrollRowIntoView(id) {
    const row = S.dom.layers.querySelector(`.d-layer[data-id="${id}"]`);
    if (row && row.scrollIntoView) row.scrollIntoView({ block: 'nearest' });
  }

  function wireLayersPanel() {
    const list = S.dom.layers;
    let dragId = null;

    list.addEventListener('click', (e) => {
      const row = e.target.closest('.d-layer');
      if (!row) return;
      const id = row.dataset.id;
      const btn = e.target.closest('button.d-ic');
      if (btn && id !== 'canvas') {
        const l = getLayer(id);
        if (!l) return;
        if (btn.classList.contains('d-eye')) {
          l.visible = l.visible === false;
          const n = S.nodes.get(id); if (n) n.visible(l.visible);
          setInteractive(l);
          if (l.visible === false && S.selectedId === id) { /* keep row selected, no handles */ }
          updateTransformer(); S.contentLayer.batchDraw(); commit();
        } else if (btn.classList.contains('d-lock')) {
          l.locked = !l.locked;
          setInteractive(l);
          updateTransformer(); commit();
        } else if (btn.classList.contains('d-del')) {
          removeLayer(id);
        }
        return;
      }
      select(id);
    });

    list.addEventListener('dblclick', (e) => {
      const nameEl = e.target.closest('.d-name');
      const row = e.target.closest('.d-layer');
      if (!nameEl || !row || row.dataset.id === 'canvas') return;
      const l = getLayer(row.dataset.id);
      if (!l) return;
      nameEl.contentEditable = 'true';
      nameEl.focus();
      const r = document.createRange(); r.selectNodeContents(nameEl);
      const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      const done = (save) => {
        nameEl.contentEditable = 'false';
        nameEl.removeEventListener('blur', onBlur);
        nameEl.removeEventListener('keydown', onKey);
        if (save) { const v = nameEl.textContent.trim(); l.name = v || null; commit(); } else renderLayers();
      };
      const onBlur = () => done(true);
      const onKey = (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); nameEl.blur(); } else if (ev.key === 'Escape') { ev.preventDefault(); nameEl.removeEventListener('blur', onBlur); done(false); } };
      nameEl.addEventListener('blur', onBlur);
      nameEl.addEventListener('keydown', onKey);
    });

    // Drag to reorder
    const clearMarks = () => list.querySelectorAll('.drop-before,.drop-after').forEach((n) => n.classList.remove('drop-before', 'drop-after'));
    list.addEventListener('dragstart', (e) => {
      const row = e.target.closest && e.target.closest('.d-layer');
      if (!row || row.dataset.id === 'canvas' || S.cropping) { e.preventDefault(); return; }
      dragId = row.dataset.id;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/x-design-layer', dragId);
      row.classList.add('dragging');
    });
    list.addEventListener('dragend', () => { dragId = null; clearMarks(); list.querySelectorAll('.dragging').forEach((n) => n.classList.remove('dragging')); });
    list.addEventListener('dragover', (e) => {
      if (!dragId) return; // file drops are handled on the wrapper
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      const row = e.target.closest('.d-layer');
      clearMarks();
      if (!row) return;
      if (row.dataset.id === 'canvas') { const last = list.querySelector('.d-layer:not(.d-layer-canvas):last-of-type'); if (last) last.classList.add('drop-after'); return; }
      const r = row.getBoundingClientRect();
      row.classList.add(e.clientY < r.top + r.height / 2 ? 'drop-before' : 'drop-after');
    });
    list.addEventListener('drop', (e) => {
      if (!dragId) return;
      e.preventDefault();
      e.stopPropagation();
      const row = e.target.closest('.d-layer');
      const id = dragId;
      dragId = null;
      clearMarks();
      if (!row) return;
      // Rows run top-of-stack first, the model runs bottom first, hence the flips.
      const from = M.layers.findIndex((l) => l.id === id);
      let target;
      if (row.dataset.id === 'canvas') {
        target = 0;
      } else {
        const ti = M.layers.findIndex((l) => l.id === row.dataset.id);
        const r = row.getBoundingClientRect();
        const above = e.clientY < r.top + r.height / 2; // dropped on the upper half = above it in the stack = higher index
        target = above ? ti + 1 : ti;
        if (from < target) target -= 1;
      }
      if (target === from) return;
      moveLayer(id, target);
      renderLayers();
    });
  }

  /* ---------------------------------------------------------------- text editing */
  function startTextEdit(id, selectAll) {
    const l = getLayer(id);
    const node = S.nodes.get(id);
    if (!l || l.type !== 'text' || !node || l.locked) return;
    if (S.editing) finishTextEdit(true);
    select(id);
    const ta = document.createElement('textarea');
    ta.className = 'd-text-editor';
    ta.value = l.text;
    ta.spellcheck = false;
    S.dom.viewport.appendChild(ta);
    S.editing = { id, ta, original: l.text };
    node.hide();
    S.transformer.nodes([]);
    S.contentLayer.batchDraw();
    S.uiLayer.batchDraw();
    positionTextEditor();
    ta.focus();
    if (selectAll) ta.select(); else ta.setSelectionRange(ta.value.length, ta.value.length);
    ta.addEventListener('input', () => {
      l.text = ta.value;
      node.text(ta.value || ' ');
      positionTextEditor();
    });
    ta.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') { e.preventDefault(); finishTextEdit(false); }
      else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); finishTextEdit(true); }
    });
    ta.addEventListener('blur', () => { if (S.editing && S.editing.ta === ta) finishTextEdit(true); });
  }

  function positionTextEditor() {
    const ed = S.editing;
    if (!ed) return;
    const l = getLayer(ed.id);
    const node = S.nodes.get(ed.id);
    if (!l || !node) return;
    const z = S.view.scale;
    const pos = node.getAbsolutePosition();
    const ta = ed.ta;
    const w = (l.autoWidth ? Math.max(node.width(), 40) : l.width) * z;
    ta.style.left = `${pos.x}px`;
    ta.style.top = `${pos.y}px`;
    ta.style.width = `${w + (l.autoWidth ? 24 : 0)}px`;
    ta.style.height = `${Math.max(node.height() * z, l.fontSize * z * 1.2) + 4}px`;
    ta.style.fontFamily = l.fontFamily;
    ta.style.fontSize = `${l.fontSize * z}px`;
    ta.style.fontWeight = /bold/.test(l.fontStyle) ? '700' : '400';
    ta.style.fontStyle = /italic/.test(l.fontStyle) ? 'italic' : 'normal';
    ta.style.color = l.fill;
    ta.style.lineHeight = '1.2';
    ta.style.whiteSpace = l.autoWidth ? 'pre' : 'pre-wrap';
    ta.style.transform = `rotate(${l.rotation || 0}deg)`;
    ta.style.transformOrigin = 'left top';
  }

  function finishTextEdit(save) {
    const ed = S.editing;
    if (!ed) return;
    S.editing = null;
    const l = getLayer(ed.id);
    const node = S.nodes.get(ed.id);
    ed.ta.remove();
    if (!l || !node) return;
    if (!save) l.text = ed.original;
    if (!l.text.trim()) { removeLayer(ed.id); return; }
    applyLayerToNode(node, l);
    node.show();
    updateTransformer();
    S.contentLayer.batchDraw();
    commit();
    syncToolbar();
  }

  /* ---------------------------------------------------------------- crop */
  // Crop works on the selected image/SVG. The picture stays put while a crop
  // window (8 handles, and it can be dragged) is moved over it; everything
  // happens in the layer's own rotated frame, so rotated images crop correctly.
  const HANDLES = [
    ['tl', 0, 0, 'nwse-resize'], ['tc', 0.5, 0, 'ns-resize'], ['tr', 1, 0, 'nesw-resize'],
    ['ml', 0, 0.5, 'ew-resize'], ['mr', 1, 0.5, 'ew-resize'],
    ['bl', 0, 1, 'nesw-resize'], ['bc', 0.5, 1, 'ns-resize'], ['br', 1, 1, 'nwse-resize'],
  ];

  function startCrop() {
    if (S.cropping) { endCrop(true); return; }
    const l = getLayer(S.selectedId);
    if (!l || (l.type !== 'image' && l.type !== 'svg') || l.locked || l.visible === false) return;
    if (S.editing) finishTextEdit(true);
    const node = S.nodes.get(l.id);
    const a = S.assets.get(l.asset);
    const c = l.crop || { x: 0, y: 0, w: 1, h: 1 };
    const fullW = l.width / c.w;
    const fullH = l.height / c.h;
    const bounds = { x: -c.x * fullW, y: -c.y * fullH, w: fullW, h: fullH };
    const r = { x: 0, y: 0, w: l.width, h: l.height };

    const group = new Konva.Group({ x: l.x, y: l.y, rotation: l.rotation || 0 });
    const dim = new Konva.Image({ image: a.view, x: bounds.x, y: bounds.y, width: bounds.w, height: bounds.h, opacity: 0.32, listening: false });
    const sw = a.view.naturalWidth || a.view.width;
    const sh = a.view.naturalHeight || a.view.height;
    const bright = new Konva.Image({ image: a.view, listening: false });
    const frame = new Konva.Rect({ stroke: ACCENT, strokeWidth: 1.5, strokeScaleEnabled: false, fill: 'rgba(0,0,0,0.001)' });
    group.add(dim, bright, frame);
    const handles = HANDLES.map(([name, fx, fy, cursor]) => {
      const h = new Konva.Rect({ name, fill: '#fff', stroke: ACCENT, strokeWidth: 1.5, strokeScaleEnabled: false, cornerRadius: 1.5 });
      h._cursor = cursor; h._fx = fx; h._fy = fy;
      group.add(h);
      return h;
    });
    S.uiLayer.add(group);
    node.hide();
    S.cropping = { id: l.id, group, dim, bright, frame, handles, r, bounds, sw, sh, drag: null, node };

    const layout = () => {
      const z = S.view.scale;
      const cr = S.cropping;
      cr.bright.setAttrs({
        x: cr.r.x, y: cr.r.y, width: cr.r.w, height: cr.r.h,
        crop: {
          x: ((cr.r.x - cr.bounds.x) / cr.bounds.w) * cr.sw, y: ((cr.r.y - cr.bounds.y) / cr.bounds.h) * cr.sh,
          width: (cr.r.w / cr.bounds.w) * cr.sw, height: (cr.r.h / cr.bounds.h) * cr.sh,
        },
      });
      cr.frame.setAttrs({ x: cr.r.x, y: cr.r.y, width: cr.r.w, height: cr.r.h });
      const hs = 11 / z;
      cr.handles.forEach((h) => h.setAttrs({ width: hs, height: hs, x: cr.r.x + cr.r.w * h._fx - hs / 2, y: cr.r.y + cr.r.h * h._fy - hs / 2 }));
      S.uiLayer.batchDraw();
    };
    S.cropping.layout = layout;
    layout();

    const localPoint = () => {
      const p = S.stage.getPointerPosition();
      return group.getAbsoluteTransform().copy().invert().point(p);
    };
    const minSize = 12;
    const onMove = () => {
      const cr = S.cropping;
      if (!cr || !cr.drag) return;
      const p = localPoint();
      const d = cr.drag;
      const b = cr.bounds;
      if (d.type === 'move') {
        cr.r.x = clamp(d.r0.x + (p.x - d.p0.x), b.x, b.x + b.w - cr.r.w);
        cr.r.y = clamp(d.r0.y + (p.y - d.p0.y), b.y, b.y + b.h - cr.r.h);
      } else {
        let { x, y, w, h } = d.r0;
        let x2 = x + w; let y2 = y + h;
        const n = d.type;
        if (n.includes('l')) x = clamp(p.x, b.x, x2 - minSize);
        if (n.includes('r')) x2 = clamp(p.x, x + minSize, b.x + b.w);
        if (n.includes('t')) y = clamp(p.y, b.y, y2 - minSize);
        if (n.includes('b')) y2 = clamp(p.y, y + minSize, b.y + b.h);
        cr.r.x = x; cr.r.y = y; cr.r.w = x2 - x; cr.r.h = y2 - y;
      }
      layout();
    };
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      if (S.cropping) S.cropping.drag = null;
    };
    const begin = (type) => (e) => {
      e.cancelBubble = true;
      S.cropping.drag = { type, p0: localPoint(), r0: { ...S.cropping.r } };
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
    };
    frame.on('mousedown', begin('move'));
    frame.on('mouseenter', () => { S.stage.container().style.cursor = 'move'; });
    frame.on('mouseleave', () => { S.stage.container().style.cursor = ''; });
    handles.forEach((h) => {
      h.on('mousedown', begin(h.name()));
      h.on('mouseenter', () => { S.stage.container().style.cursor = h._cursor; });
      h.on('mouseleave', () => { S.stage.container().style.cursor = ''; });
    });

    document.body.classList.add('design-cropping');
    updateTransformer();
    S.contentLayer.batchDraw();
    syncToolbar();
    showToast('Drag the handles to crop — Enter to apply, Esc to cancel', 2600);
  }

  function endCrop(apply) {
    const cr = S.cropping;
    if (!cr) return;
    const l = getLayer(cr.id);
    if (apply && l) {
      const b = cr.bounds;
      l.crop = { x: (cr.r.x - b.x) / b.w, y: (cr.r.y - b.y) / b.h, w: cr.r.w / b.w, h: cr.r.h / b.h };
      const p = cr.group.getTransform().point({ x: cr.r.x, y: cr.r.y });
      l.x = p.x; l.y = p.y; l.width = cr.r.w; l.height = cr.r.h;
    }
    cr.group.destroy();
    S.cropping = null;
    document.body.classList.remove('design-cropping');
    S.stage.container().style.cursor = '';
    const node = S.nodes.get(cr.id);
    if (node && l) {
      S.transformer.nodes([]);
      node.destroy();
      S.nodes.delete(cr.id);
      const fresh = addNodeFor(l);
      reorderNodes();
      fresh.show();
    } else if (node) node.show();
    updateTransformer();
    S.contentLayer.batchDraw();
    S.uiLayer.batchDraw();
    if (apply && l) commit();
    renderLayers();
    syncToolbar();
  }

  /* ---------------------------------------------------------------- centre, canvas size */
  function centerSelected(axis) {
    const l = getLayer(S.selectedId);
    const node = l && S.nodes.get(l.id);
    if (!l || !node || l.locked) return;
    const rect = node.getClientRect({ skipStroke: true, skipShadow: true, relativeTo: S.content });
    if (axis === 'h') node.x(node.x() + (M.canvas.w / 2 - (rect.x + rect.width / 2)));
    else node.y(node.y() + (M.canvas.h / 2 - (rect.y + rect.height / 2)));
    l.x = node.x(); l.y = node.y();
    S.transformer.forceUpdate();
    S.contentLayer.batchDraw();
    S.uiLayer.batchDraw();
    commit();
  }

  function setCanvasSize(w, h) {
    w = clamp(Math.round(w) || M.canvas.w, MIN_DIM, MAX_DIM);
    h = clamp(Math.round(h) || M.canvas.h, MIN_DIM, MAX_DIM);
    if (w === Math.round(M.canvas.w) && h === Math.round(M.canvas.h)) { applyCanvasToStage(); return; }
    M.canvas.w = w; M.canvas.h = h;
    applyCanvasToStage();
    S.transformer.forceUpdate();
    commit();
  }

  function wireArtboard() {
    const ab = S.artboard;
    ab.on('mousedown touchstart', () => { if (!S.cropping && !S.editing) select('canvas'); });
    ab.on('transform', () => {
      S.dom.w.value = Math.round(ab.width() * ab.scaleX());
      S.dom.h.value = Math.round(ab.height() * ab.scaleY());
    });
    ab.on('transformend', () => {
      const w = clamp(Math.round(ab.width() * ab.scaleX()), MIN_DIM, MAX_DIM);
      const h = clamp(Math.round(ab.height() * ab.scaleY()), MIN_DIM, MAX_DIM);
      ab.scale({ x: 1, y: 1 });
      M.canvas.w = w; M.canvas.h = h;
      applyCanvasToStage();
      S.transformer.forceUpdate();
      commit();
    });
  }

  /* ---------------------------------------------------------------- toolbar */
  function syncToolbar() {
    const d = S.dom;
    if (!d) return;
    const l = getLayer(S.selectedId);
    const isText = !!l && l.type === 'text';
    const colorable = !!l && (l.type === 'text' || l.type === 'shape');
    const croppable = !!l && (l.type === 'image' || l.type === 'svg') && !l.locked && l.visible !== false;
    const busy = !!S.cropping;
    d.fontBtn.disabled = !isText || busy;
    d.fontSize.disabled = !isText || busy;
    d.bold.disabled = !isText || busy;
    d.italic.disabled = !isText || busy;
    d.color.disabled = !colorable || busy;
    d.crop.disabled = !croppable && !busy;
    d.crop.classList.toggle('active', busy);
    d.crop.querySelector('span').textContent = busy ? 'Done' : 'Crop';
    d.cropCancel.hidden = !busy;
    d.clear.disabled = !M.layers.length || busy;
    d.centerH.disabled = !l || !!l.locked || busy;
    d.centerV.disabled = !l || !!l.locked || busy;
    if (isText) {
      d.fontBtn.querySelector('.d-font-name').textContent = l.fontFamily;
      d.fontBtn.dataset.family = l.fontFamily;
      d.fontBtn.style.fontFamily = l.fontFamily;
      if (document.activeElement !== d.fontSize) d.fontSize.value = Math.round(l.fontSize);
      d.bold.classList.toggle('active', /bold/.test(l.fontStyle));
      d.italic.classList.toggle('active', /italic/.test(l.fontStyle));
      d.bold.setAttribute('aria-pressed', String(/bold/.test(l.fontStyle)));
      d.italic.setAttribute('aria-pressed', String(/italic/.test(l.fontStyle)));
    } else {
      d.bold.classList.remove('active'); d.italic.classList.remove('active');
      d.bold.setAttribute('aria-pressed', 'false'); d.italic.setAttribute('aria-pressed', 'false');
    }
    if (colorable) d.color.value = l.fill;
    for (const id of ['dAddImage', 'dAddText', 'dShapesBtn', 'dApply']) { const b = $(id); if (b) b.disabled = busy; }
    syncHistoryButtons();
  }

  function selectedText() {
    const l = getLayer(S.selectedId);
    return l && l.type === 'text' ? l : null;
  }

  function updateSelectedTextLive(patch) {
    const l = selectedText();
    if (!l) return null;
    Object.assign(l, patch);
    const n = S.nodes.get(l.id);
    if (n) applyLayerToNode(n, l);
    S.transformer.forceUpdate();
    S.contentLayer.batchDraw();
    S.uiLayer.batchDraw();
    positionTextEditor();
    return l;
  }

  function setFontStyle(bold, italic) {
    const fs = [bold ? 'bold' : '', italic ? 'italic' : ''].filter(Boolean).join(' ') || 'normal';
    return fs;
  }

  /* ---------------------------------------------------------------- fonts */
  function familyCss(f) { return /^[\w-]+$/.test(f) ? f : `"${f}"`; }

  async function addFontFile(file) {
    let base = file.name.replace(/\.[^.]+$/, '').replace(/[^\w\- ]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Custom font';
    let name = base;
    let n = 2;
    while (S.customFonts.includes(name) || BUILTIN_FONTS.includes(name)) name = `${base} ${n++}`;
    const face = new FontFace(name, await file.arrayBuffer());
    await face.load();
    document.fonts.add(face);
    S.customFonts.push(name);
    S.fontBlobs.set(name, file);
    S.dom.fontBtn.dataset.family = name;
    const t = selectedText();
    if (t) { t.fontFamily = name; commit(); const nd = S.nodes.get(t.id); if (nd) applyLayerToNode(nd, t); }
    refreshTextNodes();
    syncToolbar();
    showToast(`Font “${name}” added${t ? ' and applied' : ''}`, 2600);
    return name;
  }

  function buildFontMenu() {
    const menu = S.dom.fontMenu;
    const cur = (selectedText() || {}).fontFamily || S.dom.fontBtn.dataset.family;
    menu.innerHTML = '';
    const addGroup = (title, fonts) => {
      if (!fonts.length) return;
      const h = document.createElement('div'); h.className = 'd-menu-title'; h.textContent = title; menu.appendChild(h);
      for (const f of fonts) {
        const b = document.createElement('button');
        b.type = 'button'; b.className = 'd-menu-item' + (f === cur ? ' is-current' : '');
        b.dataset.family = f; b.style.fontFamily = familyCss(f); b.textContent = f;
        menu.appendChild(b);
      }
    };
    addGroup('Your fonts', S.customFonts);
    addGroup('Fonts', BUILTIN_FONTS);
    const add = document.createElement('button');
    add.type = 'button'; add.className = 'd-menu-item d-menu-add'; add.dataset.action = 'add-font';
    add.innerHTML = '<span>＋ Add a font file…</span><small>or drag &amp; drop .ttf .otf .woff .woff2 onto the canvas</small>';
    menu.appendChild(add);
  }

  /* ---------------------------------------------------------------- popovers */
  let openMenu = null;
  function closeMenu() {
    if (!openMenu) return;
    openMenu.menu.hidden = true;
    openMenu.btn.setAttribute('aria-expanded', 'false');
    openMenu = null;
  }
  function toggleMenu(btn, menu, beforeOpen) {
    if (openMenu && openMenu.menu === menu) { closeMenu(); return; }
    closeMenu();
    if (beforeOpen) beforeOpen();
    menu.hidden = false;
    btn.setAttribute('aria-expanded', 'true');
    const b = btn.getBoundingClientRect();
    const host = menu.offsetParent ? menu.offsetParent.getBoundingClientRect() : { left: 0, top: 0 };
    menu.style.left = `${Math.max(4, Math.min(b.left - host.left, window.innerWidth - menu.offsetWidth - 8 - host.left))}px`;
    menu.style.top = `${b.bottom - host.top + 4}px`;
    openMenu = { btn, menu };
  }

  /* ---------------------------------------------------------------- files: drop, paste, pick */
  const isFontFile = (f) => /\.(ttf|otf|woff2?)$/i.test(f.name);
  const isImageLike = (f) => (f.type && f.type.startsWith('image/')) || /\.(png|jpe?g|webp|gif|bmp|avif|svg)$/i.test(f.name);

  async function handleFiles(files, at, items) {
    const list = Array.from(files || []);
    if (!list.length) return;
    let i = 0;
    let added = 0;
    let bad = 0;
    // Dropped outside the artboard? Use its centre instead.
    const dropAt = at && at.x >= 0 && at.y >= 0 && at.x <= M.canvas.w && at.y <= M.canvas.h ? at : null;
    for (const f of list) {
      try {
        if (/\.json$/i.test(f.name) && typeof isDesignTemplateFile === 'function' && await isDesignTemplateFile(f)) {
          // Same handle-capture as the homepage dropzone (see tryGetDropHandle in editor.js) —
          // best-effort, so a drag source or browser that can't supply one just opens with none.
          const handle = items && typeof tryGetDropHandle === 'function' ? await tryGetDropHandle(items[list.indexOf(f)]) : null;
          openTemplateFile(f, handle); // replaces the design on screen (asks first if it has unsaved changes)
          return;
        }
        if (isFontFile(f)) { await addFontFile(f); added++; }
        else if (isImageLike(f)) { await addImageFile(f, dropAt, i * 24); i++; added++; }
        else bad++;
      } catch (err) {
        console.error(err);
        bad++;
        showToast(`Couldn’t add “${f.name}” — ${err && err.message ? err.message : 'unsupported file'}`, 4000);
      }
    }
    if (bad && !added) showToast('Drop images, SVGs or font files (.ttf .otf .woff .woff2)', 3500);
  }

  function wireFiles() {
    const w = S.dom.wrapper;
    const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
    w.addEventListener('dragover', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
      w.classList.add('drag-over');
    });
    w.addEventListener('dragleave', (e) => {
      if (e.relatedTarget && w.contains(e.relatedTarget)) return;
      w.classList.remove('drag-over');
    });
    w.addEventListener('drop', (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      w.classList.remove('drag-over');
      if (S.cropping) { showToast('Finish cropping first'); return; }
      handleFiles(e.dataTransfer.files, stageToDesign(e.clientX, e.clientY), e.dataTransfer.items);
    });
    document.addEventListener('paste', (e) => {
      if (!S.active) return;
      const t = e.target;
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
      // Copies made with the Async Clipboard API (e.g. a web app's "Copy Image" button, which
      // is how Canva's export-to-clipboard works) don't always populate clipboardData.files —
      // some browsers only expose them through clipboardData.items. Merge both sources so
      // either kind of copy pastes correctly; dedupe on type+size in case a browser reports
      // the same blob in both places.
      const fromFiles = Array.from((e.clipboardData && e.clipboardData.files) || []);
      const fromItems = Array.from((e.clipboardData && e.clipboardData.items) || [])
        .filter((it) => it.kind === 'file')
        .map((it) => it.getAsFile())
        .filter(Boolean);
      const seen = new Set(fromFiles.map((f) => `${f.type}:${f.size}`));
      const files = fromFiles
        .concat(fromItems.filter((f) => !seen.has(`${f.type}:${f.size}`)))
        .filter(isImageLike);
      if (files.length) { e.preventDefault(); handleFiles(files, null); return; }
      // No image on the OS clipboard — fall back to the app's own layer clipboard, so
      // Ctrl/Cmd+V still duplicates whatever layer was last copied with Ctrl/Cmd+C.
      if (S.clipboard) { e.preventDefault(); pasteClipboard(); }
    });
    S.dom.imageInput.addEventListener('change', () => { handleFiles(S.dom.imageInput.files, null); S.dom.imageInput.value = ''; });
    S.dom.fontFile.addEventListener('change', () => { handleFiles(Array.from(S.dom.fontFile.files).filter(isFontFile), null); S.dom.fontFile.value = ''; });
  }

  /* ---------------------------------------------------------------- keyboard, wheel, pan */
  let nudgeTimer = 0;
  function onKeyDown(e) {
    if (!S.active) return;
    const t = e.target;
    if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable)) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); exportImage(); return; }
    if (S.cropping) {
      if (e.key === 'Enter') { e.preventDefault(); endCrop(true); }
      else if (e.key === 'Escape') { e.preventDefault(); endCrop(false); }
      return;
    }
    if (e.key === 'Escape') { closeMenu(); if (S.selectedId) select(null); return; }
    // Paste doesn't need a layer selected, so it's handled before the selection guard below.
    // Ctrl/Cmd+V is handled by the document 'paste' listener below (wireFiles), not here —
    // preventDefault()ing the keydown would cancel the browser's native paste action, so the
    // 'paste' event (which is how an OS-clipboard image, e.g. from Canva's "Copy Image", gets
    // in) would never fire. That listener falls back to pasteClipboard() itself when the OS
    // clipboard has no image.
    const l = getLayer(S.selectedId);
    if (!l) return;
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); removeLayer(l.id); return; }
    if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); duplicateLayer(l.id); return; }
    if (mod && e.key.toLowerCase() === 'c') { e.preventDefault(); copySelected(l.id); return; }
    if (e.key === 'Enter' && l.type === 'text') { e.preventDefault(); startTextEdit(l.id, false); return; }
    const step = e.shiftKey ? 10 : 1;
    const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
    const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
    if ((dx || dy) && !l.locked) {
      e.preventDefault();
      l.x += dx; l.y += dy;
      const n = S.nodes.get(l.id);
      if (n) { n.x(l.x); n.y(l.y); }
      S.transformer.forceUpdate();
      S.contentLayer.batchDraw(); S.uiLayer.batchDraw();
      clearTimeout(nudgeTimer);
      nudgeTimer = setTimeout(commit, 350);
    }
  }

  function wireViewport() {
    const vp = S.dom.viewport;
    vp.addEventListener('wheel', (e) => {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) {
        const box = vp.getBoundingClientRect();
        zoomTo(S.view.scale * Math.exp(-e.deltaY * 0.012), { x: e.clientX - box.left, y: e.clientY - box.top });
      } else {
        S.view.auto = false;
        const dx = e.shiftKey && !e.deltaX ? e.deltaY : e.deltaX;
        const dy = e.shiftKey && !e.deltaX ? 0 : e.deltaY;
        S.stage.position({ x: S.stage.x() - dx, y: S.stage.y() - dy });
        S.stage.batchDraw();
        positionTextEditor();
      }
    }, { passive: false });

    // Drag the grey area to pan; click it to deselect.
    let pan = null;
    S.stage.on('mousedown touchstart', (e) => {
      if (e.target !== S.stage || S.cropping) return;
      if (S.editing) finishTextEdit(true);
      const p = S.stage.getPointerPosition();
      pan = { x: p.x, y: p.y, sx: S.stage.x(), sy: S.stage.y(), moved: false };
    });
    window.addEventListener('mousemove', (e) => {
      if (!pan) return;
      const box = vp.getBoundingClientRect();
      const px = e.clientX - box.left; const py = e.clientY - box.top;
      if (!pan.moved && Math.hypot(px - pan.x, py - pan.y) < 3) return;
      pan.moved = true;
      S.view.auto = false;
      S.stage.position({ x: pan.sx + px - pan.x, y: pan.sy + py - pan.y });
      S.stage.batchDraw();
      vp.style.cursor = 'grabbing';
    });
    window.addEventListener('mouseup', () => {
      if (!pan) return;
      if (!pan.moved) select(null);
      pan = null;
      vp.style.cursor = '';
    });
  }

  /* ---------------------------------------------------------------- export */
  // Target size = a number plus a KB | MB switch, so the unit is never a guess.
  function readTarget() {
    const raw = ($('targetSize').value || '').trim();
    if (!raw) return { bytes: null };
    const n = parseFloat(raw.replace(',', '.'));
    if (!isFinite(n) || n <= 0) return { bytes: null, invalid: true };
    const mb = $('targetUnit').dataset.value === 'mb';
    return { bytes: Math.max(1, Math.round(n * (mb ? 1024 * 1024 : 1024))) };
  }

  function formatBytesShort(n) {
    if (n >= 1024 * 1024) return `${(n / 1024 / 1024).toFixed(2)} MB`;
    if (n >= 1024) return `${Math.round(n / 1024)} KB`;
    return `${n} B`;
  }

  // The same model, drawn again off-screen at 1:1, with none of the editing
  // chrome (handles, selection, hidden layers) — this is what gets saved.
  function renderFullCanvas(scale = 1) {
    const host = document.createElement('div');
    const stage = new Konva.Stage({ container: host, width: M.canvas.w, height: M.canvas.h });
    const layer = new Konva.Layer();
    stage.add(layer);
    layer.add(new Konva.Rect({ x: 0, y: 0, width: M.canvas.w, height: M.canvas.h, fill: M.canvas.fill }));
    const group = new Konva.Group({ clipFunc: (ctx) => ctx.rect(0, 0, M.canvas.w, M.canvas.h) });
    layer.add(group);
    for (const l of M.layers) if (l.visible !== false) group.add(makeNode(l, true));
    layer.draw();
    // pixelRatio renders at `scale`x the stage's own (1x) pixel dimensions, so
    // text/strokes/images all come out sharp instead of just being upscaled.
    const out = layer.toCanvas({ pixelRatio: scale });
    stage.destroy();
    return out;
  }

  const toBlob = (canvas, type, q) => new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('The browser could not encode the image'))), type, q);
  });

  function resample(src, scale) {
    if (scale >= 0.999) return src;
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(src.width * scale));
    c.height = Math.max(1, Math.round(src.height * scale));
    const ctx = c.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(src, 0, 0, c.width, c.height);
    return c;
  }

  // JPEG: highest quality whose file still fits. Returns null if even the lowest quality is too big.
  async function fitJpegQuality(canvas, target, progress) {
    const top = await toBlob(canvas, 'image/jpeg', 1);
    if (!target || top.size <= target) return { blob: top, quality: 1 };
    const floor = await toBlob(canvas, 'image/jpeg', 0.02);
    if (floor.size > target) return null;
    let best = { blob: floor, quality: 0.02 };
    let lo = 0.02; let hi = 1;
    for (let i = 0; i < 10; i++) {
      const q = (lo + hi) / 2;
      const b = await toBlob(canvas, 'image/jpeg', q);
      if (progress) progress(i + 1, 10);
      if (b.size <= target) { best = { blob: b, quality: q }; lo = q; if (b.size >= target * 0.97) break; } else hi = q;
    }
    return best;
  }

  // A size target is only ever used with JPG (the interface switches to JPG when one is
  // set): quality is lowered to fit and, only if even the lowest quality is too big, the pixel
  // size is reduced too. PNG (and no target) is always saved losslessly / at full quality.
  async function encodeForTarget(full, format, target, progress) {
    const type = format === 'jpg' ? 'image/jpeg' : 'image/png';
    if (!target || format !== 'jpg') {
      const blob = await toBlob(full, type, 1);
      return { blob, quality: 1, scale: 1, width: full.width, height: full.height };
    }
    let scale = 1;
    let canvas = full;
    for (let attempt = 0; attempt < 8; attempt++) {
      const r = await fitJpegQuality(canvas, target, progress);
      if (r) return { blob: r.blob, quality: r.quality, scale, width: canvas.width, height: canvas.height };
      // Even the lowest quality is too big: shrink the pixels too (size scales ~ with area).
      const floor = await toBlob(canvas, 'image/jpeg', 0.02);
      scale *= Math.max(0.1, Math.min(0.9, Math.sqrt(target / floor.size) * 0.92));
      canvas = resample(full, scale);
    }
    const blob = await toBlob(canvas, 'image/jpeg', 0.02);
    return { blob, quality: 0.02, scale, width: canvas.width, height: canvas.height };
  }

  function currentFormat() {
    const v = $('designFormat').dataset.value;
    return v === 'jpg' || v === 'template' ? v : 'png';
  }

  // Export resolution multiplier (1x/2x/3x). Doesn't apply to the Template format
  // (that's the vector layer data, not a raster render), so it's ignored there.
  function currentScale() {
    const n = parseInt($('designScale').dataset.value, 10);
    return [1, 2, 3].includes(n) ? n : 1;
  }

  const safeName = () => (state.title || 'design').replace(/[\\/:*?"<>|]+/g, '_');

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  // Save / Export does whatever the PNG | JPG | Template switch says.
  async function exportImage() {
    if (!S.active || S.export.busy) return;
    if (S.editing) finishTextEdit(true);
    if (S.cropping) endCrop(true);
    const format = currentFormat();
    if (format === 'template') return saveTemplate();
    const tgt = readTarget();
    if (tgt.invalid) { showToast('Target size: enter a number greater than 0 (then pick KB or MB)', 3500); $('targetSize').focus(); return; }
    const target = tgt.bytes;
    const scale = currentScale();
    S.export.busy = true;
    el.saveBtn.disabled = true;
    try {
      await (document.fonts && document.fonts.ready);
      showToast(target ? `Fitting to ${formatBytesShort(target)}…` : 'Exporting…', 60000);
      await new Promise((r) => setTimeout(r, 30)); // let the toast paint
      const full = renderFullCanvas(scale);
      const r = await encodeForTarget(full, format, target, (i, n) => showToast(`Fitting to ${formatBytesShort(target)}… (${i}/${n})`, 60000));
      const name = `${safeName()}.${format}`;
      downloadBlob(r.blob, name);
      // (An image export is not a save of the design itself: unsaved-changes warnings stay
      // until the template is saved.)
      const parts = [`Saved ${name}`, formatBytesShort(r.blob.size)];
      if (scale > 1) parts.push(`${scale}× (${r.width} × ${r.height}px)`);
      if (format === 'jpg') parts.push(`quality ${Math.round(r.quality * 100)}%`);
      if (r.scale < 0.999) parts.push(`scaled to ${r.width} × ${r.height}`);
      if (target && r.blob.size > target) parts.push('couldn’t reach the target');
      showToast(parts.join(' · '), 5000);
      S.export.last = { name, size: r.blob.size, quality: r.quality, scale: r.scale, width: r.width, height: r.height, format, blob: r.blob };
      return S.export.last;
    } catch (err) {
      console.error(err);
      showToast(`Export failed — ${err && err.message ? err.message : 'see the console'}`, 5000);
    } finally {
      S.export.busy = false;
      el.saveBtn.disabled = false;
    }
  }

  /* ---------------------------------------------------------------- template files */
  const blobToDataUrl = (blob) => new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error || new Error('Could not read a file'));
    r.readAsDataURL(blob);
  });

  const dataUrlToBlob = async (url) => (await fetch(url)).blob();

  // Everything needed to rebuild the design: the layers, and (embedded) every
  // picture and custom font they use, so the file works on its own anywhere.
  async function serializeTemplate() {
    const usedAssets = new Set(M.layers.filter((l) => l.asset).map((l) => l.asset));
    const assets = [];
    for (const id of usedAssets) {
      const a = S.assets.get(id);
      if (a && a.blob) assets.push({ id, kind: a.kind, mime: a.mime, data: await blobToDataUrl(a.blob) });
    }
    const usedFonts = new Set(M.layers.filter((l) => l.type === 'text').map((l) => l.fontFamily));
    const fonts = [];
    for (const [family, blob] of S.fontBlobs) {
      if (usedFonts.has(family)) fonts.push({ family, data: await blobToDataUrl(blob) });
    }
    return { type: TEMPLATE_TYPE, version: TEMPLATE_VERSION, title: state.title || 'Untitled design', savedAt: new Date().toISOString(), canvas: M.canvas, layers: M.layers, assets, fonts };
  }

  async function saveTemplate() {
    if (!S.active || S.export.busy) return;
    if (S.editing) finishTextEdit(true);
    if (S.cropping) endCrop(true);
    S.export.busy = true;
    el.saveBtn.disabled = true;
    try {
      showToast('Saving template…', 60000);
      await new Promise((r) => setTimeout(r, 30));
      const json = JSON.stringify(await serializeTemplate());
      const blob = new Blob([json], { type: 'application/json' });
      const name = `${safeName()}.design.json`;
      downloadBlob(blob, name);
      state.dirty = false; // the layered design itself is now saved
      showToast(`Saved ${name} · ${formatBytesShort(blob.size)} — reopen it with Open File or by dropping it here`, 5000);
      S.export.last = { name, size: blob.size, format: 'template', blob };
      return S.export.last;
    } catch (err) {
      console.error(`Save template failed: ${err && err.name}: ${err && err.message}`, err);
      showToast(`Couldn’t save the template — ${err && err.message ? err.message : (err && err.name) || 'see the console'}`, 5000);
    } finally {
      S.export.busy = false;
      el.saveBtn.disabled = false;
    }
  }

  async function registerFontData(family, dataUrl) {
    if (S.customFonts.includes(family) || BUILTIN_FONTS.includes(family)) return; // never shadow a built-in font
    const blob = await dataUrlToBlob(dataUrl);
    const face = new FontFace(family, await blob.arrayBuffer());
    await face.load();
    document.fonts.add(face);
    S.customFonts.push(family);
    S.fontBlobs.set(family, blob);
  }

  // Reads a template file into ready-to-use pieces WITHOUT touching the design on
  // screen, so a damaged file simply fails and leaves everything as it was.
  async function prepareTemplate(text) {
    let d;
    try { d = JSON.parse(text); } catch (e) { throw new Error('this isn’t a valid design file'); }
    if (!d || d.type !== TEMPLATE_TYPE || !Array.isArray(d.layers) || !d.canvas) throw new Error('this isn’t a design file');
    if (Number(d.version) > TEMPLATE_VERSION) throw new Error('it was saved by a newer version of this editor');
    const num = (v, dflt) => (isFinite(Number(v)) && v !== null && v !== '' ? Number(v) : dflt);
    const pos = (v, dflt) => Math.max(1, num(v, dflt));
    const colour = (v, dflt) => (typeof v === 'string' && /^#[0-9a-f]{6}$/i.test(v) ? v : dflt);
    const canvas = {
      w: clamp(Math.round(num(d.canvas.w, DEFAULT_W)), MIN_DIM, MAX_DIM),
      h: clamp(Math.round(num(d.canvas.h, DEFAULT_H)), MIN_DIM, MAX_DIM),
      fill: colour(d.canvas.fill, '#ffffff'),
    };
    const missingFonts = [];
    for (const f of Array.isArray(d.fonts) ? d.fonts : []) {
      if (!f || typeof f.family !== 'string' || typeof f.data !== 'string') continue;
      try { await registerFontData(f.family, f.data); } catch (e) { missingFonts.push(f.family); }
    }
    const assets = new Map();
    for (const a of Array.isArray(d.assets) ? d.assets : []) {
      if (!a || typeof a.id !== 'string' || typeof a.data !== 'string') continue;
      try {
        const blob = await dataUrlToBlob(a.data);
        const file = new File([blob], a.id, { type: a.kind === 'svg' ? 'image/svg+xml' : (a.mime || blob.type || 'image/png') });
        assets.set(a.id, a.kind === 'svg' ? await decodeSvgAsset(file, a.id) : await decodeRasterAsset(file, a.id));
      } catch (e) { /* a broken picture just drops the layers that use it */ }
    }
    const seen = new Set();
    const layers = [];
    let dropped = 0;
    for (const l of d.layers) {
      if (!l || typeof l !== 'object') { dropped++; continue; }
      const base = {
        id: typeof l.id === 'string' && !seen.has(l.id) ? l.id : uid('l'),
        type: l.type, name: typeof l.name === 'string' && l.name ? l.name : null,
        x: num(l.x, 0), y: num(l.y, 0), rotation: num(l.rotation, 0), visible: l.visible !== false, locked: !!l.locked,
      };
      if (l.type === 'image' || l.type === 'svg') {
        if (!assets.has(l.asset)) { dropped++; continue; }
        const c = l.crop || {};
        const crop = { x: clamp(num(c.x, 0), 0, 1), y: clamp(num(c.y, 0), 0, 1), w: clamp(num(c.w, 1), 0.001, 1), h: clamp(num(c.h, 1), 0.001, 1) };
        layers.push({ ...base, asset: l.asset, crop, width: pos(l.width, 100), height: pos(l.height, 100) });
      } else if (l.type === 'text') {
        layers.push({
          ...base, text: String(l.text == null ? '' : l.text), fontFamily: String(l.fontFamily || 'Arial'),
          fontSize: clamp(num(l.fontSize, 32), 4, 999), fontStyle: ['normal', 'bold', 'italic', 'bold italic'].includes(l.fontStyle) ? l.fontStyle : 'normal',
          fill: colour(l.fill, '#202124'), autoWidth: l.autoWidth !== false, width: pos(l.width, 200),
        });
      } else if (l.type === 'shape' && SHAPES.some((s) => s.id === l.shape)) {
        const shapeLayer = {
          ...base, shape: l.shape, fill: colour(l.fill, ACCENT), stroke: colour(l.stroke, ''), strokeWidth: Math.max(0, num(l.strokeWidth, 0)),
          width: pos(l.width, 100), height: pos(l.height, 100),
        };
        if (l.shape === 'rect') shapeLayer.radius = Math.max(0, num(l.radius, 0)); // the only shape that uses it
        layers.push(shapeLayer);
      } else { dropped++; continue; }
      seen.add(base.id);
    }
    return { canvas, layers, assets, title: typeof d.title === 'string' ? d.title : '', dropped, missingFonts };
  }

  // Wipes the design (undoable — it's one history step).
  function clearAll() {
    if (!S.active) return;
    if (S.editing) finishTextEdit(false);
    if (S.cropping) endCrop(false);
    if (!M.layers.length) { showToast('The canvas is already empty', 1800); return; }
    S.transformer.nodes([]);
    S.content.destroyChildren();
    S.nodes.clear();
    M.layers = [];
    S.selectedId = null;
    updateTransformer();
    S.contentLayer.batchDraw();
    commit();
    syncToolbar();
    // Clearing the scratchpad also drops any "scratchpad_<dropped file>" name it picked up,
    // back to the plain default (not for a design opened outside the scratchpad).
    if (state.scratchId === 'design' && state.title !== SCRATCH_DEFAULT_NAMES.design) {
      state.title = SCRATCH_DEFAULT_NAMES.design;
      el.fileName.textContent = state.title;
    }
    showToast('Cleared — Ctrl+Z brings it back', 2600);
  }

  /* ---------------------------------------------------------------- setup */
  function ensureDom() {
    if (S.dom) return S.dom;
    const d = {
      wrapper: $('designWrapper'), viewport: $('designViewport'), layers: $('designLayers'), layersEmpty: $('designLayersEmpty'),
      addImage: $('dAddImage'), addText: $('dAddText'), shapesBtn: $('dShapesBtn'), shapesMenu: $('dShapesMenu'),
      undo: $('dUndo'), redo: $('dRedo'), clear: $('dClear'), crop: $('dCrop'), cropCancel: $('dCropCancel'),
      centerH: $('dCenterH'), centerV: $('dCenterV'),
      w: $('dW'), h: $('dH'), apply: $('dApply'), canvasColor: $('dCanvasColor'),
      fontBtn: $('dFontBtn'), fontMenu: $('dFontMenu'), fontSize: $('dFontSize'), bold: $('dBold'), italic: $('dItalic'), color: $('dColor'),
      zoomOut: $('dZoomOut'), zoomIn: $('dZoomIn'), zoomLabel: $('dZoomLabel'), fit: $('dFit'),
      imageInput: $('dImageInput'), fontFile: $('dFontFile'),
    };
    S.dom = d;
    wireToolbar();
    wireLayersPanel();
    wireFiles();
    document.addEventListener('keydown', onKeyDown);
    document.addEventListener('pointerdown', (e) => {
      if (openMenu && !openMenu.menu.contains(e.target) && !openMenu.btn.contains(e.target)) closeMenu();
    }, true);
    return d;
  }

  function wireToolbar() {
    const d = S.dom;
    d.addImage.addEventListener('click', () => d.imageInput.click());
    d.addText.addEventListener('click', () => addText());

    d.shapesMenu.innerHTML = SHAPES.map((s) => `<button type="button" class="d-menu-item" data-shape="${s.id}">${s.label}</button>`).join('');
    d.shapesBtn.addEventListener('click', () => toggleMenu(d.shapesBtn, d.shapesMenu));
    d.shapesMenu.addEventListener('click', (e) => {
      const b = e.target.closest('[data-shape]');
      if (b) { closeMenu(); addShape(b.dataset.shape); }
    });

    d.undo.addEventListener('click', undo);
    d.redo.addEventListener('click', redo);
    d.crop.addEventListener('click', startCrop);
    d.cropCancel.addEventListener('click', () => endCrop(false));
    d.centerH.addEventListener('click', () => centerSelected('h'));
    d.centerV.addEventListener('click', () => centerSelected('v'));

    const applySize = () => setCanvasSize(parseFloat(d.w.value), parseFloat(d.h.value));
    d.apply.addEventListener('click', applySize);
    for (const inp of [d.w, d.h]) inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); applySize(); inp.blur(); } });

    d.canvasColor.addEventListener('input', () => { M.canvas.fill = d.canvasColor.value; S.artboard.fill(M.canvas.fill); S.bgLayer.batchDraw(); });
    d.canvasColor.addEventListener('change', () => { M.canvas.fill = d.canvasColor.value; commit(); });

    d.fontBtn.addEventListener('click', () => toggleMenu(d.fontBtn, d.fontMenu, buildFontMenu));
    d.fontMenu.addEventListener('click', (e) => {
      const b = e.target.closest('.d-menu-item');
      if (!b) return;
      if (b.dataset.action === 'add-font') { closeMenu(); d.fontFile.click(); return; }
      closeMenu();
      const fam = b.dataset.family;
      d.fontBtn.dataset.family = fam;
      if (updateSelectedTextLive({ fontFamily: fam })) { commit(); syncToolbar(); }
    });

    d.fontSize.addEventListener('input', () => {
      const v = parseFloat(d.fontSize.value);
      if (isFinite(v) && v >= 4) updateSelectedTextLive({ fontSize: Math.min(999, v) });
    });
    d.fontSize.addEventListener('change', () => {
      const l = selectedText();
      if (!l) return;
      const v = clamp(parseFloat(d.fontSize.value) || l.fontSize, 4, 999);
      updateSelectedTextLive({ fontSize: v });
      d.fontSize.value = Math.round(v);
      commit();
    });
    d.fontSize.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); d.fontSize.blur(); } });

    const toggleStyle = (which) => () => {
      const l = selectedText();
      if (!l) return;
      const bold = /bold/.test(l.fontStyle);
      const italic = /italic/.test(l.fontStyle);
      updateSelectedTextLive({ fontStyle: setFontStyle(which === 'bold' ? !bold : bold, which === 'italic' ? !italic : italic) });
      commit();
      syncToolbar();
    };
    d.bold.addEventListener('click', toggleStyle('bold'));
    d.italic.addEventListener('click', toggleStyle('italic'));

    d.color.addEventListener('input', () => {
      const l = getLayer(S.selectedId);
      if (!l || (l.type !== 'text' && l.type !== 'shape')) return;
      l.fill = d.color.value;
      const n = S.nodes.get(l.id); if (n) applyLayerToNode(n, l);
      S.contentLayer.batchDraw();
      positionTextEditor();
    });
    d.color.addEventListener('change', () => { if (getLayer(S.selectedId)) commit(); });

    d.zoomIn.addEventListener('click', () => zoomTo(S.view.scale * 1.25));
    d.zoomOut.addEventListener('click', () => zoomTo(S.view.scale / 1.25));
    d.fit.addEventListener('click', fitView);

    d.clear.addEventListener('click', clearAll);

    // PNG | JPG | Template switch (same look and behaviour as the XLSX | CSV one)
    const fmt = $('designFormat');
    const unit = $('targetUnit');
    const targetEl = $('targetSize');
    const scaleGroup = $('designScale');
    const setFmt = (v) => {
      fmt.dataset.value = v;
      fmt.querySelectorAll('button[data-value]').forEach((b) => {
        const on = b.dataset.value === v;
        b.classList.toggle('active', on); b.setAttribute('aria-checked', String(on)); b.tabIndex = on ? 0 : -1;
      });
      syncTargetUi();
    };
    // A target size can only be met by lowering JPEG quality, so setting one selects JPG
    // (and PNG is switched off until the target is cleared). Template files have no size target.
    function syncTargetUi() {
      const has = !!readTarget().bytes;
      const pngBtn = fmt.querySelector('button[data-value="png"]');
      pngBtn.classList.toggle('is-disabled', has);
      pngBtn.setAttribute('aria-disabled', String(has));
      pngBtn.title = has ? 'PNG can’t be squeezed to a size. Clear the target size to use PNG.' : 'Lossless image';
      const off = fmt.dataset.value === 'template';
      targetEl.disabled = off;
      $('targetGroup').classList.toggle('is-off', off);
      $('targetNote').textContent = has && !off ? '→ saved as JPG' : '';
      // The scale multiplier only means anything for a raster export.
      scaleGroup.classList.toggle('is-off', off);
      scaleGroup.querySelectorAll('button').forEach((b) => { b.disabled = off; });
    }
    fmt.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-value]');
      if (!b) return;
      if (b.dataset.value === 'png' && readTarget().bytes) { showToast('PNG can’t be squeezed to a size — clear the target size first, or use JPG', 3500); return; }
      setFmt(b.dataset.value);
    });
    fmt.addEventListener('keydown', (e) => {
      if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
      e.preventDefault();
      const order = ['png', 'jpg', 'template'].filter((v) => !(v === 'png' && readTarget().bytes));
      const i = order.indexOf(fmt.dataset.value);
      const step = e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 1;
      const next = order[(i + step + order.length) % order.length];
      setFmt(next);
      fmt.querySelector(`button[data-value="${next}"]`).focus();
    });
    scaleGroup.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-value]');
      if (!b || b.disabled) return;
      scaleGroup.dataset.value = b.dataset.value;
      scaleGroup.querySelectorAll('button[data-value]').forEach((x) => {
        const on = x === b;
        x.classList.toggle('active', on); x.setAttribute('aria-checked', String(on)); x.tabIndex = on ? 0 : -1;
      });
    });
    scaleGroup.addEventListener('keydown', (e) => {
      if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(e.key)) return;
      e.preventDefault();
      const order = ['1', '2', '3'];
      const i = order.indexOf(scaleGroup.dataset.value);
      const step = e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 1;
      const next = order[(i + step + order.length) % order.length];
      scaleGroup.querySelector(`button[data-value="${next}"]`).click();
      scaleGroup.querySelector(`button[data-value="${next}"]`).focus();
    });
    unit.addEventListener('click', (e) => {
      const b = e.target.closest('button[data-value]');
      if (!b || targetEl.disabled) return;
      unit.dataset.value = b.dataset.value;
      unit.querySelectorAll('button[data-value]').forEach((x) => {
        const on = x === b;
        x.classList.toggle('active', on); x.setAttribute('aria-checked', String(on)); x.tabIndex = on ? 0 : -1;
      });
      syncTargetUi();
    });
    targetEl.addEventListener('input', () => {
      if (readTarget().bytes && fmt.dataset.value !== 'jpg') setFmt('jpg'); else syncTargetUi();
    });
    targetEl.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); exportImage(); } });
    syncTargetUi();
  }

  function ensureStage() {
    if (S.stage) return;
    const d = S.dom;
    Konva.pixelRatio = Math.min(window.devicePixelRatio || 1, 2);
    S.stage = new Konva.Stage({ container: d.viewport, width: Math.max(50, d.viewport.clientWidth), height: Math.max(50, d.viewport.clientHeight) });
    S.bgLayer = new Konva.Layer();
    S.contentLayer = new Konva.Layer();
    S.uiLayer = new Konva.Layer();
    S.stage.add(S.bgLayer, S.contentLayer, S.uiLayer);

    S.artboard = new Konva.Rect({ id: 'canvas', x: 0, y: 0, width: M.canvas.w, height: M.canvas.h, fill: M.canvas.fill, stroke: '#c4ccd6', strokeWidth: 1, strokeScaleEnabled: false });
    S.bgLayer.add(S.artboard);
    S.content = new Konva.Group({ clipFunc: (ctx) => ctx.rect(0, 0, M.canvas.w, M.canvas.h) });
    S.contentLayer.add(S.content);

    S.transformer = new Konva.Transformer({
      borderStroke: ACCENT, anchorStroke: ACCENT, anchorFill: '#ffffff', anchorSize: 10, anchorCornerRadius: 2,
      borderStrokeWidth: 1.5, anchorStrokeWidth: 1.5, rotateAnchorOffset: 28, padding: 0,
      rotationSnaps: [0, 90, 180, 270], rotationSnapTolerance: 5, ignoreStroke: true, flipEnabled: false,
      boundBoxFunc: snapResize,
    });
    S.uiLayer.add(S.transformer);
    const guide = () => new Konva.Line({ points: [0, 0, 0, 0], stroke: GUIDE_COLOR, strokeWidth: 1.5, strokeScaleEnabled: false, listening: false, visible: false });
    S.guides = { v: guide(), h: guide() };
    S.uiLayer.add(S.guides.v, S.guides.h);
    wireArtboard();
    wireViewport();

    S.resizeObserver = new ResizeObserver(() => {
      if (!S.active) return;
      if (S.view.auto) fitView(); else { S.stage.size({ width: Math.max(50, d.viewport.clientWidth), height: Math.max(50, d.viewport.clientHeight) }); S.stage.batchDraw(); }
      positionTextEditor();
    });
    S.resizeObserver.observe(d.viewport);
  }

  const nextFrame = () => new Promise((r) => requestAnimationFrame(() => r()));

  function resetModel(w, h) {
    for (const a of S.assets.values()) { try { URL.revokeObjectURL(a.url); } catch (e) { /* ignore */ } }
    S.assets.clear();
    M.canvas = { w, h, fill: '#ffffff' };
    M.layers = [];
    S.selectedId = null;
    S.history = [];
    S.hIndex = -1;
    S.view = { scale: 1, auto: true };
    S.cropping = null;
    S.editing = null;
    S.content.destroyChildren();
    S.nodes.clear();
  }

  // opts: { w, h, title, imageFile } or { templateText, title }
  async function enter(opts = {}) {
    const d = ensureDom();
    let asset = null;
    let tpl = null;
    if (opts.imageFile) asset = await createRasterAsset(opts.imageFile); // throws if it isn't a readable image — before anything on screen changes
    if (opts.templateText) tpl = await prepareTemplate(opts.templateText); // likewise for a damaged design file

    state.mode = 'design';
    document.body.classList.add('mode-design');
    document.body.classList.remove('mode-home');
    await nextFrame();
    ensureStage();
    S.active = true;
    closeMenu();

    let w = opts.w || DEFAULT_W;
    let h = opts.h || DEFAULT_H;
    let clampedNote = '';
    if (asset) {
      w = asset.w; h = asset.h;
      const big = Math.max(w, h);
      if (big > MAX_DIM) {
        const k = MAX_DIM / big;
        w = Math.round(w * k); h = Math.round(h * k);
        clampedNote = ` (scaled to ${w} × ${h}, the largest canvas)`;
      }
    }
    if (tpl) { w = tpl.canvas.w; h = tpl.canvas.h; }
    // Resuming the design scratchpad (opts.resume): a stash from the last time it was left
    // (see leave()) is the authoritative "last live state" — more current than opts/tpl,
    // which after a Clear on the home-screen tile is blank because the saved record was
    // deleted. When a stash exists, its current entry replaces whatever opts would otherwise
    // have built, so what ends up on screen always matches S.history[S.hIndex] — otherwise
    // Undo can restore a snapshot that doesn't match what was actually just drawn.
    const resuming = !!(opts.resume && S.scratchHistory && S.scratchHistory.length);
    let resumedCanvas = null, resumedLayers = null;
    if (resuming) {
      const snap = JSON.parse(S.scratchHistory[S.scratchHIndex]);
      resumedCanvas = snap.canvas; resumedLayers = snap.layers;
      w = resumedCanvas.w; h = resumedCanvas.h;
    }
    // resetModel() clears assets, so the ones decoded above (or stashed by leave(), for a
    // resume) are re-registered right after it.
    resetModel(w, h);
    if (resuming && S.scratchAssets) { for (const [id, a] of S.scratchAssets) S.assets.set(id, a); }
    if (asset) S.assets.set(asset.id, asset);
    if (resuming) {
      M.canvas = resumedCanvas;
      M.layers = resumedLayers;
    } else if (tpl) {
      M.canvas = tpl.canvas;
      M.layers = tpl.layers;
      for (const a of tpl.assets.values()) S.assets.set(a.id, a);
    }
    state.title = opts.title || tpl?.title || (asset ? (opts.imageFile.name || 'image').replace(/\.[^.]+$/, '') : 'Untitled design');
    state.originalFileName = '';
    state.dirty = false;
    el.fileName.textContent = state.title;
    el.renameTitleBtn.hidden = false;
    el.saveBtn.disabled = false;

    if (asset && !resuming) {
      M.layers.push({
        id: uid('l'), type: 'image', name: state.title, asset: asset.id, crop: { x: 0, y: 0, w: 1, h: 1 },
        x: 0, y: 0, width: w, height: h, rotation: 0, visible: true, locked: false,
      });
    }
    applyCanvasToStage(); // sizes the artboard, fills the W/H boxes and fits it in the window
    for (const l of M.layers) addNodeFor(l);
    if (resuming) {
      S.history = S.scratchHistory;
      S.hIndex = S.scratchHIndex;
    } else {
      S.history = [snapshot()];
      S.hIndex = 0;
    }
    S.scratchHistory = null;
    S.scratchHIndex = -1;
    S.scratchAssets = null;
    S.boundToScratch = !!opts.resume;
    fitView();
    updateTransformer();
    renderLayers();
    syncToolbar();
    S.contentLayer.batchDraw();
    S.bgLayer.batchDraw();
    d.zoomLabel.textContent = `${Math.round(S.view.scale * 100)}%`;
    if (tpl) {
      const notes = [];
      if (tpl.dropped) notes.push(`${tpl.dropped} damaged layer${tpl.dropped === 1 ? '' : 's'} skipped`);
      if (tpl.missingFonts.length) notes.push(`font${tpl.missingFonts.length === 1 ? '' : 's'} not loaded: ${tpl.missingFonts.join(', ')}`);
      showToast(`Opened template “${state.title}” — ${M.layers.length} layer${M.layers.length === 1 ? '' : 's'}, canvas ${w} × ${h}${notes.length ? ' · ' + notes.join(' · ') : ''}`, notes.length ? 6000 : 3200);
    } else {
      showToast(asset ? `Opened “${state.title}” — canvas ${w} × ${h}${clampedNote}` : `New template — ${w} × ${h}`, 3200);
    }
    return { width: w, height: h };
  }

  function leave() {
    if (!S.active) return;
    if (S.editing) { S.editing.ta.remove(); S.editing = null; }
    if (S.cropping) { S.cropping.group.destroy(); S.cropping = null; document.body.classList.remove('design-cropping'); }
    closeMenu();
    S.active = false;
    document.body.classList.remove('mode-design');
    if (S.content) S.content.destroyChildren();
    S.nodes.clear();
    if (S.transformer) S.transformer.nodes([]);
    // Stash undo/redo — and the asset blobs those layers reference — for a scratchpad
    // session so reopening it (enter with opts.resume) can restore it fully; a non-scratch
    // session (an opened file/image/template) just discards everything as before.
    if (S.boundToScratch) {
      S.scratchHistory = S.history; S.scratchHIndex = S.hIndex;
      S.scratchAssets = S.assets; S.assets = new Map(); // hand the blobs off intact (not revoked below)
    } else {
      S.scratchHistory = null; S.scratchHIndex = -1; S.scratchAssets = null;
    }
    for (const a of S.assets.values()) { try { URL.revokeObjectURL(a.url); } catch (e) { /* ignore */ } }
    S.assets.clear();
    M.layers = [];
    S.history = [];
    S.hIndex = -1;
    S.selectedId = null;
    if (state.mode === 'design') state.mode = 'home';
  }

  window.Design = {
    enter, leave, undo, redo, exportImage, clearAll,
    isActive: () => S.active,
    DEFAULT_W, DEFAULT_H,
    // for tests / debugging
    _S: S, _M: M, _addImageFile: addImageFile, _addText: addText, _addShape: addShape, _select: select,
    _readTarget: readTarget, _serializeTemplate: serializeTemplate, _encodeForTarget: encodeForTarget, _renderFullCanvas: renderFullCanvas,
    _handleFiles: handleFiles, _startCrop: startCrop, _endCrop: endCrop, _centerSelected: centerSelected,
  };
})();
