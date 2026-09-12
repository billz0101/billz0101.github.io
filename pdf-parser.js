/**
 * PdfExtractor — Pallet Manifest PDF Parser
 *
 * Reads the PDF TEXT LAYER in reading order and reconstructs the
 * PICKING LIST rows from CONTENT (labels, batch/expiry text, dates,
 * quantities) instead of from fixed pixel positions. Rows that wrap
 * across several physical lines are merged back into a single item,
 * so the parser also works when the pack-note template differs.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.PdfExtractor = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {

  const COLS = ['stockCode', 'productName', 'specifics', 'location', 'qty', 'tick'];

  const HEADER_HINTS = {
    stockCode:   ['stock', 'sku', 'item code', 'item#', 'code'],
    productName: ['product', 'descri', 'name', 'description'],
    specifics:   ['spec', 'batch', 'b/n', 'lot', 'expiry', 'mfg'],
    location:    ['locat', 'bin', 'bay', 'aisle', 'rack'],
    qty:         ['qty', 'quant', 'units', 'ordered', 'pick'],
    tick:        ['tick', 'pallet', 'chk', 'check', 'done']
  };

  const SKIP_STOCK = /^(order|orders|comments|date|courier|part\b|status|weight|number|contact|requested|delivery|request\b|items|total|subtotal|sub\s*total|page|null)/i;

  let PDFJS = null;

  function getPdfjs() {
    if (PDFJS) return PDFJS;
    const rootObj = (typeof window !== 'undefined' ? window : globalThis);
    if (rootObj.pdfjsLib && rootObj.pdfjsLib.getDocument) PDFJS = rootObj.pdfjsLib;
    return PDFJS;
  }

  // ---------------- public API ----------------

  async function parsePdfBuffer(arrayBuffer) {
    const pdfjs = getPdfjs();
    if (!pdfjs) throw new Error('PDF.js library is not available');

    if (typeof window !== 'undefined' && window.document) {
      pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }
    const inNode = typeof window === 'undefined' || typeof document === 'undefined';
    const loadingTask = pdfjs.getDocument({ data: arrayBuffer, disableWorker: inNode });
    const pdfDoc = await loadingTask.promise;

    const rows = [];
    let carriedColumns = null;
    let carriedPitch = null;
    let prevRow = null;
    const meta = { orderNumber: '', orderDate: '' };

    for (let p = 1; p <= pdfDoc.numPages; p++) {
      const result = await parsePage(pdfDoc, p, carriedColumns, carriedPitch, meta);

      // A row split across a page boundary leaves a leading fragment at the
      // top of the next page (wrapped text with NO stock code and NO qty,
      // e.g. "Packs / Exp:17/08/2028"). Merge it back into the previous
      // page's last row instead of emitting it as a phantom row or dropping
      // it (which would lose the wrapped text). Any other stray fragment is
      // dropped.
      let pageRows = result.rows || [];
      const merged = [];
      for (const row of pageRows) {
        if (row.fragment) {
          if (merged.length === 0 && prevRow) appendFragment(prevRow, row);
          continue;
        }
        merged.push(row);
      }

      rows.push(...merged);
      if (merged.length) prevRow = merged[merged.length - 1];
      if (result.columns && result.columns.length >= 3) carriedColumns = result.columns;
      if (result.pitch) carriedPitch = result.pitch;
    }
    rows.meta = meta;
    return rows;
  }

  function appendFragment(target, frag) {
    if (frag.productName) target.productName = ((target.productName || '') + ' ' + frag.productName).trim();
    if (frag.specifics) target.specifics = ((target.specifics || '') + ' ' + frag.specifics).trim();
    if (frag.location && !target.location) target.location = frag.location;
  }

  /**
   * Detect the order number and order date from the page's text lines. Runs
   * on every page but only fills fields that are still empty (the order
   * details block repeats on each page). The detection is restricted to the
   * order-details region (between the "Order Details" heading and the
   * "Picking List" heading) so body-table dates and stock codes never
   * confuse it. Because pack note cells stack their text on staggered
   * baselines (labels and values can end up on different physical lines),
   * detection is positional, not same-line:
   *   - order number: the first code-shaped token (letters + hyphen + digits)
   *     anywhere in the region, e.g. Goods-2849, ORD-2026-9941.
   *   - order date: the leftmost date token on any line within a few points
   *     vertically of the standalone "Date" label cell (the requested
   *     delivvery-date value always sits further right).
   */
  function detectOrderMeta(lines, meta) {
    if (meta.orderNumber && meta.orderDate) return;

    // Restrict to the order-details block if it exists on this page.
    let start = -1, end = lines.length;
    for (let i = 0; i < lines.length; i++) {
      const t = lineText(lines[i]);
      if (start < 0) {
        if (/^order\s+details$/i.test(t)) start = i + 1;
      } else if (/picking\s+list/i.test(t)) {
        end = i;
        break;
      }
    }
    if (start < 0) start = 0;
    const region = lines.slice(start, end);
    if (!region.length) return;

    if (!meta.orderNumber) {
      // 1) Full value cell beside the "Order Number" label (customer + ref,
      //    e.g. "Ashford Supermart Goods-2849"). The cell text is stacked on
      //    staggered baselines, so collect every token that sits between this
      //    label's column and the next grid column, spanning this row's
      //    vertical extent.
      let labelItem = null;
      for (const line of region) {
        for (const it of line.items) {
          if (/^order\s+number$/i.test(it.s) && (!labelItem || line.y < labelItem.y)) {
            labelItem = { x: it.x, y: line.y };
          }
        }
      }
      if (labelItem) {
        const cols = [...new Set(region.flatMap(l => l.items.map(i => Math.round(i.x / 12) * 12)))].sort((a, b) => a - b);
        const labelCol = cols.reduce((p, c) => (Math.abs(c - labelItem.x) < Math.abs(p - labelItem.x) ? c : p), cols[0]);
        const ci = cols.indexOf(labelCol);
        const valCol = cols[ci + 1];
        if (valCol !== undefined) {
          const bandL = (labelCol + valCol) / 2;
          const bandR = ci + 2 < cols.length ? (valCol + cols[ci + 2]) / 2 : valCol + 220;
          let rowEndY = Infinity;
          for (const line of region) {
            if (line.y <= labelItem.y + 2) continue;
            const sameCol = line.items.find(it => Math.abs(it.x - labelItem.x) <= 10 && !/^\d/.test(it.s));
            if (sameCol) { rowEndY = line.y; break; }
          }
          const toks = [];
          for (const line of region) {
            if (line.y < labelItem.y - 16 || line.y >= rowEndY) continue;
            for (const it of line.items) {
              if (it.x > bandL && it.x < bandR) toks.push({ y: line.y, x: it.x, s: it.s });
            }
          }
          toks.sort((a, b) => a.y - b.y || a.x - b.x);
          const text = toks.map(t => t.s).join(' ').replace(/\s+/g, ' ').trim();
          if (text) meta.orderNumber = text;
        }
      }
    }
    if (!meta.orderNumber) {
      const re = /\b[A-Za-z][A-Za-z0-9]*(?:-[A-Za-z0-9]+)+\d{2,}[A-Za-z0-9\-_.\/]*/;
      for (const line of region) {
        const m = lineText(line).match(re);
        if (m) { meta.orderNumber = m[0].trim(); break; }
      }
    }

    if (!meta.orderDate) {
      const dateRe = /\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?/;
      let labelX = null, labelY = 0;
      for (const line of region) {
        for (const it of line.items) {
          if (/^date$/i.test(it.s) && (labelX === null || it.x < labelX)) { labelX = it.x; labelY = line.y; }
        }
      }
      if (labelX !== null) {
        let bestX = Infinity;
        for (const line of region) {
          if (Math.abs(line.y - labelY) > 14) continue;
          for (const it of line.items) {
            const m = it.s.match(dateRe);
            if (m && it.x < bestX) { bestX = it.x; meta.orderDate = m[0]; }
          }
        }
      } else {
        for (const line of region) {
          const t = lineText(line);
          if (/\bexpir|best\s*before|use\s*by|mfg\b|requested\s*delivery|delivery\s*date/i.test(t)) continue;
          if (!/\bdate\b/i.test(t)) continue;
          const m = t.match(/\bdate\b[^\d]{0,40}?(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?/i);
          if (m) { meta.orderDate = (m[1] + (m[2] ? ' ' + m[2] : '')).trim(); break; }
        }
      }
    }
  }

  function extractBatchExpiry(specifics) {
    const s = (specifics || '').trim();
    let batch = '';
    let expiry = '';

    const bm = s.match(/(?:batch\s*(?:no\.?|number|#)?|b\/n|b#|lot\s*(?:no\.?)?)\s*[:#]?\s*([A-Za-z0-9][A-Za-z0-9_.\-]*)/i);
    if (bm) batch = bm[1];

    const dateTok = /(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}|\d{1,2}[\/\-.]\d{2,4})/;
    const dm = s.match(/(?:exp(?:iry)?|bbe|best\s*before|use\s*by)\s*[:#]?\s*(?:N\/?A\b\s*|(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})|(\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2})|(\d{1,2}[\/\-.]\d{2,4}))/i);
    if (dm) expiry = dm[1] || dm[2] || dm[3] || '';

    if (!expiry) {
      const any = s.match(dateTok);
      if (any && !/^[0-9]{1,2}[\/\-.]$/.test(any[0])) expiry = any[1];
    }
    return { batch, expiry };
  }

  // ---------------- page parsing ----------------

  async function parsePage(pdfDoc, pageNum, fallbackColumns, pitchHint, meta) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();

    const items = content.items.map(it => {
      const tx = PDFJS.Util.transform(viewport.transform, it.transform);
      return { x: tx[4], y: tx[5], w: it.width || 0, s: (it.str || '').trim() };
    }).filter(it => it.s.length > 0);

    if (!items.length) return { rows: [], columns: null };

    items.sort((a, b) => a.y - b.y || a.x - b.x);
    const lines = groupLines(items);
    if (meta) detectOrderMeta(lines, meta);

    // 1) Try to find the picking-list header anywhere on the page (by content).
    let headerIdx = -1;
    let columns = null;
    for (let i = 0; i < lines.length; i++) {
      if (looksLikeHeader(lineText(lines[i]))) {
        headerIdx = i;
        columns = detectColumns(lines[i].items);
        break;
      }
    }

    let tableLines;
    if (headerIdx >= 0) {
      tableLines = lines.slice(headerIdx + 1);
    } else {
      // No header: fall back to the previous page's columns, or slice the
      // region after the "PICKING LIST" heading.
      columns = (fallbackColumns && fallbackColumns.length >= 3) ? fallbackColumns : null;
      let startIdx = 0;
      for (let i = 0; i < lines.length; i++) {
        if (/picking\s*list/i.test(lineText(lines[i]))) { startIdx = i + 1; break; }
      }
      tableLines = lines.slice(startIdx);
      if (startIdx === 0 && columns) {
        // no heading found either; use whole page minus obvious chrome
      }
    }

    // 2) Trim the candidate region to the actual table.
    const candidates = [];
    for (const line of tableLines) {
      const txt = lineText(line).toLowerCase();
      if (!txt) continue;
      if (/^thanks\s+for\s+your\s+money/.test(txt)) break;
      if (/^picking\s+list\s*$/.test(txt)) continue;
      if (/^\d+\s+of\s+\d+$/.test(txt) || /^page\s+\d/.test(txt)) break;
      if (/^(order\s+details|comments)\s*$/i.test(txt)) break;
      if (/^file:\/\//.test(txt) || /^chrome:\/\//.test(txt)) continue;
      candidates.push(line);
    }
    if (!candidates.length) return { rows: [], columns: columns };

    // 3) If the header still can't give usable columns, derive them straight
    //    from the data rows (template-independent, content-first).
    if (!columns || columns.length < 3) {
      columns = deriveColumnsFromRows(candidates);
      if (!columns) return { rows: [], columns: null };
    }

    // Line pitch: the vertical gap between table ROWS. Wrapped continuation
    // lines sit ~half that apart, so the largest gap on a page is the row
    // pitch (a median can land on the wrap pitch when wraps outnumber rows).
    // Computed on pages with enough lines, otherwise carried over from
    // earlier pages so rows keep clustering correctly on sparse pages (e.g.
    // a continuation page holding a single row).
    let pitch = pitchHint || null;
    const gaps = [];
    for (let i = 1; i < candidates.length; i++) gaps.push(candidates[i].y - candidates[i - 1].y);
    if (gaps.length >= 2) pitch = Math.max(...gaps);

    // 4) Merge wrapped physical lines into logical rows using the document's
    //    own line pitch (adaptive, per document).
    const rowClusters = clusterRows(candidates, columns, pitch);

    // 5) Build rows and validate by content.
    const rows = [];
    for (const cluster of rowClusters) {
      const cells = buildRow(cluster, columns);
      const stock = cleanupCode(cells.stockCode);
      const product = (cells.productName || '').trim();
      const specs = (cells.specifics || '').trim();
      const loc = (cells.location || '').trim();
      const qtyNum = parseQty(cells.qty);

      if (!stock && !product && !specs && !loc && qtyNum == null) continue;
      if (SKIP_STOCK.test(stock)) continue;

      // A cluster with no stock code AND no qty is a row fragment: wrapped
      // tail text of a row that was split across a page boundary. It is
      // surfaced to the caller (flagged) so the continuation can be merged
      // back into the previous page's last row rather than becoming a
      // phantom row or silently losing text.
      const fragment = !stock && qtyNum == null;

      rows.push({
        stockCode: stock,
        productName: product,
        specifics: specs,
        location: loc,
        qty: qtyNum == null ? (cells.qty || '').trim() : String(qtyNum),
        fragment: fragment
      });
    }
    return { rows, columns, pitch };
  }

  // ---------------- structure helpers ----------------

  function groupLines(items) {
    const lines = [];
    for (const it of items) {
      let line = lines.find(l => Math.abs(l.y - it.y) <= 3);
      if (!line) { line = { y: it.y, items: [] }; lines.push(line); }
      line.items.push(it);
    }
    lines.forEach(l => l.items.sort((a, b) => a.x - b.x));
    return lines;
  }

  function lineText(line) {
    return line.items.map(i => i.s).join(' ').replace(/\s+/g, ' ').trim();
  }

  function gapClusters(items) {
    const sorted = [...items].sort((a, b) => a.x - b.x);
    const clusters = [];
    let cur = null;
    for (const it of sorted) {
      const x1 = it.x + (it.w || 0);
      if (!cur) cur = { x0: it.x, x1: x1, text: it.s };
      else if (it.x - cur.x1 > 12) { clusters.push(cur); cur = { x0: it.x, x1: x1, text: it.s }; }
      else { cur.x1 = Math.max(cur.x1, x1); cur.text += ' ' + it.s; }
    }
    if (cur) clusters.push(cur);
    return clusters;
  }

  function looksLikeHeader(txt) {
    const t = txt.toLowerCase();
    const hasStock = /stock\s*code|stockcode|sku|item\s*code/.test(t);
    const hasQty = /\bqty\b|quantity|\bunit/.test(t);
    const hasSpec = /spec|batch|b\/n|lot|expiry/.test(t);
    const hasName = /product|descri|name/.test(t);
    return hasStock && hasQty && (hasSpec || hasName);
  }

  function detectColumns(headerItems) {
    const clusters = gapClusters(headerItems);
    if (!clusters.length) return null;

    const used = new Set();
    const picked = clusters.map(cl => {
      const lower = cl.text.toLowerCase();
      let label = null;
      for (const key of COLS) {
        if (used.has(key)) continue;
        if (HEADER_HINTS[key].some(k => lower.includes(k))) { label = key; used.add(key); break; }
      }
      return { label, x0: cl.x0, x1: cl.x1 };
    });

    const remaining = COLS.filter(k => !used.has(k));
    for (const p of picked) {
      if (!p.label && remaining.length) p.label = remaining.shift();
    }

    picked.sort((a, b) => (a.x0 + a.x1) / 2 - (b.x0 + b.x1) / 2);
    const cols = picked.filter(p => p.label).map(p => ({ label: p.label, center: (p.x0 + p.x1) / 2 }));
    if (!cols.some(c => c.label === 'stockCode') || !cols.some(c => c.label === 'qty')) return null;
    return cols;
  }

  function assignToColumns(lineItems, columns) {
    const centers = columns.map(c => c.center);
    const bounds = [-Infinity];
    for (let i = 0; i < centers.length - 1; i++) bounds.push((centers[i] + centers[i + 1]) / 2);
    bounds.push(Infinity);

    const cells = {};
    for (const it of lineItems) {
      let idx = columns.length - 1;
      for (let i = 0; i < columns.length; i++) {
        if (it.x >= bounds[i] && it.x < bounds[i + 1]) { idx = i; break; }
      }
      const label = columns[idx].label;
      cells[label] = cells[label] ? cells[label] + ' ' + it.s : it.s;
    }
    return cells;
  }

  function buildRow(cluster, columns) {
    const cells = {};
    for (const line of cluster) {
      const lineCells = assignToColumns(line.items, columns);
      for (const k of COLS) {
        if (lineCells[k]) cells[k] = cells[k] ? cells[k] + ' ' + lineCells[k] : lineCells[k];
      }
    }
    return cells;
  }

  function cleanupCode(raw) {
    return String(raw || '').replace(/\s+/g, '');
  }

  function parseQty(raw) {
    const m = String(raw || '').match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  }

  function cellQty(line, columns) {
    const cells = assignToColumns(line.items, columns);
    return parseQty(cells.qty);
  }

  function clusterRows(lines, columns, pitchHint) {
    if (!lines.length) return [];

    const gaps = [];
    for (let i = 1; i < lines.length; i++) gaps.push(lines[i].y - lines[i - 1].y);
    // Threshold for "these two lines are separate rows". Use the stricter of
    // the local mean pitch, the local max gap (row pitch) and the row pitch
    // carried from earlier pages, so wrapped lines never get mistaken for new
    // rows on sparse pages. Carried pitch gets the widest margin because it
    // is the only reliable signal there (sparse pages can't estimate it).
    let threshold = 30;
    if (gaps.length) {
      const mean = gaps.reduce((a, b) => a + b, 0) / gaps.length;
      const localMax = Math.max(...gaps);
      threshold = Math.max(mean * 0.9, localMax * 0.6, (pitchHint || 0) * 0.9, 3);
    } else if (pitchHint) {
      threshold = Math.max(pitchHint * 0.9, 3);
    }

    let clusters = [];
    let cur = [lines[0]];
    for (let i = 1; i < lines.length; i++) {
      if (lines[i].y - lines[i - 1].y < threshold) cur.push(lines[i]);
      else { clusters.push(cur); cur = [lines[i]]; }
    }
    clusters.push(cur);

    // Safety net: if a merged cluster contains two physical lines that EACH
    // carry a qty value, they are really two separate rows.
    const out = [];
    for (const cl of clusters) {
      let cur2 = [];
      let seenQty = false;
      for (const line of cl) {
        const hasQty = cellQty(line, columns) != null;
        if (hasQty && seenQty) {
          out.push(cur2);
          cur2 = [];
          seenQty = false;
        }
        cur2.push(line);
        if (hasQty) seenQty = true;
      }
      out.push(cur2);
    }
    return out.filter(c => c.length);
  }

  /**
   * Content-first fallback: when a header row can't be found, build column
   * boundaries from the x-positions that repeat across the data rows, then
   * label each column by the shape of its contents.
   */
  function deriveColumnsFromRows(lines) {
    const usable = lines.filter(l => l.items.length >= 2);
    if (usable.length < 2) return null;

    const buckets = new Map();
    for (const line of usable) {
      for (const t of line.items) {
        const b = Math.round(t.x / 10);
        buckets.set(b, (buckets.get(b) || 0) + 1);
      }
    }
    const minOccur = Math.max(2, Math.floor(usable.length * 0.4));
    let starts = [...buckets.entries()]
      .filter(([, c]) => c >= minOccur)
      .map(([b]) => b * 10)
      .sort((a, b) => a - b);
    if (starts.length < 3) return null;

    const merged = [];
    for (const s of starts) {
      if (merged.length && s - merged[merged.length - 1] < 25) {
        merged[merged.length - 1] = (merged[merged.length - 1] + s) / 2;
      } else merged.push(s);
    }
    starts = merged;
    if (starts.length < 3) return null;

    const bounds = [-Infinity];
    for (let i = 0; i < starts.length - 1; i++) bounds.push((starts[i] + starts[i + 1]) / 2);
    bounds.push(Infinity);

    const cellMap = starts.map(() => []);
    for (const line of usable) {
      for (const t of line.items) {
        let idx = starts.length - 1;
        for (let i = 0; i < starts.length; i++) {
          if (t.x >= bounds[i] && t.x < bounds[i + 1]) { idx = i; break; }
        }
        cellMap[idx].push((t.s || '').trim());
      }
    }

    const nonEmptyCount = i => cellMap[i].filter(v => v).length;

    const labels = new Array(starts.length).fill(null);
    const used = new Set();

    // qty: mostly pure numbers
    let qi = -1, qs = 0;
    for (let i = 0; i < starts.length; i++) {
      const vals = cellMap[i].filter(v => v && /^-?\d+(\.\d+)?$/.test(v));
      const r = nonEmptyCount(i) ? vals.length / nonEmptyCount(i) : 0;
      if (r >= 0.6 && r > qs && vals.length >= 2) { qs = r; qi = i; }
    }
    if (qi >= 0) { labels[qi] = 'qty'; used.add(qi); }

    // specifics: contains batch / expiry / date text
    let si = -1, ss = 0;
    for (let i = 0; i < starts.length; i++) {
      if (used.has(i)) continue;
      const vals = cellMap[i].filter(v => v && (/batch|exp|expiry|b\/n|lot|\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}/i.test(v)));
      const r = nonEmptyCount(i) ? vals.length / nonEmptyCount(i) : 0;
      if (r >= 0.4 && r > ss) { ss = r; si = i; }
    }
    if (si >= 0) { labels[si] = 'specifics'; used.add(si); }

    // location: short alphanumeric codes with a digit + letter. Location
    // columns sit to the RIGHT of the product name, so on a tie prefer the
    // rightmost candidate (the stock-code column also looks like a code).
    const looksLoc = v => {
      if (!v || !/^[A-Za-z0-9][A-Za-z0-9\-:]{1,11}$/.test(v)) return false;
      if (/^-?\d+(\.\d+)?$/.test(v)) return false;
      return /[A-Za-z]/.test(v) && /[0-9]/.test(v);
    };
    let li = -1, ls = 0;
    for (let i = 0; i < starts.length; i++) {
      if (used.has(i)) continue;
      const vals = cellMap[i].filter(v => v && looksLoc(v));
      const r = nonEmptyCount(i) ? vals.length / nonEmptyCount(i) : 0;
      if (r >= 0.7 && (r > ls || (r === ls && r > 0))) { ls = r; li = i; }
    }
    if (li >= 0) { labels[li] = 'location'; used.add(li); }

    // leftover: leftmost = stockCode, then productName
    const order = [];
    for (let i = 0; i < starts.length; i++) if (!used.has(i)) order.push(i);
    if (order.length >= 2) { labels[order[0]] = 'stockCode'; labels[order[1]] = 'productName'; }
    else if (order.length === 1) { labels[order[0]] = 'productName'; }

    const labelIdx = l => starts.findIndex((_, i) => labels[i] === l);
    const iStock = labelIdx('stockCode');
    const iProd = labelIdx('productName');
    const iSpec = labelIdx('specifics');
    const iLoc = labelIdx('location');
    const iQty = labelIdx('qty');
    // Sanity ordering (left→right): stockCode … productName … specifics …
    // location … qty. Heavily scrambled layouts return null rather than
    // emitting badly-labelled rows.
    const hasOrder = (a, b) => a >= 0 && b >= 0 && a < b;
    if (!hasOrder(iStock, iProd) && !hasOrder(iStock, iSpec)) return null;
    if (iSpec >= 0 && !hasOrder(iProd < 0 ? iStock : iProd, iSpec)) return null;
    if (iLoc >= 0 && !hasOrder(iProd < 0 ? iStock : iProd, iLoc)) return null;
    if (iQty >= 0 && !hasOrder(Math.max(iSpec, iLoc, iProd, iStock), iQty)) return null;

    const cols = [];
    for (let i = 0; i < starts.length; i++) {
      if (labels[i]) cols.push({ label: labels[i], center: starts[i] });
    }
    if (!cols.some(c => c.label === 'stockCode') || !cols.some(c => c.label === 'qty')) return null;
    return cols;
  }

  return {
    parsePdfBuffer: parsePdfBuffer,
    extractBatchExpiry: extractBatchExpiry
  };
});