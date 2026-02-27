/**
 * app.js — Autori Convertor
 *
 * Fully browser-based (no server) batch PDF processor.
 * Reads Spanish tax-authorisation PDFs, applies specific text / date
 * replacements, and offers the modified PDFs for download.
 *
 * Dependencies (loaded as global UMD scripts in index.html):
 *   • PDF.js  → window.pdfjsLib
 *   • pdf-lib → window.PDFLib
 *
 * URL flags:
 *   ?dev  — enable verbose console logging + debug UI panel + canvas overlay
 */

'use strict';

// ═══════════════════════════════════════════════════════════════════════════
// 1.  CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Exact string that MUST appear in the source PDF (after space normalisation).
 * NOTE: The source PDFs use the double surname "MENSHUTKINA MENSHUTKINA" as
 * rendered by the document editor — this is the literal text in the content
 * stream and must be matched exactly.
 */
const REQUIRED_STR_1 = 'A NATALIA MENSHUTKINA MENSHUTKINA CON DNI 79235585X';
/** Replacement for REQUIRED_STR_1. */
const REPLACE_STR_1  = 'A NALOGI ASESORES S.L.U. CIF B25973512';

/** Exact string that MUST appear in the source PDF (after space normalisation). */
const REQUIRED_STR_2 = 'A SU AUTORIZADO RED 241384 INCLUIDA';
/** Replacement for REQUIRED_STR_2. */
const REPLACE_STR_2  = 'RED 354901, INCLUIDA';

/** Maximum files accepted per batch. */
const MAX_FILES = 50;

/** Minimum / maximum delay (ms) between successive browser downloads. */
const DOWNLOAD_DELAY_MIN = 400;
const DOWNLOAD_DELAY_MAX = 800;

/**
 * How long (ms) to keep a blob URL alive after triggering a download.
 * A generous value avoids revocation before the browser fetches the blob,
 * which can happen on slow connections or with large PDFs.
 */
const URL_REVOKE_DELAY_MS = 60_000;

/** Enable verbose dev logging + debug UI when URL contains ?dev */
const DEV_MODE = new URLSearchParams(window.location.search).has('dev');

/** Accumulates debug info for the "Copy debug report" button. */
let _debugReport = [];

// ═══════════════════════════════════════════════════════════════════════════
// 2.  LOGGING HELPER
// ═══════════════════════════════════════════════════════════════════════════

/** Log only when ?dev is present in the URL. */
const devLog = (...args) => {
  if (DEV_MODE) console.log('[DEV]', ...args);
};

/** Append a line to the debug report buffer (used by dev panel). */
const devReport = (line) => {
  if (DEV_MODE) _debugReport.push(line);
};

// ═══════════════════════════════════════════════════════════════════════════
// 3.  UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Format a Date object as DD/MM/YYYY.
 * @param {Date} date
 * @returns {string}
 */
function formatDateDDMMYYYY(date) {
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = String(date.getFullYear());
  return `${d}/${m}/${y}`;
}

/**
 * Sanitise a string for use as a filename.
 * Replaces  \ / : * ? " < > |  with underscores, trims, and caps at 120 chars.
 * @param {string} name
 * @returns {string}
 */
function sanitizeFileName(name) {
  return name
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim()
    .substring(0, 120);
}

/**
 * Strip invisible / zero-width / control characters that PDF extractors
 * sometimes emit (zero-width spaces, NBSP, soft hyphens, BOM, etc.).
 * @param {string} str
 * @returns {string}
 */
function stripInvisible(str) {
  return str
    // soft hyphen, zero-width space/non-joiner/joiner, word-joiner, BOM
    .replace(/[\u00ad\u200b\u200c\u200d\u2060\ufeff]/g, '')
    // non-breaking space → regular space
    .replace(/\u00a0/g, ' ')
    // line/paragraph separators → space
    .replace(/[\u2028\u2029]/g, ' ');
}

/**
 * Collapse all whitespace sequences to a single space and trim.
 * Also strips invisible characters first.
 * @param {string} str
 * @returns {string}
 */
function normalizeSpaces(str) {
  return stripInvisible(str).replace(/\s+/g, ' ').trim();
}

/**
 * Escape HTML special characters for safe insertion into innerHTML.
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Return a random integer in [min, max].
 * @param {number} min
 * @param {number} max
 * @returns {number}
 */
function randomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// ═══════════════════════════════════════════════════════════════════════════
// 4.  SELF-CHECKS  (run on page load; assert in console)
// ═══════════════════════════════════════════════════════════════════════════

function runSelfChecks() {
  // formatDateDDMMYYYY
  console.assert(
    formatDateDDMMYYYY(new Date(2026, 1, 27)) === '27/02/2026',
    'Self-check FAIL: formatDateDDMMYYYY basic'
  );
  console.assert(
    formatDateDDMMYYYY(new Date(2024, 11, 5)) === '05/12/2024',
    'Self-check FAIL: formatDateDDMMYYYY zero-padding'
  );
  console.assert(
    formatDateDDMMYYYY(new Date(2000, 0, 1)) === '01/01/2000',
    'Self-check FAIL: formatDateDDMMYYYY Jan'
  );

  // sanitizeFileName
  console.assert(
    sanitizeFileName('BAZHAN ANDRII') === 'BAZHAN ANDRII',
    'Self-check FAIL: sanitizeFileName passthrough'
  );
  console.assert(
    sanitizeFileName('NAME/SURNAME') === 'NAME_SURNAME',
    'Self-check FAIL: sanitizeFileName slash'
  );
  console.assert(
    sanitizeFileName('A\\B:C*D?E"F<G>H|I') === 'A_B_C_D_E_F_G_H_I',
    'Self-check FAIL: sanitizeFileName all special chars'
  );
  console.assert(
    sanitizeFileName('x'.repeat(200)).length === 120,
    'Self-check FAIL: sanitizeFileName max length'
  );

  // stripInvisible
  console.assert(
    stripInvisible('A\u200bB\u00adC\u00a0D') === 'AB C D',
    'Self-check FAIL: stripInvisible'
  );

  // normalizeSpaces with invisible chars
  console.assert(
    normalizeSpaces('A\u00a0 B\u200b  C') === 'A B C',
    'Self-check FAIL: normalizeSpaces invisible'
  );

  devLog('Self-checks passed ✓');
}

// ═══════════════════════════════════════════════════════════════════════════
// 5.  TEXT PARSING
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Extract Name + Surname from a line matching "YO, NAME SURNAME, DNI/NIE,".
 * @param {string} text  Full PDF text (all pages concatenated).
 * @returns {{ name: string|null, doc: string|null }}
 */
function parseName(text) {
  const norm = normalizeSpaces(text);
  // DNI/NIE format: starts with alphanumeric [A-Z0-9], followed by 7-8 digits,
  // optionally ending with alphanumeric [A-Z0-9]? (e.g. "X1234567A", "12345678Z").
  // Separator between name and DNI can be comma+space or just whitespace.
  const match = norm.match(/\bYO,\s+(.+?)(?:,\s*|\s+)([A-Z0-9]\d{7,8}[A-Z0-9]?)\b/);
  if (!match) return { name: null, doc: null };
  return { name: match[1].trim(), doc: match[2].trim() };
}

/**
 * Extract IAE code from text.
 * @param {string} text
 * @returns {string|null}
 */
function parseIAE(text) {
  const match = text.match(/\bIAE[:\s]+[A-Z]?(\d{1,4}(?:\.\d{1,2})?)\b/);
  return match ? match[1] : null;
}

/**
 * Extract CNAE code from text (4-digit primary, decimal fallback).
 * @param {string} text
 * @returns {string|null}
 */
function parseCNAE(text) {
  const primary = text.match(/\bCNAE[:\s]+(\d{4})\b/);
  if (primary) return primary[1];
  const fallback = text.match(/\bCNAE[:\s]+(\d{4}(?:\.\d+)?)\b/);
  return fallback ? fallback[1] : null;
}

/**
 * Build a "full page text" from reconstructed lines by joining them with a
 * single space (for cross-line search).  Lines are passed in reading order.
 * @param {string[]} lineTexts
 * @returns {string}
 */
function buildFullPageText(lineTexts) {
  return normalizeSpaces(lineTexts.join(' '));
}

/**
 * Check whether both required strings are present anywhere in the provided
 * line texts (after space normalisation — exact, case-sensitive match).
 *
 * Strategy:
 *  1. Check each individual line (handles same-line occurrence).
 *  2. Check sliding windows of 2–4 consecutive lines joined with a space
 *     (handles cross-line / line-wrapped occurrences).
 *  3. Check full concatenation of the page (last resort).
 *
 * @param {string[]} lineTexts  Normalised line strings in reading order.
 * @returns {{ has1: boolean, has2: boolean,
 *             line1: number|null, line2: number|null }}
 *   line1/line2 = 0-based index of the first matching line (or window start).
 */
function checkRequiredStringsInLines(lineTexts) {
  let has1 = false, has2 = false;
  let line1 = null, line2 = null;

  // Pass 1: individual lines
  for (let i = 0; i < lineTexts.length; i++) {
    const t = lineTexts[i];
    if (!has1 && t.includes(REQUIRED_STR_1)) { has1 = true; line1 = i; }
    if (!has2 && t.includes(REQUIRED_STR_2)) { has2 = true; line2 = i; }
    if (has1 && has2) return { has1, has2, line1, line2 };
  }

  // Pass 2: sliding windows of 2–4 lines
  for (let winSize = 2; winSize <= 4; winSize++) {
    for (let i = 0; i <= lineTexts.length - winSize; i++) {
      const window = normalizeSpaces(lineTexts.slice(i, i + winSize).join(' '));
      if (!has1 && window.includes(REQUIRED_STR_1)) { has1 = true; line1 = i; }
      if (!has2 && window.includes(REQUIRED_STR_2)) { has2 = true; line2 = i; }
      if (has1 && has2) return { has1, has2, line1, line2 };
    }
  }

  return { has1, has2, line1, line2 };
}

/**
 * Check whether both required strings are present in the text
 * (after space normalisation — exact, case-sensitive match).
 * @param {string} text
 * @returns {{ has1: boolean, has2: boolean }}
 */
function checkRequiredStrings(text) {
  const norm = normalizeSpaces(text);
  return {
    has1: norm.includes(REQUIRED_STR_1),
    has2: norm.includes(REQUIRED_STR_2),
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// 6.  PDF TEXT EXTRACTION  (PDF.js)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Join line items using position proximity: items that are positionally
 * adjacent (gap < threshold) are concatenated without a space; items with
 * a significant gap get a space inserted.  This fixes date fragments
 * like "07/11" + "/2023" → "07/11/2023".
 *
 * @param {object[]} lineItems  Sorted left-to-right by x.
 * @param {number}   gapTol     Maximum gap (PDF points) to join without space.
 * @returns {string}
 */
function smartJoinLineItems(lineItems, gapTol = 2) {
  if (lineItems.length === 0) return '';
  let text = normalizeSpaces(lineItems[0].str);
  for (let i = 1; i < lineItems.length; i++) {
    const prev = lineItems[i - 1];
    const curr = lineItems[i];
    const prevEnd   = prev.transform[4] + (prev.width || 0);
    const currStart = curr.transform[4];
    const gap = currStart - prevEnd;
    if (gap >= gapTol) text += ' ';
    text += normalizeSpaces(curr.str);
  }
  return normalizeSpaces(text);
}

/**
 * Merge PDF.js text items that are positionally adjacent on the same
 * baseline into single items.  This repairs fragments created by PDF.js
 * when fonts or encodings change mid-word (e.g. dates split into
 * "07/11" + "/2023").
 *
 * @param {object[]} items     PDF.js TextContent items (with transform, width, height).
 * @param {number}   yTol      Baseline y tolerance (PDF points).
 * @param {number}   gapTol    Max horizontal gap to consider "adjacent" (PDF points).
 * @returns {object[]}         New array of (possibly merged) items.
 */
function mergeAdjacentItems(items, yTol = 4, gapTol = 2) {
  if (items.length <= 1) return items.map(it => ({ ...it }));

  // Sort top→bottom (large y first), then left→right.
  const sorted = [...items].sort((a, b) => {
    const dy = b.transform[5] - a.transform[5];
    if (Math.abs(dy) > yTol) return dy;
    return a.transform[4] - b.transform[4];
  });

  const merged = [];
  let cur = { ...sorted[0], transform: [...sorted[0].transform] };

  for (let i = 1; i < sorted.length; i++) {
    const item = sorted[i];
    const sameLine = Math.abs(item.transform[5] - cur.transform[5]) <= yTol;
    const curEnd   = cur.transform[4] + (cur.width || 0);
    const gap      = item.transform[4] - curEnd;

    if (sameLine && gap < gapTol) {
      // Merge into current
      const newEnd = item.transform[4] + (item.width || 0);
      cur.str   += item.str;
      cur.width  = newEnd - cur.transform[4];
      cur.height = Math.max(cur.height || 0, item.height || 0);
    } else {
      merged.push(cur);
      cur = { ...item, transform: [...item.transform] };
    }
  }
  merged.push(cur);
  return merged;
}

/**
 * Reconstruct "clean" date strings from fragmented line text by removing
 * interior spaces between the numeric components of a date.
 *
 * E.g. "26 / 02 /202 6" → "26/02/2026"
 *       "0 8/11 /2023"  → "08/11/2023"
 *
 * @param {string} text
 * @returns {string}
 */
function reconstructDates(text) {
  // Remove spaces inside DD / MM / YYYY style dates (with possible spaces
  // around slashes and inside digit groups).
  return text
    // spaces around slashes in DD / MM / YYYY
    .replace(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{2,4})\s*(\d*)/g,
      (_, d, m, y1, y2) => `${d.padStart(2,'0')}/${m.padStart(2,'0')}/${(y1 + y2).padStart(4,'0')}`)
    // ISO dates with spaces
    .replace(/(\d{4})\s*-\s*(\d{2})\s*-\s*(\d{2})/g, '$1-$2-$3');
}

/**
 * Extract text and item positions from all pages of a PDF.
 *
 * @param {ArrayBuffer} arrayBuffer  Raw PDF bytes.
 * @returns {Promise<{
 *   fullText: string,
 *   pageItems: Array<{ pageNum: number, items: object[] }>,
 *   pageLines: Array<Array<{ y: number, text: string, items: object[] }>>
 * }>}
 */
async function extractPDFText(arrayBuffer) {
  // Use a typed-array copy so PDF.js can own its buffer safely.
  const data = new Uint8Array(arrayBuffer.slice(0));
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;

  let fullText = '';
  const pageItems = [];
  const pageLines = []; // per-page array of { y, text, items }

  if (DEV_MODE) {
    devReport(`=== PDF: ${pdf.numPages} page(s) ===`);
  }

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent({ includeMarkedContent: false });
    // Keep only items that have non-empty string content.
    const rawItems = content.items.filter(it => it.str && typeof it.str === 'string' && it.str.trim().length > 0);

    // Merge positionally adjacent items to fix text fragmentation
    // (dates, split words, etc.).
    const items = mergeAdjacentItems(rawItems);
    pageItems.push({ pageNum, items });

    // Build reconstructed lines with y, text, and source items.
    const lineGroups = groupItemsByLine(items);
    const lines = lineGroups.map((lineItems, idx) => ({
      lineIndex: idx,
      y: lineItems[0] ? lineItems[0].transform[5] : 0,
      text: normalizeSpaces(reconstructDates(smartJoinLineItems(lineItems))),
      items: lineItems,
    }));
    pageLines.push(lines);

    // Build fullText using reconstructed line texts.
    fullText += lines.map(l => l.text).join(' ') + '\n';

    if (DEV_MODE) {
      devReport(`--- Page ${pageNum}: ${rawItems.length} rawItems, ${items.length} merged, ${lines.length} lines ---`);
      // Top-50 longest items
      const top50 = [...items]
        .sort((a, b) => (b.str || '').length - (a.str || '').length)
        .slice(0, 50);
      devReport(`  Top-${Math.min(50, top50.length)} items by length:`);
      top50.forEach(it => {
        devReport(`    [x=${Math.round(it.transform[4])},y=${Math.round(it.transform[5])}] "${normalizeSpaces(it.str)}"`);
      });
      // Reconstructed lines dump
      devReport(`  Reconstructed lines:`);
      lines.forEach(l => {
        devReport(`    [line=${l.lineIndex},y=${Math.round(l.y)}] "${l.text}"`);
      });
    }
  }

  pdf.destroy();
  return { fullText, pageItems, pageLines };
}

// ═══════════════════════════════════════════════════════════════════════════
// 7.  PDF GENERATION WITH REPLACEMENTS  (pdf-lib)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Group text items into "lines" by proximity of their baseline y-coordinate.
 * Within each line items are sorted left→right by x.
 *
 * @param {object[]} items   PDF.js TextContent items.
 * @param {number}   tol     Baseline-y tolerance in PDF points.
 * @returns {object[][]}     Array of lines; each line is an array of items.
 */
function groupItemsByLine(items, tol = 4) {
  if (items.length === 0) return [];

  // Sort top→bottom  (PDF y=0 is at bottom; larger y = higher on page)
  const sorted = [...items].sort((a, b) => b.transform[5] - a.transform[5]);

  const lines = [];
  let line = [sorted[0]];
  let lineY = sorted[0].transform[5];

  for (let i = 1; i < sorted.length; i++) {
    const item = sorted[i];
    const y    = item.transform[5];
    if (Math.abs(y - lineY) <= tol) {
      line.push(item);
    } else {
      lines.push(line.sort((a, b) => a.transform[4] - b.transform[4]));
      line  = [item];
      lineY = y;
    }
  }
  if (line.length > 0) {
    lines.push(line.sort((a, b) => a.transform[4] - b.transform[4]));
  }

  return lines;
}

/**
 * Given a sorted line of items, find the subset of items whose concatenated
 * text (single-space joined) contains the target string.
 * Returns the contributing items, or null if the target is not found.
 *
 * @param {object[]} lineItems
 * @param {string}   targetStr   Already space-normalised.
 * @returns {object[]|null}
 */
function findItemsForString(lineItems, targetStr) {
  // Build a parallel array: each element knows its source item.
  const charMap = [];
  let first = true;
  for (const item of lineItems) {
    if (!first) charMap.push({ ch: ' ', item: null }); // inter-item space
    for (const ch of item.str) charMap.push({ ch, item });
    first = false;
  }

  // Collapse multiple spaces in the reconstructed line text.
  const lineStr  = charMap.map(c => c.ch).join('');
  const normLine = lineStr.replace(/\s+/g, ' ');
  const normTgt  = normalizeSpaces(targetStr);

  const idx = normLine.indexOf(normTgt);
  if (idx === -1) return null;

  // Collect the distinct items that cover the match range.
  const matched = new Set();
  for (let i = idx; i < idx + normTgt.length; i++) {
    if (i < charMap.length && charMap[i].item !== null) {
      matched.add(charMap[i].item);
    }
  }
  return [...matched];
}

/**
 * Draw a white rectangle over `items` on `pdfPage`, then draw `newText`
 * starting at the leftmost item's x position, at the line's baseline y.
 *
 * @param {object}   pdfPage   pdf-lib PDFPage.
 * @param {object}   font      Embedded pdf-lib font.
 * @param {object[]} items     Source PDF.js items to overwrite.
 * @param {string}   newText   Replacement string.
 */
function applyTextOverlay(pdfPage, font, items, newText) {
  if (!items || items.length === 0) return;

  // Compute bounding box of all contributing items.
  let minX = Infinity, maxX = -Infinity;
  let baseY = items[0].transform[5]; // baseline of first item
  let fontSize = 0;

  for (const item of items) {
    const x = item.transform[4];
    const w = item.width  || 0;
    const h = item.height || Math.abs(item.transform[3]) || 10;

    if (x     < minX) minX = x;
    if (x + w > maxX) maxX = x + w;
    if (h > fontSize) fontSize = h;
  }

  // Fall back to a reasonable default if height data is missing.
  if (fontSize < 1) fontSize = 10;

  // White rectangle: slightly larger than the bounding box.
  const newTextWidth = font.widthOfTextAtSize(newText, fontSize);
  const rectW = Math.max(maxX - minX, newTextWidth) + 4;
  const rectH = fontSize * 1.45;

  pdfPage.drawRectangle({
    x:      minX - 1,
    y:      baseY - fontSize * 0.30,  // a little below baseline for descenders
    width:  rectW,
    height: rectH,
    color:  PDFLib.rgb(1, 1, 1),
    borderWidth: 0,
  });

  // Draw replacement text at the original baseline position.
  pdfPage.drawText(newText, {
    x:     minX,
    y:     baseY,
    size:  fontSize,
    font,
    color: PDFLib.rgb(0, 0, 0),
  });
}

/**
 * Build a modified PDF: load the original with pdf-lib, then for each page
 * overlay replacements (white-box + new text) on:
 *   • All DD/MM/YYYY and YYYY-MM-DD date occurrences → today's date.
 *   • The two required literal strings → their respective replacements.
 *
 * Cross-line strings are handled by scanning windows of 2–4 consecutive
 * lines; the first line in the window receives the overlay in that case.
 *
 * @param {ArrayBuffer} originalArrayBuffer
 * @param {Array<{ pageNum: number, items: object[] }>} pageItems
 * @param {Array<Array<{ y: number, text: string, items: object[] }>>} pageLines
 * @returns {Promise<Uint8Array>}  Bytes of the output PDF.
 */
async function generateOutputPDF(originalArrayBuffer, pageItems, pageLines) {
  const today = formatDateDDMMYYYY(new Date());

  // pdf-lib needs its own copy of the buffer.
  const pdfDoc = await PDFLib.PDFDocument.load(originalArrayBuffer.slice(0));
  const pages  = pdfDoc.getPages();

  // Embed a standard font for replacement text.
  const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  // Regexes for date detection (non-global — used only with .test()).
  const RE_DATE_DDMM = /\b\d{2}\/\d{2}\/\d{4}\b/;
  const RE_DATE_ISO  = /\b\d{4}-\d{2}-\d{2}\b/;

  let dateCount = 0;

  for (let pi = 0; pi < pages.length; pi++) {
    const pdfPage = pages[pi];
    const { items } = pageItems[pi] || { items: [] };
    const lines = pageLines ? pageLines[pi] : null;

    // Use the pre-reconstructed lines if available, otherwise rebuild.
    const lineGroups = groupItemsByLine(items);
    const lineCount = lineGroups.length;

    // ── Track which lines have been handled to avoid double-overlay ───────
    const handled = new Set();

    // ── String replacements: first try individual lines, then windows ─────
    const applyStringReplacement = (targetStr, replaceStr, label) => {
      // Pass 1: individual lines
      for (let li = 0; li < lineCount; li++) {
        const lineText = normalizeSpaces(reconstructDates(smartJoinLineItems(lineGroups[li])));
        if (lineText.includes(targetStr)) {
          const matchItems = findItemsForString(lineGroups[li], targetStr);
          if (matchItems && matchItems.length > 0) {
            devLog(`Page ${pi + 1}: replacing ${label} (line ${li})`);
            applyTextOverlay(pdfPage, font, matchItems, replaceStr);
            handled.add(li);
            return true;
          }
        }
      }
      // Pass 2: sliding windows of 2–4 lines
      for (let winSize = 2; winSize <= 4; winSize++) {
        for (let li = 0; li <= lineCount - winSize; li++) {
          const windowItems = lineGroups.slice(li, li + winSize).flat();
          const windowText = normalizeSpaces(
            reconstructDates(smartJoinLineItems(
              windowItems.sort((a, b) => {
                const dy = b.transform[5] - a.transform[5];
                return Math.abs(dy) > 4 ? dy : a.transform[4] - b.transform[4];
              })
            ))
          );
          if (windowText.includes(targetStr)) {
            const matchItems = findItemsForString(windowItems, targetStr);
            if (matchItems && matchItems.length > 0) {
              devLog(`Page ${pi + 1}: replacing ${label} (window lines ${li}–${li + winSize - 1})`);
              applyTextOverlay(pdfPage, font, matchItems, replaceStr);
              for (let k = li; k < li + winSize; k++) handled.add(k);
              return true;
            }
          }
        }
      }
      return false;
    };

    applyStringReplacement(REQUIRED_STR_1, REPLACE_STR_1, 'string 1');
    applyStringReplacement(REQUIRED_STR_2, REPLACE_STR_2, 'string 2');

    // ── Date replacement (item-by-item, skip already-handled lines) ───────
    for (let li = 0; li < lineCount; li++) {
      if (handled.has(li)) continue;
      for (const item of lineGroups[li]) {
        // Reconstruct a clean date string from the item (handles fragmentation)
        const cleanStr = reconstructDates(normalizeSpaces(item.str));
        if (!RE_DATE_DDMM.test(cleanStr) && !RE_DATE_ISO.test(cleanStr)) continue;

        const newStr = cleanStr
          .replace(/\b\d{2}\/\d{2}\/\d{4}\b/g, today)
          .replace(/\b\d{4}-\d{2}-\d{2}\b/g,   today);

        if (newStr !== cleanStr) {
          dateCount++;
          devLog(`Page ${pi + 1}: date "${item.str}" → "${newStr}"`);
          applyTextOverlay(pdfPage, font, [item], newStr);
        }
      }
    }

    // ── Also check line-level reconstructed dates (handles split dates) ───
    for (let li = 0; li < lineCount; li++) {
      if (handled.has(li)) continue;
      const lineText = normalizeSpaces(reconstructDates(smartJoinLineItems(lineGroups[li])));
      if (!RE_DATE_DDMM.test(lineText) && !RE_DATE_ISO.test(lineText)) continue;

      // Check if line text has a date that no individual item had
      // (means the date was split across items).
      const hasDateInItems = lineGroups[li].some(it => {
        const clean = reconstructDates(normalizeSpaces(it.str));
        return RE_DATE_DDMM.test(clean) || RE_DATE_ISO.test(clean);
      });
      if (!hasDateInItems) {
        // The date is formed by combining multiple items — overlay the whole line.
        const newLineText = lineText
          .replace(/\b\d{2}\/\d{2}\/\d{4}\b/g, today)
          .replace(/\b\d{4}-\d{2}-\d{2}\b/g,   today);
        if (newLineText !== lineText) {
          dateCount++;
          devLog(`Page ${pi + 1}: split date line "${lineText}" → "${newLineText}"`);
          applyTextOverlay(pdfPage, font, lineGroups[li], newLineText);
          handled.add(li);
        }
      }
    }
  }

  devLog(`Total dates replaced: ${dateCount}`);
  return pdfDoc.save();
}

// ═══════════════════════════════════════════════════════════════════════════
// 8.  MAIN PIPELINE — process a single File
// ═══════════════════════════════════════════════════════════════════════════

/**
 * @typedef {{ fileName: string, name: string, iae: string, cnae: string,
 *             status: 'done'|'error', message: string,
 *             outputBytes: Uint8Array|null, outputFileName: string|null }} Result
 */

/**
 * Process one PDF file end-to-end.
 *
 * @param {File} file
 * @returns {Promise<Result>}
 */
async function processPDFFile(file) {
  if (DEV_MODE) {
    _debugReport = [];
    devReport(`\n${'═'.repeat(60)}`);
    devReport(`FILE: ${file.name}`);
    devReport(`${'═'.repeat(60)}`);
  }

  const arrayBuffer = await file.arrayBuffer();

  // ── Extract text ──────────────────────────────────────────────────────
  const { fullText, pageItems, pageLines } = await extractPDFText(arrayBuffer);

  // ── Parse metadata ────────────────────────────────────────────────────
  const { name, doc } = parseName(fullText);
  const iae  = parseIAE(fullText);
  const cnae = parseCNAE(fullText);

  // ── Validate required strings using multi-line scanning ───────────────
  // Collect all line texts across all pages in reading order.
  const allLineTexts = (pageLines || []).flatMap(pg => pg.map(l => l.text));
  const { has1, has2, line1, line2 } = allLineTexts.length > 0
    ? checkRequiredStringsInLines(allLineTexts)
    : checkRequiredStrings(fullText);

  // ── Dev mode report ───────────────────────────────────────────────────
  devLog('─'.repeat(60));
  devLog(`File   : ${file.name}`);
  devLog(`Name   : ${name ?? '(not found)'}   Doc: ${doc ?? '(not found)'}`);
  devLog(`IAE    : ${iae  ?? '(not found)'}`);
  devLog(`CNAE   : ${cnae ?? '(not found)'}`);
  devLog(`String 1 found: ${has1}  («${REQUIRED_STR_1}»)${line1 !== null ? ` at line ${line1}` : ''}`);
  devLog(`String 2 found: ${has2}  («${REQUIRED_STR_2}»)${line2 !== null ? ` at line ${line2}` : ''}`);

  // Count dates found in text for diagnostics.
  const datesDD  = (normalizeSpaces(fullText).match(/\b\d{2}\/\d{2}\/\d{4}\b/g) || []);
  const datesISO = (normalizeSpaces(fullText).match(/\b\d{4}-\d{2}-\d{2}\b/g) || []);
  devLog(`Dates found: ${datesDD.length} DD/MM/YYYY + ${datesISO.length} ISO = ${datesDD.length + datesISO.length} total`);
  devLog(`Decision: ${has1 && has2 ? '✅ generate PDF' : '❌ manual edit required'}`);

  if (DEV_MODE) {
    devReport(`\n--- SEARCH REPORT ---`);
    // YO line
    const yoIdx = allLineTexts.findIndex(t => t.match(/\bYO[,\s]/));
    devReport(`  YO line: ${yoIdx >= 0 ? `line ${yoIdx}: "${allLineTexts[yoIdx]}"` : '(not found)'}`);
    devReport(`  Name: ${name ?? '(not found)'}   Doc: ${doc ?? '(not found)'}`);
    // IAE / CNAE
    const iaeIdx = allLineTexts.findIndex(t => t.includes('IAE'));
    const cnaeIdx = allLineTexts.findIndex(t => t.includes('CNAE'));
    devReport(`  IAE: ${iae ?? '(not found)'} — line ${iaeIdx}`);
    devReport(`  CNAE: ${cnae ?? '(not found)'} — line ${cnaeIdx}`);
    // Required strings
    devReport(`  String 1 (${has1 ? 'FOUND' : 'NOT FOUND'}): ${line1 !== null ? `line ${line1}` : '—'}`);
    devReport(`  String 2 (${has2 ? 'FOUND' : 'NOT FOUND'}): ${line2 !== null ? `line ${line2}` : '—'}`);
    // Dates
    devReport(`  Dates (DD/MM/YYYY): ${datesDD.length} — ${datesDD.slice(0, 5).join(', ')}`);
    devReport(`  Dates (ISO): ${datesISO.length} — ${datesISO.slice(0, 5).join(', ')}`);

    // If strings not found, show candidate lines (debug aid)
    if (!has1) {
      const cands = allLineTexts
        .map((t, i) => ({ t, i }))
        .filter(({ t }) => t.includes('NATALIA') || t.includes('MENSHUTKINA') || t.includes('79235585'));
      devReport(`  Candidates for String 1: ${JSON.stringify(cands)}`);
    }
    if (!has2) {
      const cands = allLineTexts
        .map((t, i) => ({ t, i }))
        .filter(({ t }) => t.includes('AUTORIZADO') || t.includes('241384') || t.includes('RED'));
      devReport(`  Candidates for String 2: ${JSON.stringify(cands)}`);
    }

    devReport(`\nDecision: ${has1 && has2 ? 'DONE' : 'MANUAL EDIT REQUIRED'}`);

    // Update dev panel in UI
    updateDevPanel(file.name, _debugReport.join('\n'));
  }

  if (!has1 || !has2) {
    const missing = [];
    if (!has1) missing.push(`«${REQUIRED_STR_1}»`);
    if (!has2) missing.push(`«${REQUIRED_STR_2}»`);
    return {
      fileName:     file.name,
      name:         name ?? '—',
      iae:          iae  ?? '—',
      cnae:         cnae ?? '—',
      status:       'error',
      message:      `Необходимо ручное редактирование. Не найдено: ${missing.join(', ')}`,
      outputBytes:  null,
      outputFileName: null,
    };
  }

  // ── Generate output PDF ───────────────────────────────────────────────
  const outputBytes = await generateOutputPDF(arrayBuffer, pageItems, pageLines);

  // Build sanitised output filename.
  const baseName = name
    ? sanitizeFileName(name)
    : sanitizeFileName(file.name.replace(/\.pdf$/i, ''));
  const outputFileName = `AUTORIZACION ${baseName}.pdf`;

  return {
    fileName:      file.name,
    name:          name ?? '—',
    iae:           iae  ?? '—',
    cnae:          cnae ?? '—',
    status:        'done',
    message:       'OK',
    outputBytes,
    outputFileName,
  };
}

// ═══════════════════════════════════════════════════════════════════════════
// 9.  DOWNLOAD HELPERS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Trigger a browser download of `bytes` as a PDF with the given filename.
 *
 * @param {Uint8Array} bytes
 * @param {string}     fileName
 */
function downloadFile(bytes, fileName) {
  const blob = new Blob([bytes], { type: 'application/pdf' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Revoke after a generous delay so the download can fully complete even on
  // slow connections or with large PDFs (see URL_REVOKE_DELAY_MS constant).
  setTimeout(() => URL.revokeObjectURL(url), URL_REVOKE_DELAY_MS);
}

/**
 * Download all successful results sequentially, with a random delay between
 * each to allow the browser to handle multiple download requests.
 *
 * @param {Result[]} results
 */
async function downloadAll(results) {
  const successes = results.filter(r => r.status === 'done' && r.outputBytes);
  for (let i = 0; i < successes.length; i++) {
    const r = successes[i];
    downloadFile(r.outputBytes, r.outputFileName);
    if (i < successes.length - 1) {
      // Delay only between files (not after the last one).
      await new Promise(resolve =>
        setTimeout(resolve, randomInt(DOWNLOAD_DELAY_MIN, DOWNLOAD_DELAY_MAX))
      );
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 10.  UI / EVENT HANDLERS
// ═══════════════════════════════════════════════════════════════════════════

const fileInput      = document.getElementById('fileInput');
const generateBtn    = document.getElementById('generateBtn');
const downloadAllBtn = document.getElementById('downloadAllBtn');
const resultsSection = document.getElementById('resultsSection');
const resultsBody    = document.getElementById('resultsBody');
const progressWrap   = document.getElementById('progressWrap');
const progressFill   = document.getElementById('progressFill');
const progressLabel  = document.getElementById('progressLabel');
const hintBanner     = document.getElementById('hint');

/** Shared results array, re-populated on each Generate run. */
let results = [];

// ── Dev panel (only rendered in ?dev mode) ────────────────────────────────
let devPanelEl = null;
let devTextEl  = null;

/**
 * Create and inject the dev panel DOM (once) if ?dev is active.
 */
function ensureDevPanel() {
  if (!DEV_MODE || devPanelEl) return;

  devPanelEl = document.createElement('div');
  devPanelEl.id = 'devPanel';
  devPanelEl.style.cssText = [
    'background:#1a1a2e', 'color:#e0e0e0', 'border-radius:8px',
    'padding:16px', 'margin-top:20px', 'font-family:monospace',
    'font-size:0.78rem', 'overflow-wrap:break-word', 'word-break:break-word',
    'max-height:500px', 'overflow-y:auto', 'box-shadow:0 2px 10px rgba(0,0,0,0.3)',
  ].join(';');

  const header = document.createElement('div');
  header.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;';
  header.innerHTML = '<span style="color:#7dd3fc;font-weight:700">🔍 DEV DEBUG PANEL</span>';

  const copyBtn = document.createElement('button');
  copyBtn.textContent = '📋 Copy debug report';
  copyBtn.style.cssText = [
    'background:#4299e1', 'color:#fff', 'border:none', 'border-radius:4px',
    'padding:4px 10px', 'cursor:pointer', 'font-size:0.78rem', 'font-weight:600',
  ].join(';');
  copyBtn.addEventListener('click', () => {
    const text = _debugReport.join('\n');
    navigator.clipboard.writeText(text).then(() => {
      copyBtn.textContent = '✅ Copied!';
      setTimeout(() => { copyBtn.textContent = '📋 Copy debug report'; }, 2000);
    }).catch(() => {
      // Fallback: show a selectable textarea
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.cssText = 'position:fixed;top:10%;left:5%;width:90%;height:70%;z-index:9999;font-size:12px;';
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.6);z-index:9998;';
      const closeBtn = document.createElement('button');
      closeBtn.textContent = '✕ Close';
      closeBtn.style.cssText = 'position:fixed;top:8%;right:6%;z-index:9999;padding:6px 14px;cursor:pointer;';
      const close = () => { ta.remove(); overlay.remove(); closeBtn.remove(); };
      overlay.addEventListener('click', close);
      closeBtn.addEventListener('click', close);
      document.body.append(overlay, ta, closeBtn);
      ta.select();
    });
  });

  header.appendChild(copyBtn);
  devPanelEl.appendChild(header);

  devTextEl = document.createElement('pre');
  devTextEl.style.cssText = 'margin:0;color:#d1fae5;font-size:0.75rem;';
  devPanelEl.appendChild(devTextEl);

  document.querySelector('.container').appendChild(devPanelEl);
}

/**
 * Update the dev panel text with the current debug report.
 * @param {string} fileName
 * @param {string} reportText
 */
function updateDevPanel(fileName, reportText) {
  if (!DEV_MODE) return;
  ensureDevPanel();
  if (devTextEl) {
    devTextEl.textContent = reportText;
    // Scroll to bottom
    devPanelEl.scrollTop = devPanelEl.scrollHeight;
  }
}

// ── File-input change: enable / disable Generate button ───────────────────
fileInput.addEventListener('change', () => {
  const n = fileInput.files.length;
  if (n > MAX_FILES) {
    alert(`Максимальное количество файлов: ${MAX_FILES}. Выбрано: ${n}.`);
    fileInput.value = '';
    generateBtn.disabled = true;
    return;
  }
  generateBtn.disabled = n === 0;
});

// ── Append a row to the results table ─────────────────────────────────────
function addResultRow(result) {
  const isOk       = result.status === 'done';
  const statusCls  = isOk ? 'status-done'  : 'status-error';
  const statusText = isOk ? '✅ Done'       : '❌ Error';
  const diagCls    = isOk ? 'diag'          : 'diag diag-error';

  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td>${escapeHtml(result.fileName)}</td>
    <td>${escapeHtml(result.name)}</td>
    <td>${escapeHtml(result.iae)}</td>
    <td>${escapeHtml(result.cnae)}</td>
    <td class="${statusCls}">${statusText}</td>
    <td class="${diagCls}">${escapeHtml(result.message)}</td>
  `;
  resultsBody.appendChild(tr);
}

// ── Generate button ────────────────────────────────────────────────────────
generateBtn.addEventListener('click', async () => {
  const files = Array.from(fileInput.files).slice(0, MAX_FILES);
  if (files.length === 0) return;

  // Reset UI state.
  results = [];
  resultsBody.innerHTML = '';
  generateBtn.disabled    = true;
  downloadAllBtn.disabled = true;
  hintBanner.style.display = 'none';
  if (devTextEl) devTextEl.textContent = '';

  // Show section and progress bar.
  resultsSection.style.display = 'block';
  progressWrap.style.display   = 'block';
  progressFill.style.width     = '0%';

  if (DEV_MODE) ensureDevPanel();

  let doneCount = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];

    progressFill.style.width = `${(i / files.length) * 100}%`;
    progressLabel.textContent = `Обработка ${i + 1} / ${files.length}: ${file.name}`;

    try {
      const result = await processPDFFile(file);
      results.push(result);
      addResultRow(result);
      if (result.status === 'done') doneCount++;
    } catch (err) {
      console.error(`Error processing "${file.name}":`, err);
      const errResult = {
        fileName:      file.name,
        name:          '—',
        iae:           '—',
        cnae:          '—',
        status:        'error',
        message:       `Ошибка обработки: ${err.message}`,
        outputBytes:   null,
        outputFileName: null,
      };
      results.push(errResult);
      addResultRow(errResult);
    }
  }

  // Final progress state.
  progressFill.style.width  = '100%';
  progressLabel.textContent = `Завершено: ${doneCount} успешно из ${files.length}`;

  if (doneCount > 0) {
    downloadAllBtn.disabled  = false;
    hintBanner.style.display = 'block';
  }

  generateBtn.disabled = false;
});

// ── Download All button ────────────────────────────────────────────────────
downloadAllBtn.addEventListener('click', async () => {
  downloadAllBtn.disabled = true;
  await downloadAll(results);
  downloadAllBtn.disabled = false;
});

// ── Dev mode indicator in page title ──────────────────────────────────────
if (DEV_MODE) {
  document.title = '[DEV] ' + document.title;
  ensureDevPanel();
}

// ═══════════════════════════════════════════════════════════════════════════
// 11.  INITIALISATION
// ═══════════════════════════════════════════════════════════════════════════

runSelfChecks();
