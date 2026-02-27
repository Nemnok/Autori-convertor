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
 * URL flag: append  ?dev  to enable verbose console logging.
 */

'use strict';

// ═══════════════════════════════════════════════════════════════════════════
// 1.  CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════

/** Exact string that MUST appear in the source PDF (after space normalisation). */
const REQUIRED_STR_1 = 'A NATALIA MENSHUTKINA CON DNI 79235585X';
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

/** Enable verbose dev logging when URL contains ?dev */
const DEV_MODE = new URLSearchParams(window.location.search).has('dev');

// ═══════════════════════════════════════════════════════════════════════════
// 2.  LOGGING HELPER
// ═══════════════════════════════════════════════════════════════════════════

/** Log only when ?dev is present in the URL. */
const devLog = (...args) => {
  if (DEV_MODE) console.log('[DEV]', ...args);
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
 * Collapse all whitespace sequences to a single space and trim.
 * @param {string} str
 * @returns {string}
 */
function normalizeSpaces(str) {
  return str.replace(/\s+/g, ' ').trim();
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
  const match = norm.match(/\bYO,\s+(.+?),\s*([A-Z0-9]\d{7,8}[A-Z0-9]?)\b/);
  if (!match) return { name: null, doc: null };
  return { name: match[1].trim(), doc: match[2].trim() };
}

/**
 * Extract IAE code from text.
 * @param {string} text
 * @returns {string|null}
 */
function parseIAE(text) {
  const match = text.match(/\bIAE[:\s]+(\d{1,4}(?:\.\d{1,2})?)\b/);
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
 * Extract text and item positions from all pages of a PDF.
 *
 * @param {ArrayBuffer} arrayBuffer  Raw PDF bytes.
 * @returns {Promise<{
 *   fullText: string,
 *   pageItems: Array<{ pageNum: number, items: object[] }>
 * }>}
 */
async function extractPDFText(arrayBuffer) {
  // Use a typed-array copy so PDF.js can own its buffer safely.
  const data = new Uint8Array(arrayBuffer.slice(0));
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;

  let fullText = '';
  const pageItems = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent({ includeMarkedContent: false });
    // Keep only items that have non-empty string content.
    const items = content.items.filter(it => it.str && it.str.length > 0);
    pageItems.push({ pageNum, items });
    fullText += items.map(it => it.str).join(' ') + '\n';
  }

  pdf.destroy();
  return { fullText, pageItems };
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
 * @param {ArrayBuffer} originalArrayBuffer
 * @param {Array<{ pageNum: number, items: object[] }>} pageItems
 * @returns {Promise<Uint8Array>}  Bytes of the output PDF.
 */
async function generateOutputPDF(originalArrayBuffer, pageItems) {
  const today = formatDateDDMMYYYY(new Date());

  // pdf-lib needs its own copy of the buffer.
  const pdfDoc = await PDFLib.PDFDocument.load(originalArrayBuffer.slice(0));
  const pages  = pdfDoc.getPages();

  // Embed a standard font for replacement text.
  const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);

  // Regexes for date detection (non-global — used only with .test()).
  const RE_DATE_DDMM = /\b\d{2}\/\d{2}\/\d{4}\b/;
  const RE_DATE_ISO  = /\b\d{4}-\d{2}-\d{2}\b/;

  for (let pi = 0; pi < pages.length; pi++) {
    const pdfPage = pages[pi];
    const { items } = pageItems[pi] || { items: [] };

    const lines = groupItemsByLine(items);

    for (const line of lines) {
      const lineText = normalizeSpaces(line.map(it => it.str).join(' '));

      // ── String replacement 1 ────────────────────────────────────────────
      if (lineText.includes(REQUIRED_STR_1)) {
        const matchItems = findItemsForString(line, REQUIRED_STR_1);
        if (matchItems && matchItems.length > 0) {
          devLog(`Page ${pi + 1}: replacing string 1`);
          applyTextOverlay(pdfPage, font, matchItems, REPLACE_STR_1);
        }
        continue; // REQUIRED_STR_1 and REQUIRED_STR_2 cannot share a line
      }

      // ── String replacement 2 ────────────────────────────────────────────
      if (lineText.includes(REQUIRED_STR_2)) {
        const matchItems = findItemsForString(line, REQUIRED_STR_2);
        if (matchItems && matchItems.length > 0) {
          devLog(`Page ${pi + 1}: replacing string 2`);
          applyTextOverlay(pdfPage, font, matchItems, REPLACE_STR_2);
        }
        continue;
      }

      // ── Date replacement (item-by-item) ─────────────────────────────────
      for (const item of line) {
        if (!RE_DATE_DDMM.test(item.str) && !RE_DATE_ISO.test(item.str)) continue;

        const newStr = item.str
          .replace(/\b\d{2}\/\d{2}\/\d{4}\b/g, today)
          .replace(/\b\d{4}-\d{2}-\d{2}\b/g,   today);

        if (newStr !== item.str) {
          devLog(`Page ${pi + 1}: date "${item.str}" → "${newStr}"`);
          applyTextOverlay(pdfPage, font, [item], newStr);
        }
      }
    }
  }

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
  const arrayBuffer = await file.arrayBuffer();

  // ── Extract text ──────────────────────────────────────────────────────
  const { fullText, pageItems } = await extractPDFText(arrayBuffer);

  // ── Parse metadata ────────────────────────────────────────────────────
  const { name, doc } = parseName(fullText);
  const iae  = parseIAE(fullText);
  const cnae = parseCNAE(fullText);

  // ── Validate required strings ─────────────────────────────────────────
  const { has1, has2 } = checkRequiredStrings(fullText);

  devLog('─'.repeat(60));
  devLog(`File   : ${file.name}`);
  devLog(`Name   : ${name ?? '(not found)'}   Doc: ${doc ?? '(not found)'}`);
  devLog(`IAE    : ${iae  ?? '(not found)'}`);
  devLog(`CNAE   : ${cnae ?? '(not found)'}`);
  devLog(`String 1 found: ${has1}  («${REQUIRED_STR_1}»)`);
  devLog(`String 2 found: ${has2}  («${REQUIRED_STR_2}»)`);
  devLog(`Decision: ${has1 && has2 ? '✅ generate PDF' : '❌ manual edit required'}`);

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
  const outputBytes = await generateOutputPDF(arrayBuffer, pageItems);

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

  // Show section and progress bar.
  resultsSection.style.display = 'block';
  progressWrap.style.display   = 'block';
  progressFill.style.width     = '0%';

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

// ═══════════════════════════════════════════════════════════════════════════
// 11.  INITIALISATION
// ═══════════════════════════════════════════════════════════════════════════

runSelfChecks();
