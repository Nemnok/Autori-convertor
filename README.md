# Autori-convertor

Fully browser-based (no server, no build step) batch PDF processor for Spanish
tax-authorisation documents.  Reads PDF files, applies specific text / date
replacements, and offers the modified PDFs for download.  Hosted on GitHub Pages.

## How It Works

1. **Upload** one or more single-page PDF files.
2. The app extracts text using **PDF.js** (`getTextContent`) — the *text pipeline*.
3. It searches the **bottom half** of each page (y ≥ 50 %) for two mandatory
   strings and replaces them, along with any dates, using **pdf-lib** overlays.
4. The upper 50 % of the page is **never modified**.
5. Output files are named `AUTORIZACION <NAME SURNAME>.pdf`.

## Automatic OCR Fallback

If the text pipeline cannot find the two required strings in the bottom half
(e.g. the PDF contains scanned images instead of selectable text), an **OCR
fallback** is triggered automatically:

| Step | Detail |
|------|--------|
| Render | PDF page → canvas via PDF.js at 2× scale |
| Crop | Only the **bottom half** of the canvas is kept |
| OCR | [Tesseract.js v5.1.1](https://github.com/nicedoc/tesseract.js) runs on the crop (`spa+eng`) |
| Match | Exact match after whitespace normalisation — **no fuzzy matching** |
| Replace | White-box + new text overlay at OCR-derived positions via pdf-lib |

* **OCR is only invoked when the text pipeline fails** — successful text
  extraction never triggers OCR.
* If the YO (name) line is not found by either pipeline, a separate full-page
  OCR pass is performed solely for name extraction (no replacements).

### Limitations

* OCR accuracy depends on scan quality; poor scans may still require manual
  editing.
* Only exact string matches (after space normalisation) are accepted — OCR
  recognition errors will result in `Необходимо ручное редактирование`.
* Tesseract.js language data is downloaded from a CDN on first use; an internet
  connection is required for the first OCR run.

## Status Messages

| Message | Meaning |
|---------|---------|
| `OK` | Text pipeline succeeded |
| `OK (OCR)` | OCR fallback succeeded |
| `Необходимо ручное редактирование` | Text pipeline failed, no OCR |
| `OCR: Необходимо ручное редактирование` | Both pipelines failed |

## Dev Mode

Append `?dev` to the URL to enable verbose logging and a debug panel:

* Why OCR fallback was triggered (which strings were missing)
* OCR timing
* OCR-recognised text and confidence
* Page cutoff line (50 %) position
