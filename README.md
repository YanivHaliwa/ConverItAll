# con — Universal File Converter

A fast, command-line file converter that handles documents, images, audio, video, archives, and data formats — with automatic Hebrew RTL text fixing for PDF conversions.

---

## Usage

```bash
con <source_file> <target_format>
```

**Examples:**
```bash
con document.pdf md
con video.mp4 mp3
con image.png jpg
con archive.zip rar
con data.json xml
con document.pdf png      # exports all pages into a folder
```

---

## Conversion Table

| From | To |
|---|---|
| **PDF** | md, txt, html, jpg, png, gif, webp, bmp, tiff, ico, svg, avif, heic |
| **DOCX / DOC / ODT / RTF** | pdf, docx, doc, odt, rtf, md, txt, html, epub |
| **MD / TXT / HTML** | pdf, docx, doc, odt, rtf, md, txt, html, epub |
| **XLSX / XLS / ODS** | pdf, xlsx, xls, ods, csv, json, xml, yaml |
| **CSV** | pdf, xlsx, xls, ods, json, xml, yaml |
| **PPTX / PPT / ODP** | pdf, pptx, ppt, odp |
| **JPG / PNG / BMP / WEBP / GIF / TIFF / ICO / SVG / AVIF / HEIC** | jpg, png, gif, webp, bmp, tiff, ico, svg, avif, heic, pdf, html |
| **MP3 / WAV / FLAC / OGG / M4A / AAC / WMA / OPUS** | mp3, wav, flac, ogg, m4a, aac, wma, opus |
| **MP4 / AVI / MKV / MOV / WEBM / FLV / WMV / MPG / M4V** | mp4, avi, mkv, mov, webm, flv, wmv, mpg, m4v, mp3, wav, flac, ogg, m4a, aac, gif, jpg, png, bmp, webp, tiff, ico, svg, avif, heic |
| **ZIP / RAR / 7Z / TAR** | zip, rar, 7z, tar |
| **JSON / XML / YAML** | json, xml, yaml, csv, xlsx, xls, ods |

> **PDF → image** exports every page into a subfolder named after the file (e.g. `document/document-page-1.png`).  
> **PDF → text** uses `pdftotext` (fast) with automatic `markitdown` fallback and smart Hebrew RTL fixing.  
> **Video → GIF** converts the first 10 seconds only.  
> **Video → image** (jpg, png, etc.) extracts the first frame.  
> **Animated GIF → image** extracts the first frame only.

> ⚠️ **Image-based PDFs** (scanned documents or PDFs converted from images) contain no real text layer. Converting them to md, txt, or html will produce empty or useless output. This also means you cannot convert an image (jpg, png, etc.) to any text format — text extraction only works on PDFs with selectable text.

---

## Technologies Used for Conversion

Install only what you need — each tool is used automatically when available:

| Tool | Purpose |
|---|---|
| `libreoffice` | Document conversions (docx, xlsx, pptx, pdf, ...) |
| `pandoc` | Text-based document conversions (md, html, epub, ...) |
| `ffmpeg` | Audio and video conversions |
| `imagemagick` | Image conversions |
| `poppler-utils` | PDF → image (`pdftoppm`, `pdftotext`) |
| `img2pdf` | Image → PDF (high quality) |
| `markitdown` | PDF text extraction fallback (`pip3 install markitdown`) |
| `heif-convert` | HEIC/HEIF image support |
| `zip` / `unzip` / `rar` / `unrar` / `7z` / `tar` | Archive conversions |
| `genisoimage` or `mkisofs` | ISO creation |
| `pyyaml` | YAML support (`pip3 install pyyaml`) |

```bash
# System tools
sudo apt install libreoffice pandoc ffmpeg imagemagick poppler-utils img2pdf p7zip-full rar unrar genisoimage

# Python packages
pip3 install --break-system-packages -r requirements.txt
```

---

## Install

```bash
sudo cp con /usr/local/bin/con
sudo chmod +x /usr/local/bin/con
```

---

## License

MIT

---

Created by [Yaniv Haliwa](https://github.com/YanivHaliwa) for security testing
