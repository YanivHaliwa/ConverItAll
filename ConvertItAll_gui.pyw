#!/usr/bin/env python3
"""
ConvertItAll - Windows GUI Universal File Converter

A graphical, Windows-ready version of the `con` command-line tool.
Converts documents, images, audio, video, archives and data formats by
orchestrating external tools (LibreOffice, pandoc, ffmpeg, ImageMagick,
poppler, 7-Zip, ...). Includes automatic Hebrew RTL fixing for PDF text.

Run with pythonw.exe (or double-click the .pyw) so no console window appears.

Created by Yaniv Haliwa (https://github.com/YanivHaliwa)
"""
#version 14.06.26

import os
import sys
import html
import base64
import struct
import shutil
import tempfile
import threading
import subprocess
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

DEBUG = False
IS_WINDOWS = os.name == "nt"

# ============================================================================
# Format tables (same coverage as the CLI `con`)
# ============================================================================

DOCUMENT_FORMATS = {
    'pdf', 'docx', 'doc', 'xlsx', 'xls', 'csv', 'pptx', 'ppt',
    'odt', 'ods', 'odp', 'txt', 'md', 'html', 'rtf', 'epub',
}
IMAGE_FORMATS = {
    'jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'svg', 'ico',
    'tiff', 'tif', 'heic', 'heif', 'avif',
}
AUDIO_FORMATS = {'mp3', 'wav', 'flac', 'ogg', 'm4a', 'aac', 'wma', 'opus'}
VIDEO_FORMATS = {'mp4', 'avi', 'mkv', 'mov', 'webm', 'flv', 'wmv', 'mpg', 'mpeg', 'm4v'}
DATA_FORMATS = {'json', 'xml', 'yaml', 'yml'}
ARCHIVE_FORMATS = {
    'zip', 'rar', '7z', 'tar', 'tar.gz', 'tar.bz2', 'tar.xz', 'tgz', 'iso',
}

# Target formats the user can pick, per source extension.
def valid_targets(source_ext):
    """Return a sorted list of allowed target formats for a given source ext."""
    e = source_ext
    img_targets = sorted(IMAGE_FORMATS - {'jpeg', 'tif', 'heif'})
    if e == 'pdf':
        return sorted(set(['md', 'txt', 'html']) | set(img_targets))
    if e in ('docx', 'doc', 'odt', 'rtf'):
        return sorted({'pdf', 'docx', 'doc', 'odt', 'rtf', 'md', 'txt', 'html', 'epub'}
                      | set(img_targets))
    if e in ('md', 'txt', 'html'):
        return sorted({'pdf', 'docx', 'doc', 'odt', 'rtf', 'md', 'txt', 'html', 'epub'})
    if e in ('xlsx', 'xls', 'ods'):
        return sorted({'pdf', 'xlsx', 'xls', 'ods', 'csv', 'json', 'xml', 'yaml'} - {e})
    if e == 'csv':
        return sorted({'pdf', 'xlsx', 'xls', 'ods', 'json', 'xml', 'yaml'})
    if e in ('pptx', 'ppt', 'odp'):
        return sorted({'pdf', 'pptx', 'ppt', 'odp'} - {e})
    if e in IMAGE_FORMATS:
        return sorted((set(img_targets) | {'pdf', 'html'}) - {e})
    if e in AUDIO_FORMATS:
        return sorted(AUDIO_FORMATS - {e})
    if e in VIDEO_FORMATS:
        out = set(VIDEO_FORMATS) | set(AUDIO_FORMATS) | set(img_targets)
        return sorted(out - {e})
    if e in DATA_FORMATS:
        return sorted({'json', 'xml', 'yaml', 'csv', 'xlsx', 'xls', 'ods'} - {e})
    if e in ARCHIVE_FORMATS:
        return sorted({'zip', 'rar', '7z', 'tar'} - {e})
    return []


def get_file_category(ext):
    ext = ext.lower().lstrip('.')
    if ext in DOCUMENT_FORMATS:
        return 'document', ext
    if ext in IMAGE_FORMATS:
        return 'image', ext
    if ext in AUDIO_FORMATS:
        return 'audio', ext
    if ext in VIDEO_FORMATS:
        return 'video', ext
    if ext in ARCHIVE_FORMATS:
        return 'archive', ext
    if ext in DATA_FORMATS:
        return 'data', ext
    return None, None


def detect_source_ext(source: Path):
    """Detect extension, honoring compound archive suffixes like .tar.gz."""
    name = source.name.lower()
    for comp in ('tar.gz', 'tar.bz2', 'tar.xz', 'tgz'):
        if name.endswith('.' + comp):
            return comp
    return source.suffix.lower().lstrip('.')


# ============================================================================
# Windows-aware tool resolution + subprocess wrapper
# ============================================================================

# Common Windows install locations searched when a tool is not on PATH.
_PROGRAM_DIRS = [
    os.environ.get('ProgramFiles', r'C:\Program Files'),
    os.environ.get('ProgramFiles(x86)', r'C:\Program Files (x86)'),
    os.environ.get('LOCALAPPDATA', ''),
]

# Where to look for each tool on Windows (glob patterns under the program dirs).
_WINDOWS_HINTS = {
    'soffice':    [r'LibreOffice\program\soffice.exe'],
    'pandoc':     [r'Pandoc\pandoc.exe'],
    'ffmpeg':     [r'ffmpeg\bin\ffmpeg.exe', r'ffmpeg\ffmpeg.exe'],
    'magick':     [r'ImageMagick*\magick.exe'],
    'pdftotext':  [r'poppler*\Library\bin\pdftotext.exe', r'poppler*\bin\pdftotext.exe'],
    'pdftoppm':   [r'poppler*\Library\bin\pdftoppm.exe', r'poppler*\bin\pdftoppm.exe'],
    '7z':         [r'7-Zip\7z.exe'],
    'img2pdf':    [r'img2pdf.exe'],
    'markitdown': [r'markitdown.exe'],
}

# Logical name -> actual executable name candidates (PATH lookup order).
_TOOL_ALIASES = {
    'libreoffice': ['soffice', 'libreoffice'],
    'convert':     ['magick', 'convert'],   # Windows uses `magick`; Linux uses `convert`
    'magick':      ['magick', 'convert'],
}

_tool_cache = {}


def resolve_tool(name):
    """Return a runnable command (string path) for a logical tool, or None.

    On Windows, also searches common Program Files install locations because
    most of these tools are not added to PATH by their installers.
    """
    if name in _tool_cache:
        return _tool_cache[name]

    candidates = _TOOL_ALIASES.get(name, [name])

    # 1) PATH lookup (works on Linux and Windows; .exe handled by shutil.which)
    for cand in candidates:
        found = shutil.which(cand)
        if found:
            _tool_cache[name] = found
            return found

    # 2) Windows install-dir hints
    if IS_WINDOWS:
        for cand in candidates:
            for pattern in _WINDOWS_HINTS.get(cand, []):
                for base in _PROGRAM_DIRS:
                    if not base:
                        continue
                    for match in Path(base).glob(pattern):
                        if match.exists():
                            _tool_cache[name] = str(match)
                            return str(match)

    _tool_cache[name] = None
    return None


def tool_ok(name):
    return resolve_tool(name) is not None


def run_hidden(cmd, timeout=120):
    """Run a subprocess with no flashing console window on Windows.

    The first element of cmd is a logical tool name and is replaced with the
    resolved executable path.
    """
    resolved = resolve_tool(cmd[0])
    if resolved:
        cmd = [resolved] + list(cmd[1:])

    kwargs = dict(capture_output=True, text=True, timeout=timeout)
    if IS_WINDOWS:
        kwargs['creationflags'] = 0x08000000  # CREATE_NO_WINDOW
    return subprocess.run(cmd, **kwargs)


# ============================================================================
# Hebrew / RTL helpers (ported from `con`)
# ============================================================================

def has_hebrew(text):
    return any('֐' <= c <= '׿' for c in text)


def reverse_hebrew_text(text):
    result, buf = [], []
    for c in text:
        if '֐' <= c <= '׿':
            buf.append(c)
        else:
            if buf:
                result.extend(reversed(buf)); buf = []
            result.append(c)
    if buf:
        result.extend(reversed(buf))
    return ''.join(result)


def generate_html_with_smart_rtl(content, title="document"):
    header = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html.escape(title)}</title>
    <style>
        body {{ font-family: Arial, Helvetica, sans-serif; direction: ltr; text-align: left;
               max-width: 1000px; margin: 40px auto; padding: 20px; line-height: 1.8;
               background:#f9f9f9; color:#333; }}
        .container {{ background:#fff; padding:40px; border-radius:8px;
                     box-shadow:0 2px 10px rgba(0,0,0,0.1); }}
        h1,h2,h3,h4 {{ margin-top:1.5em; margin-bottom:.5em; font-weight:bold; }}
        h1 {{ font-size:2em; border-bottom:2px solid #333; padding-bottom:10px; }}
        h2 {{ font-size:1.5em; color:#444; }} h3 {{ font-size:1.2em; color:#555; }}
        h4 {{ font-size:1em; color:#666; }}
        .rtl {{ direction:rtl; text-align:right; }} .ltr {{ direction:ltr; text-align:left; }}
        div {{ margin:.5em 0; }}
    </style>
</head>
<body>
    <div class="container">
"""
    footer = "\n    </div>\n</body>\n</html>\n"
    out = []
    for line in content.split('\n'):
        line = line.rstrip()
        if not line:
            out.append('<br>'); continue
        d = 'rtl' if has_hebrew(line) else 'ltr'
        if line.startswith('#### '):
            out.append(f'<h4 class="{d}">{html.escape(line[5:])}</h4>')
        elif line.startswith('### '):
            out.append(f'<h3 class="{d}">{html.escape(line[4:])}</h3>')
        elif line.startswith('## '):
            out.append(f'<h2 class="{d}">{html.escape(line[3:])}</h2>')
        elif line.startswith('# '):
            out.append(f'<h1 class="{d}">{html.escape(line[2:])}</h1>')
        else:
            out.append(f'<div class="{d}">{html.escape(line)}</div>')
    return header + '\n'.join(out) + footer


# ============================================================================
# PNG / SVG helpers for PDF page export (ported from `con`)
# ============================================================================

def get_png_dimensions(path):
    with path.open('rb') as f:
        h = f.read(24)
    if h[:8] != b'\x89PNG\r\n\x1a\n':
        raise ValueError(f"Not a PNG file: {path}")
    return struct.unpack('>II', h[16:24])


def write_embedded_svg_from_png(src_png, out_svg):
    w, h = get_png_dimensions(src_png)
    b64 = base64.b64encode(src_png.read_bytes()).decode('ascii')
    out_svg.write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" '
        f'viewBox="0 0 {w} {h}">\n'
        f'  <image width="{w}" height="{h}" href="data:image/png;base64,{b64}" />\n</svg>\n',
        encoding='utf-8')


def _page_num(path):
    try:
        return int(path.stem.rsplit('-', 1)[-1])
    except (IndexError, ValueError):
        return 0


def convert_rendered_pdf_page(temp_page, out_page, target_ext, log):
    if target_ext == 'png':
        shutil.move(str(temp_page), str(out_page)); return None
    if target_ext == 'svg':
        write_embedded_svg_from_png(temp_page, out_page); return None
    if not tool_ok('convert'):
        return "Converting PDF pages to .%s requires ImageMagick" % target_ext
    cmd = ['convert', str(temp_page)]
    if target_ext == 'ico':
        cmd += ['-resize', '256x256']
    cmd.append(str(out_page))
    r = run_hidden(cmd, timeout=120)
    if r.returncode != 0 or not out_page.exists():
        return r.stderr.strip() if r.stderr else "ImageMagick did not create output"
    return None


# ============================================================================
# Converters (ported from `con`, print() -> log callback)
# ============================================================================

def convert_document(source, target_ext, source_ext, log):
    target = source.with_suffix(f'.{target_ext}')
    office = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'csv', 'pptx', 'ppt',
              'odt', 'ods', 'odp', 'rtf', 'html'}
    if target_ext in office and tool_ok('libreoffice'):
        try:
            if target_ext == 'csv':
                conv = 'csv:Text - txt - csv (StarCalc):44,34,76'
            else:
                conv = target_ext
            cmd = ['libreoffice', '--headless', '--convert-to', conv,
                   '--outdir', str(source.parent), str(source)]
            r = run_hidden(cmd, timeout=120)
            if r.returncode == 0 and target.exists():
                return True, str(target)
            if DEBUG:
                log(f"[DEBUG] LibreOffice failed: {r.stderr}")
        except subprocess.TimeoutExpired:
            return False, "Timeout (LibreOffice)"
        except Exception as e:
            if DEBUG:
                log(f"[DEBUG] LibreOffice error: {e}")

    if (target_ext in {'md', 'html', 'txt', 'docx', 'pdf', 'epub', 'rtf'}
            and source_ext in {'md', 'html', 'txt', 'docx', 'rtf'}
            and tool_ok('pandoc')):
        try:
            r = run_hidden(['pandoc', str(source), '-o', str(target)], timeout=120)
            if r.returncode == 0 and target.exists():
                return True, str(target)
            if DEBUG:
                log(f"[DEBUG] Pandoc failed: {r.stderr}")
        except subprocess.TimeoutExpired:
            return False, "Timeout (pandoc)"
        except Exception as e:
            if DEBUG:
                log(f"[DEBUG] Pandoc error: {e}")

    return False, "No suitable converter found (install LibreOffice / pandoc)"


def convert_pdf(source, target_ext, source_ext, log):
    target = source.with_suffix(f'.{target_ext}')
    used = False
    if tool_ok('pdftotext'):
        try:
            log("Extracting text from PDF using pdftotext...")
            r = run_hidden(['pdftotext', '-layout', '-enc', 'UTF-8', str(source), '-'],
                           timeout=30)
            if r.returncode == 0 and r.stdout.strip():
                used = True
                content = r.stdout
                if target_ext == 'md':
                    target.write_text(content, encoding='utf-8')
                elif target_ext == 'txt':
                    mark = '‏' if has_hebrew(content) else ''
                    target.write_text(mark + content, encoding='utf-8')
                elif target_ext == 'html':
                    target.write_text(
                        generate_html_with_smart_rtl(content, source.stem),
                        encoding='utf-8')
                log("Conversion successful using pdftotext")
                return True, str(target)
        except subprocess.TimeoutExpired:
            if DEBUG:
                log("[DEBUG] pdftotext timed out, trying markitdown")
        except Exception as e:
            if DEBUG:
                log(f"[DEBUG] pdftotext error: {e}")

    if not used:
        if not tool_ok('markitdown'):
            return False, ("pdftotext unavailable/failed and markitdown not installed "
                           "(pip install markitdown)")
        try:
            log("Extracting text from PDF using markitdown...")
            r = run_hidden(['markitdown', str(source)], timeout=600)
            if r.returncode != 0:
                return False, f"markitdown error: {r.stderr}"
            fixed = '\n'.join(reverse_hebrew_text(l) for l in r.stdout.split('\n'))
            if target_ext == 'md':
                target.write_text(fixed, encoding='utf-8')
            elif target_ext == 'txt':
                mark = '‏' if has_hebrew(fixed) else ''
                target.write_text(mark + fixed, encoding='utf-8')
            elif target_ext == 'html':
                target.write_text(generate_html_with_smart_rtl(fixed, source.stem),
                                  encoding='utf-8')
            else:
                return False, f"Unsupported PDF target: {target_ext}"
            log("Conversion successful using markitdown")
            return True, str(target)
        except subprocess.TimeoutExpired:
            return False, "PDF conversion timed out (600s)"
        except Exception as e:
            return False, str(e)


def convert_image(source, target_ext, source_ext, log):
    if not tool_ok('convert'):
        return False, "ImageMagick not installed"
    target = source.with_suffix(f'.{target_ext}')
    try:
        if source_ext == 'gif':
            cmd = ['convert', f'{str(source)}[0]', str(target)]
        else:
            cmd = ['convert', str(source), str(target)]
        r = run_hidden(cmd, timeout=120)
        if r.returncode == 0 and target.exists():
            return True, str(target)
        return False, f"ImageMagick error: {r.stderr}"
    except subprocess.TimeoutExpired:
        return False, "Timeout"
    except Exception as e:
        return False, str(e)


def convert_audio(source, target_ext, source_ext, log):
    if not tool_ok('ffmpeg'):
        return False, "FFmpeg not installed"
    target = source.with_suffix(f'.{target_ext}')
    try:
        cmd = ['ffmpeg', '-i', str(source), '-y']
        if target_ext == 'mp3':
            cmd += ['-codec:a', 'libmp3lame', '-q:a', '2']
        elif target_ext == 'ogg':
            cmd += ['-codec:a', 'libvorbis', '-q:a', '6']
        elif target_ext == 'flac':
            cmd += ['-codec:a', 'flac']
        elif target_ext == 'aac':
            cmd += ['-codec:a', 'aac', '-b:a', '192k']
        cmd.append(str(target))
        r = run_hidden(cmd, timeout=300)
        if r.returncode == 0 and target.exists():
            return True, str(target)
        return False, "FFmpeg conversion failed"
    except subprocess.TimeoutExpired:
        return False, "Timeout (5 min)"
    except Exception as e:
        return False, str(e)


def convert_video(source, target_ext, source_ext, log):
    if not tool_ok('ffmpeg'):
        return False, "FFmpeg not installed"
    if target_ext in AUDIO_FORMATS:
        return convert_audio(source, target_ext, source_ext, log)
    target = source.with_suffix(f'.{target_ext}')
    try:
        cmd = ['ffmpeg', '-i', str(source), '-y']
        if target_ext == 'mp4':
            cmd += ['-codec:v', 'libx264', '-crf', '23', '-codec:a', 'aac', '-b:a', '192k']
        elif target_ext == 'webm':
            cmd += ['-codec:v', 'libvpx-vp9', '-crf', '30', '-codec:a', 'libopus']
        elif target_ext == 'mkv':
            cmd += ['-codec:v', 'copy', '-codec:a', 'copy']
        cmd.append(str(target))
        log("Converting video... (this may take a while)")
        r = run_hidden(cmd, timeout=600)
        if r.returncode == 0 and target.exists():
            return True, str(target)
        return False, "FFmpeg conversion failed"
    except subprocess.TimeoutExpired:
        return False, "Timeout (10 min)"
    except Exception as e:
        return False, str(e)


def convert_video_to_gif(source, target_ext, source_ext, log):
    if not tool_ok('ffmpeg'):
        return False, "FFmpeg not installed"
    target = source.with_suffix(f'.{target_ext}')
    palette = source.parent / f"{source.stem}_palette.png"
    try:
        r = run_hidden(['ffmpeg', '-i', str(source), '-t', '10',
                        '-vf', 'fps=10,scale=480:-1:flags=lanczos,palettegen',
                        '-y', str(palette)], timeout=60)
        if r.returncode != 0 or not palette.exists():
            return False, "Failed to generate color palette"
        log("Converting video to animated GIF (first 10 seconds)...")
        r = run_hidden(['ffmpeg', '-i', str(source), '-i', str(palette), '-t', '10',
                        '-lavfi', 'fps=10,scale=480:-1:flags=lanczos [x]; [x][1:v] paletteuse',
                        '-y', str(target)], timeout=120)
        palette.unlink(missing_ok=True)
        if r.returncode == 0 and target.exists():
            return True, str(target)
        return False, "FFmpeg GIF conversion failed"
    except subprocess.TimeoutExpired:
        palette.unlink(missing_ok=True)
        return False, "Timeout (2 min)"
    except Exception as e:
        palette.unlink(missing_ok=True)
        return False, str(e)


def convert_video_to_image(source, target_ext, source_ext, log):
    if not tool_ok('ffmpeg'):
        return False, "FFmpeg not installed"
    target = source.with_suffix(f'.{target_ext}')
    try:
        r = run_hidden(['ffmpeg', '-i', str(source), '-ss', '00:00:00',
                        '-vframes', '1', '-y', str(target)], timeout=60)
        if r.returncode == 0 and target.exists():
            return True, str(target)
        return False, "Frame extraction failed"
    except subprocess.TimeoutExpired:
        return False, "Timeout"
    except Exception as e:
        return False, str(e)


def convert_image_to_document(source, target_ext, source_ext, log):
    target = source.with_suffix(f'.{target_ext}')
    if target_ext == 'pdf':
        if tool_ok('img2pdf'):
            try:
                r = run_hidden(['img2pdf', str(source), '-o', str(target)], timeout=60)
                if r.returncode == 0 and target.exists():
                    return True, str(target)
            except Exception:
                pass
        if tool_ok('convert'):
            try:
                r = run_hidden(['convert', str(source), str(target)], timeout=60)
                if r.returncode == 0 and target.exists():
                    return True, str(target)
                return False, f"ImageMagick error: {r.stderr}"
            except Exception as e:
                return False, str(e)
        return False, "No PDF converter found (install img2pdf or ImageMagick)"

    if target_ext == 'html':
        try:
            data = base64.b64encode(source.read_bytes()).decode('utf-8')
            mimes = {'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'png': 'image/png',
                     'gif': 'image/gif', 'bmp': 'image/bmp', 'webp': 'image/webp',
                     'svg': 'image/svg+xml', 'ico': 'image/x-icon'}
            mime = mimes.get(source_ext, 'image/png')
            safe = html.escape(source.name)
            target.write_text(
                '<!DOCTYPE html>\n<html>\n<head>\n<meta charset="UTF-8">\n'
                f'<title>{safe}</title>\n<style>body{{margin:0;display:flex;'
                'justify-content:center;align-items:center;min-height:100vh;'
                'background:#f0f0f0;}img{max-width:100%;max-height:100vh;'
                'box-shadow:0 4px 6px rgba(0,0,0,0.1);}</style>\n</head>\n<body>\n'
                f'<img src="data:{mime};base64,{data}" alt="{safe}">\n</body>\n</html>',
                encoding='utf-8')
            return True, str(target)
        except Exception as e:
            return False, f"HTML creation error: {e}"

    return False, f"Image to {target_ext} not supported (only PDF and HTML)"


def convert_document_to_image(source, target_ext, source_ext, log, pdf_all_pages=False):
    target = source.with_suffix(f'.{target_ext}')
    if source_ext == 'pdf':
        if pdf_all_pages:
            if not tool_ok('pdftoppm'):
                return False, "PDF to image requires pdftoppm (poppler)"
            out_dir = source.parent / source.stem
            if out_dir.exists() and not out_dir.is_dir():
                return False, f"Output path exists and is not a folder: {out_dir}"
            out_dir.mkdir(exist_ok=True)
            for old in out_dir.glob(f"{source.stem}-page-*.{target_ext}"):
                old.unlink(missing_ok=True)
            with tempfile.TemporaryDirectory(prefix=f"{source.stem}-pages-") as td:
                td = Path(td)
                prefix = str(td / f"{source.stem}-page")
                r = run_hidden(['pdftoppm', '-png', str(source), prefix], timeout=120)
                pages = sorted(td.glob(f"{source.stem}-page-*.png"), key=_page_num)
                if r.returncode != 0 or not pages:
                    return False, "pdftoppm did not create page images"
                for p in pages:
                    n = _page_num(p)
                    out_page = out_dir / f"{source.stem}-page-{n}.{target_ext}"
                    err = convert_rendered_pdf_page(p, out_page, target_ext, log)
                    if err:
                        return False, err
            return True, str(out_dir)

        if tool_ok('pdftoppm'):
            try:
                prefix = str(target.with_suffix(''))
                fmt = {'png': 'png', 'jpg': 'jpeg', 'jpeg': 'jpeg',
                       'tiff': 'tiff', 'tif': 'tiff'}.get(target_ext, 'png')
                r = run_hidden(['pdftoppm', '-' + fmt, '-singlefile', str(source), prefix],
                               timeout=60)
                created = Path(f"{prefix}.{fmt if fmt != 'jpeg' else 'jpg'}")
                if r.returncode == 0 and created.exists():
                    if created != target:
                        created.rename(target)
                    return True, str(target)
            except Exception:
                pass
        if tool_ok('convert'):
            try:
                r = run_hidden(['convert', '-density', '300', f'{str(source)}[0]',
                                str(target)], timeout=60)
                if r.returncode == 0 and target.exists():
                    return True, str(target)
                return False, f"ImageMagick error: {r.stderr}"
            except Exception as e:
                return False, str(e)
        return False, "No PDF to image converter (install poppler or ImageMagick)"

    # Other documents: -> PDF first, then PDF -> image.
    if tool_ok('libreoffice'):
        try:
            r = run_hidden(['libreoffice', '--headless', '--convert-to', 'pdf',
                            '--outdir', str(source.parent), str(source)], timeout=120)
            tmp_pdf = source.with_suffix('.pdf')
            if r.returncode == 0 and tmp_pdf.exists():
                ok, res = convert_document_to_image(tmp_pdf, target_ext, 'pdf', log)
                try:
                    tmp_pdf.unlink()
                except Exception:
                    pass
                return ok, res
        except Exception as e:
            if DEBUG:
                log(f"[DEBUG] Document to PDF failed: {e}")
    return False, f"Cannot convert {source_ext} to image (LibreOffice missing)"


def convert_data(source, target_ext, source_ext, log):
    import json, csv as csv_mod, re
    import xml.etree.ElementTree as ET
    import xml.dom.minidom
    try:
        # Safe against XXE / billion-laughs when available; stdlib is the fallback.
        from defusedxml.ElementTree import parse as _safe_xml_parse
    except ImportError:
        _safe_xml_parse = ET.parse
    target = source.with_suffix(f'.{target_ext}')

    def read_json(p):
        with open(p, encoding='utf-8') as f:
            return json.load(f)

    def read_xml(p):
        root = _safe_xml_parse(p).getroot()
        if len(root) == 0:
            return {root.tag: root.text}
        result = []
        for child in root:
            item = {sc.tag: sc.text for sc in child}
            if item:
                result.append(item)
        return result if result else {c.tag: c.text for c in root}

    def read_csv(p):
        with open(p, encoding='utf-8') as f:
            return list(csv_mod.DictReader(f))

    def read_yaml(p):
        import yaml
        with open(p, encoding='utf-8') as f:
            return yaml.safe_load(f)

    def sanitize(name):
        name = re.sub(r'[^a-zA-Z0-9_.-]', '_', str(name))
        if name and (name[0].isdigit() or name[0] in '.-'):
            name = 'item_' + name
        return name or 'item'

    def write_json(d, p):
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(d, f, indent=2, ensure_ascii=False)

    def write_xml(d, p):
        root = ET.Element("root")
        if isinstance(d, list):
            for item in d:
                el = ET.SubElement(root, "item")
                if isinstance(item, dict):
                    for k, v in item.items():
                        ch = ET.SubElement(el, sanitize(k))
                        ch.text = str(v) if v is not None else ''
                else:
                    el.text = str(item) if item is not None else ''
        elif isinstance(d, dict):
            for k, v in d.items():
                ch = ET.SubElement(root, sanitize(k))
                ch.text = str(v) if v is not None else ''
        pretty = xml.dom.minidom.parseString(
            ET.tostring(root, encoding='utf-8')).toprettyxml(indent="  ", encoding='utf-8')
        with open(p, 'wb') as f:
            f.write(pretty)

    def write_csv(d, p):
        if isinstance(d, dict):
            d = [d]
        if not d:
            return
        with open(p, 'w', encoding='utf-8', newline='') as f:
            w = csv_mod.DictWriter(f, fieldnames=d[0].keys())
            w.writeheader(); w.writerows(d)

    def write_yaml(d, p):
        import yaml
        with open(p, 'w', encoding='utf-8') as f:
            yaml.dump(d, f, allow_unicode=True, default_flow_style=False)

    try:
        if source_ext == 'json':
            data = read_json(source)
        elif source_ext == 'xml':
            data = read_xml(source)
        elif source_ext == 'csv':
            data = read_csv(source)
        elif source_ext in ('yaml', 'yml'):
            data = read_yaml(source)
        else:
            return False, f"Unsupported source: {source_ext}"

        if target_ext == 'json':
            write_json(data, target)
        elif target_ext == 'xml':
            write_xml(data, target)
        elif target_ext == 'csv':
            write_csv(data, target)
        elif target_ext in ('yaml', 'yml'):
            write_yaml(data, target)
        else:
            return False, f"Unsupported target: {target_ext}"
        return (True, str(target)) if target.exists() else (False, "No output produced")
    except ImportError:
        return False, "PyYAML not installed (pip install pyyaml)"
    except Exception as e:
        return False, str(e)


# ----- Archive helpers (ported, with zip-slip protection) -----

def _check_archive_safe(source, source_ext):
    try:
        if source_ext == 'zip':
            cmd = ['unzip', '-l', str(source)]
        elif source_ext == 'rar':
            cmd = ['unrar', 'l', str(source)]
        elif source_ext in ('7z', 'iso'):
            cmd = ['7z', 'l', str(source)]
        elif source_ext in ('tar', 'tar.gz', 'tgz', 'tar.bz2', 'tar.xz'):
            cmd = ['tar', 'tf', str(source)]
        else:
            return True
        r = run_hidden(cmd, timeout=30)
        if r.returncode != 0:
            return True
        for line in r.stdout.splitlines():
            s = line.strip()
            if '../' in s or s.startswith('/'):
                return False
        return True
    except Exception:
        return True


def _extract_archive(source, temp_dir, source_ext, log):
    if not _check_archive_safe(source, source_ext):
        return False
    try:
        if source_ext == 'zip':
            cmd = ['unzip', '-q', str(source), '-d', str(temp_dir)]
        elif source_ext == 'rar':
            cmd = ['unrar', 'x', '-y', str(source), str(temp_dir) + os.sep]
        elif source_ext in ('7z', 'iso'):
            cmd = ['7z', 'x', f'-o{temp_dir}', '-y', str(source)]
        elif source_ext in ('tar', 'tar.gz', 'tgz', 'tar.bz2', 'tar.xz'):
            cmd = ['tar', 'xf', str(source), '-C', str(temp_dir)]
        else:
            return False
        return run_hidden(cmd, timeout=300).returncode == 0
    except Exception:
        return False


def _create_archive(temp_dir, target_file, target_ext, log):
    try:
        items = [i.name for i in temp_dir.iterdir()]
        if not items:
            return False
        out = str(target_file.resolve())
        if target_ext == 'zip':
            cmd = ['zip', '-r', '-q', out] + items
        elif target_ext == 'rar':
            cmd = ['rar', 'a', '-r', '-ep1', out] + items
        elif target_ext == '7z':
            cmd = ['7z', 'a', out] + items
        elif target_ext == 'tar':
            cmd = ['tar', 'cf', out] + items
        else:
            return False
        resolved = resolve_tool(cmd[0])
        if resolved:
            cmd[0] = resolved
        kwargs = dict(capture_output=True, text=True, timeout=300, cwd=str(temp_dir))
        if IS_WINDOWS:
            kwargs['creationflags'] = 0x08000000
        return subprocess.run(cmd, **kwargs).returncode == 0
    except Exception:
        return False


def convert_archive(source, target_ext, source_ext, log):
    name = source.name.lower()
    base = source.stem
    for comp in ('tar.gz', 'tar.bz2', 'tar.xz', 'tgz'):
        if name.endswith('.' + comp):
            base = source.name[:-(len(comp) + 1)]
            break
    target = source.parent / f"{base}.{target_ext}"
    temp_dir = Path(tempfile.mkdtemp(prefix='con_archive_'))
    try:
        log(f"Extracting {source.name}...")
        if not _extract_archive(source, temp_dir, source_ext, log):
            return False, f"Failed to extract {source_ext} archive"
        log(f"Creating {target_ext} archive...")
        if not _create_archive(temp_dir, target, target_ext, log):
            return False, f"Failed to create {target_ext} archive (tool installed?)"
        return (True, str(target)) if target.exists() else (False, "Archive not found")
    except Exception as e:
        return False, f"Archive error: {e}"
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# ============================================================================
# Dispatcher (mirrors `con` main routing)
# ============================================================================

def do_conversion(source_path, target_format, log):
    source = Path(source_path).resolve()
    if not source.exists():
        return False, f"Source file not found: {source_path}"

    source_ext = detect_source_ext(source)
    src_cat, src_e = get_file_category(source_ext)
    tgt_cat, tgt_e = get_file_category(target_format)
    if src_e:
        source_ext = src_e
    if tgt_e:
        target_format = tgt_e

    if not src_cat:
        return False, f"Unsupported source format '.{source_ext}'"
    if not tgt_cat:
        return False, f"Unsupported target format '.{target_format}'"
    if source_ext == target_format:
        return False, f"Source and target are the same (.{source_ext})"

    log(f"Converting: {source.name} -> .{target_format}")

    if source_ext == 'pdf' and target_format in ('md', 'txt', 'html'):
        return convert_pdf(source, target_format, source_ext, log)
    if src_cat == 'document' and tgt_cat == 'document':
        return convert_document(source, target_format, source_ext, log)
    if src_cat == 'image' and tgt_cat == 'image':
        return convert_image(source, target_format, source_ext, log)
    if src_cat == 'audio' and tgt_cat == 'audio':
        return convert_audio(source, target_format, source_ext, log)
    if src_cat == 'video' and tgt_cat == 'video':
        return convert_video(source, target_format, source_ext, log)
    if src_cat == 'archive' and tgt_cat == 'archive':
        return convert_archive(source, target_format, source_ext, log)
    if src_cat == 'data' and tgt_cat == 'data':
        return convert_data(source, target_format, source_ext, log)
    if src_cat == 'document' and tgt_cat == 'image':
        return convert_document_to_image(source, target_format, source_ext, log,
                                         pdf_all_pages=(source_ext == 'pdf'))
    if src_cat == 'image' and tgt_cat == 'document':
        return convert_image_to_document(source, target_format, source_ext, log)
    if src_cat == 'video' and tgt_cat == 'audio':
        return convert_video(source, target_format, source_ext, log)
    if src_cat == 'video' and tgt_cat == 'image' and target_format == 'gif':
        return convert_video_to_gif(source, target_format, source_ext, log)
    if src_cat == 'video' and tgt_cat == 'image':
        return convert_video_to_image(source, target_format, source_ext, log)
    if src_cat == 'data' and tgt_cat == 'document' and target_format == 'csv':
        return convert_data(source, target_format, source_ext, log)
    if src_cat == 'document' and source_ext == 'csv' and tgt_cat == 'data':
        return convert_data(source, target_format, source_ext, log)
    if src_cat == 'document' and source_ext in ('xlsx', 'xls', 'ods') and tgt_cat == 'data':
        log("Converting spreadsheet to CSV first...")
        ok, res = convert_document(source, 'csv', source_ext, log)
        if ok:
            r = convert_data(Path(res), target_format, 'csv', log)
            Path(res).unlink(missing_ok=True)
            return r
        return False, "Failed to convert spreadsheet to CSV"
    if src_cat == 'data' and tgt_cat == 'document' and target_format in ('xlsx', 'xls', 'ods'):
        log("Converting data to CSV first...")
        ok, res = convert_data(source, 'csv', source_ext, log)
        if ok:
            r = convert_document(Path(res), target_format, 'csv', log)
            Path(res).unlink(missing_ok=True)
            return r
        return False, "Failed to convert data to CSV"

    return False, f"Cannot convert {src_cat} to {tgt_cat}"


# ============================================================================
# GUI
# ============================================================================

class ConvertItAllApp:
    # File-dialog filters grouped by category.
    FILE_TYPES = [
        ("All supported", "*.pdf *.docx *.doc *.odt *.rtf *.txt *.md *.html "
                          "*.xlsx *.xls *.ods *.csv *.pptx *.ppt *.odp "
                          "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.svg *.ico "
                          "*.tiff *.tif *.heic *.heif *.avif "
                          "*.mp3 *.wav *.flac *.ogg *.m4a *.aac *.wma *.opus "
                          "*.mp4 *.avi *.mkv *.mov *.webm *.flv *.wmv *.mpg *.mpeg *.m4v "
                          "*.zip *.rar *.7z *.tar *.tar.gz *.tar.bz2 *.tar.xz *.tgz *.iso "
                          "*.json *.xml *.yaml *.yml"),
        ("Documents", "*.pdf *.docx *.doc *.odt *.rtf *.txt *.md *.html "
                      "*.xlsx *.xls *.ods *.csv *.pptx *.ppt *.odp"),
        ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.svg *.ico *.tiff *.tif "
                   "*.heic *.heif *.avif"),
        ("Audio", "*.mp3 *.wav *.flac *.ogg *.m4a *.aac *.wma *.opus"),
        ("Video", "*.mp4 *.avi *.mkv *.mov *.webm *.flv *.wmv *.mpg *.mpeg *.m4v"),
        ("Archives", "*.zip *.rar *.7z *.tar *.tar.gz *.tar.bz2 *.tar.xz *.tgz *.iso"),
        ("Data", "*.json *.xml *.yaml *.yml"),
        ("All files", "*.*"),
    ]

    def __init__(self, root):
        self.root = root
        root.title("ConvertItAll - Universal File Converter")
        root.geometry("680x520")
        root.minsize(560, 440)

        self.source_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self._build_ui()

    def _build_ui(self):
        pad = {'padx': 10, 'pady': 6}
        frm = ttk.Frame(self.root)
        frm.pack(fill='both', expand=True)

        # Source row
        row1 = ttk.Frame(frm)
        row1.pack(fill='x', **pad)
        ttk.Label(row1, text="Source file:", width=12).pack(side='left')
        self.source_entry = ttk.Entry(row1, textvariable=self.source_var)
        self.source_entry.pack(side='left', fill='x', expand=True, padx=(0, 6))
        ttk.Button(row1, text="Browse...", command=self.browse).pack(side='left')

        # Target row
        row2 = ttk.Frame(frm)
        row2.pack(fill='x', **pad)
        ttk.Label(row2, text="Convert to:", width=12).pack(side='left')
        self.target_combo = ttk.Combobox(row2, textvariable=self.target_var,
                                          state='readonly', width=14)
        self.target_combo.pack(side='left')
        self.convert_btn = ttk.Button(row2, text="Convert", command=self.start_convert)
        self.convert_btn.pack(side='left', padx=10)

        # Log
        ttk.Label(frm, text="Log:").pack(anchor='w', padx=10)
        self.log_box = scrolledtext.ScrolledText(frm, height=16, state='disabled',
                                                  wrap='word', font=('Consolas', 9))
        self.log_box.pack(fill='both', expand=True, padx=10, pady=(0, 6))

        # Status bar
        status = ttk.Label(self.root, textvariable=self.status_var, relief='sunken',
                           anchor='w')
        status.pack(fill='x', side='bottom')

        self.source_var.trace_add('write', lambda *a: self._refresh_targets())

    # ---- helpers ----
    def log(self, msg):
        self.root.after(0, self._log_main, msg)

    def _log_main(self, msg):
        self.log_box.configure(state='normal')
        self.log_box.insert('end', msg + '\n')
        self.log_box.see('end')
        self.log_box.configure(state='disabled')

    def browse(self):
        path = filedialog.askopenfilename(title="Select file to convert",
                                          filetypes=self.FILE_TYPES)
        if path:
            self.source_var.set(path)

    def _refresh_targets(self):
        path = self.source_var.get().strip()
        if not path:
            self.target_combo['values'] = []
            self.target_var.set('')
            return
        ext = detect_source_ext(Path(path))
        targets = valid_targets(ext)
        self.target_combo['values'] = targets
        if targets:
            self.target_var.set(targets[0])
            self.status_var.set(f"Detected .{ext} ({get_file_category(ext)[0]}) "
                                f"- {len(targets)} target formats available")
        else:
            self.target_var.set('')
            self.status_var.set(f"Unsupported source format: .{ext}")

    def start_convert(self):
        src = self.source_var.get().strip()
        tgt = self.target_var.get().strip()
        if not src:
            messagebox.showwarning("No file", "Please select a source file.")
            return
        if not os.path.exists(src):
            messagebox.showerror("Not found", f"File not found:\n{src}")
            return
        if not tgt:
            messagebox.showwarning("No target", "Please choose a target format.")
            return

        self.convert_btn.configure(state='disabled')
        self.status_var.set("Converting...")
        threading.Thread(target=self._worker, args=(src, tgt), daemon=True).start()

    def _worker(self, src, tgt):
        try:
            ok, result = do_conversion(src, tgt, self.log)
        except Exception as e:
            ok, result = False, str(e)
        self.root.after(0, self._done, ok, result)

    def _done(self, ok, result):
        self.convert_btn.configure(state='normal')
        if ok:
            self.log(f"SUCCESS: {result}")
            self.status_var.set("Done")
            messagebox.showinfo("Success", f"Converted successfully:\n{result}")
        else:
            self.log(f"FAILED: {result}")
            self.status_var.set("Failed")
            messagebox.showerror("Conversion failed", result)


def main():
    root = tk.Tk()
    try:
        # Crisper text on Windows high-DPI displays.
        if IS_WINDOWS:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    ConvertItAllApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
