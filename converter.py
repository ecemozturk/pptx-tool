"""Turn a PPTX deck into one HTML page and PNG per slide."""

import argparse
import html
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from urllib.parse import urlparse

import fitz
from pptx import Presentation
from pptx.oxml.ns import qn


STYLE = """body{font:16px/1.55 system-ui,sans-serif;max-width:1050px;margin:2rem auto;padding:0 1rem;color:#222}
article{margin:0 0 3rem}h2{font-size:1.25rem}.slide{position:relative;line-height:0}
.slide img{display:block;width:100%;height:auto}.slide a{position:absolute;display:block;outline-offset:2px}
.slide a:focus-visible{outline:3px solid #1675cc}.notes{white-space:pre-wrap;line-height:1.55}
.page-number{text-align:center;margin:1.4rem 0;color:#555;font-size:14px}
"""

FONTS = {
    "arial": "Arial,Helvetica,sans-serif",
    "georgia": "Georgia,serif",
    "verdana": "Verdana,Geneva,sans-serif",
    "times": "'Times New Roman',Times,serif",
}


def _safe_url(value):
    value = (value or "").strip()
    try:
        parsed = urlparse(value)
    except ValueError:
        return None
    return value if parsed.scheme.lower() in ("http", "https", "mailto") else None


URL_PATTERN = re.compile(r"(?:https?://|mailto:)[^\s<>\"']+", re.IGNORECASE)


def _note_link(label, url, new_tab):
    target = ' target="_blank" rel="noopener noreferrer"' if new_tab else ""
    return f'<a href="{html.escape(url, quote=True)}"{target}>{html.escape(label)}</a>'


def _linkify_plain_text(text, new_tab):
    parts = []
    cursor = 0
    for match in URL_PATTERN.finditer(text):
        url = match.group().rstrip(".,;:!?)")
        parts.append(html.escape(text[cursor:match.start()]))
        if _safe_url(url):
            parts.append(_note_link(url, url, new_tab))
        else:
            parts.append(html.escape(url))
        cursor = match.start() + len(url)
    parts.append(html.escape(text[cursor:]))
    return "".join(parts)


def _notes_html(frame, new_tab):
    if frame is None:
        return ""

    paragraphs = []
    for paragraph in frame.paragraphs:
        runs = iter(paragraph.runs)
        parts = []
        plain = []
        for child in paragraph._element.content_children:
            url = None
            if child.tag == qn("a:r"):
                run = next(runs)
                try:
                    url = _safe_url(run.hyperlink.address)
                except KeyError:
                    pass
            if url:
                parts.append(_linkify_plain_text("".join(plain), new_tab))
                plain.clear()
                parts.append(_note_link(child.text, url, new_tab))
            else:
                plain.append(child.text)
        parts.append(_linkify_plain_text("".join(plain), new_tab))
        paragraphs.append("".join(parts))
    return "\n".join(paragraphs)


def _shape_links(slide, width, height, new_tab):
    links = []
    for shape in slide.shapes:
        try:
            url = _safe_url(shape.click_action.hyperlink.address)
        except (AttributeError, KeyError, TypeError, ValueError):
            url = None
        if not url:
            continue
        x = max(0, shape.left / width * 100)
        y = max(0, shape.top / height * 100)
        w = min(100 - x, shape.width / width * 100)
        h = min(100 - y, shape.height / height * 100)
        if w <= 0 or h <= 0:
            continue
        target = ' target="_blank" rel="noopener noreferrer"' if new_tab else ""
        links.append(f'<a href="{html.escape(url, quote=True)}"{target} '
                     f'aria-label="Open {html.escape(url, quote=True)}" '
                     f'style="left:{x:.5f}%;top:{y:.5f}%;width:{w:.5f}%;height:{h:.5f}%"></a>')
    return "\n".join(links)


def _pdf_links(page, new_tab):
    links = []
    width, height = page.rect.width, page.rect.height
    target = ' target="_blank" rel="noopener noreferrer"' if new_tab else ""
    for link in page.get_links():
        url = _safe_url(link.get("uri"))
        if not url:
            continue
        box = link["from"] & page.rect
        if box.is_empty:
            continue
        address = html.escape(url, quote=True)
        links.append(
            f'<a href="{address}"{target} aria-label="Open {address}" '
            f'style="left:{100*box.x0/width:.5f}%;top:{100*box.y0/height:.5f}%;'
            f'width:{100*box.width/width:.5f}%;height:{100*box.height/height:.5f}%"></a>'
        )
    return "".join(links)


def _slide_text_layer(page):
    """Move selectable PDF text into HTML while retaining raster artwork."""
    parts = []
    flags = fitz.TEXTFLAGS_DICT & ~fitz.TEXT_PRESERVE_IMAGES
    for block in page.get_text("dict", flags=flags)["blocks"]:
        for line in block.get("lines", []):
            if abs(line["dir"][0] - 1) > 0.001:
                continue
            for span in line["spans"]:
                label = span["text"]
                if not label.strip():
                    continue
                x, y = span["origin"]
                width, height = page.rect.width, page.rect.height
                font = "monospace" if span["flags"] & 8 else (
                    "Georgia,serif" if span["flags"] & 4 else "Arial,sans-serif"
                )
                weight = "bold" if span["flags"] & 16 else "normal"
                style = "italic" if span["flags"] & 2 else "normal"
                parts.append(
                    f'<span style="left:{100*x/width:.5f}%;top:{100*y/height:.5f}%;'
                    f'font:{style} {weight} {100*span["size"]/width:.5f}cqw {font};'
                    f'color:#{span["color"]:06x}">{html.escape(label)}</span>'
                )
                page.add_redact_annot(span["bbox"], fill=False, cross_out=False)
    if parts:
        page.apply_redactions(images=0, graphics=0, text=0)
    return "".join(parts)


def convert(source: Path, destination: Path, *, dpi: int = 144, new_tab: bool = True,
            legacy: bool = False, image_width: int = 607, slide_font: str = "georgia",
            body_font: str = "arial", create_index: bool = False):
    source = source.resolve()
    if source.suffix.lower() != ".pptx" or not source.is_file():
        raise ValueError("Input must be an existing .pptx file")
    if not 72 <= dpi <= 300:
        raise ValueError("DPI must be between 72 and 300")
    if image_width not in (480, 607, 720) or slide_font not in FONTS or body_font not in FONTS:
        raise ValueError("Invalid display setting")
    office = shutil.which("libreoffice") or shutil.which("soffice")
    if not office:
        mac_office = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        if mac_office.is_file():
            office = str(mac_office)
    if not office:
        raise RuntimeError("LibreOffice is required to render slides")

    deck = Presentation(source)
    destination.mkdir(parents=True, exist_ok=True)
    pages = []
    with tempfile.TemporaryDirectory() as temp:
        temp = Path(temp)
        profile = (temp / "lo-profile").as_uri()
        proc = subprocess.run([office, f"-env:UserInstallation={profile}", "--headless",
                               "--convert-to", "pdf", "--outdir", str(temp), str(source)],
                              capture_output=True, text=True, timeout=180)
        pdf_path = temp / f"{source.stem}.pdf"
        if proc.returncode or not pdf_path.exists():
            raise RuntimeError(f"LibreOffice conversion failed: {proc.stderr or proc.stdout}")
        with fitz.open(pdf_path) as pdf:
            if len(pdf) != len(deck.slides):
                raise RuntimeError(f"Slide count mismatch: PPTX {len(deck.slides)}, PDF {len(pdf)}")
            for index, (slide, page) in enumerate(zip(deck.slides, pdf), 1):
                png_name = f"background{index}.png" if legacy else f"slide-{index:03}.png"
                pdf_links = _pdf_links(page, new_tab) if legacy else ""
                text_layer = _slide_text_layer(page) if legacy else ""
                page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False).save(destination / png_name)
                notes = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
                links = pdf_links or _shape_links(slide, deck.slide_width, deck.slide_height, new_tab)
                heading = "" if legacy else f'<h2>Slide {index}</h2>'
                pages.append(f'<article id="slide-{index}">{heading}'
                             f'<div class="slide"><img src="{png_name}" alt="Slide {index}">'
                             f'{text_layer}{links}</div><div class="notes">{_notes_html(notes, new_tab)}</div></article>')
    if legacy:
        style = STYLE + (
            f'body{{font-family:{FONTS[body_font]};max-width:{max(745, image_width+40)}px;margin:29px auto;padding:0 13px}}'
            f'.slide{{container-type:inline-size;width:min({image_width}px,100%);margin:auto}}'
            f'.slide span{{position:absolute;white-space:pre;transform:translateY(-.8em);font-family:{FONTS[slide_font]}!important}}'
            f'.notes{{font:17px/20px {FONTS[body_font]};margin-top:30px}}'
        )
        for index, body in enumerate(pages, 1):
            title = html.escape(source.stem)
            navigation = (f'<nav><a href="index.html">Contents</a></nav>' if create_index else '')
            markup = (
                '<!doctype html><html lang="en"><meta charset="utf-8">'
                '<meta name="viewport" content="width=device-width, initial-scale=1">'
                f'<title>{title} - page {index} of {len(pages)}</title>'
                f'<style>{style}</style><body><main id="pageTable">'
                f'<a id="page{index}"></a>{navigation}{body}'
                f'<div class="page-number">{index} / {len(pages)}</div></main></body></html>'
            )
            (destination / f"pg_{index}.html").write_text(markup, encoding="utf-8")
        if create_index:
            listing = '<ol>' + ''.join(f'<li><a href="pg_{n}.html">Page {n}</a></li>' for n in range(1, len(pages)+1)) + '</ol>'
            (destination / 'index.html').write_text(
                f'<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
                f'<title>{title} - Contents</title><style>{style}</style><body><main><h1>{title}</h1>{listing}</main></body></html>',
                encoding='utf-8')
        return len(pages)

    title = html.escape(source.stem)
    head = (f'<!doctype html>\n<html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>{title}</title><style>{STYLE}'
            f'body{{font-family:{FONTS[body_font]};max-width:{max(745, image_width+40)}px}}'
            f'.slide{{width:min({image_width}px,100%);margin:auto}}'
            f'</style></head><body>')
    for index, page in enumerate(pages, 1):
        navigation = '<nav>'
        if index > 1:
            navigation += f'<a href="pg_{index-1}.html">Previous</a> '
        if create_index:
            navigation += f'<a href="index.html">Contents</a> '
        if index < len(pages):
            navigation += f' <a href="pg_{index+1}.html">Next</a>'
        navigation += '</nav>'
        (destination / f'pg_{index}.html').write_text(
            head + navigation + page + f'<div class="page-number">{index} / {len(pages)}</div></body></html>', encoding='utf-8')
    if create_index:
        listing = '<ol>' + ''.join(f'<li><a href="pg_{n}.html">Slide {n}</a></li>' for n in range(1, len(pages)+1)) + '</ol>'
        (destination / 'index.html').write_text(head + f'<h1>{title}</h1>{listing}</body></html>', encoding='utf-8')
    return len(pages)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--dpi", type=int, default=144)
    parser.add_argument("--same-tab", action="store_true", help="Open slide links in the current tab")
    args = parser.parse_args()
    count = convert(args.pptx, args.output, dpi=args.dpi, new_tab=not args.same_tab)
    print(f"Converted {count} slides to {args.output / 'index.html'} and pg_1.html through pg_{count}.html")


if __name__ == "__main__":
    main()
