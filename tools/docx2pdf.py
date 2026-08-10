#!/usr/bin/env python3
"""Render a docx-js produced .docx to PDF via headless Chromium.

LibreOffice is unusable in this container, so we parse word/document.xml
directly (paragraphs, runs, character formatting, bullets, paragraph
borders) into HTML and let Chromium print it.
"""
import html
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"


def val(node, tag, attr="val"):
    el = node.find(f"{W}{tag}") if node is not None else None
    return el.get(f"{W}{attr}") if el is not None else None


def run_html(r):
    rpr = r.find(f"{W}rPr")
    text = "".join(t.text or "" for t in r.iter(f"{W}t"))
    if not text:
        return ""
    css = []
    if rpr is not None:
        if rpr.find(f"{W}b") is not None:
            css.append("font-weight:700")
        if rpr.find(f"{W}i") is not None:
            css.append("font-style:italic")
        color = val(rpr, "color")
        if color and color.lower() not in ("auto",):
            css.append(f"color:#{color}")
        sz = val(rpr, "sz")
        if sz:
            css.append(f"font-size:{int(sz)/2:.1f}pt")
        if rpr.find(f"{W}caps") is not None:
            css.append("text-transform:uppercase")
        spacing = val(rpr, "spacing")
        if spacing:
            css.append(f"letter-spacing:{int(spacing)/20:.2f}pt")
    style = f' style="{";".join(css)}"' if css else ""
    return f"<span{style}>{html.escape(text)}</span>"


def convert(docx_path, pdf_path):
    docx_path, pdf_path = Path(docx_path).resolve(), Path(pdf_path).resolve()
    xml = zipfile.ZipFile(docx_path).read("word/document.xml")
    body = ET.fromstring(xml).find(f"{W}body")

    parts, open_list = [], False
    for p in body.findall(f"{W}p"):
        ppr = p.find(f"{W}pPr")
        is_bullet = ppr is not None and ppr.find(f"{W}numPr") is not None
        inner = "".join(run_html(r) for r in p.findall(f"{W}r"))

        if is_bullet:
            if not open_list:
                parts.append("<ul>")
                open_list = True
            parts.append(f"<li>{inner or '&nbsp;'}</li>")
            continue
        if open_list:
            parts.append("</ul>")
            open_list = False

        css = []
        if ppr is not None:
            sp = ppr.find(f"{W}spacing")
            if sp is not None:
                if sp.get(f"{W}before"):
                    css.append(f"margin-top:{int(sp.get(f'{W}before'))/20:.1f}pt")
                if sp.get(f"{W}after"):
                    css.append(f"margin-bottom:{int(sp.get(f'{W}after'))/20:.1f}pt")
            bdr = ppr.find(f"{W}pBdr")
            if bdr is not None and bdr.find(f"{W}bottom") is not None:
                b = bdr.find(f"{W}bottom")
                sz = int(b.get(f"{W}sz", "6")) / 8.0
                col = b.get(f"{W}color", "999999")
                css.append(f"border-bottom:{max(sz,0.5):.1f}pt solid #{col}")
                css.append("padding-bottom:3pt")
        style = f' style="{";".join(css)}"' if css else ""
        parts.append(f"<p{style}>{inner or '&nbsp;'}</p>")
    if open_list:
        parts.append("</ul>")

    doc = f"""<!doctype html><html><head><meta charset="utf-8">
<style>
  @page {{ size: A4; margin: 15mm 17mm; }}
  html, body {{ margin:0; padding:0; }}
  body {{ font-family: Carlito, Calibri, "DejaVu Sans", Arial, sans-serif;
          font-size:10pt; line-height:1.32; color:#1A1A1A; }}
  p {{ margin:0 0 4.5pt; }}
  ul {{ margin:0 0 6pt; padding-left:15pt; }}
  li {{ margin:0 0 2.8pt; }}
  li::marker {{ color:#777; }}
</style></head><body>
{chr(10).join(parts)}
</body></html>"""

    tmp = pdf_path.with_suffix(".render.html")
    tmp.write_text(doc, encoding="utf-8")
    subprocess.run(
        [CHROME, "--headless", "--disable-gpu", "--no-sandbox",
         "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}", tmp.as_uri()],
        check=True, capture_output=True, timeout=180,
    )
    tmp.unlink()
    print(f"{pdf_path.name}: {pdf_path.stat().st_size} bytes")


if __name__ == "__main__":
    for src in sys.argv[1:]:
        convert(src, Path(src).with_suffix(".pdf"))
