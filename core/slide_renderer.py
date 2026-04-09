"""Render PPTX slides to images using LibreOffice headless."""

from __future__ import annotations

import os
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

# LibreOffice paths on macOS
SOFFICE_PATHS = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/local/bin/soffice",
    "/usr/bin/soffice",
    "/opt/homebrew/bin/soffice",
]


def _find_soffice() -> str | None:
    """Find the LibreOffice soffice binary."""
    for path in SOFFICE_PATHS:
        if os.path.isfile(path):
            return path
    return None


def render_slides(pptx_bytes: bytes) -> list[bytes]:
    """Render all slides in a PPTX to PNG images.

    Args:
        pptx_bytes: Raw bytes of the .pptx file

    Returns:
        List of PNG image bytes, one per slide (ordered by slide number)
    """
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice not found. Install with: brew install --cask libreoffice"
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        # Write PPTX to temp file
        pptx_path = os.path.join(tmpdir, "presentation.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes)

        # Convert to PNG using LibreOffice
        outdir = os.path.join(tmpdir, "output")
        os.makedirs(outdir)

        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to", "png",
                "--outdir", outdir,
                pptx_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

        # LibreOffice exports a single PNG for single-slide files,
        # or multiple PNGs for multi-slide. Check what we got.
        png_files = sorted(Path(outdir).glob("*.png"))

        if not png_files:
            raise RuntimeError("LibreOffice produced no output images")

        # If only one PNG was produced, LibreOffice might have merged slides.
        # For multi-slide export, we need a different approach: export to PDF first,
        # then convert PDF pages to images.
        if len(png_files) == 1:
            # Try PDF approach for multi-slide support
            return _render_via_pdf(soffice, pptx_bytes, tmpdir)

        images = []
        for png_file in png_files:
            images.append(png_file.read_bytes())

        return images


def _render_via_pdf(soffice: str, pptx_bytes: bytes, tmpdir: str) -> list[bytes]:
    """Render slides via PDF intermediate for multi-slide support."""
    pptx_path = os.path.join(tmpdir, "presentation.pptx")
    pdf_outdir = os.path.join(tmpdir, "pdf_output")
    os.makedirs(pdf_outdir, exist_ok=True)

    # Convert to PDF
    subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", pdf_outdir,
            pptx_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )

    pdf_path = os.path.join(pdf_outdir, "presentation.pdf")
    if not os.path.exists(pdf_path):
        # Fallback: return the single PNG
        png_files = sorted(Path(tmpdir).rglob("*.png"))
        if png_files:
            return [png_files[0].read_bytes()]
        raise RuntimeError("Could not render slides")

    # Convert PDF pages to PNGs using sips (macOS built-in) or LibreOffice
    # Use sips for PDF to PNG conversion on macOS
    png_outdir = os.path.join(tmpdir, "png_pages")
    os.makedirs(png_outdir, exist_ok=True)

    # Try using Python's pdf2image if available, otherwise use sips
    try:
        return _pdf_to_pngs_sips(pdf_path, png_outdir)
    except Exception:
        # Last resort: return single page render
        png_files = sorted(Path(tmpdir).rglob("*.png"))
        if png_files:
            return [png_files[0].read_bytes()]
        raise RuntimeError("Could not render slides to images")


def _pdf_to_pngs_sips(pdf_path: str, outdir: str) -> list[bytes]:
    """Convert PDF to PNGs using macOS sips (one page at a time via Preview/Quartz)."""
    # macOS: use `sips` to convert PDF to PNG (handles only first page)
    # For multi-page, use `mdls` to get page count then extract per-page
    # Actually, the simplest cross-platform approach: use subprocess with
    # `sips -s format png <pdf> --out <png>` - but sips only does first page.

    # Better approach: use Quartz (PyObjC) which is available on macOS
    try:
        import Quartz
        from CoreFoundation import CFURLCreateFromFileSystemRepresentation

        url = Quartz.CFURLCreateFromFileSystemRepresentation(
            None, pdf_path.encode(), len(pdf_path.encode()), False
        )
        pdf_doc = Quartz.CGPDFDocumentCreateWithURL(url)

        if pdf_doc is None:
            raise RuntimeError("Could not open PDF")

        page_count = Quartz.CGPDFDocumentGetNumberOfPages(pdf_doc)
        images = []

        for page_num in range(1, page_count + 1):
            page = Quartz.CGPDFDocumentGetPage(pdf_doc, page_num)
            media_box = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)

            # Scale for good resolution
            scale = 2.0
            width = int(media_box.size.width * scale)
            height = int(media_box.size.height * scale)

            color_space = Quartz.CGColorSpaceCreateDeviceRGB()
            context = Quartz.CGBitmapContextCreate(
                None, width, height, 8, width * 4,
                color_space, Quartz.kCGImageAlphaPremultipliedLast
            )

            # White background
            Quartz.CGContextSetRGBFillColor(context, 1, 1, 1, 1)
            Quartz.CGContextFillRect(context, Quartz.CGRectMake(0, 0, width, height))

            # Scale and draw
            Quartz.CGContextScaleCTM(context, scale, scale)
            Quartz.CGContextDrawPDFPage(context, page)

            # Get image
            cg_image = Quartz.CGBitmapContextCreateImage(context)

            # Save to PNG
            png_path = os.path.join(outdir, f"slide_{page_num:03d}.png")
            url_out = Quartz.CFURLCreateFromFileSystemRepresentation(
                None, png_path.encode(), len(png_path.encode()), False
            )
            dest = Quartz.CGImageDestinationCreateWithURL(url_out, "public.png", 1, None)
            Quartz.CGImageDestinationAddImage(dest, cg_image, None)
            Quartz.CGImageDestinationFinalize(dest)

            images.append(Path(png_path).read_bytes())

        return images

    except ImportError:
        # PyObjC not available, fall back to sips (first page only)
        out_path = os.path.join(outdir, "slide_001.png")
        subprocess.run(
            ["sips", "-s", "format", "png", pdf_path, "--out", out_path],
            capture_output=True, timeout=30,
        )
        if os.path.exists(out_path):
            return [Path(out_path).read_bytes()]
        raise RuntimeError("Could not convert PDF to PNG")


def render_single_slide(pptx_bytes: bytes, slide_index: int) -> bytes | None:
    """Render a single slide to PNG.

    Args:
        pptx_bytes: Raw bytes of the .pptx file
        slide_index: Zero-based slide index

    Returns:
        PNG image bytes, or None if rendering failed
    """
    try:
        images = render_slides(pptx_bytes)
        if slide_index < len(images):
            return images[slide_index]
        return None
    except Exception:
        return None
