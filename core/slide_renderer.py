"""Render PPTX slides to images using LibreOffice headless.

Cross-platform: macOS, Linux, Windows.
PDF-to-PNG conversion via Quartz (macOS preferred) or pdf2image (all platforms).
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional

# Candidate soffice paths per platform
_SOFFICE_PATHS_MACOS = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
]

_SOFFICE_PATHS_LINUX = [
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/lib/libreoffice/program/soffice",
    "/snap/bin/libreoffice",
]

_SOFFICE_PATHS_WINDOWS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]

_INSTALL_HINTS = {
    "Darwin": "brew install --cask libreoffice",
    "Linux": "sudo apt install libreoffice   # or: sudo dnf install libreoffice",
    "Windows": "Download from https://www.libreoffice.org/download/",
}


def _find_soffice() -> str | None:
    """Find the LibreOffice soffice binary on any supported OS."""
    system = platform.system()

    if system == "Darwin":
        candidates = _SOFFICE_PATHS_MACOS
    elif system == "Linux":
        candidates = _SOFFICE_PATHS_LINUX
    elif system == "Windows":
        candidates = _SOFFICE_PATHS_WINDOWS
    else:
        candidates = []

    for path in candidates:
        if os.path.isfile(path):
            return path

    # Universal fallback: check PATH
    found = shutil.which("soffice") or shutil.which("libreoffice")
    return found


def render_slides(pptx_bytes: bytes) -> list[bytes]:
    """Render all slides in a PPTX to PNG images.

    Args:
        pptx_bytes: Raw bytes of the .pptx file

    Returns:
        List of PNG image bytes, one per slide (ordered by slide number)
    """
    soffice = _find_soffice()
    if not soffice:
        system = platform.system()
        hint = _INSTALL_HINTS.get(system, "Install LibreOffice from https://www.libreoffice.org/download/")
        raise RuntimeError(f"LibreOffice not found. Install with:\n  {hint}")

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

    # Convert PDF pages to PNGs
    png_outdir = os.path.join(tmpdir, "png_pages")
    os.makedirs(png_outdir, exist_ok=True)

    # Try platform-specific methods first, then cross-platform fallback
    try:
        return _pdf_to_pngs_quartz(pdf_path, png_outdir)
    except (ImportError, Exception):
        pass

    try:
        return _pdf_to_pngs_pdf2image(pdf_path, png_outdir)
    except (ImportError, Exception):
        pass

    # Last resort: return single page render
    png_files = sorted(Path(tmpdir).rglob("*.png"))
    if png_files:
        return [png_files[0].read_bytes()]
    raise RuntimeError("Could not render slides to images. Install poppler-utils or pdf2image.")


def _pdf_to_pngs_quartz(pdf_path: str, outdir: str) -> list[bytes]:
    """Convert PDF to PNGs using macOS Quartz (highest quality, macOS only)."""
    import Quartz  # type: ignore[import-untyped]

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


def _pdf_to_pngs_pdf2image(pdf_path: str, outdir: str) -> list[bytes]:
    """Convert PDF to PNGs using pdf2image (cross-platform, requires Poppler)."""
    from pdf2image import convert_from_path  # type: ignore[import-untyped]

    pil_images = convert_from_path(pdf_path, dpi=200, fmt="png")

    images = []
    for i, pil_img in enumerate(pil_images):
        png_path = os.path.join(outdir, f"slide_{i + 1:03d}.png")
        pil_img.save(png_path, "PNG")
        images.append(Path(png_path).read_bytes())

    return images
