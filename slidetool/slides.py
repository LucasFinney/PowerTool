"""Rasterize a .pptx to one PNG per slide.

Primary path: PowerPoint COM automation via comtypes. Faithful rendering of
fonts, themes, and the final state of slide builds.

Fallback: LibreOffice headless (`soffice --headless`). Used if PowerPoint
or comtypes isn't available — handy for CI / non-Windows dev.
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

# PowerPoint's Export image format codes (PpFixedFormatType is a different
# enum). For Presentation.Export(path, "PNG"), the second arg is a filter
# name string; PowerPoint creates one image per slide in `path`.


def rasterize_pptx(
    pptx_path: str | Path,
    out_dir: str | Path,
    width: int = 1920,
    height: int = 1080,
) -> list[Path]:
    """Convert each slide to a PNG. Returns the PNG paths in slide order."""
    pptx_path = Path(pptx_path).resolve()
    out_dir = Path(out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not pptx_path.exists():
        raise FileNotFoundError(pptx_path)

    try:
        pngs = _rasterize_via_powerpoint(pptx_path, out_dir, width, height)
    except _PowerPointUnavailable:
        pngs = _rasterize_via_libreoffice(pptx_path, out_dir)

    if not pngs:
        raise RuntimeError(f"Rasterization produced no images in {out_dir}")
    return pngs


class _PowerPointUnavailable(RuntimeError):
    pass


def _rasterize_via_powerpoint(
    pptx_path: Path, out_dir: Path, width: int, height: int
) -> list[Path]:
    try:
        import comtypes.client  # type: ignore
    except ImportError as e:
        raise _PowerPointUnavailable("comtypes not installed") from e

    try:
        ppt = comtypes.client.CreateObject("PowerPoint.Application")
    except Exception as e:
        raise _PowerPointUnavailable(f"PowerPoint COM unavailable: {e}") from e

    presentation = None
    try:
        # WithWindow=False keeps PowerPoint headless-ish.
        presentation = ppt.Presentations.Open(
            str(pptx_path), ReadOnly=True, Untitled=False, WithWindow=False
        )
        # Export each slide individually so we control the filename order.
        for i, slide in enumerate(presentation.Slides, start=1):
            out = out_dir / f"slide_{i:04d}.png"
            slide.Export(str(out), "PNG", width, height)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            ppt.Quit()
        except Exception:
            pass

    return sorted(out_dir.glob("slide_*.png"))


def _rasterize_via_libreoffice(pptx_path: Path, out_dir: Path) -> list[Path]:
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        raise RuntimeError(
            "Neither PowerPoint COM nor LibreOffice (soffice) is available."
        )
    # LibreOffice converts the whole deck to a single multi-page PDF first,
    # then we'd need pdftoppm to split. Simpler: convert directly to PNG —
    # LibreOffice writes one PNG per slide using a naming scheme.
    subprocess.run(
        [soffice, "--headless", "--convert-to", "png", "--outdir",
         str(out_dir), str(pptx_path)],
        check=True,
    )
    # LibreOffice's PNG export only emits the first slide. For the fallback
    # path we go via PDF and split with pdftoppm if available.
    pngs = sorted(out_dir.glob("*.png"))
    if len(pngs) > 1:
        return pngs

    # Fallback-of-fallback: PDF + pdftoppm.
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir",
         str(out_dir), str(pptx_path)],
        check=True,
    )
    pdf = next(out_dir.glob(f"{pptx_path.stem}.pdf"), None)
    if pdf is None:
        raise RuntimeError("LibreOffice failed to produce a PDF.")
    pdftoppm = shutil.which("pdftoppm") or shutil.which("pdftoppm.exe")
    if not pdftoppm:
        raise RuntimeError(
            "LibreOffice fallback requires pdftoppm (poppler) on PATH."
        )
    subprocess.run(
        [pdftoppm, "-png", "-r", "150", str(pdf), str(out_dir / "slide")],
        check=True,
    )
    return sorted(out_dir.glob("slide-*.png"))
