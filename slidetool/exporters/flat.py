"""Flat exporter: rasterize each slide to PNG, ffmpeg-concat into MP4.

Fast (seconds for typical decks). Loses animations and transitions — every
slide is rendered in its final state. Good for quick previews and as a
fallback when PowerPoint isn't available.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Callable

from .. import assemble, slides


def export(
    pptx_path: Path,
    cue_times_s: list[float],
    total_duration_s: float,
    out_path: Path,
    on_progress: Callable[[str], None] = lambda _m: None,
    *,
    width: int = 1920,
    height: int = 1080,
    fps: int = 30,
) -> Path:
    with tempfile.TemporaryDirectory(prefix="slidetool_flat_") as td:
        on_progress("Rasterizing slides...")
        pngs = slides.rasterize_pptx(pptx_path, td, width=width, height=height)
        on_progress(f"Encoding video ({len(pngs)} slides)...")
        assemble.build_video(
            pngs, cue_times_s, total_duration_s, out_path, fps=fps
        )
    return out_path
