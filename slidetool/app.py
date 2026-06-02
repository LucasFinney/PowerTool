"""Top-level orchestration: cues + slides + assemble -> one MP4."""

from __future__ import annotations

import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from . import assemble, cues, slides


@dataclass
class BuildRequest:
    pptx_path: Path
    markers_csv: Path
    fps: float = 24.0
    total_duration_s: float | None = None
    reference_media: Path | None = None
    width: int = 1920
    height: int = 1080
    out_path: Path | None = None


@dataclass
class BuildResult:
    out_path: Path
    n_slides: int
    n_cues: int
    total_duration_s: float


def run_build(
    req: BuildRequest,
    on_progress: Callable[[str], None] = lambda _msg: None,
) -> BuildResult:
    on_progress("Parsing markers...")
    cue_times = cues.parse_resolve_markers(req.markers_csv, fps=req.fps)

    total = req.total_duration_s
    if total is None:
        if req.reference_media is None:
            raise ValueError(
                "Need either total_duration_s or a reference media file."
            )
        on_progress("Probing reference media duration...")
        total = assemble.probe_duration(req.reference_media)

    out_path = req.out_path or req.pptx_path.with_suffix(".synced.mp4")

    with tempfile.TemporaryDirectory(prefix="slidetool_") as td:
        on_progress("Rasterizing slides...")
        pngs = slides.rasterize_pptx(
            req.pptx_path, td, width=req.width, height=req.height
        )
        on_progress(
            f"Assembling video ({len(pngs)} slides, {len(cue_times)} cues)..."
        )
        assemble.build_video(pngs, cue_times, total, out_path)

    on_progress(f"Done: {out_path}")
    return BuildResult(
        out_path=out_path,
        n_slides=len(pngs),
        n_cues=len(cue_times),
        total_duration_s=total,
    )


def main() -> None:
    """Entry point: launches the GUI."""
    from .gui import launch
    launch()


if __name__ == "__main__":
    main()
