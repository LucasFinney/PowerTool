"""Top-level orchestration: cues + chosen exporter -> one MP4."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Literal

from . import assemble, cues, exporters


OutputMode = Literal["flat", "powerpoint"]


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
    output_mode: OutputMode = "powerpoint"
    encode_fps: int = 30


@dataclass
class BuildResult:
    out_path: Path
    n_cues: int
    total_duration_s: float
    output_mode: OutputMode


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

    if req.output_mode not in exporters.MODES:
        raise ValueError(
            f"Unknown output_mode {req.output_mode!r}; "
            f"expected one of {list(exporters.MODES)}"
        )
    export_fn = exporters.MODES[req.output_mode]
    export_fn(
        req.pptx_path,
        cue_times,
        total,
        out_path,
        on_progress,
        width=req.width,
        height=req.height,
        fps=req.encode_fps,
    )

    on_progress(f"Done: {out_path}")
    return BuildResult(
        out_path=out_path,
        n_cues=len(cue_times),
        total_duration_s=total,
        output_mode=req.output_mode,
    )


def main() -> None:
    """Entry point: launches the GUI."""
    from .gui import launch
    launch()


if __name__ == "__main__":
    main()
