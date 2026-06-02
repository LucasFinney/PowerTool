"""Output-side seam.

Each exporter implements the same signature:

    export(
        pptx_path: Path,
        cue_times_s: list[float],
        total_duration_s: float,
        out_path: Path,
        on_progress: Callable[[str], None] = lambda _m: None,
        *,
        width: int = 1920,
        height: int = 1080,
        fps: int = 30,
    ) -> Path

So they're interchangeable from app.py's point of view.
"""

from . import flat, powerpoint

MODES = {
    "flat": flat.export,
    "powerpoint": powerpoint.export,
}
