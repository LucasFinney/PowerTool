"""PowerPoint-native exporter: bake `advTm` into the deck, then ask
PowerPoint to render the video itself.

This preserves animations and slide transitions, which the flat exporter
can't. The trade-off is speed (PowerPoint's encoder is much slower than
ffmpeg) and the hard dependency on a Windows machine with PowerPoint
installed.

Mechanics:
  1. Compute per-slide durations from cue times.
  2. Copy the .pptx and write `advTm` into each slide via timings.py.
  3. Open the copy in PowerPoint via COM, call Presentation.CreateVideo(...)
     with UseTimingsAndNarrations=True.
  4. CreateVideo is asynchronous — poll Presentation.CreateVideoStatus
     until it reports done.

CreateVideoStatus codes (ppMediaTaskStatus):
  0 = None, 1 = InProgress, 2 = Failed, 3 = Done.
"""

from __future__ import annotations

import tempfile
import time
from pathlib import Path
from typing import Callable

from pptx import Presentation as _PptxOpen  # type: ignore

from .. import assemble, timings


_STATUS_NONE = 0
_STATUS_IN_PROGRESS = 1
_STATUS_FAILED = 2
_STATUS_DONE = 3


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
    quality: int = 85,
    poll_interval_s: float = 2.0,
) -> Path:
    n_slides = _count_slides(pptx_path)
    durations = assemble.compute_durations(
        cue_times_s, total_duration_s, n_slides
    )

    out_path = Path(out_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="slidetool_ppt_") as td:
        on_progress("Writing slide timings into deck copy...")
        deck_copy = timings.write_slide_durations(
            pptx_path, durations, Path(td) / "with_timings.pptx"
        )
        on_progress("Launching PowerPoint for export...")
        _create_video_via_com(
            deck_copy,
            out_path,
            on_progress,
            width=width,
            height=height,
            fps=fps,
            quality=quality,
            poll_interval_s=poll_interval_s,
        )

    return out_path


def _count_slides(pptx_path: Path) -> int:
    prs = _PptxOpen(str(pptx_path))
    return len(prs.slides)


def _create_video_via_com(
    deck_path: Path,
    out_path: Path,
    on_progress: Callable[[str], None],
    *,
    width: int,
    height: int,
    fps: int,
    quality: int,
    poll_interval_s: float,
) -> None:
    try:
        import comtypes.client  # type: ignore
    except ImportError as e:
        raise RuntimeError(
            "PowerPoint-native export requires comtypes. "
            "Install with: pip install comtypes"
        ) from e

    try:
        ppt = comtypes.client.CreateObject("PowerPoint.Application")
    except Exception as e:
        raise RuntimeError(
            f"Could not start PowerPoint via COM: {e}. "
            "PowerPoint must be installed on this machine."
        ) from e

    presentation = None
    try:
        presentation = ppt.Presentations.Open(
            str(deck_path.resolve()),
            ReadOnly=False,
            Untitled=False,
            WithWindow=False,
        )
        # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration,
        #             VertResolution, FramesPerSecond, Quality)
        # DefaultSlideDuration is only used for slides without an advTm;
        # we set advTm on every slide so it's effectively a fallback.
        presentation.CreateVideo(
            str(out_path),
            True,           # UseTimingsAndNarrations
            5,              # DefaultSlideDuration (s) — unused in practice
            height,         # VertResolution
            fps,            # FramesPerSecond
            quality,        # Quality 0-100
        )

        last_msg = ""
        while True:
            status = int(presentation.CreateVideoStatus)
            if status == _STATUS_DONE:
                on_progress("PowerPoint export complete.")
                break
            if status == _STATUS_FAILED:
                raise RuntimeError(
                    "PowerPoint CreateVideo reported FAILED status."
                )
            msg = (
                "Encoding via PowerPoint..."
                if status == _STATUS_IN_PROGRESS
                else "Waiting for PowerPoint to start encoding..."
            )
            if msg != last_msg:
                on_progress(msg)
                last_msg = msg
            time.sleep(poll_interval_s)
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

    if not out_path.exists():
        raise RuntimeError(
            f"PowerPoint reported done but {out_path} was not written."
        )
