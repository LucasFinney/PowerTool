"""Assemble PNG slides into a single MP4 with marker-driven slide durations.

Uses ffmpeg's concat demuxer: write a manifest listing each PNG with its
`duration`, then encode in one pass.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path


def probe_duration(media_path: str | Path) -> float:
    """Return media duration in seconds via ffprobe."""
    ffprobe = _require("ffprobe")
    out = subprocess.run(
        [ffprobe, "-v", "error", "-show_entries", "format=duration",
         "-of", "default=noprint_wrappers=1:nokey=1", str(media_path)],
        check=True, capture_output=True, text=True,
    )
    return float(out.stdout.strip())


def compute_durations(
    cue_times_s: list[float],
    total_duration_s: float,
    n_slides: int,
) -> list[float]:
    """Per-slide durations from cue times.

    Slide 1 starts at t=0 and ends at cue_times_s[0]. Slide k (1-indexed)
    ends at cue_times_s[k-1] for k < n_slides. The final slide ends at
    total_duration_s.

    If there are more cues than slide gaps (n_slides - 1), extras are ignored
    with a warning-style behavior — they're dropped silently here; the GUI
    surfaces the count mismatch. If there are fewer cues than gaps, the
    remaining slides each get an equal share of the leftover time.
    """
    if n_slides < 1:
        raise ValueError("Need at least one slide.")
    if total_duration_s <= 0:
        raise ValueError("total_duration_s must be positive.")

    cues = [t for t in cue_times_s if 0 < t < total_duration_s]
    cues.sort()
    cues = cues[: n_slides - 1]  # at most one cue per gap

    boundaries = [0.0, *cues, total_duration_s]

    # If we got fewer cues than slide gaps, pad evenly across the tail.
    while len(boundaries) - 1 < n_slides:
        tail_start = boundaries[-2]
        tail_end = boundaries[-1]
        gap = (tail_end - tail_start) / (n_slides - (len(boundaries) - 2))
        boundaries.insert(-1, tail_start + gap)

    durations = [boundaries[i + 1] - boundaries[i] for i in range(n_slides)]
    # Guard against zero/negative durations from coincident cues.
    durations = [max(d, 1.0 / 60) for d in durations]
    return durations


def build_video(
    slide_pngs: list[Path],
    cue_times_s: list[float],
    total_duration_s: float,
    out_path: str | Path,
    fps: int = 30,
) -> Path:
    """Build an MP4 where each slide PNG is held for its computed duration."""
    if not slide_pngs:
        raise ValueError("No slides to assemble.")
    out_path = Path(out_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    durations = compute_durations(
        cue_times_s, total_duration_s, len(slide_pngs)
    )

    ffmpeg = _require("ffmpeg")

    with tempfile.TemporaryDirectory() as td:
        manifest = Path(td) / "concat.txt"
        with manifest.open("w", encoding="utf-8") as f:
            for png, dur in zip(slide_pngs, durations):
                # ffmpeg concat demuxer wants forward slashes & escaped quotes.
                p = str(png.resolve()).replace("\\", "/").replace("'", r"'\''")
                f.write(f"file '{p}'\n")
                f.write(f"duration {dur:.6f}\n")
            # Concat-demuxer quirk: repeat the last file once without a
            # duration so its frames actually get emitted.
            last = str(slide_pngs[-1].resolve()).replace("\\", "/").replace("'", r"'\''")
            f.write(f"file '{last}'\n")

        cmd = [
            ffmpeg, "-y",
            "-f", "concat", "-safe", "0", "-i", str(manifest),
            "-vsync", "vfr",
            "-pix_fmt", "yuv420p",
            "-c:v", "libx264", "-preset", "medium", "-crf", "18",
            "-r", str(fps),
            "-movflags", "+faststart",
            str(out_path),
        ]
        subprocess.run(cmd, check=True)

    return out_path


def _require(tool: str) -> str:
    found = shutil.which(tool) or shutil.which(tool + ".exe")
    if not found:
        raise RuntimeError(
            f"{tool} not found on PATH. Install it or bundle it with SlideTool."
        )
    return found
