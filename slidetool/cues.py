"""Cue-timestamp sources for SlideTool.

Each function returns a list of cue times in seconds — the moments at which
the slideshow should advance to the next slide. cue_times[0] is when slide 2
appears (slide 1 starts at t=0). The downstream assembler decides when the
final slide ends, from a separately-supplied total duration.

This module is the swap point for Tier 2: any function returning list[float]
is interchangeable.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path


_TC_RE = re.compile(r"^\s*(\d+):(\d{2}):(\d{2})[:;](\d+)\s*$")


def timecode_to_seconds(tc: str, fps: float) -> float:
    """Convert HH:MM:SS:FF (or HH:MM:SS;FF for drop-frame) to seconds.

    Drop-frame is treated as non-drop here — close enough for cue placement
    at typical project rates. If precision matters we can revisit.
    """
    m = _TC_RE.match(tc)
    if not m:
        raise ValueError(f"Not a timecode: {tc!r}")
    h, mn, s, f = (int(x) for x in m.groups())
    return h * 3600 + mn * 60 + s + f / fps


def parse_resolve_markers(path: str | Path, fps: float = 24.0) -> list[float]:
    """Parse a DaVinci Resolve Edit Index CSV export.

    Resolve's Edit Index CSV has a header row; the "Record In" column holds
    the timeline timecode for each marker. We accept a few likely column
    names to be tolerant of Resolve version differences.
    """
    path = Path(path)
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError(f"No header row in {path}")

        tc_field = _pick_field(
            reader.fieldnames,
            ("Record In", "Source In", "Timecode", "TC In", "Marker In"),
        )

        times: list[float] = []
        for row in reader:
            tc = (row.get(tc_field) or "").strip()
            if not tc:
                continue
            try:
                times.append(timecode_to_seconds(tc, fps))
            except ValueError:
                continue  # skip non-timecode rows (totals, notes, etc.)

    times.sort()
    return times


def _pick_field(fields: list[str], candidates: tuple[str, ...]) -> str:
    lowered = {f.lower(): f for f in fields}
    for c in candidates:
        if c.lower() in lowered:
            return lowered[c.lower()]
    raise ValueError(
        f"No timecode column found. Looked for {candidates}, got {fields}"
    )
