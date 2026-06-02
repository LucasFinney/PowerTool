# SlideTool

**Automate the markers-to-synced-slideshow step of voiceover-narrated video editing.**

SlideTool takes a PowerPoint deck, a list of slide-change cue points exported from your video editor, and a target total duration, and produces a single MP4 in which each slide is held for exactly the right amount of time. You drop that MP4 onto a video track above your voiceover in DaVinci Resolve (or any NLE) and the slideshow is already in sync — no screen recording, no speed ramps, no freezeframes.

## Why this exists

The original workflow:

1. SMEs author a PowerPoint deck as the outline for a video.
2. Voiceover is recorded against the deck.
3. In DaVinci Resolve, the editor listens through the voiceover with the deck open on a second screen and places a timeline marker at every point the slideshow should advance.
4. The editor then either screen-records the deck live in sync with the voiceover, or exports the deck as a video and hand-tunes speed ramps + freezeframes until the slides land on the markers.

Step 4 is purely mechanical given the markers — but doing it by hand takes time and is error-prone. SlideTool replaces it.

## How it works

```
   .pptx              markers CSV          reference audio (or typed duration)
     │                    │                          │
     ▼                    ▼                          ▼
 ┌────────┐         ┌──────────┐               ┌──────────┐
 │ slides │         │   cues   │               │ assemble │
 │ (PPTX  │         │ (Resolve │               │  (ffmpeg │
 │ → PNGs)│         │ CSV →    │               │  probe + │
 └────┬───┘         │  seconds)│               │  concat) │
      │             └────┬─────┘               └────┬─────┘
      │ slide PNGs       │ cue times                │ total duration
      └──────────────────┴───────────┬──────────────┘
                                     ▼
                             one synced MP4
```

Per-slide duration math: slide 1 starts at t=0 and ends at the first cue; slide *k* ends at cue *k*; the final slide ends at the total duration. ffmpeg's concat demuxer holds each PNG for its computed duration in a single encode.

## Project layout

```
PowerTool/                 ← repo root (and your working directory)
├── README.md              ← you are here
├── LICENSE
├── requirements.txt       ← comtypes on Windows; everything else is stdlib
├── __main__.py            ← convenience entry; same as `python -m slidetool`
└── slidetool/             ← the package
    ├── __init__.py
    ├── app.py             ← orchestration: BuildRequest, run_build()
    ├── gui.py             ← Tkinter window: file pickers + Run
    ├── cues.py            ← Resolve marker CSV → list[float] of cue seconds
    ├── slides.py          ← .pptx → one PNG per slide (PowerPoint COM, LibreOffice fallback)
    └── assemble.py        ← ffmpeg concat-demuxer build; ffprobe duration probe
```

### The Tier 2 seam

`cues.py` is the deliberate swap point in the architecture. Any function that returns `list[float]` of cue timestamps in seconds is interchangeable with the current `parse_resolve_markers`. The planned Tier 2 upgrade plugs in alongside it:

```python
# Tier 2 (future): transcript-driven cues, no manual markers needed
cues.transcribe_and_match(voiceover_audio, deck) -> list[float]
```

That function would run the voiceover through faster-whisper, pull a trigger phrase per slide from PowerPoint speaker notes, fuzzy-match each phrase against the transcript, and return the matched start timestamps. Same `list[float]` contract — the rest of the pipeline doesn't change.

## Workflow

1. **In DaVinci Resolve**, place a timeline marker at every point the slideshow should advance.
2. Open the **Edit Index** panel → right-click → **Export Edit Index** → CSV. (Any CSV with a `Record In`, `Source In`, `Timecode`, `TC In`, or `Marker In` column will work.)
3. **Run SlideTool** from inside the `PowerTool/` directory:
   ```
   python -m slidetool
   ```
4. **In the GUI**:
   - Pick the `.pptx`.
   - Pick the markers CSV.
   - Either pick a reference audio/video file (its duration sets when the last slide ends) **or** type a total duration in seconds.
   - Confirm the timeline FPS (default 24).
   - Hit **Run**.
5. The output MP4 lands next to the `.pptx` as `<deckname>.synced.mp4`. Drop it on a video track above your voiceover — the slide changes will already align with the cues.

### Headless / programmatic use

You don't have to use the GUI:

```python
from pathlib import Path
from slidetool.app import BuildRequest, run_build

result = run_build(BuildRequest(
    pptx_path=Path("deck.pptx"),
    markers_csv=Path("markers.csv"),
    fps=24.0,
    reference_media=Path("voiceover.wav"),  # or total_duration_s=583.4
))
print(result.out_path, result.n_slides, result.n_cues)
```

## Requirements

- **Windows + PowerPoint installed** — used for high-fidelity slide rasterization via COM (`comtypes`).
- **ffmpeg.exe and ffprobe.exe on PATH** — for video assembly and duration probing.
- **Python 3.10+**
- Install package deps with `pip install -r requirements.txt` (just `comtypes` on Windows; everything else is stdlib including the Tkinter GUI).

**Fallback**: if PowerPoint COM is unavailable, SlideTool tries LibreOffice (`soffice`) headless. The fallback also needs `pdftoppm` (poppler) on PATH because LibreOffice's direct PNG export only emits the first slide, so we go via PDF.

## Behavior notes & edge cases

- **Hard cuts only** in v1 — no transitions, animations, or per-slide builds. The final state of each slide is what gets rendered. This matches the existing manual workflow.
- **Cue/slide count mismatch**: if there are fewer cues than slide gaps (`n_slides - 1`), the remaining slides at the tail each get an equal share of leftover time, and the GUI shows a warning. If there are *more* cues than gaps, extras beyond `n_slides - 1` are silently dropped.
- **Coincident cues** (two cues at the same instant, or one within ~1/60s of another) are clamped so every slide gets at least one frame.
- **Timecode parsing** accepts both `HH:MM:SS:FF` (non-drop) and `HH:MM:SS;FF` (drop-frame). Drop-frame is treated as non-drop — close enough for cue placement at standard rates. If you ever need frame-accurate sync across very long durations at 29.97/59.94, this is where to revisit.
- **Output**: H.264 / yuv420p MP4 at CRF 18, 30 fps container framerate, faststart-flagged. Re-imports into Resolve without transcoding hiccups.

## Tiers / roadmap

- **Tier 1 (shipped)** — marker-driven assembler. What's in this repo.
- **Tier 2 (planned, not built)** — transcript-driven cues via faster-whisper + speaker-notes trigger phrases. Plugs into `cues.py`; nothing else changes. Eliminates the manual marker step.
- **Tier 3 (not planned)** — real-time presenter mode (streaming Whisper + PowerPoint COM advancing live slides). Significantly more fiddly; post-hoc is sufficient for the current need.

## Packaging

Once the pipeline is validated against a real episode, the intended distribution is a single `.exe` via PyInstaller so non-Python teammates can run it.
