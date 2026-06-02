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

## Output modes

SlideTool ships with two exporters. They consume the same inputs (deck + cues + total duration) and produce the same kind of output (a single MP4), but make different trade-offs:

### PowerPoint native (default) — preserves animations & transitions

SlideTool writes per-slide advance timings (`<p:transition advClick="0" advTm="…">`) directly into a copy of your deck, then drives PowerPoint via COM to call its built-in **Create a Video** function with `UseTimingsAndNarrations=True`. The resulting MP4 includes every slide transition and every within-slide animation, exactly as PowerPoint would render them if you'd recorded the timings by hand.

This is the right choice when SMEs have authored animations or transitions that need to appear in the final video.

**Trade-offs:**
- Requires Windows + PowerPoint installed.
- Significantly slower than the flat path — PowerPoint's encoder is minute-scale for a typical deck.
- PowerPoint's encoder is configurable only by quality preset (0–100), resolution, and frame rate. No codec choice or CRF.

### Flat slides — fast, no animations

SlideTool rasterizes each slide to a PNG (final state only — animations and transitions are gone) and uses ffmpeg's concat demuxer to hold each PNG for its computed duration in a single H.264 encode.

This is the right choice for fast previews, for decks with no animations, or as a fallback when PowerPoint isn't available.

**Trade-offs:**
- No animations, no transitions, no within-slide builds.
- Falls back to LibreOffice for rasterization if PowerPoint COM isn't available (also needs `pdftoppm`).

## How it works

```
   .pptx              markers CSV          reference audio (or typed duration)
     │                    │                          │
     ▼                    ▼                          ▼
              ┌────────────────────────┐
              │   cues.parse_resolve_  │
              │   markers              │
              └────────────┬───────────┘
                           │ cue times (seconds)
                           ▼
              ┌────────────────────────┐
              │  app.run_build picks   │
              │  output mode           │
              └─────┬──────────────┬───┘
                    │              │
            ┌───────▼──────┐  ┌────▼────────────┐
            │ exporters/   │  │ exporters/      │
            │ flat.py      │  │ powerpoint.py   │
            │ (rasterize + │  │ (timings.py     │
            │  ffmpeg      │  │  writes advTm,  │
            │  concat)     │  │  PowerPoint     │
            │              │  │  CreateVideo)   │
            └───────┬──────┘  └────────┬────────┘
                    └──────┬───────────┘
                           ▼
                     one synced MP4
```

Per-slide duration math (shared by both exporters): slide 1 starts at t=0 and ends at the first cue; slide *k* ends at cue *k*; the final slide ends at the total duration.

## Project layout

```
PowerTool/                     ← repo root (and your working directory)
├── README.md                  ← you are here
├── LICENSE
├── requirements.txt           ← python-pptx; comtypes on Windows
├── __main__.py                ← convenience entry; same as `python -m slidetool`
└── slidetool/                 ← the package
    ├── __init__.py
    ├── app.py                 ← orchestration: BuildRequest, run_build()
    ├── gui.py                 ← Tkinter window: pickers + output-mode radio + Run
    ├── cues.py                ← INPUT SEAM: marker CSV → list[float] of cue seconds
    ├── timings.py             ← OOXML writer: stamps advTm into a .pptx copy
    ├── slides.py              ← .pptx → one PNG per slide (PowerPoint COM, LibreOffice fallback)
    ├── assemble.py            ← shared utils: compute_durations, ffprobe, ffmpeg concat
    └── exporters/             ← OUTPUT SEAM: pluggable exporters with a common signature
        ├── __init__.py        ← MODES = {"flat": ..., "powerpoint": ...}
        ├── flat.py            ← rasterize + ffmpeg-concat path
        └── powerpoint.py      ← advTm + PowerPoint CreateVideo path
```

### Architectural seams (input and output)

SlideTool has two deliberate plug-in points:

- **`cues.py` — input seam.** Any function returning `list[float]` of cue timestamps in seconds is interchangeable with the current `parse_resolve_markers`. The planned Tier 2 upgrade plugs in alongside it:
  ```python
  cues.transcribe_and_match(voiceover_audio, deck) -> list[float]
  ```
  That function would run the voiceover through faster-whisper, pull a trigger phrase per slide from PowerPoint speaker notes, fuzzy-match each phrase against the transcript, and return the matched start timestamps. Same contract — the rest of the pipeline doesn't change.

- **`exporters/` — output seam.** Each exporter implements:
  ```python
  export(
      pptx_path, cue_times_s, total_duration_s, out_path, on_progress,
      *, width=1920, height=1080, fps=30,
  ) -> Path
  ```
  Adding a new exporter is a matter of registering it in `exporters/__init__.MODES` and adding a GUI radio option.

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
   - Pick an output mode: **PowerPoint native** (default — preserves animations) or **Flat slides** (fast preview).
   - Hit **Run**.
5. The output MP4 lands next to the `.pptx` as `<deckname>.synced.mp4`. Drop it on a video track above your voiceover.

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
    output_mode="powerpoint",               # or "flat"
))
print(result.out_path, result.n_cues, result.output_mode)
```

## Requirements

- **Windows + PowerPoint installed** — required for the PowerPoint-native exporter and recommended for the flat exporter's PPTX rasterization. Accessed via COM (`comtypes`).
- **`ffmpeg.exe` and `ffprobe.exe` on PATH** — needed for the flat exporter and for probing reference-media durations.
- **Python 3.10+**
- Install package deps with `pip install -r requirements.txt` (just `python-pptx` and, on Windows, `comtypes`; everything else is stdlib including the Tkinter GUI).

**Fallbacks:** if PowerPoint COM is unavailable, the flat exporter falls back to LibreOffice (`soffice`) headless. That fallback also needs `pdftoppm` (poppler) on PATH because LibreOffice's direct PNG export only emits the first slide, so the fallback goes via PDF. The PowerPoint-native exporter has no fallback — PowerPoint is required.

## Authoring rules for animated decks

When using the PowerPoint-native exporter, a few PowerPoint behaviors carry over from PowerPoint's own "Create a Video" feature and are worth knowing about. None are SlideTool bugs — they're consequences of how PowerPoint exports.

- **Within-slide animations play at their authored timing**, not stretched to fit the slide's `advTm`. If the animation sequence is shorter than the cue gap, the slide holds on its final state for the remainder (this is what you want). If it's *longer*, PowerPoint cuts the animation off when the slide advances. A future enhancement could warn when authored animation duration exceeds the cue gap.
- **"On Click" animation triggers will not auto-fire** in exported video. Animations need to be set to "Start: After Previous" or "Start: With Previous" (with their own delays) to appear. This is a one-time content-authoring rule for SMEs.
- **Slide transitions ARE preserved.** If a slide has, e.g., a Fade or Push transition, the exported video shows it.

## Behavior notes & edge cases

- **Cue/slide count mismatch**: if there are fewer cues than slide gaps (`n_slides - 1`), the remaining slides at the tail each get an equal share of leftover time. If there are *more* cues than gaps, extras beyond `n_slides - 1` are silently dropped.
- **Coincident cues** (two cues at the same instant, or one within ~1/60s of another) are clamped so every slide gets at least one frame.
- **Timecode parsing** accepts both `HH:MM:SS:FF` (non-drop) and `HH:MM:SS;FF` (drop-frame). Drop-frame is treated as non-drop — close enough for cue placement at standard rates. If you ever need frame-accurate sync across very long durations at 29.97/59.94, this is where to revisit.
- **Output (flat mode)**: H.264 / yuv420p MP4 at CRF 18, faststart-flagged. Container framerate defaults to 30 fps. Re-imports into Resolve without transcoding hiccups.
- **Output (PowerPoint mode)**: H.264 MP4 at the resolution / fps / quality you set. PowerPoint controls the encoder.

## Tiers / roadmap

- **Tier 1 (shipped)** — marker-driven assembler, two output modes (flat + PowerPoint-native).
- **Tier 2 (planned)** — transcript-driven cues via faster-whisper + speaker-notes trigger phrases. Plugs into `cues.py`; eliminates the manual marker step.
- **Tier 3 (not planned)** — real-time presenter mode. Post-hoc is sufficient for the current need.

## Packaging

Once the pipeline is validated against a real episode, the intended distribution is a single `.exe` via PyInstaller so non-Python teammates can run it.
