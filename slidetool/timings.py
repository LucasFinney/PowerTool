"""Write per-slide advance timings (`advTm`) into a .pptx.

PowerPoint stores per-slide auto-advance timings inside each slide's
`<p:transition>` element:

    <p:transition spd="med" advClick="0" advTm="5000">
      <p:fade/>      <!-- optional transition effect -->
    </p:transition>

`advTm` is in **milliseconds**. `advClick="0"` disables the "advance on
click" requirement so the slide auto-advances during video export.

If a slide already has a `<p:transition>` (because the SME picked a fade,
push, etc.), we preserve its child effect element and just update the
`advClick`/`advTm` attributes. Otherwise we insert a new transition.

Element order inside `<p:sld>` matters in OOXML — `<p:transition>` must
come after `<p:clrMapOvr>` and before `<p:timing>`. python-pptx exposes the
underlying lxml tree, so we manipulate it directly.
"""

from __future__ import annotations

import shutil
from pathlib import Path

from pptx import Presentation  # type: ignore
from pptx.oxml.ns import qn  # type: ignore


def write_slide_durations(
    src_pptx: Path,
    durations_s: list[float],
    dst_pptx: Path,
) -> Path:
    """Copy src_pptx to dst_pptx and write each slide's advance timing.

    `durations_s[i]` is how long slide i+1 should display, in seconds.
    Must have one entry per slide in the deck.
    """
    src_pptx = Path(src_pptx)
    dst_pptx = Path(dst_pptx)
    dst_pptx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src_pptx, dst_pptx)

    prs = Presentation(str(dst_pptx))
    slides = list(prs.slides)
    if len(durations_s) != len(slides):
        raise ValueError(
            f"durations_s has {len(durations_s)} entries but deck has "
            f"{len(slides)} slides."
        )

    for slide, dur_s in zip(slides, durations_s):
        ms = max(int(round(dur_s * 1000)), 1)
        _set_slide_advance(slide.element, ms)

    prs.save(str(dst_pptx))
    return dst_pptx


def _set_slide_advance(sld_element, adv_tm_ms: int) -> None:
    """Set or insert the <p:transition> on a <p:sld> with advTm=ms."""
    transition = sld_element.find(qn("p:transition"))
    if transition is None:
        transition = _make_transition_element(adv_tm_ms)
        _insert_transition(sld_element, transition)
    else:
        transition.set("advClick", "0")
        transition.set("advTm", str(adv_tm_ms))


def _make_transition_element(adv_tm_ms: int):
    from lxml import etree  # python-pptx ships lxml as a dep

    nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    transition = etree.SubElement(
        etree.Element("placeholder", nsmap=nsmap),
        qn("p:transition"),
    )
    transition.set("spd", "med")
    transition.set("advClick", "0")
    transition.set("advTm", str(adv_tm_ms))
    # Detach from the placeholder parent before returning.
    transition.getparent().remove(transition)
    return transition


def _insert_transition(sld_element, transition) -> None:
    """Insert <p:transition> in its OOXML-valid position.

    Schema order inside <p:sld>: cSld, clrMapOvr, transition, timing, extLst.
    We insert immediately after clrMapOvr if present, else after cSld.
    """
    existing_timing = sld_element.find(qn("p:timing"))
    if existing_timing is not None:
        existing_timing.addprevious(transition)
        return

    clr_map = sld_element.find(qn("p:clrMapOvr"))
    if clr_map is not None:
        clr_map.addnext(transition)
        return

    csld = sld_element.find(qn("p:cSld"))
    if csld is not None:
        csld.addnext(transition)
        return

    sld_element.append(transition)
