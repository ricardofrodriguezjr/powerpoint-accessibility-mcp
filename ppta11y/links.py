"""Hyperlink text (WCAG 2.4.4 Link Purpose, A). A raw URL is "generally not sufficiently descriptive" (G91),
and fails when nothing in the same paragraph explains it (F63); PowerPoint's checker flags long/complex
raw URLs as "Unclear hyperlink text". The fix is descriptive visible text that keeps the link.

  link_report    every text hyperlink: slide, shape, text, address, and why it may be unclear; links with no
                 target (ppaction://noaction) are found exactly in the saved slide XML
  set_link_text  replace the visible text of one link, keep (or add) its address; verified in the deck
"""
from __future__ import annotations

import os
import re
import zipfile
import xml.etree.ElementTree as ET
from urllib.parse import unquote, urlparse

from .com import ToolError, com_err, safe
from .inventory import find_shape, get_part
from .session import checkpoint, require_editable, revert, save, save_copy, tmp_dir

URL_RX = re.compile(r"^\s*((https?|ftp)://|www\.)\S+\s*$", re.I)
BARE_DOMAIN_RX = re.compile(r"^\s*[\w-]+(\.[\w-]+)*\.(com|org|net|edu|gov|io|ai|co|uk|de|info|us|ca)(/\S*)?\s*$", re.I)
URL_TOKEN_RX = re.compile(r"((https?|ftp)://\S+|www\.\S+|\b[\w-]+(\.[\w-]+)*\.(com|org|net|edu|gov|io|ai|co|uk|de|info|us|ca)\b(/\S*)?)", re.I)
GENERIC = {"click here", "here", "link", "this link", "more", "read more", "learn more", "click", "go here",
           "see here", "this page", "website", "page"}
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def _is_url(t: str) -> bool:
    return bool(URL_RX.match(t) or BARE_DOMAIN_RX.match(t))


def _text_shapes(sl):
    """(shape id to address it by, shape with text, label, table cell or None) for every text frame: group
    children are addressed by their own id; merged table cells once."""
    for i in range(1, sl.Shapes.Count + 1):
        s = sl.Shapes(i)
        if s.Type == 6:
            for k in range(1, s.GroupItems.Count + 1):
                c = s.GroupItems(k)
                if safe(lambda: c.HasTextFrame) == -1:
                    yield c.Id, c, f"#{c.Id} in group #{s.Id}", None
        elif safe(lambda: s.HasTable) == -1:
            t = s.Table
            seen = set()
            for r in range(1, t.Rows.Count + 1):
                for c in range(1, t.Columns.Count + 1):
                    cs = t.Cell(r, c).Shape
                    key = (round(cs.Left, 1), round(cs.Top, 1), round(cs.Width, 1), round(cs.Height, 1))
                    if key in seen:
                        continue
                    seen.add(key)
                    yield s.Id, cs, f"#{s.Id} cell R{r}C{c}", [r, c]
        elif safe(lambda: s.HasTextFrame) == -1:
            yield s.Id, s, f"#{s.Id}", None


def _links_in(shape) -> list[dict]:
    """Linked character ranges of a text frame (runs with a hyperlink, merged when adjacent and equal)."""
    out = []
    tr = safe(lambda: shape.TextFrame.TextRange)
    if tr is None or not tr.Text:
        return out
    for i in range(1, tr.Runs().Count + 1):
        r = tr.Runs(i)
        a = r.ActionSettings(1)
        act = safe(lambda: a.Action, 0)
        addr = safe(lambda: a.Hyperlink.Address, "") or ""
        sub = safe(lambda: a.Hyperlink.SubAddress, "") or ""
        tip = safe(lambda: a.Hyperlink.ScreenTip, "") or ""
        if act != 7 and not addr and not sub:
            continue
        if out and out[-1]["start"] + out[-1]["length"] == r.Start and out[-1]["address"] == addr and out[-1]["sub"] == sub:
            out[-1]["length"] += r.Length
            out[-1]["text"] += r.Text
            continue
        out.append({"start": r.Start, "length": r.Length, "text": r.Text, "address": addr, "sub": sub, "tip": tip})
    return out


def _u16(t: str) -> int:
    """PowerPoint character positions count UTF-16 units (an emoji or math letter is 2)."""
    return len(t.encode("utf-16-le")) // 2


def _dead_runs(par) -> list[str]:
    """Texts of the dead links in one a:p, in order: adjacent dead runs joined, a line break inside one kept
    as PowerPoint's \x0b; anything else (a live run, a field, the paragraph end) ends it."""
    out, cur = [], None
    for ch in par:
        tag = ch.tag.rsplit("}", 1)[-1]
        if tag == "r":
            h = ch.find("a:rPr/a:hlinkClick", NS)
            t = ch.findtext("a:t", "", NS) or ""
            dead = h is not None and not h.get(f"{{{NS['r']}}}id") and \
                (h.get("action") in (None, "", "ppaction://noaction"))
            if dead:
                cur = (cur or "") + t
            elif cur is not None:
                out.append(cur)
                cur = None
        elif tag == "br":
            if cur is not None:
                cur += "\x0b"
        elif tag == "fld" and cur is not None:
            out.append(cur)
            cur = None
    if cur is not None:
        out.append(cur)
    return [x.strip("\x0b") for x in out if x.strip("\x0b").strip()]


def _dead_links_xml(pres) -> dict[int, list[dict]]:
    """{slide number: [{shape_id, cell, text}]} for text runs whose click link has no target
    (a:hlinkClick with an empty r:id and no action, or action="ppaction://noaction"), read from a saved copy
    (COM cannot tell such a run from plain text), paragraph by paragraph, in shapes, group children and
    table cells. Jumps, macros and programs are real actions, not dead."""
    from .animations import slide_part_name
    tmp = str(tmp_dir(pres) / f"links{os.getpid() % 10000}.pptx")
    out: dict[int, list[dict]] = {}
    try:
        save_copy(pres, tmp)
        with zipfile.ZipFile(tmp) as z:
            for n in range(1, pres.Slides.Count + 1):
                root = ET.fromstring(z.read(slide_part_name(z, pres.Slides(n).SlideID)))
                for sp in root.iter(f"{{{NS['p']}}}sp"):
                    cnv = sp.find("p:nvSpPr/p:cNvPr", NS)
                    sid = int(cnv.get("id")) if cnv is not None else None
                    for par in sp.findall("p:txBody/a:p", NS):
                        for t in _dead_runs(par):
                            out.setdefault(n, []).append({"shape_id": sid, "cell": None, "text": t})
                for gf in root.iter(f"{{{NS['p']}}}graphicFrame"):
                    tbl = gf.find(".//a:tbl", NS)
                    if tbl is None:
                        continue
                    cnv = gf.find("p:nvGraphicFramePr/p:cNvPr", NS)
                    sid = int(cnv.get("id")) if cnv is not None else None
                    for ri, row in enumerate(tbl.findall("a:tr", NS), 1):
                        for ci, tc in enumerate(row.findall("a:tc", NS), 1):
                            if tc.get("hMerge") or tc.get("vMerge"):
                                continue  # covered by a merged cell: its text is the merged cell's
                            for par in tc.findall("a:txBody/a:p", NS):
                                for t in _dead_runs(par):
                                    out.setdefault(n, []).append({"shape_id": sid, "cell": [ri, ci], "text": t})
    finally:
        safe(lambda: os.remove(tmp))
    return out


def _why(text: str, address: str, sub: str, paragraph: str) -> tuple[list[str], list[str]]:
    """(reasons it fails WCAG 2.4.4 / PowerPoint's check, advice only). A raw URL with explaining words in
    the same paragraph passes 2.4.4 in context (G53); one with nothing but other URLs/punctuation around it,
    or a long URL with a query string, does not (F63; PowerPoint flags it)."""
    t = text.strip()
    fail, advice = [], []
    rest = URL_TOKEN_RX.sub(" ", paragraph.replace(t, " "))
    alone = not re.search(r"[^\W\d_]{2,}", rest)  # no real word left once URLs/punctuation are removed
    if _is_url(t):
        if alone:
            fail.append("raw URL with nothing in its paragraph saying where it goes (WCAG 2.4.4, F63)")
        elif "?" in t or "&" in t:
            fail.append("long URL with a query string as link text (PowerPoint flags these as unclear)")
        else:
            advice.append("raw URL as link text; passes WCAG 2.4.4 through the surrounding sentence, descriptive "
                          "text would be clearer (2.4.9, AAA)")
    if t.lower().strip(".:!") in GENERIC:
        (fail if alone else advice).append("generic link text")
    if not address and not sub:
        fail.append("the link has no target (it does nothing when clicked)")
    return fail, advice


def suggest_text(url: str) -> str:
    """A starting point from the URL: site name + last path words. The agent should refine it to the
    destination page's real title."""
    u = urlparse(url if "://" in url else "https://" + url)
    host = (u.hostname or "").removeprefix("www.")
    words = [w for w in re.split(r"[/_\-.]+", unquote(u.path)) if w and not w.isdigit() and w.lower() not in
             ("html", "htm", "php", "aspx", "index", "en", "us", "en-us")]
    site = host.split(".")[-2].capitalize() if host.count(".") >= 1 else host
    if not words:
        return f"{site} website"
    return f"{' '.join(words[-3:]).replace('  ', ' ').title()} ({site})"


def _paragraph_of(tr, start: int) -> str:
    for p in range(1, tr.Paragraphs().Count + 1):
        pp = tr.Paragraphs(p)
        if pp.Start <= start < pp.Start + max(1, pp.Length):
            return pp.Text.strip("\r")
    return ""


def link_report(pres) -> dict:
    rows = []
    dead = _dead_links_xml(pres)
    for n in range(1, pres.Slides.Count + 1):
        sl = pres.Slides(n)
        holders = list(_text_shapes(sl))
        for sid, shape, where, cell in holders:
            tr = shape.TextFrame.TextRange
            for lk in _links_in(shape):
                why, advice = _why(lk["text"], lk["address"], lk["sub"], _paragraph_of(tr, lk["start"]))
                rows.append({"slide": n, "where": where, "shape_id": sid, "cell": cell, "start": lk["start"],
                             "text": lk["text"].strip(), "address": lk["address"], "sub_address": lk["sub"],
                             "screen_tip": lk["tip"], "unclear": bool(why), "why": why, "advice": advice,
                             **({"suggested_text": suggest_text(lk["address"] or lk["text"])} if why else {})})
        for d in dead.get(n, []):
            found = False
            for sid, shape, where, cell in holders:
                if (sid != d["shape_id"] and safe(lambda: shape.Id) != d["shape_id"]) or (cell or None) != d["cell"]:
                    continue
                full = safe(lambda: shape.TextFrame.TextRange.Text, "") or ""
                pos = full.find(d["text"])
                if pos < 0:
                    continue
                found = True
                start = _u16(full[:pos]) + 1
                if any(r["slide"] == n and r["shape_id"] == sid and r["cell"] == cell and r["start"] == start
                       for r in rows):
                    break  # PowerPoint reports it as a link already (with no address)
                why, advice = _why(d["text"], "", "", _paragraph_of(shape.TextFrame.TextRange, start))
                rows.append({"slide": n, "where": where, "shape_id": sid, "cell": cell, "start": start,
                             "text": d["text"].strip(), "address": "", "sub_address": "", "screen_tip": "",
                             "unclear": True, "why": why, "advice": advice, "dead": True,
                             "suggested_text": suggest_text(d["text"]) if _is_url(d["text"]) else d["text"].strip()})
                break
            if not found:  # never drop a dead link because its position could not be matched
                rows.append({"slide": n, "where": f"#{d['shape_id']}" + (f" cell R{d['cell'][0]}C{d['cell'][1]}"
                                                                         if d["cell"] else ""),
                             "shape_id": d["shape_id"], "cell": d["cell"], "start": None, "text": d["text"].strip(),
                             "address": "", "sub_address": "", "screen_tip": "", "unclear": True, "dead": True,
                             "why": ["the link has no target (it does nothing when clicked)"], "advice": [],
                             "note": "position not found in the live text (e.g. inside a group's table); fix by hand"})
    bad = [r for r in rows if r["unclear"]]
    return {"links": rows, "summary": {"links": len(rows), "unclear": len(bad),
                                       "unclear_at": [(r["slide"], r["where"]) for r in bad]}}


def _normalize(address: str) -> str:
    a = address.strip()
    if a and "://" not in a and not a.startswith(("mailto:", "#", "file:", "\\\\")) and \
            (a.lower().startswith("www.") or BARE_DOMAIN_RX.match(a)):
        return "https://" + a
    return a


def link_holder(sl, shape_id: int, cell=None, *texts: str):
    """The text frame a link lives in: the shape, a table cell, or - for a group id - the one piece holding
    any of `texts`."""
    s = find_shape(sl, shape_id)
    if s.Type == 6 and not cell:
        hits = [s.GroupItems(i) for i in range(1, s.GroupItems.Count + 1)
                if any(t and t in (safe(lambda i=i: s.GroupItems(i).TextFrame.TextRange.Text, "") or "") for t in texts)]
        if len(hits) != 1:
            raise ToolError(f"#{shape_id} is a group; pass the id of the piece holding the link")
        return hits[0]
    return s.Table.Cell(*cell).Shape if cell else s


def set_link_text(pres, slide: int, shape_id: int, old_text: str, new_text: str, address: str | None = None,
                  cell: list[int] | None = None, autosave: bool = True) -> dict:
    """Replace the visible link text `old_text` with `new_text` in shape `shape_id` (a group child by its own
    id, or a table by its id + cell [r, c]) and keep the link: same address/sub-address/ScreenTip, or
    `address` if given. A dead link whose text was a URL gets that URL as its address (https:// added when the
    URL had no scheme). Only those characters change; everything else in the text stays."""
    require_editable(pres)
    sl = get_part(pres, slide, "slide")
    holder = link_holder(sl, shape_id, cell, old_text)
    tr = holder.TextFrame.TextRange
    full = tr.Text
    pos = full.find(old_text)
    if pos < 0:
        raise ToolError(f"'{old_text[:60]}' not found in #{shape_id}")
    if full.find(old_text, pos + 1) >= 0:
        raise ToolError(f"'{old_text[:60]}' occurs more than once in #{shape_id}; pass a longer, unique old_text")
    if not new_text.strip():
        raise ToolError("new_text is empty")
    p16 = _u16(full[:pos]) + 1
    old = tr.Characters(p16, _u16(old_text))
    a = old.ActionSettings(1)
    keep = {"address": safe(lambda: a.Hyperlink.Address, "") or "", "sub": safe(lambda: a.Hyperlink.SubAddress, "") or "",
            "tip": safe(lambda: a.Hyperlink.ScreenTip, "") or ""}
    if address:
        keep["address"] = _normalize(address)
    elif not keep["address"] and not keep["sub"] and _is_url(old_text):
        keep["address"] = _normalize(old_text)
        keep["added_target"] = True
    if not keep["address"] and not keep["sub"]:
        raise ToolError("this text has no link target; pass address=")
    before, after = full[:pos], full[pos + len(old_text):]
    checkpoint(pres)
    result = {"slide": slide, "where": f"#{shape_id}" + (f" cell {cell}" if cell else ""), "old_text": old_text,
              "new_text": new_text, "address": keep["address"] or None, "sub_address": keep["sub"] or None}
    if keep.get("added_target"):
        result["note"] = "the link had no target; its URL text is now its address"
    failures = []
    try:
        old.Text = new_text
        rng = tr.Characters(p16, _u16(new_text))
        hl = rng.ActionSettings(1).Hyperlink
        if keep["address"]:
            hl.Address = keep["address"]
        if keep["sub"]:
            hl.SubAddress = keep["sub"]
        if keep["tip"]:
            hl.ScreenTip = keep["tip"]
        now = tr.Text
        if now != before + new_text + after:
            failures.append("text around the link changed")
        chk = tr.Characters(p16, _u16(new_text)).ActionSettings(1)
        if safe(lambda: chk.Action) != 7 or (keep["address"] and (safe(lambda: chk.Hyperlink.Address, "") or "") != keep["address"]):
            failures.append("the new text is not linked to the address")
        lk = _links_in(holder)
        if not any(l["text"].strip() == new_text.strip() for l in lk):
            failures.append("PowerPoint does not report the new text as one link")
    except Exception as exc:
        failures.append(f"{type(exc).__name__}: {com_err(exc)}")
    if failures:
        revert(pres)
        result.update(ok=False, reverted=True, failures=failures)
        return result
    if autosave:
        save(pres)
    result.update(ok=True, saved=autosave)
    return result
