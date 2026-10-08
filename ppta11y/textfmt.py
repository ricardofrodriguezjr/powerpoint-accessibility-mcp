"""Copy a shape's text and its exact formatting onto another shape, so the copy renders the same.

Used when a visible heading becomes a real title placeholder (titles.promote_to_title) and when a layout
table becomes text boxes (tables.table_to_shapes). The order of assignments matters: line rules before
spacing, AutoSize last.
"""
from __future__ import annotations

from .com import safe

TF2_PROPS = ("WordWrap", "MarginLeft", "MarginRight", "MarginTop", "MarginBottom", "VerticalAnchor",
             "HorizontalAnchor", "Orientation")
PARA_RULES = ("LineRuleWithin", "LineRuleBefore", "LineRuleAfter")  # must precede the spacing values
PARA_PROPS = ("Alignment", "SpaceWithin", "SpaceBefore", "SpaceAfter", "FirstLineIndent", "LeftIndent",
              "RightIndent", "BaselineAlignment", "WordWrap", "FarEastLineBreakLevel", "HangingPunctuation",
              "TextDirection", "IndentLevel")
FONT_PROPS = ("Name", "NameAscii", "NameFarEast", "NameComplexScript", "NameOther", "Size", "Bold", "Italic",
              "Allcaps", "Smallcaps", "Spacing", "Kerning", "BaselineOffset", "UnderlineStyle", "Strike", "Superscript",
              "Subscript")
BULLET_PROPS = ("Visible", "Type", "Style", "Character", "RelativeSize", "UseTextColor", "UseTextFont", "StartValue")


def _copy_attrs(src, dst, names, log: list, label: str) -> None:
    for k in names:
        v = safe(lambda: getattr(src, k))
        if v is None or v == -2147483648:
            continue
        cur = safe(lambda: getattr(dst, k))
        if cur == v:
            continue
        try:
            setattr(dst, k, v)
        except Exception as exc:  # noqa: BLE001 - recorded, verified by render afterwards
            log.append(f"{label}.{k}: {str(exc)[:80]}")


def _items(coll):
    return [coll.Item(i) for i in range(1, coll.Count + 1)]


MIXED = -2147483648


def color_state(cf) -> dict | None:
    """Everything that decides a ColorFormat's rendered colour: theme colour + brightness/tint, or RGB."""
    rgb = safe(lambda: cf.RGB)
    if not isinstance(rgb, int) or rgb in (MIXED,) or rgb < 0:
        return None
    return {"theme": safe(lambda: cf.ObjectThemeColor, 0) or 0, "brightness": safe(lambda: cf.Brightness, 0) or 0,
            "tint": safe(lambda: cf.TintAndShade, 0) or 0, "rgb": rgb}


def apply_color_state(cf, st: dict | None) -> None:
    """Write a colour_state back: theme colour (then brightness) when it was a theme colour, else RGB; then
    tint/shade. Theme tints such as 'Text 1, Lighter 25%' are kept instead of collapsing to the base colour."""
    if not st:
        return
    if st["theme"] and st["theme"] > 0:
        cf.ObjectThemeColor = st["theme"]
        if st["brightness"]:
            safe(lambda: setattr(cf, "Brightness", st["brightness"]))
    else:
        cf.RGB = st["rgb"]
    if st["tint"]:
        safe(lambda: setattr(cf, "TintAndShade", st["tint"]))


def copy_color(src_cf, dst_cf) -> None:
    st = color_state(src_cf)
    if st is None:
        return
    if color_state(dst_cf) == st:
        return
    apply_color_state(dst_cf, st)


def copy_shape_format(src, dst) -> None:
    """Fill, outline, shadow and other shape effects (PowerPoint's Format Painter)."""
    src.PickUp()
    dst.Apply()


def copy_text(src, dst, log: list | None = None, frame: bool = True) -> list[str]:
    """Replace dst's text with src's text, with every paragraph, run and text-frame setting copied.
    `src`/`dst` are Shapes (or table cell Shapes). Returns the list of settings PowerPoint refused."""
    log = [] if log is None else log
    s2, d2 = src.TextFrame2, dst.TextFrame2
    d2.TextRange.Text = s2.TextRange.Text
    if frame:
        _copy_attrs(s2, d2, TF2_PROPS, log, "frame")
        safe(lambda: setattr(d2.Column, "Number", s2.Column.Number))
    sp, dp = _items(s2.TextRange.Paragraphs), _items(d2.TextRange.Paragraphs)
    for i, (a, b) in enumerate(zip(sp, dp), 1):
        _copy_attrs(a.ParagraphFormat, b.ParagraphFormat, PARA_RULES, log, f"p{i}")
        _copy_attrs(a.ParagraphFormat, b.ParagraphFormat, PARA_PROPS, log, f"p{i}")
        _copy_attrs(a.ParagraphFormat.Bullet, b.ParagraphFormat.Bullet, BULLET_PROPS, log, f"p{i}.bullet")
        if safe(lambda: a.ParagraphFormat.Bullet.Visible) == -1:
            safe(lambda: setattr(b.ParagraphFormat.Bullet.Font, "Name", a.ParagraphFormat.Bullet.Font.Name))
            safe(lambda: copy_color(a.ParagraphFormat.Bullet.Font.Fill.ForeColor, b.ParagraphFormat.Bullet.Font.Fill.ForeColor))
        lang = safe(lambda: a.LanguageID)
        if lang not in (None, MIXED, -2):  # language of parts (WCAG 3.1.2)
            safe(lambda: setattr(b, "LanguageID", lang))
    # runs: same character positions in both (same text)
    for r in _items(s2.TextRange.Runs):
        start, length = r.Start, r.Length
        if length <= 0:
            continue
        t = d2.TextRange.GetCharacters(start, length)
        _copy_attrs(r.Font, t.Font, FONT_PROPS, log, f"run@{start}")
        lang = safe(lambda: r.LanguageID)
        if lang not in (None, MIXED, -2):
            safe(lambda: setattr(t, "LanguageID", lang))
            if safe(lambda: t.LanguageID) != lang:
                log.append(f"run@{start}.language not kept")
        f = r.Font.Fill
        if safe(lambda: f.Visible) == 0:
            safe(lambda: setattr(t.Font.Fill, "Visible", 0))
        else:
            try:
                copy_color(f.ForeColor, t.Font.Fill.ForeColor)
            except Exception as exc:  # noqa: BLE001
                log.append(f"run@{start}.color: {str(exc)[:80]}")
            tr_ = safe(lambda: f.Transparency, 0) or 0
            if tr_:
                safe(lambda: setattr(t.Font.Fill, "Transparency", tr_))
        if safe(lambda: r.Font.Highlight.Type, 0) not in (0, -2, None):
            safe(lambda: copy_color(r.Font.Highlight, t.Font.Highlight))
    _copy_links(src, dst, log)
    if frame:
        _copy_attrs(s2, d2, ("AutoSize",), log, "frame")
    return log


def _action_tuple(a) -> tuple | None:
    act = safe(lambda: a.Action, 0) or 0
    if not act:
        return None
    return (act, safe(lambda: a.Hyperlink.Address, "") or "", safe(lambda: a.Hyperlink.SubAddress, "") or "",
            safe(lambda: a.SlideShowName, "") or "", safe(lambda: a.Run, "") or "")


def actions(shape) -> list[tuple]:
    """Every click (1) / hover (2) action on the shape and on its text runs, as comparable tuples."""
    out = []
    for k in (1, 2):
        t = _action_tuple(safe(lambda k=k: shape.ActionSettings(k)))
        if t:
            out.append(("shape", k) + t)
    tr = safe(lambda: shape.TextFrame.TextRange)
    if tr is not None and tr.Text:
        for i in range(1, tr.Runs().Count + 1):
            r = tr.Runs(i)
            for k in (1, 2):
                t = _action_tuple(safe(lambda k=k: r.ActionSettings(k)))
                if t:
                    out.append(("text", k, r.Text) + t)
    return out


def _apply_action(dst_settings, a, log: list, label: str) -> None:
    try:
        if a.Action == 7 or (safe(lambda: a.Hyperlink.Address, "") or safe(lambda: a.Hyperlink.SubAddress, "")):
            dst_settings.Hyperlink.Address = a.Hyperlink.Address or ""
            if a.Hyperlink.SubAddress:
                dst_settings.Hyperlink.SubAddress = a.Hyperlink.SubAddress
            if a.Hyperlink.ScreenTip:
                dst_settings.Hyperlink.ScreenTip = a.Hyperlink.ScreenTip
        else:
            dst_settings.Action = a.Action
            for k in ("SlideShowName", "Run"):
                v = safe(lambda k=k: getattr(a, k))
                if v:
                    setattr(dst_settings, k, v)
    except Exception as exc:  # noqa: BLE001
        log.append(f"{label}: {str(exc)[:80]}")


def copy_actions(src, dst, log: list) -> None:
    """Shape-level click/hover actions, and run-level non-hyperlink and hover actions (run-level click
    hyperlinks are copied by copy_text)."""
    for k in (1, 2):
        a = safe(lambda k=k: src.ActionSettings(k))
        if a is not None and safe(lambda: a.Action, 0):
            _apply_action(dst.ActionSettings(k), a, log, f"shape action {k}")
    st, dt = safe(lambda: src.TextFrame.TextRange), safe(lambda: dst.TextFrame.TextRange)
    if st is None or dt is None or not st.Text:
        return
    for i in range(1, st.Runs().Count + 1):
        r = st.Runs(i)
        for k in (1, 2):
            a = r.ActionSettings(k)
            act = safe(lambda: a.Action, 0)
            if not act or (k == 1 and act == 7):
                continue
            _apply_action(dt.Characters(r.Start, r.Length).ActionSettings(k), a, log, f"run@{r.Start} action {k}")


def _copy_links(src, dst, log: list) -> None:
    """Hyperlinks live on character runs of the legacy TextRange (TextRange2 has no ActionSettings)."""
    st, dt = src.TextFrame.TextRange, dst.TextFrame.TextRange
    for i in range(1, st.Runs().Count + 1):
        r = st.Runs(i)
        a = r.ActionSettings(1)
        if safe(lambda: a.Action) != 7:  # ppActionHyperlink
            continue
        tgt = dt.Characters(r.Start, r.Length).ActionSettings(1)
        try:
            tgt.Hyperlink.Address = a.Hyperlink.Address or ""
            if a.Hyperlink.SubAddress:
                tgt.Hyperlink.SubAddress = a.Hyperlink.SubAddress
            if a.Hyperlink.ScreenTip:
                tgt.Hyperlink.ScreenTip = a.Hyperlink.ScreenTip
        except Exception as exc:  # noqa: BLE001
            log.append(f"link@{r.Start}: {str(exc)[:80]}")
