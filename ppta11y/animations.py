"""Snapshot every animation effect on a slide and rebuild lost effects exactly.

Why this exists: PowerPoint deletes a shape's animation effects when the shape is put into a group
(Shapes.Range(..).Group(), same as Ctrl+G). It does, however, let you add effects to a shape that is
*inside* a group (Sequence.AddEffect(group.GroupItems(k), ...)), and member shapes keep their Shape.Id
when grouped. So: snapshot -> group -> re-add each lost effect on the same shape inside the group ->
copy every property -> restore the original order. The result is verified by comparing the slide's
<p:timing> XML that PowerPoint itself writes before and after.
"""
from __future__ import annotations

import copy
import re
import zipfile
import xml.etree.ElementTree as ET

from .com import TRUE, ToolError, com_err, safe

# MsoAnimTriggerType
ON_CLICK, WITH_PREV, AFTER_PREV, ON_SHAPE_CLICK, ON_BOOKMARK = 1, 2, 3, 4, 5
TRIGGER_NAMES = {1: "on click", 2: "with previous", 3: "after previous", 4: "on shape click", 5: "on bookmark"}

TIMING_PROPS = ("TriggerDelayTime", "Duration", "Accelerate", "Decelerate", "AutoReverse", "RepeatCount",
                "RepeatDuration", "Restart", "RewindAtEnd", "SmoothStart", "SmoothEnd", "BounceEnd",
                "BounceEndIntensity")
PARAM_PROPS = ("Direction", "Amount", "FontName", "Relative", "Size")
BEHAVIOR_PROPS = {
    1: ("MotionEffect", ("Path", "ByX", "ByY", "FromX", "FromY", "ToX", "ToY")),
    2: ("ColorEffect", ()),
    3: ("ScaleEffect", ("ByX", "ByY", "FromX", "FromY", "ToX", "ToY")),
    4: ("RotationEffect", ("By", "From", "To")),
    5: ("PropertyEffect", ("Property", "From", "To")),
    6: ("CommandEffect", ("Type", "Command")),
    7: ("FilterEffect", ("Type", "Subtype", "Reveal")),
    8: ("SetEffect", ("Property", "To")),
}


# --------------------------------------------------------------------------------------------------
# Snapshot
# --------------------------------------------------------------------------------------------------

def _num(v):
    return v  # full precision: values are written back, so never round them


def _timing(t) -> dict:
    d = {"TriggerType": safe(lambda: t.TriggerType)}
    for p in TIMING_PROPS:
        d[p] = _num(safe(lambda p=p: getattr(t, p)))
    d["TriggerShape"] = safe(lambda: t.TriggerShape.Id)
    if d["TriggerType"] == ON_BOOKMARK:
        d["TriggerBookmark"] = safe(lambda: t.TriggerBookmark)
    return d


def _behavior(b) -> dict:
    d = {"Type": safe(lambda: b.Type), "Additive": safe(lambda: b.Additive), "Accumulate": safe(lambda: b.Accumulate),
         "Timing": _timing(b.Timing)}
    spec = BEHAVIOR_PROPS.get(d["Type"])
    if spec:
        obj = safe(lambda: getattr(b, spec[0]))
        if obj is not None:
            d["props"] = {p: _num(safe(lambda p=p: getattr(obj, p))) for p in spec[1]}
            if d["Type"] == 2:
                d["props"] = {k: safe(lambda k=k: getattr(obj, k).RGB) for k in ("By", "From", "To")}
            if d["Type"] == 5:
                pts = safe(lambda: obj.Points)
                if pts is not None and safe(lambda: pts.Count, 0):
                    d["points"] = [(_num(pts.Item(i).Time), safe(lambda i=i: pts.Item(i).Value),
                                    safe(lambda i=i: pts.Item(i).Formula)) for i in range(1, pts.Count + 1)]
    return d


def snapshot_effect(e) -> dict:
    info = e.EffectInformation
    shape = e.Shape
    d = {
        "shape_id": shape.Id,
        "shape_name": shape.Name,
        "EffectType": e.EffectType,
        "Exit": e.Exit,
        "DisplayName": safe(lambda: e.DisplayName, ""),
        "Paragraph": safe(lambda: e.Paragraph, 0),
        "TextRangeStart": safe(lambda: e.TextRangeStart),
        "TextRangeLength": safe(lambda: e.TextRangeLength),
        "Timing": _timing(e.Timing),
        "params": {p: _num(safe(lambda p=p: getattr(e.EffectParameters, p))) for p in PARAM_PROPS},
        "Color2": safe(lambda: e.EffectParameters.Color2.RGB),
        "AfterEffect": safe(lambda: info.AfterEffect),
        "Dim": safe(lambda: info.Dim.RGB),
        "AnimateBackground": safe(lambda: info.AnimateBackground),
        "AnimateTextInReverse": safe(lambda: info.AnimateTextInReverse),
        "BuildByLevelEffect": safe(lambda: info.BuildByLevelEffect),
        "TextUnitEffect": safe(lambda: info.TextUnitEffect),
        "Sound": {"Type": safe(lambda: info.SoundEffect.Type), "Name": safe(lambda: info.SoundEffect.Name)},
        "behaviors": [_behavior(e.Behaviors.Item(i)) for i in range(1, e.Behaviors.Count + 1)],
    }
    return d


def _sequences(container):
    tl = container.TimeLine
    out = [("main", tl.MainSequence)]
    for k in range(1, tl.InteractiveSequences.Count + 1):
        out.append((f"interactive{k}", tl.InteractiveSequences(k)))
    return out


def snapshot(container) -> list[dict]:
    """Every sequence on the slide as [{'seq': 'main'|'interactiveN', 'trigger_shape': id|None, 'effects': [...]}]."""
    out = []
    tl = safe(lambda: container.TimeLine)
    if tl is None:
        return out
    seen: dict[tuple, int] = {}
    for name, q in _sequences(container):
        effects = [snapshot_effect(q.Item(i)) for i in range(1, q.Count + 1)]
        trig = effects[0]["Timing"]["TriggerShape"] if (name != "main" and effects) else None
        key = None
        if name != "main" and effects:
            t0 = effects[0]["Timing"]
            base = (trig, t0.get("TriggerType"), t0.get("TriggerBookmark") or "")
            seen[base] = seen.get(base, -1) + 1
            key = base + (seen[base],)
        out.append({"seq": name, "trigger_shape": trig, "seq_key": key, "effects": effects})
    return out


def _seq_key_live(q, ordinals: dict) -> tuple | None:
    if not q.Count:
        return None
    t = q.Item(1).Timing
    base = (safe(lambda: t.TriggerShape.Id), safe(lambda: t.TriggerType), safe(lambda: t.TriggerBookmark) or "")
    ordinals[base] = ordinals.get(base, -1) + 1
    return base + (ordinals[base],)


def signature(s: dict) -> tuple:
    """What identifies an effect for alignment purposes."""
    return (s["shape_id"], s["EffectType"], s["Exit"], s["Paragraph"])


def summarize(snap: list[dict]) -> list[dict]:
    rows = []
    for seq in snap:
        for i, e in enumerate(seq["effects"], 1):
            rows.append({
                "seq": seq["seq"] + (f" (trigger #{seq['trigger_shape']})" if seq["trigger_shape"] else ""),
                "order": i, "shape_id": e["shape_id"], "shape": e["shape_name"], "effect": e["DisplayName"],
                "type": e["EffectType"], "exit": bool(e["Exit"]),
                "start": TRIGGER_NAMES.get(e["Timing"]["TriggerType"], e["Timing"]["TriggerType"]),
                "delay": e["Timing"]["TriggerDelayTime"], "duration": e["Timing"]["Duration"],
                **({"paragraph": e["Paragraph"]} if e["Paragraph"] else {}),
            })
    return rows


# --------------------------------------------------------------------------------------------------
# Rebuild
# --------------------------------------------------------------------------------------------------

def _set_if_diff(obj, prop, want, failures, label):
    if want is None:
        return
    have = safe(lambda: getattr(obj, prop))
    if isinstance(want, float) and isinstance(have, float) and abs(want - have) < 1e-4:
        return
    if have == want:
        return
    try:
        setattr(obj, prop, want)
    except Exception as exc:
        failures.append(f"{label}.{prop}={want!r}: {com_err(exc)}")


def _apply_timing(t, snap_t: dict, failures: list, label: str):
    for p in TIMING_PROPS:
        _set_if_diff(t, p, snap_t.get(p), failures, label)
    if snap_t.get("TriggerBookmark"):
        _set_if_diff(t, "TriggerBookmark", snap_t["TriggerBookmark"], failures, label)


def _apply_behaviors(e, snap: dict, failures: list, label: str):
    """Preset effects get their behaviors from EffectType + EffectParameters, so only an edited motion
    path is copied for them. Custom effects (EffectType 0) get their behavior list rebuilt."""
    want = snap["behaviors"]
    bhv = e.Behaviors
    custom = snap["EffectType"] == 0
    have_types = [safe(lambda i=i: bhv.Item(i).Type) for i in range(1, bhv.Count + 1)]
    want_types = [b["Type"] for b in want]
    if have_types != want_types:
        if not custom:
            failures.append(f"{label}: behavior types differ after AddEffect {have_types} != {want_types}")
            return
        for i in range(bhv.Count, 0, -1):
            safe(lambda i=i: bhv.Item(i).Delete())
        for b in want:
            try:
                bhv.Add(b["Type"])
            except Exception as exc:
                failures.append(f"{label}: add behavior {b['Type']}: {com_err(exc)}")
    for i, b in enumerate(want, 1):
        if i > bhv.Count:
            break
        ob = bhv.Item(i)
        blabel = f"{label}.behavior{i}"
        props = b.get("props") or {}
        if b["Type"] == 1 and props.get("Path"):  # motion path: copy the drawn path
            _set_if_diff(ob.MotionEffect, "Path", props["Path"], failures, blabel)
        if not custom:
            continue
        _set_if_diff(ob, "Additive", b.get("Additive"), failures, blabel)
        _set_if_diff(ob, "Accumulate", b.get("Accumulate"), failures, blabel)
        _apply_timing(ob.Timing, b["Timing"], failures, blabel)
        spec = BEHAVIOR_PROPS.get(b["Type"])
        sub = safe(lambda: getattr(ob, spec[0])) if spec else None
        if sub is None or not props:
            continue
        if b["Type"] == 2:
            for k, rgb in props.items():
                cf = safe(lambda k=k: getattr(sub, k))
                if rgb is not None and cf is not None and safe(lambda: cf.RGB) != rgb:
                    try:
                        cf.RGB = rgb
                    except Exception as exc:
                        failures.append(f"{blabel}.{k}: {com_err(exc)}")
            continue
        for p, v in props.items():
            if p != "Path":
                _set_if_diff(sub, p, v, failures, blabel)
        if b.get("points"):
            pts = sub.Points
            for i2 in range(pts.Count, 0, -1):
                safe(lambda i2=i2: pts.Item(i2).Delete())
            for tm, val, formula in b["points"]:
                try:
                    pt = pts.Add()
                    pt.Time, pt.Value = tm, val
                    if formula:
                        pt.Formula = formula
                except Exception as exc:
                    failures.append(f"{blabel}.points: {com_err(exc)}")


def apply_properties(seq, e, snap: dict, label: str) -> tuple[object, list[str]]:
    """Copy every recorded property onto effect `e`. Returns (effect, failures); the Convert* calls
    return a *new* Effect object, so always use the returned one."""
    failures: list[str] = []
    trig = snap["Timing"].get("TriggerType")
    if trig not in (None, ON_SHAPE_CLICK, ON_BOOKMARK):
        _set_if_diff(e.Timing, "TriggerType", trig, failures, label)
    _set_if_diff(e, "Exit", snap["Exit"], failures, label)
    if snap["Paragraph"]:
        _set_if_diff(e, "Paragraph", snap["Paragraph"], failures, label)

    info = e.EffectInformation
    lvl = snap.get("BuildByLevelEffect")
    if lvl is not None and safe(lambda: info.BuildByLevelEffect) not in (lvl, None) and not snap["Paragraph"]:
        try:
            e = seq.ConvertToBuildLevel(e, lvl)
        except Exception as exc:
            failures.append(f"{label}.BuildByLevel={lvl}: {com_err(exc)}")
    tu = snap.get("TextUnitEffect")
    if tu is not None and safe(lambda: e.EffectInformation.TextUnitEffect) not in (tu, None):
        try:
            e = seq.ConvertToTextUnitEffect(e, tu)
        except Exception as exc:
            failures.append(f"{label}.TextUnitEffect={tu}: {com_err(exc)}")
    ab = snap.get("AnimateBackground")
    if ab is not None and safe(lambda: e.EffectInformation.AnimateBackground) not in (ab, None):
        try:
            e = seq.ConvertToAnimateBackground(e, ab)
        except Exception as exc:
            failures.append(f"{label}.AnimateBackground: {com_err(exc)}")
    rev = snap.get("AnimateTextInReverse")
    if rev is not None and safe(lambda: e.EffectInformation.AnimateTextInReverse) not in (rev, None):
        try:
            e = seq.ConvertToAnimateInReverse(e, rev)
        except Exception as exc:
            failures.append(f"{label}.AnimateTextInReverse: {com_err(exc)}")
    after = snap.get("AfterEffect")
    if after not in (None, -2) and safe(lambda: e.EffectInformation.AfterEffect) != after:
        try:
            if after == 1 and snap.get("Dim") is not None:  # msoAnimAfterEffectDim
                e = seq.ConvertToAfterEffect(e, after, snap["Dim"])
            else:
                e = seq.ConvertToAfterEffect(e, after)
        except Exception as exc:
            failures.append(f"{label}.AfterEffect={after}: {com_err(exc)}")

    for p, v in snap["params"].items():
        _set_if_diff(e.EffectParameters, p, v, failures, label)
    if snap.get("Color2") is not None:
        cf = safe(lambda: e.EffectParameters.Color2)
        if cf is not None and safe(lambda: cf.RGB) != snap["Color2"]:
            try:
                cf.RGB = snap["Color2"]
            except Exception as exc:
                failures.append(f"{label}.Color2: {com_err(exc)}")

    _apply_behaviors(e, snap, failures, label)
    _apply_timing(e.Timing, snap["Timing"], failures, label)  # after behaviors: effect timing wins
    if trig not in (None, ON_SHAPE_CLICK, ON_BOOKMARK):  # Convert* may have rebuilt the effect
        _set_if_diff(e.Timing, "TriggerType", trig, failures, label)

    snd = snap.get("Sound") or {}
    if snd.get("Type") not in (None, 0) and safe(lambda: e.EffectInformation.SoundEffect.Type) != snd["Type"]:
        try:
            se = e.EffectInformation.SoundEffect
            if snd["Type"] == 1:  # ppSoundStopPrevious
                se.Type = 1
            else:  # a named sound; works for PowerPoint's built-in sounds only
                se.Name = snd["Name"]
        except Exception as exc:
            failures.append(f"{label}.Sound {snd}: {com_err(exc)} (custom sound files cannot be copied via COM)")
    return e, failures


def _align(before: list[dict], after_q) -> list:
    """Map each `before` effect to the surviving Effect object in `after_q` (or None if it was lost).
    Survivors keep their relative order, so a greedy left-to-right match is exact."""
    post = [after_q.Item(i) for i in range(1, after_q.Count + 1)]
    post_sig = [(safe(lambda e=e: e.Shape.Id), e.EffectType, e.Exit, safe(lambda e=e: e.Paragraph, 0)) for e in post]
    mapping, j = [], 0
    for s in before:
        if j < len(post) and post_sig[j] == signature(s):
            mapping.append(post[j])
            j += 1
        else:
            mapping.append(None)
    if j != len(post):
        raise ToolError(f"could not align {len(post)} remaining effects with the snapshot "
                        f"({j} matched); not rebuilding")
    return mapping


def _set_trigger_shape(e, seq_snap, resolve_shape, report, label):
    try:
        e.Timing.TriggerShape = resolve_shape(seq_snap["trigger_shape"])
    except Exception as exc:
        report["property_failures"].append(f"{label}.TriggerShape: {com_err(exc)}")


def build_groups_from_pptx(pptx_path: str, slide_id: int) -> dict:
    """{'main': [grpId per effect in order], 'interactive1': [...], ...} from a saved deck."""
    with zipfile.ZipFile(pptx_path) as z:
        xml = z.read(slide_part_name(z, slide_id))
    t = next(ET.fromstring(xml).iter(P + "timing"), None)
    out = {}
    if t is None:
        return out
    k = 0
    for seq in t.iter(P + "seq"):
        ctn = seq.find(P + "cTn")
        if ctn is None:
            continue
        name = "main" if ctn.get("nodeType") == "mainSeq" else None
        if name is None:
            k += 1
            name = f"interactive{k}"
        out[name] = [c.get("grpId") for c in ctn.iter(P + "cTn") if c.get("presetClass")]
    return out


def whole_targets_from_pptx(pptx_path: str, slide_id: int) -> dict:
    """{'main': [True if the effect targets the whole shape the way other apps write it: a bare <p:spTgt>
    (no text element) AND no <p:bldP> build entry for that shape and build group, ...], 'interactive1': [...],
    ...} from a saved deck: one entry per effect, in the order of build_groups_from_pptx. PowerPoint itself
    writes a bare target with a build entry for text builds (e.g. by letter), where COM's AnimateBackground
    is right and must be kept."""
    with zipfile.ZipFile(pptx_path) as z:
        xml = z.read(slide_part_name(z, slide_id))
    t = next(ET.fromstring(xml).iter(P + "timing"), None)
    out = {}
    if t is None:
        return out
    built = {(b.get("spid"), b.get("grpId")) for b in t.iter(P + "bldP")}
    k = 0
    for seq in t.iter(P + "seq"):
        ctn = seq.find(P + "cTn")
        if ctn is None:
            continue
        name = "main" if ctn.get("nodeType") == "mainSeq" else None
        if name is None:
            k += 1
            name = f"interactive{k}"
        flags = []
        for c in ctn.iter(P + "cTn"):
            if not c.get("presetClass"):
                continue
            tgt = c.find(f".//{P}spTgt")
            flags.append(tgt is not None and tgt.find(P + "txEl") is None
                         and (tgt.get("spid"), c.get("grpId")) not in built)
        out[name] = flags
    return out


def mark_whole_shape_effects(snap: list[dict], whole: dict, has_text_frame) -> list[int]:
    """Other apps (e.g. Google Slides) write a whole-shape effect as a bare <p:spTgt> without a build entry,
    and COM then reports AnimateBackground = False for it. Re-created with that value, PowerPoint animates
    only the shape's (possibly empty) text and the shape itself shows from the start. For such effects on
    shapes that have a text frame, set AnimateBackground = True ("as one object") in the snapshot, so rebuild
    re-creates them as whole-shape effects and compare() expects that. Returns the shape ids changed."""
    changed = []
    for q in snap:
        flags = whole.get(q["seq"])
        if not flags or len(flags) != len(q["effects"]):
            continue
        for e, is_whole in zip(q["effects"], flags):
            if (is_whole and not e.get("AnimateBackground") and not e.get("Paragraph")
                    and not e.get("BuildByLevelEffect") and has_text_frame(e["shape_id"])):
                e["AnimateBackground"] = TRUE
                changed.append(e["shape_id"])
    return changed


def rebuild(container, before: list[dict], resolve_shape, build_groups: dict | None = None) -> dict:
    """Recreate every effect in `before` that is no longer on the slide, then restore the original
    order of every sequence. `resolve_shape(id)` returns the (possibly grouped) shape with that id."""
    tl = container.TimeLine
    report = {"recreated": 0, "survived": 0, "property_failures": [], "sequences_created": 0}
    # Map existing interactive sequences by trigger shape id.
    inter = {}
    ordinals: dict = {}
    for k in range(1, tl.InteractiveSequences.Count + 1):
        q = tl.InteractiveSequences(k)
        key = _seq_key_live(q, ordinals)
        if key is not None:
            inter[key] = q

    for seq_snap in before:
        effects = seq_snap["effects"]
        if not effects:
            continue
        if seq_snap["seq"] == "main":
            q = tl.MainSequence
        else:
            key = seq_snap.get("seq_key") or (seq_snap["trigger_shape"],)
            q = inter.get(key)
            if q is None:
                q = tl.InteractiveSequences.Add()
                inter[key] = q
                report["sequences_created"] += 1
        mapping = _align(effects, q) if q.Count else [None] * len(effects)
        lost = [i for i, m in enumerate(mapping) if m is None]
        report["survived"] += len(effects) - len(lost)
        for i, cur in enumerate(mapping):
            s = effects[i]
            if cur is not None and s["Timing"]["TriggerType"] not in (ON_SHAPE_CLICK, None):
                # PowerPoint may re-flag a survivor's start when the click-leader before it is deleted.
                _set_if_diff(cur.Timing, "TriggerType", s["Timing"]["TriggerType"], report["property_failures"],
                             f"{seq_snap['seq']}#{i + 1}")
        done: set[int] = set()
        for i in lost:
            if i in done:
                continue
            s = effects[i]
            label = f"{seq_snap['seq']}#{i + 1}(shape {s['shape_id']})"
            shape = resolve_shape(s["shape_id"])
            first_in_seq = seq_snap["seq"] != "main" and q.Count == 0
            trig = ON_SHAPE_CLICK if first_in_seq else s["Timing"]["TriggerType"]
            if s.get("BuildByLevelEffect"):
                # A text build ("all paragraphs at once" / "by paragraph"): ONE AddEffect with that build level
                # creates the linked effects together (a whole-shape part, Paragraph 0, plus one per paragraph)
                # in a single build. Re-adding them one by one would make separate builds or extra effects.
                key = (s["shape_id"], s["EffectType"], s["Exit"], s["BuildByLevelEffect"])
                unit, paras = [], set()
                gl = (build_groups or {}).get(seq_snap["seq"])
                gl = gl if gl and len(gl) == len(effects) else None
                for j in lost:
                    ej = effects[j]
                    if j in done or j < i or                             (ej["shape_id"], ej["EffectType"], ej["Exit"], ej.get("BuildByLevelEffect")) != key:
                        continue
                    if gl is not None:  # exact: same build id as written by PowerPoint
                        if gl[j] == gl[i]:
                            unit.append(j)
                        continue
                    pk = ej["Paragraph"] or 0
                    if pk in paras or (pk == 0 and paras):  # repeated part, or a new whole-shape part
                        break
                    paras.add(pk)
                    unit.append(j)
                n0 = q.Count
                try:
                    q.AddEffect(shape, s["EffectType"], s["BuildByLevelEffect"], trig, -1)
                except Exception as exc:
                    raise ToolError(f"AddEffect (text build) failed for {label}: {com_err(exc)}")
                created: dict[int, list] = {}
                for k in range(n0 + 1, q.Count + 1):
                    e = q.Item(k)
                    created.setdefault(safe(lambda e=e: e.Paragraph, 0) or 0, []).append(e)
                for j in unit:
                    pool = created.get(effects[j]["Paragraph"] or 0) or []
                    e = pool.pop(0) if pool else None
                    lbl = f"{seq_snap['seq']}#{j + 1}(shape {s['shape_id']} part {effects[j]['Paragraph'] or 'shape'})"
                    if e is None:
                        report["property_failures"].append(f"{lbl}: text-build part was not recreated")
                        continue
                    if first_in_seq and j == unit[0]:
                        _set_trigger_shape(e, seq_snap, resolve_shape, report, lbl)
                    e, fails = apply_properties(q, e, effects[j], lbl)
                    report["property_failures"] += fails
                    mapping[j] = e
                    report["recreated"] += 1
                    done.add(j)
                for pool in created.values():  # parts that were not animated originally
                    for e in pool:
                        e.Delete()
                continue
            try:
                if first_in_seq and s["Timing"]["TriggerType"] == ON_BOOKMARK:
                    e = q.AddTriggerEffect(shape, s["EffectType"], ON_BOOKMARK,
                                           resolve_shape(seq_snap["trigger_shape"]),
                                           s["Timing"].get("TriggerBookmark") or "", 0)
                    first_in_seq = False
                else:
                    e = q.AddEffect(shape, s["EffectType"], 0, trig, -1)
            except Exception as exc:
                raise ToolError(f"AddEffect failed for {label}: {com_err(exc)}")
            if first_in_seq:
                _set_trigger_shape(e, seq_snap, resolve_shape, report, label)
            e, fails = apply_properties(q, e, s, label)
            report["property_failures"] += fails
            mapping[i] = e
            report["recreated"] += 1
            done.add(i)
        if any(m is None for m in mapping):
            report["property_failures"].append(f"{seq_snap['seq']}: some effects could not be recreated")
            continue
        # Restore order: put each effect at its original position.
        for pos, e in enumerate(mapping, 1):
            if e.Index != pos:
                e.MoveTo(pos)
        for pos, (s, e) in enumerate(zip(effects, mapping), 1):
            if e.Index != pos or safe(lambda: e.Shape.Id) != s["shape_id"]:
                report["property_failures"].append(f"{seq_snap['seq']}: order check failed at {pos}")
                break
    return report


COMPARE_KEYS = ("shape_id", "EffectType", "Exit", "Paragraph", "params", "Color2", "AfterEffect", "Dim",
                "AnimateBackground", "AnimateTextInReverse", "BuildByLevelEffect", "TextUnitEffect", "Sound")
COMPARE_TIMING = ("TriggerType", "TriggerDelayTime", "Duration", "Accelerate", "Decelerate", "AutoReverse",
                  "RepeatCount", "RewindAtEnd", "SmoothStart", "SmoothEnd", "Restart", "TriggerShape")


def _cmp_view(e: dict) -> dict:
    v = {k: e.get(k) for k in COMPARE_KEYS}
    v.update({f"Timing.{k}": e["Timing"].get(k) for k in COMPARE_TIMING})
    v["behavior_types"] = [b["Type"] for b in e["behaviors"]]
    v["motion_paths"] = [(b.get("props") or {}).get("Path") for b in e["behaviors"] if b["Type"] == 1]
    return v


def _close(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) < 1e-4
    if isinstance(a, dict) and isinstance(b, dict):
        return a.keys() == b.keys() and all(_close(a[k], b[k]) for k in a)
    if isinstance(a, (list, tuple)) and isinstance(b, (list, tuple)):
        return len(a) == len(b) and all(_close(x, y) for x, y in zip(a, b))
    return a == b


def compare(before: list[dict], after: list[dict]) -> list[str]:
    """Differences between two snapshots in the properties COM reports reliably. (Behavior point lists
    read back inconsistently through COM; the saved-XML check is the authority for those.)"""
    diffs = []
    main_b = next((q for q in before if q["seq"] == "main"), None)
    main_a = next((q for q in after if q["seq"] == "main"), None)
    by_trig_b = {(q.get("seq_key") or (q["trigger_shape"],)): q for q in before if q["seq"] != "main"}
    by_trig_a = {(q.get("seq_key") or (q["trigger_shape"],)): q for q in after if q["seq"] != "main"}
    pairs = [("main", main_b, main_a)] + [(f"trigger #{t}", by_trig_b.get(t), by_trig_a.get(t))
                                         for t in sorted(set(by_trig_b) | set(by_trig_a), key=str)]
    for name, b, a in pairs:
        be = (b or {}).get("effects", [])
        ae = (a or {}).get("effects", [])
        if len(be) != len(ae):
            diffs.append(f"{name}: {len(be)} effects before, {len(ae)} after")
            continue
        for i, (x, y) in enumerate(zip(be, ae), 1):
            vx, vy = _cmp_view(x), _cmp_view(y)
            for k in vx:
                if not _close(vx[k], vy[k]):
                    diffs.append(f"{name} #{i} (shape {x['shape_id']}): {k} {vx[k]!r} -> {vy[k]!r}")
    return diffs


# --------------------------------------------------------------------------------------------------
# XML fingerprint (what PowerPoint actually writes to the file)
# --------------------------------------------------------------------------------------------------

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def slide_part_name(z: zipfile.ZipFile, slide_id: int) -> str:
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    rels = ET.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rid = None
    for s in pres.iter(f"{{{NS['p']}}}sldId"):
        if int(s.get("id")) == slide_id:
            rid = s.get(f"{{{NS['r']}}}id")
    if rid is None:
        raise ToolError(f"slide id {slide_id} not found in saved file")
    for rel in rels:
        if rel.get("Id") == rid:
            target = rel.get("Target").lstrip("/")
            return target if target.startswith("ppt/") else "ppt/" + target
    raise ToolError(f"relationship {rid} not found")


P = f"{{{NS['p']}}}"


def timing_from_pptx(pptx_path: str, slide_id: int, idmap: dict[int, int] | None = None) -> tuple[str, str]:
    """(raw <p:timing> XML, normalized form) of one slide inside a saved .pptx. `idmap` renames shape ids
    (old -> new) first, for comparing a slide before and after a shape was replaced by a copy."""
    with zipfile.ZipFile(pptx_path) as z:
        xml = z.read(slide_part_name(z, slide_id))
    m = re.search(rb"<p:timing>.*?</p:timing>", xml, re.S)
    raw = m.group(0).decode("utf8") if m else ""
    t = next(ET.fromstring(xml).iter(P + "timing"), None)
    if t is not None and idmap:
        for el in t.iter():
            v = el.get("spid")
            if v is not None and int(v) in idmap:
                el.set("spid", str(idmap[int(v)]))
        for old, new in idmap.items():
            raw = raw.replace(f'spid="{old}"', f'spid="{new}"')
    if t is None:
        return raw, "", ""
    return raw, normalize_timing(t), normalize_timing(t, strip_build=True)


def timeline_model(snap: list[dict], expand: dict[int, list[int]] | None = None) -> dict:
    """What a viewer sees: for every shape, the list of (sequence, click step, start s, end s, effect type,
    exit) of the effects that touch it. `expand` maps a group id to its pieces, so an effect on a group
    counts as the same effect on each piece. Two slides with equal models animate identically."""
    expand = expand or {}
    out: dict[int, list] = {}
    for q in snap:
        name = "main" if q["seq"] == "main" else f"trigger{q['trigger_shape']}"
        step, cont_start, step_end = 0, 0.0, 0.0
        for e in q["effects"]:
            t = e["Timing"]
            trig = t.get("TriggerType")
            delay = t.get("TriggerDelayTime") or 0.0
            dur = (t.get("Duration") or 0.0) * max(1.0, float(t.get("RepeatCount") or 1))
            if trig in (ON_CLICK, ON_SHAPE_CLICK, ON_BOOKMARK):
                step, cont_start, step_end = step + 1, 0.0, 0.0
            elif trig == AFTER_PREV:
                cont_start = step_end
            start = cont_start + delay
            step_end = max(step_end, start + dur)
            for sid in expand.get(e["shape_id"], [e["shape_id"]]):
                out.setdefault(sid, []).append((name, step, round(start, 3), round(start + dur, 3),
                                                e["EffectType"], bool(e["Exit"])))
    return out


def normalize_timing(timing: ET.Element, strip_build: bool = False) -> str:
    """Canonical form: drop cTn ids, sort interactive sequences by trigger and the build list by shape,
    THEN renumber grpId per shape by first use, and sort attributes. Equal result = same animation.
    strip_build=True also drops the build list and grpId links (bookkeeping PowerPoint keeps for the
    Animation Pane's text options): equal then means the same timing tree, to be confirmed by playback."""
    root = copy.deepcopy(timing)
    for el in root.iter():
        el.attrib.pop("uiExpand", None)  # Animation Pane expand/collapse state: UI only, not playback
    if strip_build:
        for el in root.iter():
            el.attrib.pop("grpId", None)
        for parent in list(root.iter()):
            for c in list(parent):
                if c.tag == P + "bldLst":
                    parent.remove(c)
    for ctn in root.iter(P + "cTn"):
        ctn.attrib.pop("id", None)

    def trigger_key(seq):
        c = seq.find(f"./{P}cTn/{P}stCondLst/{P}cond")
        t = c.find(f"./{P}tgtEl/{P}spTgt") if c is not None else None
        if t is None:
            return (0, "", "")
        return (1, t.get("spid", "").zfill(8), ET.tostring(c, encoding="unicode"))

    for parent in root.iter():
        kids = list(parent)
        slots = [i for i, c in enumerate(kids) if c.tag == P + "seq"]
        if len(slots) > 1:
            ordered = sorted((kids[i] for i in slots), key=trigger_key)
            for i, c in zip(slots, ordered):
                parent[i] = c

    seen: dict[tuple, str] = {}
    counters: dict[str | None, int] = {}
    for el in root.iter():
        g = el.get("grpId")
        if g is None:
            continue
        spid = el.get("spid")
        if spid is None:
            tgt = el.find(f".//{P}spTgt")
            spid = tgt.get("spid") if tgt is not None else None
        key = (spid, g)
        if key not in seen:
            counters[spid] = counters.get(spid, -1) + 1
            seen[key] = str(counters[spid])
        el.set("grpId", seen[key])

    for bld in root.iter(P + "bldLst"):
        items = sorted(list(bld), key=lambda c: (c.get("spid", "").zfill(8), c.get("grpId", ""), c.tag))
        for c in list(bld):
            bld.remove(c)
        bld.extend(items)
    for el in root.iter():
        items = sorted(el.attrib.items())
        el.attrib.clear()
        el.attrib.update(items)
    return ET.tostring(root, encoding="unicode")


def xml_diff(a: str, b: str, context: int = 2, limit: int = 40) -> list[str]:
    import difflib
    al = re.sub(r">", ">\n", a).splitlines()
    bl = re.sub(r">", ">\n", b).splitlines()
    return list(difflib.unified_diff(al, bl, "before", "after", lineterm="", n=context))[:limit]
