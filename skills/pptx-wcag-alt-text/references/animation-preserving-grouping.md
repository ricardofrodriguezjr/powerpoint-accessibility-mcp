# Grouping without losing animations: how and why it works

## The problem

PowerPoint deletes a shape's animation effects when the shape is put into a group.

- This happens with Ctrl+G and with COM `ShapeRange.Group()`.
- Microsoft Q&A, BrightCarbon and PPT MVPs all describe it, with no workaround for per-piece builds.
  Public tools such as IguanaTex only move whole-group effects.
- The handoff deck lost all 8 click animations on slide 8 this way.

## What PowerPoint *does* allow (verified on 16.0.20430, Microsoft 365, Oct 2026)

| Behaviour | Test result |
|---|---|
| Members keep their `Shape.Id` after `Group()` | Yes. The ids were the same before and after on slide 8. |
| `Sequence.AddEffect(group.GroupItems(k), ...)` | Works. `Effect.Shape.ParentGroup` is the new group. |
| These effects survive save and reopen | Yes |
| They play in Slide Show | Yes. Frame-by-frame captures of the original and the grouped deck are identical at every click. |
| `Effect.Shape = other` (write) | Works for top-level shapes. However, effects re-pointed to a temporary shape were *still* deleted by a later `Group()` of the original members, so "park the effects on a temp shape" does not work. |
| Grouping unrelated shapes | Leaves other effects alone (control test). |

## The algorithm (`ops.group_shapes` and `animations.rebuild`)

1. **Checkpoint.** Save the working copy, so the file on disk is the "before" state.
2. **Snapshot** every sequence: the main sequence and each trigger sequence (keyed by trigger shape).
   For every effect, record:
   - target id, type, exit, paragraph;
   - all `Timing` members;
   - `EffectParameters` (Direction, Amount, Size, Color2, ...);
   - after-effect and dim colour, text unit, build level, animate-background, reverse, sound;
   - behaviours (motion paths).
3. **Group** with `Shapes.Range([indices]).Group()`.
4. **Align.** Effects that survived (other shapes) are matched to the snapshot in order. The rest are
   lost and get rebuilt.
5. **Rebuild** each lost effect on the same shape id, which now lives inside the group.
   - **Text builds** ("all paragraphs at once" = level 1, "by 1st-level paragraphs" = level 2) are recreated
     with **one** `AddEffect(shape, type, Level=build level)`. That call creates the linked parts
     together: a whole-shape part (`Paragraph` 0) plus one part per paragraph. Each part is then matched to
     the snapshot, and parts that were not animated are deleted. Adding parts one by one would create
     separate builds or extra effects (seen on slide 46).
   - **Single paragraph effects with build level 0** (decks exported from Google Slides) are recreated
     individually: `AddEffect(level 0)`, then `Effect.Paragraph = n`.
   - **Whole-shape effects written by other apps** (Google Slides): saved as a bare `<p:spTgt>` with no
     build entry, and COM then reports `AnimateBackground = False`. Copying that value re-creates the effect
     as text-only, so the shape itself is visible from the start. The saved target decides instead: a bare
     target on a shape with a text frame is re-created "as one object". The XML then differs only in the
     build list, so playback must confirm it (step 7).
   - Trigger sequences are recreated with `InteractiveSequences.Add()` and `Timing.TriggerShape`.
   - Properties are copied only where they differ:
     - Timing (delay, duration, easing, repeat, auto-reverse, rewind, smooth, bounce, restart);
     - parameters;
     - `ConvertToAfterEffect`, `ConvertToTextUnitEffect`, `ConvertToBuildLevel`,
       `ConvertToAnimateBackground`, `ConvertToAnimateInReverse`;
     - motion path;
     - for custom effects, the full behaviour list.
6. **Reorder** with `Effect.MoveTo` so every sequence has its original order.
7. **Verify.** All of these must pass:
   - `SaveCopyAs` a temporary file and compare the slide's `<p:timing>` XML with the checkpoint file.
     It must be byte-identical, or equal after normalisation: `cTn` ids dropped, `grpId` renumbered per
     shape, trigger sequences and build list sorted.
   - Every member shape's Left/Top/Width/Height/Rotation is unchanged (within 0.05 pt).
   - `Slide.Export` before and after: at most 20 pixels may differ, on any colour channel.
   - If the timing tree is identical and only the **build list** (`bldLst` / `grpId` bookkeeping)
     differs, the change is accepted only if a real **Slide Show playback** comparison (`compare_playback`,
     every click) is identical. COM cannot recreate effects that lack build entries, as in decks from other
     apps. `uiExpand` (the Animation Pane's expand/collapse state) is ignored.
8. **Commit or roll back.** Any failure: close without saving and reopen the checkpoint. Success: save.

## Evidence

| Case | Effects | Result |
|---|---|---|
| Deck slide 8, AI diagram (experiment) | 8 Appear (click and after-previous chain) | Timing XML byte-identical (5,948 B); playback identical |
| Deck slide 13, hexagons | 7 Appear | XML identical (5,312 B); 0.0 % render change; playback identical |
| Deck slide 8 (manual group, animations lost) | 8 restored from the original | XML equivalent to the original; group moved back up 7.01 pt; render matches the original |
| Stress deck (`tests/stress_animations.py`) | 13 effects (see the list below) | Byte-identical XML (14,943 B) |
| First stress run, before the paragraph fix | | XML check caught 3 separate builds and rolled back correctly |

The stress deck's 13 effects are on 11 grouped shapes:
- Fly-in from the left;
- Wipe from the bottom after previous, with a delay;
- a 3-paragraph Fade build;
- Spin 180° repeated twice with auto-reverse;
- a custom motion path;
- an exit Fade with acceleration and deceleration;
- Grow 150 %;
- a fill-colour change;
- Appear plus dim-after;
- Fade by letter;
- a trigger sequence fired by an ungrouped button.

An ungrouped animated shape also sits in the middle of the sequence.

| Real deck, all 49 rows applied with `apply_plan` | 76 steps | Every grouping: XML identical or equivalent, 0 px render change. Slides 31 and 46 (Google-style text builds) were confirmed by playback. All 62 slides are pixel-identical to the original, and PowerPoint's checker reports 0 "Missing alt text" (was 144). |
| Playback found a bug that XML missed | `flatten_group` v1 | Flattened text boxes animated only their text, so the coloured box showed early. Fixed: pieces with text get `AnimateBackground` ("as one object"), which is now verified. |
| Playback found a second one (a Google Slides export, Oct 2026) | `flatten_group` v2, rebuild | Pieces with an *empty* text frame (bars, ovals, arrows) and Google Slides effects were re-created text-only, so the shapes showed before their click. Render, timeline model and XML-vs-checkpoint could not see it. Fixed: every piece with a text frame animates as one object, and rebuild reads the saved targets (verified by Slide Show playback against the original). 13 animated slides then played identically to the original, including four Google-style diagrams of 31-87 effects. |

## Flattening an animated group (`flatten_group`)

- Only Appear and Fade qualify: for these, "animate the group" looks exactly like "animate every piece
  at the same moment".
- Each piece gets the group's effect. The first piece keeps the original trigger and delay; the others
  get "with previous" and the same delay.
- Pieces with a text frame, even an empty one (rectangles, ovals, arrows), are set to animate "as one
  object" (`ConvertToAnimateBackground`). Otherwise PowerPoint animates only their text and the shape is
  visible from the start.
- Verification:
  - a viewer timeline model, which compares click step, start and end per piece against the original
    expanded over the group;
  - the render;
  - an "as one object" check;
  - a `compare_playback` of the slide against the untouched original afterwards. **Required**: the render
    shows the final state and the timeline model has no notion of "as one object", so only playback sees
    a shape that appears too early.

## Hard limit: a group that is itself animated cannot be nested

COM's `Shape.GroupItems` is flat:
- A group nested inside another group can no longer be addressed. `ParentGroup` returns the outer group.
- So an effect that animates the inner group as a whole can't be re-created.

`check_grouping` blocks this case before anything changes.
- On the test deck this affects slide 17 (groups #24–#27) and slide 58 (#14–#16).
- Ask the user: (a) leave those groups out of the new group, (b) carrier, or (c) accept losing the effect.

## Limits (the XML check catches all of these, and the deck is rolled back)

- **Custom sound files on effects:** COM cannot read the embedded sound back out.
- **Media bookmarks** (`ON_BOOKMARK` triggers) on media inside the group: untested.
- **A trigger shape that is itself grouped:** handled (the trigger becomes the grouped shape) but only
  lightly tested.
- **Layouts and masters:** there is no XML check for these parts (layout animations are very rare). A
  COM-level comparison is used instead.

## Alternative that never touches the timeline (WCAG G196 "carrier")

Put the alt text on one piece and mark the others decorative.

- Use it when grouping is impossible (placeholders, tables) or rolled back.
- WCAG accepts it. Microsoft guidance and some checklists prefer a group.
