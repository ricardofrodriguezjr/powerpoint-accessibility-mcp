# How it works

This page is for the curious and for contributors. You do not need it to use the tool.

## The pieces

| Piece | What it is | Where |
|---|---|---|
| **MCP server** (`ppt-a11y`) | A small Python program your AI app talks to. It drives the copy of PowerPoint installed on your computer through Windows COM automation (pywin32). Every edit is made by PowerPoint itself. | [`ppta11y/`](../ppta11y) |
| **Agent skill** (`pptx-wcag-alt-text`) | Plain-language instructions that tell the AI how to remediate a deck: what to check, in what order, which changes are allowed, and when to stop. | [`skills/`](../skills) |
| **Installer** | Downloads the code, creates a private Python environment, registers the server with your AI app and copies the skill. | [`install.ps1`](../install.ps1) |

LibreOffice and other converters are never used. Because PowerPoint does every edit, the file stays a normal
PowerPoint file, and anything the tool does not touch stays exactly as it was.

## Safety rules enforced in code

- **The original is never saved.** `open_working_copy` copies `Lecture.pptx` to `Lecture_accessible.pptx`
  next to it, and only registered working copies can be edited or saved. A small record in
  `Lecture_accessible_a11y_checks\working_copy.json` marks which file is the copy.
- **Every edit is a transaction.** The tool saves a checkpoint, makes the change, verifies it, and keeps it
  only if every check passes. On any difference or error it reopens the checkpoint, so nothing half-done
  is ever kept.
- **Invisible changes are proven invisible.** Structural edits (grouping, restacking for reading order,
  converting a layout table to text boxes) are verified with a before and after render of the slide,
  shape positions, and the slide's saved animation XML.
- **Visible changes are limited and logged.** Only what WCAG itself requires is allowed to change visibly:
  a text or fill colour nudged to the nearest passing colour, "(2 of 3)" on repeated titles, link text,
  and repeated text in unmerged table cells. Each one is listed with before and after values.
- **Spreadsheets are never edited in place.** Results go to a copy of your alt-text spreadsheet.
- **Reports stay next to your deck** in `<deck>_a11y_checks\`: renders, audit reports and a change log.

## What gets fixed, and how

| WCAG 2.1 | Problem | What the tool does |
|---|---|---|
| 1.1.1 | Pictures, charts, diagrams without alt text | Adds alt text, or marks purely decorative objects as decorative. Multi-piece figures are grouped so they get one description, **keeping their animations** (see below). |
| 2.4.2, 2.4.6 | Slides without a title, or with the same title | Turns the visible heading into the real title placeholder with an identical look, adds a hidden title for repeated section banners, and numbers continued topics as "Topic (2 of 3)". |
| 1.3.2, 2.4.3 | Screen readers read objects in the wrong order | Restacks objects so the title is read first, content in visual order and footers last. Overlapping objects keep their stacking, so the slide cannot look different. |
| 1.4.3 | Text that is too light | Measures contrast against what PowerPoint actually draws behind each run of text, in every animation step, and picks the nearest colour that passes. |
| 1.3.1 | Tables with merged cells or no header row | Diagram tables become identical-looking text boxes. Data tables are unmerged and get a header row. |
| 2.4.4 | Raw or broken link text | Replaces it with a description of the destination and keeps the link working. |

When it is done, `wcag_audit` runs every rule again plus PowerPoint's own Accessibility Checker and reports
`done: true` only when nothing is left.

## Keeping animations when grouping

When shapes are grouped, PowerPoint deletes their animation effects (Ctrl+G and `ShapeRange.Group()`
behave the same way). PowerPoint does allow two things:

1. A grouped shape keeps its `Shape.Id`.
2. `Sequence.AddEffect(group.GroupItems(k), ...)` animates a shape *inside* a group, and the effect is
   saved and plays in Slide Show.

So `group_shapes`:

1. saves a checkpoint of the working copy;
2. snapshots every effect on the slide: main and trigger sequences, timing, parameters, by-paragraph
   builds, after-effects and motion paths;
3. groups the shapes;
4. re-adds each lost effect to the same shape inside the group, copies its properties and restores the
   original order;
5. verifies the slide's saved `<p:timing>` XML against the checkpoint, shape positions, and a before and
   after render;
6. rolls back on any difference.

`compare_playback` then plays the original and the working copy in Slide Show and compares a frame after
every click, which is exactly what your audience sees.

Evidence from testing:

- Real lecture slides with 7 and 8 effects: the saved timing XML was byte-identical after grouping, and
  Slide Show frames matched at every click.
- The stress test (`tests/stress_animations.py`) groups 11 shapes animated with Fly-in, Wipe, a
  3-paragraph text build, Spin with repeat and auto-reverse, a custom motion path, an exit Fade with
  easing, Grow, a fill-colour change, Dim-after, by-letter text and a click trigger. The saved XML stays
  byte-identical (14,943 bytes).
- A Google Slides export with 13 animated slides, including diagrams of 31 to 87 effects, played
  identically to the original after remediation.

## Reading order without visible changes

PowerPoint's reading order is its stacking order: the object at the back is read first. Changing the order
of two objects can only change the picture where they overlap, so the tool:

1. computes everything each object can paint on the slide (its box if it has a fill, outline or picture,
   its text, shadows, and the whole slide if an animation moves or resizes it);
2. moves only the objects that must move; objects that overlap keep their stacking;
3. keeps a swap only if the slide renders identically and no animated object could reveal the difference;
4. otherwise uses the closest order that keeps every overlapping pair as it was, and asks the AI to check
   that order before accepting it.

PowerPoint's own "Check reading order" item is a heuristic prompt to double-check. It can stay on for a
slide whose order is correct (for example code slides with notes whose arrows point into the code). The
report lists those slides and explains why their order is right.

## Text contrast

Contrast is measured the way WCAG defines it (technique G18): against the pixels drawn right around the
letters. The tool renders each slide, finds the letters of every text run, and samples a thin ring just
outside them. Text on gradients, pictures or split backgrounds is checked against every colour behind it.
Text near an animated object is checked at every step of the animation in which the text is visible.
Large text (18 pt, or 14 pt bold) needs 3:1, other text 4.5:1. The fix is the closest colour of the same
hue that passes; when white text sits on a mid-tone fill, darkening the fill slightly is preferred because
it changes the slide less.

## Known limits

- Windows and desktop PowerPoint (Microsoft 365) only. macOS PowerPoint has no COM automation.
- A group that is itself animated as a whole cannot be nested into a new group as is. `flatten_group`
  handles the common Appear and Fade cases; other cases are reported with alternatives.
- Custom sound files attached to effects cannot be copied through COM, and media-bookmark triggers are
  untested. In both cases the check fails and the deck is rolled back, so nothing is lost.
- Colours stored inside animations (Dim after animation, colour emphasis) are measured and reported, but
  must be changed by hand in the Animation Pane.
