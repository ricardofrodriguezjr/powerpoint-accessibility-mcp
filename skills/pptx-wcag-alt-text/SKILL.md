---
name: pptx-wcag-alt-text
description: Make any PowerPoint deck fully WCAG 2.1 AA / Section 508 compliant and ready to use, by driving desktop PowerPoint on Windows through the ppt-a11y MCP - alt text or "Mark as decorative" on every picture, chart, diagram and shape (grouping multi-piece figures without losing animations), a real and unique title on every slide, correct screen-reader reading order, sufficient text contrast, accessible tables (no merged cells, header rows) and descriptive link text - while keeping the slides' content, look and animations identical except for the few small, logged changes WCAG itself requires. Verifies with its own WCAG audit plus PowerPoint's Accessibility Checker and click-by-click Slide Show playback. Use when asked to make a .pptx accessible or WCAG/Section 508/ADA compliant, add alt text, fix reading order, slide titles, contrast, tables or links, remediate or audit a deck for accessibility, or apply an alt-text spreadsheet to a presentation.
---

# PowerPoint WCAG remediation: compliant, same content, same look, same animations

You drive real PowerPoint through the `ppt-a11y` MCP server. This skill names its tools plainly
(`open_working_copy`, `wcag_audit`, ...); your app may show them with a prefix, such as
`mcp__ppt-a11y__open_working_copy` in Claude Code or `ppt-a11y_open_working_copy` in OpenCode.
Every change is made by PowerPoint itself, on a **working copy**. The original deck and spreadsheet are
never modified. The goal is a deck a professor can use as-is: nothing left "to do by hand".

**Done means `wcag_audit()` returns `done: true`:**
- 1.1.1 every non-text object has meaningful alt text, is decorative, or is inside a group with alt text;
- every slide has a real, non-empty, unique title;
- every slide's reading order is right: title first, content in visual/logical order, footer last;
- all text meets contrast (4.5:1, or 3:1 for >= 18 pt / >= 14 pt bold);
- no table has merged cells, every table has a header row;
- every link says where it goes and has a target;
- PowerPoint's Accessibility Checker shows **0** issues;
- every slide looks identical to the original except the logged WCAG changes (section 1, rule 2);
- every animation plays identically, click by click.

If PowerPoint's checker still lists "Check reading order" on a slide whose order you verified (its rule is
an undocumented heuristic that can stay on after a correct fix), say so explicitly in the report with the
slide numbers and why the order is right. Everything else must be 0.

## 0. Preconditions

- Windows, desktop PowerPoint (Microsoft 365), the `ppt-a11y` MCP connected.
- If the tools are missing, point the user to the Quick start at
  https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp (one PowerShell command installs the
  server and this skill). Until the app is restarted, the same tools run from a terminal:
  `<install folder>\.venv\Scripts\python.exe <install folder>\cli.py <tool> "@args.json"` (the installer
  uses `%USERPROFILE%\.ppt-a11y\app`).
- Ask once, briefly, only for what you cannot find: the deck; the alt-text spreadsheet if there is one; an
  untouched original if the deck was already edited by hand.

## 1. Principles (in priority order)

1. **Work on a copy.** Start with `open_working_copy(source)`. Never save over the original.
2. **Never change content or design, except what WCAG requires.** Invisible structural changes are fine
   when needed (each verified: identical render, identical animation; each noted). The only visible changes
   allowed are these, applied by default and each listed in the report with before/after values:
   - text or fill colour nudged to the nearest passing colour (`fix_contrast`; usually ΔE < 12);
   - a repeated visible slide title made unique, e.g. "LOOCV steps (2 of 3)";
   - link text replaced by a description of its destination;
   - a merged cell in a real data table split, with its text repeated (`unmerge_table`).
   Anything else that would change the look: stop and ask.
3. **Group everything that forms one figure** (pieces, arrows, labels, legend, caption); alt text on the group.
4. **Prefer the invisible fix.** Visible heading → real title (`promote_to_title`, looks identical). Layout
   or diagram table → text boxes (`table_to_shapes`, looks identical). Text nobody can see → decorative.
5. **Use supplied alt text verbatim**; flag mismatches, never edit the spreadsheet's input columns.
6. **Never group or restack animated shapes any other way than with the tools.** `ok:false` means the deck
   was rolled back: report the failures; don't retry with random tweaks.
7. **Reading order is fixed last**, after every other structural change (grouping, moving, converting,
   titles all change the stacking order).

## 2. Workflow

### Step 1: Baseline
1. `open_working_copy(source)`.
2. `wcag_audit()` (our rules + PowerPoint's checker): the to-do list for every category.
3. If there is a spreadsheet, `read_alt_text_sheet(path)`.

### Step 2: Alt text (1.1.1)
As before: `inspect_slide` + `slide_snapshot` per slide, map rows/objects to ids, `check_grouping`,
structural fixes (section 3), plan with `"alt_text": "ROW"`, `apply_plan` in chunks, `audit_alt_text` until
`slides_objects_needing_work` is 0 and every layout/master object is decorative or moved. Write alt text
yourself (section 5) only when no sheet supplies it.

### Step 3: Slide titles
`title_report()` lists slides without a title (with the most title-like visible text as `candidate`) and
groups of slides sharing a title.
- **Visible heading that is unique in the deck** → `promote_to_title(slide, candidate_id)`. It becomes the
  real title placeholder, same look (a title placeholder is added to the layout if it has none: invisible).
- **Same visible heading on several title-less slides** (section dividers, repeated agenda/roadmap/recap
  slides, a module banner): promoting would create duplicates. Instead `add_hidden_title(slide, text)` with a
  short, **descriptive, unique** title of what distinguishes the slide, taken from its content (e.g. the
  agenda tile it highlights: "Topic 2: Cross-Validation"; first/last occurrence: "... : module topics" /
  "... : module review"). The visible banner stays as it is. Check the slide-1 title too (it must not repeat).
- **Duplicate visible titles on content slides** → `make_titles_unique()`: each run of *consecutive* slides
  sharing a title gets "Title (n of N)" (Section 508 / HHS pattern), always the whole run in one transaction
  (naming one slide of a run numbers all of it: `extended_to_whole_run`). Slides that share a title but are
  not next to each other come back in `needs_distinct_title`: give each a descriptive title with
  `set_slide_title` (e.g. "Questions: Linear Regression"). If consecutive slides clearly cover distinct sub-topics,
  "Title: sub-topic" is better than numbering. Check `warnings` / `formatting_reset` and the compare image.
- `title_state` tells why a slide has no title: `absent` / `empty` (promote the heading or add a hidden
  title), `decorative` (`mark_decorative(..., decorative=False)` on the title), `hidden` (un-hide it).
  `set_slide_title` only edits a title screen readers get; it refuses empty, hidden or decorative ones.

### Step 4: Tables
`table_report()`:
- merged cells + `hierarchy: true` (a tree/roadmap drawn with merged cells), or a layout grid that is not
  data → `table_to_shapes(slide, id)`: real text boxes, same look, read branch by branch. If the result has
  `table_alt_text_not_kept`, put that text in the slide notes when it adds information.
- merged cells in a real data table (`hierarchy: false`; numbers, a spanning header, row groups) →
  `unmerge_table(slide, id)` (visible, logged; refused if the table would run off the slide or over other
  content - then ask), then `set_table_header`;
- `multi_row_header: true` → only row 1 is announced as headers: ask before combining the two rows into one
  (visible change);
- no header row → `set_table_header(slide, id)` (look kept: fills, fonts and borders are restored).

### Step 5: Links
`link_report()`: for every `unclear` link, `set_link_text(slide, shape_id, old_text, new_text)` with the
destination's real page title (derive from the URL path or the slide's own words; `suggested_text` is only a
start). A dead link whose text is a URL gets that URL as its target automatically. `advice` entries pass
WCAG AA; leave them unless the user wants AAA.

### Step 6: Contrast
`contrast_report()` then `fix_contrast()` (strategy "auto": the smallest visible change - nearest text colour,
or a slightly darker/lighter fill for every cell/shape of that colour when white text sits on a mid-tone
fill; a slide whose fill change fails is retried with text colours only). Our check also covers text on
the slide background, inside groups and SmartArt, over other shapes and pictures (every background colour
behind the text must pass), and layout/master text, which PowerPoint's checker misses. Logos and decorative
text are exempt. Text with an animated object over or under it (a cover that leaves on click, a box that
appears behind it) is measured in both states. Layout/master text gets one colour that passes on every slide
showing it (re-measured on all of them). `unfixable`: text over a mixed light/dark area (ask about a backing
shape), or a colour stored inside an animation (dim after animation, colour emphasis): tell the user the
colour to pick in the Animation Pane (`why` has it). `not_machine_checked`: chart text, gradient text, and
colour-blending emphasis effects: look at them. Record every colour change (from, to, ratio).

### Step 7: Reading order (always last)
1. `reading_order_report(compare_to=<original>)`: per slide what is read now, the suggested order, issues.
2. For slides whose order changes, look at `reading_order_snapshot(slide, suggested=True)` (numbered
   badges). The suggestion follows the rules in `references/reading-order.md`: title, subtitle, content in
   visual order (heading/figure columns column by column), call-out notes right after what they point at,
   converted tables as one block in their own order, footer and slide number last.
   Override only for meaning (e.g. a caption that belongs to a figure elsewhere): pass your order in
   `fix_reading_order(orders={"<slide>": [ids]})` or `set_reading_order(slide, ids)`.
3. `fix_reading_order()`: text nobody can see (completely behind static opaque objects, proven by a render
   with it hidden) is marked decorative; every slide is restacked **invisibly** (overlapping objects keep
   their stacking unless swapping them renders pixel-identical and neither moves during an animation;
   verified render, positions and animation XML; rollback otherwise); then PowerPoint's checker is run and
   other valid orders (never ones that break title-first / footer-last) are tried on slides it still flags.
   - `needs_attention`: the requested order was not possible invisibly (overlapping objects near an
     animation, or a pixel change), so the closest invisible order was applied but **not accepted**. Look at
     `reads_now` (and the snapshot). If it reads sensibly, accept it with
     `set_reading_order(slide, <applied[slide].order_now>)` (exactly that order: nothing moves, it is
     recorded). If it breaks meaning or a rule (title not first, footer before content), decide with the
     user. Only accepted or rule-following orders count as done in `wcag_audit`.
   - `hidden_text_failed`: the hidden-text check failed on that slide (its text is still read): report it.

### Step 8: Verify
1. `wcag_audit(compare_to=<original>)` → `done: true` (see the reading-order caveat at the top).
2. `compare_playback(slide, source_deck=<original>)` on **every** slide where animated shapes were grouped,
   flattened, converted or restored (needs the screen; batch them; ask for a quiet screen window first).
3. Optional whole-deck render diff (fresh reopen first): only the logged visible changes may differ.

### Step 9: Record and report
`record_sheet_results(...)` columns per alt-text row: "Applied (Claude)", "Shape IDs", "Structural change",
"Question for reviewer", "Alt text note". Then `record_change_log(dest=<same results copy>, changes=[...])`:
one row per non-alt-text change (titles, tables, links, contrast, reading order, hidden text made
decorative) with Slide, Category, WCAG criterion, What changed, Kind (invisible / visible), Before, After,
Verified. Then tell the user in plain words: what was done per category, each visible change
(before/after), before/after checker numbers, and open questions.

## 3. Structural fixes (invisible; always noted)

| Situation | Tool | Notes |
|---|---|---|
| Figure contains a group animated as a whole | `flatten_group`, then `group_shapes` | Appear/Fade only; verify with `compare_playback` |
| Figure contains a picture in a content placeholder | `placeholder_to_picture`, then group `'@id'` | |
| Informative content on a layout/master | `move_to_slides`, then group/alt `'@id'` | lands at the back (read first) until Step 7 |
| Decorative content on a layout/master | `mark_decorative(part=...)` | |
| Animations lost to a manual grouping | `restore_animations_from(..., restore_positions=true)` | |
| Slide has a visible heading but no title | `promote_to_title` | layout may gain a title placeholder |
| Repeated headings / no visible heading | `add_hidden_title` | off-slide, not shown in Slide Show |
| Layout/diagram table with merged cells | `table_to_shapes` | borders become one decorative group |
| Text completely hidden behind other objects | marked decorative by `fix_reading_order` | |
| Reading order wrong | `fix_reading_order` / `set_reading_order` | restack only; render-verified |

## 4. Asking the user
Ask only for a real choice (e.g. a visible change outside rule 2, a figure whose meaning is unclear). One
message, numbered; for each: what is on the slide, why it is a WCAG problem, options with one marked
**"⭐ Recommended for WCAG"**, and whether each changes structure (invisible) or design.

## 5. Writing alt text (only when no text is supplied)
1-2 sentences giving the purpose in context; no "image of"; include words shown in the image; charts state
the takeaway; complex figures: short alt text + detail in the notes (ask first). Decorative = removing it
loses no information. Equations and native tables are read directly. Leave the Title field empty.
Tell the user: NVDA has been reported to skip group alt text in Slide Show (it reads the pieces' text);
JAWS and Narrator read it. PDF export turns layout/master objects and footers into artifacts.

## 6. Troubleshooting
| Symptom | Fix |
|---|---|
| `error: ...` | No half-done edits are kept. Fix the cause; don't loop. |
| "not a registered working copy" | `open_working_copy` first, or pass `deck=<full path>`. |
| `ok:false` + visual change | The change would be visible: rolled back. Use the alternative the tool names. |
| `promote_to_title` fails | `add_hidden_title` with the heading's text. |
| `table_to_shapes` fails | Gradient/picture cell fills or an animated table: ask (unmerge or keep). |
| PowerPoint still flags "Check reading order" | `reading_order_snapshot`; if the order is right, report it as verified (heuristic). |
| `compare_playback` differs | Look at the image; re-run if an overlay covered the slide; otherwise report it. |

References: `references/tools.md` (all tools, plan format), `references/reading-order.md` (rules and how the
restack stays invisible), `references/wcag-rules.md` (titles, contrast, tables, links: rules and checker
behaviour), `references/alt-text-rules.md`, `references/animation-preserving-grouping.md`.
