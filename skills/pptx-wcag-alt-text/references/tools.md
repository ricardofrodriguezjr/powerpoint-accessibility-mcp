# ppt-a11y tools

Your app may show these names with a server prefix (for example `mcp__ppt-a11y__<name>` in Claude Code or
`ppt-a11y_<name>` in OpenCode).

All `slide` numbers are 1-based. All `ids` are PowerPoint `Shape.Id` values, which are shown in
`inspect_slide` and `slide_snapshot`. `part` is `slide` (the default), `layout` or `master`. For the
last two, `slide` says which slide's layout or master to edit. `deck` is optional: pass the full
path when more than one working copy is open.

## Decks

| Tool | Args | Notes |
|---|---|---|
| `open_working_copy` | `source, dest?, fresh?` | Creates or resumes `<name>_accessible.pptx` and opens it. This is the only way to make a deck editable. |
| `open_deck` | `path, read_only=true` | For comparing with an original. |
| `list_open_decks` | | |
| `save_deck` / `revert_to_saved` / `close_deck` | `deck?` / `deck?` / `deck?, save?, discard?` | `close_deck` refuses to drop unsaved work unless `discard=true` (ask the user first). |

## Looking

| Tool | Args | Returns |
|---|---|---|
| `deck_overview` | | Per slide: title, layout, effect count, todo list (`#id status`). |
| `inspect_slide` | `slide, part, inherited=true` | Per shape: id, name, kind, alt, title, decorative, left/top/w/h (points), z, needs_alt, text, placeholder, animation_effects, animation_trigger, members (groups). Plus `inherited` layout/master objects and `effects_total`. |
| `slide_snapshot` | `slide, part, ids?, nested?, width?` | An image with `#id` boxes, and a text legend. |
| `list_animations` | `slide` | The ordered effects: seq, order, shape_id, effect, start, delay, duration, paragraph. |

## Editing

| Tool | Args | Notes |
|---|---|---|
| `check_grouping` | `slide, ids, part` | `ok`, `blockers` (placeholders, tables, shapes inside a group, existing groups with their own animation), `warnings` (stacking, reading order), `animation_effects_to_rebuild`. Changes nothing. |
| `group_shapes` | `slide, ids, alt_text? \| decorative?, name?, part, allow_visual_change=false` | Returns `ok`, `group_id`, `rebuild`, `xml_check`, `visual_check` (with a compare image), and `failures` with `reverted: true` when it fails. Saves on success. |
| `set_alt_text` | `slide, ids, text, part` | Clears the decorative flag. Returns lint warnings. |
| `mark_decorative` | `slide, ids, part, decorative=true` | Clears the alt text. |
| `clear_alt_text` | `slide, ids, part` | Clears the description and title. |
| `restore_animations_from` | `slide, source_deck, source_slide?, replace?, restore_positions?` | Re-creates the source slide's animations on the same ids. Verified against the source XML. |
| `flatten_group` | `slide, group_id, allow_alt_loss?` | Splits an existing group that is animated as a whole (Appear/Fade only). Every piece gets the same effect, starting together. Verified with a timeline model, the render, and that pieces with a text frame (even an empty one) animate as one object. Use it before grouping the bigger figure, then run `compare_playback` against the original (required). |
| `placeholder_to_picture` | `slide, shape_id` | Replaces a picture in a content placeholder with an identical ordinary picture, so it can be grouped. Returns `new_id` (also recorded, so `'@<old id>'` works in later plans). Uses the clipboard. |
| `move_to_slides` | `slide, ids, part='layout'\|'master'` | Moves informative layout/master objects onto every slide that shows them, behind the slide content, then removes them from the layout. Returns `{slide: {old: new}}`. Uses the clipboard. |
| `apply_plan` | `plan_path \| steps, sheet_path?, rows?, stop_on_failure=true` | Runs plan steps in order, each a verified transaction. Format is below. |
| `compare_playback` | `slide, source_deck, source_slide?` | Plays the original and the working copy in Slide Show, captures a frame after each click, and compares only the slide area. Needs the screen. |
| `lint_alt_text` | `text` | |

`set_alt_text`, `mark_decorative` and `clear_alt_text` are not saved automatically: call `save_deck`
after a batch.

## WCAG beyond alt text (v2)

Every edit below is its own transaction (checkpoint, change, verify, save or roll back).

| Tool | Args | Notes |
|---|---|---|
| `title_report` | | `missing` (with `candidate` heading id/text, `layout_has_title`), `duplicates` (slide groups). |
| `promote_to_title` | `slide, shape_id` | Visible heading → real title placeholder, same look (render-verified; other slides on a changed layout verified too). Refuses animated headings. |
| `add_hidden_title` | `slide, text` | Off-slide title placeholder (not shown in Slide Show). For repeated headings or slides with no heading. |
| `set_slide_title` | `slide, text` | Keeps formatting; `warnings` if it wraps, grows or overflows. Refuses empty, hidden or decorative title placeholders. Visible change. |
| `make_titles_unique` | `slides?, pattern="{title} ({n} of {N})"` | Every run of consecutive duplicates, whole runs only, one transaction per run. Visible change. |
| `table_report` | | Per table: size, `merged_regions`, `header_row`, `hierarchy`, `recommended`. |
| `table_to_shapes` | `slide, table_id, reading=auto\|depth_first\|row_major` | Layout/diagram tables → text boxes (tagged as one reading block) + decorative border group. Verified allowing 1-px antialiasing. |
| `unmerge_table` | `slide, table_id, repeat_text=true` | Data tables. Visible change. |
| `set_table_header` | `slide, table_id` | Header Row on, look kept (verified). |
| `link_report` | | Per link: `unclear` + `why` (fails 2.4.4 / flagged), `advice` (passes AA), `suggested_text`. |
| `set_link_text` | `slide, shape_id, old_text, new_text, address?, cell?` | Keeps or adds the link target; verified. |
| `contrast_report` | `slides?` | Per run: ratio, needed, text/background colour, `bg_from` fill or render. |
| `fix_contrast` | `slides?, strategy=auto\|text` | Nearest passing colour (or fill); re-measured; `changed` lists from/to/ratio/ΔE. |
| `reading_order_report` | `slides?, compare_to?` | Per slide: `reads_now`, `suggested_reads`, `matches_suggestion`, `issues`, `changed_vs_original`. JSON + TXT report. |
| `reading_order_snapshot` | `slide, suggested=false` | Numbered badges in reading order (green title, blue content, grey footer). |
| `set_reading_order` | `slide, order, strict=false` | Every readable top-level id once. Invisible restack, verified; `deviation` + `not_accepted` if overlaps forced the closest invisible order (re-submit exactly `order_now` to accept it). Otherwise records the order as accepted. |
| `fix_reading_order` | `slides?, orders?, use_checker=true, hidden_text="decorative"` | Hidden text → decorative; suggested (or given) order per slide; PowerPoint checker loop over variants. `checker.still_flagged` lists what PowerPoint still asks to check. |
| `wcag_audit` | `compare_to?, run_checker=true, check_contrast=true` | Everything left per category + PowerPoint's counts; `done`. JSON report. |
| `record_change_log` | `changes=[{Slide, Category, WCAG, What changed, Kind, Before, After, Verified}], dest?, sheet_name="WCAG changes"` | Change log as its own sheet of the results copy (`dest`), or by default a new workbook `<deck>_changes.xlsx` in the deck's checks folder. Returns the path. Each call **rewrites** that sheet: pass every change at once. Refuses a sheet that is not a change log. |

`group_shapes` and `move_to_slides` results now carry a `reading_order_note`: both change the stacking
order, which is why reading order is fixed last.

## Checking and recording

| Tool | Args | Notes |
|---|---|---|
| `audit_alt_text` | `include_inherited=true, include_ok=false` | Writes JSON and CSV reports into `<deck>_a11y_checks/`. |
| `run_accessibility_checker` | `itemize=["Missing alt text"]?` | Uses PowerPoint's Accessibility Assistant through UI Automation. Itemising takes about 1 s per issue and changes the selection only. |
| `read_alt_text_sheet` | `path, sheet_name?` | Returns rows with an Excel `row` number. |
| `record_sheet_results` | `source, dest, updates, sheet_name?` | Adds new columns in the copy `dest`. Input columns are protected. |
| `evaluate` | `code, slide?` | Inspection escape hatch for developers: globals `presentation`, `slide`, `find(container, id)`; set `result`. Saving, closing and opening are blocked, and changes to other open decks are reported. Not a sandbox. Off by default (the server needs `PPT_A11Y_ENABLE_EVALUATE=1`); never needed for remediation. |

## Plan format (`apply_plan`)

```json
[{"row": 2,  "op": "move_to_slides", "slide": 1, "part": "layout", "ids": [12]},
 {"row": 2,  "op": "alt", "slide": 1, "ids": ["@12"], "text": "ROW"},
 {"row": 3,  "op": "decorative", "slide": 2, "part": "master", "ids": [11, 21, 22, 23]},
 {"row": 11, "op": "placeholder_to_picture", "slide": 11, "id": 10},
 {"row": 11, "op": "group", "slide": 11, "ids": ["@10", 11, 12], "alt_text": "ROW", "name": "Outlook screenshot"},
 {"row": 16, "op": "flatten_group", "slide": 17, "id": 24},
 {"row": 25, "op": "clear_alt", "slide": 27, "ids": [30]}]
```

- `ops`:
  - `group`
  - `alt`
  - `decorative`
  - `clear_alt`
  - `flatten_group`
  - `placeholder_to_picture`
  - `move_to_slides`
  - v2, slide-level: `promote_title` (`id`), `hidden_title` (`text`), `set_title` (`text`),
    `table_to_shapes` (`id`, `reading?`), `unmerge_table` (`id`), `table_header` (`id`),
    `link_text` (`id`, `old_text`, `new_text`, `address?`, `cell?`), `reading_order` (`ids` in order)
  - v2, deck-level (no `slide`; optional `slides`): `titles_unique` (`pattern?`), `contrast`
    (`strategy?`), `fix_reading_order` (`orders?`, `use_checker?`, `hidden_text?`)
  - `note` (ignored; for comments)

```json
[{"op": "promote_title", "slide": 4, "id": 4},
 {"op": "hidden_title", "slide": 8, "text": "Topic 1: Introduction"},
 {"op": "titles_unique"},
 {"op": "table_to_shapes", "slide": 4, "id": 5},
 {"op": "link_text", "slide": 15, "id": 3, "old_text": "https://...", "new_text": "pandas Essential Training: Using read_csv (LinkedIn Learning)"},
 {"op": "contrast"},
 {"op": "fix_reading_order"}]
```
Steps that already ran are skipped on a re-run (title present, table gone, link text replaced).
- `"@N"` means "the shape #N became on this slide" after an earlier `move_to_slides` or
  `placeholder_to_picture` step. When an id exists in more than one part, use `"@layout:N"`,
  `"@master:N"` or `"@slide:N"`; an ambiguous `"@N"` is refused. Re-running a chunk skips structural
  steps that already ran. The mapping is kept in `<deck>_a11y_checks/id_map.json`, so chunks can
  run separately.
- `"ROW"` copies that spreadsheet row's Alt Text verbatim. Pass `sheet_path` to use it.
