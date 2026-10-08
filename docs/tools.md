# Tool reference

The `ppt-a11y` server exposes 46 tools. You normally never call them yourself: the skill tells your AI app
which ones to use. They are listed here for transparency and for contributors. The skill's own reference,
with arguments and plan formats, is [`skills/pptx-wcag-alt-text/references/tools.md`](../skills/pptx-wcag-alt-text/references/tools.md).

## Decks

| Tool | What it does |
|---|---|
| `open_working_copy(source)` | Copies the deck to `<name>_accessible.pptx` beside it and opens that copy. Only working copies can be edited or saved. |
| `open_deck`, `list_open_decks`, `save_deck`, `revert_to_saved`, `close_deck` | Deck housekeeping. |
| `deck_overview` | One line per slide: title, layout, what still needs work, animation count. |
| `inspect_slide(slide, part)` | Every shape: id, kind, alt text, decorative flag, text, position, stacking, animations, group members, objects inherited from the layout or master. |
| `slide_snapshot(slide, part, ids)` | PowerPoint's own render of the slide with a labelled box per shape. |

## Alt text and grouping (WCAG 1.1.1)

| Tool | What it does |
|---|---|
| `set_alt_text`, `mark_decorative`, `clear_alt_text` | Single-object edits. `part='layout'` or `'master'` edits Slide Master objects. |
| `check_grouping(slide, ids)` | Dry run: blockers, effects that will be rebuilt, stacking warnings. Changes nothing. |
| `group_shapes(slide, ids, alt_text or decorative)` | Groups and keeps every animation. Verified, rolled back on any difference. |
| `flatten_group(slide, group_id)` | Splits a group animated as a whole (Appear or Fade) so it can join a bigger figure. |
| `placeholder_to_picture(slide, id)` | Turns a picture held by a placeholder into an identical ordinary picture, so it can be grouped. |
| `move_to_slides(slide, ids, part)` | Moves informative layout or master content onto the slides that show it (screen readers skip layout content). |
| `restore_animations_from(slide, source_deck)` | Re-creates a slide's animations from another copy, for example after someone grouped by hand. |
| `list_animations(slide)` | The animation sequence in order. |
| `compare_playback(slide, source_deck)` | Plays original and copy in Slide Show and compares a frame after every click. |
| `audit_alt_text`, `lint_alt_text` | Deck-wide alt-text audit, and quality warnings for a proposed description. |
| `read_alt_text_sheet`, `record_sheet_results` | Reads an alt-text spreadsheet; writes results into new columns of a copy. |
| `apply_plan(plan_path, sheet_path, rows)` | Runs a JSON plan of steps, each a verified transaction. |

## Titles, reading order, contrast, tables, links

| Tool | What it does |
|---|---|
| `title_report`, `promote_to_title`, `add_hidden_title`, `set_slide_title`, `make_titles_unique` | Real, unique slide titles with an identical look. |
| `reading_order_report`, `reading_order_snapshot` | What a screen reader reads on each slide, the suggested order, and numbered badges on the slide. |
| `set_reading_order`, `fix_reading_order` | Invisible, verified restack to a given or suggested order. |
| `contrast_report`, `fix_contrast` | WCAG contrast against the rendered background; nearest passing colour or fill. |
| `table_report`, `table_to_shapes`, `unmerge_table`, `set_table_header` | Merged cells and header rows. |
| `link_report`, `set_link_text` | Descriptive link text that keeps the link. |

## Checking and reporting

| Tool | What it does |
|---|---|
| `wcag_audit` | Everything left per category plus PowerPoint's Accessibility Checker; `done: true` when nothing is left. |
| `run_accessibility_checker(itemize)` | Runs PowerPoint's own Accessibility Assistant and returns counts per category. |
| `record_change_log` | Writes the change log (one row per change, with before and after values) to a sheet of the results workbook. |
| `evaluate` | Inspection escape hatch for advanced use: runs Python against an open working copy. It refuses to save, close or open decks. It is not a sandbox. |

## Command line

Every tool also runs from a terminal, without an AI app:

```powershell
cd $env:USERPROFILE\.ppt-a11y\app
.venv\Scripts\python.exe cli.py list
Set-Content args.json '{"deck": "C:\\Slides\\Lecture 05_accessible.pptx"}'
.venv\Scripts\python.exe cli.py wcag_audit "@args.json"
```
