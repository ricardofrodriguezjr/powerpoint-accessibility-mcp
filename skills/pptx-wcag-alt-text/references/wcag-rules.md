# Titles, tables, contrast and links: what compliant means here

Measured behaviour is from PowerPoint for Microsoft 365, build 16.0.20430 (October 2026).

## Slide titles (2.4.2 as applied to documents, 2.4.6, 1.3.1; checker: Missing slide title = Error,
## Duplicate slide title = Tip)

- A title is a placeholder of type title / center title / vertical title (1, 3, 5) **with text**. A text box
  that looks like a title is not one. An empty title placeholder counts as missing.
- A title placed off the slide counts (Microsoft's own "Add Hidden Slide Title" does this); it is not shown
  in Slide Show. Use it only when no visible heading can become the title, or the visible heading repeats on
  several slides.
- `promote_to_title` copies geometry, fill/outline (PickUp/Apply), every text-frame, paragraph and run
  setting (line rules before spacing, AutoSize last) and links onto the title placeholder, then deletes the
  heading. Deleting a placeholder that holds text leaves an empty copy behind (as in the UI); the tool
  removes it too. On a slide whose layout has no title placeholder, one is added to the layout (layout
  placeholders are prompts only: no slide changes). `Shapes.AddTitle` is refused on "Blank"-type layouts;
  re-applying the layout instantiates the title placeholder instead.
- Unique titles: Section508.gov "every slide needs a unique title" (patterns: "Topic: Subtopic", "Topic
  (Continued)", "Topic - Part 1", "Topic (Slide # / Total #)"); HHS: "(i.e. # of #). Multiple 'Cont.' is not
  sufficient." Default: "Title (n of N)".

## Tables (1.3.1; checker: Missing table header = Error, Use of merged or split cells = Warning)

- Missing table header fires exactly when Header Row (`Table.FirstRow`) is off.
- PowerPoint's merged-cell warning fires only on tables of at least 4x4 with a merge outside row 1 /
  column 1 (measured). The HHS PowerPoint 508 checklist fails **any** merged or split cell, and
  Section508.gov's authoring guidance says to avoid them (WCAG 1.3.1 itself does not ban them as such):
  our `table_report` finds them all.
- PowerPoint cannot express scope/headers, so a table whose meaning depends on spanning cells cannot meet
  1.3.1 as a table. Layout/diagram tables (a tree drawn with merged cells) should not be tables at all
  (Microsoft, Texas HHS, ICT baseline 12.C) → `table_to_shapes` (same look, real text, read branch by
  branch; borders as a decorative group). Data tables → `unmerge_table` (Split with the exact span keeps the
  grid; text repeated in each cell) and Header Row on.
- Turning Header Row on restyles a styled table; `set_table_header` puts every fill and text colour back.

## Text contrast (1.4.3 AA)

- Ratio (L1 + 0.05) / (L2 + 0.05) with WCAG relative luminance, never rounded (4.499 fails).
- 4.5:1, or 3:1 for large text: >= 18 pt, or >= 14 pt bold (Section508.gov applies PowerPoint's point size
  directly). With "shrink text on overflow" the rendered size can be smaller, so the tool only grants the
  large-text threshold from 24 pt / 18.7 pt bold.
- Exempt: logos, pure decoration, text nobody can see.
- PowerPoint's checker only evaluates text on solid opaque fills (older builds) and skips tables without fill
  and chart text; it misses e.g. orange #E5751F text on a white slide (2.97:1). Our check uses the cell/shape
  fill (theme tints resolved) when it is solid, opaque and renders as that colour; otherwise every
  significant colour PowerPoint actually renders behind the run (gradients, pictures, shapes below), and the
  run must pass against all of them. Layout/master text and SmartArt are included; chart text is listed as
  not machine-checked. Table-cell text bounds are relative to the cell, so cells are sampled as a whole.
- Fix: the nearest colour that passes, same hue (OKLCH lightness, chroma reduced only to stay in sRGB),
  keeping light/dark polarity when possible, 0.08 above the threshold for rounding. Strategy "auto" also
  considers a darker/lighter cell or shape fill and picks it when that is a much smaller change (all text on
  that fill is re-checked). Example from a university template: burnt orange #E5751F needs #E17119 for large text on
  white, #B95800 for normal text; white text needs the orange fill darkened to #BB5900 and teal #508590 to
  #487D88.

## Links (2.4.4 A)

- A raw URL is "generally not sufficiently descriptive" (G91) and fails when nothing in the same paragraph
  explains it (F63). With an explaining sentence in the same paragraph it passes in context (G53).
- PowerPoint flagged a 98-character URL with a query string standing alone in its paragraph, and did not
  flag "Go to https://www.anaconda.com/download/success".
- Fix: descriptive visible text, preferably the destination page's title. A ScreenTip alone is not enough.
- Replacing the characters drops the link; `set_link_text` re-applies Address / SubAddress / ScreenTip to the
  new characters and verifies it. A link with no target (`ppaction://noaction`) whose text is a URL gets that
  URL as its address.
