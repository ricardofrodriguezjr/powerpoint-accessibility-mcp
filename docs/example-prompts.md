# Example prompts

Copy a prompt, change the path to your own file, and paste it into Claude Code, Codex or OpenCode. You do
not need special wording: any request to make a PowerPoint accessible or WCAG compliant starts the skill.

**Tip:** to get the full path of a file, hold **Shift**, right-click the file in File Explorer and choose
**Copy as path**. Paste it with the quotes.

Every prompt below leaves your original file untouched and writes the result next to it as
`<name>_accessible.pptx`, plus a folder `<name>_accessible_a11y_checks` with the report.

## Make one deck accessible

```text
Make "C:\Users\me\Documents\STAT 101\Lecture 05.pptx" WCAG 2.1 AA compliant.
Save the result as an _accessible copy in the same folder and give me a short
summary of what you changed.
```

## Use alt text you already wrote

If you (or a TA) wrote image descriptions in a spreadsheet, the skill uses them word for word:

```text
Make "C:\Users\me\Documents\STAT 101\Lecture 05.pptx" accessible. Use the alt text
in "C:\Users\me\Documents\STAT 101\Lecture 05 alt text.xlsx" exactly as written
(one row per figure). Put the _accessible copy in the same folder.
```

Your spreadsheet should be an Excel file (.xlsx) with column names in row 1 and one row per figure,
including the slide number and the description. A column named `Alt Text` works best. If your descriptions
are in a .csv file or a Word table, save them as .xlsx first.

## Do a whole folder

```text
Make every PowerPoint in "C:\Users\me\Documents\STAT 101\Slides" accessible,
one deck at a time. Skip files that already end in _accessible. Save each
_accessible copy next to its original and give me one summary table at the end.
```

Only .pptx and .pptm files can be processed. For older .ppt files, open them in PowerPoint, choose
**File > Save As > PowerPoint Presentation (.pptx)**, then run the prompt.

## Check first, change nothing

```text
Check "C:\Users\me\Documents\STAT 101\Lecture 05.pptx" for accessibility problems.
Do not change anything yet. List what is missing per slide and what you would fix.
```

## Review what changed

```text
Open the change log in "C:\Users\me\Documents\STAT 101\Lecture 05_accessible_a11y_checks"
and list every change I can see on the slides, with before and after values.
```

## Continue later

Work is saved after every verified step, so you can stop and pick up again:

```text
Continue making "C:\Users\me\Documents\STAT 101\Lecture 05.pptx" accessible.
The _accessible copy is already in the same folder.
```

## Calling the skill by name

Most of the time the app picks the skill by itself. To call it explicitly:

| App | Type |
|---|---|
| Claude Code | `/pptx-wcag-alt-text` followed by your request |
| Codex | `$pptx-wcag-alt-text` followed by your request |
| OpenCode | `@pptx-wcag-alt-text` (OpenCode 2) or `/pptx-wcag-alt-text` (OpenCode 1), followed by your request |

## What to expect

- The app opens PowerPoint on your computer and works on the copy. Leave PowerPoint open while it works,
  and do not edit the `_accessible` file at the same time.
- A 40-slide lecture with diagrams and animations typically takes 10 to 30 minutes.
- At the end you get a summary: what was fixed in each category, every visible change with before and
  after values, and anything that needs your decision (for example a chart whose description only you can
  write).
- If the deck has animations, the app may ask for a few minutes of a quiet screen to play the slides and
  confirm every animation still plays exactly as before.
