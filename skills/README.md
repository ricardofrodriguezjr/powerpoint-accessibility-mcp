# Skills

A skill is a folder with a `SKILL.md` file: plain-language instructions that teach your AI app how to do a
specific job with the tools of this MCP server. Claude Code, Codex and OpenCode all read the same format.

The installer copies every skill in this folder to the right place for each app it sets up. To install one
by hand, copy its folder (the folder itself, not just `SKILL.md`) into your app's skills folder; see
[Manual setup](../docs/manual-setup.md).

| Skill | What it does | Use it when |
|---|---|---|
| [`pptx-wcag-alt-text`](pptx-wcag-alt-text/SKILL.md) | Takes any PowerPoint deck to WCAG 2.1 AA / Section 508 and leaves it ready to teach with: alt text or "decorative" on every picture, chart and diagram (grouping multi-piece figures without losing their animations), a real and unique title on every slide, correct screen-reader reading order, sufficient text contrast, accessible tables and descriptive link text. Works on an `_accessible` copy next to the original, keeps the look and animations, logs every change, and checks the result with its own audit plus PowerPoint's Accessibility Checker. | You want a deck (or a folder of decks) made accessible, audited, or finished after a partial pass; you have an alt-text spreadsheet to apply. |

The name `pptx-wcag-alt-text` is kept from the first version, which only handled alt text. The skill now
covers every WCAG item listed above.

## What is inside

```text
pptx-wcag-alt-text/
├── SKILL.md                                  the workflow the AI follows, step by step
└── references/
    ├── alt-text-rules.md                     how to write good alt text, and when to mark decorative
    ├── wcag-rules.md                         the WCAG rules checked, with the exact thresholds
    ├── reading-order.md                      how reading order is decided and fixed invisibly
    ├── animation-preserving-grouping.md      how figures are grouped without losing animations
    └── tools.md                              every MCP tool and the plan format
```

## Writing your own skill

Skills that build on this server are welcome (for example one tuned to your department's slide template).
Start from a copy of `pptx-wcag-alt-text`, give it a new `name` in the `SKILL.md` front matter, keep the
description specific about when it should be used, and open a pull request.
