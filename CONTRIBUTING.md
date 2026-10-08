# Contributing

Thank you for helping make course materials accessible.

## Reporting a problem

Open an [issue](https://github.com/ricardofrodriguezjr/powerpoint-accessibility-mcp/issues) with:

- your Windows version and PowerPoint version (File > Account > About PowerPoint);
- which app you use (Claude Code, Codex or OpenCode) and its version;
- what you asked for and what happened;
- the report folder `<deck>_accessible_a11y_checks` if you can share it.

Please do not attach slides you cannot share publicly. A small deck that shows the same problem is ideal.

## Changing the code

1. Fork the repository and create a branch.
2. Set up the environment (see [Manual setup](docs/manual-setup.md)).
3. Make your change. Every tool that edits a deck must:
   - work only on a registered working copy, never the original;
   - save a checkpoint first, verify the result (render, positions, animation XML), and roll back on any
     difference;
   - report visible changes explicitly, with before and after values.
4. Run the tests in [`tests/`](tests/README.md). `wcag_v2_tests.py` and `stress_animations.py` must print
   `RESULT: PASS`.
5. Open a pull request that explains what changed and how you verified it.

## Changing the skill

The skill in [`skills/pptx-wcag-alt-text`](skills/pptx-wcag-alt-text) is written for an AI reader: short,
direct rules, in the order they should be applied. If you change what a tool does, update the skill and
its `references/tools.md` in the same pull request.
