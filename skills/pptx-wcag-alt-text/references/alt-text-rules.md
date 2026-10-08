# Alt-text rules: what "compliant" means here

## WCAG 2.1 SC 1.1.1 Non-text Content

Every non-text object needs a text alternative that serves the same purpose. If it is purely
decorative, it must be ignorable by assistive technology.

- Informative images: give a short description of their purpose in context. https://www.w3.org/WAI/tutorials/images/
- Images of text: the alt text contains the same words. https://www.w3.org/WAI/tutorials/images/textual/
- Complex images (charts, diagrams): give a short alt text, plus a long description on the slide or in
  the notes (G74). https://www.w3.org/WAI/WCAG21/Techniques/general/G74
- Groups of images: one item carries the alt text for the set, and the others are ignored (G196).
  https://www.w3.org/WAI/WCAG21/Techniques/general/G196
- Decorative objects must not carry text (F39). In PowerPoint that means ticking *Mark as decorative*
  and leaving alt text empty.
- Meaning conveyed only by an animation must also exist as text (the alt text can say "builds in N steps: ...").

## PowerPoint Accessibility Checker (Microsoft 365 "Accessibility Assistant")

Rules: https://support.microsoft.com/en-us/office/rules-for-the-accessibility-checker-651e08f2-0fc3-4e10-aaca-74b4a67101c1

What "Missing alt text" flags, as measured on the test deck (our audit matched it 140/140):

- Pictures, including pictures inside content placeholders.
- Groups without alt text (children are not flagged when the group has alt text).
- Charts, SmartArt, media, OLE objects, icons/graphics, 3D models and ink.
- **Shapes without text**: rectangles, arrows, freeforms, lines, connectors, and **text boxes that
  contain only spaces**.
- Alt text that is a file name or URL (e.g. `preencoded.png`, `https://...`).

Not flagged:
- Objects on slide layouts and the slide master. The checker does not scan them.
- Decorative objects.
- Shapes with real text.
- Native tables (these are checked for a header row instead).
- Equations.

Other categories it reports, which are outside SC 1.1.1 and should be reported to the user:
- Missing slide title, duplicate slide title;
- check reading order;
- hard-to-read text contrast;
- merged or split table cells;
- missing table header;
- missing captions.

## Writing guidance

- Microsoft recommends 1 or 2 sentences. There is no hard length limit, but over about 250 to 300 characters
  consider moving detail to the notes.
  https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2
- Don't start with "image of" or "picture of". The screen reader already announces "graphic" or "group".
- Avoid file names, default names ("Picture 3"), and "Description automatically generated".
- Use the Description field only. Leave the Title field empty: old PDF export merged the two fields,
  and WebAIM recommends leaving it blank.

## Screen reader and export caveats (tell the user, don't block on them)

- **Layouts/master:** screen readers generally skip objects that come from the layout or master, and
  in PDF export they become artifacts. Logos and bars there should be decorative. Content tiles there
  carry information only visually; offer to move them onto slides if that matters.
- **Group alt text:** JAWS and Narrator read it. NVDA was reported to skip group alt text in Slide Show
  (nvaccess/nvda#16161, closed without a fix). Edit-view reading works.
- **PDF export:** a group with alt text becomes one Figure. WebAIM reports children's text can also
  appear, so test with a PDF checker if PDFs are distributed.
- **Reading order** is the z-order (Selection pane, bottom to top). Grouping places the group at the
  front-most member's position; `move_to_slides` puts moved objects at the back (read first). Decorative
  objects drop out of the reading order (NVDA still reads the text of decorative text shapes). Fix it last
  with `fix_reading_order`; see `reading-order.md`.
