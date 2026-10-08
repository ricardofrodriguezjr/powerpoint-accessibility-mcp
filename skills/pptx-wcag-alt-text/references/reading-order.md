# Reading order (WCAG 1.3.2 Meaningful Sequence, 2.4.3 Focus Order)

## How PowerPoint decides what is read, and when

- ECMA-376 §19.3.1.45: the order of shapes in the slide's shape tree is both the z-order (first = back)
  and the navigation order. So **reading order = stacking order, back to front** (`slide.Shapes(1)` first).
  The Reading Order pane, Tab order, NVDA (source: `appModules/powerpnt.py`), JAWS, Narrator and tagged PDF
  export all follow it.
- Not read: objects marked decorative (except: NVDA still reads the *text* of a decorative text shape),
  objects on the slide layout/master, empty placeholders. Hidden objects (Selection Pane eye) are still read.
- A group is one stop: JAWS/Narrator read its alt text; NVDA in Slide Show reads the pieces in group order.
- Slide change: NVDA, VoiceOver and TalkBack announce the slide title first (from the title placeholder),
  then read the content in stacking order. So the title must also be first in the stacking order, or it
  is read twice in a confusing place.

## The rules our suggestion follows

1. The slide title, then a subtitle.
2. Everything else in visual reading order (recursive XY-cut): rows top to bottom, columns left to right
   inside a row, recursively. Columns that start with a heading over bigger content, or a figure over a
   short caption, are read column by column (heading/figure, then what belongs to it).
3. A call-out note whose arrow/line points into another object is read right after that object
   (`visual` variant: pure visual order without this rule).
4. Shapes our tools made as one unit (e.g. a table converted to text boxes, tags `A11Y_BLOCK`/`A11Y_SEQ`)
   stay together in their own order (branch by branch for a hierarchy).
5. Footer, date and slide number last.
6. Text that nobody can see (completely behind opaque objects) should not be read: mark it decorative.

Sources: WebAIM ("the title will be the first thing read"), Section508.gov (Title, then Subtitle), UGA
(title, subtitle/overview, main text, images/tables/charts, footer), Chaffey and Karen McCall (footer and
slide number last), Penn State (top-to-bottom-then-left-to-right is sometimes right). WCAG Understanding
1.3.2: "There may be more than one order that is correct"; only one correct order is needed.

## PowerPoint's "Check reading order" warning

Microsoft: a slide is listed "when the reading order of the objects on a slide doesn't match one of the
common ways objects are spatially ordered". The algorithm is undocumented and the warning may not clear after
a correct fix. Observed on Microsoft 365 16.0.20430:
- title, then subtitle, then a picture passed where title, picture, subtitle (top-to-bottom) was flagged;
- footer/slide number read before a picture was flagged on one slide, not on others;
- on code slides with notes in a left column and code on the right, "notes, then code" passed and
  "code, then notes" was flagged.
`fix_reading_order` therefore tries the valid variants (`rules`, `visual`, `placeholders`, `top_down`,
`minimal`) on flagged slides and keeps one PowerPoint accepts; if none is accepted, it keeps the `rules`
order and reports the slide.

## Why the restack is invisible

Changing the stacking order of two objects can only change the picture where their painted areas overlap.
`set_reading_order`:
1. computes each object's painted extent, clipped to the slide (rotated box if it has a fill/outline/
   picture/chart/table, its text bounds (text can overflow; PowerPoint reports them in slide space, rotation included), group children, outline
   weight, shadow/glow/soft edge, reflection; the whole slide if an animation can move, resize or spin it -
   anything but appear/fade/wipe/blinds/box/checkerboard/circle/diamond/dissolve/plus/random bars/split/
   strips/wedge/wheel/color reveal, or any x/y/width/height/rotation property animation);
2. puts the read objects into the requested order in the slots read objects already occupy (decorative and
   hidden objects keep theirs);
3. if overlapping objects would have to swap, it swaps only the read objects that must (and, when a
   decorative object such as a converted table's border lines is in the way, only the misplaced objects,
   e.g. the title); everything else keeps its stacking. That trial is kept only if neither object of a
   swapped pair moves during an animation, no other animated object is painted over their overlap (an exit
   or click reveal would show the stacking), and the slide renders identical (no channel changed by more
   than 2 levels, apart from a thin antialiasing edge a 1-px shift explains). Otherwise the closest order that
   keeps every overlapping pair as it was is used (reported as `deviation`). A Morph transition into or out
   of the slide keeps every object's stacking;
4. verifies: before/after `Slide.Export` render, positions of every shape, and the slide's saved animation
   XML (spTgt references ids, which a restack does not change); rolls back on any difference;
5. records the order as accepted (`reading_order_accepted.json`, keyed by SlideID; moved aside when the
   working copy is recreated) only if it is the order asked for and either follows the rules (title first,
   footers last) or was chosen by you with `set_reading_order`. A closest-invisible order is never accepted
   on its own: the slide stays in `slides_to_fix` / `needs_attention` until you re-submit exactly that order.

Hidden text: text completely behind static opaque objects (not animated, not triggers) is only marked
decorative after a render with that text hidden shows 0 changed pixels. A slide title or subtitle is never
marked decorative (the slide would lose its title); it is reported in `notes` instead.

Grouping (`group_shapes`) puts the group where its front-most piece was; `move_to_slides` puts moved
layout objects at the back (read first). Both are why reading order is fixed last.
