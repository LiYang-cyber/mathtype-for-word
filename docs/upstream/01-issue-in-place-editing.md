# Feature request: edit existing MathType equations in place

## Problem

The skill can create equations from `{{MATH:...}}` markers, and it can replace a document whose draft
still holds markers. It has no way to change an equation that is already a native `Equation.DSMT4`
object inside the DOCX. The current guidance is to make a working copy, insert markers at the intended
ranges, and render.

That path re-creates the marked equations from TeX. A whole-document toggle re-creates every object it
touches, so untouched equations lose their original OLE bytes and layout, and a display-flagged inline
symbol comes back as a display equation. On a long Chinese engineering document this changes the pagination
and the paragraph formatting around the equations.

## Proposal

Two bridge actions, plus one read-only preview action:

1. **`inventory`** — read-only. Lists every `Equation.DSMT4` in document order with its 1-based ordinal,
   paragraph number, numbered state, cached TeX, and layout (`inline` / `display`).
   `-ReadTexFromMathType` recovers TeX for objects that store no cache, by toggling a throwaway copy.
2. **`edit`** — takes a schema v1 edit manifest: `index`, `old_tex`, one of `new_tex` or `delete`,
   plus optional `prose` exact find/replace entries. It toggles out only the listed equations, verifies
   `old_tex`, rewrites, and toggles back with the delimiter that equation originally used, so the
   inline/display property survives. A mismatch aborts before publishing; the source is left intact.
3. **`preview`** — exports the DOCX to PDF and optional per-page PNGs through LibreOffice for visual
   inspection. Useful because structural checks cannot prove typography.

## Status

I have a working implementation in a fork, used on Chinese engineering documents with MathType 7.8.2.441
and 7.11.1.462 against Word 16: about 400 lines added to the bridge plus a separate editing module.
It was written against the 1.3-era bridge (`441daf4`), so it does not know about the 1.4.x features:
batched rendering, bridge lock, journaled preferences, orphan cleanup, `scan` / `mark`.

Before I port it onto current `main` and add tests, I would like to know:

- Is in-place editing in scope for this project?
- Would you prefer one PR (inventory + edit + preview) or a split, for example inventory first?
- Should `edit` reuse the render manifest schema with a `mode` field, or stay a separate document type
  as it is now?
- Is the `old_tex` verification design acceptable? It is a safety check against a stale index, not a
  security boundary.

I also prepared two smaller changes independently of this one:

- a documentation PR with the measured Texvc subset (renders / literal red text / black diamond /
  silently dropped) and the Chinese inline-math typography rules;
- a PR with a read-only inline-math layout checker (`--check`) plus a fix mode.

I can send the change inventory if it helps you judge the port cost.
