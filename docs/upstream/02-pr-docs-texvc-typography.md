# docs(skill): record the measured Texvc subset and Chinese inline-math typography

## What

- `skills/mathtype-for-word/references/academic-equation-style.md`
  - **MathType carrier notes**: how the LaTeX-facing rules translate to MathType — `\mathbf` selects
    the Vector character style (set it to `Times New Roman,BI` for bold italic), `\operatorname` and
    `\widebar` are outside the Texvc subset, matrices reuse the Vector style.
  - **Measured Texvc subset**: commands that render correctly, commands echoed as literal red text,
    commands rendered as a black diamond, and constructs dropped silently. Measured on MathType 7.8
    and 7.11.1.462 with Word 16.
  - **Substitutions** that keep the meaning, for each failure mode.
  - **Chinese inline-math rules**: `其中，` explanation paragraphs get zero first-line indent
    (`w:ind/@w:firstLine="0"`, no `w:firstLineChars`); inline `Equation.DSMT4` runs use Normal
    character position (`w:position="0"`); the paragraph gets `w:textAlignment="center"`; exact or
    fixed line spacing is replaced with At Least at the body line height.
- `skills/mathtype-for-word/references/troubleshooting.md`: three rows for the Texvc failure modes.
- `skills/mathtype-for-word/SKILL.md`: validation gates for `其中，` and inline-equation paragraphs;
  the failure-handling bullet now points at the subset table.

## Why

Toggle TeX accepts unknown commands instead of failing: it still creates the object, so object counts
and number fields look correct while the equation renders as red text, a black diamond, or loses a
symbol. The current reference mentions the subset only in general terms. The measured table makes the
failure predictable, and the substitution table gives the writer a fix instead of a retry.

The Chinese inline rules cover two visible defects that structural validation does not catch: a body
style's two-character indent leaking into `其中，` paragraphs, and inline OLE objects being clipped by
exact line spacing or misaligned by a non-Normal character position.

## Verification

- `pwsh -File tests/run-tests.ps1`: 105 tests, same two environment-only failures as `main`
  (`test_shared_plugin_launcher` needs a `python` that is not the Windows Store alias).
- `dist/mathtype-for-word.skill`, `dist/mathtype-for-word-plugin.zip`, and their SHA-256 files rebuilt
  because the packaging contract compares them against the working tree.

## Notes

The claims about enforcement were removed on purpose: `render` and `validate` do not check the Texvc
subset on `main`, so the document tells the user to render and inspect the pages instead. If you want
that automated, the fork has a `Get-UnsupportedTexCommand` pre-flight plus a `stored_tex` readback;
I can send it as a separate PR.
