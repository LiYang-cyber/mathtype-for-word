# feat(tools): add an inline-math layout checker

## What

- `scripts/normalize_word_inline_math_layout.py`
  - `--check` is read-only: it reports `其中，` explanation paragraphs with a nonzero first-line
    indent, and paragraphs that contain an inline `Equation.DSMT4` object but lack
    `w:textAlignment="center"`, a Normal character position (`w:position="0"`) on the object run, or
    At Least line spacing at the body line height. Exit code 1 when anything is reported.
  - With an output path it fixes those properties, rewrites `document.xml` through a temporary
    sibling with an atomic replace, and verifies that every `word/embeddings/` part is byte-identical
    before and after.
- `tests/test_normalize_layout.py`: check, patch, and OLE-preservation cases.
- `skills/mathtype-for-word/SKILL.md`: step 7 requires the check for a DOCX that contains inline
  equations, and reference routing points at the script.

## Why

Word clips an inline OLE object under exact or fixed line spacing, and a `其中，` paragraph inherits
the body style's two-character first-line indent unless the direct formatting clears it. Both defects
survive structural validation because the object and field counts are correct; today the only signal is
a human looking at the page.

The checker works on the DOCX package, so it never opens Word and never rebuilds the OLE objects.

## Dependency

The checker imports `lxml`; the test skips when it is missing, so the repository test run stays
dependency-free. If you would rather avoid the dependency, I can rewrite it on `xml.etree`, or fold the
same rules into the bridge as a PowerShell action (the bridge already edits `document.xml` for the
inline-layout fixes in this fork).

## Verification

- `pwsh -File tests/run-tests.ps1`: 108 tests (3 new), same two environment-only failures as `main`
  (`test_shared_plugin_launcher` needs a `python` that is not the Windows Store alias).
- Without `lxml` the three new tests skip; with `lxml` 6.1.3 they pass.
- `dist/mathtype-for-word.skill`, `dist/mathtype-for-word-plugin.zip`, and their SHA-256 files rebuilt
  because the packaging contract compares them against the working tree.

## Notes

`lxml` keeps the `w:` namespace prefixes and allows inserting `w:ind`, `w:spacing`, `w:textAlignment`
and `w:position` in schema order. `xml.etree` would rewrite prefixes, which is a needless risk for a
package Word has to reopen.
