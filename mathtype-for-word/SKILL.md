---
name: mathtype-for-word
description: Create, edit, replace, number, cross-reference, update, inspect, or validate mathematical equations in Microsoft Word DOCX and PowerPoint PPTX using genuine MathType 7 Equation.DSMT4 objects. Use for DOCX or PPTX equation authoring, TeX-to-MathType conversion, editing equations that already exist inside a DOCX, consistent IEEE-style mathematical typography, academic Chinese or English prose around formulas, centered displays, MathType Word equation numbers such as (1), the “equation reference goes here” workflow, or repairs that must avoid Word OMath, captions, numbered lists, images, and fake text equations.
license: MIT
metadata:
  author: Jia-Ming Zhou (Felimet)
  fork_maintainer: LiYang
  version: 2.0.0
  upstream: https://github.com/felimet/mathtype-for-word
  upstream_base: "441daf4"
---

# MathType for Word and PowerPoint

Create real MathType equations in DOCX and PPTX, preserving MathType-native Word numbering and references. Treat source Office files as user data: write a new output unless the user explicitly authorizes in-place replacement.

## Non-negotiable invariants

- Equations are `Equation.DSMT4` inline OLE objects created by MathType Toggle TeX.
- Numbered displays use MathType `MACROBUTTON MTPlaceRef` and `SEQ MTEqn` fields.
- References use `MTCommand_InsertEqnRef`, its `MTReference` placeholder, and the target number's `MTPlaceRef` action. A resolved reference is a `GOTOBUTTON` containing a nested `REF` to a `ZEqnNum...` bookmark.
- The default number format is simple Arabic equation number only: `(1)`, `(2)`, `(3)`, with chapter and section components disabled.
- Apply numbering to new numbers and the whole processed document. Update fields after edits.
- Never substitute Word numbered lists, captions, `SEQ Equation`, OMath, Unicode-only text, images, or manually typed numbers. Use those only as an explicitly accepted, clearly labeled fallback when MathType is unavailable.
- PowerPoint equations are centered, floating `Equation.DSMT4` OLE objects. PowerPoint has no MathType-native equivalent of Word's equation-number/reference fields, so do not imitate them with typed numbers or links.
- An edit run touches only the equations named in its manifest. Never round-trip a whole document through TeX to save work: untouched equations keep their original OLE bytes and layout.
- Preserve each edited equation's own inline/display property by reusing the delimiter that the equation toggles out with. Converting a display-flagged inline symbol to an inline equation (or the reverse) silently reflows the paragraph.
- Verify the TeX that an equation currently holds against the manifest's `old_tex` before replacing it. Never publish an edit whose verification failed.
- Keep each symbol's mathematical role and font style consistent across display equations, inline math, prose, captions, and definitions.
- In Chinese Word documents, every explanatory paragraph beginning with `其中，` must have a zero direct first-line indent: set `w:ind/@w:firstLine="0"` and remove or zero `w:firstLineChars`. A body style's two-character first-line indent must never be inherited by this paragraph.
- Every Word paragraph containing an inline `Equation.DSMT4` object must set all object runs to Normal character position (`w:position/@w:val="0"`) and set Chinese text alignment to Center (`w:textAlignment/@w:val="center"`). Display-only `MTDisplayEquation` paragraphs are not subjected to the inline-paragraph text-alignment rule.
- Never use exact or fixed line spacing for a paragraph containing an inline `Equation.DSMT4` object. Use At Least line spacing with the paragraph's established body line height as the minimum so the OLE object can expand the line without clipping.
- Keep Word, PowerPoint, and MathType automation silent: leave application windows hidden, suppress modal alerts, do not steal focus, and do not automate visible UI with mouse, keyboard, `AppActivate`, or `SendKeys`. If a step cannot run silently, stop and report the limitation.

## Terminal selection

- Detect the active terminal and available executables before invoking the bridge.
- Use PowerShell 7 directly when the active host is `pwsh.exe` 7 or later.
- From Bash on Windows, including Git Bash, or from CMD, invoke the same Windows `pwsh.exe` bridge commands. From WSL Bash, invoke Windows `pwsh.exe`; Linux `pwsh` cannot automate Windows Office COM.
- Never run the bridge inside Windows PowerShell 5.1. If 5.1 is active, switch to Git Bash or CMD and invoke Windows `pwsh.exe` there.
- If no supported terminal or no `pwsh.exe` is available, stop and ask the user to install PowerShell 7 using <https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6>. Do not silently downgrade the workflow.
- Read [installation-matrix.md](references/installation-matrix.md) for cross-shell commands and agent-specific installation paths.

## Environment readiness

`probe_mathtype_word` answers everything needed to run the skill. Read `checks` and `warnings`;
do not assume a failed probe means the machine is unsupported.

- The MathType executable is discovered from `MATHTYPE_EXECUTABLE_PATH`, `PATH`, the uninstall
  registry entries, and the usual install roots. MathType may live outside `Program Files`.
- MathType Commands (`MathType Commands *.dotm`) is discovered in Word's STARTUP folders and in
  MathType's own `Office Support\64` or `Office Support\32` folder. When Word does not already load
  it, the bridge adds it as a Word add-in for the current process and removes it again afterwards, so
  the skill runs without installing anything into Word. `install_mathtype_word_template` makes the
  load permanent by copying the template into the user's Word STARTUP folder.
- Word COM needs write access to `%APPDATA%\Microsoft\Word`. In a restricted sandbox Word fails with
  "Word could not fire event" (`0x800A1772`). `checks.user_profile_writable` and `warnings` surface
  that before a long run fails, and `checks.word_document_creation` is reported separately because
  opening an existing DOCX can still succeed when creating a blank one fails.
- MathType maps `\mathbf`, `\boldsymbol`, `\bold` and `\textbf` to its **Vector** character
  style, and Texvc, the language behind Toggle TeX, has no separate bold-italic command. A Vector style
  of `Times New Roman,B` therefore renders every vector upright. `configure_mathtype_word_defaults`
  sets `Times New Roman,BI` after backing up the previous value, which is what notation macros such as
  `\mathbold` expect.

## Shared output coordination

- Resolve relative output paths against the user's current working directory. Do not default to the plugin repository or create a tool-named output subdirectory.
- Obey an explicit user-supplied input, output, or directory. Never reinterpret an explicit path, and never treat an input path as permission to overwrite it.
- Operate as a complete standalone skill. Do not assume, require, install, or prompt for `endnote-for-word` merely to complete MathType work.
- Prefix retained non-document artifacts with this skill name, for example `mathtype-for-word-manifest.json` or `mathtype-for-word-validation.json`.
- Give every newly generated DOCX or PPTX a descriptive suffix before its extension. Use `-mathtype` when this skill runs alone. Use `-mathtype-endnote`, for example `<source-stem>-mathtype-endnote.docx`, only when `endnote-for-word` is available and the task requires both skills.
- Verify that `endnote-for-word` is available before planning a joint workflow. If it is unavailable, complete only the MathType scope, keep the standalone suffix, report the unperformed EndNote scope, and never claim combined completion.
- When `endnote-for-word` is available and required, plan one shared final DOCX. Do not create separate MathType and EndNote document trees or keep one document per skill. Preserve the original source, let the first skill publish the shared working copy, let the second skill continue on that same copy with explicit overwrite authorization, then validate the final shared document with both skills.
- Keep bridge-owned temporary sibling files temporary. Do not present them as outputs, and remove retained intermediates after the shared document passes both validations.

## Choosing a Word mode

Word work has two modes. Pick one before touching the document.

| Mode | Use when | Entry point |
|---|---|---|
| Create from markers | The DOCX has no MathType equations yet, or equations are still `{{MATH:id}}` markers. | `render_mathtype_word_document` |
| Edit what exists | The DOCX already contains MathType equations and only their content must change. | `inventory_mathtype_word_equations` then `edit_mathtype_word_equations` |

Rendering rebuilds objects from TeX and therefore owns numbering and references. Editing rewrites
existing objects in place and must leave numbering, references, tab stops, and paragraph formatting
exactly as they were.

## Mandatory Word workflow

1. Preserve the input DOCX and choose an explicit output path.
2. Call `probe_mathtype_word`. Stop on a failed prerequisite; report the exact failed check. Read its `warnings` too: they flag a non-writable user profile, a template that is only loaded at runtime, and a Vector style that would render vectors upright.
3. Scan the complete document twice. First inventory every equation candidate and later prose reference. Then classify each candidate as `inline`, `display`, or `display_numbered`, and each prose occurrence as `reference`, using [academic-equation-style.md](references/academic-equation-style.md). Decide every identifier, unique marker, TeX, layout, numbered state, reference marker, and target.
4. Create a schema v1 JSON manifest. Read [workflow-and-schema.md](references/workflow-and-schema.md) for the exact contract and examples.
5. For academic prose and mathematical typography, apply [academic-equation-style.md](references/academic-equation-style.md). A display equation must be introduced in the preceding prose and followed by “其中，” plus definitions of every symbol, index, superscript, subscript, unit, and physical meaning that is not already unambiguous. A paragraph beginning with “其中，” immediately after a numbered display is a continuation of the lead-in and equation: force `w:ind/@w:firstLine="0"` and clear `w:firstLineChars`, then keep all three together on the same page. For every paragraph containing inline equations, set object runs to Normal character position, set `w:textAlignment` to `center`, and change exact/fixed spacing to At Least while retaining the established body line height. Audit the whole document so variables, vectors, matrices, functions, constants, indices, and units retain one consistent style.
6. Call `render_mathtype_word_document`. Do not edit the same DOCX concurrently in Word.
7. Call `validate_mathtype_word_document` with the same manifest. A render is incomplete until validation returns `ok: true`. Also run `python scripts/normalize-word-inline-math-layout.py "$outputDocx" --check`; require `ok: true` before delivery.
8. For complex fractions, matrices, aligned systems, accents, or nested scripts, open/render the result and visually inspect it. Structural checks cannot prove typographic correctness.
9. Report the output path, counts, native number/reference mechanism, validation result, and any limitations.

## Mandatory PowerPoint workflow

1. Preserve the input PPTX and choose a distinct output path.
2. Call `probe_mathtype_powerpoint`. Stop and report the exact failed check when `powerpoint_ready` is false.
3. Put each `{{MATH:id}}` marker alone in a top-level text box. PowerPoint does not support inline external objects.
4. Create a presentation schema v1 manifest, apply [academic-equation-style.md](references/academic-equation-style.md), and read [powerpoint-workflow.md](references/powerpoint-workflow.md).
5. Call `render_mathtype_powerpoint_presentation`, then `validate_mathtype_powerpoint_presentation` with the same manifest.
6. Render every output slide and inspect the full-size image. Structural OLE checks do not prove legibility or absence of overlap.
7. Report that the result contains editable floating MathType OLE objects and that Word-style dynamic numbering/references are unavailable in PPTX.

## MCP tools

Use the bundled local MCP server when available:

- `probe_mathtype_word`: read-only prerequisite check. Discovers MathType and its Word template anywhere on the machine and reports whether the template loads at runtime, so a STARTUP install is optional.
- `configure_mathtype_word_defaults`: persist the packaged default profile and MathType warning preferences, and set the MathType Vector style to bold italic after backing up the previous value. Invoke only when the user asks to configure defaults or during an explicitly requested installation.
- `install_mathtype_word_template`: copy MathType Commands into Word's STARTUP folder so Word loads it at every start. Optional; the bridge otherwise loads it per run.
- `inventory_mathtype_word_equations`: read-only list of every MathType equation in a DOCX with its 1-based ordinal, paragraph, numbered state, cached TeX, and (on request) exact TeX and inline/display layout.
- `edit_mathtype_word_equations`: rewrite or delete the equations named in an edit manifest, verifying `old_tex` and preserving each equation's layout.
- `render_mathtype_word_preview`: render the DOCX to PDF and optional page PNGs for visual inspection.
- `render_mathtype_word_document`: marker-driven conversion and reference placement.
- `validate_mathtype_word_document`: read-only structural validation.
- `update_mathtype_word_fields`: update number/reference fields after moving, adding, or deleting equations.
- `probe_mathtype_powerpoint`: verify the desktop PowerPoint and MathType 7 integration.
- `render_mathtype_powerpoint_presentation`: replace marker-only text boxes with centered, editable MathType OLE objects.
- `validate_mathtype_powerpoint_presentation`: verify expected named OLE objects, centering, and resolved markers.

If MCP is unavailable, invoke the same bridge directly from the skill directory:

```powershell
# Resolve these variables from the user's attached or repository files; never reuse a sample path literally.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action probe
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action render -InputPath $inputDocx -OutputPath $outputDocx -ManifestPath $manifestJson
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action inventory -InputPath $inputDocx -ReadTexFromMathType
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action edit -InputPath $inputDocx -OutputPath $outputDocx -ManifestPath $editManifestJson
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action preview -InputPath $outputDocx -OutputPath $outputPdf -ImageDirectory $imageDirectory
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\mathtype-word.ps1 -Action validate -InputPath $outputDocx -ManifestPath $manifestJson
```

When the skill is installed as a standalone skill directory, call the wrapper next to this file instead:
`scripts\invoke-bridge.ps1 -Action <action> ...`. It resolves the bridge inside the skill, so no
plugin root or environment variable is needed.

## Manifest decisions

- Use stable ASCII IDs such as `energy_balance` or `eq_motion_x`.
- Use collision-resistant visible markers such as `{{MATH:eq_motion_x}}` and `{{EQREF:eq_motion_x}}`.
- Each marker must occur exactly once. Never reuse a marker.
- Map `inline` to `layout: "inline", numbered: false`.
- Map `display` to `layout: "display", numbered: false`.
- Map `display_numbered` to `layout: "display", numbered: true`.
- Map each `reference` occurrence to its own unique `references[].marker`; multiple markers may target the same numbered equation.
- Promote a display to `display_numbered` when the full-document scan finds a specific later reference to it.
- Only display equations may set `numbered: true`. Never create a reference to an unnumbered equation.
- Every reference target must exist and be numbered.
- Keep TeX on one line. Prefer MathType-compatible TeX primitives. Do not depend on document preambles, user macros, packages, `\label`, or `\ref`.
- Escape backslashes correctly in JSON. For example, TeX `\frac{a}{b}` is JSON string `"\\frac{a}{b}"`.

## Numbering and reference policy

The bundled default mirrors the requested Format Equation Numbers configuration:

- Simple Format.
- Chapter Number off.
- Section Number off.
- Equation Number on, Arabic `1,2,3,...`.
- Enclosure on, parentheses.
- Separator off.
- New equation numbers and Whole document.
- Update equation numbers automatically on.
- Warn when inserting first equation number on.
- Warn when inserting equation references off.
- Use format as default for new documents on within this agent workflow.

The bridge temporarily suppresses blocking dialogs during headless automation, restores the user's warning preferences afterward, and normalizes MathType's native field structure to `(n)`. Read [numbering-and-references.md](references/numbering-and-references.md) before diagnosing or manually repairing a field.

## Editing equations that already exist

Use this workflow whenever a DOCX already contains MathType equations and their content must change.
Do not render new equations over them, and do not toggle the whole document to save time.

1. Preserve the input DOCX and choose an explicit output path. Editing the input in place needs an
   explicit user decision and `-Overwrite`.
2. Probe first and read `warnings`.
3. Run `inventory_mathtype_word_equations` with `read_tex_from_mathtype` (CLI:
   `-Action inventory -ReadTexFromMathType`). It works on a throwaway copy, so the source is never
   modified, and reports per equation: the 1-based `ordinal` in document order, the `paragraph`,
   whether the paragraph carries a MathType number field, the exact `tex`, and `layout`
   (`inline` or `display`).
   - An equation whose OLE object stores no TeX comes back as `tex: null`,
     `tex_source: "missing"`. The toggle pass recovers it and marks it `mathtype-toggle`.
   - A one-letter symbol that reports `layout: "display"` while sitting inside running text is a
     display-flagged inline object. Leave it untouched unless its content really changes; the edit
     action reuses its delimiter.
4. Create a schema v1 edit manifest (see [workflow-and-schema.md](references/workflow-and-schema.md)):
   one entry per equation, addressed by `index`, carrying the exact `old_tex` just read plus either
   `new_tex` or `delete: true`. Add `prose` find/replace entries when the sentences around a formula
   must follow the new symbols.
5. Run `edit_mathtype_word_equations`. Each listed equation is toggled out, verified against
   `old_tex`, rewritten, and toggled back with the same delimiter. Unlisted equations are never
   touched. A failed verification aborts the run before anything is published, so the source stays
   intact and no output file appears.
6. Validate the output with `validate_mathtype_word_document`. The edit action already runs the same
   structural validation before publishing: MathType object count, MathType number fields, sequential
   `(1), (2), ...`, no Word OMath, no leftover markers.
7. Inspect the result with `render_mathtype_word_preview`, comparing against a pre-edit render when
   the change is more than a symbol swap.
8. Report the output path, the per-equation counts (`rewritten`, `deleted`, `untouched`), the prose
   results, and any verification problem.

- `index` is the document-order ordinal from the inventory, not a shape id. Re-run the inventory if
  the document changed in between, and never reuse a stale `old_tex` with verification enabled.
- Prose entries are exact, case-sensitive literal replacements with an optional `count` assertion.
- Deleting an equation leaves its paragraph and any number field in place. Re-run
  `update_mathtype_word_fields` (`-Action update`) when a numbered equation was removed.
- Use markers or the edit manifest for deterministic changes. Never use broad find/replace on TeX
  fragments or equation text.
- Do not unlink, flatten, or convert the MathType fields or OLE objects.
- Do not overwrite a locked file. Ask the user to close it or select another output path.

## Advanced ribbon operations

The screenshot-level workflow also includes symbol palettes, equation browsing, preferences, whole-document formatting, conversion, export, MathPage, and help. Read [advanced-ribbon-operations.md](references/advanced-ribbon-operations.md) before using any of them. The MCP surface intentionally automates only the deterministic subset; interactive or whole-document commands require explicit scope, a preserved source, and visual review.

## PowerPoint boundary

- Use desktop **MathType for Windows**, not the Microsoft 365 task-pane add-in.
- Treat every PowerPoint equation as a floating object; there is no inline OLE placement.
- Keep Word, PowerPoint, and MathType hidden. The bridge uses hidden Word conversion and briefly transfers the OLE object through the Windows clipboard without mouse, keyboard, focus, `AppActivate`, or `SendKeys`.
- Do not launch Word for a PPTX job; the direct PowerPoint path must leave the set of `WINWORD` process IDs unchanged.
- Do not promise MathType-native PowerPoint equation numbering or cross-references. Keep those features in Word, or ask the user to accept a clearly labeled static slide annotation.
- Do not rasterize equations. A valid PPTX output retains `Equation.DSMT4` and opens in desktop MathType when edited.

## Validation gates

Treat any of these as a failed deliverable:

- Missing `Equation.DSMT4` objects.
- Any Word built-in `OMath` object when the requested equations must all be MathType.
- A numbered equation without `MACROBUTTON MTPlaceRef`, hidden/current `SEQ MTEqn`, or parentheses.
- A current `MTSec` or `MTChap` component in a number.
- A reference without `GOTOBUTTON`, nested `REF`, or a live `ZEqnNum...` bookmark.
- Remaining `{{MATH:...}}`, `{{EQREF:...}}`, or `equation reference goes here` text.
- Non-sequential equation values starting from 1.
- `Error! Reference source not found.`
- An edit that reported a verification problem, or an output whose MathType object count no longer
  matches the inventory minus the deletions.
- An edited document in which an inline symbol now occupies its own line: the edit used the wrong
  delimiter instead of preserving the equation's original layout.
- A Chinese `其中，` explanation inherits a two-character first-line indent, or has nonzero
  `w:firstLine` or `w:firstLineChars`.
- An inline-math paragraph lacks centered Chinese text alignment, or any inline equation run has a
  nonzero or missing Normal character-position setting.
- A paragraph containing an inline equation still uses exact/fixed line spacing, or its At Least
  minimum is below the intended body line height.
- A `validate` warning that MathType could not convert a command: the TeX read back from the object
  still contains it, so that equation renders as literal red text or drops a symbol. Object counts
  and number fields stay correct in this case, so this check is the only automated signal — inspect
  `stored_tex` in the report and the rendered pages, and never ship the document as-is.

## Failure handling

- If probe fails, use [troubleshooting.md](references/troubleshooting.md); do not silently fall back.
- If Toggle TeX does not create exactly one `Equation.DSMT4` object, preserve the input, stop, and report the equation ID and TeX.
- If MathType fails to create or resolve `MTReference`, preserve the failed output only for diagnosis and do not present it as complete.
- If the MCP watchdog reports a timeout, record its isolated Word PID and cleanup result, verify the warning preferences were restored, then retry only the smallest fixture once before escalating.
- If the requested TeX is outside MathType's supported conversion subset, simplify it, split it into supported expressions, or ask for a MathType-authored source equation. Toggle TeX speaks Texvc and does **not** reject unknown commands: it creates the object anyway and either echoes the command as literal red text, renders a black diamond, or silently drops the construct. `\mathbfit`, `\bm`, `\mathbold`, `\widebar`, `\widetilde`, `\operatorname`, `\lvert`, `\rvert`, `\partial`, `\middle`, `\text` and `cases` come back as literal red text; `\Longrightarrow`, `\iff`, `\longrightarrow`, `\lceil` and `\rceil` render as a black diamond; `\mathop{...}` drops its subscript, `\limits` vanishes, and a norm written `\|x\|_2` loses its double bars. `\mathbf`, `\boldsymbol`, `\mathrm`, `\mathcal`, `\mathbb`, `\widehat`, `\bar`, `\tilde`, `\odot`, `\vartheta` and `\lVert...\rVert` are accepted. `render` refuses such a manifest upfront and `validate` re-checks the TeX read back from each object.
- If an edit reports an `old_tex` mismatch, stop and re-inventory. The document changed, the index is
  stale, or the cached TeX was never the equation's real content.
- If a rewritten equation reflows its paragraph, restore the source from the preserved copy and report
  the equation ordinal, its original delimiter, and the TeX that was applied.
- Word COM is single-user desktop automation. Do not run two render jobs concurrently in the same Windows profile.

## Reference routing

- Manifest design and end-to-end examples: [workflow-and-schema.md](references/workflow-and-schema.md)
- Native field anatomy and Format Equation Numbers behavior: [numbering-and-references.md](references/numbering-and-references.md)
- Academic placement, consistent symbol typography, definitions, units, and TeX/MathML guidance: [academic-equation-style.md](references/academic-equation-style.md)
- Chinese inline-math paragraph and character-position audit: `scripts/normalize-word-inline-math-layout.py --check`
- Terminal selection plus Claude Code, Claude Desktop, Codex, and ChatGPT Desktop setup: [installation-matrix.md](references/installation-matrix.md)
- Errors, recovery, and fallback boundary: [troubleshooting.md](references/troubleshooting.md)
- MathType ribbon capability matrix and guarded interactive commands: [advanced-ribbon-operations.md](references/advanced-ribbon-operations.md)
- PowerPoint manifest, floating-object workflow, and validation: [powerpoint-workflow.md](references/powerpoint-workflow.md)
