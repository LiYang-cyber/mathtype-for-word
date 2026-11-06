# Troubleshooting

| Symptom | Likely cause | Action |
|---|---|---|
| `word_com` is false | Word desktop missing, broken registration, or blocked automation | Repair/install Word; retry the probe in the interactive user session. |
| MathType executable or template missing | MathType or Word add-in is not installed | Repair MathType and enable its Word support. |
| `equation_dsmt4_registered` is false | MathType OLE registration is broken | Run MathType repair as the same Windows user. |
| Toggle TeX creates no object | Unsupported delimiter/TeX, add-in not loaded, or interactive dialog | Use one-line portable TeX; confirm `$...$` for inline and `\[...\]` for display; inspect Word add-ins. |
| Toggle TeX creates Word OMath | Word conversion was used instead of MathType | Remove the OMath result and rerun through `MTCommand_TeXToggle`. |
| Number is `(1.1)` | Existing MathType format includes section number | Render through the bridge; it removes current `MTSec`/`MTChap` components while preserving `MTPlaceRef`. |
| `MTReference` remains | The target `MTPlaceRef` action did not complete | Do not type over it; retry against a native numbered equation. |
| `Error! Reference source not found.` | A `ZEqnNum...` bookmark was deleted or the field was copied incorrectly | Recreate the reference through MathType; do not hand-edit the bookmark. |
| Output is locked | Word or another process has the DOCX open | Close it or choose a new output path. |
| Word remains in Task Manager after a failure | A COM call or MathType dialog blocked | Identify only the Word process started by the failed job before terminating it; never kill all Word sessions. |
| MCP returns a bridge timeout | A Word/MathType COM macro blocked longer than the watchdog | Check `isolated_word_pid` and `isolated_word_process_terminated`; confirm preferences were restored, then retry one small fixture. The default watchdog is 240 seconds and may be changed with `MATHTYPE_WORD_TIMEOUT_SECONDS`. |
| Validation reports OMath | Existing or newly inserted built-in Word math is present | Convert the requested equations to MathType; inspect unrelated legacy OMath separately before deletion. |
| MCP server emits parse errors | A launcher/log wrote to stdout | Protocol output must be JSON only; keep diagnostics on stderr. |

## Recovery rules

- Keep the input untouched.
- Record the equation ID, marker, TeX, and exact bridge error.
- Check for a blocking dialog in Word/MathType if a call times out.
- Clean only the isolated Word process created by the failed job.
- Confirm temporary warning registry values were restored.
- Rerun the smallest fixture before rerunning a large document.

## Fallback boundary

Word captions, numbered lists, manually typed numbers, and Word OMath are a fallback only. Before using one:

1. Show the failed prerequisite or conversion result.
2. Explain that the output will not be MathType-native.
3. Obtain explicit user acceptance.
4. Label the deliverable as a fallback and do not claim native MathType references.

## Environment and editing failures

| Symptom | Cause | Action |
|---|---|---|
| Probe fails with `Word 未能引发事件` / `Word could not fire event` / `0x800A1772` | Word COM cannot write `%APPDATA%\Microsoft\Word` (restricted sandbox, locked profile, or a modal Word dialog) | Check `checks.user_profile_writable` and `checks.word_document_creation_error`; widen the file sandbox or run outside it. Do not switch to Word OMath or images. |
| Probe reports `math_type_template_usable: false` and `math_type_word_template_path: null` | MathType 7 is installed but its Word commands template cannot be found | Confirm the install, or pass `-TemplatePath` / set `MATHTYPE_WORD_TEMPLATE`. |
| Probe warns that the template is loaded at runtime | Word does not load MathType Commands at startup | Harmless: the bridge loads it per run. Run `install-template` to make it permanent. |
| Probe warns about the Vector style | MathType's Vector style is bold but not italic, so `\mathbf` renders upright vectors | Run `configure-defaults` (it backs up the previous value) or accept upright vectors explicitly. |
| TeX comes back as literal red text | Toggle TeX speaks Texvc; the command is outside its subset | Use `\mathbf`, `\boldsymbol`, `\mathrm`, `\mathcal`, `\mathbb`, `\widehat`, `\bar`, `\tilde`, `\odot`, `\vartheta`. Avoid `\mathbfit`, `\bm`, `\mathbold`, `\widebar`, `\widetilde`, `\operatorname`, `\lvert`, `\rvert`, `\partial`, `\middle`, `\text`, `cases`. |
| Equation renders as a black diamond (`◆`) | The command maps to no glyph in Texvc | `\Longrightarrow`, `\iff`, `\longrightarrow`, `\lceil`, `\rceil` are the common ones; use `\Rightarrow`, `\Leftrightarrow`, `\rightarrow`, or `\mathrm{ceil}`. |
| A subscript vanished, or a norm lost its bars | The construct was dropped silently instead of echoed, so object counts and number fields still look correct | `\mathop{...}` drops its subscript, `\limits`/`\nolimits` vanish, and a norm written `\|x\|_2` or `\left\|x\right\|_2` loses its double bars. Use `\mathrm{...}_{...}` and `\lVert x\rVert_2`. |
| `render` refuses the manifest with "uses commands MathType renders as literal text" | The pre-flight check found a command from the list above in the manifest | Replace it and render again. Do not bypass the check: the failure is invisible in the output. |
| `validate` warns that MathType could not convert a command | The TeX read back from the object still contains that command, so it renders as literal text or drops a symbol | Treat the document as not delivered; fix the TeX and render again. `stored_tex` in the report shows what MathType actually holds. |
| `inventory -ReadTexFromMathType` reports `layout_error` and `tex_source: "ole-cache"` | The recovered equation count disagreed with the package | Re-run; if it persists, report the counts. The OLE cache is returned instead and may differ from what MathType renders. |
| Inventory reports `tex: null`, `tex_source: "missing"` for some equations | Those OLE objects store no TeX cache | Re-run the inventory with `-ReadTexFromMathType` to recover their TeX by toggling. |
| After an edit, inline symbols sit on their own lines | The run toggled equations that did not need changing, so display-flagged inline symbols came back as display equations | Restore the source, list only the equations that must change, and let the edit action reuse each equation's original delimiter. |
| Every equation's TeX changed after a no-op round trip | The whole document was toggled instead of the listed equations | Prefer the edit action. A round trip re-creates every object it touches, including ones that needed no change. |
| Preview output looks like Word field code soup | LibreOffice prints field codes and hidden MathType section markers | Judge line breaks, spacing, and glyph style only; verify field behaviour in Word. |
| Preview produces no PDF | `soffice.exe` is a launcher and returns before the conversion finishes, or LibreOffice is missing | The action already waits for the output file; if it still fails, install LibreOffice or set `MATHTYPE_SOFFICE_PATH`, and check the reported converter path. |
