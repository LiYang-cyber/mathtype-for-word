# Troubleshooting

| Symptom | Likely cause | Action |
|---|---|---|
| `word_com` is false | Word desktop missing, broken registration, or blocked automation | Repair/install Word; retry the probe in the interactive user session. |
| MathType executable or template missing | MathType or Word add-in is not installed | Repair MathType and enable its Word support. |
| `equation_dsmt4_registered` is false | MathType OLE registration is broken | Run MathType repair as the same Windows user. |
| Toggle TeX creates no object | Unsupported delimiter/TeX, add-in not loaded, or interactive dialog | Use one-line portable TeX; confirm `$...$` for inline and `\[...\]` for display; inspect Word add-ins. |
| Validation names an equation that `kept untranslated TeX as red text` or `drew … unmapped symbol(s) as a diamond` | Toggle TeX accepted TeX it cannot translate; the object exists and every structural check passes | Replace the command with a form from [TeX that Toggle TeX does not translate](#tex-that-toggle-tex-does-not-translate) and render again from the preserved input. |
| Toggle TeX creates no object for TeX that starts with a vertical bar | MathType 7.11.1.462 converts no TeX whose first character is a vertical bar | See [TeX that Toggle TeX does not translate](#tex-that-toggle-tex-does-not-translate). |
| Toggle TeX creates Word OMath | Word conversion was used instead of MathType | Remove the OMath result and rerun through `MTCommand_TeXToggle`. |
| Number is `(1.1)` | Existing MathType format includes section number | Render through the bridge; it removes current `MTSec`/`MTChap` components while preserving `MTPlaceRef`. |
| `MTReference` remains | The target `MTPlaceRef` action did not complete | Do not type over it; retry against a native numbered equation. |
| `Error! Reference source not found.` | A `ZEqnNum...` bookmark was deleted or the field was copied incorrectly | Recreate the reference through MathType; do not hand-edit the bookmark. |
| Output is locked | Word or another process has the DOCX open | Close it or choose a new output path. |
| Word remains in Task Manager after a failure | A COM call or MathType dialog blocked | Identify only the Word process started by the failed job before terminating it; never kill all Word sessions. |
| MCP returns a bridge timeout | A Word/MathType COM macro blocked longer than the watchdog | Check `bridge_office_pids`, `isolated_office_processes_terminated`, and `new_office_pids_left_running` (for a Word render job, in its `timeouts` list); confirm preferences were restored (`%APPDATA%\MathTypeForWordAgent\warning-preferences.json` absent means the user's values are back; present means the next render, render-pptx, or configure-defaults call restores them when it takes the bridge lock), then retry one small fixture. The default watchdog is 240 seconds and may be changed with `MATHTYPE_OFFICE_TIMEOUT_SECONDS`; batched chunks scale it with their size. |
| Equation numbering or references hang in hidden Word | A MathType dialog other than the three warning preferences, or a hidden instance left by an earlier failure. The bridge writes `NoEqnNumWarningDlg`, `NoInsertEqnRefDlg`, and `NoSlowUpdateEqnDlg` as `REG_SZ` `"1"` before every render (MathType ignores `REG_DWORD`) and restores the user's values afterwards, or on the next run after a kill | Check the log for the last equation ID and look for the blocking dialog; confirm the three values under `HKCU\Software\Design Science\DSMT7\WordCommands` match the pre-run state. Run `configure_mathtype_word_defaults` only when the user asks to configure defaults. |
| RPC error `0x800706BA`/`0x800706BE` in a large document | One Word session exhausted after hundreds of conversions | The job already halves a chunk whose Word session fails, so this error means a single-operation chunk failed. Call render again with the same paths: it resumes from the last checkpoint and retries that chunk in a fresh Word session. A job keeps the `batch_size` it started with; another `batch_size` applies only with `resume: false` (or after `cleanup: true`), which starts over. |
| `Another MathType render job is running (pid …, output …)` | A foreground or background Word render job for another output is still active | Poll `get_mathtype_render_status` with the `output` shown in the message; start this job when it has finished. The MCP PowerPoint render and defaults tools are refused the same way while a Word render job runs; a direct bridge call is not checked and can make the next chunk fail on the bridge lock. |
| `Another MathType Office job is still running after 30 s` | Another editing bridge call (a Word render chunk, `render-pptx`, `configure-defaults`, or a direct CLI call) holds the bridge lock | Wait for that call to finish, then retry. A PowerPoint render, a defaults configuration, or a CLI call has no job status to poll. Only a live `pwsh.exe` running `mathtype-word.ps1` holds the lock (another MCP client's server, or an orphaned direct CLI call); a killed bridge or a leftover hidden Office process never does. If no call you know of is running, find the holder with `Get-CimInstance Win32_Process -Filter "Name='pwsh.exe'" \| Where-Object CommandLine -like '*mathtype-word.ps1*' \| Select-Object ProcessId, CommandLine`, confirm it is not a job you still need, and end it. |
| `The presentation is already open in PowerPoint` | The user has the input or output PPTX open (for a OneDrive or SharePoint file, any open cloud presentation with the same file name) | Ask the user to save and close it, then retry. The bridge never edits or closes a presentation the user has open. |
| Validation warns that an equation is clipped | A paragraph with exact line spacing, or a table row with an exact height, is lower than the equation | Re-render with the default `exact_line_spacing: "at_least"`, or set that paragraph or row to an at-least value. When validation says a table row could not be inspected, set that row to an at-least height in Word. |
| `The Windows clipboard is held open by …` | Another program keeps the clipboard open (some chat and clipboard-manager apps do), so Word cannot hand the equation to PowerPoint | Ask the user to close or restart the program named after `held open by`. For `an unidentified program`, the last program that wrote to the clipboard, when named, is the likeliest holder; otherwise ask which clipboard-manager, chat, or remote-desktop app is running. Never ask the user to close Word for this error. Retry afterwards; the bridge never ends any program. |
| `… cross-session anchor bookmark(s) of an unfinished batched render remain` | The document is a checkpoint of a batched render, not its published output | Validate the published output, or finish the job by calling render again with the same paths. |
| Validation reports OMath | Existing or newly inserted built-in Word math is present | Convert the requested equations to MathType; inspect unrelated legacy OMath separately before deletion. |
| `pwsh.exe` was not found, so the bridge or the MCP server cannot start | This fork requires PowerShell 7 on PATH and never installs it | Install PowerShell 7 with the MSI installer (`winget install --id Microsoft.PowerShell --source winget --installer-type wix`), or set `MATHTYPE_WORD_POWERSHELL` to its absolute path, then restart the agent. |
| MCP server emits parse errors | A launcher/log wrote to stdout | Protocol output must be JSON only; keep diagnostics on stderr. |

## TeX that Toggle TeX does not translate

Toggle TeX does not reject TeX it cannot translate. It keeps an unknown command as red literal text, or draws a symbol it cannot map as a diamond, and the object passes every structural check. Validation reads the saved objects and rejects such an equation, naming its id and TeX, and for red text the untranslated command (for a diamond, only how many). Every row and every replacement below was measured on MathType 7.11.1.462 with Word 16.0 in both display and inline equations (`tests/TEST-REPORT.md`); other versions may differ, so trust the validation result over this table.

| Command | Result | Use instead |
|---|---|---|
| `\bm{x}`, `\mathbfit{x}`, `\pmb{x}` | red text | `\boldsymbol{x}`; `\mathbf{x}` for upright bold vectors |
| `\widetilde{x}` | red text | `\tilde{x}`; for a bold symbol `\boldsymbol{\tilde{M}}` (see the PDF note below) |
| `\widebar{x}` | red text | `\bar{x}` or `\overline{x}` |
| `\lvert`, `\rvert`, `\middle` | red text | see the list below |
| `\mathtt{x}` | red text | `\texttt{x}` |
| `\coloneqq` | red text | `:=` |
| `\phantom{x}`, `\hphantom{x}` | red text | spacing commands such as `\quad` or `\,` |
| `\boxed{x}`, `\fbox{x}` | red text | none; frame the equation in Word if required |
| `aligned`, `gathered`, `split` environments | red text | one display equation per line |
| `\Longrightarrow`, `\longrightarrow`, `\iff` | diamond | `\Rightarrow`, `\rightarrow` or `\to`, `\Leftrightarrow` |
| `\lceil x\rceil`, `\lfloor x\rfloor` | diamond | `\left\lceil x\right\rceil`, `\left\lfloor x\right\rfloor` |
| `\longmapsto`, `\hookrightarrow` | diamond | `\mapsto`, `\rightarrow` |
| `\mathbb{1}` | diamond | `\mathbf{1}` |

Vertical bars, kept out of the table because Markdown reads them as cell borders:

- `\lvert x\rvert` (red text): use `\left|x\right|`.
- `\left(a\middle|b\right)` (red text): use `\left(a\mid b\right)`.
- TeX whose first character is a vertical bar, such as `|x|`, creates no object at all, and render stops with `Toggle TeX did not add exactly one MathType object`. Start it with `\left|x\right|`; `y=|x|` converts.

These produce neither red text nor a diamond on 7.11.1.462 and need no substitute: `\operatorname`, `\operatorname*`, `\text`, `\textbf`, `\partial`, `\mathop`, `\limits`, `\nolimits`, `\left\|x\right\|`, `\left\lVert x\right\rVert`, `\langle`, `\mapsto`, `\triangleq`, `\overset`, `\underbrace`, `\xrightarrow`, `\boldsymbol`, `\mathscr`, `\mathfrak`, and the two-row environments `cases`, `bmatrix`, `pmatrix`, `vmatrix`, `Vmatrix`, `array` and `\substack`. Their layout still needs a visual check; for example, `\|x\|` and `\lVert x\rVert` draw two separate bars, where `\left\|x\right\|` and `\left\lVert x\right\rVert` draw a fused norm.

### Accents on bold symbols in Word's PDF export

An accent written around a bold command (`\tilde{\boldsymbol{M}}`, `\tilde{\mathbf{M}}`, `\hat{\boldsymbol{x}}`, `\hat{\mathbf{x}}`) becomes a wide-accent template. Word shows and saves it, but Word 16.0's PDF export (`ExportAsFixedFormat`) leaves the accent out, so the PDF shows a bare **M** or **x**. Validation cannot see this. Put the accent inside the bold command instead: `\boldsymbol{\tilde{M}}`, `\mathbf{\tilde{M}}`, `\boldsymbol{\hat{x}}` and `\mathbf{\hat{x}}` keep their accent in the PDF. `\bar` and `\vec` around a bold command kept theirs.

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
| `validate` reports untranslated TeX or a diamond | The saved object still holds a TeX command MathType could not translate, or an unmapped symbol it drew as a diamond | Treat the document as not delivered; replace the command from [TeX that Toggle TeX does not translate](#tex-that-toggle-tex-does-not-translate) and render again. The render result names the equation and the command. |
| `inventory -ReadTexFromMathType` reports `layout_error` and `tex_source: "ole-cache"` | The recovered equation count disagreed with the package | Re-run; if it persists, report the counts. The OLE cache is returned instead and may differ from what MathType renders. |
| Inventory reports `tex: null`, `tex_source: "missing"` for some equations | Those OLE objects store no TeX cache | Re-run the inventory with `-ReadTexFromMathType` to recover their TeX by toggling. |
| After an edit, inline symbols sit on their own lines | The run toggled equations that did not need changing, so display-flagged inline symbols came back as display equations | Restore the source, list only the equations that must change, and let the edit action reuse each equation's original delimiter. |
| Every equation's TeX changed after a no-op round trip | The whole document was toggled instead of the listed equations | Prefer the edit action. A round trip re-creates every object it touches, including ones that needed no change. |
| Preview output looks like Word field code soup | LibreOffice prints field codes and hidden MathType section markers | Judge line breaks, spacing, and glyph style only; verify field behaviour in Word. |
| Preview produces no PDF | `soffice.exe` is a launcher and returns before the conversion finishes, or LibreOffice is missing | The action already waits for the output file; if it still fails, install LibreOffice or set `MATHTYPE_SOFFICE_PATH`, and check the reported converter path. |
