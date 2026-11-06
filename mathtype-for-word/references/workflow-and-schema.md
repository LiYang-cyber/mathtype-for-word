# Workflow and manifest schema

## Contents

1. End-to-end flow
2. Schema v1
3. Complete example
4. Marker preparation
5. Output contract

## End-to-end flow

The bridge performs a deterministic transaction:

1. Open the input DOCX in an isolated hidden Word instance.
2. Require every marker to occur exactly once.
3. Replace equation markers with TeX delimited for MathType Toggle TeX.
4. Require exactly one new `Equation.DSMT4` object per equation.
5. For numbered displays, call MathType's native number command and normalize its native fields to `(n)`.
6. Replace reference markers through MathType's `MTReference` and `MTPlaceRef` workflow.
7. Update Word fields.
8. Save to a temporary sibling file and atomically move it to the output path.
9. Reopen read-only and validate.

The bridge refuses an existing output unless `overwrite` is explicitly true.

## Schema v1

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "stable_ascii_identifier",
      "marker": "{{MATH:stable_ascii_identifier}}",
      "tex": "one-line MathType-compatible TeX",
      "layout": "inline or display",
      "numbered": false
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:stable_ascii_identifier}}",
      "target": "stable_ascii_identifier"
    }
  ]
}
```

Rules:

- `schema_version` must equal `1`.
- `equations` is required. `references` may be empty.
- `id` values are unique and case-sensitive.
- All equation and reference markers are unique across the manifest.
- A marker occurs exactly once in the DOCX.
- `tex` is non-empty and contains no newline.
- `layout` is `inline` or `display`.
- `numbered: true` is legal only for `display`.
- A reference target names an existing numbered equation.

Semantic classes map to schema v1 as follows:

| Class | `layout` | `numbered` | Additional entry |
|---|---|---:|---|
| `inline` | `inline` | `false` | None |
| `display` | `display` | `false` | None |
| `display_numbered` | `display` | `true` | None |
| `reference` | Not applicable | Not applicable | One unique `references[]` item targeting a `display_numbered` equation |

Read the whole document before finalizing this table. A later semantic reference promotes its target from `display` to `display_numbered`. Multiple reference occurrences may share a target but must use different source markers.

## Complete example

Source text:

```text
The relativistic energy relation is expressed as follows:

{{MATH:mass_energy}}

where E is the energy in joules, m is the mass in kilograms, and c is the speed of light in metres per second.

As shown in Eq. {{EQREF:mass_energy_result}}, energy is proportional to mass.
```

Manifest:

```json
{
  "schema_version": 1,
  "equations": [
    {
      "id": "mass_energy",
      "marker": "{{MATH:mass_energy}}",
      "tex": "E = mc^2",
      "layout": "display",
      "numbered": true
    }
  ],
  "references": [
    {
      "marker": "{{EQREF:mass_energy_result}}",
      "target": "mass_energy"
    }
  ]
}
```

Expected output structure:

- One `Equation.DSMT4` OLE object.
- One native MathType number displayed as `(1)`.
- One native reference displayed as `(1)`.
- The target number is bookmarked as `ZEqnNum` followed by digits.
- The reference is a `GOTOBUTTON` field with a nested `REF` field.

## Marker preparation

Prefer markers that are visually obvious, unique, and absent from normal prose. Put a display marker alone in its paragraph. An inline marker remains inside the sentence.

Good:

```text
The gain is {{MATH:inline_gain}} at resonance.

{{MATH:transfer_function}}

其中，H(s) 為轉移函數，s 為複數頻率，單位為 rad/s。
```

Avoid:

- Reusing `{{MATH:eq1}}` twice.
- Embedding a display marker inside a sentence.
- Using the TeX itself as a find target.
- Using tracked-change fragments as markers without accepting or inspecting the changes.

## Output contract

A successful render response reports paths and counts. It does not replace validation. A valid output must return `ok: true` from validation with:

- MathType object count at least equal to manifest equation count.
- Native number count equal to numbered equation count.
- Native reference count equal to manifest reference count.
- Sequential equation values beginning at 1.
- No unresolved markers or placeholders.
- No Word OMath objects under strict MathType output.

## Edit manifest (schema v1)

An edit manifest rewrites equations that already exist in a DOCX. It is a different document type from
a render manifest: it never creates equations, and it addresses existing ones by the 1-based ordinal
that `inventory_mathtype_word_equations` reports.

```json
{
  "schema_version": 1,
  "mode": "edit",
  "equations": [
    {
      "index": 36,
      "old_tex": "\\tau_i=\\frac{\\kappa_i+\\operatorname{stopgrad}(\\rho_i)}{2}",
      "new_tex": "\\tau_i=\\kappa_i"
    },
    {
      "index": 39,
      "old_tex": "\\operatorname{stopgrad}(\\cdot)",
      "delete": true
    }
  ],
  "prose": [
    {
      "find": "将邻域兼容度与经梯度截断的原型集中度组合成消息信任值。",
      "replace": "以邻域兼容度作为消息信任值。",
      "count": 1
    }
  ]
}
```

| Field | Required | Meaning |
|---|---|---|
| `schema_version` | yes | Must be `1`. |
| `mode` | no | `"edit"` is informational; the action already knows the document type. |
| `equations[].index` | yes | 1-based document-order ordinal from the inventory action. Must be unique. |
| `equations[].old_tex` | recommended | Exact TeX the equation must currently hold. A mismatch aborts the whole run. |
| `equations[].new_tex` | one of | Replacement TeX, on one line. |
| `equations[].delete` | one of | `true` removes the equation object and nothing else. |
| `prose[].find` | yes | Exact, case-sensitive text to find. |
| `prose[].replace` | yes | Replacement text. |
| `prose[].count` | no | Asserted number of occurrences; a different count aborts the run. |

Working rules:

- Exactly one of `new_tex` and `delete: true` per entry.
- The action preserves each equation's own inline/display delimiter. It never promotes an inline
  symbol to a display equation or the reverse.
- Equations that are not listed are not touched at all, so numbering fields, tab stops, and paragraph
  formatting survive.
- A failed verification or an index that does not exist aborts before publishing; the source file is
  left intact and no output file is written.
- `skip_verify` disables only the `old_tex` comparison. Use it when the cached TeX is known to be
  stale, never to silence an unexpected mismatch.

Responses report `rewritten`, `deleted`, `untouched`, the per-entry prose results, and a
`validation` block produced by the same structural checks used by the validate action.
