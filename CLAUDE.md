# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Single-page browser tool for tagging "안건(agenda)" sections inside Korean HWPX (한컴 Office) documents. Operator opens a `.hwpx`, picks a standard tag from the left pane, then clicks/drags in the rendered document on the right to attach paragraphs or table cells to that tag. Output is either a tagged `.hwpx` (with HWP "누름틀" / CLICK_HERE click-here fields embedded) or a JSON sidecar.

There is **no build system, no package manager, no test suite, and no dependencies**. The entire app is one file: [index.html](index.html). To run, open `index.html` in Chrome/Edge — the rest is browser APIs.

`*.hwp` and `*.hwpx` are gitignored; sample documents at the repo root (`샘플서식.hwpx`, `샘플서식_tagged (1).hwpx`) are local-only.

## Browser requirement

Requires `DecompressionStream`/`CompressionStream` (`deflate-raw`). Chrome/Edge only — Safari/older Firefox will fail at HWPX open with an explicit error message.

## Architecture (one file, but several layers)

The script in [index.html:499-2199](index.html#L499) is organized in distinct layers; understand them as separate subsystems even though they share globals:

1. **Standard storage** ([index.html:503-552](index.html#L503)) — Tag definitions (`{ id, label }[]`) live in `localStorage` under `agenda-tagging-standard`, with a full mutation `history[]` (add/delete/modify diffs + editor + memo). `DEFAULT_TAGS` is only the seed for first-run. Everything that mutates the standard MUST go through `applyStandard()` so dependent state (active tag, `tagLocations` for deleted tags, expanded set) is reconciled and the change is persisted + rendered.

2. **ZIP read/write** ([index.html:631-785](index.html#L631)) — Hand-rolled ZIP reader and writer (no JSZip). Critical invariant: when re-saving HWPX, **unmodified entries reuse their original compressed bytes** (`entry.original.compressed` + `crc32` + `compressedSize`). Only modified section XMLs are re-compressed. This avoids byte-level drift in untouched parts (signatures, manifests, mimetype) that would cause some HWP readers to reject the file. Don't "simplify" `buildZip` to always re-deflate.

3. **HWPX parsing & style mapping** ([index.html:837-1149](index.html#L837)) — `parseStyles()` reads `Contents/header.xml` to build `Map`s for `charPr`, `paraPr`, `borderFill`, and font faces. These are rendered to inline CSS by `charShapeToStyle` / `parShapeToStyle` / `borderFillToCSS`. HWPX uses HWPUNIT-based measurements: divide by ~75 to convert to CSS px, divide `height` by 100 to get pt. Korean font names get fallback chains via `HANCOM_FONT_FALLBACK`.

4. **Render & node mapping** ([index.html:1153-1310](index.html#L1153)) — `renderHwpx()` walks each section's XML and produces HTML. The crucial output is `state.xmlNodeMap`: a flat array indexed by `xmlIdx` that links each rendered DOM element back to its source XML node, section index, and (for paragraphs inside cells) the parent cell's xmlIdx. Every selectable HTML element carries `data-xml-idx`. Cells are pushed before their inner paragraphs so xmlIdx ordering matches DOM order — both selection dedup and the export sort rely on this.

5. **Selection model** ([index.html:1312-1448](index.html#L1312)) — Two unit types: `'p'` (paragraph) and `'tc'` (table cell). Click and drag follow distinct rules — and they are user-visible behavior, documented in the in-app help modal ([index.html:317-481](index.html#L317)). Don't change them without updating both:
   - **Single-paragraph-cell promotion**: clicking a `p.p-in-cell` whose parent `td` has exactly one paragraph auto-promotes the unit to the cell ([index.html:1322-1333](index.html#L1322)).
   - **Drag promotion rules** (`rangeToUnits`): one paragraph → 1 unit; multiple paragraphs (outside tables) → N separate units; one cell with multiple paragraphs → cell unit; multiple cells or a drag that crosses a table boundary → entire table.
   - **Containment dedup** in `applyUnits`: if a `tc` is selected, any `p` inside it is dropped automatically.

6. **Field insertion** ([index.html:1468-1649](index.html#L1468)) — `insertFieldInParagraph()` wraps tagged text in HWP `fieldBegin type="CLICK_HERE"` + `fieldEnd` ctrl runs (HWP "누름틀"), preserving the original namespace prefix from the parent `<p>`. Three matching strategies, in order: full-paragraph match → run-boundary combination match → character-level match with run splitting. `partIndex/partCount` and a `metaTag` JSON payload (`schema: agenda-tagging/v1`) are embedded so the tags can be round-tripped.

7. **Export** ([index.html:1654-1754](index.html#L1654)) — `exportHwpx()` clones the parsed DOM via serialize→reparse so the in-memory render isn't mutated, walks units in `xmlIdx` order (= document order), then rebuilds the ZIP reusing original entries for everything except modified sections.

## Key state shape

Single mutable `state` object ([index.html:557-566](index.html#L557)):
- `tagLocations: { [tagId]: Unit[] }` where `Unit = { type:'p'|'tc', xmlIdx, text, rawText, sectionIdx }`
- `xmlNodeMap[xmlIdx] = { type, node, sectionIdx, parentTcIdx? }` — the bridge between DOM and XML
- `xmlSections[]` — parsed `Contents/sectionN.xml` documents (sorted numerically)
- `rawEntries[name]` — raw ZIP entry metadata, kept around so unmodified entries can be repacked verbatim
- `activeId` / `activeMode` ('replace' | 'append') drive the click/drag handlers

`TAGS` and `STANDARD` are top-level globals derived from `loadStandard()`; reassign them only via `applyStandard()`.

## When editing

- **All text and labels are Korean**; preserve Korean strings in user-facing messages and the help modal.
- The in-app help modal ([index.html:317-481](index.html#L317)) documents the selection rules to end users — keep it in sync if you change `rangeToUnits`, `elementToUnit`, or the click/append/clear button behavior.
- Don't introduce build tooling or external dependencies; the "single static HTML, no install" property is the point. There is intentionally no `package.json`.
