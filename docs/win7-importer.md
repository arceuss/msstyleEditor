# Win7 -> Win10 importer, developer notes

Code: `libmsstyle/Import/`. CLI: `msstyleImport/`. UI: `MainWindow.OnImportWin7Style` plus the result window `msstyleEditorSharp/Dialogs/Win7ImportDialog.cs`. Tests: `libmsstyleTests/Win7ImportTests.cs` (the Shine test looks for `MSSTYLE_SHINE` or the default Downloads path and is inconclusive without it).

The base (stock win7 aero) is `libmsstyleTests/styles/w7_aero.msstyles`, embedded into libmsstyle as `libmsstyle.Import.w7_aero.msstyles`. `Win7StockAero.GetPath()` writes it to `%TEMP%\msstyleEditor\w7_aero-<sha>.msstyles` because `VisualStyle.Load` goes through LoadLibraryEx. It has no .mui next to it, so base font/string properties can only be compared by id: a source that rewrote the text behind an unchanged id isn't detected (noted in every report).

In the editor the template is the open style: it's saved (standalone) to `%TEMP%\msstyleEditor-import\<time>\template\`, so unsaved edits count, the import writes `...\out\<name>-win7import.msstyles`, and on success that file replaces the open style. On failure the open style is left alone.

## Pipeline

1. `VisualStyle.Load` all three; any `LoadDiagnostics` entry blocks the run. Platform must be Vista/7 for source and base, 10/11 for the template.
2. `SemanticModel` flattens each style to `class|part name|S<state>|property#type[@n]`. Images compare by decoded pixels, string ids (FONT and non-image FILENAME) by their string-table text when both sides have a table, by id otherwise. Only IMAGEFILE..IMAGEFILE7/GLYPHIMAGEFILE are IMAGE references; resource ids use the low word like uxtheme. Numeric class ids never take part in matching.
3. `BuildThreeWayDiff` records every key (written to `<prefix>.threeway.json`).
4. `DwmImport` handles DWMWindow, then `MergeGenericClasses` everything else with the same three-way rule:
   source == base -> keep template; template == base -> take source; template == source -> nothing; otherwise keep template (`--prefer-source` overrides).
   Images that differ only by +-2 premultiplied rounding count as equal; Aero11 re-encoded nearly every stock image that way (its real retouches differ by 13+).
   When the template shows a stock picture that belongs to another part (Aero11's TOPTABITEM uses the stock TABITEMLEFTEDGE strip), the source's version of that other part is used (`FindImageAlias`, `applied-alias`), same idea as the DWM alias.
   Strip geometry stays consistent: if IMAGECOUNT/IMAGELAYOUT/SIZINGMARGINS is a kept conflict, or the source changes one of them while some image of the part stays the template's, the whole part keeps the template geometry (`LockedGeometryParts`). Images on their own are decided one by one. The validator wants every changed strip image to divide into its effective IMAGECOUNT, except where the source ships that exact picture with that count (Pachenko's 9x121 trackbar thumb in 5): win7 drew those and no uxtheme check rejecting a remainder was found, so they're warnings (`CollectFrameWaivers`).
   After the merge, `PropagateImages` walks template images nobody touched (win10-only classes, dark mode copies, IMAGEFILE4-7) that aren't still their own key's stock picture: a stock win7 picture -> the source's version of it (`applied-stock-copy`), and an IMAGE id shared with an already replaced property -> the same replacement when all replacements of that id agree (`applied-shared-id`). Ambiguous stock pictures are narrowed to the same part/state/property, then the same base class name, then the class the template inherits from (BCMAP).
   Fonts are compared as parsed `FontSpec`s (grammar from uxtheme `_ParseFont`). A font the template has no string for is added to every string table with a fresh id and the candidate is saved standalone. The validator parses every font, checks all old strings still read the same and checks caption/small caption/menu fonts against their bar heights.
   Records the source adds where stock win7 has none (text glow on the address bar, a tree view font...) are added to the template (`added`, `TryAddRecord`) when the class exists there; a missing part/state section is created. The value goes through the same translation as any applied value; the validator accepts exactly these added records on top of the template's.
5. New bytes go through `ResourceIdAllocator` (max id + 1, shared when bytes are identical; ids with several language variants are never shared). Existing resources are never rewritten.
6. `VisualStyle.Save` into a validation folder of its own, with `StyleCompanions.CopyAlongside` putting the template's `Shell` folder and `<lang>\<name>.msstyles.mui` next to it (mui-based templates only resolve their string tables that way), then `StyleValidator` against the template. On success the file is moved to the output and the companions are copied there too; the validation folder is always deleted. The editor passes the opened file as `CompanionSourcePath`, because its template is a temp snapshot.

## DWMWindow

The template atlas has to be the stock win7 atlas (pixel-compared, `TemplateAtlasIsStockWin7`). Then:

- a template rect equal to the stock rect of the same-named win7 part maps by name;
- otherwise the rect is looked up among stock win7 rects (alias). Several win7 parts may share one stock rect; that's only followed if the source gave all of them the same pixels (`ChooseAlias`);
- anything else keeps the template sprite, copied into a strip below the source atlas (`fallback`), and the part keeps its template properties.

A mapped part takes the source sprite at whatever size it has, plus the numbers that describe it (IMAGECOUNT, SIZINGMARGINS, CONTENTMARGINS, `applied-with-sprite`), even where the template had adapted them: those belonged to the template's sprite. That's the same thing as replacing a part in the editor. The fallback (template sprite copied under the source atlas) is only for sprites that can't be traced, or parts excluded by rule.

`CheckGeometry` then only warns (`GeometryWaivers`, `dwm-geometry-warning`). Strips are taken as sliced proportionally like the editor slices them (129px in 2 = 64 + 65; uDWM's rounding isn't verified), so odd sizes are fine; what gets a warning is a strip with nothing left after uDWM's 1px inset, or margins wider than the strip. Since those numbers came with the source sprite, the source shipped them like that on win7 (Dream's LEFTFRAME, Metalik's TOPFRAME). Exclude rules work here too: an excluded atlas skips DWMWindow, an excluded part's ATLASRECT keeps the template sprite.

Aero11 Seven Clear (data, checked by `Win10DwmPartNamesPointAtTheSameStockSpritesAsWin7`): its DWM atlas is byte-identical to win7 STREAM 971, and 59 of 60 shared part names sit on the same stock rect. The exception is SMALLBUTTONCLOSEGLYPH192, which points at the BUTTONRESTOREGLYPH192 sprite; the importer keeps that pairing and warns. Parts 64-91 (dark glyphs, dark caption backgrounds) reuse light stock sprites; parts 92-94 are not in the editor's table and reuse BUTTONCLOSEGLOW / BUTTONGLOW / SMALLBUTTONCLOSEGLOW.

## Binary evidence (uxtheme 10.0.19041.7725, uDWM 10.0.19041.7548, x64)

Static IDA reading only; nothing was executed. Other Win10 builds weren't checked.

| fact | where | confidence |
|---|---|---|
| load order CMAP -> RMAP -> CMAP data -> BCMAP -> AMAP | `VSLoad` 0x171B57E00 | verified |
| record = 8 dwords (name, type, class, part, state, external id, reserved, size), payload at +32, stride align8(32+size) when external id is 0, else 32 | `GetNextVSRecord` 0x171B54384, `LoadVSRecordData` 0x171B58008 | verified |
| IMAGE / STREAM resource name is the low word of +20; FONT and string ids use the full dword | `LoadImageFileRes` 0x171B59380, `LoadDiskStreamRes` 0x171B59260, `LoadFontRes` 0x171B592F0 | verified |
| class/part/state groups must be contiguous; a reopened group fails with a duplicate-section error | `LoadClassDataMap` 0x171B54E48 | verified (validator checks it) |
| no ascending property-id order required inside a group | same | not found |
| VARIANT name comes from VMAP, not a hardcoded NORMAL | `GetClassData` 0x171B541F8 | verified (importer keeps VMAP and VARIANT/NORMAL as they are) |
| PACKTHEM_VERSION/1 must be 2 bytes with value 4 | `LoadThemeLibrary` 0x171B52598 | verified (validator checks it) |
| CMAP entries 8-byte aligned on x64, BCMAP = count + parents, -1 falls back to globals | `CVSUnpack::Initialize` 0x171B549EC, `LoadBaseClassDataMap` 0x171B54DBC | verified |
| BCMAP entry *i* and parent value *p* index the **load registration order**, not CMAP order: classes with VARIANT records in record order, then AMAP-only classes. On Aero11 that's `[4..333, 336, 337, 338, 334, 335]`, so the editor's fixed "+4" (`VisualStyle.LoadBaseClassMap`, `BcMap`) is right for 4..333 and wrong for the last five (timingfunction/animations area). The importer never rewrites CMAP/BCMAP and its inheritance tie-break only looks at image classes, so it isn't affected; the editor's inheritance display and add-class save path are | 18362 (10.0.18362.1645, typed) `LoadClassDataMap` 0x171B51398, `AddIndexInternal` 0x171B3CE18, `LoadAnimationDataMap` 0x171B510EC, `CopyLocalThemeToLive` 0x171B3D4F8; same mechanism in 19041 0x171B3F7C8 / 0x171B54B98 / 0x171B3FECC | verified (formula inferred from verified code) |
| win7 (6.1.7600.16385) vs 19041: font grammar, unknown-style failure, record layout, low-word resource ids, PACKTHEM_VERSION 4, CMAP alignment and sysmetrics mapping are identical. Differences: IMAGEFILE6/7 (3009/3010) are string ids on win7 and images on win10 (the importer refuses to copy across that), an external ATLASINPUTIMAGE (8001) string must decode to 16 bytes on win10 (not copied), and win7 has no AMAP stage | 7600 `_ParseFont` 0x7FF72EF4F80, `GetThemePrimitiveID` table 0x7FF72F21F00, `VSLoad` 0x7FF72EFC610 | verified |
| uDWM opens "DWMWindow", atlas via GetThemeStream(part 0, state 0, 213); sprites via ATLASRECT 8002, SIZINGMARGINS 3601, CONTENTMARGINS 3602, all at state 0; the rect is shrunk by 1px on every side | `CreateBitmapFromAtlas` 0x180049788 | verified |
| uDWM part ids 1-91 agree with `PARTS_DWMWINDOW_WIN10`, glyph DPI order 96/120/144/192, caption backgrounds requested as 3,5,4,6,88,90,89,91 | uDWM frame/button tables | verified; names of 88-91 inferred from slot pairing |
| frame text colors come from part 46 states 1/2, property 3803 | uDWM | verified |
| font grammar: `Face, size, styles`; style words bold/italic/underline/strikeout/quality:{cleartype,cleartype-natural,antialiased,nonantialiased}, case-insensitive; an unknown word fails the whole theme load (0x80070648); size ends at the next comma (`Face,9 Bold` drops Bold); face cut at 31 chars, string at 127 | `_ParseFont` 0x171B59A6C, `LoadFontRes` 0x171B592F0 | verified |
| font height `-MulDiv(points,96,72)`, scaled to screen dpi later | `_ParseFont`, `ScaleFontForScreenDpi` 0x171B4AE18 | verified |
| sysmetrics 801-806 fonts and 1201-1210 sizes go straight into NONCLIENTMETRICS / icon font, no check of fonts against bar heights | `CThemeLoader::PackMetrics` 0x171B41E2C, `SetSystemMetrics_Worker` 0x171B424BC | verified (the importer's font/bar check is its own policy) |

Full notes with call sites: `C:\gpttmp\msstyle-port\ida_findings.md` (scratch, not in the repo).

## Parser fixes made along the way

- `PropertyStream.ReadNextProperty`: opaque (undecoded) payloads copied `data.Length - cursor` bytes into a `sizeInBytes` buffer, which throws for any record that isn't last. The load loop swallowed the exception without advancing, i.e. spun forever. Now copies exactly `sizeInBytes`, bounds-checked.
- The header scan could run past the end of the stream; it now stops and reports the skipped tail.
- `VisualStyle.LoadProperties` records every skipped or undecodable record in `LoadDiagnostics` instead of dropping it silently, and stops instead of spinning when the cursor can't advance.
- Records now step to the next 8-byte boundary when the gap is zero padding (uxtheme's `align8(32 + size)`). Before, the scan tolerated at most 4 skipped bytes, so a legal empty string (6 padding bytes) counted as damage.
- Nonzero bytes in the last 1-4 bytes of VARIANT are reported instead of ignored.
