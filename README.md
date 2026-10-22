# msstyleEditor [![Latest Release](https://img.shields.io/github/release/nptr/msstyleEditor.svg)](https://github.com/nptr/msstyleEditor/releases/latest) [![License: MIT](https://img.shields.io/badge/License-MIT-brightgreen.svg)](https://opensource.org/licenses/MIT)

The msstyleEditor is an editor for Windows Vista, 7, 8, 10 and 11 visual styles (.msstyles files). It allows you to change visual styles without using a hex editor or a PE resource tool. It lists all components, can add/remove/modify the majority of properties, as well as extract and replace images.


## Installation
The application is a single, portable executeable. It needs no installation but requires .NET Framework 4.5 on your PC. That comes preinstalled with Windows 8+, and likely exists on your Windows 7 already as well.

In order to apply the visual styles, your Windows must also be patched with one of the following tools, because Microsoft locks-out third party themes.
* SecureUxTheme: https://github.com/namazso/SecureUxTheme/releases
* UltraUXThemePatcher: https://mhoefs.eu/software_uxtheme.php


## Limitations
* can't create styles from scratch
* can't add new fonts
* no preview of items in image atlases
* supports only .png ' s

No abstraction over the internals is done. Things are named and structured as they were by the developers at Microsoft. It may be a bit puzzling at first to figure out what each class/part/property effects. The following chapter might help a bit though.

## Visual Style Structure

Understanding of the logical structure of visual styles is helpful when using this program. Fortunately its quite simple, so the figure below should give you the basic understanding required.
```bash
├─ Class 1 (e.g. Button, Window)
│   └─ Part 1 (e.g. Pushbutton, Left Frame)
│      ├─ State 1 (e.g. Pressed, Disabled)
│      └─ State 2
│          ├─ Property 1 (e.g. BackgroundColor, Margins)
│          └─ Property 2
├─ Class 2
│    ├─ Part 1
│    └─ Part 2
└─ Class 3
```

You will also encounter classes with :: in their names, such as `BrowserTab::Tab`. This means that the
`BrowserTab` class inherits (reuses) the parts, properties and states of `Tab`, and just defines a few that override existing ones or are an addition. With this concept, duplication of properties is minimized and further provides a single place for common properties.

## User Interface Description

![Ui of the msstyleEditor](https://user-images.githubusercontent.com/5485569/154944873-e6f03424-6c8f-4114-8f15-6f1af7343f1b.png)

In the treeview on the left, the classes, parts and images are listed. On selection of an image, it is shown in the middle area. Right-click the image area to change the background if images are barely visible. Export and replace of the currently visible image can be done via the menu bar.

On selection of a part, its properties are shown on the right side, grouped by their states. This is also the place where the properties can be added, removed and edited. The search function is invoked via CTRL+F; it allows to search for classes, parts and properties with specific values.

Saving the style is done via the menubar. It is recommended to save often, and to a new file, since there is no undo/redo functionality. Also remember to backup your original style and don't work in the "Themes" directory directly (probably no write permission anyways, but still).

## Importing a Windows 7 style (experimental)

Open a Windows 10 style that already works on your machine (e.g. an Aero 7 port), click Extras > Import Win7 Style and pick a custom Windows 7 `.msstyles`. The open style gets the custom style's look, and the result is opened in the editor unsaved, so you can test it, keep editing and save it like any other style.

Stock Windows 7 `aero.msstyles` is built into the editor and used as the reference: only what the custom style changed relative to stock is carried over, and only where the open style still has the stock value. Everything else in the open style (class map, inheritance, animations, dark mode parts, every untouched resource) stays byte for byte. Images are compared by pixels, so a template that re-encoded the stock images still counts as stock.

DWMWindow (frames, caption buttons, glyphs) is matched by part name, not by part number, because the numbering changed after Windows 7. This only works if the template uses the stock Windows 7 DWM atlas; otherwise the template frames are kept and the report says so. Windows 10 parts without a Windows 7 counterpart (dark glyphs, extra caption backgrounds) follow the stock sprite they pointed at. Sprites that can't be traced keep the template's pixels.

If offline validation fails, nothing changes and the report says why. Every run writes a `.md` and `.json` report plus a `.threeway.json` with the full stock / custom / template comparison ("Open report folder" in the result window). The template's `Shell` folder and `<lang>\*.msstyles.mui` files are copied next to the result (renamed to match), since Windows can't load the shellstyle without them.

The same thing runs headless:

```
msstyleImport import --source custom.msstyles --template win10.msstyles --report out\report [--write --out out\candidate.msstyles]
```

`--prefer-source Class|Part|Property` takes the custom value where the template deliberately changed it, `--exclude` skips matches (`*` wildcards).

The CLI does a dry run (reports only) unless `--write` is given; `--base` swaps in a different stock reference.

Limitations:

* Offline validation is not a render test. Try the result in a disposable VM with a snapshot before using it for real.
* Win7-only parts (Flip3D, reflection, side highlight) have no Windows 10 counterpart and are dropped.
* Win10 templates add 192-384 dpi images that a Windows 7 style never had. Where the custom style replaced a part's pictures, those extra slots get its largest picture; parts it left alone keep the template's high-dpi art.
* Classes that don't exist in the template are reported, not added. Properties the custom style adds on top of stock Windows 7 (text glow, fonts, colors...) are added to the template's matching class.
* Colors always come from the custom style, even where the template changed them on purpose, and a color the custom style deletes is deleted too. Where a Windows 10 class adds its own value on top of a class the custom style changed (e.g. Explorer's buttons on top of the normal button colors), the custom value is copied into it. Some things Windows 10 draws from a different property than Windows 7 did (navigation pane, file list text, reading pane, status bar, size grip); those get the custom style's Windows 7 value too. Dark-mode colors stay the template's, to match its dark backgrounds. Other conflicts keep the template's value unless `--prefer-source` says otherwise.
* Apply a complete `.theme` when checking system-color controls. Loading only the `.msstyles` can leave the previous theme's system palette active. The importer preserves source rendering, including plain labels that are already black on dark backgrounds on Windows 7; it does not invent replacement colors for those.
* A picture the template redrew still comes from the custom style when the part uses the custom style's frame count and margins, since that's what those numbers describe.
* Removals never leave a class without records: Windows only registers classes that have records and the class inheritance table counts them in order, so an emptied class breaks inheritance for every class after it (fonts falling back to the System font, for one). Such a color keeps the value the deletion would have fallen back to instead.
* Fonts are compared by what they parse to, not their spelling. A font the template doesn't have is added to its string tables, which means the result is saved with the strings inside the `.msstyles`. A font string Windows would reject (an unknown style word fails the whole theme) is never written, and a caption/menu font taller than its bar blocks the result.
* Pictures the template reuses from stock Windows 7 elsewhere (dark mode classes, extra DPI sizes, shared image ids) get the custom style's version too, unless the custom style gave the parts that shared that picture different looks.

## Contributors
* [Shomnipotence](https://github.com/Shomnipotence) - New application icon design