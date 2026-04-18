# Word → Brightspace HTML Converter

Converts `.docx` files into clean HTML ready to paste into Brightspace (D2L).

---

## Features

| Word Feature | HTML Output |
|---|---|
| Heading 1–6 | Configurable tag (h2–h6) via the GUI |
| Bullet lists | `<ul><li>` (nested levels supported) |
| Numbered lists | `<ol><li>` (nested levels supported) |
| Bold / Italic / Underline / Strikethrough | `<strong>`, `<em>`, `<u>`, `<del>` |
| Hyperlinks | `<a href="...">` |
| Blockquote (Quote / Block Text style) | `<blockquote><p>` |
| Accordion (single-cell table) | D2L accordion card HTML |
| Module headings (`Module 1:`, `Module 2:`, …) | One HTML file per module |

---

## Setup (one-time)

You need **Python 3.8+** installed. No other manual setup is required.

When you run the converter for the first time, it will automatically detect any missing packages and show a graphical installer window — just click **Install** and the app will handle the rest before opening.

If you prefer to install manually beforehand:

```bash
pip install python-docx
pip install tkinterdnd2   # optional — adds drag-and-drop support
```

---

## Running

```bash
python converter.py
```

On first launch, if any required packages are missing a **📦 Missing Python Packages** window will appear. It lists what needs to be installed, shows live pip output as packages download, and opens the main app automatically once installation is complete. If a package fails to install, a **Retry** button appears. You can also click **Skip** to proceed without installing (the app may not function correctly if required packages are absent).

Once open, click the file area to pick your `.docx` (or drag and drop one in), adjust settings if needed, then click **Convert**.

The output `.html` file(s) are saved in the same folder as your Word document (unless you uncheck that option).

---

## Brightspace Preview

The **Rendered** preview panel is styled to match a real Brightspace content page, so what you see in the app closely reflects what students will see:

- A dark charcoal navbar at the top showing a **Course Home › Content › filename** breadcrumb, matching D2L's standard chrome
- A grey page shell with a white content card floating inside it, matching D2L's content page layout
- Lato typeface, D2L's default link blue (`#006fbf`), and heading styles consistent with D2L's defaults

When you use **Open in Browser**, the exported HTML uses the same stylesheet, so the browser view is consistent with the in-app preview.

> The preview approximates D2L's default appearance before institution-level theming (custom colours, logos, fonts) is applied. If your institution has a heavily customised theme, you can load your own `.css` file in **Settings → Preview CSS** to override the defaults.

---

## Image Extraction

Embedded images in your `.docx` are automatically extracted and saved to an `images/` subfolder alongside your HTML output. The converted HTML references them with relative paths:

```html
<img src="images/image1.png" alt="image1.png" style="max-width:100%;">
```

**What gets extracted:** any image inserted via Word's Insert → Pictures, including inline and floating images. Duplicate images (the same file embedded multiple times) are de-duplicated automatically.

**The status bar** reports how many images were extracted after each conversion, e.g. `✅ Saved: MyDoc.html (3 images → images/)`.

**In the Rendered preview**, images appear as a placeholder showing the filename: `🖼️ [image1.png]`. Use **Open in Browser** to see the actual images rendered.

### Uploading to Brightspace

After converting, your output folder will look like:

```
MyDoc_module_01.html
MyDoc_module_02.html
images/
  image1.png
  image2.png
```

Upload both the HTML files and the `images/` folder to the **same location** in Brightspace's Manage Files. As long as the folder structure is preserved, the relative `src="images/..."` paths will resolve correctly. If Brightspace flattens the folder on upload, the image paths will break — in that case you will need to re-point the `src` attributes manually or use a zip import that preserves folder structure.

---

## Module Splitting

When the **Split into one file per Module** option is enabled (default: on), the converter detects headings that match the pattern **"Module N:"** — where N is any number from 1 to 99 — and splits the document into one HTML file per module.

**Example heading patterns detected:**
- `Module 1: Introduction`
- `Module 12: Advanced Topics`

**Output filenames** follow the pattern `MyDoc_module_01.html`, `MyDoc_module_02.html`, etc.

If no Module headings are found in the document, a single `.html` file is produced as normal.

### Module Preview Navigation

When modules are detected, a **◀ Module N of M ▶** navigation bar appears in the preview toolbar. Click the arrows to step through each module's content without re-converting.

The preview always reflects the currently selected module. Switching between **HTML Source** and **Rendered** modes also respects the current module.

---

## Accordion Support

To create accordion output, place your content inside a **single-cell Word table**. The converter detects these automatically — no special styles or plugins required.

| Element | How to create it |
|---|---|
| **Card title** | Use a **Heading 4** paragraph inside the cell |
| **Card body** | Use normal body paragraphs after the Heading 4 |

Each Heading 4 paragraph inside the cell starts a new accordion card. The body paragraphs that follow it (up to the next Heading 4, or the end of the cell) become that card's content. Multiple Heading 4 / body pairs in the same cell produce multiple cards inside a single `<div class="accordion">` wrapper.

**Example structure (all inside one single-cell table):**

```
Heading 4 → "What is Brightspace?"
Normal paragraph → "Brightspace is a learning management system..."

Heading 4 → "How do I enroll?"
Normal paragraph → "Contact your institution's registrar..."
```

**Output:**

```html
<div class="accordion">
  <div class="card">
    <div class="card-header">
      <h2 class="card-title">What is Brightspace?</h2>
    </div>
    <div class="collapse">
      <div class="card-body">
        <p>Brightspace is a learning management system...</p>
      </div>
    </div>
  </div>
  ...
</div>
```

**Changing the trigger heading**

By default, **Heading 4** is used as the card title style. You can change this in **Settings → Accordion Card Title Style** if your document uses a different heading level.

> **Note:** A table is only treated as an accordion if it is a single-cell table containing at least one paragraph using the accordion heading style. All other tables are converted as standard `<table>` HTML.

---

## Heading Map

The **Settings** tab lets you remap each Word heading level to any HTML heading tag. For example, if your Brightspace page already has an `<h1>` title, you might want:

- Heading 1 → `h2`
- Heading 2 → `h3`
- Heading 3 → `h4`
- and so on.

Choose `(skip)` to drop that heading level from the output entirely.

Your settings are saved automatically between sessions.

---

## Output Options

- **Wrap in full HTML document** — adds `<!DOCTYPE html>`, `<html>`, `<head>`, and `<body>` tags. Useful for previewing in a browser. Leave unchecked for a snippet you paste directly into Brightspace's HTML editor.
- **Save output alongside source** — saves `MyDoc.html` (or `MyDoc_module_01.html` etc.) next to the source `.docx`. Uncheck to choose a custom save location each time you convert.
- **Split into one file per Module** — when Module headings are detected, saves each module as a separate file. Uncheck to always produce a single combined file.

---

## Profiles (Presets)

The **Settings** tab includes a **Profiles** dropdown that lets you save, load, rename, and delete named sets of settings.

- **First launch:** a default **Profile 1** is created automatically.
- **Saving:** click **Save…**, enter a name (or pick an existing one to overwrite), and confirm.
- **Loading:** select a profile from the dropdown — settings apply immediately.
- **Persistence:** the last-used profile is remembered across sessions and restored automatically when you reopen the app.

Profiles store: heading map, list transforms, blockquote style, accordion heading, and CSS file path.

---

## Batch Conversion

Switch to **Batch** mode using the pill toggle at the top of the Convert tab to convert an entire folder of `.docx` files at once.

### Workflow

1. The folder bar defaults to the **current working directory** when you open the app. Click **Browse…** to pick a different folder — all `.docx` files in it are listed automatically.
2. Each file has a checkbox. Use the **Select All / None** links to quickly check or uncheck everything, then uncheck any files you want to skip.
3. Click **Convert All →**. Files are converted one at a time so the UI stays responsive. A progress bar tracks the run.

### Status indicators

Each filename gets a status badge as it is processed:

| Badge | Meaning |
|---|---|
| ⏳ | Currently converting |
| ✅ | Converted successfully (single file) |
| ✅ ×N | Converted successfully — N module files saved |
| ❌ | Conversion failed — hover the filename for the error |

Successful filenames turn green; failed ones turn red.

### Output

Batch mode always saves output **alongside each source file**, following the same settings as Single mode — module splitting, full HTML wrapping, and image extraction all apply exactly as configured.

### Previewing results

Click any filename in the list after conversion to load it into the preview panel on the right. This works the same as loading a file in Single mode — the HTML Source and Rendered views both update, and module navigation appears if modules were detected.

---

## Conversion Log

A **Conversion Log** panel sits at the bottom of the Convert tab, below the Refresh Preview button. It reports exactly what happened during each conversion run.

**What it shows:**

- A summary of every element type converted — headings, paragraphs, list items, blockquotes, tables, accordions
- Images extracted and their filenames
- Elements skipped because a style is set to `(skip)` in Settings
- Unrecognised Word styles that were treated as plain `<p>` (useful for spotting styles you may want to remap)
- File paths and sizes of every HTML file written to disk
- The path of the `images/` folder when images are saved

**Severity levels** are colour-coded:

| Symbol | Colour | Meaning |
|---|---|---|
| `·` | Muted | Info — everything converted as expected |
| `⚠` | Amber | Warning — something was skipped or fell back to a default |
| `✕` | Red | Error — a file couldn't be written or a conversion failed |

**The header badge** (`OK`, `2 warnings`, `1 error`) gives an at-a-glance summary without opening the panel.

**Auto-expand behaviour:** the panel collapses automatically after a clean conversion and expands automatically when there are any warnings or errors, so it never gets in the way during normal use but is always visible when something needs attention.

Click the **▶ CONVERSION LOG** header to toggle the panel open or closed manually at any time.

---

## Copy HTML

Click **Copy HTML 📋** in the bottom-right of the preview panel, or press **Ctrl+Shift+C**, to copy the converted HTML directly to your clipboard.

- The button briefly changes to **✓ Copied!** as confirmation, then resets.
- When modules are detected, only the **currently displayed module** is copied — since each module is pasted into its own Brightspace content page separately.
- When no modules are present, the full document body is copied.
- The clipboard always receives a **bare HTML snippet** (no `<!DOCTYPE>` wrapper), regardless of the "Wrap in full HTML document" setting — Brightspace's HTML editor expects a fragment, not a full document.

---

## Search (Ctrl+F)

Press **Ctrl+F** while the app is focused to open a search bar at the top of the preview panel. It works in both **HTML Source** and **Rendered** modes.

- Type to search incrementally — all matches are highlighted in amber.
- The current match is highlighted in orange and scrolled into view.
- Press **Enter** / **▼** to jump to the next match; **Shift+Enter** / **▲** for the previous.
- Press **Escape** or click **✕** to close the search bar.

---

## Preview CSS (optional)

You can load a `.css` file in the **Settings** tab to style the preview panel and the full-document HTML output. Loading your Brightspace theme's stylesheet gives you an accurate preview of how the converted page will look to students.

CSS is only embedded in the output when **Wrap in full HTML document** is enabled.

---

## Drag and Drop

Drag a `.docx` file from your file manager and drop it anywhere on the application window. The drop zone highlights as soon as you bring a file over the app, giving clear visual feedback before you release. Requires the `tkinterdnd2` package, which the automatic installer will offer to install on first launch.

---

## Building an Executable

A `build.py` script is included to compile the converter into a standalone `BrightspaceConverter.exe` using PyInstaller. All dependencies are bundled into the executable — end users do not need Python or any packages installed.

```bash
python build.py           # standard build
python build.py --fast    # skip cleaning, rebuild only changed files
python build.py --debug   # build with console window for troubleshooting
```

The finished executable is written to `dist/BrightspaceConverter.exe`. The dependency installer dialog does not appear in the built executable, since all packages are already bundled inside it.

**Prerequisites for building:**

```bash
pip install pyinstaller python-docx tkinterdnd2
```
