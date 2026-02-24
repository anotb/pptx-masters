---
name: pptx-masters
description: Extract PptxGenJS slide masters from corporate .pptx/.potx templates. Converts OOXML theme colors, fonts, placeholders, backgrounds, slide numbers, and static shapes into ready-to-use defineSlideMaster() JavaScript code. Use when building presentations programmatically with PptxGenJS and the user has a corporate PowerPoint template they want to replicate.
license: MIT
metadata:
  author: anotb
  version: "0.1.0"
compatibility: Requires Node.js 18+. Works on macOS, Linux, and Windows.
allowed-tools: Bash(npx:*) Bash(node:*) Bash(libreoffice:*) Bash(osascript:*) Bash(python:*) Read Write
---

# pptx-masters

Extract PptxGenJS `defineSlideMaster()` code from corporate PowerPoint templates.

## When to use this skill

Use when:
- The user has a corporate .pptx or .potx template and wants to create presentations programmatically
- The user is working with PptxGenJS and needs slide masters that match their corporate template
- The user mentions "slide masters", "PowerPoint template", "corporate slides", "PptxGenJS", or "programmatic presentations"

## Step 1: List available layouts

First, show the user what layouts are in their template:

```bash
npx pptx-masters <template.pptx> --list
```

Ask the user which layouts they want. Many corporate templates have 15-25+ layouts — half are redundant. Typical essentials: title/cover, content, content with subtitle, divider/section, and end/closing.

## Step 2: Extract selected masters

```bash
npx pptx-masters <template.pptx> -o ./slide-masters --layout "Title Slide" --layout "Content" --layout "Section Header"
```

Or by number (from `--list`):

```bash
npx pptx-masters <template.pptx> -o ./slide-masters --layout 1 --layout 2 --layout 5
```

Or extract all:

```bash
npx pptx-masters <template.pptx> -o ./slide-masters
```

This generates:

| File | Purpose |
|------|---------|
| `masters.js` | ES module with `defineSlideMaster()` calls + `createPresentation()`, `THEME`, `POS`, `CHART_COLORS`, `FONT`, `PALETTE` exports |
| `theme.json` | Full extracted theme (all colors with tint/shade variants, fonts, dimensions) |
| `SLIDE_MASTERS.md` | Agent-readable reference: master table, placeholder names, color/font/positioning docs |
| `report.md` | Detailed extraction report with warnings |
| `preview.pptx` | Sample deck using all generated masters (for visual comparison) |
| `media/` | Extracted background images, logos |

## Step 3: Review the extraction

1. Read `SLIDE_MASTERS.md` — this is the primary reference for creating presentations. It documents every master, placeholder name, color, font, and positioning constant.
2. Read `report.md` for warnings about unsupported features (gradients, grouped shapes, etc.).
3. Open `preview.pptx` alongside the original template to visually compare. See [QA: Converting to Images](#qa-converting-to-images) for automated comparison.

## Step 4: Build the presentation

Read `SLIDE_MASTERS.md` — it contains the complete Quick Start code, all master/placeholder details, and key conventions. The pattern:

```js
import { createPresentation, THEME, CHART_COLORS, FONT, POS } from './slide-masters/masters.js';

const pres = await createPresentation('Deck Title');

// Use placeholders for positioned elements
const slide = pres.addSlide({ masterName: 'CONTENT' });
slide.addText('Title', { placeholder: 'title' });

// Use POS for manual content placement
slide.addChart(pres.charts.BAR, data, { ...POS.body, chartColors: CHART_COLORS });

// Always specify fontFace — PptxGenJS masters don't cascade fonts
slide.addText('Note', { x: 0.5, y: 6, w: 5, h: 0.5, fontFace: FONT, fontSize: 10 });

await pres.writeFile({ fileName: 'output.pptx' });
```

**Key gotchas** (also documented in SLIDE_MASTERS.md):
- Always specify `fontFace: FONT` on every `addText()` call
- Footers, logos, and slide numbers are automatic — do not add them manually
- Use placeholder names from SLIDE_MASTERS.md (e.g., `{ placeholder: 'Title 1' }`)
- Use `slide.background = { color: THEME.brand }` for background overrides

## Step 5: Design first, export last

For iterative design work:
1. Create HTML previews matching slide dimensions and brand colors
2. Get user approval on the visual design
3. Only then generate the PPTX using masters.js

This avoids slow PowerPoint open/close cycles during design iteration.

## QA: Converting to Images

After generating a PPTX, convert slides to images for visual comparison with the original template. The approach depends on what software is available.

### LibreOffice (macOS, Linux, Windows — free)

Works everywhere LibreOffice is installed:

```bash
libreoffice --headless --convert-to png --outdir ./slides output.pptx
```

This creates `output-1.png`, `output-2.png`, etc. Compare side-by-side with the original template.

If output is poor quality (LibreOffice can struggle with corporate fonts), try PDF as an intermediate:

```bash
libreoffice --headless --convert-to pdf output.pptx
# Then convert PDF pages to images
```

### PowerPoint for Mac (macOS with Microsoft 365)

Higher fidelity than LibreOffice. Use AppleScript via `osascript`:

```bash
osascript -e '
tell application "Microsoft PowerPoint"
  open POSIX file "/absolute/path/to/output.pptx"
  set thePresentation to active presentation
  set slideCount to count of slides of thePresentation
  repeat with i from 1 to slideCount
    set theSlide to slide i of thePresentation
    set paddedNum to text -2 thru -1 of ("0" & i)
    save theSlide in POSIX file ("/absolute/path/to/slides/slide-" & paddedNum & ".png") as save as PNG
  end repeat
  close thePresentation saving no
end tell
'
```

**Important:** Use absolute paths. Create the output directory first.

### PowerPoint COM Automation (Windows with Microsoft 365)

Highest fidelity. Requires Python with `pywin32`:

```python
import win32com.client, os
pptx_path = os.path.abspath('output.pptx')
ppt = win32com.client.Dispatch('PowerPoint.Application')
pres = ppt.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)
for i in range(1, pres.Slides.Count + 1):
    pres.Slides(i).Export(os.path.abspath(f'slides/slide-{i:02d}.png'), 'PNG', 1920, 1080)
pres.Close()
ppt.Quit()
```

### Choosing the right approach

| Method | Fidelity | Platform | Requires |
|--------|----------|----------|----------|
| LibreOffice | Good | All | `libreoffice` CLI |
| PowerPoint for Mac | High | macOS | Microsoft 365 |
| PowerPoint COM | Highest | Windows | Python + pywin32 + Microsoft 365 |

For most QA, LibreOffice is sufficient. Use PowerPoint when exact font rendering matters.

## Limitations (v1)

- Gradient fills — dominant color used as fallback
- Pattern fills — foreground color as fallback
- Grouped shapes not extracted
- Animations/transitions not supported
- SmartArt/diagrams not supported
- 3D effects not supported

Check `report.md` for specific warnings about your template.
