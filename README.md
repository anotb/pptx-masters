# pptx-masters

Turn your corporate PowerPoint template into code that AI agents can actually use.

## The Problem

LLMs are good at writing slide content but consistently bad at corporate branding. Every company has a `.pptx` or `.potx` template with specific colors, fonts, logos, layouts, footer positions, copyright text. When you ask an LLM to "make a presentation in our brand," you get an approximation that doesn't quite match. Every generation is a fresh guess.

There are a few ways LLMs typically generate PowerPoint files, and each has real tradeoffs:

**Raw OOXML.** PowerPoint files are ZIP archives containing XML. In theory, an LLM can write the XML directly. In practice, the OOXML spec is ~6,000 pages (ECMA-376), a single slide can be 500+ lines of XML, and one wrong namespace or relationship corrupts the file. It's also extremely token-inefficient.

**python-pptx.** The most common Python library for working with .pptx files. It can load existing templates and modify them, which is a real advantage. But it can only modify placeholders that already exist in the template's layouts, and [adding new layout types is unsupported](https://github.com/scanny/python-pptx/issues/413). It's also Python-only, which means there's no way to prototype slides in HTML or do visual iteration in a browser. The repo has [500+ open issues](https://github.com/scanny/python-pptx/issues) and hasn't been updated since August 2024.

**PptxGenJS.** The best JavaScript library for generating slides. Clean API, active development, works in Node.js and browsers. Anthropic's [official pptx skill](https://github.com/anthropics/skills) uses it. The catch: it [cannot import existing .pptx templates](https://github.com/gitbrent/PptxGenJS/issues/712). It builds everything from scratch. To match your corporate template, someone has to manually translate every layout into code, measuring positions, extracting colors with tint/shade math, getting the footer placement right. Multiply that by 10-20 layout variants in a typical corporate template.

pptx-masters automates that last part. Point it at your template, get ready-to-use JavaScript code that replicates your slide masters. Your LLM handles content, your template handles branding.

For the full discussion of why LLM-generated slides are hard and how this approach works, see [LLM Slides Suck (But They Don't Have To)](https://unsol.dev/guides/llm-slides/).

## Quick Start

```bash
npx pptx-masters your-template.potx -o ./slide-masters
```

That's it. You get a folder of files your AI agent (or your own code) can use right away.

To see what layouts are in the template first:

```bash
npx pptx-masters your-template.potx --list
```

To extract only specific layouts by name or number:

```bash
npx pptx-masters your-template.potx --layout "Title Slide" --layout "Content" -o ./slide-masters
```

## What You Get

```
slide-masters/
  masters.js          JavaScript module with your slide masters as code
  theme.json          Full theme data (colors, fonts, dimensions)
  SLIDE_MASTERS.md    Reference doc for AI agents
  STYLE_GUIDE.md      Your design preferences (editable, never overwritten)
  report.md           Extraction report with any warnings
  preview.pptx        Sample deck for visual comparison
  media/              Background images, logos
```

**masters.js** is the main output. It exports functions and constants that create presentations with your corporate branding: backgrounds, placeholders, slide numbers, footers, the full color palette with tint/shade variants. Import it, call `createPresentation()`, add slides.

**SLIDE_MASTERS.md** is what makes this work with AI agents. It documents every layout, placeholder name, color value, and font in a format LLMs can read and follow. Point your agent at this file and it knows how to use your template correctly.

**STYLE_GUIDE.md** is yours to customize. Add your own dos and don'ts, typography rules, color usage notes, layout preferences. This file is never overwritten when you re-extract, so your edits persist across updates. Agents are instructed to check it.

## Using with AI Agents

The main use case. After extraction, give your AI agent access to the output folder. The agent reads `SLIDE_MASTERS.md`, imports `masters.js`, and builds presentations that match your corporate template.

Works with any agent that can run JavaScript: Claude Code, Cursor, Windsurf, Gemini CLI, etc.

This tool also ships as an [Agent Skills](https://agentskills.io)-compatible skill. Copy `skill/pptx-masters/` to your agent's skills directory for automatic template extraction support.

### Recommended Workflow

The generated `SLIDE_MASTERS.md` includes a workflow for agents:

1. **Draft content as HTML** for fast iteration and visual preview
2. **Visually review the HTML** to catch layout and spacing issues early
3. **Translate to PptxGenJS** using the exported masters, colors, fonts, positions
4. **Render to images and verify** each slide visually before delivering

This front-loads visual feedback to the cheapest part of the process (HTML) instead of regenerating PowerPoint files on every iteration.

## Using as a Library

If you're building your own tooling:

```js
import { createPresentation, THEME, CHART_COLORS, FONT, POS } from './slide-masters/masters.js';

const pres = await createPresentation('Q4 Strategy Review');

const slide = pres.addSlide({ masterName: 'TITLE' });
slide.addText('Q4 Strategy Review', { placeholder: 'title' });

await pres.writeFile({ fileName: 'output.pptx' });
```

Or use the extraction API directly:

```js
import { extract } from 'pptx-masters';

const result = await extract('template.potx', {
  layouts: ['Title Slide', 'Content'],
});

// result.masterData    PptxGenJS-ready master definitions
// result.themeColors   { dk1: '000000', lt1: 'FFFFFF', accent1: '4472C4', ... }
// result.themeFonts    { heading: 'Calibri Light', body: 'Calibri' }
// result.dimensions    { width: 13.333, height: 7.5 }
```

## What Gets Extracted

- **Theme colors.** All scheme colors (dark, light, accents, hyperlink), with full tint/shade resolution including lumMod, lumOff, and satMod calculations.
- **Fonts.** Heading and body typefaces from the theme.
- **Placeholders.** Title, body, subtitle, picture, chart, table, with exact positions, text styling, and alignment.
- **Slide numbers, footers, dates.** Position, font, size, color.
- **Backgrounds.** Solid colors, images (extracted to `media/`), gradient fallback to dominant color.
- **Static shapes.** Rectangles, lines, text boxes, images, with fill, border, rotation, shadow.
- **Extended palette.** Auto-generated tints and shades for charts and data visualization.

### Handling Broken Theme Colors

Some templates, particularly those exported from Google Slides, set all accent colors to white. The template looks fine in PowerPoint because the actual colors live in the master backgrounds, but the theme palette itself is technically useless.

pptx-masters detects this automatically. When the theme accents are all white or all identical, it falls back to colors extracted from the master slide backgrounds. The generated code and docs reflect the effective colors, not the broken theme values. Templates with working palettes are unaffected.

## CLI Reference

| Option | Description | Default |
|--------|-------------|---------|
| `<input>` | Path to `.pptx` or `.potx` file | Required |
| `-o, --output <dir>` | Output directory | `./output` |
| `--list` | List layout names and exit | |
| `--layout <name>` | Layout name or number (repeatable) | All |
| `--no-preview` | Skip `preview.pptx` generation | |
| `--no-report` | Skip `report.md` generation | |
| `-v, --verbose` | Verbose logging | |

## Limitations

Some PowerPoint features don't have clean equivalents in PptxGenJS. These are noted in the generated `report.md`:

- **Gradient fills** fall back to the dominant color
- **Pattern fills** fall back to the foreground color
- **Grouped shapes** are logged but skipped
- **Animations and transitions** are not supported by PptxGenJS
- **SmartArt, 3D effects, text warp, OLE objects** are not supported

For most corporate templates, these are edge cases. The core elements (colors, fonts, layouts, placeholders, backgrounds, slide numbers) are fully covered.

## How It Works

A `.pptx` file is a ZIP archive containing XML. pptx-masters:

1. Extracts the archive and parses the theme XML (colors, fonts, format scheme)
2. Builds a color resolution engine that follows the full OOXML chain: scheme color references through color maps to final hex values, including tint/shade math
3. Parses slide masters and layouts for backgrounds, placeholders, static shapes, and positions
4. Maps everything to [PptxGenJS](https://github.com/gitbrent/PptxGenJS) `defineSlideMaster()` objects
5. Generates JavaScript code, theme data, agent documentation, and a preview deck

Built on [PptxGenJS](https://github.com/gitbrent/PptxGenJS) for PowerPoint generation, [JSZip](https://stuk.github.io/jszip/) for archive handling, and [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) for XML parsing. Requires Node.js 18+.

## Roadmap

- **Embed theme colors in the PPTX.** The generated presentations use the correct colors in the slides, but PowerPoint's color picker dropdowns still show default theme colors instead of your template's palette. The fix is writing the extracted theme into the PPTX's `theme1.xml` so colors show up natively in the picker.
- **Windows testing.** Currently developed and tested on macOS. Node.js 18+ is the only hard requirement, so it should work on Windows, but this hasn't been verified yet. On Windows machines with PowerPoint installed, slide-to-image verification could use PowerPoint's COM interface directly instead of requiring LibreOffice.
- **Direct PowerPoint integration.** Most corporate environments already have PowerPoint installed and approved by IT. An adapter that talks to PowerPoint via [COM](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint) (Windows) or [Office.js](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/powerpoint-add-ins-reference-overview) (cross-platform, including PowerPoint Online) could use the extracted theme and master data to drive PowerPoint directly, without generating files through PptxGenJS at all.

## License

MIT
