/**
 * Brand skill generator — produces SKILL.md for Claude Code agent skills.
 *
 * Generates a complete, ready-to-use skill file that teaches LLM agents
 * how to create on-brand presentations using the extracted slide masters.
 */

import { toUpperSnakeCase, detectLimitedPalette, extractBackgroundColors } from './code.js';

/**
 * Round to at most 2 decimal places for display.
 *
 * @param {number} n
 * @returns {number}
 */
function r2(n) {
  if (typeof n !== 'number' || !isFinite(n)) return n;
  return Math.round(n * 100) / 100;
}

/**
 * Detect if a hex color is "dark" (relative luminance < 0.5).
 *
 * @param {string|undefined} hex - 6-char hex color
 * @returns {boolean}
 */
function isDark(hex) {
  if (!hex || hex.length < 6) return false;
  const clean = hex.replace('#', '');
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return (r * 0.299 + g * 0.587 + b * 0.114) < 128;
}

/**
 * Classify a master's role based on its name and properties.
 *
 * @param {object} master - Master data object
 * @param {number} index - Position in layout list
 * @param {number} total - Total number of layouts
 * @returns {{ role: string, desc: string }}
 */
function classifyMaster(master, index, total) {
  const name = (master.name || '').toLowerCase();
  const objs = master.objects || [];
  const hasTitle = objs.some((o) => o.placeholder?.options?.type === 'title');
  const hasBody = objs.some((o) => o.placeholder?.options?.type === 'body');
  const hasSub = objs.some((o) => {
    const n = (o.placeholder?.options?.name || '').toLowerCase();
    const t = o.placeholder?.options?.type || '';
    return n.includes('subtitle') || n.includes('sub') || t === 'subTitle';
  });
  const hasSlideNum = !!master.slideNumber;
  const darkBg = isDark(master.background?.color);

  // Cover/title slide: first layout, or name suggests it, typically no slide number
  if (name.includes('title slide') || name.includes('cover')) {
    return { role: 'cover', desc: 'Cover slide' };
  }
  // End/closing slide
  if (name.includes('end') || name.includes('closing') || name.includes('thank')) {
    return { role: 'end', desc: 'Closing slide' };
  }
  // Heuristic: last slide with dark bg and no slide number
  if (index === total - 1 && darkBg && !hasSlideNum) {
    return { role: 'end', desc: 'Closing slide' };
  }
  // Divider/section: dark/colored bg with only title
  if (name.includes('divider') || name.includes('section')) {
    return { role: 'divider', desc: 'Section divider' };
  }
  if (darkBg && hasTitle && !hasBody) {
    return { role: 'divider', desc: 'Section divider' };
  }
  // Content with subtitle
  if (hasTitle && hasBody && hasSub) {
    return { role: 'content-subtitled', desc: 'Content with subtitle' };
  }
  // Standard content
  if (hasTitle && hasBody) {
    return { role: 'content', desc: 'Standard content' };
  }
  // Title only (no body)
  if (hasTitle && !hasBody) {
    return { role: 'title-only', desc: 'Title only' };
  }
  // Blank or other
  if (!hasTitle && !hasBody) {
    return { role: 'blank', desc: 'Blank / decorative' };
  }
  return { role: 'other', desc: master.name };
}

/**
 * Compute content positioning areas from master placeholder data.
 *
 * @param {Array<object>} masterData
 * @param {{ width: number, height: number }} dimensions
 * @returns {Array<{ key: string, x: number, y: number, w: number, h: number, notes: string }>}
 */
function computePositions(masterData, dimensions) {
  const w = r2(dimensions?.width || 13.3333);
  const h = r2(dimensions?.height || 7.5);
  const entries = [];

  entries.push({ key: 'full', x: 0, y: 0, w, h, notes: 'Full slide' });

  // Find the primary "content" master (title + body, not cover/end)
  const classified = masterData.map((m, i) => ({
    master: m,
    cls: classifyMaster(m, i, masterData.length),
  }));

  const contentEntry = classified.find((c) => c.cls.role === 'content');
  if (contentEntry) {
    const objs = contentEntry.master.objects || [];
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    const bodyPh = objs.find((o) => o.placeholder?.options?.type === 'body');

    if (titlePh) {
      const opts = titlePh.placeholder.options;
      entries.push({
        key: 'title',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Title area',
      });
    }
    if (bodyPh) {
      const opts = bodyPh.placeholder.options;
      entries.push({
        key: 'body',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Content area',
      });
    }
  }

  // Subtitled content: extract subtitle + adjusted body
  const subtitledEntry = classified.find((c) => c.cls.role === 'content-subtitled');
  if (subtitledEntry) {
    const objs = subtitledEntry.master.objects || [];
    const subPh = objs.find((o) => {
      const n = (o.placeholder?.options?.name || '').toLowerCase();
      const t = o.placeholder?.options?.type || '';
      return n.includes('subtitle') || n.includes('sub') || t === 'subTitle';
    });
    const bodyPh = objs.find((o) => o.placeholder?.options?.type === 'body');

    if (subPh) {
      const opts = subPh.placeholder.options;
      entries.push({
        key: 'subtitle',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Subtitle (subtitled layouts)',
      });
    }
    if (bodyPh) {
      const opts = bodyPh.placeholder.options;
      entries.push({
        key: 'bodySubtitled',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Content area (subtitled layouts)',
      });
    }
  }

  // Cover slide positions
  const coverEntry = classified.find((c) => c.cls.role === 'cover');
  if (coverEntry) {
    const objs = coverEntry.master.objects || [];
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    const subPh = objs.find((o) => {
      const t = o.placeholder?.options?.type || '';
      return t === 'body' || t === 'subTitle';
    });

    if (titlePh) {
      const opts = titlePh.placeholder.options;
      entries.push({
        key: 'coverTitle',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Cover slide title',
      });
    }
    if (subPh) {
      const opts = subPh.placeholder.options;
      entries.push({
        key: 'coverSubtitle',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Cover slide subtitle',
      });
    }
  }

  // Divider text position
  const dividerEntry = classified.find((c) => c.cls.role === 'divider');
  if (dividerEntry) {
    const objs = dividerEntry.master.objects || [];
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    if (titlePh) {
      const opts = titlePh.placeholder.options;
      entries.push({
        key: 'dividerText',
        x: r2(opts.x),
        y: r2(opts.y),
        w: r2(opts.w),
        h: r2(opts.h),
        notes: 'Divider section title',
      });
    }
  }

  return entries;
}

/**
 * Extract typography info from placeholder styles across masters.
 *
 * @param {Array<object>} masterData
 * @param {{ heading: string, body: string }} themeFonts
 * @returns {Array<{ element: string, font: string, size: string, weight: string }>}
 */
function extractTypography(masterData, themeFonts, dimensions) {
  const seen = new Map(); // element → { fontFace, fontSize, bold }

  for (const master of masterData) {
    for (const obj of master.objects || []) {
      if (obj.placeholder?.options) {
        const opts = obj.placeholder.options;
        const type = opts.type;

        // Map placeholder types to typography elements
        let element;
        const name = (opts.name || '').toLowerCase();
        if (type === 'title' && name.includes('subtitle')) {
          element = 'Subtitle';
        } else if (type === 'title') {
          element = 'Slide title';
        } else if (type === 'body') {
          element = name.includes('subtitle') ? 'Subtitle' : 'Body text';
        } else {
          continue;
        }

        // Only take the first occurrence of each element type
        if (seen.has(element)) continue;
        if (opts.fontSize != null) {
          seen.set(element, {
            fontFace: opts.fontFace || themeFonts?.heading || 'Calibri',
            fontSize: opts.fontSize,
            bold: opts.bold || false,
          });
        }
      }

      // Footer text
      if (obj.text?.options && !seen.has('Footer')) {
        const opts = obj.text.options;
        const y = opts.y || 0;
        const slideH = dimensions?.height || 7.5;
        if (y > slideH * 0.85 && opts.fontSize != null) {
          seen.set('Footer', {
            fontFace: opts.fontFace || themeFonts?.body || 'Calibri',
            fontSize: opts.fontSize,
            bold: false,
          });
        }
      }
    }
  }

  const rows = [];
  for (const [element, props] of seen) {
    rows.push({
      element,
      font: props.fontFace,
      size: `${props.fontSize}pt`,
      weight: props.bold ? 'Bold' : 'Regular',
    });
  }

  return rows;
}

/**
 * Detect the footer text color from master data.
 *
 * @param {Array<object>} masterData
 * @returns {string|null} Hex color or null
 */
function detectFooterColor(masterData, dimensions) {
  const footerThreshold = (dimensions?.height || 7.5) * 0.8;
  for (const master of masterData) {
    for (const obj of master.objects || []) {
      if (obj.text?.options) {
        const opts = obj.text.options;
        if ((opts.y || 0) > footerThreshold && opts.color) {
          return opts.color;
        }
      }
    }
  }
  return null;
}

/**
 * Generate a brand skill SKILL.md file.
 *
 * @param {Array<object>} masterData
 * @param {object} options
 * @param {Record<string, string>} options.themeColors
 * @param {{ heading: string, body: string }} options.themeFonts
 * @param {{ width: number, height: number }} options.dimensions
 * @param {string} options.templateName
 * @returns {string} Markdown content
 */
export function generateBrandSkill(masterData, options) {
  const { themeColors, themeFonts, dimensions, templateName } = options;

  // Detect limited palette and extract fallback colors
  const paletteInfo = detectLimitedPalette(themeColors);
  const bgColors = extractBackgroundColors(masterData);
  const usesFallback = paletteInfo.isLimited && bgColors.length > 0;

  // Build effective theme colors
  let effectiveThemeColors = themeColors;
  if (usesFallback) {
    effectiveThemeColors = { ...themeColors };
    const accentSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    for (let i = 0; i < accentSlots.length && i < bgColors.length; i++) {
      effectiveThemeColors[accentSlots[i]] = bgColors[i];
    }
  }

  const lines = [];

  const masterTitles = masterData.map((m) => toUpperSnakeCase(m.name));
  const masterList = masterTitles.join(', ');
  const headingFont = themeFonts?.heading || 'Calibri';
  const bodyFont = themeFonts?.body || 'Calibri';
  const font = headingFont === bodyFont ? headingFont : headingFont;
  const w = r2(dimensions?.width || 13.3333);
  const h = r2(dimensions?.height || 7.5);

  // Classify masters
  const classified = masterData.map((m, i) => ({
    master: m,
    title: toUpperSnakeCase(m.name),
    cls: classifyMaster(m, i, masterData.length),
  }));

  // --- YAML Header ---
  lines.push('---');
  lines.push('name: brand-slides');
  lines.push(`description: "Brand guidelines for presentations. Load this skill when creating .pptx files. Masters: ${masterList}. Trigger on: presentation, slides, deck, .pptx, PptxGenJS."`);
  lines.push('---');
  lines.push('');

  // --- Title ---
  lines.push('# Brand Slide Masters');
  lines.push('');

  // --- The Rule ---
  lines.push('## The Rule');
  lines.push('');
  lines.push(`**Use \`masters.js\` for all new presentations.** It defines ${masterData.length} PptxGenJS slide masters (${masterList}) that reproduce the brand visual system. Footers, logos, slide numbers, and decorative elements are baked into the masters.`);
  lines.push('');
  lines.push('---');
  lines.push('');

  // --- Workflow ---
  lines.push('## Workflow');
  lines.push('');
  lines.push(`**Design first, export last.** Create HTML previews (${w}" \u00D7 ${h}" dimensions, brand colors/fonts) for rapid iteration. Only generate PPTX once the visual design is approved.`);
  lines.push('');
  lines.push('**HTML to PPTX conversion:** positions `px / 96` → inches, font sizes `px × 0.85` → pt (rounded to 0.5pt). Leave 10-15% extra width in text boxes. For custom SVG charts, screenshot as PNG and embed with `addImage()` — do not build complex charts from shape primitives (line shapes with negative dimensions corrupt the file).');
  lines.push('');
  lines.push('```javascript');
  if (usesFallback) {
    lines.push("import { createPresentation, THEME, PALETTE, CHART_COLORS, FONT, POS, SLIDE_COLORS } from './masters.js';");
  } else {
    lines.push("import { createPresentation, THEME, PALETTE, CHART_COLORS, FONT, POS } from './masters.js';");
  }
  lines.push('');
  lines.push("const pres = await createPresentation('Deck Title');");
  lines.push('');

  // Generate one example per master
  for (let i = 0; i < classified.length; i++) {
    const { master, title, cls } = classified[i];
    const objs = master.objects || [];
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    const bodyPh = objs.find((o) => o.placeholder?.options?.type === 'body');
    const subPh = objs.find((o) => {
      const n = (o.placeholder?.options?.name || '').toLowerCase();
      return n.includes('subtitle') || n.includes('sub');
    });

    lines.push(`// ${cls.desc}`);
    if (i === 0) {
      lines.push(`let slide = pres.addSlide({ masterName: '${title}' });`);
    } else {
      lines.push(`slide = pres.addSlide({ masterName: '${title}' });`);
    }

    if (titlePh) {
      const phName = titlePh.placeholder.options.name || 'title';
      const text = cls.role === 'cover'
        ? 'Deck Title'
        : cls.role === 'divider'
          ? 'Section Title'
          : 'Slide Title';
      lines.push(`slide.addText('${text}', { placeholder: '${phName}' });`);
    }
    if (subPh) {
      const phName = subPh.placeholder.options.name || 'subtitle';
      lines.push(`slide.addText('Subtitle text', { placeholder: '${phName}' });`);
    }
    if (bodyPh && cls.role !== 'cover' && cls.role !== 'end') {
      if (cls.role === 'content') {
        lines.push('slide.addChart(pres.charts.BAR, data, { ...POS.body, chartColors: CHART_COLORS });');
      } else {
        const phName = bodyPh.placeholder.options.name || 'body';
        lines.push(`slide.addText('Content goes here', { placeholder: '${phName}' });`);
      }
    }

    // Show background override example on the second content slide
    if (cls.role === 'content' && i > 0) {
      lines.push('// Background override: slide.background = { color: THEME.accent1 };');
    }

    lines.push('');
  }

  lines.push("await pres.writeFile({ fileName: 'output.pptx' });");
  lines.push('```');
  lines.push('');
  lines.push('Run with: `node script.mjs`');
  lines.push('');
  lines.push('---');
  lines.push('');

  // --- Masters Reference ---
  lines.push('## Masters Reference');
  lines.push('');
  lines.push('| Master | Background | Footer | Slide # | Use |');
  lines.push('|--------|-----------|--------|---------|-----|');

  for (const { master, title, cls } of classified) {
    const bgDesc = master.background?.color
      ? `#${master.background.color}`
      : master.background?.path
        ? 'Image'
        : 'None';
    const hasFooter = (master.objects || []).some(
      (o) => o.text && (o.text.options?.y || 0) > (h * 0.85),
    );
    const footerDesc = hasFooter ? 'Yes' : 'No';
    const slideNumDesc = master.slideNumber ? 'Yes' : 'No';

    lines.push(`| \`${title}\` | ${bgDesc} | ${footerDesc} | ${slideNumDesc} | ${cls.desc} |`);
  }

  lines.push('');
  lines.push('---');
  lines.push('');

  // --- Positioning ---
  const positions = computePositions(masterData, dimensions);
  if (positions.length > 0) {
    lines.push('## Positioning (`POS`)');
    lines.push('');
    lines.push('Use these when placing content on slides:');
    lines.push('');
    lines.push('| Key | x | y | w | h | Notes |');
    lines.push('|-----|---|---|---|---|-------|');

    for (const pos of positions) {
      lines.push(`| \`${pos.key}\` | ${pos.x} | ${pos.y} | ${pos.w} | ${pos.h} | ${pos.notes} |`);
    }

    lines.push('');
    lines.push('---');
    lines.push('');
  }

  // --- Color Reference ---
  lines.push('## Color Reference');
  lines.push('');
  if (usesFallback) {
    lines.push('> **Note:** Theme accents are limited (all near-white/identical). Colors below are derived from master slide backgrounds.');
    lines.push('');
  }

  // Theme colors with semantic roles
  lines.push('### Theme Colors (`THEME.*`)');
  lines.push('');
  lines.push('| Key | Hex | Role |');
  lines.push('|-----|-----|------|');

  if (themeColors) {
    const semanticColors = [
      { key: 'text', hex: themeColors.dk1, role: 'Default text' },
      { key: 'background', hex: themeColors.lt1, role: 'Default background' },
      { key: 'brand', hex: effectiveThemeColors.accent1, role: usesFallback ? 'From master backgrounds' : 'Primary brand accent' },
    ];

    // Detect footer color
    const footerColor = detectFooterColor(masterData, dimensions);
    if (footerColor) {
      semanticColors.push({ key: 'footer', hex: footerColor, role: 'Footer text' });
    }

    // Add all accents (use effective colors)
    for (let i = 2; i <= 6; i++) {
      const slot = `accent${i}`;
      const hex = effectiveThemeColors[slot];
      if (hex) {
        const source = usesFallback && hex !== themeColors[slot] ? 'From master backgrounds' : `Accent ${i}`;
        semanticColors.push({ key: slot, hex, role: source });
      }
    }

    for (const { key, hex, role } of semanticColors) {
      if (hex) {
        lines.push(`| \`${key}\` | ${hex} | ${role} |`);
      }
    }
  }

  lines.push('');

  // Chart colors
  lines.push('### Chart Colors (`CHART_COLORS`)');
  lines.push('');
  lines.push('Use in order for data visualization:');
  lines.push('');
  lines.push('| # | Hex | Source |');
  lines.push('|---|-----|--------|');

  if (usesFallback) {
    let num = 1;
    for (const color of bgColors) {
      lines.push(`| ${num} | ${color} | master background |`);
      num++;
    }
  } else if (themeColors) {
    const chartSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    let num = 1;
    for (const slot of chartSlots) {
      if (themeColors[slot]) {
        lines.push(`| ${num} | ${themeColors[slot]} | ${slot} |`);
        num++;
      }
    }
  }

  lines.push('');

  // Extended palette
  lines.push('### Extended Palette (`PALETTE.*`)');
  lines.push('');
  lines.push('Access tints/shades: `PALETTE.accent1.lighter60`, `PALETTE.dk2.darker25`, etc.');
  lines.push('');
  lines.push('Variants per color: `base`, `lighter80`, `lighter60`, `lighter40`, `lighter25`, `darker25`, `darker50`');
  lines.push('');
  lines.push('---');
  lines.push('');

  // --- Typography ---
  const typography = extractTypography(masterData, themeFonts, dimensions);
  if (typography.length > 0) {
    lines.push('## Typography');
    lines.push('');
    lines.push('| Element | Font | Size | Weight |');
    lines.push('|---------|------|------|--------|');

    for (const row of typography) {
      lines.push(`| ${row.element} | ${row.font} | ${row.size} | ${row.weight} |`);
    }

    lines.push('');
    if (headingFont !== bodyFont) {
      lines.push(`Fallback fonts: ${headingFont}, ${bodyFont}, Arial.`);
    } else {
      lines.push(`Fallback fonts: ${headingFont}, Arial.`);
    }
    lines.push('');
    lines.push('---');
    lines.push('');
  }

  // --- Key Conventions ---
  lines.push('## Key Conventions');
  lines.push('');
  lines.push(`- **Always specify \`fontFace: FONT\`** on every text element. PptxGenJS masters don't cascade fonts like OOXML themes.`);
  lines.push('- **Footers are automatic.** Masters handle copyright, slide numbers, and decorative elements.');
  lines.push("- **Use placeholder names** for titles and body: `slide.addText('...', { placeholder: 'title' })`.");
  lines.push(`- **Background overrides:** Set \`slide.background = { color: THEME.accent1 }\` on any slide.`);
  lines.push(`- **Slide dimensions:** ${w}" \u00D7 ${h}". POS values are pre-calculated for this size.`);
  lines.push(`- **Body text color:** \`THEME.text\` (${themeColors?.dk1 || '000000'}).`);
  if (usesFallback) {
    lines.push('- **Dark backgrounds:** Use `THEME.background` (white) for text on dark master backgrounds.');
  }
  lines.push('');

  return lines.join('\n');
}

export { classifyMaster, computePositions, extractTypography, detectFooterColor };
