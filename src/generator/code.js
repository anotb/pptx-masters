/**
 * Code generator — produces ES module JavaScript code and theme JSON
 * from mapped PptxGenJS master data.
 *
 * Also generates agent instructions (SLIDE_MASTERS.md) for teaching
 * LLM agents how to use the extracted masters.
 */

import { applyColorModifiers } from '../parser/colors.js';

/**
 * Convert a layout name to UPPER_SNAKE_CASE.
 * e.g., "Title Slide" → "TITLE_SLIDE"
 *
 * @param {string} name - Layout name
 * @returns {string}
 */
export function toUpperSnakeCase(name) {
  return name
    .replace(/[_-]/g, ' ')
    .replace(/[^a-zA-Z0-9\s]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .toUpperCase();
}

/**
 * Round a number to at most 4 decimal places.
 *
 * @param {number} n
 * @returns {number}
 */
function round4(n) {
  if (typeof n !== 'number' || !isFinite(n)) return n;
  return Math.round(n * 10000) / 10000;
}

/**
 * Deep-clean an object: remove undefined/null values and round numbers.
 *
 * @param {*} val
 * @returns {*}
 */
function cleanValue(val) {
  if (val === undefined || val === null) return undefined;
  if (typeof val === 'number') return round4(val);
  if (Array.isArray(val)) {
    return val.map(cleanValue).filter((v) => v !== undefined);
  }
  if (typeof val === 'object') {
    const cleaned = {};
    for (const [k, v] of Object.entries(val)) {
      const cv = cleanValue(v);
      if (cv !== undefined) {
        cleaned[k] = cv;
      }
    }
    return Object.keys(cleaned).length > 0 ? cleaned : undefined;
  }
  return val;
}

/**
 * Stringify a value with nice formatting, indented to a given depth.
 *
 * @param {*} val
 * @param {number} indent - Base indentation level (in spaces)
 * @returns {string}
 */
function prettyStringify(val, indent = 4) {
  const cleaned = cleanValue(val);
  if (cleaned === undefined) return 'undefined';
  const json = JSON.stringify(cleaned, null, 2);
  // Re-indent each line to the base indentation
  const pad = ' '.repeat(indent);
  return json
    .split('\n')
    .map((line, i) => (i === 0 ? line : pad + line))
    .join('\n');
}

/**
 * Round to 2 decimal places for display.
 *
 * @param {number} n
 * @returns {number}
 */
function round2(n) {
  if (typeof n !== 'number' || !isFinite(n)) return n;
  return Math.round(n * 100) / 100;
}

/**
 * Detect if a hex color is "dark" based on relative luminance.
 *
 * @param {string|undefined} hex
 * @returns {boolean}
 */
function isDarkColor(hex) {
  if (!hex || hex.length < 6) return false;
  const clean = hex.replace('#', '');
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return (r * 0.299 + g * 0.587 + b * 0.114) < 128;
}

/**
 * Compute weighted brightness (0-255) for a hex color.
 *
 * @param {string} hex - 6-character hex color (with or without #)
 * @returns {number}
 */
function hexBrightness(hex) {
  if (!hex || typeof hex !== 'string') return 0;
  const clean = hex.replace('#', '');
  if (clean.length < 6) return 0;
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return r * 0.299 + g * 0.587 + b * 0.114;
}

/**
 * Detect whether a theme's accent colors are limited — all near-white,
 * near-black, or identical, making them useless for data visualization.
 *
 * @param {Record<string, string>|null} themeColors
 * @returns {{ isLimited: boolean, usableAccents: string[] }}
 */
export function detectLimitedPalette(themeColors) {
  if (!themeColors) return { isLimited: false, usableAccents: [] };

  const accentSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  const accents = accentSlots.map((s) => themeColors[s]).filter(Boolean);

  // Filter out neutrals: brightness > 245 (near-white) or < 10 (near-black)
  const nonNeutral = accents.filter((hex) => {
    const b = hexBrightness(hex);
    return b <= 245 && b >= 10;
  });

  // Unique non-neutral accents
  const unique = [...new Set(nonNeutral)];

  return {
    isLimited: unique.length < 2,
    usableAccents: unique,
  };
}

/**
 * Extract non-neutral background colors from master slide data,
 * sorted by frequency (descending) then darkness (descending).
 *
 * @param {Array<object>|null} masterData
 * @returns {string[]}
 */
export function extractBackgroundColors(masterData) {
  if (!masterData || !Array.isArray(masterData) || masterData.length === 0) return [];

  // Collect all solid background colors
  const freq = {};
  for (const master of masterData) {
    const color = master.background?.color;
    if (!color) continue; // skip image backgrounds and missing
    // Filter neutrals: brightness > 245 (near-white) or < 10 (near-black)
    const b = hexBrightness(color);
    if (b > 245 || b < 10) continue;
    freq[color] = (freq[color] || 0) + 1;
  }

  const colors = Object.keys(freq);
  if (colors.length === 0) return [];

  // Sort by frequency desc, then darkness desc (lower brightness = darker)
  colors.sort((a, b) => {
    const freqDiff = freq[b] - freq[a];
    if (freqDiff !== 0) return freqDiff;
    return hexBrightness(a) - hexBrightness(b); // darker first
  });

  return colors;
}

/**
 * Detect footer text color from master data by scanning for
 * text objects in the bottom 15% of the slide.
 *
 * @param {Array<object>} masterData
 * @returns {string|null}
 */
function detectMasterFooterColor(masterData, dimensions) {
  const footerThreshold = (dimensions?.height || 7.5) * 0.8;
  // Prefer footer color from light-background masters (more common, more useful)
  const lightBgMasters = masterData.filter((m) => !isDarkColor(m.background?.color));
  const darkBgMasters = masterData.filter((m) => isDarkColor(m.background?.color));

  for (const masters of [lightBgMasters, darkBgMasters]) {
    for (const master of masters) {
      for (const obj of master.objects || []) {
        if (obj.text?.options) {
          const opts = obj.text.options;
          if ((opts.y || 0) > footerThreshold && opts.color) {
            return opts.color;
          }
        }
      }
    }
  }
  return null;
}

/**
 * Classify a master's role based on its name and properties.
 *
 * @param {object} master
 * @param {number} index
 * @param {number} total
 * @returns {{ role: string, desc: string }}
 */
function classifyMasterRole(master, index, total) {
  const name = (master.name || '').toLowerCase();
  const objs = master.objects || [];
  const hasTitle = objs.some((o) => o.placeholder?.options?.type === 'title');
  const bodyPhs = objs.filter((o) => o.placeholder?.options?.type === 'body');
  const hasBody = bodyPhs.length > 0;
  const hasSub = objs.some((o) => {
    const n = (o.placeholder?.options?.name || '').toLowerCase();
    const t = o.placeholder?.options?.type || '';
    return n.includes('subtitle') || n.includes('sub') || t === 'subTitle';
  });
  // Also detect subtitled layouts by name or by having multiple body placeholders
  const nameHasSubtitle = name.includes('subtitle');
  const hasMultipleBodies = bodyPhs.length >= 2;
  const isSubtitled = hasSub || nameHasSubtitle || hasMultipleBodies;
  const hasSlideNum = !!master.slideNumber;
  const darkBg = isDarkColor(master.background?.color);

  if (name.includes('title slide') || name.includes('cover')) return { role: 'cover', desc: 'Cover slide' };
  if (name.includes('end') || name.includes('closing') || name.includes('thank')) return { role: 'end', desc: 'Closing slide' };
  if (index === total - 1 && darkBg && !hasSlideNum) return { role: 'end', desc: 'Closing slide' };
  if (name.includes('divider') || name.includes('section')) return { role: 'divider', desc: 'Section divider' };
  if (darkBg && hasTitle && !hasBody) return { role: 'divider', desc: 'Section divider' };
  if (name.includes('team') || name.includes('profile')) return { role: 'content-subtitled', desc: 'Team/profile layout' };
  if (name.includes('qualif')) return { role: 'content-subtitled', desc: 'Qualifications layout' };
  if (hasTitle && hasBody && isSubtitled) return { role: 'content-subtitled', desc: 'Content with subtitle' };
  if (hasTitle && hasBody) return { role: 'content', desc: 'Standard content' };
  if (hasTitle && !hasBody) return { role: 'title-only', desc: 'Title only' };
  return { role: 'other', desc: master.name };
}

/**
 * Compute content positioning areas from master placeholder data.
 *
 * @param {Array<object>} masterData
 * @param {{ width: number, height: number }} dimensions
 * @returns {Array<{ key: string, x: number, y: number, w: number, h: number, notes: string }>}
 */
function computeMasterPositions(masterData, dimensions) {
  const w = round2(dimensions?.width || 13.3333);
  const h = round2(dimensions?.height || 7.5);
  const entries = [];

  entries.push({ key: 'full', x: 0, y: 0, w, h, notes: 'Full slide' });

  const classified = masterData.map((m, i) => ({
    master: m,
    cls: classifyMasterRole(m, i, masterData.length),
  }));

  // Collect all body placeholders from content masters for POS.body
  const contentRoles = new Set(['content', 'content-subtitled', 'title-only']);
  const contentMasters = classified.filter((c) => contentRoles.has(c.cls.role));

  // Title: from any content master with a title placeholder
  const titleSource = contentMasters.find((c) =>
    (c.master.objects || []).some((o) => o.placeholder?.options?.type === 'title'),
  );
  if (titleSource) {
    const titlePh = titleSource.master.objects.find((o) => o.placeholder?.options?.type === 'title');
    const opts = titlePh.placeholder.options;
    entries.push({ key: 'title', x: round2(opts.x), y: round2(opts.y), w: round2(opts.w), h: round2(opts.h), notes: 'Title area' });
  }

  // Body: find the largest body placeholder across all content masters (by area)
  // This ensures we get the main content area, not a small subtitle placeholder
  let bestBody = null;
  let bestBodyArea = 0;
  for (const c of contentMasters) {
    for (const o of c.master.objects || []) {
      if (o.placeholder?.options?.type !== 'body') continue;
      const opts = o.placeholder.options;
      const area = (opts.w || 0) * (opts.h || 0);
      if (area > bestBodyArea) {
        bestBodyArea = area;
        bestBody = opts;
      }
    }
  }
  if (bestBody) {
    entries.push({ key: 'body', x: round2(bestBody.x), y: round2(bestBody.y), w: round2(bestBody.w), h: round2(bestBody.h), notes: 'Content area below title+subtitle' });
  }

  // bodyFull: for title-only layouts — starts right after title, more vertical space
  const titleOnlyEntry = classified.find((c) => c.cls.role === 'title-only');
  if (titleOnlyEntry) {
    const objs = titleOnlyEntry.master.objects || [];
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    if (titlePh) {
      const titleOpts = titlePh.placeholder.options;
      const startY = round2((titleOpts.y || 0) + (titleOpts.h || 0) + 0.12); // 0.12" gap after title
      const bodyX = bestBody ? round2(bestBody.x) : 0.5;
      const bodyW = bestBody ? round2(bestBody.w) : round2(w - 1);
      const footerY = round2(h - 0.55); // reserve ~0.55" for footer zone
      const fullH = round2(footerY - startY);
      entries.push({ key: 'bodyFull', x: bodyX, y: startY, w: bodyW, h: fullH, notes: 'Full content area (title-only, no subtitle)' });
    }
  }

  // Subtitle: small body placeholder near title in subtitled layouts
  const subtitledEntries = classified.filter((c) => c.cls.role === 'content-subtitled');
  if (subtitledEntries.length > 0) {
    // Find subtitle: body placeholder closest to title with small height
    for (const entry of subtitledEntries) {
      const objs = entry.master.objects || [];
      const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
      const bodyPhs = objs.filter((o) => o.placeholder?.options?.type === 'body');
      if (titlePh && bodyPhs.length >= 1) {
        const titleBottom = (titlePh.placeholder.options.y || 0) + (titlePh.placeholder.options.h || 0);
        // Subtitle = body placeholder closest to title bottom, with height < 1.5"
        const subtitle = bodyPhs
          .filter((o) => (o.placeholder.options.h || 0) < 1.5)
          .sort((a, b) => Math.abs((a.placeholder.options.y || 0) - titleBottom) - Math.abs((b.placeholder.options.y || 0) - titleBottom))[0];
        if (subtitle) {
          const opts = subtitle.placeholder.options;
          entries.push({ key: 'subtitle', x: round2(opts.x), y: round2(opts.y), w: round2(opts.w), h: round2(opts.h), notes: 'Subtitle area' });
          break;
        }
      }
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
      entries.push({ key: 'coverTitle', x: round2(opts.x), y: round2(opts.y), w: round2(opts.w), h: round2(opts.h), notes: 'Cover slide title' });
    }
    if (subPh) {
      const opts = subPh.placeholder.options;
      entries.push({ key: 'coverSubtitle', x: round2(opts.x), y: round2(opts.y), w: round2(opts.w), h: round2(opts.h), notes: 'Cover slide subtitle' });
    }
  }

  // Divider text position
  const dividerEntry = classified.find((c) => c.cls.role === 'divider');
  if (dividerEntry) {
    const objs = dividerEntry.master.objects || [];
    const bodyPh = objs.find((o) => o.placeholder?.options?.type === 'body');
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    const ph = bodyPh || titlePh;
    if (ph) {
      const opts = ph.placeholder.options;
      entries.push({ key: 'dividerText', x: round2(opts.x), y: round2(opts.y), w: round2(opts.w), h: round2(opts.h), notes: 'Divider section title' });
    }
  }

  return entries;
}

/**
 * Generate a valid ES module JavaScript file containing PptxGenJS
 * defineSlideMaster() calls.
 *
 * @param {Array<object>} masterData - Array of layout objects with mapped PptxGenJS data
 * @param {object} options - Generation options
 * @param {string} options.templateName - Source template filename
 * @param {{ width: number, height: number }} options.dimensions - Slide dimensions
 * @param {Record<string, string>} options.themeColors - Theme color map
 * @param {{ heading: string, body: string }} options.themeFonts - Theme fonts
 * @returns {string} Generated JavaScript source code
 */
export function generateMastersCode(masterData, options) {
  const { templateName, themeColors, themeFonts, dimensions } = options;
  const date = new Date().toISOString().slice(0, 10);

  // Detect limited palette and extract fallback colors
  const { isLimited } = detectLimitedPalette(themeColors);
  const bgColors = extractBackgroundColors(masterData);
  const usesFallback = isLimited && bgColors.length > 0;

  // Build effective theme colors by patching bgColors into accent slots
  let effectiveThemeColors = themeColors;
  if (usesFallback && themeColors) {
    effectiveThemeColors = { ...themeColors };
    const accentSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    for (let i = 0; i < accentSlots.length; i++) {
      if (i < bgColors.length) {
        effectiveThemeColors[accentSlots[i]] = bgColors[i];
      }
    }
  }

  const lines = [];

  // Header
  lines.push('// Generated by pptx-masters (https://github.com/anotb/pptx-masters)');
  lines.push(`// Template: ${templateName}`);
  lines.push(`// Date: ${date}`);
  lines.push('');

  // Theme colors (base scheme)
  lines.push('// Theme Colors');
  lines.push('const THEME_COLORS = {');
  if (themeColors) {
    const colorSlots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    for (const slot of colorSlots) {
      if (themeColors[slot] != null) {
        lines.push(`  ${slot}: '${themeColors[slot]}',`);
      }
    }
  }
  lines.push('};');
  lines.push('');

  // Extended palette (PowerPoint color picker tints/shades)
  if (themeColors) {
    const palette = generateExtendedPalette(usesFallback ? effectiveThemeColors : themeColors);
    lines.push('// Extended Palette (PowerPoint color picker tints/shades)');
    lines.push('const PALETTE = {');
    for (const [slot, entry] of Object.entries(palette)) {
      const vals = Object.entries(entry)
        .map(([k, v]) => `${k}: '${v}'`)
        .join(', ');
      lines.push(`  ${slot}: { ${vals} },`);
    }
    lines.push('};');
    lines.push('');
  }

  // Theme fonts
  lines.push('// Theme Fonts');
  lines.push(`const HEADING_FONT = '${themeFonts?.heading || 'Calibri'}';`);
  lines.push(`const BODY_FONT = '${themeFonts?.body || 'Calibri'}';`);
  lines.push('');

  // registerMasters function
  lines.push('/**');
  lines.push(' * Register all slide masters with a PptxGenJS instance.');
  lines.push(' * @param {PptxGenJS} pptx - PptxGenJS instance');
  lines.push(' */');
  lines.push('export function registerMasters(pptx) {');

  for (let i = 0; i < masterData.length; i++) {
    const master = masterData[i];
    const title = toUpperSnakeCase(master.name);

    if (i > 0) lines.push('');

    lines.push(`  // ${master.name}`);
    lines.push('  pptx.defineSlideMaster({');
    lines.push(`    title: '${title}',`);

    // Background
    if (master.background) {
      const bgCleaned = cleanValue(master.background);
      if (bgCleaned) {
        lines.push(`    background: ${prettyStringify(bgCleaned, 4)},`);
      }
    }

    // Slide number
    if (master.slideNumber) {
      const snCleaned = cleanValue(master.slideNumber);
      if (snCleaned) {
        lines.push(`    slideNumber: ${prettyStringify(snCleaned, 4)},`);
      }
    }

    // Objects
    if (master.objects && master.objects.length > 0) {
      lines.push('    objects: [');
      for (const obj of master.objects) {
        const objCleaned = cleanValue(obj);
        if (objCleaned) {
          lines.push(`      ${prettyStringify(objCleaned, 6)},`);
        }
      }
      lines.push('    ],');
    } else {
      lines.push('    objects: [],');
    }

    lines.push('  });');
  }

  lines.push('}');
  lines.push('');

  // Primary font
  lines.push('// Primary Font');
  lines.push('const FONT = HEADING_FONT;');
  lines.push('');

  // Semantic theme colors
  lines.push('// Semantic Theme Colors');
  if (usesFallback) {
    lines.push('// NOTE: Theme accents are limited (all near-white/identical). Using master background colors.');
  }
  lines.push('const THEME = {');
  if (themeColors) {
    if (themeColors.dk1) lines.push(`  text: THEME_COLORS.dk1, // ${themeColors.dk1}`);
    if (themeColors.lt1) lines.push(`  background: THEME_COLORS.lt1, // ${themeColors.lt1}`);
    if (usesFallback) {
      // Use literal hex values from bgColors for brand/accents
      lines.push(`  brand: '${bgColors[0]}', // from master backgrounds`);
      for (let i = 1; i < bgColors.length && i < 6; i++) {
        lines.push(`  accent${i + 1}: '${bgColors[i]}', // from master backgrounds`);
      }
    } else {
      if (themeColors.accent1) lines.push(`  brand: THEME_COLORS.accent1, // ${themeColors.accent1}`);
      for (let i = 2; i <= 6; i++) {
        const slot = `accent${i}`;
        if (themeColors[slot]) lines.push(`  ${slot}: THEME_COLORS.${slot}, // ${themeColors[slot]}`);
      }
    }
    // Detect footer color from masterData
    const footerColor = detectMasterFooterColor(masterData, dimensions);
    if (footerColor) {
      lines.push(`  footer: '${footerColor}',`);
    }
  }
  lines.push('};');
  lines.push('');

  // Content positioning areas
  const posEntries = computeMasterPositions(masterData, options.dimensions);
  lines.push('// Content Positioning Areas');
  lines.push('const POS = {');
  for (const pos of posEntries) {
    lines.push(`  ${pos.key}: { x: ${pos.x}, y: ${pos.y}, w: ${pos.w}, h: ${pos.h} }, // ${pos.notes}`);
  }
  lines.push('};');
  lines.push('');

  // Chart colors
  if (usesFallback) {
    lines.push('// Chart Colors (from master backgrounds — theme accents are limited)');
    lines.push('const CHART_COLORS = [');
    for (const color of bgColors) {
      lines.push(`  '${color}', // from master backgrounds`);
    }
  } else {
    lines.push('// Chart Colors (accent-based, ordered for data visualization)');
    lines.push('const CHART_COLORS = [');
    if (themeColors) {
      const chartSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
      for (const slot of chartSlots) {
        if (themeColors[slot]) lines.push(`  '${themeColors[slot]}', // ${slot}`);
      }
    }
  }
  lines.push('];');
  lines.push('');

  // createPresentation convenience function
  const slideW = round4(options.dimensions?.width || 13.3333);
  const slideH = round4(options.dimensions?.height || 7.5);
  lines.push('/**');
  lines.push(' * Create a new presentation with all masters pre-registered.');
  lines.push(' * @param {string} [title] - Presentation title');
  lines.push(' * @returns {Promise<PptxGenJS>} Configured PptxGenJS instance');
  lines.push(' */');
  lines.push('export async function createPresentation(title = \'\') {');
  lines.push("  const { default: PptxGenJS } = await import('pptxgenjs');");
  lines.push('  const pptx = new PptxGenJS();');
  lines.push(`  pptx.defineLayout({ name: 'LAYOUT_CUSTOM', width: ${slideW}, height: ${slideH} });`);
  lines.push("  pptx.layout = 'LAYOUT_CUSTOM';");
  lines.push('  if (title) pptx.title = title;');
  lines.push('  registerMasters(pptx);');
  lines.push('  return pptx;');
  lines.push('}');
  lines.push('');

  // SLIDE_COLORS (only when using fallback)
  if (usesFallback) {
    lines.push('// Slide Colors (from master backgrounds — for data visualization fallback)');
    lines.push('const SLIDE_COLORS = [');
    for (const color of bgColors) {
      lines.push(`  '${color}',`);
    }
    lines.push('];');
    lines.push('');
  }

  // Exports
  if (usesFallback) {
    lines.push('export { THEME_COLORS, PALETTE, HEADING_FONT, BODY_FONT, THEME, POS, CHART_COLORS, FONT, SLIDE_COLORS };');
  } else {
    lines.push('export { THEME_COLORS, PALETTE, HEADING_FONT, BODY_FONT, THEME, POS, CHART_COLORS, FONT };');
  }
  lines.push('');

  return lines.join('\n');
}

/**
 * Generate a theme.json object containing colors, fonts, and dimensions.
 *
 * @param {Record<string, string>} themeColors - Theme color map
 * @param {{ heading: string, body: string }} themeFonts - Theme fonts
 * @param {{ width: number, height: number }} dimensions - Slide dimensions
 * @param {Array<object>} [masterData] - Optional master data (for limited palette detection)
 * @returns {object} Theme JSON object
 */
export function generateThemeJson(themeColors, themeFonts, dimensions, masterData) {
  const { isLimited } = detectLimitedPalette(themeColors);
  const bgColors = masterData ? extractBackgroundColors(masterData) : [];
  const usesFallback = isLimited && bgColors.length > 0;

  // Build effective theme colors for palette generation
  let paletteColors = themeColors;
  if (usesFallback && themeColors) {
    paletteColors = { ...themeColors };
    const accentSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    for (let i = 0; i < accentSlots.length; i++) {
      if (i < bgColors.length) {
        paletteColors[accentSlots[i]] = bgColors[i];
      }
    }
  }

  const result = {
    colors: { ...themeColors },
    palette: generateExtendedPalette(paletteColors),
    fonts: {
      heading: themeFonts?.heading || 'Calibri',
      body: themeFonts?.body || 'Calibri',
    },
    dimensions: {
      width: dimensions?.width || 10,
      height: dimensions?.height || 7.5,
    },
  };

  if (usesFallback) {
    result.slideColors = bgColors;
    result.paletteSource = 'background-fallback';
  }

  return result;
}

/**
 * Pick the closest emoji for a hex color.
 *
 * @param {string} hex - 6-character hex color
 * @returns {string}
 */
/**
 * Generate the extended color palette matching PowerPoint's color picker.
 *
 * For each theme color, generates the standard 5 tint/shade variants:
 *   lighter80, lighter60, lighter40, lighter25, darker25, darker50
 *
 * @param {Record<string, string>} themeColors - Base theme color map
 * @returns {Record<string, { base: string, lighter80: string, lighter60: string, lighter40: string, lighter25: string, darker25: string, darker50: string }>}
 */
function generateExtendedPalette(themeColors) {
  if (!themeColors) return {};

  // PowerPoint's standard tint/shade modifiers (from the color picker grid)
  const variants = [
    { name: 'lighter80', mods: { lumMod: 20000, lumOff: 80000 } },
    { name: 'lighter60', mods: { lumMod: 40000, lumOff: 60000 } },
    { name: 'lighter40', mods: { lumMod: 60000, lumOff: 40000 } },
    { name: 'lighter25', mods: { lumMod: 75000, lumOff: 25000 } },
    { name: 'darker25', mods: { lumMod: 75000 } },
    { name: 'darker50', mods: { lumMod: 50000 } },
  ];

  const palette = {};
  const colorSlots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

  for (const slot of colorSlots) {
    const baseHex = themeColors[slot];
    if (!baseHex) continue;

    const entry = { base: baseHex };
    for (const { name, mods } of variants) {
      entry[name] = applyColorModifiers(baseHex, mods);
    }
    palette[slot] = entry;
  }

  return palette;
}

function colorEmoji(hex) {
  if (!hex || typeof hex !== 'string') return '';

  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const brightness = (r * 299 + g * 587 + b * 114) / 1000;

  // Very dark
  if (brightness < 50) return '\u2B1B';
  // Very light
  if (brightness > 220) return '\u2B1C';

  // Determine dominant color channel for accents
  const max = Math.max(r, g, b);
  if (max === r && r > g * 1.3 && r > b * 1.3) return '\uD83D\uDFE5'; // red
  if (max === g && g > r * 1.3 && g > b * 1.3) return '\uD83D\uDFE9'; // green
  if (max === b && b > r * 1.3 && b > g * 1.3) return '\uD83D\uDFE6'; // blue
  if (r > 200 && g > 150 && b < 100) return '\uD83D\uDFE7'; // orange
  if (r > 200 && g > 200 && b < 100) return '\uD83D\uDFE8'; // yellow
  if (r > 100 && b > 100 && g < 100) return '\uD83D\uDFEA'; // purple

  // Mid-range gray or mixed
  if (brightness < 128) return '\u2B1B';
  return '\u2B1C';
}

/**
 * Format a dimension aspect ratio label.
 *
 * @param {number} width
 * @param {number} height
 * @returns {string}
 */
function aspectLabel(width, height) {
  if (width === 10 && height === 7.5) return 'Widescreen';
  if (width === 13.333 || (width === 13.3333 && height === 7.5)) return 'Widescreen 16:9';
  if (width === 7.5 && height === 10) return 'Portrait';
  return 'Custom';
}

/**
 * Describe a background for the report.
 *
 * @param {object|null} bg
 * @returns {string}
 */
function describeBackground(bg) {
  if (!bg) return 'None';
  if (bg.color) return `Solid #${bg.color}`;
  if (bg.path) return `Image: ${bg.path}`;
  return 'Unknown';
}

/**
 * Describe a placeholder for the report.
 *
 * @param {object} phObj - Placeholder object ({ placeholder: { options } })
 * @returns {string}
 */
function describePlaceholder(phObj) {
  const opts = phObj.placeholder?.options || {};
  let desc = `${opts.name || opts.type}: (${round4(opts.x || 0)}", ${round4(opts.y || 0)}") ${round4(opts.w || 0)}" \u00D7 ${round4(opts.h || 0)}"`;

  const parts = [];
  if (opts.fontFace) parts.push(opts.fontFace);
  if (opts.fontSize != null) parts.push(`${opts.fontSize}pt`);
  if (opts.color) parts.push(`#${opts.color}`);
  if (parts.length > 0) desc += ` \u2014 ${parts.join(' ')}`;

  return desc;
}

/**
 * Generate agent instructions markdown (SLIDE_MASTERS.md) teaching
 * future agents how to use the extracted slide masters.
 *
 * @param {Array<object>} masterData - Array of layout objects with mapped PptxGenJS data
 * @param {Record<string, string>} themeColors - Theme color map
 * @param {{ heading: string, body: string }} themeFonts - Theme fonts
 * @param {{ width: number, height: number }} dimensions - Slide dimensions
 * @param {string} templateName - Source template filename
 * @returns {string} Markdown content
 */
export function generateAgentInstructions(masterData, themeColors, themeFonts, dimensions, templateName) {
  const lines = [];
  const headingFont = themeFonts?.heading || 'Calibri';
  const bodyFont = themeFonts?.body || 'Calibri';
  const slideW = round4(dimensions?.width || 13.3333);
  const slideH = round4(dimensions?.height || 7.5);

  // Detect limited palette and extract fallback colors
  const { isLimited } = detectLimitedPalette(themeColors);
  const bgColors = extractBackgroundColors(masterData);
  const usesFallback = isLimited && bgColors.length > 0;

  // Build effective theme colors
  let effectiveThemeColors = themeColors;
  if (usesFallback && themeColors) {
    effectiveThemeColors = { ...themeColors };
    const accentSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    for (let i = 0; i < accentSlots.length; i++) {
      if (i < bgColors.length) {
        effectiveThemeColors[accentSlots[i]] = bgColors[i];
      }
    }
  }

  lines.push(`# Slide Masters: ${templateName}`);
  lines.push('');

  // --- Available Masters table ---
  lines.push('## Available Masters');
  lines.push('');
  lines.push('| Master | Title PH | Body PH | Slide # | Background |');
  lines.push('|--------|----------|---------|---------|------------|');

  for (const master of masterData) {
    const title = toUpperSnakeCase(master.name);
    const objects = master.objects || [];
    const titlePh = objects.find((o) => o.placeholder?.options?.type === 'title');
    const bodyPh = objects.find((o) => o.placeholder?.options?.type === 'body');
    const titlePhDesc = titlePh
      ? `\u2713 (${round4(titlePh.placeholder.options.x)},${round4(titlePh.placeholder.options.y)} ${round4(titlePh.placeholder.options.w)}\u00D7${round4(titlePh.placeholder.options.h)})`
      : '\u2717';
    const bodyPhDesc = bodyPh
      ? `\u2713 (${round4(bodyPh.placeholder.options.x)},${round4(bodyPh.placeholder.options.y)} ${round4(bodyPh.placeholder.options.w)}\u00D7${round4(bodyPh.placeholder.options.h)})`
      : '\u2717';
    const slideNumDesc = master.slideNumber ? '\u2713' : '\u2717';
    const bgDesc = master.background?.color ? `#${master.background.color}` : master.background?.path || 'None';
    lines.push(`| ${title} | ${titlePhDesc} | ${bodyPhDesc} | ${slideNumDesc} | ${bgDesc} |`);
  }
  lines.push('');

  // --- Quick Start ---
  lines.push('## Quick Start');
  lines.push('');
  lines.push('```js');
  if (usesFallback) {
    lines.push("import { createPresentation, THEME, THEME_COLORS, PALETTE, CHART_COLORS, FONT, POS, SLIDE_COLORS } from './masters.js';");
  } else {
    lines.push("import { createPresentation, THEME, THEME_COLORS, PALETTE, CHART_COLORS, FONT, POS } from './masters.js';");
  }
  lines.push('');
  lines.push("const pres = await createPresentation('Deck Title');");
  lines.push('');

  // One example per classified master
  const classified = masterData.map((m, i) => ({
    master: m,
    title: toUpperSnakeCase(m.name),
    cls: classifyMasterRole(m, i, masterData.length),
  }));

  let first = true;
  for (const { master, title, cls } of classified) {
    const objs = master.objects || [];
    const titlePh = objs.find((o) => o.placeholder?.options?.type === 'title');
    const bodyPh = objs.find((o) => o.placeholder?.options?.type === 'body');

    lines.push(`// ${cls.desc} (${master.name})`);
    if (first) {
      lines.push(`let slide = pres.addSlide({ masterName: '${title}' });`);
      first = false;
    } else {
      lines.push(`slide = pres.addSlide({ masterName: '${title}' });`);
    }

    if (titlePh) {
      const phName = titlePh.placeholder.options.name || 'title';
      const text = cls.role === 'cover' ? 'Deck Title' : cls.role === 'divider' ? 'Section Title' : 'Slide Title';
      lines.push(`slide.addText('${text}', { placeholder: '${phName}' });`);
    }
    if (bodyPh && cls.role !== 'end') {
      const phName = bodyPh.placeholder.options.name || 'body';
      if (cls.role === 'content') {
        lines.push(`slide.addText('Body content', { placeholder: '${phName}' });`);
      } else if (cls.role === 'divider') {
        lines.push(`slide.addText('Section subtitle', { placeholder: '${phName}' });`);
      } else if (cls.role === 'cover') {
        lines.push(`slide.addText('Subtitle', { placeholder: '${phName}' });`);
      }
    }
    lines.push('');
  }

  lines.push("await pres.writeFile({ fileName: 'output.pptx' });");
  lines.push('```');
  lines.push('');
  lines.push('Run with: `node script.mjs` (requires `pptxgenjs` installed).');
  lines.push('');

  // --- Exports Reference ---
  lines.push('## Exports from masters.js');
  lines.push('');
  lines.push('| Export | Type | Description |');
  lines.push('|--------|------|-------------|');
  lines.push('| `createPresentation(title?)` | `async function` | Returns PptxGenJS instance with masters registered and layout set |');
  lines.push('| `registerMasters(pptx)` | `function` | Register masters on an existing PptxGenJS instance |');
  lines.push('| `THEME` | `object` | Semantic color aliases: `text`, `background`, `brand`, `accent2`\u2013`accent6`, `footer` |');
  lines.push('| `THEME_COLORS` | `object` | Raw theme slots: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`\u2013`accent6`, `hlink`, `folHlink` |');
  lines.push('| `PALETTE` | `object` | Extended palette with tint/shade variants per color slot |');
  lines.push('| `CHART_COLORS` | `string[]` | Accent hex values ordered for chart data series |');
  lines.push('| `FONT` | `string` | Primary font name (heading font) |');
  lines.push('| `HEADING_FONT` | `string` | Heading font name |');
  lines.push('| `BODY_FONT` | `string` | Body font name |');
  lines.push('| `POS` | `object` | Pre-calculated content positioning areas |');
  if (usesFallback) {
    lines.push('| `SLIDE_COLORS` | `string[]` | Background colors extracted from master slides (fallback for limited accent palettes) |');
  }
  lines.push('');

  // --- POS table ---
  const posEntries = computeMasterPositions(masterData, dimensions);
  if (posEntries.length > 0) {
    lines.push('## Positioning (`POS`)');
    lines.push('');
    lines.push('Pre-calculated content areas. Spread into PptxGenJS options: `{ ...POS.body, chartColors: CHART_COLORS }`');
    lines.push('');
    lines.push('| Key | x | y | w | h | Notes |');
    lines.push('|-----|---|---|---|---|-------|');
    for (const pos of posEntries) {
      lines.push(`| \`${pos.key}\` | ${pos.x} | ${pos.y} | ${pos.w} | ${pos.h} | ${pos.notes} |`);
    }
    lines.push('');
  }

  // --- Theme Colors ---
  lines.push('## Theme Colors');
  lines.push('');
  if (usesFallback) {
    lines.push('> **Note:** Theme accents are limited (all near-white or identical). Colors below are derived from master slide backgrounds.');
    lines.push('');
  }

  lines.push('### THEME (semantic aliases)');
  lines.push('');
  if (themeColors) {
    if (themeColors.dk1) lines.push(`- \`THEME.text\` = \`#${themeColors.dk1}\` (default text)`);
    if (themeColors.lt1) lines.push(`- \`THEME.background\` = \`#${themeColors.lt1}\` (default background)`);
    if (usesFallback) {
      lines.push(`- \`THEME.brand\` = \`#${bgColors[0]}\` (from master backgrounds)`);
      for (let i = 1; i < bgColors.length && i < 6; i++) {
        lines.push(`- \`THEME.accent${i + 1}\` = \`#${bgColors[i]}\` (from master backgrounds)`);
      }
    } else {
      if (themeColors.accent1) lines.push(`- \`THEME.brand\` = \`#${themeColors.accent1}\` (primary brand accent)`);
      for (let i = 2; i <= 6; i++) {
        const slot = `accent${i}`;
        if (themeColors[slot]) lines.push(`- \`THEME.${slot}\` = \`#${themeColors[slot]}\``);
      }
    }
    const footerColor = detectMasterFooterColor(masterData, dimensions);
    if (footerColor) lines.push(`- \`THEME.footer\` = \`#${footerColor}\` (footer text)`);
  }
  lines.push('');

  lines.push('### THEME_COLORS (raw slots)');
  lines.push('');
  const colorLabels = {
    dk1: 'Dark 1 (text)',
    lt1: 'Light 1 (background)',
    dk2: 'Dark 2',
    lt2: 'Light 2',
    accent1: 'Accent 1 (primary)',
    accent2: 'Accent 2',
    accent3: 'Accent 3',
    accent4: 'Accent 4',
    accent5: 'Accent 5',
    accent6: 'Accent 6',
    hlink: 'Hyperlink',
    folHlink: 'Followed Hyperlink',
  };
  if (themeColors) {
    for (const [slot, label] of Object.entries(colorLabels)) {
      if (themeColors[slot]) {
        lines.push(`- ${label}: \`#${themeColors[slot]}\``);
      }
    }
  }
  lines.push('');

  // --- Chart Colors ---
  lines.push('### CHART_COLORS');
  lines.push('');
  if (usesFallback) {
    lines.push('Colors from master backgrounds (theme accents are limited): `slide.addChart(type, data, { chartColors: CHART_COLORS })`');
    lines.push('');
    let num = 1;
    for (const color of bgColors) {
      lines.push(`${num}. \`#${color}\` (from master backgrounds)`);
      num++;
    }
  } else {
    lines.push('Accent colors ordered for data visualization: `slide.addChart(type, data, { chartColors: CHART_COLORS })`');
    lines.push('');
    if (themeColors) {
      const chartSlots = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
      let num = 1;
      for (const slot of chartSlots) {
        if (themeColors[slot]) {
          lines.push(`${num}. \`#${themeColors[slot]}\` (${slot})`);
          num++;
        }
      }
    }
  }
  lines.push('');

  // --- Extended Palette ---
  lines.push('### Extended Palette (`PALETTE`)');
  lines.push('');
  lines.push('Each theme color has tint/shade variants matching PowerPoint\'s color picker:');
  lines.push('');
  lines.push('```js');
  lines.push('PALETTE.accent1.base       // base color');
  lines.push('PALETTE.accent1.lighter80  // 80% lighter (near white)');
  lines.push('PALETTE.accent1.lighter60  // 60% lighter');
  lines.push('PALETTE.accent1.lighter40  // 40% lighter');
  lines.push('PALETTE.accent1.lighter25  // 25% lighter');
  lines.push('PALETTE.accent1.darker25   // 25% darker');
  lines.push('PALETTE.accent1.darker50   // 50% darker (near black)');
  lines.push('```');
  lines.push('');
  // Show sample values for accent1
  const paletteAccent1 = usesFallback ? effectiveThemeColors?.accent1 : themeColors?.accent1;
  if (paletteAccent1) {
    const palette = generateExtendedPalette({ accent1: paletteAccent1 });
    if (palette.accent1) {
      const e = palette.accent1;
      lines.push(`Example (accent1 \`#${paletteAccent1}\`): lighter80=\`#${e.lighter80}\`, lighter40=\`#${e.lighter40}\`, darker25=\`#${e.darker25}\`, darker50=\`#${e.darker50}\``);
      lines.push('');
    }
  }
  lines.push('Available slots: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`\u2013`accent6`, `hlink`, `folHlink`. See `theme.json` for all computed values.');
  lines.push('');

  // --- Theme Fonts ---
  lines.push('## Theme Fonts');
  lines.push('');
  lines.push(`- Headings: ${headingFont}`);
  lines.push(`- Body: ${bodyFont}`);
  lines.push(`- \`FONT\` export = \`'${headingFont}'\``);
  lines.push('');

  // --- Master Details ---
  lines.push('## Master Details');
  lines.push('');

  for (const master of masterData) {
    const title = toUpperSnakeCase(master.name);
    lines.push(`### ${title}`);
    lines.push(`Original name: "${master.name}"`);

    const objects = master.objects || [];
    const placeholders = objects.filter((o) => o.placeholder);
    const statics = objects.filter((o) => !o.placeholder);

    if (placeholders.length > 0) {
      lines.push('');
      lines.push('Placeholders (use with `{ placeholder: \'name\' }`):');
      for (const phObj of placeholders) {
        const opts = phObj.placeholder.options;
        const parts = [`(${round4(opts.x)}", ${round4(opts.y)}") ${round4(opts.w)}" \u00D7 ${round4(opts.h)}"`];
        if (opts.fontFace) parts.push(opts.fontFace);
        if (opts.fontSize != null) parts.push(`${opts.fontSize}pt`);
        if (opts.color) parts.push(`#${opts.color}`);
        if (opts.align) parts.push(`${opts.align}-aligned`);
        lines.push(`- **\`${opts.name || opts.type}\`** (${opts.type}): ${parts.join(', ')}`);
      }
    } else {
      lines.push('');
      lines.push('No placeholders. Add content with explicit positioning: `slide.addText(text, { x, y, w, h })`');
    }

    if (master.slideNumber) {
      const sn = master.slideNumber;
      const parts = [];
      if (sn.fontFace) parts.push(sn.fontFace);
      if (sn.fontSize != null) parts.push(`${sn.fontSize}pt`);
      lines.push(`- Slide number: ${sn.align || 'right'}, ${parts.join(' ')} (automatic)`);
    }

    if (statics.length > 0) {
      lines.push(`- Static shapes: ${statics.length} (logos, bars, footer text \u2014 automatic)`);
    }

    lines.push('');
  }

  // --- Master Usage Guide ---
  lines.push('## Master Usage Guide');
  lines.push('');
  const classifiedForGuide = masterData.map((m, i) => ({
    master: m,
    title: toUpperSnakeCase(m.name),
    cls: classifyMasterRole(m, i, masterData.length),
  }));
  // Generate smart descriptions based on layout name and structure
  function getMasterDescription(master, cls) {
    const name = (master.name || '').toLowerCase();
    const objs = master.objects || [];
    const bodyPhs = objs.filter((o) => o.placeholder?.options?.type === 'body');
    // Count "large" content columns (h > 2" = actual content areas, not subtitle)
    const contentCols = bodyPhs.filter((o) => (o.placeholder.options.h || 0) > 2);
    const hasPicPh = objs.some((o) => o.placeholder?.options?.type === 'pic');

    // Background color indicator
    const bgColor = master.background?.color;
    const bgTag = bgColor ? (isDarkColor(bgColor) ? ' (dark bg)' : ' (light bg)') : '';

    let desc;
    if (cls.role === 'cover') {
      desc = hasPicPh
        ? 'Opening slide. Title/subtitle at bottom, picture placeholder for hero image. No slide number.'
        : 'Opening slide. Title and subtitle at bottom. No slide number.';
    } else if (cls.role === 'divider') {
      desc = 'Section break between major topics. Bold text centered on dark background.';
    } else if (cls.role === 'end') {
      desc = 'Closing slide. Branding elements only — do not add content.';
    } else if (cls.role === 'title-only') {
      desc = 'Title bar at top, full body area open. Best for charts, diagrams, full-bleed content.';
    } else if (name.includes('team') || name.includes('profile')) {
      desc = 'Team/bio layout with picture + text pairs. Use manual positioning for each member.';
    } else if (name.includes('qualif')) {
      desc = 'Qualifications/credentials layout. Use for credential summaries.';
    } else if (contentCols.length >= 3) {
      desc = `3-column layout. Title + subtitle, then 3 content columns (${round2(contentCols[0].placeholder.options.w)}" each). Use manual positioning per column.`;
    } else if (contentCols.length >= 2) {
      desc = `2-column layout. Title + subtitle, then ${contentCols.length} content columns (${round2(contentCols[0].placeholder.options.w)}" each). Use manual positioning per column.`;
    } else if (contentCols.length === 1) {
      desc = 'Title + subtitle + full-width content area. Best for tables, charts, detailed text.';
    } else if (bodyPhs.length === 1) {
      desc = 'Title + subtitle area only. Use for slides where the subtitle IS the content (key message, overview statement).';
    } else {
      desc = 'General content. Title at top, body below.';
    }

    return desc + bgTag;
  }
  lines.push('| Master | Role | When to Use |');
  lines.push('|--------|------|-------------|');
  for (const { master, title, cls } of classifiedForGuide) {
    const desc = getMasterDescription(master, cls);
    lines.push(`| \`${title}\` | ${cls.desc} | ${desc} |`);
  }
  lines.push('');

  // Check if any masters have multi-column content placeholders (h > 2" = real content, not subtitle)
  const multiColMasters = classifiedForGuide.filter(({ master }) => {
    const contentCols = (master.objects || [])
      .filter((o) => o.placeholder?.options?.type === 'body' && (o.placeholder.options.h || 0) > 2);
    return contentCols.length >= 2;
  });
  if (multiColMasters.length > 0) {
    lines.push('### Multi-Column Layouts');
    lines.push('');
    lines.push('Column placeholders share the same name but differ by x position. PptxGenJS fills by name (first match), so use **manual positioning** for multi-column content:');
    lines.push('');
    lines.push('```js');
    for (const { title, master } of multiColMasters) {
      const contentCols = (master.objects || [])
        .filter((o) => o.placeholder?.options?.type === 'body' && (o.placeholder.options.h || 0) > 2)
        .sort((a, b) => (a.placeholder.options.x || 0) - (b.placeholder.options.x || 0));
      if (contentCols.length >= 2) {
        lines.push(`// ${title} — ${contentCols.length} columns:`);
        for (let i = 0; i < contentCols.length; i++) {
          const opts = contentCols[i].placeholder.options;
          lines.push(`const col${i + 1} = { x: ${round2(opts.x)}, y: ${round2(opts.y)}, w: ${round2(opts.w)}, h: ${round2(opts.h)} };`);
        }
        lines.push(`slide.addText(content, { ...col1, fontFace: FONT, fontSize: 11, valign: 'top' });`);
        lines.push('');
      }
    }
    lines.push('```');
    lines.push('');
  }

  // --- Footer Customization ---
  // Detect footer text in static shapes
  const footerTexts = new Set();
  for (const master of masterData) {
    for (const obj of master.objects || []) {
      if (obj.text?.options && !obj.placeholder && (obj.text.options.y || 0) > (dimensions?.height || 7.5) * 0.8) {
        const text = typeof obj.text.text === 'string' ? obj.text.text : '';
        if (text && text.length > 0) footerTexts.add(text);
      }
    }
  }
  if (footerTexts.size > 0) {
    lines.push('## Footer Customization');
    lines.push('');
    lines.push('Masters include automatic footer text from the template. Default values are **placeholder text that should be updated** in `masters.js`:');
    lines.push('');
    for (const text of footerTexts) {
      lines.push(`- \`"${text}"\` — find and replace in masters.js with actual value`);
    }
    lines.push('');
    lines.push('Search masters.js for these strings and replace them before generating presentations. Example:');
    lines.push('');
    lines.push('```js');
    lines.push("// In masters.js, find and replace:");
    const firstFooter = [...footerTexts][0];
    if (firstFooter) {
      lines.push(`// "${firstFooter}" → "© ${new Date().getFullYear()} Company Name. All rights reserved."`);
    }
    lines.push('```');
    lines.push('');
  }

  // --- Typography ---
  lines.push('## Typography');
  lines.push('');
  lines.push(`Font: \`${headingFont}\` (use \`FONT\` constant). Always specify \`fontFace: FONT\` on every \`addText()\`, \`addTable()\`, and \`addChart()\` call.`);
  lines.push('');
  lines.push('Recommended sizes for manual content (placeholders handle their own sizing):');
  lines.push('');
  lines.push('| Element | Size | Weight | Color |');
  lines.push('|---------|------|--------|-------|');
  lines.push('| Section header (in body) | 14–16pt | Bold | `THEME.brand` |');
  lines.push('| Body text | 11–12pt | Regular | `THEME.text` |');
  lines.push('| Table header row | 10–11pt | Bold | `THEME.background` (on colored bg) |');
  lines.push('| Table body | 9–10pt | Regular | `THEME.text` |');
  lines.push('| Callout / highlight | 11–12pt | Bold | `THEME.accent2` or `THEME.brand` |');
  lines.push('| Footnote / source | 7–8pt | Regular | `THEME.footer` or gray |');
  lines.push('| Chart axis labels | 9–10pt | Regular | `THEME.text` |');
  lines.push('');

  // --- Content Best Practices ---
  lines.push('## Content Best Practices');
  lines.push('');
  lines.push('### Design Approach');
  lines.push('');
  lines.push('**Content first, then layout.** Decide what you need to communicate, then pick the master that fits:');
  lines.push('- Choose layouts based on content structure — let the material guide the design');
  lines.push('- Dense data and analysis → full-width layouts (TITLE_ONLY or single-column) for maximum space');
  lines.push('- Side-by-side comparisons or parallel concepts → multi-column layouts');
  lines.push('- Emphasis or transition → subtitle or divider layouts');
  lines.push('');
  const bodyH = round2(posEntries.find((p) => p.key === 'body')?.h || 5);
  const bodyW = round2(posEntries.find((p) => p.key === 'body')?.w || 12);
  const hasBodyFull = posEntries.some((p) => p.key === 'bodyFull');
  const bodyFullH = round2(posEntries.find((p) => p.key === 'bodyFull')?.h || bodyH);
  const capacityH = hasBodyFull ? bodyFullH : bodyH;
  const tableRows10pt = Math.floor(Number(capacityH) / 0.32);
  const tableRows9pt = Math.floor(Number(capacityH) / 0.28);
  const textLines11pt = Math.floor(Number(capacityH) / 0.22);
  if (hasBodyFull) {
    lines.push('**CRITICAL: Fill the content area.** Two positioning options:');
    lines.push(`- \`POS.body\` — ${bodyW}" × ${bodyH}" — use for layouts WITH a subtitle`);
    lines.push(`- \`POS.bodyFull\` — ${bodyW}" × ${bodyFullH}" — use for TITLE_ONLY layouts (starts right after title, more space)`);
    lines.push('');
    lines.push('Spread the appropriate POS into every content element. For TITLE_ONLY, always prefer `POS.bodyFull` for maximum density.');
  } else {
    lines.push('**CRITICAL: Fill the content area.** Use `POS.body` for positioning:');
    lines.push(`- \`POS.body\` — ${bodyW}" × ${bodyH}" — content area below title`);
    lines.push('');
    lines.push('Spread `{ ...POS.body }` into every content element.');
  }
  lines.push('');
  lines.push(`**Content capacity** (~${capacityH}" tall):`);
  lines.push(`- Table at 10pt: ~${tableRows10pt} rows | Table at 9pt: ~${tableRows9pt} rows`);
  lines.push(`- Text at 11pt: ~${textLines11pt} lines | Text at 9pt: ~${Math.floor(Number(capacityH) / 0.18)} lines`);
  lines.push(`- Chart: fills area automatically`);
  lines.push('');
  lines.push('### Tables');
  lines.push('');
  lines.push('```js');
  lines.push('slide.addTable(rows, {');
  lines.push('  ...POS.body,');
  lines.push('  fontFace: FONT, fontSize: 10, color: THEME.text,');
  lines.push('  border: { type: \'solid\', pt: 0.5, color: \'E0E0E0\' },');
  lines.push('  autoPage: true,');
  lines.push('});');
  lines.push('// Header row: { fill: { color: THEME.brand }, color: THEME.background, bold: true }');
  lines.push(`// Alt rows:   { fill: { color: PALETTE.accent1.lighter80 } }`);
  lines.push('```');
  lines.push('');
  lines.push('### Charts');
  lines.push('');
  lines.push('```js');
  lines.push('slide.addChart(pres.charts.BAR, chartData, {');
  lines.push('  ...POS.body, chartColors: CHART_COLORS,');
  lines.push('  showValue: true, valueFontSize: 9,');
  lines.push('  catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT,');
  lines.push('  valAxisLabelFontFace: FONT,');
  lines.push('});');
  lines.push('```');
  lines.push('');

  // --- PptxGenJS API Patterns ---
  lines.push('## PptxGenJS API Patterns');
  lines.push('');
  lines.push('### Text with Bullets');
  lines.push('');
  lines.push('Use a **flat array** with `bullet: true` per item. Do NOT use nested arrays — they produce corrupt files.');
  lines.push('');
  lines.push('```js');
  lines.push('// CORRECT: flat array, bullet per paragraph');
  lines.push('slide.addText([');
  lines.push("  { text: 'First point', options: { bullet: true, fontSize: 10, fontFace: FONT } },");
  lines.push("  { text: 'Second point', options: { bullet: true, fontSize: 10, fontFace: FONT } },");
  lines.push('], { x: 0.5, y: 1, w: 11, h: 3, valign: \'top\' });');
  lines.push('');
  lines.push('// CORRECT: bold header + normal text on same line');
  lines.push('slide.addText([');
  lines.push("  { text: 'Header: ', options: { bold: true, bullet: true, fontSize: 10, fontFace: FONT, color: THEME.brand } },");
  lines.push("  { text: 'Normal text continues here', options: { fontSize: 10, fontFace: FONT } },");
  lines.push("  { text: 'Next header: ', options: { bold: true, bullet: true, fontSize: 10, fontFace: FONT, color: THEME.brand, breakType: 'break' } },");
  lines.push("  { text: 'More text', options: { fontSize: 10, fontFace: FONT } },");
  lines.push('], { x: 0.5, y: 1, w: 11, h: 3, valign: \'top\' });');
  lines.push('');
  lines.push("// WRONG: nested arrays produce corrupt PPTX files");
  lines.push("// slide.addText([[{text: 'a'}, {text: 'b'}]], { bullet: true }); // DON'T DO THIS");
  lines.push('```');
  lines.push('');
  lines.push('### Sizing and Overflow');
  lines.push('');
  lines.push('- PptxGenJS does **not** auto-shrink text to fit. If text overflows the box, it clips or spills.');
  lines.push('- Tables do **not** auto-expand rows. Set `rowH` explicitly or let PptxGenJS calculate (omit `rowH`).');
  lines.push('- When placing multiple elements vertically, **calculate y positions manually**: `nextY = prevY + prevH + gap`.');
  lines.push(`- Body area: ${bodyW}" × ${bodyH}" starting at (${round2(posEntries.find((p) => p.key === 'body')?.x || 0.5)}, ${round2(posEntries.find((p) => p.key === 'body')?.y || 1.84)}). Plan your layout to fill this rectangle.`);
  lines.push('');
  lines.push('### Shapes');
  lines.push('');
  lines.push('```js');
  lines.push('// Colored box (KPI card, callout, section divider)');
  lines.push('slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {');
  lines.push('  x: 0.5, y: 1, w: 3, h: 0.6, fill: { color: PALETTE.accent1.lighter80 }, rectRadius: 0.05,');
  lines.push('});');
  lines.push('// Horizontal rule');
  lines.push('slide.addShape(pres.shapes.LINE, { x: 0.5, y: 4, w: 11.5, h: 0, line: { color: \'D0D0D0\', width: 0.5 } });');
  lines.push('```');
  lines.push('');

  // --- Key Conventions ---
  lines.push('## Key Conventions');
  lines.push('');
  lines.push(`- **CRITICAL: \`fontFace: FONT\`** on EVERY \`addText()\` / \`addTable()\` / \`addChart()\` call. PptxGenJS does not cascade fonts from masters.`);
  if (hasBodyFull) {
    lines.push('- **CRITICAL: \`{ ...POS.bodyFull }\`** for TITLE_ONLY layouts, **\`{ ...POS.body }\`** for subtitled layouts. This ensures content fills the available space.');
  } else {
    lines.push('- **CRITICAL: \`{ ...POS.body }\`** on content elements. This ensures content fills the available space.');
  }
  lines.push('- **Footers and logos are automatic.** Masters include copyright, slide numbers, and branding. Do not add these manually.');
  lines.push('- **Placeholders** handle font, size, color, and position: `slide.addText(text, { placeholder: \'Name\' })`.');
  lines.push('- **Suppress unused placeholders:** Fill any placeholder you don\'t need with a space: `slide.addText(\' \', { placeholder: \'Name\' })`. This prevents "Click to add..." prompts and empty boxes in exports.');
  lines.push(`- **Slide dimensions:** ${slideW}" \u00D7 ${slideH}".`);
  lines.push('- **Color usage:** `THEME.text` for body text, `THEME.brand` for emphasis/headers, accents for differentiation, `PALETTE.accent1.lighter80` for alternating table rows, `PALETTE.accent1.lighter60` for subtle backgrounds.');
  if (usesFallback) {
    lines.push('- **Dark backgrounds:** Use `THEME.background` (white) for text on dark master backgrounds.');
  }
  lines.push('');

  // --- Common Pitfalls ---
  lines.push('## Common Pitfalls');
  lines.push('');
  lines.push('- **Text overflow**: PptxGenJS does not auto-shrink text. If content exceeds the text box, it silently clips or overflows. Split long bullet lists into columns or across slides.');
  lines.push('- **Empty space**: If the bottom third of your slide is unused, the slide looks unfinished. Distribute content across the full body area or choose a smaller layout.');
  lines.push('- **Color contrast**: Dark text on dark backgrounds (or light text on light backgrounds) is unreadable. Check the Available Masters table for each layout\'s background color and choose text colors accordingly.');
  lines.push(`- **Layout monotony**: Using the same master layout repeatedly makes the deck feel generic. This template has ${masterData.length} layouts — vary them.`);
  lines.push('- **Missing fontFace**: PptxGenJS does not cascade fonts from masters. Every `addText()`, `addTable()`, `addChart()` call needs `fontFace: FONT` or text renders in Calibri.');
  lines.push('- **Table column overflow**: Column widths must sum to the container width. Mismatched widths cause silent overflow.');
  lines.push('- **Concatenated text blocks**: Build each logical section as a separate `addText()` call with its own positioning, not as one giant concatenated string.');
  lines.push('');

  // --- Workflow ---
  lines.push('## Workflow');
  lines.push('');
  lines.push('### 1. Design with HTML preview');
  lines.push('');
  lines.push('Create an HTML file with one `<div>` per slide to prototype layout and spacing before writing PptxGenJS code:');
  lines.push('');
  lines.push('```html');
  lines.push(`<div style="width:${slideW}in; height:${slideH}in; border:1px solid #ccc; position:relative; font-family:${headingFont},sans-serif; overflow:hidden;">`);
  lines.push('  <!-- Title bar -->');
  lines.push(`  <div style="position:absolute; left:0.5in; top:0.38in; width:12.3in; height:0.37in; font-size:21pt;">Title</div>`);
  const bfPos = posEntries.find((p) => p.key === 'bodyFull');
  if (bfPos) {
    lines.push(`  <!-- Body area -->`);
    lines.push(`  <div style="position:absolute; left:${bfPos.x}in; top:${bfPos.y}in; width:${bfPos.w}in; height:${bfPos.h}in; border:1px dashed #${effectiveThemeColors?.accent1 || '666666'};">`);
  } else {
    const bPos = posEntries.find((p) => p.key === 'body');
    if (bPos) {
      lines.push(`  <div style="position:absolute; left:${bPos.x}in; top:${bPos.y}in; width:${bPos.w}in; height:${bPos.h}in; border:1px dashed #${effectiveThemeColors?.accent1 || '666666'};">`);
    }
  }
  lines.push('    <!-- Your content layout here: tables, charts, text blocks -->');
  lines.push('  </div>');
  lines.push('</div>');
  lines.push('```');
  lines.push('');
  lines.push('### 2. Visually inspect HTML');
  lines.push('');
  lines.push('Render the HTML preview and **look at it**. Check every slide for visual quality — does content fill the body area? Is text readable? Is the layout balanced? Fix in HTML until every slide looks right. Iterating here is much faster than re-generating PPTX.');
  lines.push('');
  lines.push('### 3. Translate to PptxGenJS');
  lines.push('');
  lines.push('Convert the finalized HTML layouts to `addText()`, `addTable()`, `addChart()`, `addShape()` calls using the same coordinates.');
  lines.push('');
  lines.push('### 4. Render PPTX and visually verify');
  lines.push('');
  lines.push('```bash');
  lines.push('node slides.mjs                                      # Generate PPTX');
  lines.push('soffice --headless --convert-to pdf output.pptx       # Convert to PDF');
  lines.push('pdftoppm -png -r 150 output.pdf slide                 # Split into per-slide PNGs');
  lines.push('```');
  lines.push('');
  lines.push('**View every slide PNG.** Look for:');
  lines.push('- Text clipping or overflow (content cut off at edges)');
  lines.push('- Large empty areas (unused lower half of slide)');
  lines.push('- Unreadable text (wrong color on background, too small)');
  lines.push('- Visual monotony (same layout repeated)');
  lines.push('- Missing fonts (text appears in default Calibri instead of template font)');
  lines.push('');
  lines.push('Fix issues and re-render until every slide looks right.');
  lines.push('');

  // --- Style Guide reference ---
  lines.push('## Style Guide');
  lines.push('');
  lines.push('If a `STYLE_GUIDE.md` file exists in this directory, read it before building slides. It contains user-specific design preferences (dos/don\'ts, typography, color usage, layout preferences) that override or supplement the defaults in this document.');
  lines.push('');

  return lines.join('\n');
}
