/**
 * Report generator — produces a Markdown extraction report
 * summarizing theme, layouts, placeholders, and warnings.
 */

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
 * Pick the closest emoji for a hex color.
 *
 * @param {string} hex - 6-character hex color
 * @returns {string}
 */
function colorEmoji(hex) {
  if (!hex || typeof hex !== 'string') return '';

  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const brightness = (r * 299 + g * 587 + b * 114) / 1000;

  if (brightness < 50) return '\u2B1B';
  if (brightness > 220) return '\u2B1C';

  const max = Math.max(r, g, b);
  if (max === r && r > g * 1.3 && r > b * 1.3) return '\uD83D\uDFE5';
  if (max === g && g > r * 1.3 && g > b * 1.3) return '\uD83D\uDFE9';
  if (max === b && b > r * 1.3 && b > g * 1.3) return '\uD83D\uDFE6';
  if (r > 200 && g > 150 && b < 100) return '\uD83D\uDFE7';
  if (r > 200 && g > 200 && b < 100) return '\uD83D\uDFE8';
  if (r > 100 && b > 100 && g < 100) return '\uD83D\uDFEA';

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
  if ((width === 13.333 || width === 13.3333) && height === 7.5) return 'Widescreen 16:9';
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
 * Generate an extraction report in Markdown format.
 *
 * @param {object} extractionResult - Complete extraction result
 * @param {string} extractionResult.templateName - Source template filename
 * @param {{ width: number, height: number }} extractionResult.dimensions - Slide dimensions
 * @param {Record<string, string>} extractionResult.themeColors - Theme color map
 * @param {{ heading: string, body: string }} extractionResult.themeFonts - Theme fonts
 * @param {Array<object>} extractionResult.layouts - Layout extraction data
 * @param {string[]} extractionResult.allWarnings - All accumulated warnings
 * @returns {string} Markdown report
 */
export function generateReport(extractionResult) {
  const {
    templateName,
    dimensions,
    themeColors,
    themeFonts,
    layouts,
    allWarnings,
  } = extractionResult;

  const date = new Date().toISOString().slice(0, 10);
  const lines = [];

  // Header
  lines.push('# pptx-masters Extraction Report');
  lines.push('');
  lines.push(`**Template:** ${templateName}`);
  lines.push(`**Date:** ${date}`);
  lines.push(`**Slide Dimensions:** ${dimensions?.width || 10}" \u00D7 ${dimensions?.height || 7.5}" (${aspectLabel(dimensions?.width || 10, dimensions?.height || 7.5)})`);
  lines.push('');

  // Theme section
  lines.push('## Theme');
  lines.push('');

  // Colors table
  lines.push('### Colors');
  lines.push('| Slot | Hex | Preview |');
  lines.push('|------|-----|---------|');

  const colorSlots = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  if (themeColors) {
    for (const slot of colorSlots) {
      if (themeColors[slot] != null) {
        lines.push(`| ${slot} | #${themeColors[slot]} | ${colorEmoji(themeColors[slot])} |`);
      }
    }
  }

  lines.push('');

  // Fonts
  lines.push('### Fonts');
  lines.push(`- **Heading:** ${themeFonts?.heading || 'Calibri'}`);
  lines.push(`- **Body:** ${themeFonts?.body || 'Calibri'}`);
  lines.push('');

  // Layouts section
  lines.push('## Layouts');
  lines.push('');

  if (layouts && layouts.length > 0) {
    for (let i = 0; i < layouts.length; i++) {
      const layout = layouts[i];
      lines.push(`### ${i + 1}. ${layout.name}`);
      lines.push(`- **Background:** ${describeBackground(layout.background)}`);

      // Placeholders
      if (layout.placeholders && layout.placeholders.length > 0) {
        lines.push('- **Placeholders:**');
        for (const ph of layout.placeholders) {
          const pos = ph.position || {};
          let desc = `  - ${ph.name || ph.type}: (${round4(pos.x || 0)}", ${round4(pos.y || 0)}") ${round4(pos.w || 0)}" \u00D7 ${round4(pos.h || 0)}"`;
          const parts = [];
          if (ph.fontFace) parts.push(ph.fontFace);
          if (ph.fontSize != null) parts.push(`${ph.fontSize}pt`);
          if (ph.color) parts.push(`#${ph.color}`);
          if (parts.length > 0) desc += ` \u2014 ${parts.join(' ')}`;
          lines.push(desc);
        }
      } else {
        lines.push('- **Placeholders:** None');
      }

      // Slide number
      if (layout.slideNumber) {
        const sn = layout.slideNumber;
        const parts = [`(${round4(sn.x || 0)}", ${round4(sn.y || 0)}")`];
        if (sn.fontFace) parts.push(sn.fontFace);
        if (sn.fontSize != null) parts.push(`${sn.fontSize}pt`);
        if (sn.color) parts.push(`#${sn.color}`);
        lines.push(`- **Slide Number:** ${parts.join(' ')}`);
      }

      // Static shapes count
      const staticCount = layout.staticShapes?.length || 0;
      lines.push(`- **Static Shapes:** ${staticCount}`);

      // Footer objects
      if (layout.footerObjects && layout.footerObjects.length > 0) {
        lines.push(`- **Footer Objects:** ${layout.footerObjects.length}`);
      }

      // Per-layout warnings
      if (layout.warnings && layout.warnings.length > 0) {
        lines.push('- **Warnings:**');
        for (const w of layout.warnings) {
          lines.push(`  - ${w}`);
        }
      }

      lines.push('');
    }
  }

  // Warnings section
  lines.push('## Warnings');
  lines.push('');

  if (allWarnings && allWarnings.length > 0) {
    for (const w of allWarnings) {
      lines.push(`- ${w}`);
    }
  } else {
    lines.push('No warnings.');
  }

  lines.push('');

  // What's Not Supported section
  lines.push("## What's Not Supported (v1)");
  lines.push('- Gradient fills (dominant color used as fallback)');
  lines.push('- Pattern fills (foreground color used as fallback)');
  lines.push('- Grouped shapes');
  lines.push('- Animations and transitions');
  lines.push('- SmartArt / diagrams');
  lines.push('- 3D effects, text warp');
  lines.push('');

  return lines.join('\n');
}
