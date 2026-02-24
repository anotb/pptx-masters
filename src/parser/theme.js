/**
 * Theme parser — extracts colors, fonts, and format scheme from ppt/theme/theme1.xml.
 */

/** The 12 standard OOXML scheme color slot names. */
const SCHEME_COLOR_NAMES = [
  'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
];

/**
 * Parse a theme XML object (already parsed by fast-xml-parser) and extract
 * scheme colors, fonts, and format scheme.
 *
 * @param {object} themeXml - Parsed ppt/theme/theme1.xml
 * @returns {{
 *   colors: Record<string, string>,
 *   fonts: { heading: string, body: string },
 *   formatScheme: { fillStyleLst: any, lnStyleLst: any, effectStyleLst: any, bgFillStyleLst: any } | null
 * }}
 */
export function parseTheme(themeXml) {
  const theme = themeXml['a:theme'];
  if (!theme) {
    throw new Error('Invalid theme XML: missing a:theme root element');
  }

  const elements = theme['a:themeElements'];
  if (!elements) {
    throw new Error('Invalid theme XML: missing a:themeElements');
  }

  const colors = extractSchemeColors(elements['a:clrScheme']);
  const fonts = extractFontScheme(elements['a:fontScheme']);
  const formatScheme = extractFormatScheme(elements['a:fmtScheme']);

  return { colors, fonts, formatScheme };
}

/**
 * Extract the 12 scheme colors from a:clrScheme.
 *
 * Each slot contains either:
 *   - a:srgbClr val="RRGGBB" → direct hex
 *   - a:sysClr val="windowText" lastClr="000000" → use lastClr
 *
 * @param {object} clrScheme - Parsed a:clrScheme element
 * @returns {Record<string, string>} Map of slot name → hex ('RRGGBB')
 */
function extractSchemeColors(clrScheme) {
  if (!clrScheme) {
    return {};
  }

  const colors = {};

  for (const name of SCHEME_COLOR_NAMES) {
    const slot = clrScheme[`a:${name}`];
    if (!slot) continue;

    if (slot['a:srgbClr']) {
      colors[name] = slot['a:srgbClr']['@_val'];
    } else if (slot['a:sysClr']) {
      // System colors (windowText, window, etc.) — use lastClr for the resolved value
      colors[name] = slot['a:sysClr']['@_lastClr'] || '000000';
    }
  }

  return colors;
}

/**
 * Extract heading and body font faces from a:fontScheme.
 *
 * @param {object} fontScheme - Parsed a:fontScheme element
 * @returns {{ heading: string, body: string }}
 */
function extractFontScheme(fontScheme) {
  const result = { heading: '', body: '' };

  if (!fontScheme) return result;

  const majorFont = fontScheme['a:majorFont'];
  if (majorFont?.['a:latin']) {
    result.heading = majorFont['a:latin']['@_typeface'] || '';
  }

  const minorFont = fontScheme['a:minorFont'];
  if (minorFont?.['a:latin']) {
    result.body = minorFont['a:latin']['@_typeface'] || '';
  }

  return result;
}

/**
 * Extract format scheme sections (stored for reference resolution by downstream parsers).
 *
 * @param {object} fmtScheme - Parsed a:fmtScheme element
 * @returns {{ fillStyleLst: any, lnStyleLst: any, effectStyleLst: any, bgFillStyleLst: any } | null}
 */
function extractFormatScheme(fmtScheme) {
  if (!fmtScheme) return null;

  return {
    fillStyleLst: fmtScheme['a:fillStyleLst'] || null,
    lnStyleLst: fmtScheme['a:lnStyleLst'] || null,
    effectStyleLst: fmtScheme['a:effectStyleLst'] || null,
    bgFillStyleLst: fmtScheme['a:bgFillStyleLst'] || null,
  };
}

/**
 * Parse a p:clrMap element from a slide master into a lookup object.
 *
 * @param {object} clrMapEl - Parsed p:clrMap element with @_ prefixed attributes
 * @returns {Record<string, string>} Map of logical name → scheme slot (e.g. { bg1: 'lt1', tx1: 'dk1' })
 */
export function parseClrMap(clrMapEl) {
  if (!clrMapEl) return {};

  const map = {};
  for (const [key, value] of Object.entries(clrMapEl)) {
    if (key.startsWith('@_')) {
      map[key.slice(2)] = value;
    }
  }
  return map;
}
