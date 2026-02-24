/**
 * Color resolution engine for PPTX theme colors.
 *
 * Handles the full color resolution chain:
 *   schemeClr val → clrMap lookup → theme slot → hex → HSL modifiers → final hex
 */

// --- OOXML Preset Color Table (ECMA-376 ST_PresetColorVal) ---

/**
 * Complete mapping of OOXML preset color names to hex RGB values.
 * OOXML uses camelCase abbreviations: dk=dark, lt=light, med=medium.
 * Values match the CSS/W3C named color standard.
 */
const PRESET_COLORS = {
  aliceBlue: 'F0F8FF',
  antiqueWhite: 'FAEBD7',
  aqua: '00FFFF',
  aquamarine: '7FFFD4',
  azure: 'F0FFFF',
  beige: 'F5F5DC',
  bisque: 'FFE4C4',
  black: '000000',
  blanchedAlmond: 'FFEBCD',
  blue: '0000FF',
  blueViolet: '8A2BE2',
  brown: 'A52A2A',
  burlyWood: 'DEB887',
  cadetBlue: '5F9EA0',
  chartreuse: '7FFF00',
  chocolate: 'D2691E',
  coral: 'FF7F50',
  cornflowerBlue: '6495ED',
  cornsilk: 'FFF8DC',
  crimson: 'DC143C',
  cyan: '00FFFF',
  dkBlue: '00008B',
  dkCyan: '008B8B',
  dkGoldenrod: 'B8860B',
  dkGray: 'A9A9A9',
  dkGreen: '006400',
  dkKhaki: 'BDB76B',
  dkMagenta: '8B008B',
  dkOliveGreen: '556B2F',
  dkOrange: 'FF8C00',
  dkOrchid: '9932CC',
  dkRed: '8B0000',
  dkSalmon: 'E9967A',
  dkSeaGreen: '8FBC8F',
  dkSlateBlue: '483D8B',
  dkSlateGray: '2F4F4F',
  dkTurquoise: '00CED1',
  dkViolet: '9400D3',
  deepPink: 'FF1493',
  deepSkyBlue: '00BFFF',
  dimGray: '696969',
  dodgerBlue: '1E90FF',
  firebrick: 'B22222',
  floralWhite: 'FFFAF0',
  forestGreen: '228B22',
  fuchsia: 'FF00FF',
  gainsboro: 'DCDCDC',
  ghostWhite: 'F8F8FF',
  gold: 'FFD700',
  goldenrod: 'DAA520',
  gray: '808080',
  green: '008000',
  greenYellow: 'ADFF2F',
  honeydew: 'F0FFF0',
  hotPink: 'FF69B4',
  indianRed: 'CD5C5C',
  indigo: '4B0082',
  ivory: 'FFFFF0',
  khaki: 'F0E68C',
  lavender: 'E6E6FA',
  lavenderBlush: 'FFF0F5',
  lawnGreen: '7CFC00',
  lemonChiffon: 'FFFACD',
  ltBlue: 'ADD8E6',
  ltCoral: 'F08080',
  ltCyan: 'E0FFFF',
  ltGoldenrodYellow: 'FAFAD2',
  ltGray: 'D3D3D3',
  ltGreen: '90EE90',
  ltPink: 'FFB6C1',
  ltSalmon: 'FFA07A',
  ltSeaGreen: '20B2AA',
  ltSkyBlue: '87CEFA',
  ltSlateGray: '778899',
  ltSteelBlue: 'B0C4DE',
  ltYellow: 'FFFFE0',
  lime: '00FF00',
  limeGreen: '32CD32',
  linen: 'FAF0E6',
  magenta: 'FF00FF',
  maroon: '800000',
  medAquamarine: '66CDAA',
  medBlue: '0000CD',
  medOrchid: 'BA55D3',
  medPurple: '9370DB',
  medSeaGreen: '3CB371',
  medSlateBlue: '7B68EE',
  medSpringGreen: '00FA9A',
  medTurquoise: '48D1CC',
  medVioletRed: 'C71585',
  midnightBlue: '191970',
  mintCream: 'F5FFFA',
  mistyRose: 'FFE4E1',
  moccasin: 'FFE4B5',
  navajoWhite: 'FFDEAD',
  navy: '000080',
  oldLace: 'FDF5E6',
  olive: '808000',
  oliveDrab: '6B8E23',
  orange: 'FFA500',
  orangeRed: 'FF4500',
  orchid: 'DA70D6',
  paleGoldenrod: 'EEE8AA',
  paleGreen: '98FB98',
  paleTurquoise: 'AFEEEE',
  paleVioletRed: 'DB7093',
  papayaWhip: 'FFEFD5',
  peachPuff: 'FFDAB9',
  peru: 'CD853F',
  pink: 'FFC0CB',
  plum: 'DDA0DD',
  powderBlue: 'B0E0E6',
  purple: '800080',
  red: 'FF0000',
  rosyBrown: 'BC8F8F',
  royalBlue: '4169E1',
  saddleBrown: '8B4513',
  salmon: 'FA8072',
  sandyBrown: 'F4A460',
  seaGreen: '2E8B57',
  seaShell: 'FFF5EE',
  sienna: 'A0522D',
  silver: 'C0C0C0',
  skyBlue: '87CEEB',
  slateBlue: '6A5ACD',
  slateGray: '708090',
  snow: 'FFFAFA',
  springGreen: '00FF7F',
  steelBlue: '4682B4',
  tan: 'D2B48C',
  teal: '008080',
  thistle: 'D8BFD8',
  tomato: 'FF6347',
  turquoise: '40E0D0',
  violet: 'EE82EE',
  wheat: 'F5DEB3',
  white: 'FFFFFF',
  whiteSmoke: 'F5F5F5',
  yellow: 'FFFF00',
  yellowGreen: '9ACD32',
};

// --- RGB / HSL conversion utilities ---

/**
 * Convert hex string to RGB array.
 * @param {string} hex - 'RRGGBB' (no hash)
 * @returns {[number, number, number]} [r, g, b] each 0-255
 */
export function hexToRgb(hex) {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return [r, g, b];
}

/**
 * Convert RGB to hex string.
 * @param {number} r - 0-255
 * @param {number} g - 0-255
 * @param {number} b - 0-255
 * @returns {string} 'RRGGBB' uppercase
 */
export function rgbToHex(r, g, b) {
  const toHex = (v) => Math.round(Math.max(0, Math.min(255, v)))
    .toString(16)
    .padStart(2, '0')
    .toUpperCase();
  return toHex(r) + toHex(g) + toHex(b);
}

/**
 * Convert RGB to HSL.
 * @param {number} r - 0-255
 * @param {number} g - 0-255
 * @param {number} b - 0-255
 * @returns {[number, number, number]} [h, s, l] h in 0-360, s/l in 0-1
 */
export function rgbToHsl(r, g, b) {
  r /= 255;
  g /= 255;
  b /= 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;

  if (max === min) {
    return [0, 0, l];
  }

  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

  let h;
  if (max === r) {
    h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
  } else if (max === g) {
    h = ((b - r) / d + 2) / 6;
  } else {
    h = ((r - g) / d + 4) / 6;
  }

  return [h * 360, s, l];
}

/**
 * Convert HSL to RGB.
 * @param {number} h - 0-360
 * @param {number} s - 0-1
 * @param {number} l - 0-1
 * @returns {[number, number, number]} [r, g, b] each 0-255
 */
export function hslToRgb(h, s, l) {
  h /= 360;

  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }

  const hue2rgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;

  const r = hue2rgb(p, q, h + 1 / 3);
  const g = hue2rgb(p, q, h);
  const b = hue2rgb(p, q, h - 1 / 3);

  return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}

// --- Color modifier application ---

/**
 * Apply OOXML color modifiers to a hex color.
 * Supports tint, shade, lumMod, lumOff, satMod, satOff, hueMod, hueOff.
 *
 * @param {string} hex - Base color 'RRGGBB'
 * @param {object} modifiers - Modifier values in OOXML units (100000 = 100%)
 * @returns {string} Modified color 'RRGGBB'
 */
export function applyColorModifiers(hex, modifiers) {
  let [r, g, b] = hexToRgb(hex);

  // Shade first per OOXML spec: scale toward black (shade=75000 retains 75%)
  if (modifiers.shade != null) {
    const s = modifiers.shade / 100000;
    r = Math.round(r * s);
    g = Math.round(g * s);
    b = Math.round(b * s);
  }

  // Tint second: mix toward white (tint=50000 moves 50% toward white)
  if (modifiers.tint != null) {
    const t = modifiers.tint / 100000;
    r = Math.round(r + (255 - r) * t);
    g = Math.round(g + (255 - g) * t);
    b = Math.round(b + (255 - b) * t);
  }

  let [h, sat, l] = rgbToHsl(r, g, b);

  // Hue modifiers
  if (modifiers.hueMod != null) {
    h = (h * (modifiers.hueMod / 100000)) % 360;
  }
  if (modifiers.hueOff != null) {
    // hueOff is in 60000ths of a degree
    h = ((h + modifiers.hueOff / 60000) % 360 + 360) % 360;
  }

  // Saturation modifiers
  if (modifiers.satMod != null) {
    sat = sat * (modifiers.satMod / 100000);
    sat = Math.max(0, Math.min(1, sat));
  }
  if (modifiers.satOff != null) {
    sat = sat + (modifiers.satOff / 100000);
    sat = Math.max(0, Math.min(1, sat));
  }

  // Luminance modifiers
  if (modifiers.lumMod != null) {
    l = l * (modifiers.lumMod / 100000);
  }
  if (modifiers.lumOff != null) {
    l = l + (modifiers.lumOff / 100000);
  }

  l = Math.max(0, Math.min(1, l));

  const [nr, ng, nb] = hslToRgb(h, sat, l);
  return rgbToHex(nr, ng, nb);
}

// --- Color resolver factory ---

/**
 * Default color map (standard OOXML mapping).
 * Slide masters override this with their own p:clrMap.
 */
const DEFAULT_CLR_MAP = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hlink: 'hlink',
  folHlink: 'folHlink',
};

/**
 * Create a color resolver bound to theme colors and a color map.
 *
 * @param {Record<string, string>} themeColors - Scheme color slots → hex values
 *   e.g. { dk1: '000000', lt1: 'FFFFFF', accent1: '4472C4', ... }
 * @param {Record<string, string>} [clrMap] - Logical names → scheme slot names
 *   e.g. { bg1: 'lt1', tx1: 'dk1', ... }
 * @param {{ heading: string, body: string }} [themeFonts] - Theme font faces
 * @returns {{ resolve, resolveSchemeColor, resolveFontRef }}
 */
export function createColorResolver(themeColors, clrMap, themeFonts) {
  const map = clrMap || DEFAULT_CLR_MAP;
  const fonts = themeFonts || { heading: 'Calibri Light', body: 'Calibri' };

  /**
   * Resolve a scheme color name (e.g. 'tx1') to its hex value.
   * Walks: schemeName → clrMap slot → themeColors hex.
   * @param {string} schemeName
   * @returns {string} 'RRGGBB'
   */
  function resolveSchemeColor(schemeName) {
    // First check if the name is in the clrMap (e.g. tx1 → dk1)
    const mapped = map[schemeName] || schemeName;
    // Then look up in themeColors
    const hex = themeColors[mapped];
    if (!hex) {
      // Fallback: try the original name directly in themeColors
      return themeColors[schemeName] || '000000';
    }
    return hex;
  }

  /**
   * Resolve a font reference like '+mj-lt' or '+mn-lt'.
   * @param {string} fontRef - Font reference string
   * @returns {string} Font face name
   */
  function resolveFontRef(fontRef) {
    if (fontRef === '+mj-lt' || fontRef === '+mj-ea' || fontRef === '+mj-cs') {
      return fonts.heading;
    }
    if (fontRef === '+mn-lt' || fontRef === '+mn-ea' || fontRef === '+mn-cs') {
      return fonts.body;
    }
    // Not a theme font ref — return as-is
    return fontRef;
  }

  /**
   * Resolve a parsed color element to its final hex and transparency.
   *
   * Handles:
   * - a:srgbClr with optional modifiers
   * - a:schemeClr with optional modifiers
   * - a:sysClr
   * - a:prstClr (preset named colors → hex lookup)
   *
   * @param {object} colorElement - Parsed XML color element (from fast-xml-parser)
   * @returns {{ color: string, transparency?: number } | null}
   */
  function resolve(colorElement) {
    if (!colorElement) return null;

    // Direct sRGB color
    if (colorElement['a:srgbClr']) {
      return resolveSrgbClr(colorElement['a:srgbClr']);
    }

    // Scheme color
    if (colorElement['a:schemeClr']) {
      return resolveSchemeClrElement(colorElement['a:schemeClr']);
    }

    // System color
    if (colorElement['a:sysClr']) {
      return {
        color: colorElement['a:sysClr']['@_lastClr'] || '000000',
      };
    }

    // Preset color — resolve name to hex via lookup table
    if (colorElement['a:prstClr']) {
      const name = colorElement['a:prstClr']['@_val'];
      const hex = PRESET_COLORS[name] || '000000';
      const modifiers = extractModifiers(colorElement['a:prstClr']);
      let finalHex = hex;
      if (hasModifiers(modifiers)) {
        finalHex = applyColorModifiers(hex, modifiers);
      }
      const result = { color: finalHex };
      if (modifiers.alpha != null) {
        result.transparency = Math.round(100 - modifiers.alpha / 1000);
      }
      return result;
    }

    return null;
  }

  /**
   * Resolve an a:srgbClr element (may have inline modifiers).
   * @param {object} el
   * @returns {{ color: string, transparency?: number }}
   */
  function resolveSrgbClr(el) {
    let hex = el['@_val'] || '000000';
    const modifiers = extractModifiers(el);
    if (hasModifiers(modifiers)) {
      hex = applyColorModifiers(hex, modifiers);
    }
    const result = { color: hex };
    if (modifiers.alpha != null) {
      result.transparency = Math.round(100 - modifiers.alpha / 1000);
    }
    return result;
  }

  /**
   * Resolve an a:schemeClr element with the full resolution chain.
   * @param {object} el
   * @returns {{ color: string, transparency?: number }}
   */
  function resolveSchemeClrElement(el) {
    const schemeName = el['@_val'];
    // phClr (placeholder color) cannot be resolved without shape context
    if (schemeName === 'phClr') {
      return { color: '000000' };
    }

    let hex = resolveSchemeColor(schemeName);
    const modifiers = extractModifiers(el);
    if (hasModifiers(modifiers)) {
      hex = applyColorModifiers(hex, modifiers);
    }
    const result = { color: hex };
    if (modifiers.alpha != null) {
      result.transparency = Math.round(100 - modifiers.alpha / 1000);
    }
    return result;
  }

  return { resolve, resolveSchemeColor, resolveFontRef };
}

// --- Modifier extraction helpers ---

/**
 * Extract OOXML color modifiers from a parsed color element.
 * @param {object} el - Parsed element that may contain modifier child elements
 * @returns {object} Extracted modifier values in OOXML units
 */
function extractModifiers(el) {
  const mods = {};
  if (el['a:tint']) {
    mods.tint = Number(el['a:tint']['@_val']);
  }
  if (el['a:shade']) {
    mods.shade = Number(el['a:shade']['@_val']);
  }
  if (el['a:hueMod']) {
    mods.hueMod = Number(el['a:hueMod']['@_val']);
  }
  if (el['a:hueOff']) {
    mods.hueOff = Number(el['a:hueOff']['@_val']);
  }
  if (el['a:satMod']) {
    mods.satMod = Number(el['a:satMod']['@_val']);
  }
  if (el['a:satOff']) {
    mods.satOff = Number(el['a:satOff']['@_val']);
  }
  if (el['a:lumMod']) {
    mods.lumMod = Number(el['a:lumMod']['@_val']);
  }
  if (el['a:lumOff']) {
    mods.lumOff = Number(el['a:lumOff']['@_val']);
  }
  if (el['a:alpha']) {
    mods.alpha = Number(el['a:alpha']['@_val']);
  }
  return mods;
}

/**
 * Check if any color-modifying properties (excluding alpha) are present.
 * @param {object} mods
 * @returns {boolean}
 */
function hasModifiers(mods) {
  return mods.tint != null || mods.shade != null
    || mods.lumMod != null || mods.lumOff != null
    || mods.satMod != null || mods.satOff != null
    || mods.hueMod != null || mods.hueOff != null;
}

export { PRESET_COLORS };
