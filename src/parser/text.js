/**
 * Text property extraction from OOXML text body elements (a:txBody).
 *
 * Used by master.js, layout.js, and shape mappers to extract
 * text content and styling from PowerPoint elements.
 */

import { emuToInches, emuAngleToDegrees } from '../mapper/units.js';

// Default body margins in EMU (OOXML spec defaults)
const DEFAULT_LINS = 91440;
const DEFAULT_TINS = 45720;
const DEFAULT_RINS = 91440;
const DEFAULT_BINS = 45720;

// Vertical alignment mapping
const ANCHOR_MAP = {
  t: 'top',
  ctr: 'middle',
  b: 'bottom',
};

// Paragraph alignment mapping
const ALIGN_MAP = {
  l: 'left',
  ctr: 'center',
  r: 'right',
  just: 'justify',
};

/**
 * Extract text properties from a txBody element.
 *
 * @param {object} txBody - Parsed a:txBody element (from fast-xml-parser)
 * @param {{ resolve: Function, resolveFontRef: Function }} colorResolver - Color resolver from colors.js
 * @param {{ heading: string, body: string }} themeFonts - Theme font faces
 * @returns {{ bodyProps: object, paragraphs: Array, plainText: string, lstStyleProps: object|null }}
 */
export function extractTextProps(txBody, colorResolver, themeFonts) {
  if (!txBody) {
    return {
      bodyProps: buildDefaultBodyProps(),
      paragraphs: [],
      plainText: '',
      lstStyleProps: null,
    };
  }

  const bodyProps = extractBodyProps(txBody['a:bodyPr']);
  const lstStyleProps = extractLstStyle(txBody['a:lstStyle'], colorResolver, themeFonts);
  const paragraphs = extractParagraphs(txBody, colorResolver, themeFonts);

  // Build plainText from all runs
  const plainText = paragraphs
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('\n');

  return { bodyProps, paragraphs, plainText, lstStyleProps };
}

/**
 * Extract default text style from a defRPr element.
 * Useful for extracting styling from placeholder defaults.
 *
 * @param {object} defRPr - Parsed a:defRPr element
 * @param {{ resolve: Function, resolveFontRef: Function }} colorResolver
 * @param {{ heading: string, body: string }} themeFonts
 * @returns {object} Run properties shape
 */
export function extractDefaultTextStyle(defRPr, colorResolver, themeFonts) {
  return extractRunProps(defRPr, colorResolver, themeFonts);
}

// --- Internal helpers ---

/**
 * Build default body properties.
 * @returns {object}
 */
function buildDefaultBodyProps() {
  return {
    margin: [
      emuToInches(DEFAULT_TINS),
      emuToInches(DEFAULT_RINS),
      emuToInches(DEFAULT_BINS),
      emuToInches(DEFAULT_LINS),
    ],
    valign: undefined,
    rotation: undefined,
    vert: undefined,
    autoFit: undefined,
  };
}

/**
 * Extract body properties from a:bodyPr element.
 * @param {object} bodyPr - Parsed a:bodyPr element
 * @returns {object}
 */
function extractBodyProps(bodyPr) {
  if (!bodyPr) return buildDefaultBodyProps();

  // Internal margins (EMU → inches)
  const lIns = bodyPr['@_lIns'] != null ? Number(bodyPr['@_lIns']) : DEFAULT_LINS;
  const tIns = bodyPr['@_tIns'] != null ? Number(bodyPr['@_tIns']) : DEFAULT_TINS;
  const rIns = bodyPr['@_rIns'] != null ? Number(bodyPr['@_rIns']) : DEFAULT_RINS;
  const bIns = bodyPr['@_bIns'] != null ? Number(bodyPr['@_bIns']) : DEFAULT_BINS;

  const margin = [
    emuToInches(tIns),
    emuToInches(rIns),
    emuToInches(bIns),
    emuToInches(lIns),
  ];

  // Vertical alignment
  const anchor = bodyPr['@_anchor'];
  const valign = ANCHOR_MAP[anchor] || undefined;

  // Rotation
  const rot = bodyPr['@_rot'] != null ? emuAngleToDegrees(Number(bodyPr['@_rot'])) : undefined;

  // Text direction
  const vert = bodyPr['@_vert'] || undefined;

  // Auto-fit
  // Note: fast-xml-parser parses self-closing tags like <a:spAutoFit/> as ""
  // (empty string), which is falsy. Must check with != null instead of truthiness.
  let autoFit;
  if (bodyPr['a:normAutofit'] != null) {
    autoFit = 'shrink';
  } else if (bodyPr['a:spAutoFit'] != null) {
    autoFit = 'resize';
  } else if (bodyPr['a:noAutofit'] != null) {
    autoFit = 'none';
  }

  return { margin, valign, rotation: rot, vert, autoFit };
}

/**
 * Extract all paragraphs from a txBody element.
 * @param {object} txBody
 * @param {object} colorResolver
 * @param {object} themeFonts
 * @returns {Array}
 */
function extractParagraphs(txBody, colorResolver, themeFonts) {
  const pElements = txBody['a:p'];
  if (!pElements) return [];

  const items = Array.isArray(pElements) ? pElements : [pElements];
  return items.map((p) => extractParagraph(p, colorResolver, themeFonts));
}

/**
 * Extract a single paragraph.
 * @param {object} p - Parsed a:p element
 * @param {object} colorResolver
 * @param {object} themeFonts
 * @returns {object}
 */
function extractParagraph(p, colorResolver, themeFonts) {
  const pPr = p['a:pPr'] || {};

  // Alignment — track whether it was explicitly set for lstStyle merge
  const _explicitAlign = ALIGN_MAP[pPr['@_algn']] || undefined;
  const align = _explicitAlign || 'left';

  // Paragraph level (0-based in XML, 1-based for lstStyle key)
  const level = pPr['@_lvl'] != null ? Number(pPr['@_lvl']) + 1 : 1;

  // RTL
  const rtlMode = pPr['@_rtl'] === '1' || pPr['@_rtl'] === 'true' || pPr['@_rtl'] === true;

  // Margin left (EMU → inches)
  const marginLeft = pPr['@_marL'] != null ? emuToInches(Number(pPr['@_marL'])) : 0;

  // Indent (EMU → inches)
  const indent = pPr['@_indent'] != null ? emuToInches(Number(pPr['@_indent'])) : 0;

  // Line spacing
  const { lineSpacing, lineSpacingMultiple } = extractLineSpacing(pPr['a:lnSpc']);

  // Paragraph spacing
  const paraSpaceBefore = extractSpacing(pPr['a:spcBef']);
  const paraSpaceAfter = extractSpacing(pPr['a:spcAft']);

  // Bullets
  const bullet = extractBullet(pPr, colorResolver);

  // Default run properties
  const defRPr = pPr['a:defRPr']
    ? extractRunProps(pPr['a:defRPr'], colorResolver, themeFonts)
    : undefined;

  // Runs
  const runs = extractRuns(p, colorResolver, themeFonts);

  return {
    align,
    _explicitAlign,
    level,
    rtlMode,
    marginLeft,
    indent,
    lineSpacing,
    lineSpacingMultiple,
    paraSpaceBefore,
    paraSpaceAfter,
    bullet,
    runs,
    defaultRunProps: defRPr,
  };
}

/**
 * Extract line spacing from a:lnSpc element.
 * @param {object} lnSpc
 * @returns {{ lineSpacing: number | undefined, lineSpacingMultiple: number | undefined }}
 */
function extractLineSpacing(lnSpc) {
  if (!lnSpc) return { lineSpacing: undefined, lineSpacingMultiple: undefined };

  // a:spcPts → spacing in hundredths of a point
  if (lnSpc['a:spcPts']) {
    const val = Number(lnSpc['a:spcPts']['@_val'] || 0);
    return { lineSpacing: val / 100, lineSpacingMultiple: undefined };
  }

  // a:spcPct → spacing as percentage (e.g. 150000 = 150% = 1.5)
  if (lnSpc['a:spcPct']) {
    const val = Number(lnSpc['a:spcPct']['@_val'] || 0);
    return { lineSpacing: undefined, lineSpacingMultiple: val / 100000 };
  }

  return { lineSpacing: undefined, lineSpacingMultiple: undefined };
}

/**
 * Extract spacing value from a:spcBef or a:spcAft element.
 * Returns points.
 * @param {object} spc
 * @returns {number | undefined}
 */
function extractSpacing(spc) {
  if (!spc) return undefined;

  if (spc['a:spcPts']) {
    return Number(spc['a:spcPts']['@_val'] || 0) / 100;
  }

  // spcPct for before/after — less common but possible
  if (spc['a:spcPct']) {
    return Number(spc['a:spcPct']['@_val'] || 0) / 100000;
  }

  return undefined;
}

/**
 * Extract bullet properties from paragraph properties.
 * @param {object} pPr - Paragraph properties
 * @param {object} colorResolver
 * @returns {{ type: string, characterCode?: string, numberType?: string, fontFace?: string, color?: string, sizePercent?: number } | false | undefined}
 */
function extractBullet(pPr, colorResolver) {
  if (!pPr) return undefined;

  // Explicit no bullet
  if (pPr['a:buNone'] != null) return false;

  const isBuChar = pPr['a:buChar'] != null;
  const isBuAutoNum = pPr['a:buAutoNum'] != null;

  if (!isBuChar && !isBuAutoNum) return undefined;

  const bullet = {};

  // Bullet font
  if (pPr['a:buFont']) {
    bullet.fontFace = pPr['a:buFont']['@_typeface'] || '';
  }

  // Bullet color
  if (pPr['a:buClr']) {
    const resolved = colorResolver?.resolve(pPr['a:buClr']);
    bullet.color = resolved?.color || '000000';
  }

  // Bullet size percent
  if (pPr['a:buSzPct']) {
    bullet.sizePercent = Number(pPr['a:buSzPct']['@_val'] || 0) / 1000;
  }

  if (isBuChar) {
    const char = pPr['a:buChar']['@_char'] || '';
    bullet.type = 'char';
    bullet.characterCode = char.codePointAt(0)?.toString(16).toUpperCase() || '';
    return bullet;
  }

  if (isBuAutoNum) {
    bullet.type = 'number';
    bullet.numberType = pPr['a:buAutoNum']['@_type'] || '';
    return bullet;
  }

  return undefined;
}

/**
 * Extract all runs from a paragraph element.
 * Handles a:r (text runs), a:fld (field elements like slide numbers),
 * and a:br (line break elements).
 *
 * @param {object} p - Parsed a:p element
 * @param {object} colorResolver
 * @param {object} themeFonts
 * @returns {Array}
 */
function extractRuns(p, colorResolver, themeFonts) {
  const contentRuns = [];

  // Regular text runs (a:r)
  if (p['a:r']) {
    const items = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']];
    for (const r of items) {
      const text = r['a:t'] != null ? String(r['a:t']) : '';
      const rPr = r['a:rPr'] || {};
      const props = extractRunProps(rPr, colorResolver, themeFonts);
      contentRuns.push({ text, ...props });
    }
  }

  // Field elements (a:fld) — slide numbers, dates, etc.
  if (p['a:fld']) {
    const items = Array.isArray(p['a:fld']) ? p['a:fld'] : [p['a:fld']];
    for (const fld of items) {
      const text = fld['a:t'] != null ? String(fld['a:t']) : '';
      const rPr = fld['a:rPr'] || {};
      const props = extractRunProps(rPr, colorResolver, themeFonts);
      contentRuns.push({ text, ...props, isField: true, fieldType: fld['@_type'] || '' });
    }
  }

  // Line breaks (a:br) — interleave between content runs
  if (p['a:br']) {
    const brItems = Array.isArray(p['a:br']) ? p['a:br'] : [p['a:br']];
    const breakCount = brItems.length;

    // Common case: N breaks between N+1 runs (e.g., run1 <br/> run2)
    if (breakCount > 0 && contentRuns.length > 1 && breakCount <= contentRuns.length - 1) {
      const result = [];
      for (let i = 0; i < contentRuns.length; i++) {
        result.push(contentRuns[i]);
        if (i < breakCount) {
          result.push({ text: '\n', isBreak: true });
        }
      }
      return result;
    }

    // Fallback: append breaks after content
    for (let i = 0; i < breakCount; i++) {
      contentRuns.push({ text: '\n', isBreak: true });
    }
  }

  return contentRuns;
}

/**
 * Extract run-level properties from a:rPr or a:defRPr element.
 * @param {object} rPr - Run properties element
 * @param {object} colorResolver
 * @param {object} themeFonts
 * @returns {object}
 */
function extractRunProps(rPr, colorResolver, themeFonts) {
  if (!rPr) {
    return {
      fontSize: undefined,
      bold: undefined,
      italic: undefined,
      underline: undefined,
      strike: undefined,
      fontFace: undefined,
      color: undefined,
      highlight: undefined,
      superscript: false,
      subscript: false,
      charSpacing: undefined,
    };
  }

  // Font size: hundredths of a point → points (via /100)
  const fontSize = rPr['@_sz'] != null ? Number(rPr['@_sz']) / 100 : undefined;

  // Bold / italic — undefined when not specified (distinguishes from explicit false)
  const bold = rPr['@_b'] != null
    ? (rPr['@_b'] === '1' || rPr['@_b'] === 'true' || rPr['@_b'] === true)
    : undefined;
  const italic = rPr['@_i'] != null
    ? (rPr['@_i'] === '1' || rPr['@_i'] === 'true' || rPr['@_i'] === true)
    : undefined;

  // Underline
  const underline = rPr['@_u'] || undefined;

  // Strike
  const strike = rPr['@_strike'] || undefined;

  // Superscript / subscript (baseline attribute, positive = super, negative = sub)
  const baseline = rPr['@_baseline'] != null ? Number(rPr['@_baseline']) : 0;
  const superscript = baseline > 0;
  const subscript = baseline < 0;

  // Character spacing (spc attribute, in hundredths of a point)
  const charSpacing = rPr['@_spc'] != null ? Number(rPr['@_spc']) / 100 : undefined;

  // Font face
  let fontFace;
  if (rPr['a:latin']) {
    const typeface = rPr['a:latin']['@_typeface'];
    if (typeface) {
      // Resolve theme font references
      if (typeface.startsWith('+')) {
        fontFace = colorResolver?.resolveFontRef
          ? colorResolver.resolveFontRef(typeface)
          : resolveThemeFont(typeface, themeFonts);
      } else {
        fontFace = typeface;
      }
    }
  }

  // Color: look for a:solidFill child
  let color;
  if (rPr['a:solidFill']) {
    const resolved = colorResolver?.resolve(rPr['a:solidFill']);
    color = resolved?.color;
  }

  // Highlight
  let highlight;
  if (rPr['a:highlight']) {
    const resolved = colorResolver?.resolve(rPr['a:highlight']);
    highlight = resolved?.color;
  }

  return {
    fontSize,
    bold,
    italic,
    underline,
    strike,
    fontFace,
    color,
    highlight,
    superscript,
    subscript,
    charSpacing,
  };
}

/**
 * Resolve a theme font reference when no colorResolver.resolveFontRef is available.
 * @param {string} ref - Font reference like '+mj-lt'
 * @param {{ heading: string, body: string }} themeFonts
 * @returns {string}
 */
function resolveThemeFont(ref, themeFonts) {
  if (!themeFonts) return ref;
  if (ref === '+mj-lt' || ref === '+mj-ea' || ref === '+mj-cs') {
    return themeFonts.heading;
  }
  if (ref === '+mn-lt' || ref === '+mn-ea' || ref === '+mn-cs') {
    return themeFonts.body;
  }
  return ref;
}

/**
 * Extract list style properties from a:lstStyle element.
 * Contains level-based paragraph properties (a:lvl1pPr through a:lvl9pPr)
 * that define default text formatting for placeholders.
 *
 * @param {object} lstStyle - Parsed a:lstStyle element
 * @param {object} colorResolver
 * @param {object} themeFonts
 * @returns {Record<number, object>|null} Level-keyed properties (1-9), or null
 */
function extractLstStyle(lstStyle, colorResolver, themeFonts) {
  if (!lstStyle) return null;

  const levels = {};

  // Default paragraph properties
  if (lstStyle['a:defPPr']) {
    levels[0] = extractLevelProps(lstStyle['a:defPPr'], colorResolver, themeFonts);
  }

  // Level-specific properties (1-9)
  for (let i = 1; i <= 9; i++) {
    const lvlPr = lstStyle[`a:lvl${i}pPr`];
    if (lvlPr) {
      levels[i] = extractLevelProps(lvlPr, colorResolver, themeFonts);
    }
  }

  return Object.keys(levels).length > 0 ? levels : null;
}

/**
 * Extract paragraph-level properties from a level element (a:lvl1pPr, etc.).
 *
 * @param {object} lvlPr - Level paragraph properties element
 * @param {object} colorResolver
 * @param {object} themeFonts
 * @returns {object}
 */
function extractLevelProps(lvlPr, colorResolver, themeFonts) {
  const align = ALIGN_MAP[lvlPr['@_algn']] || undefined;
  const marginLeft = lvlPr['@_marL'] != null ? emuToInches(Number(lvlPr['@_marL'])) : undefined;
  const indent = lvlPr['@_indent'] != null ? emuToInches(Number(lvlPr['@_indent'])) : undefined;

  const { lineSpacing, lineSpacingMultiple } = extractLineSpacing(lvlPr['a:lnSpc']);
  const paraSpaceBefore = extractSpacing(lvlPr['a:spcBef']);
  const paraSpaceAfter = extractSpacing(lvlPr['a:spcAft']);
  const bullet = extractBullet(lvlPr, colorResolver);

  const defRPr = lvlPr['a:defRPr']
    ? extractRunProps(lvlPr['a:defRPr'], colorResolver, themeFonts)
    : undefined;

  return {
    align,
    marginLeft,
    indent,
    lineSpacing,
    lineSpacingMultiple,
    paraSpaceBefore,
    paraSpaceAfter,
    bullet,
    defaultRunProps: defRPr,
  };
}
