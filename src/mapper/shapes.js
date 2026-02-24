/**
 * Shape mapper — converts parsed OOXML shapes into PptxGenJS-compatible objects.
 *
 * Handles rectangles, lines, text boxes, and pictures.
 */

import { emuToPoints } from './units.js';

// Dash type mapping: OOXML → PptxGenJS
const DASH_TYPE_MAP = {
  solid: 'solid',
  dash: 'dash',
  dashDot: 'dashDot',
  lgDash: 'lgDash',
  lgDashDot: 'lgDashDot',
  lgDashDotDot: 'lgDashDotDot',
  sysDash: 'sysDash',
  sysDot: 'sysDot',
  sysDashDot: 'sysDashDot',
  sysDashDotDot: 'sysDashDotDot',
  dot: 'dot',
};

// Arrow type mapping: OOXML → PptxGenJS
const ARROW_TYPE_MAP = {
  none: 'none',
  triangle: 'triangle',
  stealth: 'stealth',
  diamond: 'diamond',
  oval: 'oval',
  arrow: 'arrow',
};

/**
 * Resolve a fill element to PptxGenJS fill properties.
 *
 * @param {{ type: string, element: object }|null} fill - Parsed fill from layout.js
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @returns {{ result: { color: string, transparency?: number }|null, warnings: string[] }}
 */
export function resolveFill(fill, colorResolver) {
  const warnings = [];

  if (!fill) return { result: null, warnings };

  const { type, element } = fill;

  if (type === 'noFill') {
    return { result: null, warnings };
  }

  if (type === 'solidFill') {
    const resolved = colorResolver?.resolve(element);
    if (resolved) {
      const result = { color: resolved.color };
      if (resolved.transparency != null) {
        result.transparency = resolved.transparency;
      }
      return { result, warnings };
    }
    return { result: null, warnings };
  }

  if (type === 'gradFill') {
    // Use first gradient stop as fallback
    const gsLst = element?.['a:gsLst'];
    if (gsLst) {
      const gsItems = Array.isArray(gsLst['a:gs']) ? gsLst['a:gs'] : gsLst['a:gs'] ? [gsLst['a:gs']] : [];
      if (gsItems.length > 0) {
        // Use first stop
        const firstStop = gsItems[0];
        const resolved = colorResolver?.resolve(firstStop);
        if (resolved) {
          warnings.push('Gradient fill not fully supported, using first stop color as fallback');
          const result = { color: resolved.color };
          if (resolved.transparency != null) {
            result.transparency = resolved.transparency;
          }
          return { result, warnings };
        }
      }
    }
    warnings.push('Gradient fill not supported, could not resolve fallback color');
    return { result: null, warnings };
  }

  if (type === 'blipFill') {
    // Image fill — return a marker; caller handles image resolution
    const rId = element?.['a:blip']?.['@_r:embed'];
    if (rId) {
      return { result: { imageRef: rId }, warnings };
    }
    warnings.push('blipFill without image reference');
    return { result: null, warnings };
  }

  if (type === 'pattFill') {
    // Use foreground color as fallback
    const fgClr = element?.['a:fgClr'];
    if (fgClr) {
      const resolved = colorResolver?.resolve(fgClr);
      if (resolved) {
        warnings.push('Pattern fill not supported, using foreground color as fallback');
        return { result: { color: resolved.color }, warnings };
      }
    }
    warnings.push('Pattern fill not supported');
    return { result: null, warnings };
  }

  return { result: null, warnings };
}

/**
 * Resolve a line element (a:ln) to PptxGenJS line properties.
 *
 * @param {object|null} lineEl - Raw a:ln element from parser
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @returns {{ color?: string, width?: number, dashType?: string, beginArrowType?: string, endArrowType?: string }|null}
 */
export function resolveLine(lineEl, colorResolver) {
  if (!lineEl) return null;

  const result = {};
  let hasProps = false;

  // Width: EMU → points
  if (lineEl['@_w'] != null) {
    result.width = emuToPoints(Number(lineEl['@_w']));
    hasProps = true;
  }

  // Color from solidFill
  if (lineEl['a:solidFill']) {
    const resolved = colorResolver?.resolve(lineEl['a:solidFill']);
    if (resolved) {
      result.color = resolved.color;
      hasProps = true;
    }
  }

  // Dash type
  if (lineEl['a:prstDash']) {
    const dashVal = lineEl['a:prstDash']['@_val'];
    result.dashType = DASH_TYPE_MAP[dashVal] || dashVal || 'solid';
    hasProps = true;
  }

  // Arrow heads
  if (lineEl['a:headEnd']) {
    const headType = lineEl['a:headEnd']['@_type'];
    if (headType && headType !== 'none') {
      result.beginArrowType = ARROW_TYPE_MAP[headType] || headType;
      hasProps = true;
    }
  }

  if (lineEl['a:tailEnd']) {
    const tailType = lineEl['a:tailEnd']['@_type'];
    if (tailType && tailType !== 'none') {
      result.endArrowType = ARROW_TYPE_MAP[tailType] || tailType;
      hasProps = true;
    }
  }

  return hasProps ? result : null;
}

/**
 * Resolve shadow effects from an effectLst element.
 *
 * @param {object|null} effectLst - Raw a:effectLst element
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @returns {object|null} PptxGenJS shadow object
 */
export function resolveShadow(effectLst, colorResolver) {
  if (!effectLst) return null;

  // Outer shadow
  const outerShdw = effectLst['a:outerShdw'];
  if (outerShdw) {
    return buildShadow(outerShdw, 'outer', colorResolver);
  }

  // Inner shadow
  const innerShdw = effectLst['a:innerShdw'];
  if (innerShdw) {
    return buildShadow(innerShdw, 'inner', colorResolver);
  }

  return null;
}

/**
 * Build a PptxGenJS shadow object from a shadow element.
 * @param {object} shdwEl - Shadow element (a:outerShdw or a:innerShdw)
 * @param {string} type - 'outer' or 'inner'
 * @param {{ resolve: Function }} colorResolver
 * @returns {object}
 */
function buildShadow(shdwEl, type, colorResolver) {
  const shadow = { type };

  // Blur radius in points
  if (shdwEl['@_blurRad'] != null) {
    shadow.blur = emuToPoints(Number(shdwEl['@_blurRad']));
  }

  // Distance in points
  if (shdwEl['@_dist'] != null) {
    shadow.offset = emuToPoints(Number(shdwEl['@_dist']));
  }

  // Direction in degrees (EMU angle)
  if (shdwEl['@_dir'] != null) {
    shadow.angle = Number(shdwEl['@_dir']) / 60000;
  }

  // Shadow color
  const resolved = colorResolver?.resolve(shdwEl);
  if (resolved) {
    shadow.color = resolved.color;
    if (resolved.transparency != null) {
      shadow.opacity = (100 - resolved.transparency) / 100;
    }
  }

  return shadow;
}

/**
 * Convert textProps into PptxGenJS-compatible text options.
 * Merges properties from paragraphs, lstStyle (layout-specific defaults),
 * and run properties with proper fallback chain.
 *
 * @param {{ bodyProps: object, paragraphs: Array, plainText: string, lstStyleProps?: object }|null} textProps
 * @returns {object} PptxGenJS text options
 */
export function mapTextPropsToOptions(textProps) {
  if (!textProps) return {};

  const opts = {};

  // Body props
  if (textProps.bodyProps) {
    const bp = textProps.bodyProps;
    if (bp.margin) {
      opts.margin = bp.margin;
    }
    // OOXML default anchor is 't' (top), but PptxGenJS defaults to 'ctr'.
    // Always set valign explicitly to match the original template behavior.
    opts.valign = bp.valign || 'top';
  }

  // lstStyle level 1 defaults (layout-specific text formatting)
  const lstLevel = textProps.lstStyleProps?.[1] || textProps.lstStyleProps?.[0] || null;

  // First paragraph styling with lstStyle fallback
  if (textProps.paragraphs && textProps.paragraphs.length > 0) {
    const para = textProps.paragraphs[0];

    // Alignment: explicit paragraph → lstStyle → default 'left'
    if (para._explicitAlign) {
      opts.align = para._explicitAlign;
    } else if (lstLevel?.align) {
      opts.align = lstLevel.align;
    } else if (para.align) {
      opts.align = para.align;
    }

    // Spacing: paragraph → lstStyle fallback
    opts.lineSpacing = para.lineSpacing ?? lstLevel?.lineSpacing ?? undefined;
    opts.lineSpacingMultiple = para.lineSpacingMultiple ?? lstLevel?.lineSpacingMultiple ?? undefined;
    opts.paraSpaceBefore = para.paraSpaceBefore ?? lstLevel?.paraSpaceBefore ?? undefined;
    opts.paraSpaceAfter = para.paraSpaceAfter ?? lstLevel?.paraSpaceAfter ?? undefined;

    // Clean up undefined spacing values
    if (opts.lineSpacing == null) delete opts.lineSpacing;
    if (opts.lineSpacingMultiple == null) delete opts.lineSpacingMultiple;
    if (opts.paraSpaceBefore == null) delete opts.paraSpaceBefore;
    if (opts.paraSpaceAfter == null) delete opts.paraSpaceAfter;

    // Run styling: paragraph defRPr → lstStyle defRPr → first run
    const paraDefRPr = para.defaultRunProps || {};
    const lstDefRPr = lstLevel?.defaultRunProps || {};
    const firstRun = (para.runs && para.runs[0]) || {};

    const fontFace = paraDefRPr.fontFace ?? lstDefRPr.fontFace ?? firstRun.fontFace;
    const fontSize = paraDefRPr.fontSize ?? lstDefRPr.fontSize ?? firstRun.fontSize;
    const color = paraDefRPr.color ?? lstDefRPr.color ?? firstRun.color;
    const bold = paraDefRPr.bold ?? lstDefRPr.bold ?? firstRun.bold;
    const italic = paraDefRPr.italic ?? lstDefRPr.italic ?? firstRun.italic;

    if (fontFace) opts.fontFace = fontFace;
    if (fontSize != null) opts.fontSize = fontSize;
    if (color) opts.color = color;
    if (bold) opts.bold = bold;
    if (italic) opts.italic = italic;
  } else if (lstLevel) {
    // No paragraphs — use lstStyle directly
    if (lstLevel.align) opts.align = lstLevel.align;
    if (lstLevel.lineSpacing != null) opts.lineSpacing = lstLevel.lineSpacing;
    if (lstLevel.lineSpacingMultiple != null) opts.lineSpacingMultiple = lstLevel.lineSpacingMultiple;
    if (lstLevel.paraSpaceBefore != null) opts.paraSpaceBefore = lstLevel.paraSpaceBefore;
    if (lstLevel.paraSpaceAfter != null) opts.paraSpaceAfter = lstLevel.paraSpaceAfter;

    const rp = lstLevel.defaultRunProps || {};
    if (rp.fontFace) opts.fontFace = rp.fontFace;
    if (rp.fontSize != null) opts.fontSize = rp.fontSize;
    if (rp.color) opts.color = rp.color;
    if (rp.bold) opts.bold = rp.bold;
    if (rp.italic) opts.italic = rp.italic;
  }

  return opts;
}

/**
 * Map a parsed static shape into a PptxGenJS-compatible object.
 *
 * @param {object} parsedShape - Parsed shape from layout.js staticShapes
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @param {{ heading: string, body: string }} themeFonts - Theme fonts
 * @param {Record<string, { type: string, target: string }>} relationships - Resolved relationships
 * @returns {{ object: object, warnings: string[] }}
 */
export function mapShape(parsedShape, colorResolver, themeFonts, relationships) {
  const warnings = [];

  if (!parsedShape) {
    return { object: null, warnings: ['Null shape provided'] };
  }

  const pos = parsedShape.position || {};

  // Image shapes
  if (parsedShape.type === 'picture' && parsedShape.imageRef) {
    const rel = relationships?.[parsedShape.imageRef];
    if (rel) {
      const imagePath = `./media/${rel.target.split('/').pop()}`;
      const imgObj = {
        x: pos.x,
        y: pos.y,
        w: pos.w,
        h: pos.h,
      };
      if (imagePath) imgObj.path = imagePath;
      if (parsedShape.rotation) imgObj.rotate = parsedShape.rotation;
      return { object: { image: imgObj }, warnings };
    }
    warnings.push(`Could not resolve image reference ${parsedShape.imageRef}`);
  }

  // Line shapes
  if (parsedShape.geometry === 'line') {
    const lineProps = resolveLine(parsedShape.line, colorResolver) || {};
    const lineObj = {
      x: pos.x,
      y: pos.y,
      w: pos.w,
      h: pos.h,
    };
    if (lineProps.color) lineObj.line = lineProps;
    else if (Object.keys(lineProps).length > 0) lineObj.line = lineProps;
    if (parsedShape.rotation) lineObj.rotate = parsedShape.rotation;
    return { object: { line: lineObj }, warnings };
  }

  // Text box shapes (shapes with text content)
  if (parsedShape.textProps && parsedShape.textProps.plainText) {
    const textOptions = mapTextPropsToOptions(parsedShape.textProps);
    textOptions.x = pos.x;
    textOptions.y = pos.y;
    textOptions.w = pos.w;
    textOptions.h = pos.h;

    // Fill
    const { result: fillResult, warnings: fillWarnings } = resolveFill(parsedShape.fill, colorResolver);
    warnings.push(...fillWarnings);
    if (fillResult && fillResult.color) {
      textOptions.fill = fillResult;
    }

    // Line
    const lineResult = resolveLine(parsedShape.line, colorResolver);
    if (lineResult) {
      textOptions.line = lineResult;
    }

    if (parsedShape.rotation) textOptions.rotate = parsedShape.rotation;

    // Build text content — flatten to string because PptxGenJS defineSlideMaster
    // wraps text.text in [{ text: ... }], so arrays get stringified to [object Object]
    const textContent = flattenTextContent(parsedShape.textProps);

    return { object: { text: { text: textContent, options: textOptions } }, warnings };
  }

  // Rectangle / generic shapes
  const rectObj = {
    x: pos.x,
    y: pos.y,
    w: pos.w,
    h: pos.h,
  };

  // Fill
  const { result: fillResult, warnings: fillWarnings } = resolveFill(parsedShape.fill, colorResolver);
  warnings.push(...fillWarnings);
  if (fillResult && fillResult.color) {
    rectObj.fill = fillResult;
  }

  // Line
  const lineResult = resolveLine(parsedShape.line, colorResolver);
  if (lineResult) {
    rectObj.line = lineResult;
  }

  // Rounded rect radius — extract from avLst if available
  if (parsedShape.geometry === 'roundRect') {
    const avLst = parsedShape.avLst;
    if (avLst?.adj != null) {
      // OOXML adj ranges 0-50000, representing 0-50% of min(w,h)
      // So adj/100000 gives the fraction of the shorter dimension
      const adjVal = Number(avLst.adj);
      const minDim = Math.min(pos.w || 1, pos.h || 1);
      rectObj.rectRadius = Math.round((adjVal / 100000) * minDim * 100) / 100 || 0.1;
    } else {
      rectObj.rectRadius = 0.1;
    }
  }

  if (parsedShape.rotation) rectObj.rotate = parsedShape.rotation;

  // Shadow (if spPr has effectLst — this comes through fill.element sometimes)
  // PptxGenJS shadow is set directly on the shape

  return { object: { rect: rectObj }, warnings };
}

/**
 * Build PptxGenJS text content from textProps paragraphs.
 *
 * @param {{ paragraphs: Array }} textProps
 * @returns {Array|string} PptxGenJS text array or plain string
 */
/**
 * Flatten text content to a plain string for PptxGenJS defineSlideMaster.
 *
 * PptxGenJS defineSlideMaster wraps text.text in [{ text: ... }] internally,
 * so array text runs get stringified to "[object Object]". This function
 * produces a flat string, using \n for line/paragraph breaks.
 *
 * Also strips template editing instructions (e.g. "[To edit, click View > ...]").
 *
 * @param {{ paragraphs?: Array, plainText?: string }} textProps
 * @returns {string} Flat text string
 */
function flattenTextContent(textProps) {
  if (!textProps?.paragraphs || textProps.paragraphs.length === 0) {
    return stripEditInstructions(textProps?.plainText || '');
  }

  const lines = [];
  for (const para of textProps.paragraphs) {
    if (!para.runs || para.runs.length === 0) {
      lines.push('');
      continue;
    }
    let paraText = '';
    for (const run of para.runs) {
      if (run.isBreak) {
        paraText += '\n';
      } else if (run.text) {
        paraText += run.text;
      }
    }
    lines.push(paraText);
  }

  return stripEditInstructions(lines.join('\n'));
}

/**
 * Strip template editing instructions from text content.
 * Corporate templates include "[To edit, click View > Slide Master > ...]"
 * which shouldn't appear in the generated output.
 *
 * @param {string} text
 * @returns {string}
 */
export function stripEditInstructions(text) {
  if (!text) return text;
  return text
    .split('\n')
    .filter((line) => !line.match(/^\[To edit[,.]|^\[Click .+Slide Master/i))
    .join('\n')
    .trim();
}

