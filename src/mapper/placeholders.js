/**
 * Placeholder mapper — converts parsed OOXML placeholders into PptxGenJS
 * placeholder definitions for defineSlideMaster().
 */

import { resolveFill, resolveLine, mapTextPropsToOptions } from './shapes.js';

// Type mapping: OOXML placeholder type → PptxGenJS placeholder type
const TYPE_MAP = {
  title: 'title',
  ctrTitle: 'title',
  subTitle: 'body',
  body: 'body',
  obj: 'body',
  pic: 'pic',
  chart: 'chart',
  tbl: 'tbl',
  media: 'media',
  clipArt: 'pic',
  dgm: 'chart',
};

// Types that are handled by slideNumber mapper, not here
const SKIP_TYPES = new Set(['sldNum', 'ftr', 'dt', 'hdr']);

/**
 * Map a parsed placeholder into a PptxGenJS-compatible placeholder object.
 *
 * @param {object} parsedPlaceholder - Parsed placeholder from layout.js
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @param {{ heading: string, body: string }} themeFonts - Theme fonts
 * @returns {{ placeholder: object, warnings: string[] }|null}
 */
export function mapPlaceholder(parsedPlaceholder, colorResolver, themeFonts) {
  if (!parsedPlaceholder) return null;

  // Skip slide-number-related types
  if (SKIP_TYPES.has(parsedPlaceholder.type)) return null;

  const warnings = [];
  const phType = TYPE_MAP[parsedPlaceholder.type];

  if (!phType && parsedPlaceholder.type != null) {
    warnings.push(`Unknown placeholder type: ${parsedPlaceholder.type}`);
  }

  const pos = parsedPlaceholder.position || {};
  const options = {
    name: parsedPlaceholder.name || '',
    type: phType || 'body',
    x: pos.x ?? 0,
    y: pos.y ?? 0,
    w: pos.w ?? 0,
    h: pos.h ?? 0,
  };

  // Text styling from textProps (includes lstStyle fallback)
  if (parsedPlaceholder.textProps) {
    const textOpts = mapTextPropsToOptions(parsedPlaceholder.textProps);

    if (textOpts.fontFace) options.fontFace = textOpts.fontFace;
    if (textOpts.fontSize != null) options.fontSize = textOpts.fontSize;
    if (textOpts.color) options.color = textOpts.color;
    if (textOpts.bold) options.bold = textOpts.bold;
    if (textOpts.italic) options.italic = textOpts.italic;
    if (textOpts.align) options.align = textOpts.align;
    if (textOpts.valign) options.valign = textOpts.valign;
    if (textOpts.margin) options.margin = textOpts.margin;
    if (textOpts.lineSpacing != null) options.lineSpacing = textOpts.lineSpacing;
    if (textOpts.lineSpacingMultiple != null) options.lineSpacingMultiple = textOpts.lineSpacingMultiple;
    if (textOpts.paraSpaceBefore != null) options.paraSpaceBefore = textOpts.paraSpaceBefore;
    if (textOpts.paraSpaceAfter != null) options.paraSpaceAfter = textOpts.paraSpaceAfter;
  }

  // Center title default alignment — use center unless XML or lstStyle explicitly sets alignment
  if (parsedPlaceholder.type === 'ctrTitle') {
    const tp = parsedPlaceholder.textProps;
    const hasExplicitAlign = tp?.paragraphs?.[0]?._explicitAlign
      || tp?.lstStyleProps?.[1]?.align
      || tp?.lstStyleProps?.[0]?.align;
    if (!hasExplicitAlign) {
      options.align = 'center';
    }
  }

  // Shape properties (fill, line)
  if (parsedPlaceholder.shapeProps) {
    const { fill, line } = parsedPlaceholder.shapeProps;

    if (fill) {
      const { result: fillResult, warnings: fillWarnings } = resolveFill(fill, colorResolver);
      warnings.push(...fillWarnings);
      if (fillResult && fillResult.color) {
        options.fill = fillResult;
      }
    }

    if (line) {
      const lineResult = resolveLine(line, colorResolver);
      if (lineResult) {
        options.line = lineResult;
      }
    }
  }

  // Rotation
  if (parsedPlaceholder.rotation) {
    options.rotate = parsedPlaceholder.rotation;
  }

  return { placeholder: { options }, warnings };
}
