/**
 * Slide number and footer mapper — extracts slide number position/styling
 * and footer/header/date text objects from layout placeholders.
 */

import { mapTextPropsToOptions, stripEditInstructions } from './shapes.js';

// Placeholder types handled by this mapper
const SLIDE_NUM_TYPES = new Set(['sldNum', 'ftr', 'dt', 'hdr']);

/**
 * Extract slide number and footer elements from layout placeholders.
 *
 * @param {Array} placeholders - All placeholders from layout parser
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @param {{ heading: string, body: string }} themeFonts - Theme fonts
 * @returns {{ slideNumber: object|null, footerObjects: Array }}
 */
export function mapSlideNumberAndFooters(placeholders, colorResolver, themeFonts) {
  if (!placeholders || placeholders.length === 0) {
    return { slideNumber: null, footerObjects: [] };
  }

  let slideNumber = null;
  const footerObjects = [];

  for (const ph of placeholders) {
    if (!ph) continue;

    // Slide number placeholder
    if (ph.type === 'sldNum') {
      slideNumber = extractSlideNumberProps(ph);
      continue;
    }

    // Footer, date, header placeholders
    if (ph.type === 'ftr' || ph.type === 'dt' || ph.type === 'hdr') {
      const textContent = ph.textProps?.plainText || '';
      const pos = ph.position;
      // Skip empty footer/date objects with no content and zero/missing position
      if (!textContent && (!pos || (pos.x === 0 && pos.y === 0 && pos.w === 0 && pos.h === 0))) {
        continue;
      }
      const textObj = extractFooterTextObject(ph);
      if (textObj) {
        footerObjects.push(textObj);
      }
      continue;
    }

    // Detect copyright text in non-typed shapes
    if (!ph.type && ph.textProps?.plainText) {
      const text = ph.textProps.plainText.toLowerCase();
      if (text.includes('\u00a9') || text.includes('copyright')) {
        const textObj = extractFooterTextObject(ph);
        if (textObj) {
          footerObjects.push(textObj);
        }
      }
    }
  }

  // Null out slide number if it has zero size (no valid position)
  if (slideNumber && slideNumber.w === 0 && slideNumber.h === 0) {
    slideNumber = null;
  }

  return { slideNumber, footerObjects };
}

/**
 * Extract slide number properties from a sldNum placeholder.
 *
 * @param {object} ph - Parsed sldNum placeholder
 * @returns {object} Position and styling for PptxGenJS slideNumber
 */
function extractSlideNumberProps(ph) {
  const pos = ph.position || {};
  const result = {
    x: pos.x ?? 0,
    y: pos.y ?? 0,
    w: pos.w ?? 0,
    h: pos.h ?? 0,
  };

  // Zero-size check before normalization
  if (result.w === 0 && result.h === 0) return result;

  // Extract text styling
  if (ph.textProps) {
    const textOpts = mapTextPropsToOptions(ph.textProps);
    if (textOpts.fontFace) result.fontFace = textOpts.fontFace;
    if (textOpts.fontSize != null) result.fontSize = textOpts.fontSize;
    if (textOpts.color) result.color = textOpts.color;
    if (textOpts.align) result.align = textOpts.align;
  }

  // Normalize height to match footer text boxes (spAutoFit shapes have tiny h)
  normalizeSlideNumberHeight(result);

  return result;
}

/**
 * Check if a static shape contains a slidenum field and extract it
 * as a PptxGenJS slideNumber object. Some corporate templates put
 * slide numbers in non-placeholder text boxes (a:fld type="slidenum")
 * rather than p:ph type="sldNum" placeholders.
 *
 * @param {object} shape - Parsed static shape from layout/master
 * @returns {object|null} slideNumber object or null if not a slide number shape
 */
export function extractSlideNumberFromShape(shape) {
  if (!shape?.textProps?.paragraphs) return null;

  // Check all runs across all paragraphs for a slidenum field
  let hasSlideNumField = false;
  let hasOtherContent = false;

  for (const para of shape.textProps.paragraphs) {
    if (!para.runs) continue;
    for (const run of para.runs) {
      if (run.isField && run.fieldType === 'slidenum') {
        hasSlideNumField = true;
      } else if (run.isBreak) {
        // line breaks are fine
      } else if (run.text && run.text.trim()) {
        hasOtherContent = true;
      }
    }
  }

  // Only treat as slide number if it ONLY contains a slidenum field
  if (!hasSlideNumField || hasOtherContent) return null;

  const pos = shape.position || {};
  const result = {
    x: pos.x ?? 0,
    y: pos.y ?? 0,
    w: pos.w ?? 0,
    h: pos.h ?? 0,
  };

  // Don't return zero-size slide numbers (check before normalization)
  if (result.w === 0 && result.h === 0) return null;

  // Extract text styling
  if (shape.textProps) {
    const textOpts = mapTextPropsToOptions(shape.textProps);
    if (textOpts.fontFace) result.fontFace = textOpts.fontFace;
    if (textOpts.fontSize != null) result.fontSize = textOpts.fontSize;
    if (textOpts.color) result.color = textOpts.color;
    if (textOpts.align) result.align = textOpts.align;
    if (textOpts.margin) result.margin = textOpts.margin;
  }

  // Normalize height to match footer text boxes
  normalizeSlideNumberHeight(result);

  return result;
}

/**
 * Normalize slideNumber height — original templates often use spAutoFit
 * on slide number shapes, giving them a tiny h (e.g., 0.13"). PptxGenJS
 * slideNumber doesn't auto-fit, so ensure height is at least enough for
 * the font to render cleanly and align with neighboring footer text.
 *
 * @param {object} result - Slide number props (mutated in place)
 */
function normalizeSlideNumberHeight(result) {
  const fontSize = result.fontSize || 8;
  // Minimum height: ~2.5x the font size in inches (matches typical footer text box heights)
  const minHeight = Math.round(((fontSize * 2.5) / 72) * 10000) / 10000;
  if (result.h < minHeight) {
    result.h = minHeight;
  }
}

/**
 * Extract a footer/header/date placeholder as a PptxGenJS text object.
 *
 * @param {object} ph - Parsed placeholder
 * @returns {object|null} PptxGenJS text object for objects[] array
 */
function extractFooterTextObject(ph) {
  const pos = ph.position || {};
  const textContent = stripEditInstructions(ph.textProps?.plainText || '');

  const options = {
    x: pos.x ?? 0,
    y: pos.y ?? 0,
    w: pos.w ?? 0,
    h: pos.h ?? 0,
  };

  // Extract text styling
  if (ph.textProps) {
    const textOpts = mapTextPropsToOptions(ph.textProps);
    if (textOpts.fontFace) options.fontFace = textOpts.fontFace;
    if (textOpts.fontSize != null) options.fontSize = textOpts.fontSize;
    if (textOpts.color) options.color = textOpts.color;
    if (textOpts.align) options.align = textOpts.align;
    if (textOpts.margin) options.margin = textOpts.margin;
  }

  // Always top-align footer elements for consistent baseline with slideNumber
  options.valign = 'top';

  return {
    text: { text: textContent, options },
  };
}
