/**
 * Shared parser utilities — common helpers used by both master.js and layout.js.
 */

import { emuToInches } from '../mapper/units.js';

/**
 * Normalize a value to an array.
 * fast-xml-parser returns a single child as an object, multiple as an array.
 * @param {*} el
 * @returns {Array}
 */
export function toArray(el) {
  return Array.isArray(el) ? el : el ? [el] : [];
}

/**
 * Extract the background element from a p:bg element.
 * Looks for p:bgPr (inline fill) or p:bgRef (theme reference).
 *
 * @param {object} bg - Parsed p:bg element
 * @returns {object|null} Raw fill element for downstream background mapper
 */
export function extractBackground(bg) {
  if (!bg) return null;

  if (bg['p:bgPr']) {
    return bg['p:bgPr'];
  }

  if (bg['p:bgRef']) {
    return { bgRef: bg['p:bgRef'] };
  }

  return null;
}

/**
 * Extract position from an a:xfrm element.
 *
 * @param {object} xfrm - Parsed a:xfrm element
 * @returns {{ x: number, y: number, w: number, h: number }|null}
 */
/**
 * Extract adjustment values (a:avLst) from a preset geometry element.
 *
 * @param {object} spPr - Parsed p:spPr element
 * @returns {object|null} Map of adjustment name → numeric value, or null
 */
export function extractAvLst(spPr) {
  const avLstEl = spPr?.['a:prstGeom']?.['a:avLst'];
  if (!avLstEl) return null;

  const gdItems = Array.isArray(avLstEl['a:gd']) ? avLstEl['a:gd'] : avLstEl['a:gd'] ? [avLstEl['a:gd']] : [];
  const avLst = {};
  for (const gd of gdItems) {
    const name = gd['@_name'];
    const fmla = gd['@_fmla'] || '';
    const match = fmla.match(/^val\s+(\d+)/);
    if (name && match) {
      avLst[name] = Number(match[1]);
    }
  }
  return Object.keys(avLst).length > 0 ? avLst : null;
}

export function extractPosition(xfrm) {
  if (!xfrm) return null;

  const off = xfrm['a:off'];
  const ext = xfrm['a:ext'];

  if (!off && !ext) return null;

  return {
    x: off ? emuToInches(Number(off['@_x'] || 0)) : 0,
    y: off ? emuToInches(Number(off['@_y'] || 0)) : 0,
    w: ext ? emuToInches(Number(ext['@_cx'] || 0)) : 0,
    h: ext ? emuToInches(Number(ext['@_cy'] || 0)) : 0,
  };
}
