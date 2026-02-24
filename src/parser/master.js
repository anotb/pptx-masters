/**
 * Slide master parser — extracts clrMap, background, shape tree,
 * text styles, and relationships from a slideMaster XML.
 */

import { parseRelationships } from './relationships.js';
import { parseClrMap } from './theme.js';
import { emuAngleToDegrees } from '../mapper/units.js';
import { extractTextProps } from './text.js';
import { toArray, extractBackground, extractPosition, extractAvLst } from './utils.js';

/**
 * Extract text styles from p:txStyles element.
 * Each style section (title, body, other) contains level paragraph
 * properties (a:lvl1pPr, a:lvl2pPr, etc.).
 *
 * @param {object} txStyles - Parsed p:txStyles element
 * @returns {{ title: object|null, body: object|null, other: object|null }}
 */
function extractTextStyles(txStyles) {
  if (!txStyles) {
    return { title: null, body: null, other: null };
  }

  return {
    title: txStyles['p:titleStyle'] || null,
    body: txStyles['p:bodyStyle'] || null,
    other: txStyles['p:otherStyle'] || null,
  };
}

/**
 * Extract shape tree children from p:spTree element.
 * Returns the raw child elements (p:sp, p:pic, p:grpSp, etc.).
 *
 * @param {object} spTree - Parsed p:spTree element
 * @returns {Array} Raw shape elements
 */
function extractShapeTree(spTree) {
  if (!spTree) return [];

  const shapes = [];

  // Collect all shape types
  const shapeTypes = ['p:sp', 'p:pic', 'p:grpSp', 'p:cxnSp', 'p:graphicFrame'];
  for (const type of shapeTypes) {
    const items = toArray(spTree[type]);
    for (const item of items) {
      shapes.push({ type, element: item });
    }
  }

  return shapes;
}

/**
 * Extract placeholder defaults from the master's shape tree.
 * Walks all p:sp shapes looking for those with p:ph elements,
 * extracts their xfrm position and text properties, and returns
 * them keyed by placeholder type (or `idx:<n>` for numbered ones).
 *
 * @param {Array} shapes - Shape tree items from extractShapeTree()
 * @param {object|null} colorResolver - Optional color resolver
 * @param {object|null} themeFonts - Optional theme fonts
 * @returns {Record<string, { position: object|null, textProps: object|null }>}
 */
function extractPlaceholderDefaults(shapes, colorResolver, themeFonts) {
  const defaults = {};

  for (const shape of shapes) {
    if (shape.type !== 'p:sp') continue;

    const sp = shape.element;
    if (!sp) continue;

    const nvSpPr = sp['p:nvSpPr'] || {};
    const nvPr = nvSpPr['p:nvPr'];
    if (!nvPr) continue;

    const ph = nvPr['p:ph'];
    if (!ph) continue;

    const phType = ph['@_type'] || undefined;
    const phIdx = ph['@_idx'] != null ? ph['@_idx'] : undefined;

    const spPr = sp['p:spPr'] || {};
    const position = extractPosition(spPr['a:xfrm']);

    // Extract text props if txBody exists
    let textProps = null;
    if (sp['p:txBody']) {
      textProps = extractTextProps(sp['p:txBody'], colorResolver, themeFonts);
    }

    // Key by type first, then by idx
    if (phType) {
      defaults[phType] = { position, textProps };
    }
    if (phIdx != null) {
      defaults[`idx:${phIdx}`] = { position, textProps };
    }
  }

  return defaults;
}

/**
 * Parse a slide master XML and extract its key components.
 *
 * @param {object} masterXml - Parsed slideMaster XML (from fast-xml-parser)
 * @param {object} masterRels - Parsed .rels XML for this master
 * @param {object} [pptxArchive] - PPTX archive (unused for now, reserved for future)
 * @param {{ colorResolver?: object, themeFonts?: object }} [options] - Optional resolvers
 * @returns {{
 *   clrMap: Record<string, string>,
 *   background: object|null,
 *   shapes: Array,
 *   textStyles: { title: object|null, body: object|null, other: object|null },
 *   relationships: Record<string, { type: string, target: string }>,
 *   placeholderDefaults: Record<string, { position: object|null, textProps: object|null }>,
 * }}
 */
export function parseSlideMaster(masterXml, masterRels, pptxArchive, options) {
  const master = masterXml?.['p:sldMaster'];
  if (!master) {
    return {
      clrMap: {},
      background: null,
      shapes: [],
      textStyles: { title: null, body: null, other: null },
      relationships: {},
      placeholderDefaults: {},
    };
  }

  // 1. Color map
  const clrMap = parseClrMap(master['p:clrMap']);

  // 2. Background from p:cSld -> p:bg
  const cSld = master['p:cSld'];
  const background = extractBackground(cSld?.['p:bg']);

  // 3. Shape tree children
  const shapes = extractShapeTree(cSld?.['p:spTree']);

  // 4. Text styles
  const textStyles = extractTextStyles(master['p:txStyles']);

  // 5. Relationships
  const relationships = parseRelationships(masterRels);

  // 6. Placeholder defaults (position + text props for inheritance)
  const colorResolver = options?.colorResolver || null;
  const themeFonts = options?.themeFonts || null;
  const placeholderDefaults = extractPlaceholderDefaults(shapes, colorResolver, themeFonts);

  // 7. Static (non-placeholder) shapes for showMasterSp inheritance
  const staticShapes = processNonPlaceholderShapes(shapes, colorResolver, themeFonts);

  return {
    clrMap,
    background,
    shapes,
    textStyles,
    relationships,
    placeholderDefaults,
    staticShapes,
  };
}

/**
 * Process non-placeholder shapes from the master's shape tree.
 * These are decorative elements (text boxes, images, lines) that layouts
 * inherit when showMasterSp is true (default).
 *
 * @param {Array} shapes - Shape tree items from extractShapeTree()
 * @param {object|null} colorResolver
 * @param {object|null} themeFonts
 * @returns {Array} Static shapes in the same format as layout.js staticShapes
 */
function processNonPlaceholderShapes(shapes, colorResolver, themeFonts) {
  const staticShapes = [];

  for (const shape of shapes) {
    if (shape.type === 'p:sp') {
      const sp = shape.element;
      if (!sp) continue;

      const nvSpPr = sp['p:nvSpPr'] || {};
      const nvPr = nvSpPr['p:nvPr'];

      // Skip placeholder shapes
      if (nvPr?.['p:ph']) continue;

      const cNvPr = nvSpPr['p:cNvPr'] || {};
      const spPr = sp['p:spPr'] || {};
      const txBody = sp['p:txBody'];

      const name = cNvPr['@_name'] || '';
      const xfrm = spPr['a:xfrm'];
      const position = extractPosition(xfrm);
      const rotation = xfrm?.['@_rot'] != null
        ? emuAngleToDegrees(Number(xfrm['@_rot']))
        : undefined;

      // Geometry
      const geometry = spPr['a:prstGeom']?.['@_prst'] || undefined;

      // Adjustment values (e.g., corner radius for roundRect)
      const avLst = extractAvLst(spPr);

      // Fill
      let fill = null;
      const fillTypes = ['a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:noFill'];
      for (const ft of fillTypes) {
        if (spPr[ft] != null) {
          fill = { type: ft.replace('a:', ''), element: spPr[ft] };
          break;
        }
      }

      // Line
      const line = spPr['a:ln'] || null;

      // Text props
      let textProps = null;
      if (txBody) {
        textProps = extractTextProps(txBody, colorResolver, themeFonts);
      }

      staticShapes.push({
        type: 'shape',
        name,
        position,
        rotation,
        geometry,
        fill,
        line,
        avLst,
        textProps,
        imageRef: undefined,
      });
    } else if (shape.type === 'p:pic') {
      const pic = shape.element;
      if (!pic) continue;

      const nvPicPr = pic['p:nvPicPr'] || {};
      const nvPr = nvPicPr['p:nvPr'];

      // Skip placeholder pics
      if (nvPr?.['p:ph']) continue;

      const cNvPr = nvPicPr['p:cNvPr'] || {};
      const spPr = pic['p:spPr'] || {};
      const blipFill = pic['p:blipFill'] || {};

      const name = cNvPr['@_name'] || '';
      const xfrm = spPr['a:xfrm'];
      const position = extractPosition(xfrm);
      const rotation = xfrm?.['@_rot'] != null
        ? emuAngleToDegrees(Number(xfrm['@_rot']))
        : undefined;
      const imageRef = blipFill['a:blip']?.['@_r:embed'] || undefined;

      staticShapes.push({
        type: 'picture',
        name,
        position,
        rotation,
        geometry: undefined,
        fill: null,
        line: null,
        textProps: null,
        imageRef,
      });
    }
  }

  return staticShapes;
}
