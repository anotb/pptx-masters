/**
 * Slide layout parser — extracts layout metadata, placeholders, static shapes,
 * and background from a slideLayout XML.
 *
 * Also exports parsePresentation() for extracting slide dimensions.
 */

import { parseRelationships } from './relationships.js';
import { extractTextProps } from './text.js';
import { emuToInches, emuAngleToDegrees } from '../mapper/units.js';
import { toArray, extractBackground, extractPosition, extractAvLst } from './utils.js';

/**
 * Extract color map override from p:clrMapOvr element.
 *
 * @param {object} clrMapOvr - Parsed p:clrMapOvr element
 * @returns {null|object} null if inheriting master, otherwise override map
 */
function extractClrMapOverride(clrMapOvr) {
  if (!clrMapOvr) return null;

  // a:masterClrMapping means "inherit master's clrMap"
  if (clrMapOvr['a:masterClrMapping'] != null) {
    return null;
  }

  // a:overrideClrMapping contains the override attributes
  const override = clrMapOvr['a:overrideClrMapping'];
  if (!override) return null;

  const map = {};
  for (const [key, value] of Object.entries(override)) {
    if (key.startsWith('@_')) {
      map[key.slice(2)] = value;
    }
  }
  return Object.keys(map).length > 0 ? map : null;
}

/**
 * Extract rotation from an a:xfrm element.
 *
 * @param {object} xfrm - Parsed a:xfrm element
 * @returns {number|undefined} Rotation in degrees
 */
function extractRotation(xfrm) {
  if (!xfrm || xfrm['@_rot'] == null) return undefined;
  return emuAngleToDegrees(Number(xfrm['@_rot']));
}

/**
 * Extract shape properties (geometry, fill, line) from p:spPr.
 *
 * @param {object} spPr - Parsed p:spPr element
 * @returns {{ geometry: string|undefined, fill: object|null, line: object|null }}
 */
function extractShapeProps(spPr) {
  if (!spPr) {
    return { geometry: undefined, fill: null, line: null };
  }

  // Geometry
  const geometry = spPr['a:prstGeom']?.['@_prst'] || undefined;

  // Adjustment values (e.g., corner radius for roundRect)
  const avLst = extractAvLst(spPr);

  // Fill — find whichever fill type is present
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

  return { geometry, fill, line, avLst };
}

/**
 * Check if a shape element has a placeholder marker.
 *
 * @param {object} nvPr - The p:nvPr element from nvSpPr/nvPicPr
 * @returns {{ type: string|undefined, idx: number|string|undefined, sz: string|undefined }|null}
 */
function extractPlaceholder(nvPr) {
  if (!nvPr) return null;

  const ph = nvPr['p:ph'];
  if (!ph) return null;

  return {
    type: ph['@_type'] || undefined,
    idx: ph['@_idx'] != null ? ph['@_idx'] : undefined,
    sz: ph['@_sz'] || undefined,
  };
}

/**
 * Process a p:sp (shape) element.
 *
 * @param {object} sp - Parsed p:sp element
 * @param {object|null} colorResolver - Optional color resolver
 * @param {object|null} themeFonts - Optional theme fonts
 * @param {Record<string, { position: object|null, textProps: object|null }>|null} masterDefaults - Master placeholder defaults for inheritance
 * @returns {{ isPlaceholder: boolean, data: object }}
 */
function processShape(sp, colorResolver, themeFonts, masterDefaults) {
  const nvSpPr = sp['p:nvSpPr'] || {};
  const cNvPr = nvSpPr['p:cNvPr'] || {};
  const nvPr = nvSpPr['p:nvPr'];
  const spPr = sp['p:spPr'] || {};
  const txBody = sp['p:txBody'];

  const name = cNvPr['@_name'] || '';
  const ph = extractPlaceholder(nvPr);
  const xfrm = spPr['a:xfrm'];
  let position = extractPosition(xfrm);
  const rotation = extractRotation(xfrm);
  const { geometry, fill, line, avLst } = extractShapeProps(spPr);

  // Extract text props if txBody exists
  let textProps = null;
  if (txBody) {
    textProps = extractTextProps(txBody, colorResolver, themeFonts);
  }

  if (ph) {
    // Inherit position from master if not defined locally
    if (!position && masterDefaults) {
      const masterPh = (ph.type && masterDefaults[ph.type])
        || (ph.idx != null && masterDefaults[`idx:${ph.idx}`])
        || null;
      if (masterPh?.position) {
        position = masterPh.position;
      }
      // Also inherit text props if not present locally
      if (!textProps && masterPh?.textProps) {
        textProps = masterPh.textProps;
      }
    }

    // Placeholder shape
    return {
      isPlaceholder: true,
      data: {
        type: ph.type,
        idx: ph.idx,
        name,
        position,
        rotation,
        textProps,
        shapeProps: { geometry, fill, line, avLst },
      },
    };
  }

  // Static decoration shape
  return {
    isPlaceholder: false,
    data: {
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
    },
  };
}

/**
 * Process a p:pic (picture) element.
 *
 * @param {object} pic - Parsed p:pic element
 * @returns {{ isPlaceholder: boolean, data: object }}
 */
function processPicture(pic) {
  const nvPicPr = pic['p:nvPicPr'] || {};
  const cNvPr = nvPicPr['p:cNvPr'] || {};
  const nvPr = nvPicPr['p:nvPr'];
  const spPr = pic['p:spPr'] || {};
  const blipFill = pic['p:blipFill'] || {};

  const name = cNvPr['@_name'] || '';
  const ph = extractPlaceholder(nvPr);
  const xfrm = spPr['a:xfrm'];
  const position = extractPosition(xfrm);
  const rotation = extractRotation(xfrm);

  // Image reference
  const imageRef = blipFill['a:blip']?.['@_r:embed'] || undefined;

  const { geometry, fill, line } = extractShapeProps(spPr);

  if (ph) {
    return {
      isPlaceholder: true,
      data: {
        type: ph.type || 'pic',
        idx: ph.idx,
        name,
        position,
        rotation,
        textProps: null,
        shapeProps: { geometry, fill, line },
        imageRef,
      },
    };
  }

  return {
    isPlaceholder: false,
    data: {
      type: 'picture',
      name,
      position,
      rotation,
      geometry,
      fill,
      line,
      textProps: null,
      imageRef,
    },
  };
}

/**
 * Parse a slide layout XML and extract its components.
 *
 * @param {object} layoutXml - Parsed slideLayout XML (from fast-xml-parser)
 * @param {object} layoutRels - Parsed .rels XML for this layout
 * @param {object} [pptxArchive] - PPTX archive (unused for now)
 * @param {{ colorResolver?: object, themeFonts?: object, masterDefaults?: Record<string, object> }} [options] - Optional resolvers
 * @returns {{
 *   name: string,
 *   type: string|undefined,
 *   clrMapOverride: object|null,
 *   background: object|null,
 *   placeholders: Array,
 *   staticShapes: Array,
 *   warnings: string[],
 *   relationships: Record<string, { type: string, target: string }>,
 * }}
 */
export function parseSlideLayout(layoutXml, layoutRels, pptxArchive, options) {
  const layout = layoutXml?.['p:sldLayout'];
  if (!layout) {
    return {
      name: '',
      type: undefined,
      clrMapOverride: null,
      background: null,
      placeholders: [],
      staticShapes: [],
      warnings: [],
      relationships: {},
    };
  }

  const colorResolver = options?.colorResolver || null;
  const themeFonts = options?.themeFonts || null;
  const masterDefaults = options?.masterDefaults || null;
  const cSld = layout['p:cSld'] || {};

  // 1. Layout name — check p:sldLayout @_name first, then p:cSld @_name
  const name = layout['@_name'] || cSld['@_name'] || '';

  // 2. Layout type
  const type = layout['@_type'] || undefined;

  // Show master shapes — true when not explicitly set to "0"
  const showMasterSp = layout['@_showMasterSp'] !== '0';

  // 3. Color map override
  const clrMapOverride = extractClrMapOverride(layout['p:clrMapOvr']);

  // 4. Background
  const background = extractBackground(cSld['p:bg']);

  // 5. Walk shape tree
  const spTree = cSld['p:spTree'] || {};
  const placeholders = [];
  const staticShapes = [];
  const warnings = [];

  // Process p:sp elements
  const spElements = toArray(spTree['p:sp']);
  for (const sp of spElements) {
    const result = processShape(sp, colorResolver, themeFonts, masterDefaults);
    if (result.isPlaceholder) {
      placeholders.push(result.data);
    } else {
      staticShapes.push(result.data);
    }
  }

  // Process p:pic elements
  const picElements = toArray(spTree['p:pic']);
  for (const pic of picElements) {
    const result = processPicture(pic);
    if (result.isPlaceholder) {
      placeholders.push(result.data);
    } else {
      staticShapes.push(result.data);
    }
  }

  // Log warnings for unsupported shape types
  const grpSpElements = toArray(spTree['p:grpSp']);
  if (grpSpElements.length > 0) {
    warnings.push(`Found ${grpSpElements.length} grouped shape(s) (not supported in v1)`);
  }

  const cxnSpElements = toArray(spTree['p:cxnSp']);
  if (cxnSpElements.length > 0) {
    warnings.push(`Found ${cxnSpElements.length} connection shape(s) (not supported in v1)`);
  }

  const graphicFrameElements = toArray(spTree['p:graphicFrame']);
  if (graphicFrameElements.length > 0) {
    warnings.push(`Found ${graphicFrameElements.length} graphic frame(s) (tables/charts, not supported in v1)`);
  }

  // 6. Relationships
  const relationships = parseRelationships(layoutRels);

  return {
    name,
    type,
    showMasterSp,
    clrMapOverride,
    background,
    placeholders,
    staticShapes,
    warnings,
    relationships,
  };
}

/**
 * Parse presentation.xml to extract slide dimensions.
 *
 * @param {object} presentationXml - Parsed presentation.xml
 * @returns {{ width: number, height: number }} Dimensions in inches
 */
export function parsePresentation(presentationXml) {
  const pres = presentationXml?.['p:presentation'];
  if (!pres) {
    return { width: 10, height: 7.5 }; // Default 4:3 dimensions
  }

  const sldSz = pres['p:sldSz'];
  if (!sldSz) {
    return { width: 10, height: 7.5 };
  }

  const width = emuToInches(Number(sldSz['@_cx'] || 9144000));
  const height = emuToInches(Number(sldSz['@_cy'] || 6858000));

  return { width, height };
}
