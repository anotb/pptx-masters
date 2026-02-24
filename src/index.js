/**
 * pptx-masters — library exports.
 *
 * Provides the main extract() function and re-exports useful internals
 * for programmatic usage.
 */

import { resolve, basename } from 'path';
import { extractPptx } from './parser/zip.js';
import { parseTheme, parseClrMap } from './parser/theme.js';
import { createColorResolver } from './parser/colors.js';
import { parseRelationships, resolveRelPath } from './parser/relationships.js';
import { parseSlideMaster } from './parser/master.js';
import { parseSlideLayout, parsePresentation } from './parser/layout.js';
import { mapShape } from './mapper/shapes.js';
import { mapPlaceholder } from './mapper/placeholders.js';
import { mapBackground } from './mapper/backgrounds.js';
import { mapSlideNumberAndFooters, extractSlideNumberFromShape } from './mapper/slideNumber.js';
import { generateMastersCode, generateThemeJson, toUpperSnakeCase, generateAgentInstructions } from './generator/code.js';
import { generateReport } from './generator/report.js';
import { generatePreview } from './generator/preview.js';

/**
 * Extract slide masters from a PPTX/POTX file.
 *
 * @param {string} inputPath - Path to the .pptx or .potx file
 * @param {{ layouts?: string[], output?: string, preview?: boolean, report?: boolean }} [options]
 * @returns {Promise<{
 *   masterData: Array<object>,
 *   themeColors: Record<string, string>,
 *   themeFonts: { heading: string, body: string },
 *   dimensions: { width: number, height: number },
 *   warnings: string[],
 *   layouts: Array<object>,
 *   mediaFiles: Array<{ archivePath: string, filename: string }>,
 * }>}
 */
export async function extract(inputPath, options = {}) {
  const absPath = resolve(inputPath);
  const templateName = basename(absPath);
  const allWarnings = [];

  // 1. Extract PPTX archive
  const archive = await extractPptx(absPath);
  const allFiles = archive.listFiles();

  // 2. Parse theme — resolve path from presentation.xml.rels, fallback to theme1.xml
  let themePath = 'ppt/theme/theme1.xml';
  try {
    const presRelsXml = await archive.getXml('ppt/_rels/presentation.xml.rels');
    const presRels = parseRelationships(presRelsXml);
    const themeRel = Object.values(presRels).find((r) => r.type === 'theme');
    if (themeRel) {
      themePath = resolveRelPath('ppt/presentation.xml', themeRel.target);
    }
  } catch {
    // No presentation.xml.rels — use default
  }
  const themeXml = await archive.getXml(themePath);
  const { colors: themeColors, fonts: themeFonts } = parseTheme(themeXml);

  // 3. Parse presentation dimensions
  const presXml = await archive.getXml('ppt/presentation.xml');
  const dimensions = parsePresentation(presXml);

  // 4. Find and parse slide masters
  const masterFiles = allFiles
    .filter((f) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f));

  // First pass: parse masters without colorResolver (need clrMap first)
  const masters = [];
  for (const masterFile of masterFiles) {
    const masterXml = await archive.getXml(masterFile);
    const masterRelsPath = masterFile.replace(
      'ppt/slideMasters/',
      'ppt/slideMasters/_rels/',
    ) + '.rels';

    let masterRels = {};
    try {
      masterRels = await archive.getXml(masterRelsPath);
    } catch {
      // No rels file — use empty
    }

    const master = parseSlideMaster(masterXml, masterRels);
    masters.push({ ...master, _file: masterFile });
  }

  // 5. Build color resolver from theme + first master's clrMap
  const defaultClrMap = masters.length > 0 ? masters[0].clrMap : {};
  const defaultColorResolver = createColorResolver(themeColors, defaultClrMap, themeFonts);

  // Re-parse placeholder defaults now that we have color resolvers
  for (const master of masters) {
    const masterResolver = createColorResolver(themeColors, master.clrMap, themeFonts);
    const masterXml = await archive.getXml(master._file);
    const masterRelsPath = master._file.replace(
      'ppt/slideMasters/',
      'ppt/slideMasters/_rels/',
    ) + '.rels';
    let masterRels = {};
    try {
      masterRels = await archive.getXml(masterRelsPath);
    } catch {
      // No rels file — use empty
    }
    const reParsed = parseSlideMaster(masterXml, masterRels, null, {
      colorResolver: masterResolver,
      themeFonts,
    });
    master.placeholderDefaults = reParsed.placeholderDefaults;
    master.staticShapes = reParsed.staticShapes;
    master.relationships = reParsed.relationships;
  }

  // 6. Find and parse slide layouts
  const layoutFiles = allFiles
    .filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f));

  const parsedLayouts = [];
  for (const layoutFile of layoutFiles) {
    const layoutXml = await archive.getXml(layoutFile);
    const layoutRelsPath = layoutFile.replace(
      'ppt/slideLayouts/',
      'ppt/slideLayouts/_rels/',
    ) + '.rels';

    let layoutRels = {};
    try {
      layoutRels = await archive.getXml(layoutRelsPath);
    } catch {
      // No rels file — use empty
    }

    // Resolve which master this layout belongs to via relationships
    const layoutRelsParsed = parseRelationships(layoutRels);
    const masterRel = Object.values(layoutRelsParsed).find(
      (r) => r.type === 'slideMaster',
    );
    const masterFile = masterRel
      ? resolveRelPath(layoutFile, masterRel.target)
      : null;
    const masterForLayout = masters.find((m) => m._file === masterFile)
      || masters[0]
      || null;

    // Pre-extract clrMapOverride from raw XML before full parsing
    const layoutRoot = layoutXml?.['p:sldLayout'];
    const clrMapOvr = layoutRoot?.['p:clrMapOvr'];
    let layoutClrMapOverride = null;
    if (clrMapOvr && clrMapOvr['a:masterClrMapping'] == null) {
      const overrideEl = clrMapOvr['a:overrideClrMapping'];
      if (overrideEl) {
        layoutClrMapOverride = {};
        for (const [key, value] of Object.entries(overrideEl)) {
          if (key.startsWith('@_')) {
            layoutClrMapOverride[key.slice(2)] = value;
          }
        }
      }
    }

    // Build color resolver for this layout's master, applying clrMapOverride if present
    let layoutClrMap = masterForLayout?.clrMap || {};
    if (layoutClrMapOverride) {
      layoutClrMap = { ...layoutClrMap, ...layoutClrMapOverride };
    }
    const layoutColorResolver = masterForLayout
      ? createColorResolver(themeColors, layoutClrMap, themeFonts)
      : defaultColorResolver;

    const layout = parseSlideLayout(layoutXml, layoutRels, null, {
      colorResolver: layoutColorResolver,
      themeFonts,
      masterDefaults: masterForLayout?.placeholderDefaults || null,
    });

    parsedLayouts.push({
      ...layout,
      _file: layoutFile,
      _masterFile: masterForLayout?._file || null,
    });
  }

  // 7. Filter layouts if --layouts specified (supports names or 1-based numbers)
  let filteredLayouts = parsedLayouts;
  if (options.layouts && options.layouts.length > 0) {
    const numericFilters = [];
    const nameFilters = [];
    for (const f of options.layouts) {
      const n = parseInt(f, 10);
      if (!isNaN(n) && String(n) === f.trim()) {
        numericFilters.push(n);
      } else {
        nameFilters.push(f.toLowerCase());
      }
    }

    filteredLayouts = parsedLayouts.filter((layout, idx) => {
      // Match by 1-based number
      if (numericFilters.includes(idx + 1)) return true;
      // Match by name (partial, case-insensitive)
      if (nameFilters.length > 0) {
        const lowerName = layout.name.toLowerCase();
        return nameFilters.some((f) => lowerName.includes(f));
      }
      return false;
    });
  }

  // 8. Map each layout to PptxGenJS master data
  const masterData = [];
  const mediaFiles = [];

  for (const layout of filteredLayouts) {
    const layoutWarnings = [...layout.warnings];

    // Resolve layout's master for background inheritance and color resolver
    const masterForLayout = masters.find((m) => m._file === layout._masterFile)
      || masters[0]
      || null;
    let masterClrMap = masterForLayout?.clrMap || {};
    if (layout.clrMapOverride) {
      masterClrMap = { ...masterClrMap, ...layout.clrMapOverride };
    }
    const layoutColorResolver = masterForLayout
      ? createColorResolver(themeColors, masterClrMap, themeFonts)
      : defaultColorResolver;

    // Background — inherit from master if layout has none
    // Use the correct relationships for image resolution:
    // layout.relationships when using layout's bg, master's when inheriting master's bg
    const usingMasterBg = !layout.background && masterForLayout?.background;
    const bgElement = layout.background
      || (masterForLayout?.background ?? null);
    const bgRelationships = usingMasterBg
      ? masterForLayout.relationships
      : layout.relationships;
    const { background, warnings: bgWarnings } = mapBackground(
      bgElement,
      layoutColorResolver,
      bgRelationships,
    );
    layoutWarnings.push(...bgWarnings);

    // Slide number and footers
    let { slideNumber, footerObjects } = mapSlideNumberAndFooters(
      layout.placeholders,
      layoutColorResolver,
      themeFonts,
    );

    // Placeholders (non-footer/sldNum)
    const placeholderObjects = [];
    for (const ph of layout.placeholders) {
      const result = mapPlaceholder(ph, layoutColorResolver, themeFonts);
      if (result) {
        const { warnings: phWarnings, ...phObject } = result;
        placeholderObjects.push(phObject);
        layoutWarnings.push(...phWarnings);
      }
    }

    // Static shapes from layout
    const shapeObjects = [];
    for (const shape of layout.staticShapes) {
      // Detect non-placeholder slidenum field shapes
      const slideNumFromShape = extractSlideNumberFromShape(shape);
      if (slideNumFromShape) {
        if (!slideNumber) slideNumber = slideNumFromShape;
        continue;
      }

      const { object, warnings: shapeWarnings } = mapShape(
        shape,
        layoutColorResolver,
        themeFonts,
        layout.relationships,
      );
      if (object) {
        shapeObjects.push(object);
      }
      layoutWarnings.push(...shapeWarnings);
    }

    // Inherit master static shapes when showMasterSp is true (default)
    const masterStaticShapes = [];
    if (layout.showMasterSp !== false && masterForLayout?.staticShapes) {
      for (const shape of masterForLayout.staticShapes) {
        // Detect non-placeholder slidenum field shapes from master
        const slideNumFromShape = extractSlideNumberFromShape(shape);
        if (slideNumFromShape) {
          if (!slideNumber) slideNumber = slideNumFromShape;
          continue;
        }

        const { object, warnings: shapeWarnings } = mapShape(
          shape,
          layoutColorResolver,
          themeFonts,
          masterForLayout.relationships, // Use master's relationships for image resolution
        );
        if (object) {
          masterStaticShapes.push(object);
        }
        layoutWarnings.push(...shapeWarnings);
      }
    }

    // Collect media references
    if (layout.relationships) {
      for (const [, rel] of Object.entries(layout.relationships)) {
        if (rel.type === 'image') {
          const archivePath = resolveRelPath(layout._file, rel.target);
          const filename = rel.target.split('/').pop();
          if (!mediaFiles.some((m) => m.archivePath === archivePath)) {
            mediaFiles.push({ archivePath, filename });
          }
        }
      }
    }

    allWarnings.push(...layoutWarnings.map((w) => `[${layout.name}] ${w}`));

    // Collect master media files when inheriting shapes
    if (masterStaticShapes.length > 0 && masterForLayout?.relationships) {
      for (const [, rel] of Object.entries(masterForLayout.relationships)) {
        if (rel.type === 'image') {
          const archivePath = resolveRelPath(masterForLayout._file, rel.target);
          const filename = rel.target.split('/').pop();
          if (!mediaFiles.some((m) => m.archivePath === archivePath)) {
            mediaFiles.push({ archivePath, filename });
          }
        }
      }
    }

    // Deduplicate text/image objects — some templates have duplicate footer shapes
    // in the layout XML, plus inherited master shapes. Layout shapes take priority
    // (they may override master colors, e.g. white text on dark backgrounds).
    const allObjects = [...shapeObjects, ...footerObjects, ...placeholderObjects, ...masterStaticShapes];
    const deduped = deduplicateObjects(allObjects);

    // Clean up footer-zone text: strip paraSpaceBefore/After and ensure color
    cleanupFooterZone(deduped, dimensions, background);

    // Fill missing placeholder colors from master txStyles defaults.
    // OOXML placeholders inherit text color from p:txStyles (title/body/other)
    // which PptxGenJS doesn't replicate — resolve and apply as fallback.
    const defaultTextColor = resolveDefaultTextColor(layoutColorResolver, masterForLayout);
    applyPlaceholderColorDefaults(deduped, defaultTextColor);

    // Ensure slideNumber has a color — some templates omit it, relying on
    // theme/master inheritance that PptxGenJS doesn't replicate
    if (slideNumber && !slideNumber.color) {
      slideNumber.color = isDarkBackground(background) ? 'FFFFFF' : '000000';
    }

    masterData.push({
      name: layout.name,
      title: toUpperSnakeCase(layout.name),
      background,
      slideNumber,
      objects: deduped,
    });
  }

  return {
    masterData,
    themeColors,
    themeFonts,
    dimensions,
    warnings: allWarnings,
    layouts: filteredLayouts,
    mediaFiles,
    templateName,
  };
}

/**
 * Deduplicate objects array — some templates have duplicate static shapes
 * (e.g., footer text appearing multiple times in layout XML and inherited
 * from master). Removes duplicates by comparing text content + position.
 * Placeholders are never deduped (they may have same text but different names).
 *
 * @param {Array<object>} objects
 * @returns {Array<object>}
 */
function deduplicateObjects(objects) {
  const seen = new Set();
  const result = [];

  for (const obj of objects) {
    // Only dedup text and image objects — never placeholders
    if (obj.placeholder) {
      result.push(obj);
      continue;
    }

    // Build a dedup key from object type + content + approximate position
    let key = null;
    if (obj.text) {
      const t = obj.text.text || '';
      const o = obj.text.options || {};
      key = `text:${t}:${round2(o.x)}:${round2(o.y)}:${round2(o.w)}:${round2(o.h)}`;
    } else if (obj.image) {
      key = `image:${obj.image.path || ''}:${round2(obj.image.x)}:${round2(obj.image.y)}:${round2(obj.image.w)}:${round2(obj.image.h)}`;
    } else if (obj.rect) {
      key = `rect:${round2(obj.rect.x)}:${round2(obj.rect.y)}:${round2(obj.rect.w)}:${round2(obj.rect.h)}`;
    }

    if (key && seen.has(key)) continue;
    if (key) seen.add(key);
    result.push(obj);
  }

  return result;
}

/**
 * Check if a background is dark (for choosing contrasting text color).
 * @param {object|null} background
 * @returns {boolean}
 */
function isDarkBackground(background) {
  if (!background?.color) return false;
  const hex = background.color.replace('#', '');
  if (hex.length !== 6) return false;
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  // Relative luminance threshold
  return (r * 0.299 + g * 0.587 + b * 0.114) < 128;
}

/**
 * Clean up footer-zone text objects: strip paraSpaceBefore/After (breaks
 * alignment with slideNumber) and ensure color is set (some templates
 * rely on theme inheritance that PptxGenJS doesn't replicate).
 *
 * @param {Array<object>} objects - Objects array (mutated in place)
 * @param {{ width: number, height: number }} dimensions - Slide dimensions
 * @param {object|null} background - Slide background for color fallback
 */
function cleanupFooterZone(objects, dimensions, background) {
  const footerThreshold = (dimensions?.height || 7.5) * 0.9;
  const fallbackColor = isDarkBackground(background) ? 'FFFFFF' : '000000';

  for (const obj of objects) {
    if (!obj.text?.options) continue;
    const opts = obj.text.options;
    if ((opts.y ?? 0) < footerThreshold) continue;

    delete opts.paraSpaceBefore;
    delete opts.paraSpaceAfter;

    if (!opts.color) {
      opts.color = fallbackColor;
    }
  }
}

/**
 * Resolve the default text color from the master's txStyles.
 * OOXML masters define p:txStyles with default colors for title, body, other.
 * These are typically scheme:tx1 which maps through clrMap to a theme color.
 *
 * @param {{ resolve: Function }} colorResolver
 * @param {object|null} masterForLayout
 * @returns {string|null} Hex color string or null
 */
function resolveDefaultTextColor(colorResolver, masterForLayout) {
  if (!masterForLayout?.textStyles) return null;

  // Try title style first (most common), then body, then other
  for (const style of [masterForLayout.textStyles.title, masterForLayout.textStyles.body, masterForLayout.textStyles.other]) {
    if (!style) continue;
    const lvl1 = style['a:lvl1pPr'];
    const fill = lvl1?.['a:defRPr']?.['a:solidFill'];
    if (fill) {
      const resolved = colorResolver.resolve(fill);
      if (resolved?.color) return resolved.color;
    }
  }

  return null;
}

/**
 * Apply default text color to placeholders that don't have one.
 * PptxGenJS doesn't inherit from master txStyles, so we need to
 * bake the default color into each placeholder explicitly.
 *
 * @param {Array<object>} objects - Objects array
 * @param {string|null} defaultColor - Fallback hex color
 */
function applyPlaceholderColorDefaults(objects, defaultColor) {
  if (!defaultColor) return;

  for (const obj of objects) {
    if (!obj.placeholder?.options) continue;
    if (!obj.placeholder.options.color) {
      obj.placeholder.options.color = defaultColor;
    }
  }
}

/** Round to 2 decimal places for dedup comparison */
function round2(n) {
  return n != null ? Math.round(n * 100) / 100 : '';
}

// Re-export useful internals
export { extractPptx } from './parser/zip.js';
export { parseTheme, parseClrMap } from './parser/theme.js';
export { createColorResolver } from './parser/colors.js';
export { parseSlideMaster } from './parser/master.js';
export { parseSlideLayout, parsePresentation } from './parser/layout.js';
export { generateMastersCode, generateThemeJson, generateAgentInstructions } from './generator/code.js';
export { generateReport } from './generator/report.js';
export { generatePreview } from './generator/preview.js';
export { generateBrandSkill } from './generator/skill.js';
