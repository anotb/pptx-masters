/**
 * Background mapper — converts parsed OOXML background elements into
 * PptxGenJS-compatible background properties.
 */

/**
 * Map a background element to PptxGenJS background.
 *
 * @param {object|null} bgElement - Raw background element from parser (p:bgPr or { bgRef })
 * @param {{ resolve: Function }} colorResolver - Color resolver
 * @param {Record<string, { type: string, target: string }>} relationships - Resolved relationships
 * @param {string} [basePath] - Base path for resolving relative media paths
 * @returns {{ background: object|null, warnings: string[] }}
 */
export function mapBackground(bgElement, colorResolver, relationships, basePath) {
  const warnings = [];

  if (!bgElement) {
    return { background: null, warnings };
  }

  // bgRef — theme format scheme reference
  if (bgElement.bgRef) {
    const bgRef = bgElement.bgRef;
    // Try to resolve the color from bgRef
    const resolved = colorResolver?.resolve(bgRef);
    if (resolved) {
      const bg = { color: resolved.color };
      if (resolved.transparency != null) {
        bg.transparency = resolved.transparency;
      }
      return { background: bg, warnings };
    }
    warnings.push('bgRef (theme format scheme reference) not fully supported');
    return { background: null, warnings };
  }

  // Solid fill
  if (bgElement['a:solidFill']) {
    const resolved = colorResolver?.resolve(bgElement['a:solidFill']);
    if (resolved) {
      const bg = { color: resolved.color };
      if (resolved.transparency != null) {
        bg.transparency = resolved.transparency;
      }
      return { background: bg, warnings };
    }
    return { background: null, warnings };
  }

  // Image fill (blipFill)
  if (bgElement['a:blipFill']) {
    const blip = bgElement['a:blipFill']['a:blip'];
    const rId = blip?.['@_r:embed'];
    if (rId && relationships?.[rId]) {
      const rel = relationships[rId];
      const filename = rel.target.split('/').pop();
      return { background: { path: `./media/${filename}` }, warnings };
    }
    warnings.push('Background image reference could not be resolved');
    return { background: null, warnings };
  }

  // Gradient fill — use dominant stop as fallback
  if (bgElement['a:gradFill']) {
    const gsLst = bgElement['a:gradFill']['a:gsLst'];
    if (gsLst) {
      const gsItems = Array.isArray(gsLst['a:gs']) ? gsLst['a:gs'] : gsLst['a:gs'] ? [gsLst['a:gs']] : [];
      if (gsItems.length > 0) {
        // Use first stop as the dominant color
        const firstStop = gsItems[0];
        const resolved = colorResolver?.resolve(firstStop);
        if (resolved) {
          warnings.push('Gradient background not fully supported, using first stop color as fallback');
          return { background: { color: resolved.color }, warnings };
        }
      }
    }
    warnings.push('Gradient background not supported');
    return { background: null, warnings };
  }

  // Pattern fill — use foreground color as fallback
  if (bgElement['a:pattFill']) {
    const fgClr = bgElement['a:pattFill']['a:fgClr'];
    if (fgClr) {
      const resolved = colorResolver?.resolve(fgClr);
      if (resolved) {
        warnings.push('Pattern background not supported, using foreground color as fallback');
        return { background: { color: resolved.color }, warnings };
      }
    }
    warnings.push('Pattern background not supported');
    return { background: null, warnings };
  }

  // No fill
  if (bgElement['a:noFill'] != null) {
    return { background: null, warnings };
  }

  return { background: null, warnings };
}
