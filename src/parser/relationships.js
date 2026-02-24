/**
 * Relationship resolver for OOXML .rels files.
 *
 * Parses the XML relationship files found at paths like:
 *   ppt/_rels/presentation.xml.rels
 *   ppt/slideMasters/_rels/slideMaster1.xml.rels
 *
 * Each .rels file maps relationship IDs (rId1, rId2, ...) to target paths
 * and relationship types.
 */

/**
 * Parse a .rels XML object (already parsed by fast-xml-parser) into a
 * lookup of relationship ID -> { type, target }.
 *
 * Type is simplified to the last segment of the Type URL:
 *   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
 *   becomes "slideLayout"
 *
 * @param {object} relsXml - Parsed .rels XML object
 * @returns {Record<string, { type: string, target: string }>}
 */
export function parseRelationships(relsXml) {
  const result = {};

  const root = relsXml?.['Relationships'];
  if (!root) return result;

  const rels = root['Relationship'];
  if (!rels) return result;

  const items = Array.isArray(rels) ? rels : [rels];

  for (const rel of items) {
    const id = rel['@_Id'];
    const typeUrl = rel['@_Type'] || '';
    const target = rel['@_Target'] || '';

    // Extract last segment of Type URL
    const lastSlash = typeUrl.lastIndexOf('/');
    const type = lastSlash >= 0 ? typeUrl.slice(lastSlash + 1) : typeUrl;

    result[id] = { type, target };
  }

  return result;
}

/**
 * Resolve a relative target path against a base file path.
 *
 * Examples:
 *   resolveRelPath("ppt/slideLayouts/slideLayout1.xml", "../media/image1.png")
 *     → "ppt/media/image1.png"
 *   resolveRelPath("ppt/slideMasters/slideMaster1.xml", "../slideLayouts/slideLayout2.xml")
 *     → "ppt/slideLayouts/slideLayout2.xml"
 *   resolveRelPath("ppt/presentation.xml", "slides/slide1.xml")
 *     → "ppt/slides/slide1.xml"
 *
 * @param {string} basePath - Path of the file containing the .rels reference
 * @param {string} relTarget - Relative target path from the .rels file
 * @returns {string} Resolved absolute path within the archive
 */
export function resolveRelPath(basePath, relTarget) {
  // If relTarget is already absolute (starts with /), strip the leading slash
  if (relTarget.startsWith('/')) {
    return relTarget.slice(1);
  }

  // Get the directory of the base path
  const lastSlash = basePath.lastIndexOf('/');
  const baseDir = lastSlash >= 0 ? basePath.slice(0, lastSlash) : '';

  // Split into segments
  const segments = baseDir ? baseDir.split('/') : [];
  const relSegments = relTarget.split('/');

  // Process each segment of the relative path
  for (const seg of relSegments) {
    if (seg === '..') {
      segments.pop();
    } else if (seg !== '.' && seg !== '') {
      segments.push(seg);
    }
  }

  return segments.join('/');
}
