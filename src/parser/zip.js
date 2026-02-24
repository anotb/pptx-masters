import { readFile } from 'fs/promises';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
});

/**
 * Default extractor using JSZip.
 * Implements the extractor interface:
 *   { extract(filePath) → Promise<{ getFile(path) → Promise<Buffer>, listFiles() → string[] }> }
 */
const defaultExtractor = {
  async extract(filePath) {
    const buffer = await readFile(filePath);
    const zip = await JSZip.loadAsync(buffer);

    return {
      async getFile(path) {
        const file = zip.file(path);
        if (!file) {
          throw new Error(`File not found in archive: ${path}`);
        }
        return file.async('nodebuffer');
      },

      listFiles() {
        return Object.keys(zip.files).filter((name) => !zip.files[name].dir);
      },
    };
  },
};

/**
 * Extract and parse a PPTX/POTX file.
 *
 * @param {string} filePath - Path to the .pptx or .potx file
 * @param {{ extractor?: { extract(filePath: string): Promise<{ getFile(path: string): Promise<Buffer>, listFiles(): string[] }> } }} [options]
 * @returns {Promise<{files: Map, getXml: Function, getBuffer: Function, listFiles: Function}>}
 */
export async function extractPptx(filePath, { extractor } = {}) {
  const ext = extractor || defaultExtractor;
  const archive = await ext.extract(filePath);

  return {
    /**
     * Read a file from the archive as text and parse as XML.
     * @param {string} path - File path within the archive
     * @returns {Promise<object>} Parsed XML object
     */
    async getXml(path) {
      const buffer = await archive.getFile(path);
      const text = buffer.toString('utf-8');
      return xmlParser.parse(text);
    },

    /**
     * Read a file from the archive as a buffer.
     * @param {string} path - File path within the archive
     * @returns {Promise<Buffer>} File contents as buffer
     */
    async getBuffer(path) {
      return archive.getFile(path);
    },

    /**
     * List all file paths in the archive.
     * @returns {string[]} Array of file paths
     */
    listFiles() {
      return archive.listFiles();
    },
  };
}
