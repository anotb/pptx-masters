import { describe, it, expect } from 'vitest';
import { extractPptx } from '../src/parser/zip.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = join(__dirname, 'fixtures', 'minimal.pptx');

describe('extractPptx', () => {
  it('lists files in the PPTX archive', async () => {
    const pptx = await extractPptx(fixturePath);
    const files = pptx.listFiles();

    expect(files).toContain('ppt/presentation.xml');
    expect(files).toContain('[Content_Types].xml');
    expect(files.some((f) => f.startsWith('ppt/theme/'))).toBe(true);
    expect(files.some((f) => f.startsWith('ppt/slideMasters/'))).toBe(true);
    expect(files.some((f) => f.startsWith('ppt/slideLayouts/'))).toBe(true);
  });

  it('parses XML from the archive', async () => {
    const pptx = await extractPptx(fixturePath);
    const presentation = await pptx.getXml('ppt/presentation.xml');

    expect(presentation).toBeDefined();
    expect(typeof presentation).toBe('object');
    // presentation.xml should have a p:presentation root
    expect(presentation['p:presentation']).toBeDefined();
  });

  it('reads a buffer from the archive', async () => {
    const pptx = await extractPptx(fixturePath);
    const buffer = await pptx.getBuffer('[Content_Types].xml');

    expect(Buffer.isBuffer(buffer)).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
  });

  it('throws on non-existent file', async () => {
    await expect(extractPptx('/tmp/nonexistent-file.pptx')).rejects.toThrow();
  });

  it('throws when requesting a missing file from the archive', async () => {
    const pptx = await extractPptx(fixturePath);
    await expect(pptx.getXml('nonexistent/path.xml')).rejects.toThrow(
      'File not found in archive'
    );
  });
});
