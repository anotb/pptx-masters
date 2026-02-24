import { describe, it, expect } from 'vitest';
import { parseTheme, parseClrMap } from '../src/parser/theme.js';
import { extractPptx } from '../src/parser/zip.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = join(__dirname, 'fixtures', 'minimal.pptx');

describe('parseTheme', () => {
  it('extracts all 12 scheme colors from minimal.pptx', async () => {
    const pptx = await extractPptx(fixturePath);
    const themeXml = await pptx.getXml('ppt/theme/theme1.xml');
    const theme = parseTheme(themeXml);

    const expectedSlots = [
      'dk1', 'lt1', 'dk2', 'lt2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hlink', 'folHlink',
    ];

    for (const slot of expectedSlots) {
      expect(theme.colors[slot]).toBeDefined();
      expect(theme.colors[slot]).toMatch(/^[0-9A-F]{6}$/i);
    }

    expect(Object.keys(theme.colors)).toHaveLength(12);
  });

  it('resolves sysClr via lastClr (dk1 = windowText)', async () => {
    const pptx = await extractPptx(fixturePath);
    const themeXml = await pptx.getXml('ppt/theme/theme1.xml');
    const theme = parseTheme(themeXml);

    expect(theme.colors.dk1).toBe('000000');
    expect(theme.colors.lt1).toBe('FFFFFF');
  });

  it('resolves srgbClr directly', async () => {
    const pptx = await extractPptx(fixturePath);
    const themeXml = await pptx.getXml('ppt/theme/theme1.xml');
    const theme = parseTheme(themeXml);

    expect(theme.colors.dk2).toBe('44546A');
    expect(theme.colors.accent1).toBe('4472C4');
    expect(theme.colors.hlink).toBe('0563C1');
  });

  it('extracts heading font (majorFont latin typeface)', async () => {
    const pptx = await extractPptx(fixturePath);
    const themeXml = await pptx.getXml('ppt/theme/theme1.xml');
    const theme = parseTheme(themeXml);

    expect(theme.fonts.heading).toBe('Calibri Light');
  });

  it('extracts body font (minorFont latin typeface)', async () => {
    const pptx = await extractPptx(fixturePath);
    const themeXml = await pptx.getXml('ppt/theme/theme1.xml');
    const theme = parseTheme(themeXml);

    expect(theme.fonts.body).toBe('Calibri');
  });

  it('extracts format scheme sections', async () => {
    const pptx = await extractPptx(fixturePath);
    const themeXml = await pptx.getXml('ppt/theme/theme1.xml');
    const theme = parseTheme(themeXml);

    expect(theme.formatScheme).toBeDefined();
    expect(theme.formatScheme.fillStyleLst).toBeDefined();
    expect(theme.formatScheme.lnStyleLst).toBeDefined();
    expect(theme.formatScheme.effectStyleLst).toBeDefined();
    expect(theme.formatScheme.bgFillStyleLst).toBeDefined();
  });

  it('throws on invalid theme XML', () => {
    expect(() => parseTheme({})).toThrow('Invalid theme XML');
  });

  it('throws when a:themeElements is missing', () => {
    expect(() => parseTheme({ 'a:theme': {} })).toThrow('missing a:themeElements');
  });

  it('returns empty colors when clrScheme is missing', () => {
    const theme = parseTheme({
      'a:theme': {
        'a:themeElements': {},
      },
    });
    expect(theme.colors).toEqual({});
  });
});

describe('parseClrMap', () => {
  it('parses a p:clrMap element from slide master', async () => {
    const pptx = await extractPptx(fixturePath);
    const masterXml = await pptx.getXml('ppt/slideMasters/slideMaster1.xml');
    const masterRoot = masterXml['p:sldMaster'];
    const clrMapEl = masterRoot['p:clrMap'];
    const clrMap = parseClrMap(clrMapEl);

    expect(clrMap.bg1).toBe('lt1');
    expect(clrMap.tx1).toBe('dk1');
    expect(clrMap.bg2).toBe('lt2');
    expect(clrMap.tx2).toBe('dk2');
    expect(clrMap.accent1).toBe('accent1');
    expect(clrMap.hlink).toBe('hlink');
    expect(clrMap.folHlink).toBe('folHlink');
  });

  it('returns empty object for null input', () => {
    expect(parseClrMap(null)).toEqual({});
  });

  it('returns empty object for undefined input', () => {
    expect(parseClrMap(undefined)).toEqual({});
  });

  it('strips @_ prefix from attribute keys', () => {
    const result = parseClrMap({
      '@_bg1': 'lt1',
      '@_tx1': 'dk1',
    });
    expect(result).toEqual({ bg1: 'lt1', tx1: 'dk1' });
  });
});
