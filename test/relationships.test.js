import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { parseRelationships, resolveRelPath } from '../src/parser/relationships.js';
import { extractPptx } from '../src/parser/zip.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = join(__dirname, 'fixtures', 'minimal.pptx');

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
});

// --- parseRelationships ---

describe('parseRelationships', () => {
  it('parses multiple relationships from XML string', () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
    </Relationships>`;

    const parsed = xmlParser.parse(xml);
    const rels = parseRelationships(parsed);

    expect(Object.keys(rels)).toHaveLength(3);
    expect(rels.rId1).toEqual({ type: 'slideLayout', target: '../slideLayouts/slideLayout1.xml' });
    expect(rels.rId2).toEqual({ type: 'image', target: '../media/image1.png' });
    expect(rels.rId3).toEqual({ type: 'theme', target: '../theme/theme1.xml' });
  });

  it('simplifies type URL to last segment', () => {
    const xml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
    </Relationships>`;

    const parsed = xmlParser.parse(xml);
    const rels = parseRelationships(parsed);

    expect(rels.rId1.type).toBe('slideMaster');
  });

  it('handles single relationship (not array)', () => {
    const xml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
    </Relationships>`;

    const parsed = xmlParser.parse(xml);
    const rels = parseRelationships(parsed);

    expect(Object.keys(rels)).toHaveLength(1);
    expect(rels.rId1).toEqual({ type: 'slide', target: 'slides/slide1.xml' });
  });

  it('handles various relationship types', () => {
    const xml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster" Target="handoutMasters/handoutMaster1.xml"/>
    </Relationships>`;

    const parsed = xmlParser.parse(xml);
    const rels = parseRelationships(parsed);

    expect(rels.rId1.type).toBe('theme');
    expect(rels.rId2.type).toBe('notesMaster');
    expect(rels.rId3.type).toBe('handoutMaster');
  });

  it('returns empty object for null input', () => {
    expect(parseRelationships(null)).toEqual({});
  });

  it('returns empty object for undefined input', () => {
    expect(parseRelationships(undefined)).toEqual({});
  });

  it('returns empty object for missing Relationships root', () => {
    expect(parseRelationships({})).toEqual({});
  });

  it('returns empty object when no Relationship children', () => {
    expect(parseRelationships({ Relationships: {} })).toEqual({});
  });

  it('parses .rels from real pptx fixture', async () => {
    const pptx = await extractPptx(fixturePath);
    const relsXml = await pptx.getXml('ppt/_rels/presentation.xml.rels');
    const rels = parseRelationships(relsXml);

    // Should have at least one relationship
    expect(Object.keys(rels).length).toBeGreaterThan(0);

    // All values should have type and target
    for (const [id, rel] of Object.entries(rels)) {
      expect(id).toMatch(/^rId\d+$/);
      expect(rel.type).toBeTruthy();
      expect(rel.target).toBeTruthy();
    }
  });

  it('finds theme relationship in real pptx fixture', async () => {
    const pptx = await extractPptx(fixturePath);
    const relsXml = await pptx.getXml('ppt/_rels/presentation.xml.rels');
    const rels = parseRelationships(relsXml);

    const themeRel = Object.values(rels).find((r) => r.type === 'theme');
    expect(themeRel).toBeDefined();
    expect(themeRel.target).toContain('theme');
  });
});

// --- resolveRelPath ---

describe('resolveRelPath', () => {
  it('resolves ../ from slideLayouts to media', () => {
    expect(
      resolveRelPath('ppt/slideLayouts/slideLayout1.xml', '../media/image1.png'),
    ).toBe('ppt/media/image1.png');
  });

  it('resolves ../ from slideMasters to slideLayouts', () => {
    expect(
      resolveRelPath('ppt/slideMasters/slideMaster1.xml', '../slideLayouts/slideLayout2.xml'),
    ).toBe('ppt/slideLayouts/slideLayout2.xml');
  });

  it('resolves relative path without ../', () => {
    expect(
      resolveRelPath('ppt/presentation.xml', 'slides/slide1.xml'),
    ).toBe('ppt/slides/slide1.xml');
  });

  it('handles multiple ../ segments', () => {
    expect(
      resolveRelPath('ppt/slideMasters/_rels/slideMaster1.xml.rels', '../../media/image1.png'),
    ).toBe('ppt/media/image1.png');
  });

  it('handles file at root level', () => {
    expect(
      resolveRelPath('[Content_Types].xml', 'ppt/presentation.xml'),
    ).toBe('ppt/presentation.xml');
  });

  it('handles absolute target (leading /)', () => {
    expect(
      resolveRelPath('ppt/slides/slide1.xml', '/ppt/media/image1.png'),
    ).toBe('ppt/media/image1.png');
  });

  it('resolves . segments', () => {
    expect(
      resolveRelPath('ppt/presentation.xml', './slides/slide1.xml'),
    ).toBe('ppt/slides/slide1.xml');
  });

  it('resolves ../ from slideMasters to theme', () => {
    expect(
      resolveRelPath('ppt/slideMasters/slideMaster1.xml', '../theme/theme1.xml'),
    ).toBe('ppt/theme/theme1.xml');
  });
});
