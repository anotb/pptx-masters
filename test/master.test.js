import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { parseSlideMaster } from '../src/parser/master.js';
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

// --- Tests against minimal.pptx fixture ---

describe('parseSlideMaster (fixture)', () => {
  let pptx;
  let masterXml;
  let masterRels;
  let result;

  // Load fixture once
  beforeAll(async () => {
    pptx = await extractPptx(fixturePath);
    masterXml = await pptx.getXml('ppt/slideMasters/slideMaster1.xml');
    masterRels = await pptx.getXml('ppt/slideMasters/_rels/slideMaster1.xml.rels');
    result = parseSlideMaster(masterXml, masterRels, pptx);
  });

  // --- clrMap ---

  it('extracts clrMap with all 12 standard slots', () => {
    expect(Object.keys(result.clrMap)).toHaveLength(12);
    expect(result.clrMap.bg1).toBe('lt1');
    expect(result.clrMap.tx1).toBe('dk1');
    expect(result.clrMap.bg2).toBe('lt2');
    expect(result.clrMap.tx2).toBe('dk2');
    expect(result.clrMap.accent1).toBe('accent1');
    expect(result.clrMap.accent2).toBe('accent2');
    expect(result.clrMap.accent3).toBe('accent3');
    expect(result.clrMap.accent4).toBe('accent4');
    expect(result.clrMap.accent5).toBe('accent5');
    expect(result.clrMap.accent6).toBe('accent6');
    expect(result.clrMap.hlink).toBe('hlink');
    expect(result.clrMap.folHlink).toBe('folHlink');
  });

  // --- background ---

  it('returns null background when master has no p:bg', () => {
    // PptxGenJS minimal.pptx master has no background
    expect(result.background).toBeNull();
  });

  // --- shapes ---

  it('extracts shape tree with children', () => {
    expect(result.shapes.length).toBeGreaterThan(0);
  });

  it('placeholderDefaults has sldNum with position', () => {
    expect(result.placeholderDefaults.sldNum).toBeDefined();
    expect(result.placeholderDefaults.sldNum.position).toBeDefined();
    expect(result.placeholderDefaults.sldNum.position.x).toBe(9);
    expect(result.placeholderDefaults.sldNum.position.w).toBe(0.8);
  });

  it('includes the slide number placeholder shape', () => {
    const spShapes = result.shapes.filter((s) => s.type === 'p:sp');
    expect(spShapes.length).toBeGreaterThan(0);

    // The master has a slide number placeholder
    const sldNumShape = spShapes.find((s) => {
      const ph = s.element?.['p:nvSpPr']?.['p:nvPr']?.['p:ph'];
      return ph?.['@_type'] === 'sldNum';
    });
    expect(sldNumShape).toBeDefined();
  });

  // --- textStyles ---

  it('extracts text styles section', () => {
    expect(result.textStyles).toBeDefined();
    expect(result.textStyles.title).not.toBeNull();
    expect(result.textStyles.body).not.toBeNull();
  });

  it('title style has a font size defined in level 1', () => {
    const lvl1 = result.textStyles.title['a:lvl1pPr'];
    expect(lvl1).toBeDefined();
    const fontSize = lvl1['a:defRPr']?.['@_sz'];
    expect(fontSize).toBeDefined();
    // Font size is in hundredths of a point, should be a reasonable value
    expect(Number(fontSize)).toBeGreaterThan(0);
  });

  it('body style has multiple level properties', () => {
    expect(result.textStyles.body['a:lvl1pPr']).toBeDefined();
    expect(result.textStyles.body['a:lvl2pPr']).toBeDefined();
  });

  it('placeholderDefaults sldNum has textProps with font info', () => {
    const sldNum = result.placeholderDefaults.sldNum;
    expect(sldNum.textProps).toBeDefined();
    expect(sldNum.textProps.paragraphs.length).toBeGreaterThan(0);
    const lstStyle = sldNum.textProps.lstStyleProps;
    expect(lstStyle).toBeDefined();
    expect(lstStyle[1].defaultRunProps.fontFace).toBe('Arial');
    expect(lstStyle[1].defaultRunProps.fontSize).toBe(10);
  });

  // --- relationships ---

  it('extracts relationships with layout references', () => {
    expect(Object.keys(result.relationships).length).toBeGreaterThan(0);

    // Should have layout relationships
    const layoutRels = Object.values(result.relationships).filter(
      (r) => r.type === 'slideLayout'
    );
    expect(layoutRels.length).toBeGreaterThan(0);
  });

  it('extracts theme relationship', () => {
    const themeRel = Object.values(result.relationships).find(
      (r) => r.type === 'theme'
    );
    expect(themeRel).toBeDefined();
    expect(themeRel.target).toContain('theme1.xml');
  });

  it('extracts placeholderDefaults from fixture', () => {
    // The minimal.pptx master has at least a sldNum placeholder
    expect(Object.keys(result.placeholderDefaults).length).toBeGreaterThan(0);
    const sldNum = result.placeholderDefaults.sldNum;
    if (sldNum) {
      expect(sldNum.position).toBeDefined();
    }
  });
});

// --- Tests with mock XML ---

describe('parseSlideMaster (mock)', () => {
  it('returns empty result for null input', () => {
    const result = parseSlideMaster(null, null);
    expect(result.clrMap).toEqual({});
    expect(result.background).toBeNull();
    expect(result.shapes).toEqual([]);
    expect(result.textStyles).toEqual({ title: null, body: null, other: null });
    expect(result.relationships).toEqual({});
  });

  it('returns empty result for missing p:sldMaster root', () => {
    const result = parseSlideMaster({}, null);
    expect(result.clrMap).toEqual({});
    expect(result.shapes).toEqual([]);
  });

  it('extracts inline background (p:bgPr)', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld>
          <p:bg>
            <p:bgPr>
              <a:solidFill>
                <a:srgbClr val="003366"/>
              </a:solidFill>
            </p:bgPr>
          </p:bg>
          <p:spTree/>
        </p:cSld>
        <p:clrMap bg1="lt1" tx1="dk1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.background).not.toBeNull();
    expect(result.background['a:solidFill']).toBeDefined();
  });

  it('extracts background reference (p:bgRef)', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld>
          <p:bg>
            <p:bgRef idx="1001">
              <a:schemeClr val="bg1"/>
            </p:bgRef>
          </p:bg>
          <p:spTree/>
        </p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.background).not.toBeNull();
    expect(result.background.bgRef).toBeDefined();
    expect(result.background.bgRef['@_idx']).toBe('1001');
  });

  it('extracts multiple shapes from shape tree', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="Shape A"/></p:nvSpPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="3" name="Shape B"/></p:nvSpPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.shapes.length).toBe(2);
    expect(result.shapes[0].type).toBe('p:sp');
    expect(result.shapes[1].type).toBe('p:sp');
  });

  it('handles empty text styles', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld><p:spTree/></p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.textStyles).toEqual({ title: null, body: null, other: null });
  });

  it('extracts otherStyle when present', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld><p:spTree/></p:cSld>
        <p:clrMap bg1="lt1"/>
        <p:txStyles>
          <p:otherStyle>
            <a:lvl1pPr>
              <a:defRPr sz="1200"/>
            </a:lvl1pPr>
          </p:otherStyle>
        </p:txStyles>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.textStyles.other).not.toBeNull();
    expect(result.textStyles.other['a:lvl1pPr']).toBeDefined();
    expect(result.textStyles.title).toBeNull();
    expect(result.textStyles.body).toBeNull();
  });

  it('extracts p:pic elements from shape tree', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld>
          <p:spTree>
            <p:pic>
              <p:nvPicPr><p:cNvPr id="5" name="Picture 1"/></p:nvPicPr>
            </p:pic>
          </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.shapes.length).toBe(1);
    expect(result.shapes[0].type).toBe('p:pic');
  });

  it('parses relationships from rels XML', () => {
    const masterXml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld><p:spTree/></p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const relsXml = xmlParser.parse(`
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
      </Relationships>
    `);

    const result = parseSlideMaster(masterXml, relsXml);
    expect(result.relationships.rId1.type).toBe('slideLayout');
    expect(result.relationships.rId2.type).toBe('theme');
  });

  it('returns empty placeholderDefaults for null input', () => {
    const result = parseSlideMaster(null, null);
    expect(result.placeholderDefaults).toEqual({});
  });

  it('returns empty placeholderDefaults when no placeholders exist', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Decoration"/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="914400" cy="914400"/>
                </a:xfrm>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.placeholderDefaults).toEqual({});
  });

  it('extracts placeholderDefaults keyed by type', () => {
    const xml = xmlParser.parse(`
      <p:sldMaster>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title Placeholder"/>
                <p:nvPr>
                  <p:ph type="title"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="457200" y="274638"/>
                  <a:ext cx="8229600" cy="1143000"/>
                </a:xfrm>
              </p:spPr>
            </p:sp>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="3" name="Slide Number"/>
                <p:nvPr>
                  <p:ph type="sldNum" idx="12"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="6553200" y="6356350"/>
                  <a:ext cx="2133600" cy="365125"/>
                </a:xfrm>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
        <p:clrMap bg1="lt1"/>
      </p:sldMaster>
    `);

    const result = parseSlideMaster(xml, null);
    expect(result.placeholderDefaults.title).toBeDefined();
    expect(result.placeholderDefaults.title.position).toBeDefined();
    expect(result.placeholderDefaults.title.position.x).toBeGreaterThan(0);
    expect(result.placeholderDefaults.title.position.w).toBeGreaterThan(0);

    expect(result.placeholderDefaults.sldNum).toBeDefined();
    expect(result.placeholderDefaults.sldNum.position).toBeDefined();
    expect(result.placeholderDefaults['idx:12']).toBeDefined();
    expect(result.placeholderDefaults['idx:12'].position).toEqual(
      result.placeholderDefaults.sldNum.position,
    );
  });

});
