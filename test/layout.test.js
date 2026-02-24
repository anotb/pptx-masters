import { describe, it, expect, beforeAll } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { parseSlideLayout, parsePresentation } from '../src/parser/layout.js';
import { extractPptx } from '../src/parser/zip.js';
import { emuToInches } from '../src/mapper/units.js';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const fixturePath = join(__dirname, 'fixtures', 'minimal.pptx');

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
});

// --- parseSlideLayout against minimal.pptx ---

describe('parseSlideLayout (fixture - TITLE_SLIDE layout)', () => {
  let pptx;
  let result;

  beforeAll(async () => {
    pptx = await extractPptx(fixturePath);
    const layoutXml = await pptx.getXml('ppt/slideLayouts/slideLayout2.xml');
    const layoutRels = await pptx.getXml('ppt/slideLayouts/_rels/slideLayout2.xml.rels');
    result = parseSlideLayout(layoutXml, layoutRels, pptx);
  });

  it('extracts layout name from p:cSld @_name', () => {
    expect(result.name).toBe('TITLE_SLIDE');
  });

  it('returns undefined type when not specified', () => {
    expect(result.type).toBeUndefined();
  });

  it('returns null clrMapOverride when inheriting master', () => {
    // Layout uses a:masterClrMapping, so override is null
    expect(result.clrMapOverride).toBeNull();
  });

  it('extracts background from p:bgPr', () => {
    expect(result.background).not.toBeNull();
    expect(result.background['a:solidFill']).toBeDefined();
  });

  it('detects title placeholder', () => {
    const title = result.placeholders.find((p) => p.type === 'title');
    expect(title).toBeDefined();
    expect(title.name).toBeTruthy();
  });

  it('detects body placeholder', () => {
    const body = result.placeholders.find((p) => p.type === 'body');
    expect(body).toBeDefined();
  });

  it('extracts placeholder position in inches', () => {
    const title = result.placeholders.find((p) => p.type === 'title');
    expect(title.position).toBeDefined();
    expect(title.position.x).toBe(emuToInches(457200));
    expect(title.position.y).toBe(emuToInches(457200));
    expect(title.position.w).toBe(emuToInches(8229600));
    expect(title.position.h).toBe(emuToInches(1371600));
  });

  it('extracts placeholder idx', () => {
    const title = result.placeholders.find((p) => p.type === 'title');
    expect(title.idx).toBeDefined();
  });

  it('has no static shapes (TITLE_SLIDE has only placeholders)', () => {
    expect(result.staticShapes).toHaveLength(0);
  });

  it('has no warnings', () => {
    expect(result.warnings).toHaveLength(0);
  });

  it('extracts relationships', () => {
    expect(Object.keys(result.relationships).length).toBeGreaterThan(0);
  });
});

describe('parseSlideLayout (fixture - CONTENT_SLIDE layout)', () => {
  let pptx;
  let result;

  beforeAll(async () => {
    pptx = await extractPptx(fixturePath);
    const layoutXml = await pptx.getXml('ppt/slideLayouts/slideLayout3.xml');
    const layoutRels = await pptx.getXml('ppt/slideLayouts/_rels/slideLayout3.xml.rels');
    result = parseSlideLayout(layoutXml, layoutRels, pptx);
  });

  it('extracts layout name', () => {
    expect(result.name).toBe('CONTENT_SLIDE');
  });

  it('extracts background fill', () => {
    expect(result.background).not.toBeNull();
    expect(result.background['a:solidFill']).toBeDefined();
  });

  it('detects title and body placeholders', () => {
    const title = result.placeholders.find((p) => p.type === 'title');
    const body = result.placeholders.find((p) => p.type === 'body');
    expect(title).toBeDefined();
    expect(body).toBeDefined();
  });

  it('detects slide number placeholder', () => {
    const sldNum = result.placeholders.find((p) => p.type === 'sldNum');
    expect(sldNum).toBeDefined();
  });

  it('detects static decoration shapes', () => {
    // Layout 3 has 2 static shapes (header bar rect + text overlay)
    expect(result.staticShapes.length).toBeGreaterThan(0);
  });

  it('static shapes have type "shape"', () => {
    for (const ss of result.staticShapes) {
      expect(ss.type).toBe('shape');
    }
  });

  it('static shape has name', () => {
    const firstStatic = result.staticShapes[0];
    expect(firstStatic.name).toBeTruthy();
  });

  it('static shape has position', () => {
    const firstStatic = result.staticShapes[0];
    expect(firstStatic.position).toBeDefined();
    expect(typeof firstStatic.position.x).toBe('number');
    expect(typeof firstStatic.position.y).toBe('number');
    expect(typeof firstStatic.position.w).toBe('number');
    expect(typeof firstStatic.position.h).toBe('number');
  });

  it('static shape has geometry', () => {
    const rect = result.staticShapes.find((s) => s.geometry === 'rect');
    expect(rect).toBeDefined();
  });

  it('static shape with fill has fill data', () => {
    const filled = result.staticShapes.find((s) => s.fill !== null);
    expect(filled).toBeDefined();
    expect(filled.fill.type).toBe('solidFill');
  });

  it('static shape with text has textProps', () => {
    const withText = result.staticShapes.find((s) => s.textProps !== null);
    expect(withText).toBeDefined();
    expect(withText.textProps.paragraphs).toBeDefined();
  });
});

describe('parseSlideLayout (fixture - DEFAULT layout)', () => {
  let result;

  beforeAll(async () => {
    const pptx = await extractPptx(fixturePath);
    const layoutXml = await pptx.getXml('ppt/slideLayouts/slideLayout1.xml');
    const layoutRels = await pptx.getXml('ppt/slideLayouts/_rels/slideLayout1.xml.rels');
    result = parseSlideLayout(layoutXml, layoutRels, pptx);
  });

  it('extracts name "DEFAULT"', () => {
    expect(result.name).toBe('DEFAULT');
  });

  it('has background reference (p:bgRef)', () => {
    expect(result.background).not.toBeNull();
    expect(result.background.bgRef).toBeDefined();
  });

  it('has no placeholders in empty layout', () => {
    expect(result.placeholders).toHaveLength(0);
  });
});

// --- Mock XML tests ---

describe('parseSlideLayout (mock)', () => {
  it('returns empty result for null input', () => {
    const result = parseSlideLayout(null, null);
    expect(result.name).toBe('');
    expect(result.type).toBeUndefined();
    expect(result.clrMapOverride).toBeNull();
    expect(result.background).toBeNull();
    expect(result.placeholders).toEqual([]);
    expect(result.staticShapes).toEqual([]);
    expect(result.warnings).toEqual([]);
    expect(result.relationships).toEqual({});
  });

  it('returns empty result for missing p:sldLayout root', () => {
    const result = parseSlideLayout({}, null);
    expect(result.name).toBe('');
  });

  it('extracts name from p:sldLayout @_name', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout name="Title Slide">
        <p:cSld><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.name).toBe('Title Slide');
  });

  it('extracts type from p:sldLayout @_type', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout type="title">
        <p:cSld><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.type).toBe('title');
  });

  it('falls back to p:cSld @_name when p:sldLayout @_name is missing', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld name="My Layout"><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.name).toBe('My Layout');
  });

  it('prefers p:sldLayout @_name over p:cSld @_name', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout name="Layout Name">
        <p:cSld name="CSld Name"><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.name).toBe('Layout Name');
  });

  it('detects a:overrideClrMapping', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld><p:spTree/></p:cSld>
        <p:clrMapOvr>
          <a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"
            accent1="accent1" accent2="accent2" accent3="accent3"
            accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
        </p:clrMapOvr>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.clrMapOverride).not.toBeNull();
    expect(result.clrMapOverride.bg1).toBe('dk1');
    expect(result.clrMapOverride.tx1).toBe('lt1');
  });

  it('returns null clrMapOverride for a:masterClrMapping', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld><p:spTree/></p:cSld>
        <p:clrMapOvr>
          <a:masterClrMapping/>
        </p:clrMapOvr>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.clrMapOverride).toBeNull();
  });

  it('extracts placeholder with type and idx', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title 1"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="title" idx="0"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="457200" y="274638"/>
                  <a:ext cx="8229600" cy="1143000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.placeholders).toHaveLength(1);
    expect(result.placeholders[0].type).toBe('title');
    expect(result.placeholders[0].idx).toBe('0');
    expect(result.placeholders[0].name).toBe('Title 1');
    expect(result.placeholders[0].position.x).toBe(emuToInches(457200));
    expect(result.placeholders[0].position.y).toBe(emuToInches(274638));
    expect(result.placeholders[0].position.w).toBe(emuToInches(8229600));
    expect(result.placeholders[0].position.h).toBe(emuToInches(1143000));
    expect(result.placeholders[0].shapeProps.geometry).toBe('rect');
  });

  it('extracts rotation from xfrm', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Rotated"/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm rot="5400000">
                  <a:off x="0" y="0"/>
                  <a:ext cx="914400" cy="914400"/>
                </a:xfrm>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.staticShapes[0].rotation).toBe(90);
  });

  it('separates placeholders from static shapes', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title 1"/>
                <p:nvPr><p:ph type="title" idx="0"/></p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="3" name="Decoration"/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.placeholders).toHaveLength(1);
    expect(result.placeholders[0].type).toBe('title');
    expect(result.staticShapes).toHaveLength(1);
    expect(result.staticShapes[0].name).toBe('Decoration');
  });

  it('handles p:pic elements', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="4" name="Logo"/>
                <p:cNvPicPr/>
                <p:nvPr/>
              </p:nvPicPr>
              <p:blipFill>
                <a:blip r:embed="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
              </p:blipFill>
              <p:spPr>
                <a:xfrm>
                  <a:off x="100000" y="200000"/>
                  <a:ext cx="500000" cy="500000"/>
                </a:xfrm>
              </p:spPr>
            </p:pic>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.staticShapes).toHaveLength(1);
    expect(result.staticShapes[0].type).toBe('picture');
    expect(result.staticShapes[0].name).toBe('Logo');
    expect(result.staticShapes[0].imageRef).toBe('rId5');
    expect(result.staticShapes[0].position).toBeDefined();
  });

  it('handles p:pic placeholder', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:pic>
              <p:nvPicPr>
                <p:cNvPr id="4" name="Picture Placeholder"/>
                <p:cNvPicPr/>
                <p:nvPr>
                  <p:ph type="pic" idx="1"/>
                </p:nvPr>
              </p:nvPicPr>
              <p:blipFill>
                <a:blip r:embed="rId6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
              </p:blipFill>
              <p:spPr/>
            </p:pic>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.placeholders).toHaveLength(1);
    expect(result.placeholders[0].type).toBe('pic');
    expect(result.placeholders[0].imageRef).toBe('rId6');
  });

  it('warns about grouped shapes', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:grpSp>
              <p:nvGrpSpPr><p:cNvPr id="10" name="Group 1"/></p:nvGrpSpPr>
            </p:grpSp>
            <p:grpSp>
              <p:nvGrpSpPr><p:cNvPr id="11" name="Group 2"/></p:nvGrpSpPr>
            </p:grpSp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.warnings).toHaveLength(1);
    expect(result.warnings[0]).toContain('2 grouped shape(s)');
  });

  it('warns about connection shapes', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:cxnSp>
              <p:nvCxnSpPr><p:cNvPr id="12" name="Connector 1"/></p:nvCxnSpPr>
            </p:cxnSp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.warnings.length).toBeGreaterThan(0);
    expect(result.warnings[0]).toContain('connection shape');
  });

  it('warns about graphic frames', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="13" name="Table 1"/></p:nvGraphicFramePr>
            </p:graphicFrame>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.warnings.length).toBeGreaterThan(0);
    expect(result.warnings[0]).toContain('graphic frame');
  });

  it('extracts shape fill types', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="NoFill Shape"/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr>
                <a:noFill/>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.staticShapes[0].fill.type).toBe('noFill');
  });

  it('handles shapes with no spPr', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Bare"/>
                <p:nvPr/>
              </p:nvSpPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.staticShapes).toHaveLength(1);
    expect(result.staticShapes[0].position).toBeNull();
    expect(result.staticShapes[0].geometry).toBeUndefined();
  });

  it('handles empty p:nvPr (no placeholder)', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Static"/>
                <p:nvPr/>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null);
    expect(result.placeholders).toHaveLength(0);
    expect(result.staticShapes).toHaveLength(1);
  });
});

// --- Position inheritance from master ---

describe('parseSlideLayout (position inheritance)', () => {
  it('inherits position from masterDefaults when placeholder has no xfrm', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title 1"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="title" idx="0"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const masterDefaults = {
      title: {
        position: { x: 0.5, y: 0.3, w: 9, h: 1.5 },
        textProps: null,
      },
    };

    const result = parseSlideLayout(xml, null, null, { masterDefaults });
    expect(result.placeholders).toHaveLength(1);
    expect(result.placeholders[0].position).toEqual({ x: 0.5, y: 0.3, w: 9, h: 1.5 });
  });

  it('prefers local position over masterDefaults', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title 1"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="title" idx="0"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr>
                <a:xfrm>
                  <a:off x="914400" y="914400"/>
                  <a:ext cx="7315200" cy="914400"/>
                </a:xfrm>
              </p:spPr>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const masterDefaults = {
      title: {
        position: { x: 0.5, y: 0.3, w: 9, h: 1.5 },
        textProps: null,
      },
    };

    const result = parseSlideLayout(xml, null, null, { masterDefaults });
    expect(result.placeholders[0].position.x).toBe(emuToInches(914400));
    expect(result.placeholders[0].position.y).toBe(emuToInches(914400));
  });

  it('inherits position by idx when type is not set', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Content"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph idx="10"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const masterDefaults = {
      'idx:10': {
        position: { x: 1, y: 2, w: 8, h: 4 },
        textProps: null,
      },
    };

    const result = parseSlideLayout(xml, null, null, { masterDefaults });
    expect(result.placeholders[0].position).toEqual({ x: 1, y: 2, w: 8, h: 4 });
  });

  it('inherits textProps from masterDefaults when not present locally', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Footer"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="ftr"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const masterDefaults = {
      ftr: {
        position: { x: 3, y: 6.8, w: 4, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [{ align: 'center', runs: [{ text: '' }] }],
          plainText: '',
        },
      },
    };

    const result = parseSlideLayout(xml, null, null, { masterDefaults });
    expect(result.placeholders[0].position).toEqual({ x: 3, y: 6.8, w: 4, h: 0.3 });
    expect(result.placeholders[0].textProps).toBeDefined();
    expect(result.placeholders[0].textProps.paragraphs).toBeDefined();
  });

  it('does not inherit when masterDefaults is null', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld>
          <p:spTree>
            <p:sp>
              <p:nvSpPr>
                <p:cNvPr id="2" name="Title"/>
                <p:cNvSpPr/>
                <p:nvPr>
                  <p:ph type="title"/>
                </p:nvPr>
              </p:nvSpPr>
              <p:spPr/>
            </p:sp>
          </p:spTree>
        </p:cSld>
      </p:sldLayout>
    `);

    const result = parseSlideLayout(xml, null, null, {});
    expect(result.placeholders[0].position).toBeNull();
  });
});

// --- showMasterSp behavior ---

describe('parseSlideLayout (showMasterSp)', () => {
  it('defaults showMasterSp to true when attribute is absent', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout>
        <p:cSld><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.showMasterSp).toBe(true);
  });

  it('sets showMasterSp to true when attribute is "1"', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout showMasterSp="1">
        <p:cSld><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.showMasterSp).toBe(true);
  });

  it('sets showMasterSp to false when attribute is "0"', () => {
    const xml = xmlParser.parse(`
      <p:sldLayout showMasterSp="0">
        <p:cSld><p:spTree/></p:cSld>
      </p:sldLayout>
    `);
    const result = parseSlideLayout(xml, null);
    expect(result.showMasterSp).toBe(false);
  });
});

// --- parsePresentation ---

describe('parsePresentation', () => {
  it('extracts slide dimensions from minimal.pptx', async () => {
    const pptx = await extractPptx(fixturePath);
    const presXml = await pptx.getXml('ppt/presentation.xml');
    const dims = parsePresentation(presXml);

    // 9144000 EMU = 10 inches, 5143500 EMU = 5.625 inches (16:9)
    expect(dims.width).toBe(emuToInches(9144000));
    expect(dims.height).toBe(emuToInches(5143500));
  });

  it('returns default dimensions for null input', () => {
    const dims = parsePresentation(null);
    expect(dims.width).toBe(10);
    expect(dims.height).toBe(7.5);
  });

  it('returns default dimensions for missing p:presentation', () => {
    const dims = parsePresentation({});
    expect(dims.width).toBe(10);
    expect(dims.height).toBe(7.5);
  });

  it('returns default dimensions for missing p:sldSz', () => {
    const dims = parsePresentation({ 'p:presentation': {} });
    expect(dims.width).toBe(10);
    expect(dims.height).toBe(7.5);
  });

  it('parses custom dimensions', () => {
    const dims = parsePresentation({
      'p:presentation': {
        'p:sldSz': {
          '@_cx': '12192000',  // 13.333 inches
          '@_cy': '6858000',   // 7.5 inches
        },
      },
    });

    expect(dims.width).toBe(emuToInches(12192000));
    expect(dims.height).toBe(emuToInches(6858000));
  });

  it('handles standard 4:3 dimensions', () => {
    const dims = parsePresentation({
      'p:presentation': {
        'p:sldSz': {
          '@_cx': '9144000',  // 10 inches
          '@_cy': '6858000',  // 7.5 inches
        },
      },
    });

    expect(dims.width).toBe(10);
    expect(dims.height).toBe(7.5);
  });
});
