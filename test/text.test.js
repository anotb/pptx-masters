import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { extractTextProps, extractDefaultTextStyle } from '../src/parser/text.js';
import { createColorResolver } from '../src/parser/colors.js';
import { emuToInches } from '../src/mapper/units.js';

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
});

// Shared test fixtures
const themeColors = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

const clrMap = {
  bg1: 'lt1',
  tx1: 'dk1',
  bg2: 'lt2',
  tx2: 'dk2',
  accent1: 'accent1',
};

const themeFonts = { heading: 'Calibri Light', body: 'Calibri' };

function makeResolver() {
  return createColorResolver(themeColors, clrMap, themeFonts);
}

/**
 * Helper to parse an XML string and extract the a:txBody element.
 */
function parseTxBody(xmlStr) {
  const parsed = xmlParser.parse(xmlStr);
  return parsed['a:txBody'];
}

// --- Body properties ---

describe('extractTextProps - body properties', () => {
  it('extracts default margins when no attributes specified', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);

    expect(result.bodyProps.margin).toEqual([
      emuToInches(45720),  // top
      emuToInches(91440),  // right
      emuToInches(45720),  // bottom
      emuToInches(91440),  // left
    ]);
  });

  it('extracts custom margins from bodyPr attributes', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.margin).toEqual([0, 0, 0, 0]);
  });

  it('extracts vertical alignment top', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr anchor="t"/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.valign).toBe('top');
  });

  it('extracts vertical alignment center', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.valign).toBe('middle');
  });

  it('extracts vertical alignment bottom', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr anchor="b"/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.valign).toBe('bottom');
  });

  it('returns undefined valign when no anchor', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.valign).toBeUndefined();
  });

  it('extracts rotation from rot attribute', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr rot="-5400000"/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.rotation).toBe(-90);
  });

  it('extracts text direction from vert attribute', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr vert="vert270"/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.vert).toBe('vert270');
  });

  it('detects normAutofit as shrink', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr><a:normAutofit fontScale="90000"/></a:bodyPr>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.autoFit).toBe('shrink');
  });

  it('detects spAutoFit as resize', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr><a:spAutoFit/></a:bodyPr>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.autoFit).toBe('resize');
  });

  it('detects noAutofit as none', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr><a:noAutofit/></a:bodyPr>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.autoFit).toBe('none');
  });

  it('returns undefined autoFit when none specified', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>X</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.bodyProps.autoFit).toBeUndefined();
  });
});

// --- Paragraph properties ---

describe('extractTextProps - paragraphs', () => {
  it('extracts paragraph alignment', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:t>Centered</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].align).toBe('center');
  });

  it('defaults alignment to left', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>Default</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].align).toBe('left');
  });

  it('extracts justify alignment', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="just"/>
          <a:r><a:t>Justified</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].align).toBe('justify');
  });

  it('extracts right alignment', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="r"/>
          <a:r><a:t>Right</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].align).toBe('right');
  });

  it('extracts rtl mode', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr rtl="1"/>
          <a:r><a:t>RTL</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].rtlMode).toBe(true);
  });

  it('defaults rtl to false', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>LTR</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].rtlMode).toBe(false);
  });

  it('extracts line spacing in points', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:lnSpc><a:spcPts val="1800"/></a:lnSpc>
          </a:pPr>
          <a:r><a:t>Spaced</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].lineSpacing).toBe(18);
    expect(result.paragraphs[0].lineSpacingMultiple).toBeUndefined();
  });

  it('extracts line spacing as multiplier', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:lnSpc><a:spcPct val="150000"/></a:lnSpc>
          </a:pPr>
          <a:r><a:t>1.5x</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].lineSpacing).toBeUndefined();
    expect(result.paragraphs[0].lineSpacingMultiple).toBe(1.5);
  });

  it('extracts space before and after', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:spcBef><a:spcPts val="600"/></a:spcBef>
            <a:spcAft><a:spcPts val="1200"/></a:spcAft>
          </a:pPr>
          <a:r><a:t>Spaced</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].paraSpaceBefore).toBe(6);
    expect(result.paragraphs[0].paraSpaceAfter).toBe(12);
  });

  it('extracts margin left and indent', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr marL="457200" indent="-228600"/>
          <a:r><a:t>Indented</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].marginLeft).toBe(emuToInches(457200));
    expect(result.paragraphs[0].indent).toBe(emuToInches(-228600));
  });

  it('handles multiple paragraphs', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="l"/>
          <a:r><a:t>First</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r><a:t>Second</a:t></a:r>
        </a:p>
        <a:p>
          <a:pPr algn="r"/>
          <a:r><a:t>Third</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs).toHaveLength(3);
    expect(result.paragraphs[0].align).toBe('left');
    expect(result.paragraphs[1].align).toBe('center');
    expect(result.paragraphs[2].align).toBe('right');
  });
});

// --- Run properties ---

describe('extractTextProps - runs', () => {
  it('extracts text from runs', () => {
    // Note: fast-xml-parser trims whitespace by default, so trailing
    // spaces in <a:t> are stripped. "Hello " becomes "Hello".
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>Hello</a:t></a:r>
          <a:r><a:t>World</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs).toHaveLength(2);
    expect(result.paragraphs[0].runs[0].text).toBe('Hello');
    expect(result.paragraphs[0].runs[1].text).toBe('World');
  });

  it('extracts font size (hundredths of point / 100)', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr sz="2400"/>
            <a:t>Big</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].fontSize).toBe(24);
  });

  it('extracts bold and italic', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr b="1" i="1"/>
            <a:t>Bold Italic</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].bold).toBe(true);
    expect(result.paragraphs[0].runs[0].italic).toBe(true);
  });

  it('defaults bold and italic to undefined when not specified', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr/>
            <a:t>Normal</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].bold).toBeUndefined();
    expect(result.paragraphs[0].runs[0].italic).toBeUndefined();
  });

  it('extracts explicit bold=false', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr b="0" i="0"/>
            <a:t>Explicit not bold</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].bold).toBe(false);
    expect(result.paragraphs[0].runs[0].italic).toBe(false);
  });

  it('extracts underline', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr u="sng"/>
            <a:t>Underlined</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].underline).toBe('sng');
  });

  it('extracts strikethrough', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr strike="sngStrike"/>
            <a:t>Struck</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].strike).toBe('sngStrike');
  });

  it('extracts font face from a:latin', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr>
              <a:latin typeface="Arial"/>
            </a:rPr>
            <a:t>Arial text</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].fontFace).toBe('Arial');
  });

  it('resolves +mj-lt to heading font', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr>
              <a:latin typeface="+mj-lt"/>
            </a:rPr>
            <a:t>Heading</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].fontFace).toBe('Calibri Light');
  });

  it('resolves +mn-lt to body font', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr>
              <a:latin typeface="+mn-lt"/>
            </a:rPr>
            <a:t>Body</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].fontFace).toBe('Calibri');
  });

  it('extracts color from a:solidFill with srgbClr', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr>
              <a:solidFill>
                <a:srgbClr val="FF5500"/>
              </a:solidFill>
            </a:rPr>
            <a:t>Orange</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].color).toBe('FF5500');
  });

  it('extracts color from a:solidFill with schemeClr', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr>
              <a:solidFill>
                <a:schemeClr val="accent1"/>
              </a:solidFill>
            </a:rPr>
            <a:t>Accent</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].color).toBe('4472C4');
  });

  it('extracts superscript', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr baseline="30000"/>
            <a:t>sup</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].superscript).toBe(true);
    expect(result.paragraphs[0].runs[0].subscript).toBe(false);
  });

  it('extracts subscript', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr baseline="-25000"/>
            <a:t>sub</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].superscript).toBe(false);
    expect(result.paragraphs[0].runs[0].subscript).toBe(true);
  });

  it('extracts character spacing', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r>
            <a:rPr spc="300"/>
            <a:t>Spaced</a:t>
          </a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].charSpacing).toBe(3);
  });
});

// --- Bullets ---

describe('extractTextProps - bullets', () => {
  it('extracts character bullet', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:buChar char="\u2022"/>
            <a:buFont typeface="Arial"/>
          </a:pPr>
          <a:r><a:t>Bullet item</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    const bullet = result.paragraphs[0].bullet;
    expect(bullet.type).toBe('char');
    expect(bullet.characterCode).toBe('2022');
    expect(bullet.fontFace).toBe('Arial');
  });

  it('extracts numbered bullet', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:buAutoNum type="arabicPeriod"/>
          </a:pPr>
          <a:r><a:t>Item 1</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    const bullet = result.paragraphs[0].bullet;
    expect(bullet.type).toBe('number');
    expect(bullet.numberType).toBe('arabicPeriod');
  });

  it('extracts bullet color', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:buChar char="-"/>
            <a:buClr>
              <a:srgbClr val="FF0000"/>
            </a:buClr>
          </a:pPr>
          <a:r><a:t>Red bullet</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].bullet.color).toBe('FF0000');
  });

  it('extracts bullet size percent', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:buChar char="\u2022"/>
            <a:buSzPct val="75000"/>
          </a:pPr>
          <a:r><a:t>Small bullet</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].bullet.sizePercent).toBe(75);
  });

  it('detects buNone as false', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr>
            <a:buNone/>
          </a:pPr>
          <a:r><a:t>No bullet</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].bullet).toBe(false);
  });

  it('returns undefined bullet when no bullet elements', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="l"/>
          <a:r><a:t>Plain</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].bullet).toBeUndefined();
  });
});

// --- plainText ---

describe('extractTextProps - plainText', () => {
  it('concatenates all runs across paragraphs', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>Hello</a:t></a:r>
          <a:r><a:t>World</a:t></a:r>
        </a:p>
        <a:p>
          <a:r><a:t>Second line</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.plainText).toBe('HelloWorld\nSecond line');
  });

  it('handles single paragraph', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>Only one</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.plainText).toBe('Only one');
  });

  it('handles paragraph with no runs', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:pPr algn="ctr"/>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.plainText).toBe('');
    expect(result.paragraphs[0].runs).toEqual([]);
  });
});

// --- Edge cases ---

describe('extractTextProps - edge cases', () => {
  it('handles null txBody', () => {
    const result = extractTextProps(null, makeResolver(), themeFonts);
    expect(result.paragraphs).toEqual([]);
    expect(result.plainText).toBe('');
    expect(result.bodyProps).toBeDefined();
  });

  it('handles undefined txBody', () => {
    const result = extractTextProps(undefined, makeResolver(), themeFonts);
    expect(result.paragraphs).toEqual([]);
    expect(result.plainText).toBe('');
  });

  it('handles txBody with missing bodyPr', () => {
    const txBody = { 'a:p': { 'a:r': { 'a:t': 'Bare text' } } };
    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs).toHaveLength(1);
    expect(result.paragraphs[0].runs[0].text).toBe('Bare text');
  });

  it('handles run with no rPr', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:t>No props</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    const run = result.paragraphs[0].runs[0];
    expect(run.text).toBe('No props');
    expect(run.bold).toBeUndefined();
    expect(run.italic).toBeUndefined();
    expect(run.fontSize).toBeUndefined();
    expect(run.fontFace).toBeUndefined();
    expect(run.color).toBeUndefined();
  });

  it('handles numeric text content (fast-xml-parser parses numbers)', () => {
    // fast-xml-parser may parse "2024" as a number
    const txBody = { 'a:bodyPr': {}, 'a:p': { 'a:r': { 'a:t': 2024 } } };
    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].text).toBe('2024');
    expect(result.plainText).toBe('2024');
  });
});

// --- extractDefaultTextStyle ---

describe('extractDefaultTextStyle', () => {
  it('extracts run properties from a defRPr element', () => {
    const defRPr = {
      '@_sz': '1800',
      '@_b': '1',
      'a:solidFill': {
        'a:schemeClr': { '@_val': 'tx1' },
      },
      'a:latin': { '@_typeface': '+mj-lt' },
    };

    const result = extractDefaultTextStyle(defRPr, makeResolver(), themeFonts);
    expect(result.fontSize).toBe(18);
    expect(result.bold).toBe(true);
    expect(result.color).toBe('000000');
    expect(result.fontFace).toBe('Calibri Light');
  });

  it('handles null defRPr', () => {
    const result = extractDefaultTextStyle(null, makeResolver(), themeFonts);
    expect(result.bold).toBeUndefined();
    expect(result.italic).toBeUndefined();
    expect(result.fontSize).toBeUndefined();
  });

  it('handles defRPr with no styling', () => {
    const result = extractDefaultTextStyle({}, makeResolver(), themeFonts);
    expect(result.bold).toBeUndefined();
    expect(result.italic).toBeUndefined();
    expect(result.underline).toBeUndefined();
    expect(result.strike).toBeUndefined();
    expect(result.superscript).toBe(false);
    expect(result.subscript).toBe(false);
  });
});

// --- lstStyle parsing ---

describe('extractTextProps - lstStyle', () => {
  it('extracts lstStyle level 1 properties', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:lstStyle>
          <a:lvl1pPr algn="l">
            <a:defRPr sz="3200" b="0">
              <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
              <a:latin typeface="Aptos"/>
            </a:defRPr>
            <a:lnSpc><a:spcPts val="3200"/></a:lnSpc>
          </a:lvl1pPr>
        </a:lstStyle>
        <a:p><a:r><a:rPr lang="en-US"/><a:t>Title</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.lstStyleProps).toBeDefined();
    expect(result.lstStyleProps[1]).toBeDefined();
    expect(result.lstStyleProps[1].align).toBe('left');
    expect(result.lstStyleProps[1].lineSpacing).toBe(32);
    expect(result.lstStyleProps[1].defaultRunProps.fontSize).toBe(32);
    expect(result.lstStyleProps[1].defaultRunProps.bold).toBe(false);
    expect(result.lstStyleProps[1].defaultRunProps.fontFace).toBe('Aptos');
    expect(result.lstStyleProps[1].defaultRunProps.color).toBe('4472C4');
  });

  it('extracts multiple lstStyle levels', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:lstStyle>
          <a:lvl1pPr>
            <a:defRPr sz="2400"><a:latin typeface="Arial"/></a:defRPr>
          </a:lvl1pPr>
          <a:lvl2pPr>
            <a:defRPr sz="2000"><a:latin typeface="Arial"/></a:defRPr>
          </a:lvl2pPr>
          <a:lvl3pPr marL="365760">
            <a:defRPr sz="1800"><a:latin typeface="Arial"/></a:defRPr>
          </a:lvl3pPr>
        </a:lstStyle>
        <a:p><a:r><a:t>Content</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.lstStyleProps[1].defaultRunProps.fontSize).toBe(24);
    expect(result.lstStyleProps[2].defaultRunProps.fontSize).toBe(20);
    expect(result.lstStyleProps[3].defaultRunProps.fontSize).toBe(18);
    expect(result.lstStyleProps[3].marginLeft).toBeCloseTo(0.4, 1);
  });

  it('returns null lstStyleProps when no lstStyle', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>No lstStyle</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.lstStyleProps).toBeNull();
  });

  it('extracts lstStyle with line spacing percentage', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:lstStyle>
          <a:lvl1pPr>
            <a:defRPr sz="3600" b="1"><a:latin typeface="Aptos"/></a:defRPr>
            <a:lnSpc><a:spcPct val="95000"/></a:lnSpc>
          </a:lvl1pPr>
        </a:lstStyle>
        <a:p><a:r><a:t>Title</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.lstStyleProps[1].lineSpacingMultiple).toBe(0.95);
    expect(result.lstStyleProps[1].defaultRunProps.bold).toBe(true);
  });

  it('extracts lstStyle with buNone (no bullets)', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:lstStyle>
          <a:lvl1pPr>
            <a:buNone/>
            <a:defRPr sz="1800"><a:latin typeface="Aptos"/></a:defRPr>
          </a:lvl1pPr>
        </a:lstStyle>
        <a:p><a:r><a:t>No bullets</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.lstStyleProps[1].bullet).toBe(false);
  });
});

// --- a:fld (field) parsing ---

describe('extractTextProps - field elements', () => {
  it('extracts a:fld slide number field', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:fld type="slidenum">
            <a:rPr sz="800">
              <a:solidFill><a:schemeClr val="dk1"/></a:solidFill>
              <a:latin typeface="Aptos"/>
            </a:rPr>
            <a:t>\u2039#\u203A</a:t>
          </a:fld>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs).toHaveLength(1);
    const run = result.paragraphs[0].runs[0];
    expect(run.text).toBe('\u2039#\u203A');
    expect(run.isField).toBe(true);
    expect(run.fieldType).toBe('slidenum');
    expect(run.fontSize).toBe(8);
    expect(run.fontFace).toBe('Aptos');
    expect(run.color).toBe('000000');
    expect(result.plainText).toBe('\u2039#\u203A');
  });

  it('extracts datetime field', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:fld type="datetime1">
            <a:rPr sz="1000"/>
            <a:t>2/24/2026</a:t>
          </a:fld>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].runs[0].isField).toBe(true);
    expect(result.paragraphs[0].runs[0].fieldType).toBe('datetime1');
    expect(result.paragraphs[0].runs[0].text).toBe('2/24/2026');
  });

  it('extracts runs and fields together', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr/><a:t>Page </a:t></a:r>
          <a:fld type="slidenum"><a:rPr/><a:t>1</a:t></a:fld>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    const runs = result.paragraphs[0].runs;
    expect(runs).toHaveLength(2);
    expect(runs[0].text).toBe('Page');
    expect(runs[0].isField).toBeUndefined();
    expect(runs[1].text).toBe('1');
    expect(runs[1].isField).toBe(true);
  });
});

// --- a:br (line break) parsing ---

describe('extractTextProps - line breaks', () => {
  it('extracts a:br between two runs', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:r><a:rPr/><a:t>Line 1</a:t></a:r>
          <a:br><a:rPr/></a:br>
          <a:r><a:rPr/><a:t>Line 2</a:t></a:r>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    const runs = result.paragraphs[0].runs;
    expect(runs).toHaveLength(3);
    expect(runs[0].text).toBe('Line 1');
    expect(runs[1].text).toBe('\n');
    expect(runs[1].isBreak).toBe(true);
    expect(runs[2].text).toBe('Line 2');
    expect(result.plainText).toBe('Line 1\nLine 2');
  });

  it('handles paragraph with no runs and only a:br', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p>
          <a:br><a:rPr/></a:br>
        </a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    const runs = result.paragraphs[0].runs;
    expect(runs).toHaveLength(1);
    expect(runs[0].isBreak).toBe(true);
  });
});

// --- paragraph level ---

describe('extractTextProps - paragraph level', () => {
  it('extracts paragraph level attribute', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:pPr lvl="0"/><a:r><a:t>Level 1</a:t></a:r></a:p>
        <a:p><a:pPr lvl="1"/><a:r><a:t>Level 2</a:t></a:r></a:p>
        <a:p><a:pPr lvl="2"/><a:r><a:t>Level 3</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].level).toBe(1);
    expect(result.paragraphs[1].level).toBe(2);
    expect(result.paragraphs[2].level).toBe(3);
  });

  it('defaults level to 1', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>Default</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0].level).toBe(1);
  });

  it('tracks explicit alignment separately', () => {
    const txBody = parseTxBody(`
      <a:txBody>
        <a:bodyPr/>
        <a:p><a:pPr algn="r"/><a:r><a:t>Right</a:t></a:r></a:p>
        <a:p><a:r><a:t>Default</a:t></a:r></a:p>
      </a:txBody>
    `);

    const result = extractTextProps(txBody, makeResolver(), themeFonts);
    expect(result.paragraphs[0]._explicitAlign).toBe('right');
    expect(result.paragraphs[0].align).toBe('right');
    expect(result.paragraphs[1]._explicitAlign).toBeUndefined();
    expect(result.paragraphs[1].align).toBe('left'); // default
  });
});
