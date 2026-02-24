import { describe, it, expect } from 'vitest';
import {
  generateMastersCode,
  generateThemeJson,
  generateAgentInstructions,
  toUpperSnakeCase,
  detectLimitedPalette,
  extractBackgroundColors,
} from '../src/generator/code.js';
import { generateReport } from '../src/generator/report.js';
import { generatePreview } from '../src/generator/preview.js';

// --- Shared fixtures ---

const themeColors = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '1F497D',
  lt2: 'EEECE1',
  accent1: '4F81BD',
  accent2: 'C0504D',
  accent3: '9BBB59',
  accent4: '8064A2',
  accent5: '4BACC6',
  accent6: 'F79646',
  hlink: '0000FF',
  folHlink: '800080',
};

const themeFonts = { heading: 'Calibri Light', body: 'Calibri' };

const dimensions = { width: 10, height: 7.5 };

function makeMasterData() {
  return [
    {
      name: 'Title Slide',
      background: { color: '003366' },
      slideNumber: {
        x: 9.0,
        y: 6.9,
        w: 0.8,
        h: 0.3,
        fontFace: 'Arial',
        fontSize: 10,
        color: 'FFFFFF',
        align: 'right',
      },
      objects: [
        {
          placeholder: {
            options: {
              name: 'title',
              type: 'title',
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 1.5,
              fontFace: 'Arial',
              fontSize: 36,
              color: 'FFFFFF',
              bold: true,
              align: 'center',
            },
          },
        },
        {
          placeholder: {
            options: {
              name: 'body',
              type: 'body',
              x: 0.5,
              y: 2.5,
              w: 9,
              h: 2,
              fontFace: 'Arial',
              fontSize: 18,
              color: 'CCCCCC',
              align: 'center',
            },
          },
        },
      ],
    },
    {
      name: 'Title and Content',
      background: { color: 'FFFFFF' },
      slideNumber: null,
      objects: [
        {
          placeholder: {
            options: {
              name: 'title',
              type: 'title',
              x: 0.5,
              y: 0.5,
              w: 9,
              h: 0.8,
              fontFace: 'Calibri',
              fontSize: 28,
            },
          },
        },
        {
          placeholder: {
            options: {
              name: 'content',
              type: 'body',
              x: 0.5,
              y: 1.5,
              w: 9,
              h: 5.0,
            },
          },
        },
        {
          rect: { x: 0, y: 7.0, w: 10, h: 0.5, fill: { color: '003366' } },
        },
      ],
    },
  ];
}

function makeExtractionResult() {
  return {
    templateName: 'corporate-template.potx',
    dimensions,
    themeColors,
    themeFonts,
    layouts: [
      {
        name: 'Title Slide',
        background: { color: '003366' },
        placeholders: [
          { type: 'title', name: 'title', position: { x: 0.5, y: 0.5, w: 9, h: 1.5 }, fontFace: 'Arial', fontSize: 36 },
          { type: 'body', name: 'body', position: { x: 0.5, y: 2.5, w: 9, h: 2 }, fontFace: 'Arial', fontSize: 18 },
        ],
        staticShapes: [],
        slideNumber: { x: 9.0, y: 6.9, fontFace: 'Arial', fontSize: 10, color: 'FFFFFF' },
        footerObjects: [],
        warnings: [],
      },
      {
        name: 'Content Slide',
        background: null,
        placeholders: [
          { type: 'title', name: 'title', position: { x: 0.5, y: 0.5, w: 9, h: 0.8 } },
          { type: 'body', name: 'content', position: { x: 0.5, y: 1.5, w: 9, h: 5 } },
        ],
        staticShapes: [{ type: 'rect' }, { type: 'line' }],
        slideNumber: null,
        footerObjects: [],
        warnings: ['Found 2 grouped shapes (not supported in v1)'],
      },
    ],
    allWarnings: [
      'Layout "Content Slide": Found 2 grouped shapes (not supported in v1)',
      'Layout "Divider": Gradient fill used for background — using dominant color #003366 as fallback',
    ],
  };
}

// --- toUpperSnakeCase ---

describe('toUpperSnakeCase', () => {
  it('converts space-separated words', () => {
    expect(toUpperSnakeCase('Title Slide')).toBe('TITLE_SLIDE');
  });

  it('converts multi-word names', () => {
    expect(toUpperSnakeCase('Title and Content')).toBe('TITLE_AND_CONTENT');
  });

  it('handles single word', () => {
    expect(toUpperSnakeCase('Blank')).toBe('BLANK');
  });

  it('strips special characters', () => {
    expect(toUpperSnakeCase('Section Header (v2)')).toBe('SECTION_HEADER_V2');
  });

  it('handles multiple spaces', () => {
    expect(toUpperSnakeCase('Two  Part  Name')).toBe('TWO_PART_NAME');
  });

  it('handles leading/trailing spaces', () => {
    expect(toUpperSnakeCase('  Padded  ')).toBe('PADDED');
  });
});

// --- generateMastersCode ---

describe('generateMastersCode', () => {
  it('generates valid JavaScript (syntax check via Function constructor)', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    // Replace export keywords so Function constructor can parse it
    const testable = code
      .replace(/^export /gm, '')
      .replace(/^import .*/gm, '');

    // Should not throw — validates JavaScript syntax
    expect(() => new Function(testable)).not.toThrow();
  });

  it('contains "Generated by pptx-masters" header comment', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('Generated by pptx-masters');
  });

  it('contains template name in header', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'corporate-template.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('corporate-template.potx');
  });

  it('exports THEME_COLORS', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('export { THEME_COLORS, PALETTE, HEADING_FONT, BODY_FONT, THEME, POS, CHART_COLORS, FONT }');
  });

  it('exports registerMasters function', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('export function registerMasters(pptx)');
  });

  it('uses UPPER_SNAKE_CASE for master titles', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain("title: 'TITLE_SLIDE'");
    expect(code).toContain("title: 'TITLE_AND_CONTENT'");
  });

  it('includes original layout name as comment', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('// Title Slide');
    expect(code).toContain('// Title and Content');
  });

  it('includes all 12 theme color slots', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    for (const slot of ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
      expect(code).toContain(`${slot}: '`);
    }
  });

  it('includes theme font constants', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain("HEADING_FONT = 'Calibri Light'");
    expect(code).toContain("BODY_FONT = 'Calibri'");
  });

  it('includes background in master definition', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('background:');
    expect(code).toContain('003366');
  });

  it('includes slideNumber in master definition', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('slideNumber:');
  });

  it('includes placeholder objects', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('placeholder');
    expect(code).toContain('"type": "title"');
  });

  it('includes static shape objects', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('"rect"');
  });

  it('omits undefined/null values', () => {
    const masterData = [{
      name: 'Test',
      background: null,
      slideNumber: null,
      objects: [
        {
          placeholder: {
            options: {
              name: 'title',
              type: 'title',
              x: 1,
              y: 1,
              w: 8,
              h: 1,
              bold: undefined,
              italic: undefined,
              color: undefined,
            },
          },
        },
      ],
    }];

    const code = generateMastersCode(masterData, {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).not.toContain('undefined');
    expect(code).not.toContain(': null');
  });

  it('rounds numbers to 4 decimal places max', () => {
    const masterData = [{
      name: 'Test',
      background: null,
      slideNumber: null,
      objects: [
        {
          placeholder: {
            options: {
              name: 'title',
              type: 'title',
              x: 0.123456789,
              y: 1.9999999,
              w: 8,
              h: 1,
            },
          },
        },
      ],
    }];

    const code = generateMastersCode(masterData, {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('0.1235');
    expect(code).toContain('2');
    expect(code).not.toContain('0.123456789');
  });

  it('handles empty masterData array', () => {
    const code = generateMastersCode([], {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('export function registerMasters(pptx)');
    expect(code).toContain('Generated by pptx-masters');
  });

  it('generates code for single master', () => {
    const code = generateMastersCode([makeMasterData()[0]], {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain("title: 'TITLE_SLIDE'");
    expect(code).not.toContain('TITLE_AND_CONTENT');
  });

  it('handles image background', () => {
    const masterData = [{
      name: 'Image BG',
      background: { path: './media/bg.png' },
      slideNumber: null,
      objects: [],
    }];

    const code = generateMastersCode(masterData, {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });

    expect(code).toContain('./media/bg.png');
  });
});

// --- generateThemeJson ---

describe('generateThemeJson', () => {
  it('produces valid JSON with all fields', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions);

    expect(result).toHaveProperty('colors');
    expect(result).toHaveProperty('fonts');
    expect(result).toHaveProperty('dimensions');
  });

  it('includes all theme colors', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions);

    expect(result.colors.dk1).toBe('000000');
    expect(result.colors.lt1).toBe('FFFFFF');
    expect(result.colors.accent1).toBe('4F81BD');
    expect(result.colors.hlink).toBe('0000FF');
    expect(result.colors.folHlink).toBe('800080');
  });

  it('includes font names', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions);

    expect(result.fonts.heading).toBe('Calibri Light');
    expect(result.fonts.body).toBe('Calibri');
  });

  it('includes dimensions', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions);

    expect(result.dimensions.width).toBe(10);
    expect(result.dimensions.height).toBe(7.5);
  });

  it('defaults fonts to Calibri when not provided', () => {
    const result = generateThemeJson(themeColors, null, dimensions);

    expect(result.fonts.heading).toBe('Calibri');
    expect(result.fonts.body).toBe('Calibri');
  });

  it('defaults dimensions when not provided', () => {
    const result = generateThemeJson(themeColors, themeFonts, null);

    expect(result.dimensions.width).toBe(10);
    expect(result.dimensions.height).toBe(7.5);
  });

  it('can be serialized to valid JSON string', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions);
    const jsonStr = JSON.stringify(result);
    const parsed = JSON.parse(jsonStr);

    expect(parsed.colors.dk1).toBe('000000');
    expect(parsed.fonts.heading).toBe('Calibri Light');
    expect(parsed.dimensions.width).toBe(10);
  });
});

// --- generateReport ---

describe('generateReport', () => {
  it('generates markdown with Template section', () => {
    const report = generateReport(makeExtractionResult());
    expect(report).toContain('**Template:** corporate-template.potx');
  });

  it('generates markdown with Theme section', () => {
    const report = generateReport(makeExtractionResult());
    expect(report).toContain('## Theme');
  });

  it('generates markdown with Layouts section', () => {
    const report = generateReport(makeExtractionResult());
    expect(report).toContain('## Layouts');
  });

  it('generates markdown with Warnings section', () => {
    const report = generateReport(makeExtractionResult());
    expect(report).toContain('## Warnings');
  });

  it('includes color table with all 12 color slots', () => {
    const report = generateReport(makeExtractionResult());

    for (const slot of ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
      expect(report).toContain(`| ${slot} |`);
    }
  });

  it('includes hex values with # prefix in color table', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('#000000');
    expect(report).toContain('#FFFFFF');
    expect(report).toContain('#4F81BD');
  });

  it('includes emoji previews in color table', () => {
    const report = generateReport(makeExtractionResult());

    // Should have some emoji characters (black/white squares at minimum)
    expect(report).toMatch(/[\u2B1B\u2B1C\uD83D\uDFE5\uD83D\uDFE6\uD83D\uDFE9\uD83D\uDFE8\uD83D\uDFE7\uD83D\uDFEA]/);
  });

  it('includes theme font names', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('**Heading:** Calibri Light');
    expect(report).toContain('**Body:** Calibri');
  });

  it('includes slide dimensions', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('10"');
    expect(report).toContain('7.5"');
    expect(report).toContain('Widescreen');
  });

  it('lists layout names with numbers', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('### 1. Title Slide');
    expect(report).toContain('### 2. Content Slide');
  });

  it('lists placeholders for each layout', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('**Placeholders:**');
    expect(report).toContain('title:');
    expect(report).toContain('body:');
  });

  it('includes background description', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('**Background:** Solid #003366');
    expect(report).toContain('**Background:** None');
  });

  it('includes slide number info', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('**Slide Number:**');
    expect(report).toContain('Arial');
    expect(report).toContain('10pt');
  });

  it('includes static shapes count', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('**Static Shapes:** 0');
    expect(report).toContain('**Static Shapes:** 2');
  });

  it('includes all warnings in warnings section', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain('Found 2 grouped shapes');
    expect(report).toContain('Gradient fill used for background');
  });

  it('includes per-layout warnings', () => {
    const report = generateReport(makeExtractionResult());

    // Content Slide has a layout-level warning
    expect(report).toContain('**Warnings:**');
    expect(report).toContain('grouped shapes');
  });

  it('includes "What\'s Not Supported" section', () => {
    const report = generateReport(makeExtractionResult());

    expect(report).toContain("What's Not Supported (v1)");
    expect(report).toContain('Gradient fills');
    expect(report).toContain('Pattern fills');
    expect(report).toContain('Grouped shapes');
    expect(report).toContain('Animations and transitions');
    expect(report).toContain('SmartArt');
    expect(report).toContain('3D effects');
  });

  it('handles extraction result with no warnings', () => {
    const result = makeExtractionResult();
    result.allWarnings = [];
    const report = generateReport(result);

    expect(report).toContain('No warnings.');
  });

  it('handles extraction result with no layouts', () => {
    const result = makeExtractionResult();
    result.layouts = [];
    const report = generateReport(result);

    expect(report).toContain('## Layouts');
    // Should not crash
    expect(report).toContain('## Warnings');
  });

  it('handles layout with no placeholders', () => {
    const result = makeExtractionResult();
    result.layouts = [{
      name: 'Blank',
      background: null,
      placeholders: [],
      staticShapes: [],
      slideNumber: null,
      footerObjects: [],
      warnings: [],
    }];
    const report = generateReport(result);

    expect(report).toContain('Blank');
    expect(report).toContain('**Placeholders:** None');
  });
});

// --- generatePreview ---

describe('generatePreview', () => {
  it('generates a valid PPTX buffer with ZIP signature', async () => {
    const result = await generatePreview(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
    );

    expect(result).toBeInstanceOf(Buffer);
    expect(result.length).toBeGreaterThan(0);
    // PPTX is a ZIP file, should start with PK (0x50, 0x4B)
    expect(result[0]).toBe(0x50);
    expect(result[1]).toBe(0x4b);
  });

  it('handles empty master data and null dimensions', async () => {
    const result = await generatePreview(
      [],
      themeColors,
      themeFonts,
      null,
    );

    expect(result).toBeInstanceOf(Buffer);
    expect(result[0]).toBe(0x50);
    expect(result[1]).toBe(0x4b);
  });

  it('handles master with only static shapes (no placeholders)', async () => {
    const masterData = [{
      name: 'Blank',
      background: { color: 'FFFFFF' },
      slideNumber: null,
      objects: [
        { rect: { x: 0, y: 0, w: 10, h: 7.5, fill: { color: 'EEEEEE' } } },
      ],
    }];

    const result = await generatePreview(
      masterData,
      themeColors,
      themeFonts,
      dimensions,
    );

    expect(result).toBeInstanceOf(Buffer);
    expect(result[0]).toBe(0x50);
    expect(result[1]).toBe(0x4b);
  });
});

// --- generateAgentInstructions ---

describe('generateAgentInstructions', () => {
  it('includes template name in title', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'corporate-template.potx',
    );

    expect(md).toContain('# Slide Masters: corporate-template.potx');
  });

  it('includes Available Masters table', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('## Available Masters');
    expect(md).toContain('| Master | Title PH | Body PH | Slide # | Background |');
    expect(md).toContain('TITLE_SLIDE');
    expect(md).toContain('TITLE_AND_CONTENT');
  });

  it('marks placeholders with check/cross in table', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    // Title Slide has both title and body placeholders
    expect(md).toContain('\u2713');
  });

  it('includes Quick Start code block', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('## Quick Start');
    expect(md).toContain("createPresentation");
    expect(md).toContain("from './masters.js'");
    expect(md).toContain("masterName: 'TITLE_SLIDE'");
  });

  it('includes Theme Colors section', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('## Theme Colors');
    expect(md).toContain('Dark 1 (text)');
    expect(md).toContain('#000000');
    expect(md).toContain('#FFFFFF');
    expect(md).toContain('#4F81BD');
  });

  it('includes Theme Fonts section', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('## Theme Fonts');
    expect(md).toContain('Headings: Calibri Light');
    expect(md).toContain('Body: Calibri');
  });

  it('includes Master Details sections', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('### TITLE_SLIDE');
    expect(md).toContain('Original name: "Title Slide"');
    expect(md).toContain('### TITLE_AND_CONTENT');
    expect(md).toContain('Original name: "Title and Content"');
  });

  it('lists placeholder details for each master', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('`title`');
    expect(md).toContain('`body`');
  });

  it('includes slide number info in master details', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('Slide number:');
    expect(md).toContain('Arial');
    expect(md).toContain('10pt');
  });

  it('includes static shapes count', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    // Title and Content has one rect static shape
    expect(md).toContain('Static shapes: 1');
  });

  it('handles empty master data', () => {
    const md = generateAgentInstructions(
      [],
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('# Slide Masters: test.potx');
    expect(md).toContain('## Available Masters');
  });

  it('handles master without placeholders', () => {
    const masterData = [{
      name: 'Blank',
      background: null,
      slideNumber: null,
      objects: [],
    }];

    const md = generateAgentInstructions(
      masterData,
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );

    expect(md).toContain('BLANK');
    expect(md).toContain('\u2717'); // cross marks for no placeholders
  });
});

// --- Limited Palette Detection ---

// Fixtures for limited palette tests
const limitedThemeColors = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '444444',
  lt2: 'EEEEEE',
  accent1: 'FFFFFF',
  accent2: 'FFFFFF',
  accent3: 'FFFFFF',
  accent4: 'FFFFFF',
  accent5: 'FFFFFF',
  accent6: 'FFFFFF',
  hlink: '0563C1',
  folHlink: '954F72',
};

function makeLimitedMasterData() {
  return [
    {
      name: 'Title Slide',
      background: { color: '038DAF' },
      slideNumber: null,
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 4, w: 12, h: 1.5, fontSize: 36 } } },
        { placeholder: { options: { name: 'subtitle', type: 'body', x: 0.5, y: 5.5, w: 12, h: 1, fontSize: 18 } } },
      ],
    },
    {
      name: 'Content',
      background: { color: 'FFFFFF' },
      slideNumber: { x: 12, y: 7, w: 0.5, h: 0.3, fontSize: 10 },
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.4, w: 12, h: 0.5, fontSize: 28 } } },
        { placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 1.5, w: 12, h: 5, fontSize: 14 } } },
      ],
    },
    {
      name: 'Section Header',
      background: { color: '963C50' },
      slideNumber: null,
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: 1, y: 3, w: 11, h: 1.5, fontSize: 40 } } },
      ],
    },
    {
      name: 'Two Content',
      background: { color: '038DAF' },
      slideNumber: { x: 12, y: 7, w: 0.5, h: 0.3, fontSize: 10 },
      objects: [
        { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.4, w: 12, h: 0.5, fontSize: 28 } } },
        { placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 1.5, w: 5.5, h: 5, fontSize: 14 } } },
        { placeholder: { options: { name: 'body', type: 'body', x: 6.5, y: 1.5, w: 5.5, h: 5, fontSize: 14 } } },
      ],
    },
    {
      name: 'Closing',
      background: { color: '4C216D' },
      slideNumber: null,
      objects: [],
    },
  ];
}

describe('detectLimitedPalette', () => {
  it('detects all-white accents as limited', () => {
    const result = detectLimitedPalette(limitedThemeColors);
    expect(result.isLimited).toBe(true);
    expect(result.usableAccents).toHaveLength(0);
  });

  it('detects healthy palette as non-limited', () => {
    const result = detectLimitedPalette(themeColors);
    expect(result.isLimited).toBe(false);
    expect(result.usableAccents.length).toBeGreaterThanOrEqual(2);
  });

  it('detects all-identical non-white accents as limited', () => {
    const sameColor = {
      ...themeColors,
      accent1: '336699', accent2: '336699', accent3: '336699',
      accent4: '336699', accent5: '336699', accent6: '336699',
    };
    const result = detectLimitedPalette(sameColor);
    expect(result.isLimited).toBe(true);
    expect(result.usableAccents).toEqual(['336699']);
  });

  it('treats near-black accents as neutral', () => {
    const nearBlack = {
      ...themeColors,
      accent1: '050505', accent2: '050505', accent3: '050505',
      accent4: '050505', accent5: '050505', accent6: '050505',
    };
    const result = detectLimitedPalette(nearBlack);
    expect(result.isLimited).toBe(true);
    expect(result.usableAccents).toHaveLength(0);
  });

  it('returns non-limited with 2+ real accents', () => {
    const twoReal = {
      ...limitedThemeColors,
      accent1: '336699',
      accent2: 'CC3333',
    };
    const result = detectLimitedPalette(twoReal);
    expect(result.isLimited).toBe(false);
    expect(result.usableAccents.length).toBe(2);
  });

  it('handles null themeColors', () => {
    const result = detectLimitedPalette(null);
    expect(result.isLimited).toBe(false);
    expect(result.usableAccents).toHaveLength(0);
  });
});

describe('extractBackgroundColors', () => {
  it('extracts unique non-neutral background colors', () => {
    const colors = extractBackgroundColors(makeLimitedMasterData());
    expect(colors).toContain('038DAF');
    expect(colors).toContain('963C50');
    expect(colors).toContain('4C216D');
    // FFFFFF should be filtered out as near-white
    expect(colors).not.toContain('FFFFFF');
  });

  it('sorts by frequency then darkness', () => {
    const colors = extractBackgroundColors(makeLimitedMasterData());
    // 038DAF appears 2x (Title Slide + Two Content), others 1x
    expect(colors[0]).toBe('038DAF');
  });

  it('skips image backgrounds', () => {
    const masterData = [
      { name: 'Image BG', background: { path: './media/bg.png' }, objects: [] },
      { name: 'Color BG', background: { color: 'FF0000' }, objects: [] },
    ];
    const colors = extractBackgroundColors(masterData);
    expect(colors).toEqual(['FF0000']);
  });

  it('returns empty array for empty masterData', () => {
    expect(extractBackgroundColors([])).toEqual([]);
    expect(extractBackgroundColors(null)).toEqual([]);
  });

  it('returns empty when all backgrounds are white', () => {
    const masterData = [
      { name: 'A', background: { color: 'FFFFFF' }, objects: [] },
      { name: 'B', background: { color: 'FEFEFE' }, objects: [] },
    ];
    const colors = extractBackgroundColors(masterData);
    expect(colors).toHaveLength(0);
  });

  it('returns empty when only image backgrounds', () => {
    const masterData = [
      { name: 'A', background: { path: 'a.png' }, objects: [] },
      { name: 'B', background: { path: 'b.png' }, objects: [] },
    ];
    const colors = extractBackgroundColors(masterData);
    expect(colors).toHaveLength(0);
  });
});

describe('generateMastersCode (limited palette)', () => {
  it('adds limited comment when palette is limited', () => {
    const code = generateMastersCode(makeLimitedMasterData(), {
      templateName: 'canyon.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    expect(code).toContain('Theme accents are limited');
  });

  it('uses background colors for THEME.brand', () => {
    const code = generateMastersCode(makeLimitedMasterData(), {
      templateName: 'canyon.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    expect(code).toContain("brand: '038DAF'");
    expect(code).toContain('from master backgrounds');
  });

  it('generates SLIDE_COLORS array', () => {
    const code = generateMastersCode(makeLimitedMasterData(), {
      templateName: 'canyon.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    expect(code).toContain('const SLIDE_COLORS = [');
    expect(code).toContain("'038DAF'");
  });

  it('exports SLIDE_COLORS when limited', () => {
    const code = generateMastersCode(makeLimitedMasterData(), {
      templateName: 'canyon.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    expect(code).toContain('SLIDE_COLORS');
    expect(code).toMatch(/export.*SLIDE_COLORS/);
  });

  it('uses background colors for CHART_COLORS', () => {
    const code = generateMastersCode(makeLimitedMasterData(), {
      templateName: 'canyon.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    expect(code).toContain('from master backgrounds');
    // Chart colors should contain bg colors, not white
    const chartSection = code.split('CHART_COLORS')[1];
    expect(chartSection).toContain('038DAF');
    expect(chartSection).not.toMatch(/'FFFFFF',\s*\/\/ accent/);
  });

  it('does NOT add SLIDE_COLORS for healthy palette', () => {
    const code = generateMastersCode(makeMasterData(), {
      templateName: 'test.potx',
      dimensions,
      themeColors,
      themeFonts,
    });
    expect(code).not.toContain('SLIDE_COLORS');
    expect(code).not.toContain('accents are limited');
  });

  it('generates valid JavaScript when limited', () => {
    const code = generateMastersCode(makeLimitedMasterData(), {
      templateName: 'canyon.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    const testable = code.replace(/^export /gm, '').replace(/^import .*/gm, '');
    expect(() => new Function(testable)).not.toThrow();
  });
});

describe('generateAgentInstructions (limited palette)', () => {
  it('adds limited note in Theme Colors section', () => {
    const md = generateAgentInstructions(
      makeLimitedMasterData(),
      limitedThemeColors,
      themeFonts,
      dimensions,
      'canyon.pptx',
    );
    expect(md).toContain('Theme accents are limited');
    expect(md).toContain('derived from master slide backgrounds');
  });

  it('shows SLIDE_COLORS in exports table', () => {
    const md = generateAgentInstructions(
      makeLimitedMasterData(),
      limitedThemeColors,
      themeFonts,
      dimensions,
      'canyon.pptx',
    );
    expect(md).toContain('SLIDE_COLORS');
  });

  it('shows background colors in CHART_COLORS section', () => {
    const md = generateAgentInstructions(
      makeLimitedMasterData(),
      limitedThemeColors,
      themeFonts,
      dimensions,
      'canyon.pptx',
    );
    expect(md).toContain('038DAF');
    expect(md).toContain('Colors from master backgrounds');
  });

  it('adds dark background note in Key Conventions', () => {
    const md = generateAgentInstructions(
      makeLimitedMasterData(),
      limitedThemeColors,
      themeFonts,
      dimensions,
      'canyon.pptx',
    );
    expect(md).toContain('THEME.background');
    expect(md).toContain('Dark backgrounds');
  });

  it('does NOT add limited notes for healthy palette', () => {
    const md = generateAgentInstructions(
      makeMasterData(),
      themeColors,
      themeFonts,
      dimensions,
      'test.potx',
    );
    expect(md).not.toContain('accents are limited');
    expect(md).not.toContain('SLIDE_COLORS');
  });
});

describe('generateThemeJson (limited palette)', () => {
  it('includes slideColors when limited', () => {
    const result = generateThemeJson(limitedThemeColors, themeFonts, dimensions, makeLimitedMasterData());
    expect(result.slideColors).toBeDefined();
    expect(result.slideColors).toContain('038DAF');
    expect(result.paletteSource).toBe('background-fallback');
  });

  it('does NOT include slideColors for healthy palette', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions, makeMasterData());
    expect(result.slideColors).toBeUndefined();
    expect(result.paletteSource).toBeUndefined();
  });

  it('palette uses effective colors when limited', () => {
    const result = generateThemeJson(limitedThemeColors, themeFonts, dimensions, makeLimitedMasterData());
    // accent1 in palette should not be white
    expect(result.palette.accent1.base).not.toBe('FFFFFF');
  });

  it('backwards compatible when called without masterData', () => {
    const result = generateThemeJson(themeColors, themeFonts, dimensions);
    expect(result.colors.dk1).toBe('000000');
    expect(result.palette).toBeDefined();
  });
});

describe('getMasterDescription improvements', () => {
  it('includes background color indicator in descriptions', () => {
    const md = generateAgentInstructions(
      makeLimitedMasterData(),
      limitedThemeColors,
      themeFonts,
      dimensions,
      'canyon.pptx',
    );
    // Dark bg masters should have "(dark bg)" tag
    expect(md).toContain('(dark bg)');
    // Light bg masters should have "(light bg)" tag
    expect(md).toContain('(light bg)');
  });

  it('includes column count for multi-column layouts', () => {
    const md = generateAgentInstructions(
      makeLimitedMasterData(),
      limitedThemeColors,
      themeFonts,
      dimensions,
      'canyon.pptx',
    );
    // Two Content has 2 body placeholders with h > 2"
    expect(md).toContain('2-column');
  });
});

describe('edge cases', () => {
  it('handles limited palette with no background colors', () => {
    const masterData = [{
      name: 'Only Images',
      background: { path: './media/bg.png' },
      slideNumber: null,
      objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1 } } }],
    }];
    const code = generateMastersCode(masterData, {
      templateName: 'test.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    // Should not crash, should not add SLIDE_COLORS (no bg colors to use)
    expect(code).not.toContain('SLIDE_COLORS');
    expect(code).not.toContain('accents are limited');
  });

  it('handles single background color fallback', () => {
    const masterData = [
      { name: 'A', background: { color: '336699' }, slideNumber: null, objects: [] },
      { name: 'B', background: { color: 'FFFFFF' }, slideNumber: null, objects: [] },
    ];
    const code = generateMastersCode(masterData, {
      templateName: 'test.pptx',
      dimensions,
      themeColors: limitedThemeColors,
      themeFonts,
    });
    expect(code).toContain("brand: '336699'");
    expect(code).toContain('SLIDE_COLORS');
  });
});
