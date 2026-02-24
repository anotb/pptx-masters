import { describe, it, expect } from 'vitest';
import { createColorResolver } from '../src/parser/colors.js';
import {
  mapShape,
  resolveFill,
  resolveLine,
  resolveShadow,
  mapTextPropsToOptions,
} from '../src/mapper/shapes.js';
import { mapPlaceholder } from '../src/mapper/placeholders.js';
import { mapBackground } from '../src/mapper/backgrounds.js';
import { mapSlideNumberAndFooters, extractSlideNumberFromShape } from '../src/mapper/slideNumber.js';

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

// --- resolveFill ---

describe('resolveFill', () => {
  it('resolves solidFill with srgbClr to color', () => {
    const fill = {
      type: 'solidFill',
      element: { 'a:srgbClr': { '@_val': 'FF5500' } },
    };
    const { result, warnings } = resolveFill(fill, makeResolver());
    expect(result).toEqual({ color: 'FF5500' });
    expect(warnings).toHaveLength(0);
  });

  it('resolves solidFill with schemeClr', () => {
    const fill = {
      type: 'solidFill',
      element: { 'a:schemeClr': { '@_val': 'accent1' } },
    };
    const { result } = resolveFill(fill, makeResolver());
    expect(result).toEqual({ color: '4472C4' });
  });

  it('resolves solidFill with alpha transparency', () => {
    const fill = {
      type: 'solidFill',
      element: {
        'a:srgbClr': {
          '@_val': 'FF0000',
          'a:alpha': { '@_val': '50000' },
        },
      },
    };
    const { result } = resolveFill(fill, makeResolver());
    expect(result.color).toBe('FF0000');
    expect(result.transparency).toBe(50);
  });

  it('resolves gradFill using first stop + warning', () => {
    const fill = {
      type: 'gradFill',
      element: {
        'a:gsLst': {
          'a:gs': [
            { '@_pos': '0', 'a:srgbClr': { '@_val': 'FF0000' } },
            { '@_pos': '100000', 'a:srgbClr': { '@_val': '0000FF' } },
          ],
        },
      },
    };
    const { result, warnings } = resolveFill(fill, makeResolver());
    expect(result.color).toBe('FF0000');
    expect(warnings.length).toBeGreaterThan(0);
    expect(warnings[0]).toContain('Gradient');
  });

  it('resolves gradFill with single stop', () => {
    const fill = {
      type: 'gradFill',
      element: {
        'a:gsLst': {
          'a:gs': { '@_pos': '0', 'a:srgbClr': { '@_val': '00FF00' } },
        },
      },
    };
    const { result, warnings } = resolveFill(fill, makeResolver());
    expect(result.color).toBe('00FF00');
    expect(warnings.length).toBeGreaterThan(0);
  });

  it('returns null for noFill', () => {
    const fill = { type: 'noFill', element: {} };
    const { result, warnings } = resolveFill(fill, makeResolver());
    expect(result).toBeNull();
    expect(warnings).toHaveLength(0);
  });

  it('returns null for null fill', () => {
    const { result } = resolveFill(null, makeResolver());
    expect(result).toBeNull();
  });

  it('resolves blipFill with image reference', () => {
    const fill = {
      type: 'blipFill',
      element: {
        'a:blip': { '@_r:embed': 'rId2' },
      },
    };
    const { result } = resolveFill(fill, makeResolver());
    expect(result).toEqual({ imageRef: 'rId2' });
  });

  it('resolves pattFill using foreground color + warning', () => {
    const fill = {
      type: 'pattFill',
      element: {
        'a:fgClr': { 'a:srgbClr': { '@_val': 'AABBCC' } },
      },
    };
    const { result, warnings } = resolveFill(fill, makeResolver());
    expect(result.color).toBe('AABBCC');
    expect(warnings[0]).toContain('Pattern');
  });

  it('handles pattFill with no foreground color', () => {
    const fill = {
      type: 'pattFill',
      element: {},
    };
    const { result, warnings } = resolveFill(fill, makeResolver());
    expect(result).toBeNull();
    expect(warnings[0]).toContain('Pattern');
  });
});

// --- resolveLine ---

describe('resolveLine', () => {
  it('resolves line width from EMU to points', () => {
    const lineEl = { '@_w': '25400' }; // 2pt
    const result = resolveLine(lineEl, makeResolver());
    expect(result.width).toBe(2);
  });

  it('resolves line color from solidFill', () => {
    const lineEl = {
      'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.color).toBe('FF0000');
  });

  it('resolves dash type', () => {
    const lineEl = {
      '@_w': '12700',
      'a:prstDash': { '@_val': 'dash' },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.dashType).toBe('dash');
    expect(result.width).toBe(1);
  });

  it('resolves sysDot dash type', () => {
    const lineEl = {
      'a:prstDash': { '@_val': 'sysDot' },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.dashType).toBe('sysDot');
  });

  it('resolves arrow heads', () => {
    const lineEl = {
      '@_w': '12700',
      'a:headEnd': { '@_type': 'triangle' },
      'a:tailEnd': { '@_type': 'stealth' },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.beginArrowType).toBe('triangle');
    expect(result.endArrowType).toBe('stealth');
  });

  it('ignores none arrow type', () => {
    const lineEl = {
      '@_w': '12700',
      'a:headEnd': { '@_type': 'none' },
      'a:tailEnd': { '@_type': 'none' },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.beginArrowType).toBeUndefined();
    expect(result.endArrowType).toBeUndefined();
  });

  it('returns null for null lineEl', () => {
    const result = resolveLine(null, makeResolver());
    expect(result).toBeNull();
  });

  it('returns null for empty lineEl with no properties', () => {
    const result = resolveLine({}, makeResolver());
    expect(result).toBeNull();
  });

  it('resolves scheme color in line', () => {
    const lineEl = {
      'a:solidFill': { 'a:schemeClr': { '@_val': 'tx1' } },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.color).toBe('000000');
  });

  it('resolves all properties together', () => {
    const lineEl = {
      '@_w': '38100', // 3pt
      'a:solidFill': { 'a:srgbClr': { '@_val': '0000FF' } },
      'a:prstDash': { '@_val': 'lgDash' },
      'a:headEnd': { '@_type': 'arrow' },
      'a:tailEnd': { '@_type': 'diamond' },
    };
    const result = resolveLine(lineEl, makeResolver());
    expect(result.width).toBe(3);
    expect(result.color).toBe('0000FF');
    expect(result.dashType).toBe('lgDash');
    expect(result.beginArrowType).toBe('arrow');
    expect(result.endArrowType).toBe('diamond');
  });
});

// --- resolveShadow ---

describe('resolveShadow', () => {
  it('resolves outer shadow', () => {
    const effectLst = {
      'a:outerShdw': {
        '@_blurRad': '50800',  // ~4pt
        '@_dist': '38100',     // ~3pt
        '@_dir': '2700000',    // 45 degrees
        'a:srgbClr': { '@_val': '000000', 'a:alpha': { '@_val': '40000' } },
      },
    };
    const result = resolveShadow(effectLst, makeResolver());
    expect(result.type).toBe('outer');
    expect(result.blur).toBe(4);
    expect(result.offset).toBe(3);
    expect(result.angle).toBe(45);
    expect(result.color).toBe('000000');
    expect(result.opacity).toBe(0.4);
  });

  it('resolves inner shadow', () => {
    const effectLst = {
      'a:innerShdw': {
        '@_blurRad': '25400',
        '@_dist': '12700',
        '@_dir': '5400000',
        'a:srgbClr': { '@_val': '333333' },
      },
    };
    const result = resolveShadow(effectLst, makeResolver());
    expect(result.type).toBe('inner');
    expect(result.blur).toBe(2);
    expect(result.offset).toBe(1);
    expect(result.angle).toBe(90);
    expect(result.color).toBe('333333');
  });

  it('returns null for null effectLst', () => {
    expect(resolveShadow(null, makeResolver())).toBeNull();
  });

  it('returns null for empty effectLst', () => {
    expect(resolveShadow({}, makeResolver())).toBeNull();
  });

  it('prefers outer shadow over inner shadow when both present', () => {
    const effectLst = {
      'a:outerShdw': {
        '@_blurRad': '25400',
        'a:srgbClr': { '@_val': 'FF0000' },
      },
      'a:innerShdw': {
        '@_blurRad': '12700',
        'a:srgbClr': { '@_val': '0000FF' },
      },
    };
    const result = resolveShadow(effectLst, makeResolver());
    expect(result.type).toBe('outer');
    expect(result.color).toBe('FF0000');
  });
});

// --- mapTextPropsToOptions ---

describe('mapTextPropsToOptions', () => {
  it('extracts body props and first paragraph styling', () => {
    const textProps = {
      bodyProps: {
        margin: [0.05, 0.1, 0.05, 0.1],
        valign: 'middle',
      },
      paragraphs: [
        {
          align: 'center',
          lineSpacing: 18,
          paraSpaceBefore: 6,
          paraSpaceAfter: 12,
          runs: [
            { text: 'Hello', fontFace: 'Arial', fontSize: 24, color: 'FF0000', bold: true, italic: false },
          ],
        },
      ],
      plainText: 'Hello',
    };

    const opts = mapTextPropsToOptions(textProps);
    expect(opts.margin).toEqual([0.05, 0.1, 0.05, 0.1]);
    expect(opts.valign).toBe('middle');
    expect(opts.align).toBe('center');
    expect(opts.fontFace).toBe('Arial');
    expect(opts.fontSize).toBe(24);
    expect(opts.color).toBe('FF0000');
    expect(opts.bold).toBe(true);
    expect(opts.lineSpacing).toBe(18);
    expect(opts.paraSpaceBefore).toBe(6);
    expect(opts.paraSpaceAfter).toBe(12);
  });

  it('prefers defaultRunProps over first run', () => {
    const textProps = {
      bodyProps: { margin: [0, 0, 0, 0] },
      paragraphs: [
        {
          align: 'left',
          defaultRunProps: { fontFace: 'Helvetica', fontSize: 18, bold: false, italic: true },
          runs: [
            { text: 'Text', fontFace: 'Arial', fontSize: 24, bold: true, italic: false },
          ],
        },
      ],
      plainText: 'Text',
    };

    const opts = mapTextPropsToOptions(textProps);
    expect(opts.fontFace).toBe('Helvetica');
    expect(opts.fontSize).toBe(18);
    expect(opts.italic).toBe(true);
  });

  it('returns empty object for null textProps', () => {
    const opts = mapTextPropsToOptions(null);
    expect(opts).toEqual({});
  });

  it('handles empty paragraphs array', () => {
    const textProps = {
      bodyProps: { margin: [0, 0, 0, 0], valign: 'top' },
      paragraphs: [],
      plainText: '',
    };
    const opts = mapTextPropsToOptions(textProps);
    expect(opts.margin).toEqual([0, 0, 0, 0]);
    expect(opts.valign).toBe('top');
    expect(opts.fontFace).toBeUndefined();
  });

  it('uses lstStyle properties directly when paragraphs array is empty', () => {
    const textProps = {
      bodyProps: { margin: [0.05, 0.1, 0.05, 0.1], valign: 'top' },
      paragraphs: [],
      plainText: '',
      lstStyleProps: {
        1: {
          align: 'center',
          lineSpacing: 20,
          paraSpaceBefore: 4,
          paraSpaceAfter: 8,
          defaultRunProps: {
            fontFace: 'Helvetica',
            fontSize: 14,
            color: 'FF0000',
            bold: true,
          },
        },
      },
    };

    const opts = mapTextPropsToOptions(textProps);
    expect(opts.margin).toEqual([0.05, 0.1, 0.05, 0.1]);
    expect(opts.valign).toBe('top');
    expect(opts.align).toBe('center');
    expect(opts.fontFace).toBe('Helvetica');
    expect(opts.fontSize).toBe(14);
    expect(opts.color).toBe('FF0000');
    expect(opts.bold).toBe(true);
    expect(opts.lineSpacing).toBe(20);
    expect(opts.paraSpaceBefore).toBe(4);
    expect(opts.paraSpaceAfter).toBe(8);
  });

  it('extracts lineSpacingMultiple', () => {
    const textProps = {
      bodyProps: {},
      paragraphs: [
        {
          align: 'left',
          lineSpacingMultiple: 1.5,
          runs: [{ text: 'X' }],
        },
      ],
      plainText: 'X',
    };
    const opts = mapTextPropsToOptions(textProps);
    expect(opts.lineSpacingMultiple).toBe(1.5);
  });
});

// --- mapShape ---

describe('mapShape', () => {
  const relationships = {
    rId1: { type: 'image', target: '../media/image1.png' },
    rId2: { type: 'image', target: '../media/logo.jpg' },
  };

  it('maps a rect shape', () => {
    const shape = {
      type: 'shape',
      name: 'Rectangle 1',
      position: { x: 1, y: 2, w: 3, h: 4 },
      geometry: 'rect',
      fill: { type: 'solidFill', element: { 'a:srgbClr': { '@_val': 'AABBCC' } } },
      line: null,
      textProps: null,
    };

    const { object, warnings } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.rect).toBeDefined();
    expect(object.rect.x).toBe(1);
    expect(object.rect.y).toBe(2);
    expect(object.rect.w).toBe(3);
    expect(object.rect.h).toBe(4);
    expect(object.rect.fill.color).toBe('AABBCC');
  });

  it('maps a roundRect with rectRadius', () => {
    const shape = {
      type: 'shape',
      name: 'RoundRect',
      position: { x: 0, y: 0, w: 2, h: 1 },
      geometry: 'roundRect',
      fill: { type: 'solidFill', element: { 'a:srgbClr': { '@_val': 'FF0000' } } },
      line: null,
      textProps: null,
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.rect.rectRadius).toBe(0.1);
    expect(object.rect.fill.color).toBe('FF0000');
  });

  it('maps a line shape', () => {
    const shape = {
      type: 'shape',
      name: 'Line 1',
      position: { x: 0.5, y: 1, w: 5, h: 0 },
      geometry: 'line',
      fill: null,
      line: {
        '@_w': '25400',
        'a:solidFill': { 'a:srgbClr': { '@_val': '333333' } },
      },
      textProps: null,
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.line).toBeDefined();
    expect(object.line.x).toBe(0.5);
    expect(object.line.y).toBe(1);
    expect(object.line.w).toBe(5);
    expect(object.line.h).toBe(0);
    expect(object.line.line.color).toBe('333333');
    expect(object.line.line.width).toBe(2);
  });

  it('maps a text box shape', () => {
    const shape = {
      type: 'shape',
      name: 'TextBox 1',
      position: { x: 1, y: 1, w: 4, h: 2 },
      geometry: 'rect',
      fill: { type: 'solidFill', element: { 'a:srgbClr': { '@_val': 'EEEEEE' } } },
      line: null,
      textProps: {
        bodyProps: { margin: [0.05, 0.1, 0.05, 0.1], valign: 'middle' },
        paragraphs: [
          {
            align: 'center',
            runs: [{ text: 'Hello World', fontFace: 'Arial', fontSize: 18, color: '000000', bold: true, italic: false }],
          },
        ],
        plainText: 'Hello World',
      },
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.text).toBeDefined();
    expect(object.text.options.x).toBe(1);
    expect(object.text.options.y).toBe(1);
    expect(object.text.options.w).toBe(4);
    expect(object.text.options.h).toBe(2);
    expect(object.text.options.fill.color).toBe('EEEEEE');
    expect(object.text.options.align).toBe('center');
    // Text is flattened to string for PptxGenJS defineSlideMaster compatibility
    expect(typeof object.text.text).toBe('string');
    expect(object.text.text).toBe('Hello World');
  });

  it('maps a picture shape', () => {
    const shape = {
      type: 'picture',
      name: 'Picture 1',
      position: { x: 2, y: 3, w: 4, h: 3 },
      geometry: undefined,
      fill: null,
      line: null,
      textProps: null,
      imageRef: 'rId1',
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.image).toBeDefined();
    expect(object.image.path).toBe('./media/image1.png');
    expect(object.image.x).toBe(2);
    expect(object.image.y).toBe(3);
    expect(object.image.w).toBe(4);
    expect(object.image.h).toBe(3);
  });

  it('includes rotation on shapes', () => {
    const shape = {
      type: 'shape',
      name: 'Rotated',
      position: { x: 0, y: 0, w: 2, h: 2 },
      rotation: 45,
      geometry: 'rect',
      fill: null,
      line: null,
      textProps: null,
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.rect.rotate).toBe(45);
  });

  it('includes rotation on images', () => {
    const shape = {
      type: 'picture',
      name: 'Rotated Image',
      position: { x: 0, y: 0, w: 2, h: 2 },
      rotation: 90,
      geometry: undefined,
      fill: null,
      line: null,
      textProps: null,
      imageRef: 'rId2',
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.image.rotate).toBe(90);
  });

  it('warns on unresolved image reference', () => {
    const shape = {
      type: 'picture',
      name: 'Missing Image',
      position: { x: 0, y: 0, w: 1, h: 1 },
      geometry: undefined,
      fill: null,
      line: null,
      textProps: null,
      imageRef: 'rId999',
    };

    const { warnings } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(warnings.some((w) => w.includes('rId999'))).toBe(true);
  });

  it('maps text box with multiple paragraphs using paragraph breaks', () => {
    const shape = {
      type: 'shape',
      name: 'MultiPara',
      position: { x: 1, y: 1, w: 8, h: 3 },
      geometry: 'rect',
      fill: null,
      line: null,
      textProps: {
        bodyProps: {},
        paragraphs: [
          {
            align: 'left',
            runs: [{ text: 'First paragraph', fontFace: 'Arial', fontSize: 12 }],
          },
          {
            align: 'left',
            runs: [{ text: 'Second paragraph', fontFace: 'Arial', fontSize: 12 }],
          },
        ],
        plainText: 'First paragraph\nSecond paragraph',
      },
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, {});
    expect(object.text).toBeDefined();
    // Text is flattened to string with \n for paragraph breaks
    expect(typeof object.text.text).toBe('string');
    expect(object.text.text).toBe('First paragraph\nSecond paragraph');
  });

  it('maps text box with line breaks between runs', () => {
    const shape = {
      type: 'shape',
      name: 'LineBreak',
      position: { x: 1, y: 1, w: 8, h: 2 },
      geometry: 'rect',
      fill: null,
      line: null,
      textProps: {
        bodyProps: {},
        paragraphs: [
          {
            align: 'left',
            runs: [
              { text: 'Before break', fontSize: 12 },
              { text: '', isBreak: true },
              { text: 'After break', fontSize: 12 },
            ],
          },
        ],
        plainText: 'Before break\nAfter break',
      },
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, {});
    // Text is flattened with \n for line breaks
    expect(typeof object.text.text).toBe('string');
    expect(object.text.text).toBe('Before break\nAfter break');
  });

  it('maps text box with field runs (e.g. slide number field)', () => {
    const shape = {
      type: 'shape',
      name: 'FieldBox',
      position: { x: 1, y: 1, w: 4, h: 1 },
      geometry: 'rect',
      fill: null,
      line: null,
      textProps: {
        bodyProps: {},
        paragraphs: [
          {
            align: 'center',
            runs: [
              { text: 'Page ', fontSize: 10 },
              { text: '5', fontSize: 10, isField: true, fieldType: 'slidenum' },
            ],
          },
        ],
        plainText: 'Page 5',
      },
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, {});
    // Text is flattened to string (field text included as literal)
    expect(typeof object.text.text).toBe('string');
    expect(object.text.text).toBe('Page 5');
  });

  it('handles null shape', () => {
    const { object, warnings } = mapShape(null, makeResolver(), themeFonts, relationships);
    expect(object).toBeNull();
    expect(warnings.length).toBeGreaterThan(0);
  });

  it('maps rect with line properties', () => {
    const shape = {
      type: 'shape',
      name: 'Bordered Rect',
      position: { x: 0, y: 0, w: 3, h: 2 },
      geometry: 'rect',
      fill: null,
      line: {
        '@_w': '19050',
        'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
        'a:prstDash': { '@_val': 'dash' },
      },
      textProps: null,
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.rect.line).toBeDefined();
    expect(object.rect.line.color).toBe('000000');
    expect(object.rect.line.width).toBe(1.5);
    expect(object.rect.line.dashType).toBe('dash');
  });

  it('handles shape with noFill', () => {
    const shape = {
      type: 'shape',
      name: 'NoFill',
      position: { x: 0, y: 0, w: 1, h: 1 },
      geometry: 'rect',
      fill: { type: 'noFill', element: {} },
      line: null,
      textProps: null,
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.rect.fill).toBeUndefined();
  });

  it('maps shape with gradient fill + warning', () => {
    const shape = {
      type: 'shape',
      name: 'Gradient',
      position: { x: 0, y: 0, w: 2, h: 2 },
      geometry: 'rect',
      fill: {
        type: 'gradFill',
        element: {
          'a:gsLst': {
            'a:gs': [
              { '@_pos': '0', 'a:srgbClr': { '@_val': 'FF0000' } },
              { '@_pos': '100000', 'a:srgbClr': { '@_val': '0000FF' } },
            ],
          },
        },
      },
      line: null,
      textProps: null,
    };

    const { object, warnings } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.rect.fill.color).toBe('FF0000');
    expect(warnings.some((w) => w.includes('Gradient'))).toBe(true);
  });

  it('maps line shape with no line properties', () => {
    const shape = {
      type: 'shape',
      name: 'Bare Line',
      position: { x: 0, y: 0, w: 5, h: 0 },
      geometry: 'line',
      fill: null,
      line: null,
      textProps: null,
    };

    const { object } = mapShape(shape, makeResolver(), themeFonts, relationships);
    expect(object.line).toBeDefined();
    expect(object.line.x).toBe(0);
    expect(object.line.w).toBe(5);
  });
});

// --- mapPlaceholder ---

describe('mapPlaceholder', () => {
  it('maps title placeholder', () => {
    const ph = {
      type: 'title',
      idx: 0,
      name: 'Title 1',
      position: { x: 0.5, y: 0.5, w: 9, h: 1.5 },
      textProps: {
        bodyProps: { margin: [0.05, 0.1, 0.05, 0.1], valign: 'bottom' },
        paragraphs: [
          {
            align: 'left',
            runs: [{ text: '', fontFace: 'Arial', fontSize: 36, color: '000000', bold: true, italic: false }],
          },
        ],
        plainText: '',
      },
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result).not.toBeNull();
    expect(result.placeholder.options.type).toBe('title');
    expect(result.placeholder.options.name).toBe('Title 1');
    expect(result.placeholder.options.x).toBe(0.5);
    expect(result.placeholder.options.y).toBe(0.5);
    expect(result.placeholder.options.w).toBe(9);
    expect(result.placeholder.options.h).toBe(1.5);
    expect(result.placeholder.options.fontFace).toBe('Arial');
    expect(result.placeholder.options.fontSize).toBe(36);
    expect(result.placeholder.options.bold).toBe(true);
  });

  it('maps ctrTitle with align:center', () => {
    const ph = {
      type: 'ctrTitle',
      idx: 0,
      name: 'Center Title',
      position: { x: 1, y: 2, w: 8, h: 1 },
      textProps: {
        bodyProps: {},
        paragraphs: [
          { align: 'left', runs: [{ text: '', fontSize: 40 }] },
        ],
        plainText: '',
      },
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('title');
    expect(result.placeholder.options.align).toBe('center');
  });

  it('maps body placeholder', () => {
    const ph = {
      type: 'body',
      idx: 1,
      name: 'Content Placeholder',
      position: { x: 0.5, y: 2, w: 9, h: 4.5 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('body');
  });

  it('maps subTitle to body type', () => {
    const ph = {
      type: 'subTitle',
      idx: 1,
      name: 'Subtitle',
      position: { x: 1, y: 3, w: 8, h: 1 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('body');
  });

  it('maps obj to body type', () => {
    const ph = {
      type: 'obj',
      idx: 2,
      name: 'Object Placeholder',
      position: { x: 0, y: 0, w: 5, h: 5 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('body');
  });

  it('maps pic placeholder', () => {
    const ph = {
      type: 'pic',
      idx: 3,
      name: 'Picture Placeholder',
      position: { x: 1, y: 1, w: 4, h: 4 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('pic');
  });

  it('maps chart placeholder', () => {
    const ph = {
      type: 'chart',
      idx: 4,
      name: 'Chart Placeholder',
      position: { x: 0, y: 0, w: 5, h: 3 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('chart');
  });

  it('maps tbl placeholder', () => {
    const ph = {
      type: 'tbl',
      idx: 5,
      name: 'Table Placeholder',
      position: { x: 0, y: 0, w: 6, h: 4 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('tbl');
  });

  it('maps media placeholder', () => {
    const ph = {
      type: 'media',
      idx: 6,
      name: 'Media Placeholder',
      position: { x: 0, y: 0, w: 3, h: 2 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('media');
  });

  it('maps clipArt to pic', () => {
    const ph = {
      type: 'clipArt',
      idx: 7,
      name: 'ClipArt',
      position: { x: 0, y: 0, w: 2, h: 2 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('pic');
  });

  it('maps dgm to chart', () => {
    const ph = {
      type: 'dgm',
      idx: 8,
      name: 'Diagram',
      position: { x: 0, y: 0, w: 4, h: 3 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('chart');
  });

  it('skips sldNum type', () => {
    const ph = {
      type: 'sldNum',
      idx: 10,
      name: 'Slide Number',
      position: { x: 9, y: 7, w: 0.5, h: 0.3 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result).toBeNull();
  });

  it('skips ftr type', () => {
    const ph = {
      type: 'ftr',
      idx: 11,
      name: 'Footer',
      position: { x: 3, y: 7, w: 3, h: 0.3 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    expect(mapPlaceholder(ph, makeResolver(), themeFonts)).toBeNull();
  });

  it('skips dt type', () => {
    const ph = {
      type: 'dt',
      idx: 12,
      name: 'Date',
      position: { x: 0, y: 7, w: 2, h: 0.3 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    expect(mapPlaceholder(ph, makeResolver(), themeFonts)).toBeNull();
  });

  it('skips hdr type', () => {
    const ph = {
      type: 'hdr',
      idx: 13,
      name: 'Header',
      position: { x: 0, y: 0, w: 3, h: 0.3 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    expect(mapPlaceholder(ph, makeResolver(), themeFonts)).toBeNull();
  });

  it('returns null for null placeholder', () => {
    expect(mapPlaceholder(null, makeResolver(), themeFonts)).toBeNull();
  });

  it('includes fill and line from shapeProps', () => {
    const ph = {
      type: 'title',
      idx: 0,
      name: 'Styled Title',
      position: { x: 0.5, y: 0.5, w: 9, h: 1 },
      textProps: null,
      shapeProps: {
        geometry: 'rect',
        fill: { type: 'solidFill', element: { 'a:srgbClr': { '@_val': 'F0F0F0' } } },
        line: {
          '@_w': '12700',
          'a:solidFill': { 'a:srgbClr': { '@_val': '999999' } },
        },
      },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.fill.color).toBe('F0F0F0');
    expect(result.placeholder.options.line.color).toBe('999999');
    expect(result.placeholder.options.line.width).toBe(1);
  });

  it('includes rotation', () => {
    const ph = {
      type: 'body',
      idx: 1,
      name: 'Rotated Body',
      position: { x: 0, y: 0, w: 5, h: 3 },
      rotation: 15,
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.rotate).toBe(15);
  });

  it('maps placeholder with valign from textProps', () => {
    const ph = {
      type: 'title',
      idx: 0,
      name: 'Title',
      position: { x: 0, y: 0, w: 10, h: 1.5 },
      textProps: {
        bodyProps: { valign: 'bottom', margin: [0.05, 0.1, 0.05, 0.1] },
        paragraphs: [{ align: 'left', runs: [] }],
        plainText: '',
      },
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.valign).toBe('bottom');
    expect(result.placeholder.options.margin).toEqual([0.05, 0.1, 0.05, 0.1]);
  });

  it('handles unknown placeholder type with warning', () => {
    const ph = {
      type: 'unknownType',
      idx: 99,
      name: 'Unknown',
      position: { x: 0, y: 0, w: 1, h: 1 },
      textProps: null,
      shapeProps: { geometry: 'rect', fill: null, line: null },
    };

    const result = mapPlaceholder(ph, makeResolver(), themeFonts);
    expect(result.placeholder.options.type).toBe('body');
    expect(result.warnings.some((w) => w.includes('unknownType'))).toBe(true);
  });
});

// --- mapBackground ---

describe('mapBackground', () => {
  const relationships = {
    rId1: { type: 'image', target: '../media/bg-image.png' },
  };

  it('maps solid fill background', () => {
    const bg = {
      'a:solidFill': { 'a:srgbClr': { '@_val': '003366' } },
    };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toEqual({ color: '003366' });
    expect(warnings).toHaveLength(0);
  });

  it('maps solid fill with scheme color', () => {
    const bg = {
      'a:solidFill': { 'a:schemeClr': { '@_val': 'accent1' } },
    };
    const { background } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toEqual({ color: '4472C4' });
  });

  it('maps solid fill with transparency', () => {
    const bg = {
      'a:solidFill': {
        'a:srgbClr': {
          '@_val': 'FF0000',
          'a:alpha': { '@_val': '50000' },
        },
      },
    };
    const { background } = mapBackground(bg, makeResolver(), relationships);
    expect(background.color).toBe('FF0000');
    expect(background.transparency).toBe(50);
  });

  it('maps image fill background', () => {
    const bg = {
      'a:blipFill': {
        'a:blip': { '@_r:embed': 'rId1' },
      },
    };
    const { background } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toEqual({ path: './media/bg-image.png' });
  });

  it('warns on unresolved image reference', () => {
    const bg = {
      'a:blipFill': {
        'a:blip': { '@_r:embed': 'rId999' },
      },
    };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toBeNull();
    expect(warnings.length).toBeGreaterThan(0);
  });

  it('maps gradient fill with fallback + warning', () => {
    const bg = {
      'a:gradFill': {
        'a:gsLst': {
          'a:gs': [
            { '@_pos': '0', 'a:srgbClr': { '@_val': '1A2B3C' } },
            { '@_pos': '100000', 'a:srgbClr': { '@_val': 'FFFFFF' } },
          ],
        },
      },
    };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background.color).toBe('1A2B3C');
    expect(warnings.some((w) => w.includes('Gradient'))).toBe(true);
  });

  it('maps pattern fill with foreground color + warning', () => {
    const bg = {
      'a:pattFill': {
        'a:fgClr': { 'a:srgbClr': { '@_val': 'DDEEFF' } },
      },
    };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background.color).toBe('DDEEFF');
    expect(warnings.some((w) => w.includes('Pattern'))).toBe(true);
  });

  it('returns null for noFill', () => {
    const bg = { 'a:noFill': '' };
    const { background } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toBeNull();
  });

  it('returns null for null element', () => {
    const { background } = mapBackground(null, makeResolver(), relationships);
    expect(background).toBeNull();
  });

  it('handles bgRef with color resolution', () => {
    const bg = {
      bgRef: { 'a:srgbClr': { '@_val': 'ABCDEF' } },
    };
    const { background } = mapBackground(bg, makeResolver(), relationships);
    expect(background.color).toBe('ABCDEF');
  });

  it('warns on bgRef without resolvable color', () => {
    const bg = { bgRef: {} };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toBeNull();
    expect(warnings.some((w) => w.includes('bgRef'))).toBe(true);
  });

  it('handles gradient with single stop', () => {
    const bg = {
      'a:gradFill': {
        'a:gsLst': {
          'a:gs': { '@_pos': '0', 'a:srgbClr': { '@_val': 'AABB00' } },
        },
      },
    };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background.color).toBe('AABB00');
    expect(warnings.length).toBeGreaterThan(0);
  });

  it('handles pattern fill without foreground color', () => {
    const bg = {
      'a:pattFill': {},
    };
    const { background, warnings } = mapBackground(bg, makeResolver(), relationships);
    expect(background).toBeNull();
    expect(warnings.some((w) => w.includes('Pattern'))).toBe(true);
  });
});

// --- mapSlideNumberAndFooters ---

describe('mapSlideNumberAndFooters', () => {
  it('extracts slide number with position and styling', () => {
    const placeholders = [
      {
        type: 'sldNum',
        idx: 10,
        name: 'Slide Number',
        position: { x: 8.5, y: 6.8, w: 1, h: 0.3 },
        textProps: {
          bodyProps: { margin: [0, 0, 0, 0] },
          paragraphs: [
            {
              align: 'right',
              runs: [{ text: '', fontFace: 'Calibri', fontSize: 10, color: '888888' }],
            },
          ],
          plainText: '',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber, footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).not.toBeNull();
    expect(slideNumber.x).toBe(8.5);
    expect(slideNumber.y).toBe(6.8);
    expect(slideNumber.w).toBe(1);
    // Height is normalized to minimum 2.5x font size (10pt → 0.3472")
    expect(slideNumber.h).toBeCloseTo(0.3472, 3);
    expect(slideNumber.fontFace).toBe('Calibri');
    expect(slideNumber.fontSize).toBe(10);
    expect(slideNumber.color).toBe('888888');
    expect(slideNumber.align).toBe('right');
    expect(footerObjects).toHaveLength(0);
  });

  it('extracts footer text with content', () => {
    const placeholders = [
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 3, y: 6.8, w: 4, h: 0.3 },
        textProps: {
          bodyProps: { margin: [0.05, 0.1, 0.05, 0.1], valign: 'middle' },
          paragraphs: [
            {
              align: 'center',
              runs: [{ text: 'Confidential', fontFace: 'Arial', fontSize: 8, color: '666666' }],
            },
          ],
          plainText: 'Confidential',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber, footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).toBeNull();
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('Confidential');
    expect(footerObjects[0].text.options.x).toBe(3);
    expect(footerObjects[0].text.options.y).toBe(6.8);
    expect(footerObjects[0].text.options.fontFace).toBe('Arial');
    expect(footerObjects[0].text.options.fontSize).toBe(8);
    expect(footerObjects[0].text.options.align).toBe('center');
  });

  it('extracts date placeholder', () => {
    const placeholders = [
      {
        type: 'dt',
        idx: 12,
        name: 'Date',
        position: { x: 0.5, y: 6.8, w: 2, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [
            {
              align: 'left',
              runs: [{ text: '2024-01-15', fontFace: 'Calibri', fontSize: 8 }],
            },
          ],
          plainText: '2024-01-15',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('2024-01-15');
  });

  it('extracts header placeholder', () => {
    const placeholders = [
      {
        type: 'hdr',
        idx: 13,
        name: 'Header',
        position: { x: 0, y: 0, w: 5, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [
            { align: 'left', runs: [{ text: 'Company Name' }] },
          ],
          plainText: 'Company Name',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('Company Name');
  });

  it('returns null slideNumber when none present', () => {
    const placeholders = [
      {
        type: 'title',
        idx: 0,
        name: 'Title',
        position: { x: 0, y: 0, w: 10, h: 1 },
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber, footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).toBeNull();
    expect(footerObjects).toHaveLength(0);
  });

  it('handles mixed: sldNum + ftr', () => {
    const placeholders = [
      {
        type: 'title',
        idx: 0,
        name: 'Title',
        position: { x: 0, y: 0, w: 10, h: 1 },
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
      {
        type: 'sldNum',
        idx: 10,
        name: 'Slide Number',
        position: { x: 9, y: 7, w: 0.5, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [{ align: 'right', runs: [{ text: '', fontSize: 9 }] }],
          plainText: '',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 3, y: 7, w: 4, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [{ align: 'center', runs: [{ text: 'Footer Text' }] }],
          plainText: 'Footer Text',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber, footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).not.toBeNull();
    expect(slideNumber.x).toBe(9);
    expect(slideNumber.fontSize).toBe(9);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('Footer Text');
  });

  it('handles empty placeholders array', () => {
    const { slideNumber, footerObjects } = mapSlideNumberAndFooters([], makeResolver(), themeFonts);
    expect(slideNumber).toBeNull();
    expect(footerObjects).toHaveLength(0);
  });

  it('handles null placeholders', () => {
    const { slideNumber, footerObjects } = mapSlideNumberAndFooters(null, makeResolver(), themeFonts);
    expect(slideNumber).toBeNull();
    expect(footerObjects).toHaveLength(0);
  });

  it('detects copyright text in non-typed placeholders', () => {
    const placeholders = [
      {
        type: undefined,
        idx: undefined,
        name: 'TextBox 5',
        position: { x: 0.5, y: 7, w: 3, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [
            { align: 'left', runs: [{ text: '\u00a9 2024 ACME Corp', fontSize: 7 }] },
          ],
          plainText: '\u00a9 2024 ACME Corp',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('\u00a9 2024 ACME Corp');
  });

  it('detects "copyright" text in non-typed placeholders', () => {
    const placeholders = [
      {
        type: undefined,
        idx: undefined,
        name: 'TextBox 6',
        position: { x: 0.5, y: 7, w: 4, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [
            { align: 'left', runs: [{ text: 'Copyright 2024 ACME' }] },
          ],
          plainText: 'Copyright 2024 ACME',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('Copyright 2024 ACME');
  });

  it('does not detect non-copyright text in non-typed placeholders', () => {
    const placeholders = [
      {
        type: undefined,
        idx: undefined,
        name: 'TextBox 7',
        position: { x: 0, y: 0, w: 5, h: 1 },
        textProps: {
          bodyProps: {},
          paragraphs: [
            { align: 'left', runs: [{ text: 'Just some random text' }] },
          ],
          plainText: 'Just some random text',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(0);
  });

  it('handles slide number without textProps', () => {
    const placeholders = [
      {
        type: 'sldNum',
        idx: 10,
        name: 'Slide Number',
        position: { x: 9, y: 7, w: 0.5, h: 0.3 },
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).not.toBeNull();
    expect(slideNumber.x).toBe(9);
    expect(slideNumber.y).toBe(7);
    expect(slideNumber.fontFace).toBeUndefined();
  });

  it('handles footer without textProps', () => {
    const placeholders = [
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 3, y: 7, w: 4, h: 0.3 },
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('');
    expect(footerObjects[0].text.options.x).toBe(3);
  });

  it('extracts valign and margin from footer', () => {
    const placeholders = [
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 0, y: 7, w: 3, h: 0.3 },
        textProps: {
          bodyProps: { valign: 'middle', margin: [0.02, 0.05, 0.02, 0.05] },
          paragraphs: [
            { align: 'center', runs: [{ text: 'Footer' }] },
          ],
          plainText: 'Footer',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    // Footer elements always get valign='top' for consistent baseline alignment
    expect(footerObjects[0].text.options.valign).toBe('top');
    expect(footerObjects[0].text.options.margin).toEqual([0.02, 0.05, 0.02, 0.05]);
  });

  it('nulls out slide number with zero size', () => {
    const placeholders = [
      {
        type: 'sldNum',
        idx: 10,
        name: 'Slide Number',
        position: { x: 0, y: 0, w: 0, h: 0 },
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).toBeNull();
  });

  it('keeps slide number with valid size', () => {
    const placeholders = [
      {
        type: 'sldNum',
        idx: 10,
        name: 'Slide Number',
        position: { x: 9, y: 7, w: 0.5, h: 0.3 },
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { slideNumber } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(slideNumber).not.toBeNull();
    expect(slideNumber.w).toBe(0.5);
    expect(slideNumber.h).toBe(0.3);
  });

  it('skips empty footer with zero position', () => {
    const placeholders = [
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 0, y: 0, w: 0, h: 0 },
        textProps: {
          bodyProps: {},
          paragraphs: [],
          plainText: '',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(0);
  });

  it('skips empty date with null position', () => {
    const placeholders = [
      {
        type: 'dt',
        idx: 12,
        name: 'Date',
        position: null,
        textProps: null,
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(0);
  });

  it('keeps footer with content even at zero position', () => {
    const placeholders = [
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 0, y: 0, w: 0, h: 0 },
        textProps: {
          bodyProps: {},
          paragraphs: [
            { align: 'left', runs: [{ text: 'Important Footer' }] },
          ],
          plainText: 'Important Footer',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
    expect(footerObjects[0].text.text).toBe('Important Footer');
  });

  it('keeps footer with valid position even without text', () => {
    const placeholders = [
      {
        type: 'ftr',
        idx: 11,
        name: 'Footer',
        position: { x: 3, y: 7, w: 4, h: 0.3 },
        textProps: {
          bodyProps: {},
          paragraphs: [],
          plainText: '',
        },
        shapeProps: { geometry: 'rect', fill: null, line: null },
      },
    ];

    const { footerObjects } = mapSlideNumberAndFooters(placeholders, makeResolver(), themeFonts);
    expect(footerObjects).toHaveLength(1);
  });
});

// --- extractSlideNumberFromShape ---

describe('extractSlideNumberFromShape', () => {
  it('extracts slideNumber from a non-placeholder shape with slidenum field', () => {
    const shape = {
      type: 'shape',
      name: 'TextBox 9',
      position: { x: 12.5, y: 7.13, w: 0.34, h: 0.13 },
      textProps: {
        bodyProps: { margin: [0, 0, 0, 0] },
        paragraphs: [{
          align: 'right',
          _explicitAlign: 'right',
          level: 1,
          runs: [{
            text: '‹#›',
            fontFace: 'Aptos',
            fontSize: 8,
            color: '000000',
            isField: true,
            fieldType: 'slidenum',
          }],
        }],
        plainText: '‹#›',
        lstStyleProps: null,
      },
    };

    const result = extractSlideNumberFromShape(shape);
    expect(result).not.toBeNull();
    expect(result.x).toBe(12.5);
    expect(result.y).toBe(7.13);
    expect(result.w).toBe(0.34);
    // Height is normalized to minimum 2.5x font size (8pt → 0.2778")
    expect(result.h).toBeCloseTo(0.2778, 3);
    expect(result.fontFace).toBe('Aptos');
    expect(result.fontSize).toBe(8);
    expect(result.color).toBe('000000');
    expect(result.align).toBe('right');
  });

  it('returns null for shapes without slidenum field', () => {
    const shape = {
      type: 'shape',
      position: { x: 1, y: 1, w: 5, h: 1 },
      textProps: {
        bodyProps: {},
        paragraphs: [{
          align: 'left',
          level: 1,
          runs: [{ text: 'Regular text', fontSize: 12 }],
        }],
        plainText: 'Regular text',
        lstStyleProps: null,
      },
    };

    expect(extractSlideNumberFromShape(shape)).toBeNull();
  });

  it('returns null for shapes with slidenum field AND other text content', () => {
    const shape = {
      type: 'shape',
      position: { x: 1, y: 1, w: 5, h: 1 },
      textProps: {
        bodyProps: {},
        paragraphs: [{
          align: 'left',
          level: 1,
          runs: [
            { text: 'Page ', fontSize: 12 },
            { text: '‹#›', isField: true, fieldType: 'slidenum', fontSize: 12 },
          ],
        }],
        plainText: 'Page ‹#›',
        lstStyleProps: null,
      },
    };

    expect(extractSlideNumberFromShape(shape)).toBeNull();
  });

  it('returns null for null or shapeless input', () => {
    expect(extractSlideNumberFromShape(null)).toBeNull();
    expect(extractSlideNumberFromShape({})).toBeNull();
    expect(extractSlideNumberFromShape({ textProps: null })).toBeNull();
  });

  it('returns null for zero-size slidenum shapes', () => {
    const shape = {
      type: 'shape',
      position: { x: 0, y: 0, w: 0, h: 0 },
      textProps: {
        bodyProps: {},
        paragraphs: [{
          align: 'right',
          level: 1,
          runs: [{ text: '‹#›', isField: true, fieldType: 'slidenum' }],
        }],
        plainText: '‹#›',
        lstStyleProps: null,
      },
    };

    expect(extractSlideNumberFromShape(shape)).toBeNull();
  });
});
