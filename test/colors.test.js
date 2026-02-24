import { describe, it, expect } from 'vitest';
import {
  hexToRgb,
  rgbToHex,
  rgbToHsl,
  hslToRgb,
  applyColorModifiers,
  createColorResolver,
} from '../src/parser/colors.js';

// --- hexToRgb ---

describe('hexToRgb', () => {
  it('converts white', () => {
    expect(hexToRgb('FFFFFF')).toEqual([255, 255, 255]);
  });

  it('converts black', () => {
    expect(hexToRgb('000000')).toEqual([0, 0, 0]);
  });

  it('converts a dark blue', () => {
    expect(hexToRgb('003366')).toEqual([0, 51, 102]);
  });

  it('converts pure red', () => {
    expect(hexToRgb('FF0000')).toEqual([255, 0, 0]);
  });

  it('converts pure green', () => {
    expect(hexToRgb('00FF00')).toEqual([0, 255, 0]);
  });

  it('converts pure blue', () => {
    expect(hexToRgb('0000FF')).toEqual([0, 0, 255]);
  });
});

// --- rgbToHex ---

describe('rgbToHex', () => {
  it('converts white', () => {
    expect(rgbToHex(255, 255, 255)).toBe('FFFFFF');
  });

  it('converts black', () => {
    expect(rgbToHex(0, 0, 0)).toBe('000000');
  });

  it('converts a dark blue', () => {
    expect(rgbToHex(0, 51, 102)).toBe('003366');
  });

  it('converts with single-digit hex values', () => {
    expect(rgbToHex(1, 2, 3)).toBe('010203');
  });

  it('clamps values above 255', () => {
    expect(rgbToHex(300, 0, 0)).toBe('FF0000');
  });

  it('clamps values below 0', () => {
    expect(rgbToHex(-10, 0, 0)).toBe('000000');
  });
});

// --- rgbToHsl / hslToRgb roundtrip ---

describe('rgbToHsl', () => {
  it('converts white', () => {
    const [h, s, l] = rgbToHsl(255, 255, 255);
    expect(l).toBeCloseTo(1, 5);
    expect(s).toBeCloseTo(0, 5);
  });

  it('converts black', () => {
    const [h, s, l] = rgbToHsl(0, 0, 0);
    expect(l).toBeCloseTo(0, 5);
    expect(s).toBeCloseTo(0, 5);
  });

  it('converts pure red', () => {
    const [h, s, l] = rgbToHsl(255, 0, 0);
    expect(h).toBeCloseTo(0, 0);
    expect(s).toBeCloseTo(1, 5);
    expect(l).toBeCloseTo(0.5, 5);
  });

  it('converts pure green', () => {
    const [h, s, l] = rgbToHsl(0, 255, 0);
    expect(h).toBeCloseTo(120, 0);
    expect(s).toBeCloseTo(1, 5);
    expect(l).toBeCloseTo(0.5, 5);
  });

  it('converts pure blue', () => {
    const [h, s, l] = rgbToHsl(0, 0, 255);
    expect(h).toBeCloseTo(240, 0);
    expect(s).toBeCloseTo(1, 5);
    expect(l).toBeCloseTo(0.5, 5);
  });

  it('converts a mid-tone gray', () => {
    const [h, s, l] = rgbToHsl(128, 128, 128);
    expect(s).toBeCloseTo(0, 5);
    expect(l).toBeCloseTo(0.502, 1);
  });
});

describe('hslToRgb', () => {
  it('converts white', () => {
    expect(hslToRgb(0, 0, 1)).toEqual([255, 255, 255]);
  });

  it('converts black', () => {
    expect(hslToRgb(0, 0, 0)).toEqual([0, 0, 0]);
  });

  it('converts pure red', () => {
    expect(hslToRgb(0, 1, 0.5)).toEqual([255, 0, 0]);
  });

  it('converts pure green', () => {
    expect(hslToRgb(120, 1, 0.5)).toEqual([0, 255, 0]);
  });

  it('converts pure blue', () => {
    expect(hslToRgb(240, 1, 0.5)).toEqual([0, 0, 255]);
  });
});

describe('RGB-HSL roundtrip', () => {
  const testCases = [
    { name: 'red', rgb: [255, 0, 0] },
    { name: 'green', rgb: [0, 255, 0] },
    { name: 'blue', rgb: [0, 0, 255] },
    { name: 'white', rgb: [255, 255, 255] },
    { name: 'black', rgb: [0, 0, 0] },
    { name: 'mid-tone', rgb: [100, 150, 200] },
    { name: 'dark blue', rgb: [0, 51, 102] },
  ];

  for (const { name, rgb } of testCases) {
    it(`roundtrips ${name}`, () => {
      const [h, s, l] = rgbToHsl(...rgb);
      const result = hslToRgb(h, s, l);
      // Allow +-1 for rounding
      expect(result[0]).toBeCloseTo(rgb[0], 0);
      expect(result[1]).toBeCloseTo(rgb[1], 0);
      expect(result[2]).toBeCloseTo(rgb[2], 0);
    });
  }
});

// --- Tint / shade calculations ---

describe('applyColorModifiers', () => {
  it('produces a lighter tint with lumMod=40000 lumOff=60000', () => {
    // base='003366' (dark blue) → should become significantly lighter
    const result = applyColorModifiers('003366', { lumMod: 40000, lumOff: 60000 });
    const [r, g, b] = hexToRgb(result);
    const [, , origL] = rgbToHsl(0, 51, 102);
    const [, , newL] = rgbToHsl(r, g, b);

    // New luminance should be much higher than the original
    expect(newL).toBeGreaterThan(origL);
    expect(newL).toBeGreaterThan(0.6); // Should be quite light
  });

  it('produces a darker shade with lumMod=75000', () => {
    // base='003366' → should become darker
    const result = applyColorModifiers('003366', { lumMod: 75000 });
    const [r, g, b] = hexToRgb(result);
    const [, , origL] = rgbToHsl(0, 51, 102);
    const [, , newL] = rgbToHsl(r, g, b);

    expect(newL).toBeLessThan(origL);
  });

  it('preserves color with lumMod=100000 (100%)', () => {
    const result = applyColorModifiers('4472C4', { lumMod: 100000 });
    expect(result).toBe('4472C4');
  });

  it('goes to black with lumMod=0', () => {
    const result = applyColorModifiers('4472C4', { lumMod: 0 });
    // Luminance 0 = black
    expect(result).toBe('000000');
  });

  it('applies saturation modification', () => {
    // Reduce saturation by half
    const result = applyColorModifiers('FF0000', { satMod: 50000 });
    const [r, g, b] = hexToRgb(result);
    const [, s] = rgbToHsl(r, g, b);
    // Original pure red has S=1, after 50% satMod should be ~0.5
    expect(s).toBeCloseTo(0.5, 1);
  });

  it('handles combined lumMod + lumOff + satMod', () => {
    const result = applyColorModifiers('003366', { lumMod: 50000, lumOff: 30000, satMod: 80000 });
    // Should produce some valid hex
    expect(result).toMatch(/^[0-9A-F]{6}$/);
    const [r, g, b] = hexToRgb(result);
    expect(r).toBeGreaterThanOrEqual(0);
    expect(g).toBeGreaterThanOrEqual(0);
    expect(b).toBeGreaterThanOrEqual(0);
  });

  // --- Tint modifier ---

  it('tint=50000 on pure red → moves halfway to white', () => {
    const result = applyColorModifiers('FF0000', { tint: 50000 });
    const [r, g, b] = hexToRgb(result);
    // R stays 255 (already at max), G and B move to ~128
    expect(r).toBe(255);
    expect(g).toBeCloseTo(128, -1);
    expect(b).toBeCloseTo(128, -1);
  });

  it('tint=50000 on black → mid-gray', () => {
    const result = applyColorModifiers('000000', { tint: 50000 });
    const [r, g, b] = hexToRgb(result);
    expect(r).toBeCloseTo(128, -1);
    expect(g).toBeCloseTo(128, -1);
    expect(b).toBeCloseTo(128, -1);
  });

  it('tint=100000 on black → white', () => {
    const result = applyColorModifiers('000000', { tint: 100000 });
    expect(result).toBe('FFFFFF');
  });

  it('tint=0 preserves color', () => {
    const result = applyColorModifiers('4472C4', { tint: 0 });
    expect(result).toBe('4472C4');
  });

  // --- Shade modifier ---

  it('shade=50000 on pure red → half intensity', () => {
    const result = applyColorModifiers('FF0000', { shade: 50000 });
    const [r, g, b] = hexToRgb(result);
    expect(r).toBeCloseTo(128, -1);
    expect(g).toBe(0);
    expect(b).toBe(0);
  });

  it('shade=50000 on white → mid-gray', () => {
    const result = applyColorModifiers('FFFFFF', { shade: 50000 });
    const [r, g, b] = hexToRgb(result);
    expect(r).toBeCloseTo(128, -1);
    expect(g).toBeCloseTo(128, -1);
    expect(b).toBeCloseTo(128, -1);
  });

  it('shade=100000 preserves color', () => {
    const result = applyColorModifiers('4472C4', { shade: 100000 });
    expect(result).toBe('4472C4');
  });

  it('shade=0 → black', () => {
    const result = applyColorModifiers('FF0000', { shade: 0 });
    expect(result).toBe('000000');
  });

  // --- Shade then tint order ---

  it('shade then tint: shade=50000 + tint=50000 on white', () => {
    // Shade first: white (255,255,255) * 0.5 = (128,128,128)
    // Tint second: (128,128,128) + (127,127,127)*0.5 = (192,192,192)
    const result = applyColorModifiers('FFFFFF', { shade: 50000, tint: 50000 });
    const [r, g, b] = hexToRgb(result);
    expect(r).toBeCloseTo(191, -1);
    expect(g).toBeCloseTo(191, -1);
    expect(b).toBeCloseTo(191, -1);
  });

  // --- Hue modifiers ---

  it('hueMod rotates hue', () => {
    // Pure red (hue=0), hueMod=50000 (50%) should keep hue at 0
    const result = applyColorModifiers('FF0000', { hueMod: 50000 });
    const [r, g, b] = hexToRgb(result);
    const [h] = rgbToHsl(r, g, b);
    expect(h).toBeCloseTo(0, 0);
  });

  it('hueOff shifts hue by degrees', () => {
    // Pure red (hue=0), hueOff=7200000 (120 degrees) → green
    const result = applyColorModifiers('FF0000', { hueOff: 7200000 });
    const [r, g, b] = hexToRgb(result);
    const [h] = rgbToHsl(r, g, b);
    expect(h).toBeCloseTo(120, 0);
  });

  // --- satOff modifier ---

  it('satOff adjusts saturation', () => {
    // Pure red (sat=1.0), satOff=-50000 should reduce to ~0.5
    const result = applyColorModifiers('FF0000', { satOff: -50000 });
    const [r, g, b] = hexToRgb(result);
    const [, s] = rgbToHsl(r, g, b);
    expect(s).toBeCloseTo(0.5, 1);
  });
});

// --- Color resolver ---

describe('createColorResolver', () => {
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
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink',
  };

  const fonts = { heading: 'Calibri Light', body: 'Calibri' };

  describe('resolveSchemeColor', () => {
    it('resolves tx1 → dk1 → 000000', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveSchemeColor('tx1')).toBe('000000');
    });

    it('resolves bg1 → lt1 → FFFFFF', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveSchemeColor('bg1')).toBe('FFFFFF');
    });

    it('resolves accent1 directly', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveSchemeColor('accent1')).toBe('4472C4');
    });

    it('resolves dk1 directly (bypasses clrMap when same)', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveSchemeColor('dk1')).toBe('000000');
    });

    it('returns 000000 for unknown scheme name', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveSchemeColor('unknown')).toBe('000000');
    });
  });

  describe('resolveFontRef', () => {
    it('resolves +mj-lt to heading font', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveFontRef('+mj-lt')).toBe('Calibri Light');
    });

    it('resolves +mn-lt to body font', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveFontRef('+mn-lt')).toBe('Calibri');
    });

    it('resolves +mj-ea to heading font', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveFontRef('+mj-ea')).toBe('Calibri Light');
    });

    it('resolves +mn-cs to body font', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveFontRef('+mn-cs')).toBe('Calibri');
    });

    it('passes through non-theme font names', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolveFontRef('Arial')).toBe('Arial');
    });
  });

  describe('resolve', () => {
    it('resolves a:srgbClr', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:srgbClr': { '@_val': 'FF5500' },
      });
      expect(result).toEqual({ color: 'FF5500' });
    });

    it('resolves a:schemeClr without modifiers', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:schemeClr': { '@_val': 'accent1' },
      });
      expect(result).toEqual({ color: '4472C4' });
    });

    it('resolves a:schemeClr tx1 through clrMap', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:schemeClr': { '@_val': 'tx1' },
      });
      expect(result).toEqual({ color: '000000' });
    });

    it('resolves a:schemeClr with lumMod/lumOff modifiers', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:schemeClr': {
          '@_val': 'accent1',
          'a:lumMod': { '@_val': '40000' },
          'a:lumOff': { '@_val': '60000' },
        },
      });
      // Should be a lighter version of accent1 (4472C4)
      expect(result.color).toMatch(/^[0-9A-F]{6}$/);
      // Verify it's lighter than the original
      const [, , origL] = rgbToHsl(...hexToRgb('4472C4'));
      const [, , newL] = rgbToHsl(...hexToRgb(result.color));
      expect(newL).toBeGreaterThan(origL);
    });

    it('resolves a:sysClr using lastClr', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:sysClr': { '@_val': 'windowText', '@_lastClr': '000000' },
      });
      expect(result).toEqual({ color: '000000' });
    });

    it('extracts alpha as transparency', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:schemeClr': {
          '@_val': 'accent1',
          'a:alpha': { '@_val': '50000' },
        },
      });
      expect(result.color).toBe('4472C4');
      expect(result.transparency).toBe(50);
    });

    it('alpha 100000 means fully opaque (no transparency)', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:schemeClr': {
          '@_val': 'accent1',
          'a:alpha': { '@_val': '100000' },
        },
      });
      expect(result.transparency).toBe(0);
    });

    it('alpha 0 means fully transparent', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:srgbClr': {
          '@_val': 'FF0000',
          'a:alpha': { '@_val': '0' },
        },
      });
      expect(result.transparency).toBe(100);
    });

    it('returns null for null input', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolve(null)).toBeNull();
    });

    it('returns null for empty object', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      expect(resolver.resolve({})).toBeNull();
    });

    it('resolves phClr as fallback black', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:schemeClr': { '@_val': 'phClr' },
      });
      expect(result).toEqual({ color: '000000' });
    });

    it('resolves a:prstClr with named color value', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:prstClr': { '@_val': 'red' },
      });
      expect(result).toEqual({ color: 'FF0000' });
    });

    it('resolves a:prstClr with missing val as black fallback', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:prstClr': {},
      });
      expect(result).toEqual({ color: '000000' });
    });

    it('resolves a:srgbClr with modifiers', () => {
      const resolver = createColorResolver(themeColors, clrMap, fonts);
      const result = resolver.resolve({
        'a:srgbClr': {
          '@_val': 'FF0000',
          'a:lumMod': { '@_val': '50000' },
        },
      });
      // Red with halved luminance should be darker
      const [, , newL] = rgbToHsl(...hexToRgb(result.color));
      expect(newL).toBeLessThan(0.5);
    });
  });
});
