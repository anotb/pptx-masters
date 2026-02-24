import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { resolve, join } from 'path';
import { mkdir, rm, readFile, access } from 'fs/promises';
import { execFile } from 'child_process';
import { promisify } from 'util';
import { extract } from '../src/index.js';
import { generateMastersCode, generateThemeJson, generateAgentInstructions } from '../src/generator/code.js';
import { generateReport } from '../src/generator/report.js';
import { generatePreview } from '../src/generator/preview.js';

const execFileAsync = promisify(execFile);

const FIXTURE = resolve('test/fixtures/minimal.pptx');
const CLI = resolve('src/cli.js');
const OUTPUT_DIR = resolve('test/output-integration');
const OUTPUT_FILTERED = resolve('test/output-filtered');
const OUTPUT_MINIMAL = resolve('test/output-minimal');

async function fileExists(path) {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

// --- Full pipeline tests ---

describe('extract() full pipeline', () => {
  let result;

  beforeAll(async () => {
    result = await extract(FIXTURE);
  });

  it('returns masterData with known master names', () => {
    expect(result.masterData).toBeDefined();
    expect(result.masterData).toHaveLength(3);
    const names = result.masterData.map((m) => m.name);
    expect(names).toContain('DEFAULT');
    expect(names).toContain('TITLE_SLIDE');
    expect(names).toContain('CONTENT_SLIDE');
  });

  it('returns themeColors with 12 keys', () => {
    expect(result.themeColors).toBeDefined();
    expect(Object.keys(result.themeColors)).toHaveLength(12);
  });

  it('returns themeFonts with heading and body', () => {
    expect(result.themeFonts).toBeDefined();
    expect(result.themeFonts.heading).toBeTruthy();
    expect(result.themeFonts.body).toBeTruthy();
  });

  it('returns reasonable dimensions', () => {
    expect(result.dimensions).toBeDefined();
    expect(result.dimensions.width).toBe(10);
    expect(result.dimensions.height).toBeCloseTo(5.625, 2);
  });

  it('returns layout names', () => {
    const names = result.layouts.map((l) => l.name);
    expect(names.length).toBeGreaterThan(0);
    expect(names).toContain('TITLE_SLIDE');
  });

  it('masterData entries have expected structure with known values', () => {
    for (const entry of result.masterData) {
      expect(entry.name).toMatch(/^(DEFAULT|TITLE_SLIDE|CONTENT_SLIDE)$/);
      expect(entry.title).toMatch(/^(DEFAULT|TITLE_SLIDE|CONTENT_SLIDE)$/);
      expect(Array.isArray(entry.objects)).toBe(true);
      // background and slideNumber can be null
    }
  });

  it('returns empty warnings array for well-formed minimal fixture', () => {
    expect(Array.isArray(result.warnings)).toBe(true);
    expect(result.warnings).toHaveLength(0);
  });

  it('returns templateName', () => {
    expect(result.templateName).toBe('minimal.pptx');
  });
});

// --- Code generation integration ---

describe('code generation integration', () => {
  let result;

  beforeAll(async () => {
    result = await extract(FIXTURE);
  });

  it('generates valid JavaScript code', () => {
    const code = generateMastersCode(result.masterData, {
      templateName: result.templateName,
      dimensions: result.dimensions,
      themeColors: result.themeColors,
      themeFonts: result.themeFonts,
    });

    expect(code).toBeTruthy();
    expect(code).toContain('export function registerMasters');
    expect(code).toContain('defineSlideMaster');
    expect(code).toContain('THEME_COLORS');
  });

  it('generated code has valid JavaScript syntax', () => {
    const code = generateMastersCode(result.masterData, {
      templateName: result.templateName,
      dimensions: result.dimensions,
      themeColors: result.themeColors,
      themeFonts: result.themeFonts,
    });

    // Strip export keywords for eval
    const evalCode = code.replace(/^export /gm, '');
    expect(() => new Function(evalCode)).not.toThrow();
  });

  it('generated code contains all layout titles', () => {
    const code = generateMastersCode(result.masterData, {
      templateName: result.templateName,
      dimensions: result.dimensions,
      themeColors: result.themeColors,
      themeFonts: result.themeFonts,
    });

    for (const master of result.masterData) {
      expect(code).toContain(master.title);
    }
  });
});

// --- Theme.json generation ---

describe('theme.json generation', () => {
  let result;

  beforeAll(async () => {
    result = await extract(FIXTURE);
  });

  it('generates valid theme JSON', () => {
    const themeJson = generateThemeJson(
      result.themeColors,
      result.themeFonts,
      result.dimensions,
    );

    expect(themeJson.colors).toBeDefined();
    expect(Object.keys(themeJson.colors).length).toBe(12);
    expect(themeJson.fonts).toBeDefined();
    expect(themeJson.fonts.heading).toBeTruthy();
    expect(themeJson.fonts.body).toBeTruthy();
    expect(themeJson.dimensions).toBeDefined();
    expect(themeJson.dimensions.width).toBe(10);
  });
});

// --- Report generation ---

describe('report generation', () => {
  let result;

  beforeAll(async () => {
    result = await extract(FIXTURE);
  });

  it('generates markdown report with template info', () => {
    const report = generateReport({
      templateName: result.templateName,
      dimensions: result.dimensions,
      themeColors: result.themeColors,
      themeFonts: result.themeFonts,
      layouts: result.layouts.map((l) => ({
        name: l.name,
        background: result.masterData.find((m) => m.name === l.name)?.background,
        placeholders: l.placeholders,
        staticShapes: l.staticShapes,
        warnings: l.warnings,
      })),
      allWarnings: result.warnings,
    });

    expect(report).toContain('Extraction Report');
    expect(report).toContain(result.templateName);
  });

  it('report contains theme section', () => {
    const report = generateReport({
      templateName: result.templateName,
      dimensions: result.dimensions,
      themeColors: result.themeColors,
      themeFonts: result.themeFonts,
      layouts: [],
      allWarnings: [],
    });

    expect(report).toContain('## Theme');
    expect(report).toContain('### Colors');
    expect(report).toContain('### Fonts');
  });

  it('report contains layout sections', () => {
    const report = generateReport({
      templateName: result.templateName,
      dimensions: result.dimensions,
      themeColors: result.themeColors,
      themeFonts: result.themeFonts,
      layouts: result.layouts.map((l) => ({
        name: l.name,
        background: null,
        placeholders: l.placeholders,
        staticShapes: l.staticShapes,
        warnings: l.warnings,
      })),
      allWarnings: result.warnings,
    });

    expect(report).toContain('## Layouts');
    for (const layout of result.layouts) {
      expect(report).toContain(layout.name);
    }
  });
});

// --- Preview generation ---

describe('preview generation', () => {
  let result;

  beforeAll(async () => {
    result = await extract(FIXTURE);
  });

  it('generates a valid ZIP buffer (starts with PK)', async () => {
    const buffer = await generatePreview(
      result.masterData,
      result.themeColors,
      result.themeFonts,
      result.dimensions,
    );

    expect(Buffer.isBuffer(buffer)).toBe(true);
    expect(buffer.length).toBeGreaterThan(0);
    // PPTX is a ZIP file — starts with PK signature (0x50 0x4B)
    expect(buffer[0]).toBe(0x50);
    expect(buffer[1]).toBe(0x4B);
  });
});

// --- Agent instructions generation ---

describe('agent instructions generation', () => {
  let result;

  beforeAll(async () => {
    result = await extract(FIXTURE);
  });

  it('generates SLIDE_MASTERS.md with all masters', () => {
    const md = generateAgentInstructions(
      result.masterData,
      result.themeColors,
      result.themeFonts,
      result.dimensions,
      result.templateName,
    );

    expect(md).toContain('# Slide Masters');
    expect(md).toContain('## Available Masters');
    expect(md).toContain('## Quick Start');
    for (const master of result.masterData) {
      expect(md).toContain(master.title);
    }
  });
});

// --- Layout filtering ---

describe('layout filtering', () => {
  it('filters layouts by name (case-insensitive substring)', async () => {
    const result = await extract(FIXTURE, { layouts: ['title'] });

    expect(result.masterData.length).toBe(1);
    expect(result.masterData[0].name).toBe('TITLE_SLIDE');
    expect(result.layouts.length).toBe(1);
  });

  it('returns all layouts when no filter', async () => {
    const result = await extract(FIXTURE);
    expect(result.layouts.length).toBe(3);
    expect(result.masterData.length).toBe(3);
  });

  it('returns empty when filter matches nothing', async () => {
    const result = await extract(FIXTURE, { layouts: ['nonexistent'] });
    expect(result.masterData.length).toBe(0);
    expect(result.layouts.length).toBe(0);
  });
});

// --- Error handling ---

describe('extract() error handling', () => {
  it('throws a useful error for a non-existent file path', async () => {
    await expect(extract('/tmp/does-not-exist.pptx')).rejects.toThrow();
  });
});

// --- Media files structure ---

describe('extract() mediaFiles', () => {
  it('returns mediaFiles as an array', async () => {
    const result = await extract(FIXTURE);
    expect(Array.isArray(result.mediaFiles)).toBe(true);
    // Minimal fixture may have no images, but the array must exist
  });
});

// --- CLI tests (spawn process) ---

describe('CLI', () => {
  let cliStdout;

  beforeAll(async () => {
    // Run full extraction once — all dependent tests use OUTPUT_DIR
    const { stdout } = await execFileAsync('node', [
      CLI,
      FIXTURE,
      '-o',
      OUTPUT_DIR,
    ]);
    cliStdout = stdout;
  });

  afterAll(async () => {
    // Clean up all test output directories
    await rm(OUTPUT_DIR, { recursive: true, force: true });
    await rm(OUTPUT_FILTERED, { recursive: true, force: true });
    await rm(OUTPUT_MINIMAL, { recursive: true, force: true });
  });

  it('runs full extraction and exits 0', () => {
    expect(cliStdout).toContain('Extracted');
    expect(cliStdout).toContain('layout(s)');
  });

  it('creates all expected output files', async () => {
    expect(await fileExists(join(OUTPUT_DIR, 'masters.js'))).toBe(true);
    expect(await fileExists(join(OUTPUT_DIR, 'theme.json'))).toBe(true);
    expect(await fileExists(join(OUTPUT_DIR, 'SLIDE_MASTERS.md'))).toBe(true);
    expect(await fileExists(join(OUTPUT_DIR, 'report.md'))).toBe(true);
    expect(await fileExists(join(OUTPUT_DIR, 'preview.pptx'))).toBe(true);
  });

  it('masters.js contains valid code', async () => {
    const code = await readFile(join(OUTPUT_DIR, 'masters.js'), 'utf-8');
    expect(code).toContain('registerMasters');
    expect(code).toContain('defineSlideMaster');
  });

  it('theme.json is valid JSON', async () => {
    const raw = await readFile(join(OUTPUT_DIR, 'theme.json'), 'utf-8');
    const parsed = JSON.parse(raw);
    expect(parsed.colors).toBeDefined();
    expect(parsed.fonts).toBeDefined();
    expect(parsed.dimensions).toBeDefined();
  });

  it('--list prints layout names', async () => {
    const { stdout } = await execFileAsync('node', [
      CLI,
      FIXTURE,
      '--list',
    ]);

    expect(stdout).toContain('DEFAULT');
    expect(stdout).toContain('TITLE_SLIDE');
    expect(stdout).toContain('CONTENT_SLIDE');
    // Should be numbered
    expect(stdout).toMatch(/1\./);
    expect(stdout).toMatch(/2\./);
    expect(stdout).toMatch(/3\./);
  });

  it('--layouts filters to matching layout', async () => {
    const { stdout } = await execFileAsync('node', [
      CLI,
      FIXTURE,
      '--layouts',
      'TITLE',
      '-o',
      OUTPUT_FILTERED,
    ]);

    expect(stdout).toContain('1 layout(s)');

    const code = await readFile(join(OUTPUT_FILTERED, 'masters.js'), 'utf-8');
    expect(code).toContain('TITLE_SLIDE');
    expect(code).not.toContain("'DEFAULT'");
    expect(code).not.toContain("'CONTENT_SLIDE'");
  });

  it('--no-preview --no-report skips those files', async () => {
    await execFileAsync('node', [
      CLI,
      FIXTURE,
      '--no-preview',
      '--no-report',
      '-o',
      OUTPUT_MINIMAL,
    ]);

    expect(await fileExists(join(OUTPUT_MINIMAL, 'masters.js'))).toBe(true);
    expect(await fileExists(join(OUTPUT_MINIMAL, 'theme.json'))).toBe(true);
    expect(await fileExists(join(OUTPUT_MINIMAL, 'SLIDE_MASTERS.md'))).toBe(true);
    expect(await fileExists(join(OUTPUT_MINIMAL, 'report.md'))).toBe(false);
    expect(await fileExists(join(OUTPUT_MINIMAL, 'preview.pptx'))).toBe(false);
  });

  it('nonexistent file exits with error', async () => {
    try {
      await execFileAsync('node', [CLI, 'nonexistent.pptx']);
      // Should not reach here
      expect.unreachable('Expected error');
    } catch (err) {
      expect(err.stderr).toContain('Error');
      expect(err.stderr).toContain('not found');
    }
  });

  it('non-PPTX file exits with error', async () => {
    try {
      await execFileAsync('node', [CLI, 'package.json']);
      expect.unreachable('Expected error');
    } catch (err) {
      expect(err.stderr).toContain('Error');
      expect(err.stderr).toContain('Invalid file type');
    }
  });

  it('--version prints version', async () => {
    const { stdout } = await execFileAsync('node', [CLI, '--version']);
    expect(stdout.trim()).toMatch(/^\d+\.\d+\.\d+$/);
  });
});
