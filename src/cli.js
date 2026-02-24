#!/usr/bin/env node

/**
 * pptx-masters CLI — extract PptxGenJS slide master code from PPTX/POTX templates.
 */

import { resolve, basename } from 'path';
import { mkdir, writeFile, access, stat } from 'fs/promises';
import { createRequire } from 'module';
import { Command } from 'commander';
import { extract } from './index.js';
import { generateMastersCode, generateThemeJson, generateAgentInstructions } from './generator/code.js';
import { generateReport } from './generator/report.js';
import { generatePreview } from './generator/preview.js';
import { extractPptx } from './parser/zip.js';

const require = createRequire(import.meta.url);
const { version } = require('../package.json');

const program = new Command();

function log(msg) {
  process.stderr.write(msg + '\n');
}

program
  .name('pptx-masters')
  .description('Extract PptxGenJS defineSlideMaster() code from .pptx/.potx templates')
  .version(version)
  .argument('<input>', 'Path to .pptx or .potx file')
  .option('-o, --output <dir>', 'Output directory', './output')
  .option('--list', 'List all layout names and exit')
  .option('--layouts <names>', 'Comma-separated layout names to include (deprecated, use --layout)')
  .option('--layout <name>', 'Layout name or number to include (repeatable, use --list to see numbers)', (val, prev) => prev.concat(val), [])
  .option('--no-preview', 'Skip preview.pptx generation')
  .option('--no-report', 'Skip report.md generation')
  .option('-v, --verbose', 'Verbose logging')
  .action(async (input, opts) => {
    try {
      await run(input, opts);
    } catch (err) {
      process.stderr.write(`Error: ${err.message}\n`);
      process.exit(1);
    }
  });

program.parse();

/**
 * Main pipeline.
 *
 * @param {string} input - Input file path
 * @param {object} opts - CLI options
 */
async function run(input, opts) {
  const verbose = opts.verbose || false;

  // 1. Validate input file
  const inputPath = resolve(input);
  const ext = inputPath.split('.').pop().toLowerCase();

  if (ext !== 'pptx' && ext !== 'potx') {
    throw new Error(`Invalid file type: .${ext}. Expected .pptx or .potx`);
  }

  try {
    await access(inputPath);
  } catch {
    throw new Error(`File not found: ${inputPath}`);
  }

  const inputStat = await stat(inputPath);
  if (!inputStat.isFile()) {
    throw new Error(`Not a file: ${inputPath}`);
  }

  if (verbose) log(`Input: ${inputPath}`);

  // 2. Parse layouts filter (--layout repeatable takes priority, --layouts comma-sep is deprecated)
  let layoutFilters;
  if (opts.layout && opts.layout.length > 0) {
    layoutFilters = opts.layout.map((s) => s.trim()).filter(Boolean);
  } else if (opts.layouts) {
    layoutFilters = opts.layouts.split(',').map((s) => s.trim()).filter(Boolean);
  }

  // 3. Run extraction
  if (verbose) log('Extracting PPTX archive...');

  const result = await extract(inputPath, { layouts: layoutFilters });

  const {
    masterData,
    themeColors,
    themeFonts,
    dimensions,
    warnings,
    layouts,
    mediaFiles,
    templateName,
  } = result;

  if (verbose) {
    log(`Theme: ${Object.keys(themeColors).length} colors, fonts: ${themeFonts.heading}/${themeFonts.body}`);
    log(`Dimensions: ${dimensions.width}" x ${dimensions.height}"`);
    log(`Layouts found: ${layouts.length}`);
  }

  // 4. --list mode
  if (opts.list) {
    for (let i = 0; i < layouts.length; i++) {
      process.stdout.write(`${i + 1}. ${layouts[i].name}\n`);
    }
    return;
  }

  if (verbose) {
    for (const layout of layouts) {
      log(`  Mapped layout: ${layout.name} (${layout.placeholders.length} placeholders, ${layout.staticShapes.length} shapes)`);
    }
  }

  // 5. Create output directory
  const outputDir = resolve(opts.output);
  await mkdir(outputDir, { recursive: true });

  if (verbose) log(`Output directory: ${outputDir}`);

  // 6. Generate masters.js
  if (verbose) log('Generating masters.js...');
  const mastersCode = generateMastersCode(masterData, {
    templateName,
    dimensions,
    themeColors,
    themeFonts,
  });
  await writeFile(resolve(outputDir, 'masters.js'), mastersCode, 'utf-8');

  // 6b. Generate package.json (ESM required for masters.js)
  const pkgJsonPath = resolve(outputDir, 'package.json');
  try {
    await access(pkgJsonPath);
    // package.json exists — don't overwrite
  } catch {
    const pkgJson = {
      name: templateName.replace(/[^a-z0-9-]/gi, '-').toLowerCase(),
      version: '1.0.0',
      type: 'module',
      dependencies: { pptxgenjs: '^4.0.1' },
    };
    await writeFile(pkgJsonPath, JSON.stringify(pkgJson, null, 2) + '\n', 'utf-8');
  }

  // 7. Generate theme.json
  if (verbose) log('Generating theme.json...');
  const themeJson = generateThemeJson(themeColors, themeFonts, dimensions, masterData);
  await writeFile(
    resolve(outputDir, 'theme.json'),
    JSON.stringify(themeJson, null, 2) + '\n',
    'utf-8',
  );

  // 8. Generate SLIDE_MASTERS.md
  if (verbose) log('Generating SLIDE_MASTERS.md...');
  const agentInstructions = generateAgentInstructions(
    masterData,
    themeColors,
    themeFonts,
    dimensions,
    templateName,
  );
  await writeFile(resolve(outputDir, 'SLIDE_MASTERS.md'), agentInstructions, 'utf-8');

  // 8b. Generate STYLE_GUIDE.md (user-customizable, don't overwrite if exists)
  const styleGuidePath = resolve(outputDir, 'STYLE_GUIDE.md');
  try {
    await access(styleGuidePath);
    // Exists — don't overwrite user edits
  } catch {
    const styleGuide = [
      '# Style Guide',
      '',
      '> Customize this file with your design preferences. SLIDE_MASTERS.md references it.',
      '> This file is never overwritten by pptx-masters — your edits are safe.',
      '',
      '## Dos',
      '',
      '- <!-- Add your style preferences here, e.g.: "Use data-dense layouts for financial content" -->',
      '',
      '## Don\'ts',
      '',
      '- <!-- Add things to avoid, e.g.: "Don\'t use clip art or stock photos" -->',
      '',
      '## Typography preferences',
      '',
      '- <!-- e.g.: "Body text should be 11-14pt", "Use bold sparingly" -->',
      '',
      '## Color usage',
      '',
      '- <!-- e.g.: "Reserve accent2 (blue) for hyperlinks", "Use darker shades for headers" -->',
      '',
      '## Layout preferences',
      '',
      '- <!-- e.g.: "Always use section dividers between topics", "Prefer 2-column over single column" -->',
      '',
    ].join('\n');
    await writeFile(styleGuidePath, styleGuide, 'utf-8');
  }

  // 9. Generate report.md (unless --no-report)
  if (opts.report !== false) {
    if (verbose) log('Generating report.md...');
    const reportMd = generateReport({
      templateName,
      dimensions,
      themeColors,
      themeFonts,
      layouts: layouts.map((l) => ({
        name: l.name,
        background: masterData.find((m) => m.name === l.name)?.background,
        placeholders: l.placeholders,
        staticShapes: l.staticShapes,
        slideNumber: masterData.find((m) => m.name === l.name)?.slideNumber,
        footerObjects: masterData.find((m) => m.name === l.name)?.objects?.filter((o) => o.text) || [],
        warnings: l.warnings,
      })),
      allWarnings: warnings,
    });
    await writeFile(resolve(outputDir, 'report.md'), reportMd, 'utf-8');
  }

  // 10. Copy media files
  if (mediaFiles.length > 0) {
    if (verbose) log(`Copying ${mediaFiles.length} media file(s)...`);
    const mediaDir = resolve(outputDir, 'media');
    await mkdir(mediaDir, { recursive: true });

    const archive = await extractPptx(inputPath);
    for (const { archivePath, filename } of mediaFiles) {
      try {
        const buffer = await archive.getBuffer(archivePath);
        await writeFile(resolve(mediaDir, filename), buffer);
        if (verbose) log(`  Copied: ${filename}`);
      } catch (err) {
        if (verbose) log(`  Warning: Could not copy ${archivePath}: ${err.message}`);
        warnings.push(`Could not copy media file: ${archivePath}`);
      }
    }
  }

  // 11. Generate preview.pptx (unless --no-preview)
  if (opts.preview !== false) {
    if (verbose) log('Generating preview.pptx...');
    const previewBuffer = await generatePreview(
      masterData,
      themeColors,
      themeFonts,
      dimensions,
      outputDir,
    );
    await writeFile(resolve(outputDir, 'preview.pptx'), previewBuffer);
  }

  // 12. Print summary
  const outputFiles = ['masters.js', 'theme.json', 'SLIDE_MASTERS.md'];
  if (opts.report !== false) outputFiles.push('report.md');
  if (opts.preview !== false) outputFiles.push('preview.pptx');
  if (mediaFiles.length > 0) outputFiles.push(`media/ (${mediaFiles.length} files)`);

  process.stdout.write(`\nExtracted ${masterData.length} layout(s) from ${templateName}\n`);
  process.stdout.write(`Output: ${outputDir}\n`);
  process.stdout.write(`Files: ${outputFiles.join(', ')}\n`);

  if (warnings.length > 0) {
    process.stdout.write(`Warnings: ${warnings.length}\n`);
  }
}
