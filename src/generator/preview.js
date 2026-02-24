/**
 * Preview generator — creates a sample PPTX file using PptxGenJS
 * that demonstrates all extracted slide masters with sample content.
 */

import { resolve } from 'path';
import { readFile } from 'fs/promises';
import PptxGenJS from 'pptxgenjs';
import { toUpperSnakeCase } from './code.js';

/**
 * Generate a preview PPTX file demonstrating all slide masters.
 *
 * @param {Array<object>} masterData - Array of layout objects with mapped PptxGenJS data
 * @param {Record<string, string>} themeColors - Theme color map
 * @param {{ heading: string, body: string }} themeFonts - Theme fonts
 * @param {{ width: number, height: number }} dimensions - Slide dimensions
 * @param {string} [outputDir] - Output directory for resolving image paths
 * @returns {Promise<Buffer>} PPTX file as a Node buffer
 */
export async function generatePreview(masterData, themeColors, themeFonts, dimensions, outputDir) {
  const pptx = new PptxGenJS();

  // Define custom layout dimensions (must defineLayout before setting pptx.layout)
  pptx.defineLayout({
    name: 'LAYOUT_CUSTOM',
    width: dimensions?.width || 10,
    height: dimensions?.height || 7.5,
  });
  pptx.layout = 'LAYOUT_CUSTOM';

  // Pre-load images as base64 data for embedding in preview
  const imageCache = new Map();
  if (outputDir) {
    for (const master of masterData) {
      for (const obj of master.objects || []) {
        if (obj.image?.path) {
          const absPath = resolve(outputDir, obj.image.path);
          if (!imageCache.has(absPath)) {
            try {
              const buf = await readFile(absPath);
              const ext = absPath.split('.').pop().toLowerCase();
              const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
              imageCache.set(absPath, `data:${mime};base64,${buf.toString('base64')}`);
            } catch {
              // Image not found — skip it
            }
          }
        }
      }
      if (master.background?.path) {
        const absPath = resolve(outputDir, master.background.path);
        if (!imageCache.has(absPath)) {
          try {
            const buf = await readFile(absPath);
            const ext = absPath.split('.').pop().toLowerCase();
            const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
            imageCache.set(absPath, `data:${mime};base64,${buf.toString('base64')}`);
          } catch {
            // Image not found — skip it
          }
        }
      }
    }
  }

  for (const master of masterData) {
    const title = toUpperSnakeCase(master.name);

    // Build the master definition, resolving images to base64 data when available
    const safeObjects = [];
    for (const obj of master.objects || []) {
      if (obj.image?.path && outputDir) {
        const absPath = resolve(outputDir, obj.image.path);
        const dataUri = imageCache.get(absPath);
        if (dataUri) {
          // Replace file path with embedded base64 data
          safeObjects.push({
            image: { ...obj.image, data: dataUri, path: undefined },
          });
          continue;
        }
        // Image not available — skip it
        continue;
      }
      safeObjects.push(obj);
    }

    const masterDef = {
      title,
      objects: safeObjects,
    };

    if (master.background) {
      if (master.background.path && outputDir) {
        const absPath = resolve(outputDir, master.background.path);
        const dataUri = imageCache.get(absPath);
        if (dataUri) {
          masterDef.background = { data: dataUri };
        } else {
          masterDef.background = { color: 'EEEEEE' };
        }
      } else if (master.background.path) {
        masterDef.background = { color: 'EEEEEE' };
      } else {
        masterDef.background = master.background;
      }
    }

    if (master.slideNumber) {
      masterDef.slideNumber = master.slideNumber;
    }

    // Register this master
    pptx.defineSlideMaster(masterDef);

    // Add a sample slide using this master
    const slide = pptx.addSlide({ masterName: title });

    // Fill placeholders with sample content
    const objects = master.objects || [];
    const titlePh = objects.find((o) => o.placeholder?.options?.type === 'title');
    if (titlePh) {
      slide.addText(master.name, { placeholder: titlePh.placeholder.options.name });
    }

    const bodyPh = objects.find((o) => o.placeholder?.options?.type === 'body');
    if (bodyPh) {
      slide.addText(`Sample content for "${master.name}" layout`, {
        placeholder: bodyPh.placeholder.options.name,
      });
    }
  }

  return pptx.write({ outputType: 'nodebuffer' });
}
