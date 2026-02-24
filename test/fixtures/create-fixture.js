import PptxGenJS from 'pptxgenjs';
import { writeFileSync } from 'fs';

const pptx = new PptxGenJS();

// Define a slide master
pptx.defineSlideMaster({
  title: 'TITLE_SLIDE',
  background: { color: '003366' },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.5, w: 9, h: 1.5, fontFace: 'Arial', fontSize: 36, color: 'FFFFFF', align: 'center' } } },
    { placeholder: { options: { name: 'subtitle', type: 'body', x: 0.5, y: 2.5, w: 9, h: 2, fontFace: 'Arial', fontSize: 18, color: 'CCCCCC', align: 'center' } } },
  ],
  slideNumber: { x: 9.0, y: 6.9, w: 0.8, h: 0.3, fontFace: 'Arial', fontSize: 10, color: 'FFFFFF', align: 'right' },
});

pptx.defineSlideMaster({
  title: 'CONTENT_SLIDE',
  background: { color: 'FFFFFF' },
  objects: [
    { rect: { x: 0, y: 0, w: '100%', h: 0.75, fill: { color: '003366' } } },
    { text: { text: 'Company Name', options: { x: 0.5, y: 0.15, w: 5, h: 0.45, fontFace: 'Arial', fontSize: 14, color: 'FFFFFF', bold: true } } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 1.0, w: 9, h: 0.8, fontFace: 'Arial', fontSize: 28, color: '003366' } } },
    { placeholder: { options: { name: 'body', type: 'body', x: 0.5, y: 2.0, w: 9, h: 4.5, fontFace: 'Arial', fontSize: 14, color: '333333' } } },
  ],
  slideNumber: { x: 9.0, y: 6.9, w: 0.8, h: 0.3, fontFace: 'Arial', fontSize: 10, color: '666666', align: 'right' },
});

// Add slides using the masters
const slide1 = pptx.addSlide({ masterName: 'TITLE_SLIDE' });
slide1.addText('Test Presentation', { placeholder: 'title' });
slide1.addText('Generated for testing pptx-masters', { placeholder: 'subtitle' });

const slide2 = pptx.addSlide({ masterName: 'CONTENT_SLIDE' });
slide2.addText('Content Slide', { placeholder: 'title' });
slide2.addText('This is sample body text for testing extraction.', { placeholder: 'body' });

// Write to file
const data = await pptx.write({ outputType: 'nodebuffer' });
writeFileSync(new URL('./minimal.pptx', import.meta.url), data);
console.log('Created test/fixtures/minimal.pptx');
