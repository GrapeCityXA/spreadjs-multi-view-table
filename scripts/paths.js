import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const projectRoot = fileURLToPath(new URL('../', import.meta.url));
export const publicRoot = path.join(projectRoot, 'public');
export const distRoot = path.join(projectRoot, 'dist');

const spreadRoot = path.join(projectRoot, 'node_modules/@grapecity-software/spread-sheets');
// Development and build share these exact asset paths.
export const vendorAssets = new Map([
  ['/vendor/spreadjs/gc.spread.sheets.all.min.js', path.join(spreadRoot, 'dist/gc.spread.sheets.all.min.js')],
  ['/vendor/spreadjs/gc.spread.sheets.excel2013white.css', path.join(spreadRoot, 'styles/gc.spread.sheets.excel2013white.css')]
]);

// Local-only files served straight from the project root. They are deliberately
// not copied into dist/, so npm run build never bundles large media.
export const devAssets = new Map([
  ['/intro.mp4', path.join(projectRoot, 'intro.mp4')]
]);
