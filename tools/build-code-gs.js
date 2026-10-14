/* Copies export-excel.js into apps-script/Code.gs (between the LAYOUT markers),
 * so the Google Sheet tab and the Excel export always use the same seat plan layout.
 * Run after changing export-excel.js:  node tools/build-code-gs.js
 */
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const gsPath = path.join(root, 'apps-script', 'Code.gs');
const layout = fs.readFileSync(path.join(root, 'export-excel.js'), 'utf8').trim();
const gs = fs.readFileSync(gsPath, 'utf8');
const start = '// ===== LAYOUT: copied from export-excel.js by tools/build-code-gs.js (do not edit here) =====';
const end = '// ===== END LAYOUT =====';
const a = gs.indexOf(start), b = gs.indexOf(end);
if (a < 0 || b < a) throw new Error('LAYOUT markers not found in Code.gs');
fs.writeFileSync(gsPath, gs.slice(0, a + start.length) + '\n' + layout + '\n' + gs.slice(b));
console.log('Code.gs updated from export-excel.js');
