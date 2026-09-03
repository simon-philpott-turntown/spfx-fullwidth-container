const fs = require('fs');
const path = require('path');

const tsPath = path.join(__dirname, '../src/webparts/fullWidthContainer/components/TtSvgIconCollection.ts');
const code = fs.readFileSync(tsPath, 'utf8');

const match = code.match(/export const TT_SVG_ICONS: ITtSvgIconEntry\[\] = (\[[\s\S]*?\]);\s*$/);
if (!match) {
  console.error('Could not find TT_SVG_ICONS in file');
  process.exit(1);
}

const arr = eval(match[1]);
console.log('Successfully parsed TT icons count:', arr.length);

const minimal = arr.map(i => ({
  key: i.id,
  label: i.name,
  cat: 'custom',
  isCustomSvg: true,
  svg: i.svgMarkup
}));

const outPath = path.join(__dirname, '../preview/tt_icons.js');
const jsContent = `window.TT_EXTRA_ICONS = ${JSON.stringify(minimal)};\n`;
fs.writeFileSync(outPath, jsContent, 'utf8');
console.log('Saved preview/tt_icons.js with size:', fs.statSync(outPath).size);
