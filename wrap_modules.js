/**
 * wrap_modules.js
 * Envuelve el contenido de cada Mod_*.html con <script>...</script>
 * si aún no lo tiene, para que GAS pueda servirlo como HTML válido.
 */
const fs   = require('fs');
const path = require('path');

const BASE  = __dirname;
const files = fs.readdirSync(BASE)
  .filter(f => /^Mod_.*\.html$/.test(f))
  .sort();

let modified = 0;

console.log(`Procesando ${files.length} archivos Mod_*.html...\n`);

for (const file of files) {
  const filePath = path.join(BASE, file);
  const content  = fs.readFileSync(filePath, 'utf8');

  // Verificar si ya tiene <script> al inicio (ignorando espacios/BOM)
  if (/^\s*<script[\s>]/i.test(content)) {
    console.log(`  (ya tiene <script>) ${file}`);
    continue;
  }

  const wrapped = '<script>\n' + content + '\n</script>\n';
  fs.writeFileSync(filePath, wrapped, 'utf8');
  modified++;
  console.log(`  ✓ ${file}`);
}

console.log(`\nTotal modificados: ${modified} de ${files.length} archivos.`);
