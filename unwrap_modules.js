/**
 * unwrap_modules.js
 * Elimina las etiquetas <script> y </script> envolventes de cada Mod_*.html
 * dejando solo el contenido JS puro.
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

  // Verificar si empieza con <script> (ignorando espacios/BOM)
  const openMatch  = content.match(/^(\s*<script[^>]*>\n?)/i);
  const closeMatch = content.match(/\n?<\/script>\s*\n?$/i);

  if (!openMatch || !closeMatch) {
    console.log(`  (sin <script> envolvente) ${file}`);
    continue;
  }

  // Quitar etiqueta de apertura y cierre
  let inner = content.slice(openMatch[0].length);
  inner = inner.replace(/<\/script>\s*\n?$/i, '');

  fs.writeFileSync(filePath, inner, 'utf8');
  modified++;
  console.log(`  ✓ ${file}`);
}

console.log(`\nTotal modificados: ${modified} de ${files.length} archivos.`);
