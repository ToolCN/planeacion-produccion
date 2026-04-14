/**
 * fix_backticks.js
 * Convierte template literals (backticks) en strings con concatenación
 * Solo procesa el contenido dentro de bloques <script>...</script>
 */
const fs   = require('fs');
const path = require('path');

const BASE = __dirname;

// ─── Convierte template literals en un fragmento de código JS ────────────────
function convertTemplateLiterals(source) {
  let out   = '';
  let i     = 0;
  let count = 0;
  const n   = source.length;

  function ch(offset = 0) { return i + offset < n ? source[i + offset] : ''; }

  while (i < n) {
    // ── Comilla simple: saltar sin modificar ──────────────────────────────────
    if (ch() === "'") {
      out += ch(); i++;
      while (i < n && ch() !== "'") {
        if (ch() === '\\') { out += ch() + ch(1); i += 2; }
        else                { out += ch(); i++; }
      }
      if (i < n) { out += ch(); i++; }

    // ── Comilla doble: saltar sin modificar ───────────────────────────────────
    } else if (ch() === '"') {
      out += ch(); i++;
      while (i < n && ch() !== '"') {
        if (ch() === '\\') { out += ch() + ch(1); i += 2; }
        else                { out += ch(); i++; }
      }
      if (i < n) { out += ch(); i++; }

    // ── Comentario de línea //: saltar ───────────────────────────────────────
    } else if (ch() === '/' && ch(1) === '/') {
      out += ch(); i++;
      while (i < n && ch() !== '\n') { out += ch(); i++; }

    // ── Comentario de bloque /* */: saltar ───────────────────────────────────
    } else if (ch() === '/' && ch(1) === '*') {
      out += ch() + ch(1); i += 2;
      while (i < n && !(ch() === '*' && ch(1) === '/')) { out += ch(); i++; }
      if (i < n) { out += ch() + ch(1); i += 2; }

    // ── Template literal ─────────────────────────────────────────────────────
    } else if (ch() === '`') {
      count++;
      i++;          // saltar backtick de apertura
      let tmpl = '';

      while (i < n) {
        if (ch() === '\\') {
          // Carácter escapado dentro del template
          const nx = ch(1);
          if      (nx === '`')  { tmpl += '`';   i += 2; }   // \` → `
          else if (nx === "'")  { tmpl += "\\'";  i += 2; }   // \' → \'
          else if (nx === '\n') {                 i += 2; }   // line continuation → borrar
          else                  { tmpl += ch() + ch(1); i += 2; }

        } else if (ch() === '`') {
          // Fin del template literal
          i++;
          break;

        } else if (ch() === '$' && ch(1) === '{') {
          // Interpolación ${expr}
          tmpl += "' + (";
          i += 2;   // saltar ${
          let depth = 1;
          let expr  = '';
          while (i < n && depth > 0) {
            if      (ch() === '{') { depth++; expr += ch(); i++; }
            else if (ch() === '}') {
              depth--;
              if (depth === 0) { i++; break; }
              expr += ch(); i++;
            } else { expr += ch(); i++; }
          }
          tmpl += expr.trim() + ") + '";

        } else if (ch() === "'") {
          // Comilla simple dentro del template → escapar
          tmpl += "\\'"; i++;

        } else if (ch() === '\r') {
          i++;   // ignorar CR

        } else if (ch() === '\n') {
          // Salto de línea real → secuencia de escape
          tmpl += '\\n'; i++;

        } else {
          tmpl += ch(); i++;
        }
      }

      // Limpiar concatenaciones vacías al inicio/fin: '' + ... y ... + ''
      let val = "'" + tmpl + "'";
      val = val.replace(/^'' \+ \(/, '(').replace(/\) \+ ''$/, ')');
      // Si quedó solo '' es una cadena vacía, dejarla
      out += val;

    } else {
      out += ch(); i++;
    }
  }

  return { result: out, count };
}

// ─── Procesa un archivo HTML ──────────────────────────────────────────────────
// Si tiene etiquetas <script>, solo convierte dentro de ellas.
// Si es JS puro (sin <script>), convierte el archivo completo.
function processHtml(content) {
  const hasScriptTags = /<script\b/i.test(content);

  if (!hasScriptTags) {
    // Archivo JS puro — procesar todo
    return convertTemplateLiterals(content);
  }

  // Archivo HTML mixto — solo los bloques <script>
  const parts = content.split(/(<script\b[^>]*>|<\/script\s*>)/gi);
  let result  = '';
  let inScript = false;
  let count   = 0;

  for (const part of parts) {
    if (/^<script\b/i.test(part)) {
      result += part;
      inScript = true;
    } else if (/^<\/script/i.test(part)) {
      result += part;
      inScript = false;
    } else if (inScript) {
      const { result: converted, count: c } = convertTemplateLiterals(part);
      result += converted;
      count  += c;
    } else {
      result += part;
    }
  }

  return { result, count };
}

// ─── Main ────────────────────────────────────────────────────────────────────
const files = fs.readdirSync(BASE)
  .filter(f => /^Mod_.*\.html$/.test(f))
  .sort();

console.log(`Procesando ${files.length} archivos Mod_*.html...\n`);

let totalFiles     = 0;
let totalConverted = 0;

for (const file of files) {
  const filePath = path.join(BASE, file);
  const original = fs.readFileSync(filePath, 'utf8');
  const { result, count } = processHtml(original);

  if (count > 0) {
    fs.writeFileSync(filePath, result, 'utf8');
    totalFiles++;
    totalConverted += count;
    console.log(`  ✓ ${file}: ${count} template literal(s) convertido(s)`);
  } else {
    console.log(`    ${file}: sin template literals`);
  }
}

console.log(`\nResumen: ${totalConverted} template literals convertidos en ${totalFiles} archivos.`);
