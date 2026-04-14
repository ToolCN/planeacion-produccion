/**
 * fix_script_tags.js
 * Hace DOS cosas en MainApp_PlaneacionHTML.html:
 *
 * 1. CORRECCIÓN ESTRUCTURAL:
 *    - Elimina el <script> duplicado en la línea que separa el JS global
 *      de los módulos (fue introducido por rebuild_main.js).
 *    - Elimina el </script> huérfano que quedó después de </body>.
 *    Resultado: un único bloque <script>...</script> con todo el JS.
 *
 * 2. ESCAPE DE </script> EN STRINGS:
 *    - Dentro del bloque <script> resultante, busca </script> dentro de
 *      strings JavaScript (comillas simples o dobles) y los reemplaza
 *      por <\/script> para que el navegador no los interprete como cierre.
 */

const fs   = require('fs');
const path = require('path');

const FILE = path.join(__dirname, 'MainApp_PlaneacionHTML.html');
let content = fs.readFileSync(FILE, 'utf8');
const lines  = content.split('\n');

console.log(`Líneas originales: ${lines.length}`);

// ── PASO 1: Corrección estructural ───────────────────────────────────────────

// 1a. Eliminar el <script> huérfano que está DENTRO del bloque JS principal.
//     Es la línea que contiene exactamente "<script>" y está precedida por
//     comentarios del dashboard y seguida de funciones JS (no la de la línea 1083).
//
//     Estrategia: buscar todas las líneas que sean exactamente "<script>" (trim),
//     y eliminar todas excepto la PRIMERA (que es la apertura legítima del bloque).

let scriptOpenCount = 0;
let structFixes = 0;
const fixedLines = lines.filter((line, idx) => {
  if (line.trim() === '<script>') {
    scriptOpenCount++;
    if (scriptOpenCount > 1) {
      console.log(`  [fix-struct] Eliminado <script> extra en línea ${idx + 1}`);
      structFixes++;
      return false; // eliminar
    }
  }
  return true;
});

// 1b. Eliminar el </script> huérfano que viene DESPUÉS de </body>.
//     La secuencia incorrecta es: </script> → </body> → \n → </script> → </html>
//     El segundo </script> es el huérfano a eliminar.
let bodyFound     = false;
let orphanRemoved = false;
const fixedLines2 = fixedLines.filter((line) => {
  if (line.trim() === '</body>') bodyFound = true;
  if (bodyFound && !orphanRemoved && line.trim() === '</script>') {
    console.log(`  [fix-struct] Eliminado </script> huérfano tras </body>`);
    structFixes++;
    orphanRemoved = true;
    return false; // eliminar
  }
  return true;
});

// ── PASO 2: Escapar </script> dentro de strings JS ───────────────────────────

// Une las líneas y busca el bloque <script>...</script> principal.
// Procesamos todo el contenido del bloque con un state-machine.

const joined = fixedLines2.join('\n');

// Divide en: [antes-del-script, tag-apertura, contenido-js, tag-cierre, resto]
// La idea: split por la primera <script> standalone y la primera </script> tras ella.
const OPEN_RE  = /(<script>)\n/;
const CLOSE_RE = /\n(<\/script>)/;

const openMatch = OPEN_RE.exec(joined);
if (!openMatch) {
  console.error('ERROR: No se encontró <script> principal.');
  process.exit(1);
}

const beforeScript = joined.slice(0, openMatch.index + openMatch[0].length);
const afterOpen    = joined.slice(openMatch.index + openMatch[0].length);

const closeMatch = CLOSE_RE.exec(afterOpen);
if (!closeMatch) {
  console.error('ERROR: No se encontró </script> de cierre.');
  process.exit(1);
}

const jsContent   = afterOpen.slice(0, closeMatch.index);
const afterScript = afterOpen.slice(closeMatch.index);

console.log(`\nBloque JS: ${jsContent.split('\n').length} líneas`);

// State-machine: reemplaza </script> solo dentro de strings
function fixScriptTagsInStrings(src) {
  let out = '';
  let i   = 0;
  let hits = 0;
  const n  = src.length;
  function ch(o = 0) { return i + o < n ? src[i + o] : ''; }

  while (i < n) {
    // Comilla simple
    if (ch() === "'") {
      out += ch(); i++;
      while (i < n && ch() !== "'") {
        if (ch() === '\\') { out += ch() + ch(1); i += 2; continue; }
        if (src.slice(i, i + 9).toLowerCase() === '</script>') {
          out += '<\\/script>'; i += 9; hits++; continue;
        }
        out += ch(); i++;
      }
      if (i < n) { out += ch(); i++; }

    // Comilla doble
    } else if (ch() === '"') {
      out += ch(); i++;
      while (i < n && ch() !== '"') {
        if (ch() === '\\') { out += ch() + ch(1); i += 2; continue; }
        if (src.slice(i, i + 9).toLowerCase() === '</script>') {
          out += '<\\/script>'; i += 9; hits++; continue;
        }
        out += ch(); i++;
      }
      if (i < n) { out += ch(); i++; }

    // Comentario de línea
    } else if (ch() === '/' && ch(1) === '/') {
      out += ch(); i++;
      while (i < n && ch() !== '\n') { out += ch(); i++; }

    // Comentario de bloque
    } else if (ch() === '/' && ch(1) === '*') {
      out += ch() + ch(1); i += 2;
      while (i < n && !(ch() === '*' && ch(1) === '/')) { out += ch(); i++; }
      if (i < n) { out += ch() + ch(1); i += 2; }

    } else {
      out += ch(); i++;
    }
  }
  return { result: out, hits };
}

const { result: fixedJs, hits } = fixScriptTagsInStrings(jsContent);

// ── Resultado final ───────────────────────────────────────────────────────────
const final = beforeScript + fixedJs + afterScript;
fs.writeFileSync(FILE, final, 'utf8');

const finalLines = final.split('\n').length;
console.log(`\n✓ Correcciones estructurales: ${structFixes}`);
console.log(`✓ </script> escapados en strings: ${hits}`);
console.log(`✓ Líneas finales: ${finalLines}`);

// Verificar estructura resultante
const scriptOpens  = (final.match(/^<script>$/mg)  || []).length;
const scriptCloses = (final.match(/^<\/script>$/mg) || []).length;
console.log(`\nVerificación: <script> standalone = ${scriptOpens}, </script> standalone = ${scriptCloses}`);
if (scriptOpens === scriptCloses) {
  console.log('✓ Estructura <script> balanceada.');
} else {
  console.log('⚠ Estructura desbalanceada, revisar manualmente.');
}
