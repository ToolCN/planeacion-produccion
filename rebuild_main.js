/**
 * rebuild_main.js
 * Sustituye el bloque <script> con includes de GAS en MainApp_PlaneacionHTML.html
 * por el contenido real de cada Mod_*.html concatenado, dentro de un único <script>.
 */
const fs   = require('fs');
const path = require('path');

const BASE      = __dirname;
const MAIN_PATH = path.join(BASE, 'MainApp_PlaneacionHTML.html');

// Orden exacto de los módulos
const ORDEN = [
  'Mod_Dashboard',
  'Mod_Metricas',
  'Mod_Validador',
  'Mod_Tracking',
  'Mod_NuevoPedido',
  'Mod_PedidoINT',
  'Mod_Generador',
  'Mod_Planificador',
  'Mod_Programador',
  'Mod_ProgInteractivo',
  'Mod_Wip',
  'Mod_Captura',
  'Mod_EditorProd',
  'Mod_NuevaRuta',
  'Mod_EditorRutas',
  'Mod_TableroSupervisor',
  'Mod_GestionDocs',
  'Mod_Usuarios',
  'Mod_TrackingModule',
  'Mod_MrpMp',
  'Mod_SalidaAlambres',
  'Mod_MisRegistros',
  'Mod_PedidosPendientes',
];

// 1. Leer MainApp
let main = fs.readFileSync(MAIN_PATH, 'utf8');

// 2. Encontrar el bloque <script> que contiene los includes
//    Captura desde <script>\n hasta </script> inclusive, cuando contiene includes GAS
const BLOCK_RE = /<script>\n(?:[ \t]*<\?!=\s*include\([^)]+\);\s*\?>\n)+<\/script>/;
const match = main.match(BLOCK_RE);

if (!match) {
  console.error('ERROR: No se encontró el bloque <script> con includes. Abortando.');
  process.exit(1);
}

console.log(`Bloque encontrado (${match[0].split('\n').length} líneas):`);
console.log(match[0].slice(0, 200) + '...\n');

// 3. Leer y concatenar el contenido de cada módulo
let jsUnido = '';
for (const mod of ORDEN) {
  const filePath = path.join(BASE, mod + '.html');
  if (!fs.existsSync(filePath)) {
    console.warn(`  AVISO: ${mod}.html no existe, se omite.`);
    continue;
  }
  let contenido = fs.readFileSync(filePath, 'utf8');
  // Asegurar que termina con salto de línea
  if (!contenido.endsWith('\n')) contenido += '\n';
  jsUnido += contenido;
  console.log(`  + ${mod}.html (${contenido.split('\n').length - 1} líneas)`);
}

// 4. Construir el bloque de reemplazo
const REEMPLAZO = '<script>\n' + jsUnido + '</script>';

// 5. Sustituir en el MainApp
main = main.replace(BLOCK_RE, REEMPLAZO);

// 6. Guardar
fs.writeFileSync(MAIN_PATH, main, 'utf8');

const lineas = main.split('\n').length;
console.log(`\nMainApp_PlaneacionHTML.html reconstruido: ${lineas} líneas.`);
