/**
 * rebuild_main.js
 * Regenera el bloque de módulos en MainApp_PlaneacionHTML.html.
 *
 * Busca el marcador:
 *   // ════════════════════════ MÓDULOS — rebuild_main.js ════════════════════════ //
 * y reemplaza todo lo que hay desde ese marcador hasta </script>
 * con el contenido real de cada Mod_*.html, en el orden definido.
 */
const fs   = require('fs');
const path = require('path');

const BASE      = __dirname;
const MAIN_PATH = path.join(BASE, 'MainApp_PlaneacionHTML.html');
const MARKER    = '// ════════════════════════ MÓDULOS — rebuild_main.js ════════════════════════ //';

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
const main = fs.readFileSync(MAIN_PATH, 'utf8');

// 2. Encontrar el marcador
const markerIdx = main.indexOf(MARKER);
if (markerIdx === -1) {
  console.error('ERROR: Marcador no encontrado en MainApp_PlaneacionHTML.html');
  console.error('  Marcador esperado: ' + MARKER);
  process.exit(1);
}

// 3. Todo lo que está antes del marcador se conserva intacto
const before = main.slice(0, markerIdx);

// 4. Encontrar el </script> que cierra el bloque (el primero que hay después del marcador)
const closeTag  = '\n</script>';
const closeIdx  = main.indexOf(closeTag, markerIdx);
if (closeIdx === -1) {
  console.error('ERROR: No se encontró </script> después del marcador.');
  process.exit(1);
}
const after = main.slice(closeIdx); // incluye \n</script>\n</body>...

// 5. Leer y concatenar módulos
let jsUnido = '';
for (const mod of ORDEN) {
  const filePath = path.join(BASE, mod + '.html');
  if (!fs.existsSync(filePath)) {
    console.warn(`  AVISO: ${mod}.html no existe, se omite.`);
    continue;
  }
  let contenido = fs.readFileSync(filePath, 'utf8');
  if (!contenido.endsWith('\n')) contenido += '\n';
  jsUnido += contenido;
  console.log(`  + ${mod}.html (${contenido.split('\n').length - 1} líneas)`);
}

// 6. Construir resultado: before + marcador + \n + JS módulos + </script>...
const result = before + MARKER + '\n' + jsUnido + after;

// 7. Guardar
fs.writeFileSync(MAIN_PATH, result, 'utf8');

const lineas = result.split('\n').length;
console.log(`\nMainApp_PlaneacionHTML.html reconstruido: ${lineas} líneas.`);
