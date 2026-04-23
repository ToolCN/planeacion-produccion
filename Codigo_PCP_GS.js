// =================================================================================
// CODIGO_PCP_GS — Script completo
// Web App URL: https://script.google.com/macros/s/AKfycbyDxTSHPygL1Ba93kHFisiInmZZf6zHM63jOHfqPGv45bdbOJ1Av2lg-DigAGjH50Co1w/exec
// =================================================================================

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// CONFIGURACION GLOBAL
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
var ID_HOJA_CALCULO = "1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk";
var HOJA_PEDIDOS    = "PEDIDOS";
var HOJA_ORDENES    = "ORDENES";
var ID_HOJA_OM = "1v21_Glgvk3ZV4SYpsMGbqNc97I7MO98BqkwmiJvYvnI"; 
var TOKEN = "7947767393:AAFmZUcSTnV5gvP6u_UsBcSHlz-0s9x1kSQ";
var CHAT_ID_CALIDAD = "-1003608646187";
var THREAD_ID_SELLOS = "1250"; 
var TOKEN_TELEGRAM = "7947767393:AAFmZUcSTnV5gvP6u_UsBcSHlz-0s9x1kSQ";
var CHAT_ID_AVISOS = "625827165";


// =================================================================================
// doGet — Sirve el MainApp_PlaneacionHTML como Web App
// =================================================================================
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("MainApp_PlaneacionHTML")
    .setTitle("PCP — ToolCN")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Devuelve la URL de la propia Web App (la usa el HTML internamente)
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// =================================================================================
// USUARIOS — Login
// =================================================================================

function obtenerUsuariosPCP() {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var usuarios = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue;

      var nombre        = String(row[1]).trim().toUpperCase(); // col B
      var password      = String(row[2] || "").trim();         // col C
      var rol           = String(row[3] || "").trim();         // col D = ROL (Admin / "")
      var celdaPermisos = String(row[5] || "").trim();         // col F = DEPARTAMENTO (permisos)

      // Separar todos los permisos de la celda
      var todos = celdaPermisos.split(',').map(function(p){ return p.trim(); }).filter(function(p){ return p; });

      // Solo los P_ son de esta app
      var permisosPCP = todos.filter(function(p){ return p.indexOf('P_') === 0; });

      // Si ROL = Admin -> acceso total: agregar P_Admin si no está
      if (rol.toLowerCase() === 'admin') {
        if (permisosPCP.indexOf('P_Admin') < 0) permisosPCP.push('P_Admin');
      }

      // Los permisos que NO son P_ (de otras apps) se guardan aparte para no perderlos
      var otrosPermisos = todos.filter(function(p){ return p.indexOf('P_') !== 0; });

      usuarios.push({
        nombre:         nombre,
        password:       password,
        rol:            rol,
        permisos:       permisosPCP.join(','),   // solo P_ → lo usa el HTML
        _otrosPermisos: otrosPermisos.join(',')  // otros → para conservar al guardar
      });
    }
    return usuarios;
  } catch (e) {
    return [];
  }
}

function guardarPermisosPCP(nombre, nuevosPCP) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return JSON.stringify({ success: false, msg: "Hoja USUARIOS no encontrada" });

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() !== nombre.toUpperCase()) continue;

      // Celda actual col F = DEPARTAMENTO
      var celdaActual = String(data[i][5] || "").trim();

      // Conservar los que NO son P_
      var otros = celdaActual.split(',')
        .map(function(p){ return p.trim(); })
        .filter(function(p){ return p && p.indexOf('P_') !== 0; });

      // Limpiar los P_ nuevos
      var nuevosP = nuevosPCP.split(',')
        .map(function(p){ return p.trim(); })
        .filter(function(p){ return p && p.indexOf('P_') === 0; });

      // Resultado: otros primero, luego los P_ nuevos
      var resultado = otros.concat(nuevosP).join(',');

      sheet.getRange(i + 1, 6).setValue(resultado); // col F = DEPARTAMENTO
      return JSON.stringify({ success: true, msg: "Permisos guardados" });
    }

    return JSON.stringify({ success: false, msg: "Usuario no encontrado: " + nombre });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function cambiarPasswordPCP(nombre, nuevaPass) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return JSON.stringify({ success: false, msg: "Hoja USUARIOS no encontrada" });

    if (!nuevaPass || nuevaPass.trim().length < 1) {
      return JSON.stringify({ success: false, msg: "La contraseña no puede estar vacía" });
    }

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var nombreFila = String(data[i][1]).trim().toUpperCase();
      if (nombreFila !== nombre.toUpperCase().trim()) continue;

      // Guardar nueva contraseña en col C (índice 2, columna 3)
      sheet.getRange(i + 1, 3).setValue(nuevaPass.trim());
      return JSON.stringify({ success: true, msg: "Contraseña actualizada" });
    }

    return JSON.stringify({ success: false, msg: "Usuario no encontrado: " + nombre });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function crearUsuarioPCP(data) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return JSON.stringify({ success: false, msg: "Hoja USUARIOS no encontrada" });

    var nombre   = String(data.nombre   || "").trim().toUpperCase();
    var password = String(data.password || "").trim();
    var permisos = String(data.permisos || "").trim();

    if (!nombre)   return JSON.stringify({ success: false, msg: "El nombre es obligatorio" });
    if (!password) return JSON.stringify({ success: false, msg: "La contraseña es obligatoria" });

    // Verificar que no exista ya
    var existentes = sheet.getDataRange().getValues();
    for (var i = 1; i < existentes.length; i++) {
      if (String(existentes[i][1]).trim().toUpperCase() === nombre) {
        return JSON.stringify({ success: false, msg: "Ya existe un usuario con ese nombre" });
      }
    }

    // Determinar ROL: si tiene P_Admin → guardar "Admin" en col D
    var rol = permisos.split(',').map(function(p){ return p.trim(); }).indexOf('P_Admin') > -1
              ? 'Admin'
              : '';

    // Agregar fila: [colA vacía, NOMBRE, PASSWORD, ROL, col E vacía, PERMISOS en col F]
    sheet.appendRow(['', nombre, password, rol, '', permisos]);

    return JSON.stringify({ success: true, msg: "Usuario creado: " + nombre });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function actualizarUsuarioPCP(data) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return JSON.stringify({ success: false, msg: "Hoja USUARIOS no encontrada" });

    var nombre       = String(data.nombre   || "").trim().toUpperCase();
    var nuevosP      = String(data.permisos || "").trim();
    var nuevaPass    = data.password ? String(data.password).trim() : null;

    if (!nombre) return JSON.stringify({ success: false, msg: "Nombre obligatorio" });

    var filas = sheet.getDataRange().getValues();

    for (var i = 1; i < filas.length; i++) {
      var nombreFila = String(filas[i][1]).trim().toUpperCase();
      if (nombreFila !== nombre) continue;

      // ── CONTRASEÑA (col C) ──
      if (nuevaPass && nuevaPass.length > 0) {
        sheet.getRange(i + 1, 3).setValue(nuevaPass);
      }

      // ── ROL (col D) ── actualizar según si tiene P_Admin o no
      var listaP = nuevosP.split(',').map(function(p){ return p.trim(); }).filter(Boolean);
      var esAdmin = listaP.indexOf('P_Admin') > -1;
      sheet.getRange(i + 1, 4).setValue(esAdmin ? 'Admin' : '');

      // ── PERMISOS (col F = DEPARTAMENTO) — conservar los que NO son P_ ──
      var celdaActual = String(filas[i][5] || "").trim();
      var otros = celdaActual.split(',')
        .map(function(p){ return p.trim(); })
        .filter(function(p){ return p && p.indexOf('P_') !== 0; });

      var soloP = listaP.filter(function(p){ return p.indexOf('P_') === 0; });

      var resultado = otros.concat(soloP).join(',');
      sheet.getRange(i + 1, 6).setValue(resultado);

      return JSON.stringify({ success: true, msg: "Usuario actualizado: " + nombre });
    }

    return JSON.stringify({ success: false, msg: "Usuario no encontrado: " + nombre });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function eliminarUsuarioPCP(nombre) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return JSON.stringify({ success: false, msg: "Hoja USUARIOS no encontrada" });

    var nombreBuscar = String(nombre || "").trim().toUpperCase();
    if (!nombreBuscar) return JSON.stringify({ success: false, msg: "Nombre obligatorio" });

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var nombreFila = String(data[i][1]).trim().toUpperCase();
      if (nombreFila !== nombreBuscar) continue;

      // Eliminar la fila completa (i+1 porque los valores empiezan en 1)
      sheet.deleteRow(i + 1);
      return JSON.stringify({ success: true, msg: "Usuario eliminado: " + nombreBuscar });
    }

    return JSON.stringify({ success: false, msg: "Usuario no encontrado: " + nombre });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS COMPARTIDAS POR VARIOS MODULOS
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function obtenerIconoSVG(tipo) {
  // Limpiamos el texto para evitar errores por espacios o nulos
  var t = String(tipo || "").toUpperCase().trim();
  var svgStart = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" style="width:100%; height:100%;">';
  var path = "";

  // --- 1. BLOQUE DE CLAVOS ---
  if (t.includes("CLAVO")) {
      if (t.includes("COLATADO") || t.includes("ROOFING")) {
          path = '<rect x="6" y="4" width="2" height="16" fill="#9e9e9e"/><rect x="5" y="3" width="4" height="1" fill="#757575"/><path d="M6 20 L7 22 L8 20 Z" fill="#9e9e9e"/><rect x="16" y="4" width="2" height="16" fill="#9e9e9e"/><rect x="15" y="3" width="4" height="1" fill="#757575"/><path d="M16 20 L17 22 L18 20 Z" fill="#9e9e9e"/><line x1="6" y1="8" x2="18" y2="8" stroke="#ff8f00" stroke-width="1"/><line x1="6" y1="14" x2="18" y2="14" stroke="#ff8f00" stroke-width="1"/>';
      }
      else if (t.includes("BOYA")) {
          path = '<rect x="8" y="4" width="8" height="16" fill="#212121"/><rect x="6" y="2" width="12" height="2" fill="#000"/><path d="M8 20 L12 24 L16 20 Z" fill="#212121"/><line x1="9" y1="6" x2="15" y2="8" stroke="#616161" stroke-width="0.5"/><line x1="9" y1="9" x2="15" y2="11" stroke="#616161" stroke-width="0.5"/><line x1="9" y1="12" x2="15" y2="14" stroke="#616161" stroke-width="0.5"/><line x1="9" y1="15" x2="15" y2="17" stroke="#616161" stroke-width="0.5"/>';
      }
      else if (t.includes("CONCRETO") && (t.includes("LISO") || t.includes("LISA"))) {
          path = '<rect x="10" y="4" width="4" height="16" fill="#212121"/><rect x="8" y="2" width="8" height="2" fill="#000"/><path d="M10 20 L12 24 L14 20 Z" fill="#212121"/>';
      }
      else if (t.includes("CONCRETO")) {
          path = '<rect x="10" y="4" width="4" height="16" fill="#212121"/><rect x="8" y="2" width="8" height="2" fill="#000"/><path d="M10 20 L12 24 L14 20 Z" fill="#212121"/><line x1="10" y1="6" x2="14" y2="7" stroke="#757575" stroke-width="0.5"/><line x1="10" y1="9" x2="14" y2="10" stroke="#757575" stroke-width="0.5"/><line x1="10" y1="12" x2="14" y2="13" stroke="#757575" stroke-width="0.5"/><line x1="10" y1="15" x2="14" y2="16" stroke="#757575" stroke-width="0.5"/>';
      }
      else {
          path = '<rect x="11" y="4" width="2" height="17" fill="#9e9e9e"/><rect x="9" y="3" width="6" height="1" fill="#616161"/><path d="M11 21 L12 24 L13 21 Z" fill="#9e9e9e"/>';
      }
  }

  // --- 2. BLOQUE DE TORNILLOS Y OTROS (ELSE IF PARA EVITAR SOBREESCRITURA) ---
  else if (t.includes("CUA") && t.includes("CN")) {
    path = '<rect x="4" y="2" width="16" height="8" rx="1" fill="#78909c" stroke="#37474f"/> <rect x="9" y="10" width="6" height="12" fill="#b0bec5"/> <line x1="9" y1="14" x2="15" y2="14" stroke="#78909c"/> <line x1="9" y1="18" x2="15" y2="18" stroke="#78909c"/> <text x="12" y="8" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle" fill="black">CN</text>';
  }
  else if (t.includes("CUA") && t.includes("CV")) {
    path = '<rect x="4" y="2" width="16" height="8" rx="1" fill="#78909c" stroke="#37474f"/> <rect x="9" y="10" width="6" height="12" fill="#b0bec5"/> <line x1="9" y1="14" x2="15" y2="14" stroke="#78909c"/> <line x1="9" y1="18" x2="15" y2="18" stroke="#78909c"/> <text x="12" y="8" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle" fill="black">CV</text>';
  }
  else if (t.includes("CAR")) {
    path = '<path d="M4 9 Q12 1 20 9" fill="#90a4ae" stroke="#546e7a"/> <rect x="8" y="9" width="8" height="4" fill="#607d8b"/> <rect x="9" y="13" width="6" height="10" fill="#cfd8dc"/> <line x1="9" y1="16" x2="15" y2="16" stroke="#90a4ae"/> <line x1="9" y1="19" x2="15" y2="19" stroke="#90a4ae"/>';
  }
  else if (t.includes("B7") || t.includes("B-7")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle" fill="#333">B-7</text>';
  }
  else if (t.includes("A394") && (t.includes("T0") || t.includes("T-0"))) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle" fill="#333">T-0</text>';
  }
  else if (t.includes("RED") && t.includes("RAN")) {
    path = '<path d="M16 6a4 4 0 0 0-8 0" fill="#78909c"/> <rect x="11" y="2" width="2" height="4" fill="#333"/> <rect x="10" y="6" width="4" height="16" fill="#b0bec5"/> <line x1="10" y1="10" x2="14" y2="10" stroke="#78909c"/> <line x1="10" y1="14" x2="14" y2="14" stroke="#78909c"/> <line x1="10" y1="18" x2="14" y2="18" stroke="#78909c"/>';
  }
  else if (t.includes("PIJ") || t.includes("PIJA") || t.includes("PHI")) {
    path = '<path d="M17 5H7V3h10v2z M10 5l2 14 2-14" fill="#757575"/> <path d="M11 9h2 M11 12h2 M11 15h2" stroke="#424242"/>';
  }
  else if (t.includes("MAC")) {
    path = '<rect x="6" y="8" width="12" height="14" fill="#90a4ae" stroke="#546e7a"/> <rect x="10" y="2" width="4" height="6" fill="#cfd8dc" stroke="#546e7a"/>';
  }
  else if (t.includes("HEM")) {
    path = '<path d="M6 2h12v20H6z M10 2v6h4V2" fill="#90a4ae" fill-rule="evenodd" stroke="#546e7a"/>';
  }
  else if (t.includes("HEX") && t.includes("G2")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle" fill="#333">G2</text>';
  }
  else if (t.includes("G5")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <line x1="12" y1="12" x2="12" y2="4" stroke="black" stroke-width="1.5"/> <line x1="12" y1="12" x2="5" y2="16" stroke="black" stroke-width="1.5"/> <line x1="12" y1="12" x2="19" y2="16" stroke="black" stroke-width="1.5"/>';
  }
  else if (t.includes("G8")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <line x1="12" y1="6" x2="12" y2="4" stroke="black" stroke-width="2"/> <line x1="12" y1="18" x2="12" y2="20" stroke="black" stroke-width="2"/> <line x1="6" y1="9" x2="4" y2="8" stroke="black" stroke-width="2"/> <line x1="18" y1="9" x2="20" y2="8" stroke="black" stroke-width="2"/> <line x1="6" y1="15" x2="4" y2="16" stroke="black" stroke-width="2"/> <line x1="18" y1="15" x2="20" y2="16" stroke="black" stroke-width="2"/>';
  }
  else if (t.includes("A325") || t.includes("A-325")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle">A325</text>';
  }
  else if (t.includes("A490") || t.includes("A-490")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle">A490</text>';
  }
  else if (t.includes("T-1") || t.includes("T1")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle">T-1</text>';
  }
  else if (t.includes("BIS HEM")) {
    path = '<rect x="4" y="4" width="10" height="16" fill="#90a4ae"/> <circle cx="18" cy="12" r="3" fill="none" stroke="#546e7a" stroke-width="2"/>';
  }
  else if (t.includes("BIS MAC")) {
    path = '<rect x="10" y="4" width="10" height="16" fill="#90a4ae"/> <rect x="4" y="10" width="6" height="4" fill="#546e7a"/>';
  }
  else if (t.includes("ARM")) {
    path = '<path d="M12 2a6 6 0 1 0 0 12 6 6 0 0 0 0-12zm0 2a4 4 0 1 1 0 8 4 4 0 0 1 0-8z M12 14v8" stroke="#616161" stroke-width="2" fill="none"/>';
  }
  else if (t.includes("BIR") || t.includes("BIRLO")) {
    path = '<rect x="8" y="2" width="8" height="20" fill="#bdbdbd"/> <line x1="8" y1="6" x2="16" y2="6" stroke="#424242"/> <line x1="8" y1="10" x2="16" y2="10" stroke="#424242"/> <line x1="8" y1="14" x2="16" y2="14" stroke="#424242"/> <line x1="8" y1="18" x2="16" y2="18" stroke="#424242"/>';
  }
  else if (t.includes("GRA") || t.includes("GRAPA")) {
    path = '<path d="M6 18V8a6 6 0 0 1 12 0v10" fill="none" stroke="#616161" stroke-width="3"/> <path d="M6 18l-2-2 M18 18l2-2" stroke="#616161" stroke-width="2"/>';
  }
  else if (t.includes("PER") || t.includes("PERNO")) {
  path = '<rect x="5" y="2" width="14" height="4" rx="1" fill="#424242"/>' + // Cabeza plana y ancha
             '<rect x="9" y="6" width="6" height="15" fill="#757575"/>' +      // Cuerpo
             '<path d="M9 21 Q12 23 15 21 Z" fill="#757575"/>';               // Base levemente redondeada
  }
  else if (t.includes("REM") || t.includes("REMACHE")) {
    path = '<path d="M12 4a4 4 0 0 0-4 4v2h8V8a4 4 0 0 0-4-4z" fill="#78909c"/> <rect x="10" y="10" width="4" height="10" fill="#b0bec5"/>';
  }
  else if (t.includes("VARILLA")) {
     path = '<rect x="7" y="2" width="10" height="20" fill="#bdbdbd" stroke="#424242" stroke-width="0.5"/>' +
             '<line x1="7" y1="4" x2="17" y2="6" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="6" x2="17" y2="8" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="8" x2="17" y2="10" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="10" x2="17" y2="12" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="12" x2="17" y2="14" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="14" x2="17" y2="16" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="16" x2="17" y2="18" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="18" x2="17" y2="20" stroke="#616161" stroke-width="0.8"/>';
  }
  else if (t.includes("BARRA") || t.includes("COLD") || t.includes("REDONDO")) {
    var fill = "#bdbdbd";
    if (t.includes("COLD")) fill = "#a5d6a7";
    else if (t.includes("REDONDO")) fill = "#fff59d";
    path = '<rect x="4" y="2" width="16" height="20" fill="' + fill + '" stroke="black" stroke-width="1.0"/>';
  }
  else if (t.includes("ALA") || t.includes("ALAMBRE") || t.includes("ROLLO")) {
      path = '<g fill="none" stroke="#f57c00" stroke-width="2"><circle cx="12" cy="12" r="11"/><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="7"/></g>';
  }
  
  // --- 3. FALLBACK FINAL: SI NO COINCIDIÓ NADA ---
  if (path === "") {
      path = '<rect x="2" y="6" width="20" height="12" fill="#cfd8dc" stroke="#78909c" stroke-width="2"/> <text x="12" y="15" font-family="Arial" font-weight="bold" font-size="7" text-anchor="middle" fill="#455a64">ESP</text>';
  }

  return svgStart + path + "</svg>";
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS PARA Dashboard
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
/**
 * Devuelve los KPIs del Dashboard para el mes/año indicado.
 * @param {number} mes   1-12
 * @param {number} anio  ej. 2026
 */
function obtenerDatosDashboardPCP(mes, anio) {
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetPed = ss.getSheetByName(HOJA_PEDIDOS);
    var sheetOrd = ss.getSheetByName(HOJA_ORDENES);
    var sheetCod = ss.getSheetByName("CODIGOS");
    if (!sheetPed || !sheetOrd) throw new Error("Hojas no encontradas");

    var dataPed = sheetPed.getDataRange().getValues();
    var dataOrd = sheetOrd.getDataRange().getValues();

    // ── Semana ISO ─────────────────────────────────────────────────
    function isoWeek(d) {
      var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      var day = t.getUTCDay() || 7;
      t.setUTCDate(t.getUTCDate() + 4 - day);
      var y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
      return Math.ceil((((t - y0) / 86400000) + 1) / 7);
    }
    function isoWeekYear(d) {
      var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      var day = t.getUTCDay() || 7;
      t.setUTCDate(t.getUTCDate() + 4 - day);
      return t.getUTCFullYear();
    }
    function weekKey(d) { return isoWeekYear(d) + '-W' + ('0' + isoWeek(d)).slice(-2); }
    function getMondayOfISOWeek(wy, wn) {
      var jan4 = new Date(Date.UTC(wy, 0, 4));
      var dow4 = jan4.getUTCDay() || 7;
      var mon  = new Date(jan4.getTime());
      mon.setUTCDate(jan4.getUTCDate() - (dow4 - 1) + (wn - 1) * 7);
      return mon;
    }
    function fmtDia(d) { return ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2); }

    // Ventana 8 semanas hacia atrás desde hoy
    var hoy = new Date();
    var curWY = isoWeekYear(hoy), curWN = isoWeek(hoy);
    var semanaKeys = [];
    var tmpWY = curWY, tmpWN = curWN;
    for (var w = 0; w < 8; w++) {
      semanaKeys.unshift(tmpWY + '-W' + ('0' + tmpWN).slice(-2));
      tmpWN--;
      if (tmpWN < 1) {
        tmpWY--;
        var d31 = new Date(Date.UTC(tmpWY, 11, 31));
        tmpWN = isoWeek(d31);
      }
    }

    // ── PEDIDOS ────────────────────────────────────────────────────
    var pedRecibidos = 0, pedAbiertos = 0, pedEnProceso = 0, pedTerminados = 0, pedCancelados = 0;
    var SERIES_PIE    = ['ZEQ','ZRR','MPR','MAQ','QSQ'];
    var COLORES_SERIE = { ZEQ:'#f8bbd0', ZRR:'#ffe0b2', MPR:'#fff9c4', MAQ:'#bbdefb', QSQ:'#c8e6c9' };
    var EXCLUIR       = ['INT','TEM'];
    var serieCount    = {};
    var wkPed         = {};
    semanaKeys.forEach(function(k){ wkPed[k] = 0; });

    // Alertas: listas de folios
    var alertZeqAnt   = [];
    var alertZeqAbi   = [];
    var alertZeqPlan  = [];
    var alertMprPend  = [];
    var alertQsqAbi   = [];

    var hoyMes  = hoy.getMonth() + 1;
    var hoyAnio = hoy.getFullYear();

    for (var i = 1; i < dataPed.length; i++) {
      var row      = dataPed[i];
      if (!row[1]) continue;
      var folio    = String(row[1] || '').trim();
      var fechaRaw = row[2];
      var fecha    = (fechaRaw instanceof Date) ? fechaRaw : new Date(fechaRaw);
      if (isNaN(fecha.getTime())) continue;
      var fMes  = fecha.getMonth() + 1;
      var fAnio = fecha.getFullYear();
      var est   = String(row[8] || '').toUpperCase().trim();
      var prefijo = folio.split('-')[0];

      // KPIs mes seleccionado
      if (fMes === mes && fAnio === anio) {
        pedRecibidos++;
        if (est === 'ABIERTO')         pedAbiertos++;
        else if (est === 'EN PROCESO') pedEnProceso++;
        else if (est === 'TERMINADO')  pedTerminados++;
        else if (est === 'CANCELADO')  pedCancelados++;
        if (EXCLUIR.indexOf(prefijo) < 0 && SERIES_PIE.indexOf(prefijo) > -1) {
          serieCount[prefijo] = (serieCount[prefijo] || 0) + 1;
        }
      }

      // Semanas históricas
      if (EXCLUIR.indexOf(prefijo) < 0) {
        var wk = weekKey(fecha);
        if (wkPed.hasOwnProperty(wk)) wkPed[wk]++;
      }

      // ── 5 ALERTAS ──────────────────────────────────────────────
      var noTerminado = (est !== 'CANCELADO' && est !== 'TERMINADO');

      // 1. ZEQ meses anteriores sin cerrar
      if (prefijo === 'ZEQ' && noTerminado &&
          (fAnio < hoyAnio || (fAnio === hoyAnio && fMes < hoyMes))) {
        alertZeqAnt.push({ f:folio, d:String(row[4]||'').trim(), c:Number(row[6]||0), u:String(row[7]||'').trim() });
      }
      // 2. ZEQ mes actual, ABIERTO
      if (prefijo === 'ZEQ' && est === 'ABIERTO' &&
          fMes === hoyMes && fAnio === hoyAnio) {
        alertZeqAbi.push({ f:folio, d:String(row[4]||'').trim(), c:Number(row[6]||0), u:String(row[7]||'').trim() });
      }
      // 3. ZEQ mes actual, no CANCELADO/TERMINADO
      if (prefijo === 'ZEQ' && noTerminado &&
          fMes === hoyMes && fAnio === hoyAnio) {
        alertZeqPlan.push({ f:folio, d:String(row[4]||'').trim(), c:Number(row[6]||0), u:String(row[7]||'').trim() });
      }
      // 4. MPR cualquier fecha, no CANCELADO/TERMINADO
      if (prefijo === 'MPR' && noTerminado) {
        alertMprPend.push({ f:folio, d:String(row[4]||'').trim(), c:Number(row[6]||0), u:String(row[7]||'').trim() });
      }
      // 5. QSQ cualquier fecha, ABIERTO
      if (prefijo === 'QSQ' && est === 'ABIERTO') {
        alertQsqAbi.push({ f:folio, d:String(row[4]||'').trim(), c:Number(row[6]||0), u:String(row[7]||'').trim() });
      }
    }

    // ── ÓRDENES ────────────────────────────────────────────────────
    var ordGeneradas = 0, ordTerminadas = 0, ordEnProceso = 0, ordCanceladas = 0;
    var ordenesVistas = {};
    var wkOrdGen  = {}, wkOrdTerm = {};
    semanaKeys.forEach(function(k){ wkOrdGen[k] = {}; wkOrdTerm[k] = {}; });

    // Mapa codigo->familia
    var mapCodFam = {};
    if (sheetCod) {
      var dataCod = sheetCod.getDataRange().getValues();
      var hCod = dataCod[0].map(function(x){ return String(x).toUpperCase().trim(); });
      var icCod = hCod.indexOf('CODIGO'); if (icCod < 0) icCod = 0;
      var icFam = hCod.indexOf('FAMILY'); if (icFam < 0) icFam = hCod.indexOf('FAMILIA'); if (icFam < 0) icFam = 8;
      for (var c = 1; c < dataCod.length; c++) {
        var ck = String(dataCod[c][icCod] || '').trim();
        var fk = String(dataCod[c][icFam] || 'OTROS').toUpperCase().trim();
        if (ck) mapCodFam[ck] = fk;
      }
    }
    var porFamilia = {};

    for (var i = 1; i < dataOrd.length; i++) {
      var row       = dataOrd[i];
      if (!row[1]) continue;
      var nOrden    = String(row[5]  || '').trim();
      var fechaORaw = row[3];
      var fechaO    = (fechaORaw instanceof Date) ? fechaORaw : new Date(fechaORaw);
      if (isNaN(fechaO.getTime())) continue;
      var fOMes  = fechaO.getMonth() + 1;
      var fOAnio = fechaO.getFullYear();
      var estOrd = String(row[15] || '').toUpperCase().trim();
      var serie5 = String(row[4]  || '').trim().split('-')[0];
      if (!nOrden) continue;

      if (fOMes === mes && fOAnio === anio) {
        if (!ordenesVistas[nOrden]) {
          ordenesVistas[nOrden] = estOrd;
          ordGeneradas++;
          if (estOrd === 'TERMINADO' || estOrd === 'SOBREPRODUCCION') ordTerminadas++;
          else if (estOrd === 'EN PROCESO' || estOrd === 'ACTIVE')    ordEnProceso++;
          else if (estOrd === 'CANCELADO')                             ordCanceladas++;
        }
        if (EXCLUIR.indexOf(serie5) < 0) {
          var codItem  = String(row[6] || '').replace(/^'/, '').trim();
          var pedFolio = String(row[1] || '').trim();
          var fam      = (codItem && mapCodFam[codItem]) ? mapCodFam[codItem] : 'OTROS';
          if (!porFamilia[fam]) porFamilia[fam] = { pedidosSet: {}, solicitado: 0, producido: 0 };
          if (pedFolio) porFamilia[fam].pedidosSet[pedFolio] = true;
          porFamilia[fam].solicitado += Number(row[13] || 0);
          porFamilia[fam].producido  += Number(row[14] || 0);
        }
      }

      var wkO = weekKey(fechaO);
      if (wkOrdGen.hasOwnProperty(wkO) && !wkOrdGen[wkO][nOrden]) {
        wkOrdGen[wkO][nOrden] = true;
        if (estOrd === 'TERMINADO' || estOrd === 'SOBREPRODUCCION') wkOrdTerm[wkO][nOrden] = true;
      }
    }

    var ordBase = ordGeneradas - ordCanceladas;
    var pctCumplimiento = ordBase > 0 ? Math.round(ordTerminadas / ordBase * 100) : 0;

    // ── SEMANAS ────────────────────────────────────────────────────
    var semanasArr = semanaKeys.map(function(wk, idx) {
      var parts = wk.split('-W');
      var wy = parseInt(parts[0]), wn = parseInt(parts[1]);
      var mon = getMondayOfISOWeek(wy, wn);
      var sun = new Date(mon.getTime()); sun.setUTCDate(mon.getUTCDate() + 6);
      var label = 'S' + wn + ' (' + fmtDia(mon) + '-' + fmtDia(sun) + ')';
      var ords  = wkOrdGen[wk] ? Object.keys(wkOrdGen[wk]).length : 0;
      var terms = wkOrdTerm[wk] ? Object.keys(wkOrdTerm[wk]).length : 0;
      var peds  = wkPed[wk] || 0;
      var cumpl = ords > 0 ? Math.round(terms / ords * 100) : 0;
      return { semana: label, semNum: wn, pedidos: peds, ordenes: ords, ordenesTerminadas: terms, cumplimiento: cumpl };
    });

    // ── PIE ────────────────────────────────────────────────────────
    var graficaPastel = SERIES_PIE
      .filter(function(s){ return (serieCount[s] || 0) > 0; })
      .map(function(s){ return { label: s, value: serieCount[s], color: COLORES_SERIE[s] }; });

    // ── FAMILIAS ───────────────────────────────────────────────────
    var familiasArr = Object.keys(porFamilia).map(function(f) {
      var obj = porFamilia[f];
      var pct = obj.solicitado > 0 ? Math.round(obj.producido / obj.solicitado * 100) : 0;
      return { familia: f, pedidos: Object.keys(obj.pedidosSet).length,
               solicitado: Math.round(obj.solicitado), producido: Math.round(obj.producido), pct: pct };
    }).sort(function(a, b){ return b.solicitado - a.solicitado; });

    return JSON.stringify({
      success: true,
      pedidos:  { recibidos: pedRecibidos, abiertos: pedAbiertos, enProceso: pedEnProceso,
                  terminados: pedTerminados, cancelados: pedCancelados },
      ordenes:  { generadas: ordGeneradas, terminadas: ordTerminadas, enProceso: ordEnProceso,
                  canceladas: ordCanceladas, pctCumplimiento: pctCumplimiento },
      graficaPastel: graficaPastel,
      semanasArr:    semanasArr,
      familiasArr:   familiasArr,
      alertas: {
        zeqAnt:  { cnt: alertZeqAnt.length,  lista: alertZeqAnt  },
        zeqAbi:  { cnt: alertZeqAbi.length,  lista: alertZeqAbi  },
        zeqPlan: { cnt: alertZeqPlan.length, lista: alertZeqPlan },
        mprPend: { cnt: alertMprPend.length, lista: alertMprPend },
        qsqAbi:  { cnt: alertQsqAbi.length,  lista: alertQsqAbi  }
      }
    });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function obtenerDashboardStock() {
  try {
    var ss  = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var tz  = ss.getSpreadsheetTimeZone();
    var res = obtenerFamiliasMRP();
    var lista = res.lista || [], stats = res.stats || {};
    var resultado = lista.map(function(fam) {
      var s   = stats[fam] || {};
      var pct = (s.sumMax > 0) ? Math.round(s.sumExis / s.sumMax * 100) : 0;
      var ef  = (s.count  > 0) ? Math.round(s.good   / s.count  * 100) : 0;
      var fechaStr = '';
      if (s.lastUpdate) fechaStr = Utilities.formatDate(new Date(s.lastUpdate), tz, 'dd/MM/yyyy HH:mm');
      return { familia: fam, fechaHora: fechaStr, pctExistencia: pct,
               buenos: s.good||0, malos: s.bad||0, efectividad: ef, programaKg: s.progKg||0 };
    });
    return JSON.stringify({ success: true, lista: resultado });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS PARA TableroMetricasHTML (MENU TABLERO METRICAS)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

// 1. OBTENER DATOS TABLERO (CORREGIDO)
function obtenerDatosTablero(fechaInicioStr, fechaFinStr) {
  try {
    Logger.log("📞 obtenerDatosTablero llamada con: " + fechaInicioStr + " - " + fechaFinStr);
    
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var tz = ss.getSpreadsheetTimeZone();
    var sheet = ss.getSheetByName("TABLERO");
    
    if (!sheet) {
      Logger.log("❌ Hoja TABLERO no existe");
      return [];
    }
    
    Logger.log("📅 Consultando desde: " + fechaInicioStr + " hasta: " + fechaFinStr);
    
    var data = sheet.getDataRange().getValues();
    var resultados = [];

    for(var i = 1; i < data.length; i++) {
      var celdaFecha = data[i][8]; // Col I: FECHA_DATA
      if (!(celdaFecha instanceof Date)) continue;

      var fechaCeldaStr = Utilities.formatDate(celdaFecha, tz, "yyyy-MM-dd");
      
      if(fechaCeldaStr >= fechaInicioStr && fechaCeldaStr <= fechaFinStr) {
        // ✅ SANITIZAR TODOS LOS VALORES
        var resultado = data[i][5];
        var dato1 = data[i][6];
        var meta = data[i][11];
        
        // Convertir a valores seguros para JSON
        if(resultado === null || resultado === undefined || isNaN(resultado)) {
          resultado = "";
        } else if(resultado instanceof Date) {
          resultado = Utilities.formatDate(resultado, tz, "yyyy-MM-dd");
        } else {
          resultado = Number(resultado) || String(resultado);
        }
        
        if(dato1 === null || dato1 === undefined) {
          dato1 = "";
        } else if(dato1 instanceof Date) {
          dato1 = Utilities.formatDate(dato1, tz, "yyyy-MM-dd HH:mm:ss");
        } else {
          dato1 = String(dato1);
        }
        
        if(meta === null || meta === undefined || isNaN(meta)) {
          meta = 0;
        } else {
          meta = Number(meta);
        }
        
        resultados.push({
          indicador: String(data[i][4] || ""),
          resultado: resultado,
          dato1: dato1,
          fecha: fechaCeldaStr,
          meta: meta
        });
      }
    }
    
    Logger.log("✅ Encontrados " + resultados.length + " registros");
    
    // ✅ VERIFICAR QUE SE PUEDE SERIALIZAR
    try {
      JSON.stringify(resultados);
      Logger.log("✅ Datos validados como JSON correcto");
    } catch(jsonError) {
      Logger.log("❌ ERROR en JSON.stringify: " + jsonError.toString());
      return [];
    }
    
    return resultados;
    
  } catch(e) {
    Logger.log("❌ ERROR en obtenerDatosTablero: " + e.toString());
    Logger.log("Stack: " + e.stack);
    return [];
  }
}

// 2. GUARDAR DATO (CELDA) - CORREGIDO PARA EVITAR BLOQUEO
function guardarDatoTablero(payload) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var tz = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName("TABLERO");
  
  if(!sheet) {
     sheet = ss.insertSheet("TABLERO");
     sheet.appendRow(["ID", "FECHA_REG", "SEMANA", "MES", "INDICADOR", "RESULTADO", "DATO_1", "DATO_2", "FECHA_DATA", "DATO_3", "DATO_4", "META"]);
  }

  var data = sheet.getDataRange().getValues();
  var fechaTarget = payload.fecha; // Viene como "yyyy-MM-dd"
  var indicador = payload.indicador;
  
  var filaEncontrada = -1;

  // Buscar coincidencia usando comparación de texto (Inmune a desfases horarios)
  for(var i=1; i<data.length; i++) {
      if(data[i][8] instanceof Date) {
         var fRow = Utilities.formatDate(data[i][8], tz, "yyyy-MM-dd");
         if(fRow == fechaTarget && String(data[i][4]) == indicador) {
            filaEncontrada = i + 1;
            break;
         }
      }
  }

  // Preparar fecha del dato (Mediodía para evitar saltos de día)
  var partes = fechaTarget.split("-");
  var fechaObj = new Date(partes[0], partes[1]-1, partes[2], 12, 0, 0);
  
  var semana = Utilities.formatDate(fechaObj, tz, "w");
  var mes = Utilities.formatDate(fechaObj, tz, "MMMM").toUpperCase();

  try {
    if(filaEncontrada > -1) {
        // Actualizar fila existente
        sheet.getRange(filaEncontrada, 6).setValue(payload.resultado);
        sheet.getRange(filaEncontrada, 7).setValue(payload.dato1 || "");
        sheet.getRange(filaEncontrada, 12).setValue(payload.meta || 0);
        sheet.getRange(filaEncontrada, 2).setValue(new Date()); // Fecha de registro actual
    } else {
        // Crear fila nueva
        var id = Utilities.getUuid();
        var row = [
            id, 
            new Date(), 
            semana, 
            mes, 
            indicador, 
            payload.resultado, 
            payload.dato1 || "", 
            "", // DATO_2
            fechaObj, 
            "", "", // DATO_3 y 4
            payload.meta || 0
        ];
        sheet.appendRow(row);
    }
    return "OK";
  } catch(e) {
    return "ERROR: " + e.toString();
  }
}

// =================================================================================
// 3. RESUMEN PRODUCCION (AJUSTE COLATADO, EMPAQUE Y TORNILLERÍA)
// =================================================================================
function obtenerResumenProduccionDetallado(fechaIniStr, fechaFinStr) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetStd = ss.getSheetByName("ESTANDARES");

  var COL_ORDEN_ID = 2;   var COL_MAQUINA  = 4;   var COL_FECHA    = 5;   
  var COL_TURNO    = 6;   var COL_KILOS    = 10;  var COL_HORAS    = 13;  
  var timeZone = ss.getSpreadsheetTimeZone(); 

  var dataOrd = sheetOrd.getDataRange().getValues();
  var mapOrd = {};
  for(var o=1; o<dataOrd.length; o++) {
     mapOrd[String(dataOrd[o][0])] = { tipo: String(dataOrd[o][19]).toUpperCase().trim(), peso: Number(dataOrd[o][18]) || 0 };
  }

  var dataStd = sheetStd.getDataRange().getValues();
  var mapStd = {};
  for(var s=1; s<dataStd.length; s++) {
     var maq = String(dataStd[s][3]).toUpperCase().trim();
     mapStd[maq] = { proc: String(dataStd[s][2]).toUpperCase().trim(), vel: Number(dataStd[s][4]) || 0, grupo: String(dataStd[s][9]).toUpperCase().trim() };
  }

  var dataProd = sheetProd.getDataRange().getValues();
  var estructura = {}; 
  var controlDuplicados = {}; 

  for(var i=1; i<dataProd.length; i++) {
      var rawDate = dataProd[i][COL_FECHA]; 
      if(!rawDate) continue;
      var fechaFilaStr = (rawDate instanceof Date) ? Utilities.formatDate(rawDate, timeZone, "yyyy-MM-dd") : String(rawDate).substring(0, 10); 
      
      if(fechaFilaStr >= fechaIniStr && fechaFilaStr <= fechaFinStr) {
          var idOrd = String(dataProd[i][COL_ORDEN_ID]);
          var infoOrd = mapOrd[idOrd];
          if(!infoOrd) continue; 

          var maquina = String(dataProd[i][COL_MAQUINA]).toUpperCase().trim();
          var turno = String(dataProd[i][COL_TURNO]).replace("T","").trim();
          var kilos = Number(dataProd[i][COL_KILOS]) || 0;
          var horasReales = Number(dataProd[i][COL_HORAS]) || (kilos > 0 ? 7.5 : 0); 

          var std = mapStd[maquina] || { proc: "DESCONOCIDO", vel: 0, grupo: "OTROS" };
          var procesoMaq = std.proc;
          var categoria = "";
          
          // 1. DETERMINAR CATEGORIA ESPECIFICA
          if (maquina === "EMPAQUE CLAVO") categoria = "EMPAQUE CLAVO";
          else if (["BANCO I", "BANCO II", "UGAROLA"].includes(maquina)) categoria = "BARRA";
          else if (["MOL EST 5150 02", "MOL EST 5150 03", "MOL EST 5150 04", "MOL EST 6200 01", "MOL EST 6200 02", "MOL EST 6200 03", "MOL EST 6250 01"].includes(maquina)) categoria = "TORNILLO - PIJAS TOOL";
          else if (maquina === "CEV ERE 0750 01" || maquina === "HAN ERE 1000 01") categoria = "TORNILLO - GRANDES";
          else if(procesoMaq == "ESTAMPADO") categoria = infoOrd.tipo.includes("CONCRETO") ? "CLAVO CONCRETO" : "CLAVO MADERA";
          else if(procesoMaq.includes("ROSCADO") && infoOrd.tipo.includes("VARILLA")) categoria = "VARILLA";
          else if(procesoMaq.includes("FORJA") || procesoMaq.includes("PUNTEADO") || procesoMaq.includes("ROLADO") || procesoMaq.includes("ROSCADO")) {
              if(infoOrd.tipo.includes("TORNILLO") || infoOrd.tipo.includes("PZA") || infoOrd.tipo.includes("G2") || infoOrd.tipo.includes("G5") || infoOrd.tipo.includes("PIJA")) {
                  categoria = "TORNILLO - " + (std.grupo || "GENERAL");
              }
          }
          else if(procesoMaq.includes("TREFILADO")) categoria = "TREFILADO";
          else if(procesoMaq.includes("COLATADO")) categoria = "COLATADO";
          else if(procesoMaq.includes("TEMPLE") || procesoMaq.includes("REVENIDO") || procesoMaq.includes("HORNO")) categoria = "TRATAMIENTOS TÉRMICOS";

          if(categoria === "") continue; 

          // --- LOGICA DE DOBLE ASIGNACION PARA EL GRUPO "TORNILLO" ---
          var categoriasADestino = [categoria];
          if (categoria.startsWith("TORNILLO - ") && 
              categoria !== "TORNILLO - ROSCADO N1" && 
              categoria !== "TORNILLO - ROSCADO TOOL") {
              categoriasADestino.push("TORNILLO");
          }

          // --- CALCULO UNIDADES REALES (AJUSTADO) ---
          // Si es EMPAQUE, TRATAMIENTOS o COLATADO, el valor reportado ya es la unidad final (Kilos o Rollos)
          var unidadesReales = kilos;
          if (categoria !== "EMPAQUE CLAVO" && categoria !== "TRATAMIENTOS TÉRMICOS" && categoria !== "COLATADO") {
              if(std.vel > 0 && infoOrd.peso > 0) {
                  unidadesReales = kilos / infoOrd.peso;
              }
          }

          categoriasADestino.forEach(function(catNombre) {
              if(!estructura[catNombre]) estructura[catNombre] = {};
              if(!estructura[catNombre][fechaFilaStr]) {
                  estructura[catNombre][fechaFilaStr] = { 1:{k:0,uReal:0,uTeo:0}, 2:{k:0,uReal:0,uTeo:0}, 3:{k:0,uReal:0,uTeo:0}, machines: {} };
              }
              
              var dayObj = estructura[catNombre][fechaFilaStr];
              dayObj[turno].k += kilos; 
              dayObj[turno].uReal += unidadesReales; 
          
              if(!dayObj.machines[maquina]) dayObj.machines[maquina] = { 1:{k:0,uR:0,uT:0}, 2:{k:0,uR:0,uT:0}, 3:{k:0,uR:0,uT:0} };
              dayObj.machines[maquina][turno].k += kilos;
              dayObj.machines[maquina][turno].uR += unidadesReales;

              var keyUnica = catNombre + "_" + maquina + "_" + fechaFilaStr + "_" + turno;
              if(!controlDuplicados[keyUnica]) {
                  var metaTeorica = 0;
                  if(std.vel > 0) {
                      // EMPAQUE y TRATAMIENTOS: estándar es KG/HR (estándar * horas)
                      if (catNombre === "EMPAQUE CLAVO" || catNombre === "TRATAMIENTOS TÉRMICOS") {
                          metaTeorica = std.vel * horasReales; 
                      } 
                      // COLATADO y DEMÁS: estándar es por MINUTO (estándar * 60 * horas)
                      // Nota: Colatado entra aquí porque su estándar es 5 ROL/MIN
                      else {
                          metaTeorica = std.vel * 60 * horasReales; 
                      }
                  }
                  dayObj[turno].uTeo += metaTeorica; 
                  dayObj.machines[maquina][turno].uT += metaTeorica; 
                  controlDuplicados[keyUnica] = true;
              }
          });
      }
  }
  return estructura;
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS PARA ValidadorHTML (MENU VALIDADOR)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// 1. VALIDADOR — CARGA DE DATOS COMPLETOS
function obtenerDatosCompletosValidador() {
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd  = ss.getSheetByName(HOJA_ORDENES);
    var sheetPed  = ss.getSheetByName(HOJA_PEDIDOS);
    var sheetInv  = ss.getSheetByName("INVENTARIO_EXTERNO");
    var sheetCod  = ss.getSheetByName("CODIGOS");
    var sheetRutas = ss.getSheetByName("RUTAS");

    if (!sheetOrd || !sheetPed) throw new Error("No se encontraron las hojas ORDENES o PEDIDOS.");

    var dataOrd   = sheetOrd.getDataRange().getValues();
    var dataPed   = sheetPed.getDataRange().getValues();
    var dataInv   = sheetInv   ? sheetInv.getDataRange().getValues()   : [];
    var dataCod   = sheetCod   ? sheetCod.getDataRange().getValues()   : [];
    var dataRutas = sheetRutas ? sheetRutas.getDataRange().getValues() : [];

    // 1. Mapa CODIGO → CODIGO_VENTA
    var mapaCodVenta = {};
    for (var i = 1; i < dataCod.length; i++) {
      if (dataCod[i][0]) mapaCodVenta[String(dataCod[i][0]).trim()] = String(dataCod[i][5]).trim();
    }

    // 2. Mapa Inventario Externo
    var mapaInv = {};
    if (dataInv.length > 0) {
      var hInv = dataInv[0];
      var iiCod = 0, iiExt = 1, iiMin = 2, iiMax = 3, iiBack = -1;
      for (var h = 0; h < hInv.length; h++) {
        var hName = String(hInv[h]).toUpperCase().trim();
        if (hName === "BACKORDER")  iiBack = h;
        if (hName === "CODIGO")     iiCod  = h;
        if (hName === "EXISTENCIA") iiExt  = h;
        if (hName === "MINIMO")     iiMin  = h;
        if (hName === "MAXIMO")     iiMax  = h;
      }
      for (var i = 1; i < dataInv.length; i++) {
        var c = String(dataInv[i][iiCod]).trim();
        mapaInv[c] = {
          ext:  dataInv[i][iiExt],
          min:  dataInv[i][iiMin],
          max:  dataInv[i][iiMax],
          back: iiBack > -1 ? Number(dataInv[i][iiBack]) || 0 : 0
        };
      }
    }

    // 3. Mapa Códigos con Family
    var mapaCodigos = {};
    for (var i = 1; i < dataCod.length; i++) {
      if (dataCod[i][0]) {
        mapaCodigos[String(dataCod[i][0]).trim()] = {
          family: String(dataCod[i][8] || "").trim()
        };
      }
    }

    // 4. Mapa Rutas con Diámetro y Longitud
    var mapaRutas = {};
    for (var i = 1; i < dataRutas.length; i++) {
      if (dataRutas[i][1]) {
        mapaRutas[String(dataRutas[i][1]).trim()] = {
          diametro: String(dataRutas[i][7] || "").trim(),
          longitud: String(dataRutas[i][8] || "").trim()
        };
      }
    }

    // 5. Mapa de Pedidos
    var mapaPed  = {};
    var timeZone = Session.getScriptTimeZone();

    for (var p = 1; p < dataPed.length; p++) {
      var rowP = dataPed[p];
      if (!rowP[1]) continue;

      var codProd     = String(rowP[3]).trim();
      var codBusqueda = mapaCodVenta[codProd] || codProd;
      var inv         = mapaInv[codBusqueda] || { ext: 0, min: 0, max: 0 };
      var keyPed      = String(rowP[1]).trim() + "_" + Math.trunc(rowP[5]);

      var fechaEntVal = rowP[11] instanceof Date
        ? Utilities.formatDate(rowP[11], timeZone, "dd/MM/yyyy")
        : String(rowP[11] || "").trim();

      var fechaPedVal = rowP[2] instanceof Date
        ? Utilities.formatDate(rowP[2], timeZone, "dd/MM/yy")
        : String(rowP[2] || "").trim();

      mapaPed[keyPed] = {
        fila:     p + 1,
        nombre:   String(rowP[1]).trim(),
        partida:  rowP[5],
        desc:     String(rowP[4]).trim(),
        cod:      codProd,
        cant:     Number(rowP[6]),
        uni:      String(rowP[7]).toUpperCase().trim(),
        est:      String(rowP[8]).toUpperCase().trim(),
        inv:      inv,
        fechaEnt: fechaEntVal,
        fechaPed: fechaPedVal,
        family:   (mapaCodigos[codProd] || {}).family   || "",
        diametro: (mapaRutas[codProd]   || {}).diametro || "",
        longitud: (mapaRutas[codProd]   || {}).longitud || ""
      };
    }

    // 6. Agrupar Órdenes
    var grupos = {};
    for (var i = 1; i < dataOrd.length; i++) {
      var rowO      = dataOrd[i];
      var pedNombre = String(rowO[1]).trim();
      var pedPart   = Math.trunc(rowO[2]);
      var keyPed    = pedNombre + "_" + pedPart;
      var keyOrd    = rowO[4] + "." + String(rowO[5]).padStart(4, '0');

      if (!grupos[keyOrd]) {
        grupos[keyOrd] = {
          key:      keyOrd,
          keyPed:   keyPed,
          rows:     [],
          maxSec:   -1,
          lastProc: null,
          estOrd:   String(rowO[15]).toUpperCase()
        };
      }

      var obj = {
        idx:     i,
        idOrden: String(rowO[0]),
        sec:     Number(rowO[10]),
        proc:    rowO[11],
        sol:     Number(rowO[13]),
        prod:    Number(rowO[14]),
        est:     String(rowO[15]).toUpperCase(),
        peso:    Number(rowO[18]),
        tipo:    String(rowO[19]),
        long:    String(rowO[21]),
        uni:     String(rowO[9]).toUpperCase(),
        cantOrd: Number(rowO[8]),
        cuerpo:  String(rowO[23] || "").trim(),
        cuerda:  String(rowO[22] || "").trim(),
        acero:   String(rowO[24] || "").trim()
      };
      grupos[keyOrd].rows.push(obj);
      if (obj.sec > grupos[keyOrd].maxSec) {
        grupos[keyOrd].maxSec   = obj.sec;
        grupos[keyOrd].lastProc = obj;
      }
    }

    // Enriquecer mapaPed con tipo/cuerpo/cuerda/acero de la primera orden
    for (var k in grupos) {
      var g = grupos[k];
      if (!mapaPed[g.keyPed]) continue;
      if (mapaPed[g.keyPed].cuerpo !== undefined) continue; // ya enriquecido
      var firstRow = g.rows[0] || {};
      mapaPed[g.keyPed].tipo   = String(firstRow.tipo   || "").trim();
      mapaPed[g.keyPed].cuerpo = String(firstRow.cuerpo || "").trim();
      mapaPed[g.keyPed].cuerda = String(firstRow.cuerda || "").trim();
      mapaPed[g.keyPed].acero  = String(firstRow.acero  || "").trim();
    }

    // 7. Candidatos para cierre

    // CRITERIO A: Último proceso ≥ 95% del pedido
    var candidatos  = [];
    var pedidoInfo  = {};

    for (var k in grupos) {
      var g   = grupos[k];
      var lp  = g.lastProc;
      var ped = mapaPed[g.keyPed];
      if (!ped || !lp || lp.est === "CANCELADO") continue;

      var divisor = 1;
      if (lp.uni.includes("PZA") || lp.uni.includes("CTO")) {
        var longNum = lp.tipo.includes("VARILLA")
          ? (parseFloat((lp.long.match(/[\d.]+/) || ["1"])[0]))
          : 1;
        divisor = (lp.peso * longNum) || 1;
      }
      if (!pedidoInfo[g.keyPed]) pedidoInfo[g.keyPed] = { totalProdLastProc: 0, divisor: divisor };
      pedidoInfo[g.keyPed].totalProdLastProc += lp.prod;
    }

    for (var keyPed in pedidoInfo) {
      var ped = mapaPed[keyPed];
      if (!ped || ped.est.match(/TERMINADO|CANCELADO|CERRADO/)) continue;
      var info   = pedidoInfo[keyPed];
      var avance = ped.cant > 0 ? (info.totalProdLastProc / (info.divisor || 1) / ped.cant) * 100 : 0;
      if (avance >= 95) candidatos.push(keyPed);
    }

    // CRITERIO B: Todas las órdenes cerradas
    var keysYaAgregadas   = new Set(candidatos);
    var pedidosConOrdenes = {};

    for (var k in grupos) {
      var g    = grupos[k];
      var keyP = g.keyPed;
      var ped  = mapaPed[keyP];
      if (!ped || keysYaAgregadas.has(keyP) || ped.est.match(/TERMINADO|CANCELADO|CERRADO/)) continue;

      if (!pedidosConOrdenes[keyP]) pedidosConOrdenes[keyP] = { todasCerradas: true, tieneOrdenes: false };
      pedidosConOrdenes[keyP].tieneOrdenes = true;

      for (var r = 0; r < g.rows.length; r++) {
        var estProc = g.rows[r].est;
        if (estProc !== "TERMINADO" && estProc !== "CANCELADO" && estProc !== "CERRADA") {
          pedidosConOrdenes[keyP].todasCerradas = false;
          break;
        }
      }
    }

    for (var keyP in pedidosConOrdenes) {
      if (pedidosConOrdenes[keyP].tieneOrdenes && pedidosConOrdenes[keyP].todasCerradas) {
        candidatos.push(keyP);
      }
    }

    return JSON.stringify({
      pedidos:    mapaPed,
      ordenes:    grupos,
      candidatos: candidatos,
      familias:   _extraerFamilias(mapaCodigos),
      diametros:  _extraerDiametros(mapaRutas),
      longitudes: _extraerLongitudes(mapaRutas)
    });

  } catch (e) {
    return JSON.stringify({ error: e.toString(), candidatos: [] });
  }
}

function _extraerFamilias(mapaCodigos) {
  var s = new Set();
  for (var cod in mapaCodigos) { if (mapaCodigos[cod].family) s.add(mapaCodigos[cod].family); }
  return Array.from(s).sort();
}
function _extraerDiametros(mapaRutas) {
  var s = new Set();
  for (var cod in mapaRutas) { if (mapaRutas[cod].diametro) s.add(mapaRutas[cod].diametro); }
  return Array.from(s).sort();
}
function _extraerLongitudes(mapaRutas) {
  var s = new Set();
  for (var cod in mapaRutas) { if (mapaRutas[cod].longitud) s.add(mapaRutas[cod].longitud); }
  return Array.from(s).sort();
}

// 2. VALIDADOR — GUARDAR CAMBIOS
function aplicarCambiosValidador(listaCambios) {
  var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sPed = ss.getSheetByName(HOJA_PEDIDOS);
  var sOrd = ss.getSheetByName(HOJA_ORDENES);

  try {
    listaCambios.forEach(function(c) {
      if (c.tipo === "PEDIDO") {
        if (c.nuevaCant !== undefined) sPed.getRange(c.fila, 7).setValue(c.nuevaCant);
        if (c.nuevoEstado)             sPed.getRange(c.fila, 9).setValue(c.nuevoEstado);
      }
      if (c.tipo === "ORDEN") {
        if (c.nuevaCant !== undefined) {
          c.procesos.forEach(function(p) {
            sOrd.getRange(p.idx + 1, 9).setValue(c.nuevaCant);
            sOrd.getRange(p.idx + 1, 14).setValue(p.nuevoSol);
          });
        }
        if (c.nuevoEstado) {
          c.procesos.forEach(function(p) {
            sOrd.getRange(p.idx + 1, 16).setValue(c.nuevoEstado);
          });
        }
      }
      if (c.tipo === "PROCESO") {
        sOrd.getRange(c.procIdx + 1, 16).setValue(c.nuevoEstado);
      }
    });
    return "✅ Sincronización Exitosa.";
  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}

// 3. VALIDADOR — BÚSQUEDAS ESPECIALES
function buscarPedidos80Porciento() { return _buscarPorAvance(80); }
function buscarPedidos60Porciento() { return _buscarPorAvance(60); }

function _buscarPorAvance(umbral) {
  try {
    var resultado = JSON.parse(obtenerDatosCompletosValidador());
    if (resultado.error) return JSON.stringify(resultado);

    var candidatos    = [];
    var keysAgregadas = new Set();

    for (var k in resultado.ordenes) {
      var g   = resultado.ordenes[k];
      var lp  = g.lastProc;
      var ped = resultado.pedidos[g.keyPed];
      if (!ped || !lp || keysAgregadas.has(g.keyPed)) continue;
      if (ped.est.match(/TERMINADO|CANCELADO|CERRADO/)) continue;

      var divisor = 1;
      if (lp.uni.includes("PZA") || lp.uni.includes("CTO")) {
        var longNum = lp.tipo.includes("VARILLA") ? (parseFloat(lp.long.match(/[\d.]+/)) || 1) : 1;
        divisor = lp.peso * longNum;
      }
      var avance = ped.cant > 0 ? (lp.prod / (divisor || 1) / ped.cant) * 100 : 0;

      if (avance >= umbral) {
        candidatos.push(g.keyPed);
        keysAgregadas.add(g.keyPed);
      }
    }

    resultado.candidatos = candidatos;
    return JSON.stringify(resultado);
  } catch (e) {
    return JSON.stringify({ error: e.toString(), candidatos: [] });
  }
}

function buscarPedidosConSaldos() {
  try {
    var resultado = JSON.parse(obtenerDatosCompletosValidador());
    if (resultado.error) return JSON.stringify(resultado);

    var candidatos    = [];
    var keysAgregadas = new Set();
    var ahora         = new Date().getTime();
    var quinceDiasMs  = 15 * 24 * 60 * 60 * 1000;

    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd = ss.getSheetByName("PRODUCCION");
    var dataProd  = sheetProd.getDataRange().getValues();
    var hProd     = dataProd[0].map(function(h) { return String(h).toUpperCase().trim(); });
    var ipOrden   = hProd.indexOf("ID_ORDEN"); if (ipOrden < 0) ipOrden = hProd.indexOf("ORDEN");
    var ipFecha   = hProd.indexOf("FECHA");

    var ultimaFechaProd = {};
    for (var p = 1; p < dataProd.length; p++) {
      var idOrd = String(dataProd[p][ipOrden]).trim();
      var fRaw  = dataProd[p][ipFecha];
      if (fRaw instanceof Date) {
        var ts = fRaw.getTime();
        if (!ultimaFechaProd[idOrd] || ts > ultimaFechaProd[idOrd]) ultimaFechaProd[idOrd] = ts;
      }
    }

    for (var k in resultado.ordenes) {
      var g   = resultado.ordenes[k];
      var ped = resultado.pedidos[g.keyPed];
      if (!ped || keysAgregadas.has(g.keyPed)) continue;
      if (ped.est.match(/TERMINADO|CANCELADO|CERRADO/)) continue;
      if (g.estOrd === "CANCELADO") continue;

      var rows = g.rows.slice().sort(function(a, b) { return a.sec - b.sec; });
      if (rows.length < 2) continue;

      var avancePrimer = rows[0].sol > 0 ? (rows[0].prod / rows[0].sol * 100) : 0;
      if (avancePrimer < 90) continue;

      var ultimaFechaOrd = 0;
      for (var r = 0; r < rows.length; r++) {
        var fProc = ultimaFechaProd[rows[r].idOrden] || 0;
        if (fProc > ultimaFechaOrd) ultimaFechaOrd = fProc;
      }
      if (ultimaFechaOrd > 0 && (ahora - ultimaFechaOrd) < quinceDiasMs) continue;

      candidatos.push(g.keyPed);
      keysAgregadas.add(g.keyPed);
    }

    resultado.candidatos = candidatos;
    return JSON.stringify(resultado);
  } catch (e) {
    return JSON.stringify({ error: e.toString(), candidatos: [] });
  }
}

// 4. VALIDADOR — MODAL TRANSFERENCIA DE PRODUCCIÓN
function obtenerProduccionDeOrden(idsOrdenJson) {
  try {
    var ss         = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd  = ss.getSheetByName("PRODUCCION");
    var sheetLotes = ss.getSheetByName("LOTES");
    var sheetOrd   = ss.getSheetByName(HOJA_ORDENES);
    var tz         = Session.getScriptTimeZone();

    var ids;
    try { ids = JSON.parse(idsOrdenJson); } catch(e) { ids = [idsOrdenJson]; }
    if (!Array.isArray(ids)) ids = [ids];
    var setIds = {};
    ids.forEach(function(id) { setIds[String(id)] = true; });

    var dataProd  = sheetProd.getDataRange().getValues();
    var dataLotes = sheetLotes.getDataRange().getValues();
    var dataOrd   = sheetOrd.getDataRange().getValues();

    var mapaOrdenes = {};
    for (var i = 1; i < dataOrd.length; i++) {
      var idO = String(dataOrd[i][0]);
      if (!setIds[idO]) continue;
      var serieO = String(dataOrd[i][4]).trim();
      var numO   = ("0000" + Number(dataOrd[i][5])).slice(-4);
      mapaOrdenes[idO] = {
        proceso: String(dataOrd[i][11]).trim(),
        codigo:  String(dataOrd[i][6]).trim(),
        prefijo: serieO + "." + numO + "."
      };
    }

    var registros = [];
    for (var p = 1; p < dataProd.length; p++) {
      var row   = dataProd[p];
      var idOrd = String(row[2]);
      if (!setIds[idOrd]) continue;

      var fechaStr = row[5] instanceof Date
        ? Utilities.formatDate(row[5], tz, "yyyy-MM-dd")
        : String(row[5] || "").trim();

      var oInfo = mapaOrdenes[idOrd] || {};
      registros.push({
        id:       String(row[0]),
        fila:     p + 1,
        idOrden:  idOrd,
        fecha:    fechaStr,
        turno:    String(row[6]),
        lote:     String(row[3]),
        maquina:  String(row[4]),
        pesoI:    Number(row[8])  || 0,
        pesoF:    Number(row[9])  || 0,
        pesoTina: Number(row[22]) || 0,
        producido: Number(row[10]) || 0,
        proceso:  oInfo.proceso || ""
      });
    }

    var lotesDisponibles = [];
    var setLotes = {};
    for (var l = 1; l < dataLotes.length; l++) {
      var nomLote = String(dataLotes[l][4]).trim();
      for (var idOx in mapaOrdenes) {
        var pref = mapaOrdenes[idOx].prefijo;
        if (pref && nomLote.indexOf(pref) === 0 && !setLotes[nomLote]) {
          lotesDisponibles.push(nomLote);
          setLotes[nomLote] = true;
        }
      }
    }

    var sheetEst = ss.getSheetByName("ESTANDARES");
    var maquinasDelProceso = {};
    if (sheetEst) {
      var dataEst = sheetEst.getDataRange().getValues();
      for (var e = 1; e < dataEst.length; e++) {
        var procE = String(dataEst[e][2]).trim();
        var maqE  = String(dataEst[e][3]).trim();
        if (!procE || !maqE) continue;
        if (!maquinasDelProceso[procE]) maquinasDelProceso[procE] = [];
        if (maquinasDelProceso[procE].indexOf(maqE) < 0) maquinasDelProceso[procE].push(maqE);
      }
    }

    var descripcionOrden = "";
    var nombreOrden      = "";
    for (var i = 1; i < dataOrd.length; i++) {
      if (setIds[String(dataOrd[i][0])]) {
        descripcionOrden = String(dataOrd[i][7] || "").trim();
        nombreOrden = String(dataOrd[i][4]).trim() + "." + ("0000" + Number(dataOrd[i][5])).slice(-4);
        break;
      }
    }

    return JSON.stringify({
      success:            true,
      registros:          registros,
      lotes:              lotesDisponibles,
      maquinasDelProceso: maquinasDelProceso,
      descripcionOrden:   descripcionOrden,
      nombreOrden:        nombreOrden
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString(), registros: [], lotes: [] });
  }
}

function guardarEdicionProduccion(payload) {
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd = ss.getSheetByName("PRODUCCION");

    var mapaColumnas = {
      fecha:    6,
      turno:    7,
      lote:     4,
      maquina:  5,
      pesoI:    9,
      pesoF:    10,
      pesoTina: 23,
      producido: 11
    };

    var colNum = mapaColumnas[payload.campo];
    if (!colNum) return JSON.stringify({ success: false, msg: "Campo no válido: " + payload.campo });

    var dataProd  = sheetProd.getDataRange().getValues();
    var filaVerif = payload.fila - 1;
    if (filaVerif < 1 || filaVerif >= dataProd.length)
      return JSON.stringify({ success: false, msg: "Fila fuera de rango" });
    if (String(dataProd[filaVerif][0]) !== String(payload.id))
      return JSON.stringify({ success: false, msg: "ID no coincide con la fila" });

    var valor = payload.valor;
    if (payload.campo === "fecha") {
      valor = new Date(valor + "T12:00:00");
    } else if (["pesoI", "pesoF", "pesoTina", "producido"].indexOf(payload.campo) >= 0) {
      valor = Number(valor) || 0;
    }

    sheetProd.getRange(payload.fila, colNum).setValue(valor);

    if (payload.campo === "producido") {
      recalcularEstadoOrdenMaestro(String(dataProd[filaVerif][2]));
    }

    return JSON.stringify({ success: true, msg: "Guardado" });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function transferirProduccion(payload) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return JSON.stringify({ success: false, msg: "Servidor ocupado." });

  try {
    var ss         = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd  = ss.getSheetByName("PRODUCCION");
    var sheetLotes = ss.getSheetByName("LOTES");
    var sheetOrd   = ss.getSheetByName(HOJA_ORDENES);

    var dataProd = sheetProd.getDataRange().getValues();
    var filaReg  = payload.filaRegistro - 1;

    if (filaReg < 1 || filaReg >= dataProd.length) {
      lock.releaseLock();
      return JSON.stringify({ success: false, msg: "Fila fuera de rango" });
    }
    if (String(dataProd[filaReg][0]) !== String(payload.idRegistro)) {
      lock.releaseLock();
      return JSON.stringify({ success: false, msg: "ID no coincide con la fila" });
    }

    var rowOrig      = dataProd[filaReg].slice();
    var idOrdenOrig  = String(rowOrig[2]);
    var loteName     = String(rowOrig[3]);
    var cantOrig     = Number(rowOrig[10]);

    var dataOrd     = sheetOrd.getDataRange().getValues();
    var codigoOrig  = "";
    var procesoOrig = "";
    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][0]) === idOrdenOrig) {
        codigoOrig  = String(dataOrd[i][6]).trim();
        procesoOrig = String(dataOrd[i][11]).trim();
        break;
      }
    }

    var candidatas = [];
    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][0]) === idOrdenOrig) continue;
      var estOrd = String(dataOrd[i][15]).toUpperCase();
      if (estOrd === "CANCELADO" || estOrd === "TERMINADO" || estOrd === "SOBREPRODUCCION") continue;
      if (String(dataOrd[i][6]).trim()  !== codigoOrig)  continue;
      if (String(dataOrd[i][11]).trim() !== procesoOrig) continue;
      var serieI  = String(dataOrd[i][4]).trim();
      var numOrdI = ("0000" + Number(dataOrd[i][5])).slice(-4);
      candidatas.push({
        id:     String(dataOrd[i][0]),
        nombre: serieI + "." + numOrdI,
        sol:    Number(dataOrd[i][13]),
        prod:   Number(dataOrd[i][14])
      });
    }

    if (payload.modo === "CONSULTAR") {
      lock.releaseLock();
      return JSON.stringify({ success: true, candidatas: candidatas });
    }

    var idOrdenDestino = String(payload.idOrdenDestino);

    if (payload.modo === "CREAR_DOS") {
      var cantTransferir = Number(payload.cantTransferir);
      if (cantTransferir <= 0 || cantTransferir >= cantOrig) {
        lock.releaseLock();
        return JSON.stringify({ success: false, msg: "Cantidad inválida (debe ser > 0 y < " + cantOrig + ")" });
      }
      sheetProd.getRange(payload.filaRegistro, 11).setValue(cantOrig - cantTransferir);
      var nuevoReg = rowOrig.slice();
      nuevoReg[0]  = Utilities.getUuid();
      nuevoReg[2]  = idOrdenDestino;
      nuevoReg[10] = cantTransferir;
      sheetProd.appendRow(nuevoReg);

    } else if (payload.modo === "MODIFICAR") {
      sheetProd.getRange(payload.filaRegistro, 3).setValue(idOrdenDestino);
    }

    var dataLotes = sheetLotes.getDataRange().getValues();
    for (var l = 1; l < dataLotes.length; l++) {
      if (String(dataLotes[l][4]).trim() === loteName) {
        sheetLotes.getRange(l + 1, 3).setValue(idOrdenDestino);
        break;
      }
    }

    recalcularEstadoOrdenMaestro(idOrdenOrig);
    recalcularEstadoOrdenMaestro(idOrdenDestino);
    SpreadsheetApp.flush();

    var dataOrdPost   = sheetOrd.getDataRange().getValues();
    var estadoOrigen  = { idOrden: idOrdenOrig,   nuevoEstado: "", keyOrden: "" };
    var estadoDestino = { idOrden: idOrdenDestino, nuevoEstado: "", keyOrden: "" };
    var pedidosActualizar = {};

    for (var i = 1; i < dataOrdPost.length; i++) {
      var idI  = String(dataOrdPost[i][0]);
      var estI = String(dataOrdPost[i][15]).toUpperCase();
      var keyI = String(dataOrdPost[i][4]).trim() + "." + ("0000" + Number(dataOrdPost[i][5])).slice(-4);
      if (idI === idOrdenOrig)    { estadoOrigen.nuevoEstado  = estI; estadoOrigen.keyOrden  = keyI; }
      if (idI === idOrdenDestino) { estadoDestino.nuevoEstado = estI; estadoDestino.keyOrden = keyI; }
      if (idI === idOrdenOrig || idI === idOrdenDestino) {
        var pedNom  = String(dataOrdPost[i][1]).trim();
        var pedPart = Math.trunc(Number(dataOrdPost[i][2]));
        pedidosActualizar[pedNom + "_" + pedPart] = { nom: pedNom, part: pedPart };
      }
    }

    var sheetPed = ss.getSheetByName(HOJA_PEDIDOS);
    var dataPed  = sheetPed.getDataRange().getValues();
    for (var keyP in pedidosActualizar) {
      var pa = pedidosActualizar[keyP];
      var todasTerminadas = true;
      var algunaProd      = false;
      for (var i = 1; i < dataOrdPost.length; i++) {
        if (String(dataOrdPost[i][1]).trim() !== pa.nom) continue;
        if (Math.trunc(Number(dataOrdPost[i][2])) !== pa.part) continue;
        var estOi = String(dataOrdPost[i][15]).toUpperCase();
        if (estOi !== "TERMINADO" && estOi !== "CANCELADO") todasTerminadas = false;
        if (Number(dataOrdPost[i][14]) > 0) algunaProd = true;
      }
      var nuevoPedEst = todasTerminadas ? "TERMINADO" : (algunaProd ? "EN PROCESO" : "ABIERTO");
      for (var p = 1; p < dataPed.length; p++) {
        if (String(dataPed[p][1]).trim() === pa.nom && Math.trunc(Number(dataPed[p][5])) === pa.part) {
          sheetPed.getRange(p + 1, 9).setValue(nuevoPedEst);
          break;
        }
      }
    }
    SpreadsheetApp.flush();

    lock.releaseLock();
    return JSON.stringify({
      success:       true,
      msg:           "Transferencia completada",
      candidatas:    candidatas,
      estadoOrigen:  estadoOrigen,
      estadoDestino: estadoDestino
    });

  } catch (e) {
    try { lock.releaseLock(); } catch(ex) {}
    return JSON.stringify({ success: false, msg: e.toString(), candidatas: [] });
  }
}

// =================================================================================
// STUB — recalcularEstadoOrdenMaestro
// Si ya tienes esta función en otro GS vinculado al mismo proyecto, borra este stub.
// Si NO la tienes, déjalo y adáptalo a tu lógica.
// =================================================================================
function recalcularEstadoOrdenMaestro(idOrdenInput) {
  // Esta función debe recalcular el estado de una orden según su producción.
  // Implementación mínima — reemplaza con tu lógica real si ya existe.
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd = ss.getSheetByName(HOJA_ORDENES);
    var sheetProd = ss.getSheetByName("PRODUCCION");
    if (!sheetOrd || !sheetProd) return;

    var dataOrd  = sheetOrd.getDataRange().getValues();
    var dataProd = sheetProd.getDataRange().getValues();

    // Buscar la(s) fila(s) de la orden
    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][0]) !== String(idOrdenInput)) continue;
      var sol  = Number(dataOrd[i][13]);
      var idO  = String(dataOrd[i][0]);

      // Sumar producción total de esta orden
      var totalProd = 0;
      for (var p = 1; p < dataProd.length; p++) {
        if (String(dataProd[p][2]) === idO) totalProd += Number(dataProd[p][10]) || 0;
      }

      // Calcular nuevo estado
      var nuevoEstado;
      if (totalProd <= 0)           nuevoEstado = "ABIERTA";
      else if (totalProd >= sol)    nuevoEstado = "TERMINADO";
      else                          nuevoEstado = "EN PROCESO";

      sheetOrd.getRange(i + 1, 16).setValue(nuevoEstado);
      return nuevoEstado;
    }
  } catch(e) {
    // Silencioso — no interrumpir el flujo principal
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA TrackingHTML (MENU TRACKING)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function rastreoBuscar(q) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shPed = ss.getSheetByName('PEDIDOS');
    var shOrd = ss.getSheetByName('ORDENES');
    if (!shPed) return JSON.stringify({ pedidos: [], _err: 'Hoja PEDIDOS no encontrada' });
    if (!shOrd) return JSON.stringify({ pedidos: [], _err: 'Hoja ORDENES no encontrada' });
    var dataPed = shPed.getDataRange().getValues();
    var dataOrd = shOrd.getDataRange().getValues();
    var qU = String(q || '').toUpperCase().trim();
    if (!qU) return JSON.stringify({ pedidos: [] });

    var tz = ss.getSpreadsheetTimeZone();

    // Mapa: codigo → tipo
    var mapaCodigoTipo = {};
    for (var o = 1; o < dataOrd.length; o++) {
      var cod = String(dataOrd[o][6]).trim();
      var tip = String(dataOrd[o][19]).trim();
      if (cod && tip && !mapaCodigoTipo[cod]) mapaCodigoTipo[cod] = tip;
    }

    // Mapa: pedido → { producidoKg, peso, tipo, longitud } — último proceso (maxSec) de ORDENES
    // Col B=idx1:pedido, G=idx6:codigo, K=idx10:sec, O=idx14:producido, S=idx18:peso, T=idx19:tipo, V=idx21:longitud
    var mapaPedidoProducido = {};
    for (var o2 = 1; o2 < dataOrd.length; o2++) {
      var rO = dataOrd[o2];
      var ped2 = String(rO[1] || '').trim();
      var sec2 = Number(rO[10]) || 0;
      var prod2 = Number(rO[14]) || 0;
      var peso2 = parseFloat(rO[18]) || 0;
      var tipo2 = String(rO[19] || '');
      var long2 = parseFloat(String(rO[21] || '').replace(/[^\d.]/g,'')) || 0;
      if (!ped2) continue;
      if (!mapaPedidoProducido[ped2] || sec2 > mapaPedidoProducido[ped2].sec) {
        mapaPedidoProducido[ped2] = { producidoKg: prod2, peso: peso2, tipo: tipo2, longitud: long2, sec: sec2 };
      }
    }

    // Mapa: pedido → enviadoKg — suma de col J=idx9 de ENVIADO
    var shEnv2 = ss.getSheetByName('ENVIADO');
    var mapaEnviado = {};
    if (shEnv2) {
      var dataEnv2 = shEnv2.getDataRange().getValues();
      for (var e2 = 1; e2 < dataEnv2.length; e2++) {
        var pedEnv = String(dataEnv2[e2][5] || '').trim(); // col F=idx5: pedido
        var kgEnv  = Number(dataEnv2[e2][9]) || 0;        // col J=idx9: kilos
        if (!pedEnv) continue;
        mapaEnviado[pedEnv] = (mapaEnviado[pedEnv] || 0) + kgEnv;
      }
    }

    var resultados = [];
    for (var i = 1; i < dataPed.length; i++) {
      var row = dataPed[i];
      var pedido = String(row[1] || '').toUpperCase();
      var codigo = String(row[3] || '').toUpperCase();
      var desc   = String(row[4] || '').toUpperCase();
      var estado = String(row[8] || '').toUpperCase();
      if (pedido.includes(qU) || codigo.includes(qU) || desc.includes(qU) || estado.includes(qU)) {
        var tipoReal = mapaCodigoTipo[String(row[3]).trim()] || '';
        // Serializar fechas a string para evitar problemas de serialización GAS
        var fechaStr = '';
        if (row[2] instanceof Date) {
          fechaStr = Utilities.formatDate(row[2], tz, 'yyyy-MM-dd');
        } else {
          fechaStr = String(row[2] || '');
        }
        var fechaEntStr = '';
        if (row[11] instanceof Date) {
          fechaEntStr = Utilities.formatDate(row[11], tz, 'yyyy-MM-dd');
        } else {
          fechaEntStr = String(row[11] || '');
        }
        var pedNom = String(row[1] || '');
        var infoOrd = mapaPedidoProducido[pedNom] || { producidoKg:0, peso:0, tipo:'', longitud:0 };
        resultados.push({
          id: String(row[0] || ''),
          pedido: pedNom,
          fecha: fechaStr,
          codigo: String(row[3] || ''),
          descripcion: String(row[4] || ''),
          partida: String(row[5] || ''),
          cantidad: String(row[6] || ''),
          unidad: String(row[7] || ''),
          estado: String(row[8] || ''),
          fechaEntrega: fechaEntStr,
          svgIcono: obtenerIconoSVG(tipoReal),
          producidoKg: infoOrd.producidoKg,
          enviadoKg:   mapaEnviado[pedNom] || 0,
          peso:        infoOrd.peso,
          tipo:        infoOrd.tipo,
          longitud:    infoOrd.longitud
        });
      }
    }
    return JSON.stringify({ pedidos: resultados });
  } catch(e) {
    return JSON.stringify({ pedidos: [], _err: e.message + ' | stack: ' + (e.stack||'').substring(0, 200) });
  }
}

function rastreoDetalle(pedidoNom, partida) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOrd  = ss.getSheetByName('ORDENES');
  var shProd = ss.getSheetByName('PRODUCCION');
  var shEnv  = ss.getSheetByName('ENVIADO');
  var dataOrd  = shOrd.getDataRange().getValues();
  var dataProd = shProd.getDataRange().getValues();
  var dataEnv  = shEnv ? shEnv.getDataRange().getValues() : [];
  var tz = ss.getSpreadsheetTimeZone();

  // Ordenes del pedido
  var ordenes = [];
  var maxSec  = {};
  var ordenIds = {};
  for (var i = 1; i < dataOrd.length; i++) {
    var r = dataOrd[i];
    if (String(r[1]).trim() !== String(pedidoNom).trim()) continue;
    var key = r[4] + '.' + String(r[5]).padStart(4,'0');
    if (!maxSec[key] || Number(r[10]) > maxSec[key]) maxSec[key] = Number(r[10]);
    ordenIds[String(r[0])] = true;
    ordenes.push({
      id: r[0], pedido: r[1], partida: r[2],
      serie: r[4], orden: r[5], codigo: r[6],
      descripcion: r[7], cantidad: r[8],
      unidad: r[9], sec: r[10], proceso: r[11],
      maquina: r[12], solicitado: r[13], producido: r[14],
      estado: r[15], pt: r[16],
      peso: parseFloat(r[18]) || 0,
      tipo: String(r[19] || ''),
      longitud: parseFloat(String(r[21]).replace(/[^\d.]/g,'')) || 0,
      fechaInicio: r[27], fechaFin: r[28],
      produccion: [],
      _dbg: 'r18=' + r[18] + ' r19=' + r[19] + ' r21=' + r[21]
    });
  }
  ordenes.forEach(function(o) {
    var key = o.serie + '.' + String(o.orden).padStart(4,'0');
    o.maxSec = maxSec[key] || o.sec;
  });

  // Produccion por orden
  for (var p = 1; p < dataProd.length; p++) {
    var pr = dataProd[p];
    if (!ordenIds[String(pr[2])]) continue;
    var idxOrd = ordenes.findIndex(function(o){return String(o.id)===String(pr[2]);});
    if (idxOrd >= 0) {
      ordenes[idxOrd].produccion.push({
        fecha: pr[5], turno: pr[6],
        lote: pr[3]||'—',
        operador: pr[17]||pr[7],
        producido: pr[10],
        comentario: pr[15],
        sello: pr[14]||''
      });
    }
  }

  // Enviados del pedido
  var enviados = [];
  for (var e = 1; e < dataEnv.length; e++) {
    var ev = dataEnv[e];
    if (String(ev[5]).trim() === String(pedidoNom).trim()) {
      enviados.push({
        fecha: ev[2], remision: ev[3],
        kilos: ev[9], codigo: ev[6],
        descripcion: ev[7], piezas: ev[10],
        envio: String(ev[12] || ''),
        url:   String(ev[16] || '')
      });
    }
  }

  // Serializar fechas en ordenes para evitar problemas de transferencia GAS→cliente
  ordenes = ordenes.map(function(o) {
    return {
      id: String(o.id||''), pedido: String(o.pedido||''), partida: String(o.partida||''),
      serie: String(o.serie||''), orden: String(o.orden||''), codigo: String(o.codigo||''),
      descripcion: String(o.descripcion||''), cantidad: String(o.cantidad||''),
      unidad: String(o.unidad||''), sec: Number(o.sec)||0, proceso: String(o.proceso||''),
      maquina: String(o.maquina||''), solicitado: Number(o.solicitado)||0, producido: Number(o.producido)||0,
      estado: String(o.estado||''), pt: String(o.pt||''),
      peso: Number(o.peso)||0,
      tipo: String(o.tipo||''),
      longitud: Number(o.longitud)||0,
      fechaInicio: o.fechaInicio instanceof Date ? Utilities.formatDate(o.fechaInicio, tz, 'yyyy-MM-dd') : String(o.fechaInicio||''),
      fechaFin:    o.fechaFin    instanceof Date ? Utilities.formatDate(o.fechaFin,    tz, 'yyyy-MM-dd') : String(o.fechaFin||''),
      maxSec: Number(o.maxSec)||0,
      produccion: (o.produccion||[]).map(function(pr){
        return {
          fecha: pr.fecha instanceof Date ? Utilities.formatDate(pr.fecha, tz, 'yyyy-MM-dd') : String(pr.fecha||''),
          turno: String(pr.turno||''), lote: String(pr.lote||''),
          operador: String(pr.operador||''), producido: Number(pr.producido)||0,
          comentario: String(pr.comentario||''), sello: String(pr.sello||'')
        };
      })
    };
  });
  return JSON.stringify({ ordenes: ordenes, enviados: enviados });
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// EXPORTAR PRODUCCION POR RANGO DE FECHAS (MENU TRACKING)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
function rastreoExportarProduccion(fechaIniStr, fechaFinStr) {
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd   = ss.getSheetByName('ORDENES');
    var shProd  = ss.getSheetByName('PRODUCCION');
    var dataOrd  = shOrd.getDataRange().getValues();
    var dataProd = shProd.getDataRange().getValues();
    var tz = ss.getSpreadsheetTimeZone();

    // Construir mapa id → fila de ORDENES  (col A = índice 0)
    var mapaOrdenes = {};
    for (var i = 1; i < dataOrd.length; i++) {
      var r = dataOrd[i];
      var idOrd = String(r[0]).trim();
      if (idOrd) mapaOrdenes[idOrd] = r;
    }

    // Rango de fechas (inclusive)
    var dIni = new Date(fechaIniStr + 'T00:00:00');
    var dFin = new Date(fechaFinStr + 'T23:59:59');

    var filas = [];
    for (var p = 1; p < dataProd.length; p++) {
      var pr = dataProd[p];
      // Col F (índice 5) = FECHA
      var fProd = pr[5] ? new Date(pr[5]) : null;
      if (!fProd || isNaN(fProd)) continue;
      if (fProd < dIni || fProd > dFin) continue;

      // Buscar orden enlazada — Col C (índice 2) de PRODUCCION = ID de ORDEN
      var idOrden = String(pr[2]).trim();
      var ord     = mapaOrdenes[idOrden] || [];

      // Formatear fecha como texto dd/mm/yyyy
      var fechaTxt = Utilities.formatDate(fProd, tz, 'dd/MM/yyyy');

      filas.push({
        // De PRODUCCION
        FECHA:      fechaTxt,
        TURNO:      pr[6]  || '',
        // De ORDENES
        PEDIDO:     ord[1]  || '',
        SERIE:      ord[4]  || '',
        ORDEN:      ord[5]  || '',
        CODIGO:     ord[6]  || '',
        DESCRIPCION:ord[7]  || '',
        PROCESO:    ord[11] || '',
        PESO:       Number(ord[18]) || '',
        TIPO:       ord[19] || '',
        DIAMETRO:   ord[20] || '',
        LONGITUD:   ord[21] || '',
        CUERDA:     ord[22] || '',
        CUERPO:     ord[23] || '',
        ACERO:      ord[24] || '',
        // De PRODUCCION (continuación)
        LOTE:       pr[3]  || '',
        MAQUINA:    pr[4]  || '',
        PESO_I:     Number(pr[8])  || '',
        PESO_F:     Number(pr[9])  || '',
        PESO_TINA:  Number(pr[22]) || '',
        PRODUCIDO:  Number(pr[10]) || '',
        SELLO:      pr[14] || '',
        OPERADOR:   pr[17] || '',
        USER:       pr[24] || '',
        CAMBIOS:    pr[26] || ''
      });
    }

    return JSON.stringify({ success: true, filas: filas });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA NUEVO PEDIDO (MENU PEDIDO)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
/**
 * Obtener todos los PEDIDO+PARTIDA existentes en la hoja PEDIDOS
 * para validar duplicados en el cliente.
 */
function npObtenerPedidosExistentes() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("PEDIDOS");
    var data = sh.getDataRange().getValues();
    var datos = [];
    for (var i = 1; i < data.length; i++) {
      var ped = String(data[i][1]).trim();          // Col B (índice 1): PEDIDO
      var par = String(Math.round(Number(data[i][5])||1)); // Col F (índice 5): PARTIDA como "1","2"...
      if (ped) datos.push({ pedido: ped.toUpperCase(), partida: par });
    }
    return JSON.stringify({ success: true, datos: datos });
  } catch(e) {
    return JSON.stringify({ success: false, datos: [], msg: e.toString() });
  }
}

/**
 * Devuelve el catálogo completo de RUTAS como mapa { codNormalizado: {desc, unidad} }
 * para que el frontend haga búsqueda local instantánea sin roundtrips al backend.
 */
function npObtenerCatalogoRutas() {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("RUTAS");
    var data = sh.getDataRange().getValues();
    var mapa = {};
    for (var i = 1; i < data.length; i++) {
      var codHoja = String(data[i][1]).replace(/[^0-9]/g, ''); // Col B
      if (codHoja.length === 8) codHoja = '0' + codHoja;
      if (codHoja && !mapa[codHoja]) {
        mapa[codHoja] = {
          desc:   String(data[i][2]  || '').trim(), // Col C: DESCRIPCION
          unidad: String(data[i][16] || 'PZA').trim() // Col Q: UNIDAD
        };
      }
    }
    return JSON.stringify({ success: true, mapa: mapa });
  } catch(e) {
    return JSON.stringify({ success: false, mapa: {}, msg: e.toString() });
  }
}

/**
 * Buscar un código en Col B de RUTAS.
 * Devuelve Col C (DESCRIPCION, índice 2) y Col Q (UNIDAD, índice 16).
 */
function npBuscarCodigo(codigoRaw) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("RUTAS");
    var data = sh.getDataRange().getValues();
    // Normalizar entrada: solo dígitos, si tiene 8 dígitos agregar cero a la izquierda
    var codNum = String(codigoRaw).replace(/[^0-9]/g, '');
    if (codNum.length === 8) codNum = '0' + codNum;
    for (var i = 1; i < data.length; i++) {
      var codHoja = String(data[i][1]).replace(/[^0-9]/g, ''); // Col B, índice 1
      if (codHoja.length === 8) codHoja = '0' + codHoja;
      if (codHoja === codNum) {
        var desc   = String(data[i][2]  || '').trim(); // Col C, índice 2: DESCRIPCION
        var unidad = String(data[i][16] || 'PZA').trim(); // Col Q, índice 16: UNIDAD
        return JSON.stringify({ success: true, encontrado: true, codigo: codNum, desc: desc, unidad: unidad });
      }
    }
    return JSON.stringify({ success: true, encontrado: false });
  } catch(e) {
    return JSON.stringify({ success: false, encontrado: false, msg: e.toString() });
  }
}

/**
 * Validar múltiples códigos a la vez (importación masiva).
 * Devuelve mapa { codigo: {desc, unidad} }.
 */
function npValidarCodigosMultiples(codigos) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("RUTAS");
    var data = sh.getDataRange().getValues();
    // Construir mapa desde RUTAS una sola vez
    var rutaMap = {};
    for (var i = 1; i < data.length; i++) {
      var codHoja = String(data[i][1]).replace(/[^0-9]/g, ''); // Col B
      if (codHoja.length === 8) codHoja = '0' + codHoja;
      if (codHoja && !rutaMap[codHoja]) {
        rutaMap[codHoja] = {
          desc:   String(data[i][2]  || '').trim(), // Col C: DESCRIPCION
          unidad: String(data[i][16] || 'PZA').trim() // Col Q: UNIDAD
        };
      }
    }
    var mapa = {};
    codigos.forEach(function(c) {
      var cn = String(c).replace(/[^0-9]/g, '');
      if (cn.length === 8) cn = '0' + cn;
      if (rutaMap[cn]) mapa[c] = rutaMap[cn];
    });
    return JSON.stringify({ success: true, mapa: mapa });
  } catch(e) {
    return JSON.stringify({ success: false, mapa: {}, msg: e.toString() });
  }
}

/**
 * Guardar UN pedido individual en la hoja PEDIDOS.
 */
function npGuardarPedido(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("PEDIDOS");
    var data = sh.getDataRange().getValues();
    var pedUp  = String(d.pedido).trim().toUpperCase();
    var parStr = String(d.partida).trim();
    // Verificar duplicado
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === pedUp && String(data[i][5]).trim() === parStr) {
        return JSON.stringify({ success: false, msg: 'Duplicado: PEDIDO + PARTIDA ya existe en Sheets.' });
      }
    }
    var codNum = String(d.codigo).replace(/[^0-9]/g, '');
    if (codNum.length === 8) codNum = '0' + codNum;
    var fechaDate = d.fecha ? new Date(d.fecha + 'T12:00:00') : new Date();
    sh.appendRow([
      Utilities.getUuid().substring(0, 8), // A: ID
      pedUp,                                // B: PEDIDO
      fechaDate,                            // C: FECHA
      "'" + codNum,                         // D: CODIGO
      String(d.desc || '').trim(),          // E: DESCRIPCION
      Number(d.partida) || 1,               // F: PARTIDA
      Number(d.cantidad) || 0,              // G: CANTIDAD
      String(d.unidad || 'PZA').trim(),     // H: UNIDAD
      'ABIERTO'                             // I: ESTADO
    ]);
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Guardar MÚLTIPLES pedidos (cola o importación masiva).
 */
function npGuardarPedidosMasivos(lista) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("PEDIDOS");
    var data = sh.getDataRange().getValues();
    var existentes = {};
    for (var i = 1; i < data.length; i++) {
      var k = String(data[i][1]).trim().toUpperCase() + '||' + String(data[i][5]).trim();
      existentes[k] = true;
    }
    var rows = [], guardados = 0, errores = [];
    lista.forEach(function(d) {
      var pedUp  = String(d.pedido).trim().toUpperCase();
      var parStr = String(d.partida).trim();
      var k = pedUp + '||' + parStr;
      try {
        if (existentes[k]) { errores.push({ pedido: pedUp, partida: parStr, razon: 'Duplicado en hoja' }); return; }
        existentes[k] = true;
        var codNum = String(d.codigo).replace(/[^0-9]/g, '');
        if (codNum.length === 8) codNum = '0' + codNum;
        var fechaDate = d.fecha ? new Date(d.fecha + 'T12:00:00') : new Date();
        rows.push([
          Utilities.getUuid().substring(0, 8),
          pedUp,
          fechaDate,
          "'" + codNum,
          String(d.desc || '').trim(),
          Number(parStr) || 1,
          Number(d.cantidad) || 0,
          String(d.unidad || 'PZA').trim(),
          'ABIERTO'
        ]);
        guardados++;
      } catch(eRow) {
        errores.push({ pedido: pedUp, partida: parStr, razon: eRow.toString() });
      }
    });
    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
      SpreadsheetApp.flush();
    }
    return JSON.stringify({ success: true, guardados: guardados, errores: errores });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA NUEVO PEDIDO INT (MENU PEDIDO INT)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// ── 1. CARGA INICIAL: catálogo de códigos con PEDIDO_INT == "SI" ─────────────────
// Llamada desde: loadPedidoINT() en el HTML al iniciar el módulo
function obtenerCatalogosPedidoINT() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetCod = ss.getSheetByName("CODIGOS");
  var dataCod = sheetCod.getDataRange().getValues();

  var listaValidos = [];
  for (var i = 1; i < dataCod.length; i++) {
    if (String(dataCod[i][6]).toUpperCase().trim() === "SI") {
      listaValidos.push({
        codigo:      String(dataCod[i][0]),
        descripcion: dataCod[i][1],
        unidad:      dataCod[i][2],
        serie:       dataCod[i][3],
        peso:        dataCod[i][4]
      });
    }
  }
  return listaValidos;
}

// ── 2. BUSCA LA RUTA DE UN CÓDIGO (incluye PT, VENTA, PESO, etc.) ────────────────
// Llamada desde: pintCargarRuta() en el HTML al seleccionar un código
function obtenerRutaPedidoINT(codigo) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetRut = ss.getSheetByName('RUTAS');
  var dataRut  = sheetRut.getDataRange().getValues();

  var PROCESOS_MP = ['ESTAMPADO','FORJA','ROSCADO','ESTIRADO','TREFILADO','COLATADO'];

  var pasos = [];
  for (var i = 1; i < dataRut.length; i++) {
    if (String(dataRut[i][1]).trim() === codigo) {
      var proc = String(dataRut[i][4] || '').trim().toUpperCase();
      var mpVal = String(dataRut[i][17] || '').trim();  // Col R = índice 17
      var primeraMP = (PROCESOS_MP.indexOf(proc) >= 0 && mpVal)
                      ? mpVal.split(',')[0].trim().toUpperCase()
                      : '';
      pasos.push({
        sec:      dataRut[i][3],
        proceso:  proc,
        maquinas: String(dataRut[i][5]).split(',').map(function(m){ return m.trim(); }),
        cantLote: parseFloat(dataRut[i][22]) || 0,
        pt:       dataRut[i][12],
        venta:    dataRut[i][13],
        peso:     dataRut[i][15],
        tipo:     dataRut[i][6],
        diam:     dataRut[i][7],
        long:     dataRut[i][8],
        cuerda:   dataRut[i][9],
        cuerpo:   dataRut[i][10],
        acero:    dataRut[i][11],
        mp:       primeraMP  // ← NUEVO
      });
    }
  }
  return pasos.sort(function(a, b){ return a.sec - b.sec; });
}

// ── 3. NÚMERO DE ORDEN SIGUIENTE (MAX+1 por SERIE) ───────────────────────────────
// Función auxiliar interna, usada por crearPedidoInternoCompleto()
function obtenerUltimasOrdenesPorSerie() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("ORDENES");
  var data  = sheet.getDataRange().getValues();
  var maxPorSerie = {};
  for (var i = 1; i < data.length; i++) {
    var serie = String(data[i][4] || '').trim().toUpperCase();
    var num   = parseInt(data[i][5]);
    if (!serie || isNaN(num)) continue;
    if (!maxPorSerie[serie] || num > maxPorSerie[serie]) maxPorSerie[serie] = num;
  }
  return JSON.stringify(maxPorSerie);
}

function getSiguienteNumeroOrden(serie) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("ORDENES");
  var data = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toUpperCase() === serie.toUpperCase()) { // Col E: SERIE
      var num = parseInt(data[i][5]); // Col F: ORDEN
      if (!isNaN(num) && num > max) max = num;
    }
  }
  return max + 1;
}

// ── 4. FOLIO SIGUIENTE (ej: INT-0001, INT-0002, ...) ─────────────────────────────
// Función auxiliar interna, usada por crearPedidoInternoCompleto()
function getSiguienteFolio(prefijo) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("PEDIDOS");
  var data = sheet.getDataRange().getValues();
  var max = 0;

  for (var i = 1; i < data.length; i++) {
    var folioStr = String(data[i][1]); // Columna B = FOLIO
    if (folioStr.startsWith(prefijo)) {
      var partes = folioStr.split("-");
      if (partes.length > 1) {
        var num = parseInt(partes[1]);
        if (!isNaN(num) && num > max) max = num;
      }
    }
  }
  return prefijo + "-" + ("0000" + (max + 1)).slice(-4);
}

// ── 5. GUARDADO COMPLETO: escribe en PEDIDOS + ORDENES + LOTES ───────────────────
// Llamada desde: pintGuardar() en el HTML al confirmar el pedido
function crearPedidoInternoCompleto(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetPed = ss.getSheetByName("PEDIDOS");
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetLot = ss.getSheetByName("LOTES");

    var hoySoloFecha = new Date();
    hoySoloFecha.setHours(0, 0, 0, 0);

    // A. PEDIDOS — Estado: PLANEADO
    var prefijoPedido = (payload.prefijo && payload.prefijo.trim()) ? payload.prefijo.trim().toUpperCase() : "INT";
    var folioINT = getSiguienteFolio(prefijoPedido);
    sheetPed.appendRow([
      Utilities.getUuid().substring(0, 8),
      folioINT,
      hoySoloFecha,
      "'" + String(payload.codigo).padStart(9, "0"),
      payload.descripcion,
      1,
      payload.cantidad,
      payload.unidad,
      "PLANEADO"
    ]);

    // B. ORDENES — Una fila por cada paso de la ruta
    var nOrden = getSiguienteNumeroOrden(payload.serie);
    var firstProcessId = "";
    var resumenOrdenes = [];

    payload.ruta.forEach(function(paso, index) {
      var idProceso = Utilities.getUuid().substring(0, 8);
      if (index === 0) firstProcessId = idProceso;

      // Si la unidad NO es KG/ROL, la cantidad en KG = cantidad × peso unitario
      var solicitado = payload.cantidad;
      if (!["KG", "KILOS", "ROL"].includes(String(payload.unidad).toUpperCase())) {
        solicitado = payload.cantidad * (paso.peso || 0);
      }

      sheetOrd.appendRow([
        idProceso,
        folioINT,
        1,
        hoySoloFecha,
        payload.serie,
        nOrden,
        "'" + String(payload.codigo).padStart(9, "0"),
        payload.descripcion,
        payload.cantidad,
        payload.unidad,
        paso.sec,
        paso.proceso,
        paso.maquinaSeleccionada,
        solicitado,
        0,
        "ABIERTO",
        paso.pt,
        paso.venta,
        paso.peso,
        paso.tipo,
        "'" + String(paso.diam),
        "'" + String(paso.long),
        paso.cuerda,
        paso.cuerpo,
        paso.acero,
        new Date(),
        "",             // AA=26: PRIORIDAD (vacío, se llena aparte)
        "",             // AB=27: FECHA_INICIO_PROG
        "",             // AC=28: FECHA_FIN_PROG
        paso.mp || ""   // AD=29: ALERTA_SECUENCIA = MP
      ]);
      
      resumenOrdenes.push({ proc: paso.proceso, maq: paso.maquinaSeleccionada });
    });

    // C. LOTES — División por cantLote del primer paso
    var cantLoteBase = payload.ruta[0].cantLote || payload.cantidad;
    var nLotes = Math.ceil(payload.cantidad / cantLoteBase);
    var listaLotesGenerados = [];

    for (var l = 1; l <= nLotes; l++) {
      var nombreLote = payload.serie + "." + ("0000" + nOrden).slice(-4) + "." + (l < 100 ? ("00" + l).slice(-2) : String(l));
      sheetLot.appendRow([
        Utilities.getUuid().substring(0, 8),
        payload.serie,
        firstProcessId,
        l,
        nombreLote,
        cantLoteBase,
        0,
        "ABIERTO",
        new Date()
      ]);
      listaLotesGenerados.push(nombreLote);
    }

    return {
      success:  true,
      folio:    folioINT,
      ordenNum: nOrden,
      ordenes:  resumenOrdenes,
      lotes:    listaLotesGenerados
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA MENU MRP
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

var ID_CIRCULANTE = '1v21_Glgvk3ZV4SYpsMGbqNc97I7MO98BqkwmiJvYvnI';

// ────────────────────────────────────────────────────────────
// FUNCIÓN PRINCIPAL
// serie: 'T' | 'P' | 'F' | 'TODAS' (filtro de series)
// ────────────────────────────────────────────────────────────
function mrpmpObtenerDatos(serie) {
  try {
    serie = (serie || 'TODAS').toString().toUpperCase().trim();

    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd = ss.getSheetByName('ORDENES');
    var shEst = ss.getSheetByName('ESTANDARES');
    var shRut = ss.getSheetByName('RUTAS');
    var shInv = ss.getSheetByName('INVENTARIO_EXTERNO');
    var shProd = ss.getSheetByName('PRODUCCION');

    if (!shOrd || !shEst) throw new Error('Hojas ORDENES o ESTANDARES no encontradas');

    var PROCESOS_MP = ['ESTAMPADO','FORJA','ROSCADO','ESTIRADO','TREFILADO','COLATADO'];

    // ── 0. DESCRIPCIÓN DE ALAMBRES DESDE RUTAS ────────────────
    var mapaDesc = {};
    if (shRut) {
      var dataRutDesc = shRut.getDataRange().getValues();
      var hRutDesc    = dataRutDesc[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var iRutCod  = hRutDesc.indexOf('CODIGO');
      var iRutDesc = hRutDesc.indexOf('DESCRIPCION');
      for (var rd = 1; rd < dataRutDesc.length; rd++) {
        var codRD  = String(dataRutDesc[rd][iRutCod]  || '').trim().toUpperCase();
        var descRD = String(dataRutDesc[rd][iRutDesc] || '').trim();
        if (codRD && descRD && !mapaDesc[codRD]) mapaDesc[codRD] = descRD;
      }
    }

    // ── 0b. ESTÁNDARES ─────────────────────────────────────────
    var dataEst = shEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ieMAQ   = hEst.indexOf('MAQUINA');
    var ieVEL   = hEst.indexOf('VELOCIDAD');
    var ieEFIC  = hEst.indexOf('EFICIENCIA');
    var ieTURN  = hEst.indexOf('TURNOS');
    var ieUNID  = hEst.indexOf('UNIDAD_VEL');
    var iePROC  = hEst.indexOf('PROCESO');

    var mapaEst     = {};
    var maqsPermitidas = {};

    for (var ei = 1; ei < dataEst.length; ei++) {
      var eProc = String(dataEst[ei][iePROC] || '').toUpperCase().trim();
      var eMaq  = String(dataEst[ei][ieMAQ]  || '').trim().toUpperCase();
      if (!eMaq) continue;
      mapaEst[eMaq] = {
        vel:    Number(dataEst[ei][ieVEL]  || 0),
        efic:   Number(dataEst[ei][ieEFIC] || 1),
        turnos: Number(dataEst[ei][ieTURN] || 1),
        unid:   String(dataEst[ei][ieUNID] || '').toUpperCase(),
        proc:   eProc
      };
      if (PROCESOS_MP.indexOf(eProc) >= 0) {
        maqsPermitidas[eMaq] = true;
      }
    }

    // ── 1. LEER ÓRDENES VIVAS ─────────────────────────────────
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iOrden  = hOrd.indexOf('ORDEN');
    var iCodigo = hOrd.indexOf('CODIGO');
    var iSec    = hOrd.indexOf('SEC');
    var iSol    = hOrd.indexOf('SOLICITADO');
    var iProd   = hOrd.indexOf('PRODUCIDO');
    var iEstado = hOrd.indexOf('ESTADO');
    var iSerie  = hOrd.indexOf('SERIE');
    var iMaq    = hOrd.indexOf('MAQUINA');
    var iPesoO  = hOrd.indexOf('PESO');
    var iUnidO  = hOrd.indexOf('UNIDAD');
    var iPrio   = hOrd.indexOf('PRIORIDAD');
    var iAD     = hOrd.indexOf('ALERTA_SECUENCIA');  // ← Col AD: MP guardada

    var ESTADOS_MUERTOS = ['TERMINADO','CANCELADO','SOBREPRODUCCION'];
    var mapaOrdenes = {};

    for (var i = 1; i < dataOrd.length; i++) {
      var fila   = dataOrd[i];
      var estado = String(fila[iEstado] || '').trim().toUpperCase();
      var orden  = String(fila[iOrden]  || '').trim().toUpperCase();
      var codigo = String(fila[iCodigo] || '').trim().toUpperCase();
      var serieF = String(fila[iSerie]  || '').trim().toUpperCase();
      var sec    = Number(fila[iSec]    || 9999);
      var sol    = Number(fila[iSol]    || 0);
      var prod   = Number(fila[iProd]   || 0);
      var maq    = String(fila[iMaq]    || '').trim().toUpperCase();
      var pesoU  = Number(fila[iPesoO]  || 0);
      var unidO  = String(fila[iUnidO]  || '').trim().toUpperCase();
      var prio   = Number(fila[iPrio]   || 9999);
      // ↓ MP leída directo de Col AD
      var mpOrden = iAD >= 0 ? String(fila[iAD] || '').trim().toUpperCase() : '';

      if (!orden || !codigo) continue;
      if (ESTADOS_MUERTOS.indexOf(estado) >= 0) continue;
      if (serie !== 'TODAS' && serieF !== serie) continue;
      if (!maqsPermitidas[maq]) continue;

      if (!mapaOrdenes[orden] || sec < mapaOrdenes[orden].minSec) {
        mapaOrdenes[orden] = {
          codigo: codigo, serie: serieF, minSec: sec,
          sol: sol, prod: prod, maq: maq,
          pesoU: pesoU, unidO: unidO, prio: prio,
          mp: mpOrden  // ← guardamos MP de la orden
        };
      }
    }

    var ordenesVivas = [];
    var keysOrd = Object.keys(mapaOrdenes);
    for (var j = 0; j < keysOrd.length; j++) {
      var ov   = mapaOrdenes[keysOrd[j]];
      var pend = Math.max(0, ov.sol - ov.prod);
      if (pend > 0) {
        ordenesVivas.push({
          orden:     keysOrd[j],
          codigo:    ov.codigo,
          maq:       ov.maq,
          pendiente: pend,
          prio:      ov.prio,
          pesoU:     ov.pesoU || 0,
          unidO:     ov.unidO || 'KG',
          mp:        ov.mp    // ← ya viene de Col AD
        });
      }
    }

    // ── 2. YA NO SE NECESITA MAPA MP DESDE RUTAS ─────────────
    // La MP viene directo de ov.mp (Col AD de ORDENES)

    // ── 3. SIMULADOR DE 5 DÍAS POR MÁQUINA ───────────────────
    var ahora  = new Date();
    var TZ     = Session.getScriptTimeZone();
    var NOMBRES_DIA = ['DOMINGO','LUNES','MARTES','MIÉRCOLES','JUEVES','VIERNES','SÁBADO'];
    var MESES       = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];

    function mrpmp_labelDia(d, idx) {
      var dow   = d.getDay();
      var dia   = d.getDate();
      var mes   = MESES[d.getMonth()];
      var fecha = ('0'+dia).slice(-2) + '-' + mes;
      var nom   = NOMBRES_DIA[dow];
      if (idx === 0) return 'HOY (' + nom + ' ' + fecha + ')';
      if (idx === 1) return 'MAÑANA (' + nom + ' ' + fecha + ')';
      return nom + ' ' + fecha;
    }

    function mrpmp_fechasLaborales(n) {
      var dias = [], d = new Date(ahora);
      d.setHours(0, 0, 0, 0);
      var idx = 0;
      while (dias.length < n) {
        if (d.getDay() !== 0) {
          dias.push({
            fecha: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
            label: mrpmp_labelDia(d, idx)
          });
          idx++;
        }
        d.setDate(d.getDate() + 1);
      }
      return dias;
    }
    var diasLab = mrpmp_fechasLaborales(5);

    var planMP    = [{},{},{},{},{}];
    var detalleMaq = {};

    var colasMaq = {};
    ordenesVivas.forEach(function(ov) {
      if (!ov.maq || !ov.mp) return;  // sin MP asignada → omitir
      if (!colasMaq[ov.maq]) colasMaq[ov.maq] = [];
      colasMaq[ov.maq].push({
        orden:     ov.orden,
        pendiente: ov.pendiente,
        mp:        ov.mp,
        prio:      ov.prio,
        pesoU:     ov.pesoU,
        unidO:     ov.unidO
      });
    });
    Object.keys(colasMaq).forEach(function(maq) {
      colasMaq[maq].sort(function(a, b) { return a.prio - b.prio; });
    });

    Object.keys(colasMaq).forEach(function(maq) {
      var std = mapaEst[maq] || null;
      if (!std || std.vel <= 0) return;

      var vel  = std.vel;
      var efic = std.efic;
      var hpd  = std.turnos === 3 ? 22.5 : std.turnos === 2 ? 14.5 : 7.5;

      var cola = colasMaq[maq].map(function(o) {
        var kgPH = 0;
        if (vel > 0) {
          if (std.unid.indexOf('PZA') >= 0 || std.unid.indexOf('MIN') >= 0) {
            var pesoU = o.pesoU || 0;
            if (pesoU > 0) {
              kgPH = vel * pesoU * 60 * efic;
            }
          } else {
            kgPH = vel * efic;
          }
        }
        return { orden: o.orden, restante: o.pendiente, mp: o.mp, kgPH: kgPH };
      });
      var colaIdx = 0;

      if (!detalleMaq[maq]) detalleMaq[maq] = [{},{},{},{},{}];

      for (var dIdx = 0; dIdx < 5; dIdx++) {
        var horasRest = hpd;
        while (horasRest > 0.001 && colaIdx < cola.length) {
          var item = cola[colaIdx];
          if (item.kgPH <= 0) { colaIdx++; continue; }
          var horasOrden = item.restante / item.kgPH;
          var horasUsar  = Math.min(horasOrden, horasRest);
          var consumir   = horasUsar * item.kgPH;
          var mpKey = item.mp;
          if (!planMP[dIdx][mpKey]) planMP[dIdx][mpKey] = 0;
          planMP[dIdx][mpKey] += consumir;
          if (!detalleMaq[maq][dIdx][mpKey]) detalleMaq[maq][dIdx][mpKey] = 0;
          detalleMaq[maq][dIdx][mpKey] += consumir;
          item.restante -= consumir;
          horasRest     -= horasUsar;
          if (item.restante <= 0.001) colaIdx++;
        }
      }
    });

    // ── 4. EXPLOSIÓN TOTAL ────────────────────────────────────
    var reqAlambron = {};
    var reqAlambres = {};

    ordenesVivas.forEach(function(ov) {
      var mp = ov.mp;  // ← directamente de Col AD
      if (!mp) return;
      var mpU = mp.toUpperCase();

      if (mpU.indexOf('MM') >= 0) {
        if (!reqAlambron[mpU]) reqAlambron[mpU] = { material: mpU, kg: 0, ordenes: 0, dias: [0,0,0,0,0] };
        reqAlambron[mpU].kg      += ov.pendiente;
        reqAlambron[mpU].ordenes += 1;
      } else {
        var desc = mapaDesc[mpU] || mpU;
        if (!reqAlambres[mpU]) reqAlambres[mpU] = { codigo: mpU, descripcion: desc, kg: 0, ordenes: 0, dias: [0,0,0,0,0] };
        reqAlambres[mpU].kg      += ov.pendiente;
        reqAlambres[mpU].ordenes += 1;
      }
    });

    for (var dI = 0; dI < 5; dI++) {
      Object.keys(planMP[dI]).forEach(function(mpK) {
        var kgD = planMP[dI][mpK];
        if (mpK.indexOf('MM') >= 0) {
          if (reqAlambron[mpK]) reqAlambron[mpK].dias[dI] += kgD;
        } else {
          if (reqAlambres[mpK]) reqAlambres[mpK].dias[dI] += kgD;
        }
      });
    }

    function sumarDias(arr) { return arr.reduce(function(a,b){ return a+b; }, 0); }
    Object.values(reqAlambron).forEach(function(r) {
      r.kgDia    = Math.round(r.dias[0]);
      r.kgSemana = Math.round(sumarDias(r.dias));
      r.dias     = r.dias.map(Math.round);
    });
    Object.values(reqAlambres).forEach(function(r) {
      r.kgDia    = Math.round(r.dias[0]);
      r.kgSemana = Math.round(sumarDias(r.dias));
      r.dias     = r.dias.map(Math.round);
    });

    var arrAlambron = Object.values(reqAlambron).sort(function(a,b){ return b.kg - a.kg; });
    var arrAlambres = Object.values(reqAlambres).sort(function(a,b){ return b.kg - a.kg; });

    var detalleFinal = [];
    Object.keys(detalleMaq).sort().forEach(function(maq) {
      var filasDias = detalleMaq[maq].map(function(dObj) {
        return Object.keys(dObj).map(function(mp) {
          return { mp: mp, kg: Math.round(dObj[mp]) };
        }).filter(function(x){ return x.kg > 0; })
          .sort(function(a,b){ return b.kg - a.kg; });
      });
      detalleFinal.push({ maquina: maq, dias: filasDias });
    });

    var resumenDias = diasLab.map(function(d, idx) {
      var mps = planMP[idx];
      return {
        label: d.label,
        fecha: d.fecha,
        mps: Object.keys(mps).map(function(mp) {
          return { mp: mp, kg: Math.round(mps[mp]) };
        }).sort(function(a,b){ return b.kg - a.kg; })
      };
    });

    var existencias = mrpmpCalcularExistencias(ss, shInv, shProd, arrAlambres, mapaDesc);
    var salidas     = mrpmpLeerSalidas(shInv);

    return JSON.stringify({
      success:     true,
      serie:       serie,
      existencias: existencias,
      reqAlambron: arrAlambron,
      reqAlambres: arrAlambres,
      salidas:     salidas,
      detalleMaq:  detalleFinal,
      resumenDias: resumenDias
    });

  } catch (e) {
    Logger.log('mrpmpObtenerDatos ERROR: ' + e.message);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ────────────────────────────────────────────────────────────
// CALCULAR EXISTENCIAS
// ────────────────────────────────────────────────────────────
function mrpmpCalcularExistencias(ss, shInv, shProd, arrAlambresReq, mapaDesc) {
  var FECHA_CORTE = new Date('2026-03-01T00:00:00');
  mapaDesc = mapaDesc || {};

  var codigos = {};

  function initCod(cod) {
    var codU = cod.toUpperCase();
    if (!codigos[codU]) {
      codigos[codU] = {
        codigo:       codU,
        descripcion:  mapaDesc[codU] || codU,
        entCompra:    0,
        entTransform: 0,
        salidas:      0
      };
    }
    return codU;
  }

  // ── A) COMPRAS desde CIRCULANTE ──────────────────────────
  // Col B(1)=FECHA  Col D(3)=CODIGO  Col M(12)=KILOS  Col S(18)=MOVIMIENTO
  try {
    var ssCir = SpreadsheetApp.openById(ID_CIRCULANTE);
    var shCir = ssCir.getSheetByName('CIRCULANTE');
    if (shCir) {
      var dataCir = shCir.getDataRange().getValues();
      for (var c = 1; c < dataCir.length; c++) {
        var fechaCir = dataCir[c][1];
        if (fechaCir instanceof Date && fechaCir < FECHA_CORTE) continue;
        var movCir = String(dataCir[c][18] || '').trim().toUpperCase();
        if (movCir !== 'ALAMBRES') continue;
        var codCir = String(dataCir[c][3]  || '').trim().toUpperCase();
        var kgCir  = Number(dataCir[c][12] || 0);
        if (!codCir || kgCir <= 0) continue;
        var k = initCod(codCir);
        codigos[k].entCompra += kgCir;
      }
    }
  } catch (e) {
    Logger.log('mrpmpCalcularExistencias - CIRCULANTE: ' + e.message);
  }

  // ── B) TRANSFORMACIONES desde PRODUCCION ─────────────────
  try {
    var shOrd2   = ss.getSheetByName('ORDENES');
    var dataOrd2 = shOrd2.getDataRange().getValues();
    var hOrd2    = dataOrd2[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iOrd2Ord = hOrd2.indexOf('ORDEN');
    var iOrd2Cod = hOrd2.indexOf('CODIGO');
    var iOrd2Est = hOrd2.indexOf('ESTADO');

    var mapaOrdenT = {};
    for (var o2 = 1; o2 < dataOrd2.length; o2++) {
      var ord2 = String(dataOrd2[o2][iOrd2Ord] || '').trim().toUpperCase();
      var est2 = String(dataOrd2[o2][iOrd2Est] || '').trim().toUpperCase();
      var cod2 = String(dataOrd2[o2][iOrd2Cod] || '').trim().toUpperCase();
      if (!ord2 || !cod2) continue;
      if (ord2.indexOf('T.') !== 0) continue;
      if (est2 !== 'TERMINADO') continue;
      mapaOrdenT[ord2] = cod2;
    }

    var dataProd = shProd.getDataRange().getValues();
    var hProd    = dataProd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iProdFec = hProd.indexOf('FECHA');
    var iProdOrd = hProd.indexOf('ORDEN');
    var iProdKg  = hProd.indexOf('PRODUCIDO');
    var iProdPT  = hProd.indexOf('PT');

    for (var p = 1; p < dataProd.length; p++) {
      var ordP  = String(dataProd[p][iProdOrd] || '').trim().toUpperCase();
      var ptP   = String(dataProd[p][iProdPT]  || '').trim().toUpperCase();
      var kgP   = Number(dataProd[p][iProdKg]  || 0);
      if (!ordP || ptP !== 'SI' || kgP <= 0) continue;
      var fechaP = dataProd[p][iProdFec >= 0 ? iProdFec : 0];
      if (fechaP instanceof Date && fechaP < FECHA_CORTE) continue;
      var codAlambre = mapaOrdenT[ordP];
      if (!codAlambre) continue;
      var k = initCod(codAlambre);
      codigos[k].entTransform += kgP;
    }
  } catch (e) {
    Logger.log('mrpmpCalcularExistencias - transformaciones: ' + e.message);
  }

  // ── C) SALIDAS desde INV_ALAMBRES ────────────────────────
  try {
    var dataInv = shInv.getDataRange().getValues();
    for (var inv = 1; inv < dataInv.length; inv++) {
      var fechaInv = dataInv[inv][0];
      if (fechaInv instanceof Date && fechaInv < FECHA_CORTE) continue;
      var tipoInv = String(dataInv[inv][1] || '').trim().toUpperCase();
      var codInv  = String(dataInv[inv][4] || '').trim().toUpperCase();
      var kgInv   = Number(dataInv[inv][6] || 0);
      if (!codInv || kgInv <= 0) continue;
      var k = initCod(codInv);
      if (tipoInv === 'SALIDA') codigos[k].salidas += kgInv;
    }
  } catch (e) {
    Logger.log('mrpmpCalcularExistencias - INV_ALAMBRES: ' + e.message);
  }

  // ── D) Asegurar que alambres requeridos aparezcan ─────────
  arrAlambresReq.forEach(function(a) { initCod(a.codigo); });

  var lista = [];
  Object.values(codigos).forEach(function(c) {
    c.exist = Math.max(0, (c.entCompra + c.entTransform) - c.salidas);
    lista.push(c);
  });
  lista.sort(function(a, b){ return b.exist - a.exist; });
  return lista;
}

// ────────────────────────────────────────────────────────────
// LEER ÚLTIMAS SALIDAS
// ────────────────────────────────────────────────────────────
function mrpmpLeerSalidas(shInv) {
  var salidas = [];
  try {
    var data = shInv.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var tipo = String(data[i][1] || '').trim().toUpperCase();
      if (tipo !== 'SALIDA') continue;
      var fechaRaw = data[i][0];
      var fechaStr = '';
      if (fechaRaw instanceof Date) {
        var d = fechaRaw, m = d.getMonth()+1, dd = d.getDate();
        fechaStr = d.getFullYear()+'-'+(m<10?'0':'')+m+'-'+(dd<10?'0':'')+dd;
      } else {
        fechaStr = String(fechaRaw || '').substring(0, 10);
      }
      salidas.push({
        fecha:  fechaStr,
        codigo: String(data[i][4] || '').trim().toUpperCase(),
        kg:     Number(data[i][6] || 0),
        ref:    String(data[i][3] || '').trim().toUpperCase()
      });
    }
  } catch (e) {
    Logger.log('mrpmpLeerSalidas ERROR: ' + e.message);
  }
  return salidas;
}

// ────────────────────────────────────────────────────────────
// GUARDAR SALIDA en INV_ALAMBRES
// ────────────────────────────────────────────────────────────
function mrpmpGuardarSalida(datos) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shInv = ss.getSheetByName('INV_ALAMBRES');
    if (!shInv) throw new Error('No existe la hoja INV_ALAMBRES.');

    var fecha   = new Date(datos.fecha + 'T12:00:00');
    var usuario = Session.getActiveUser().getEmail() || 'SISTEMA';

    shInv.appendRow([
      fecha,
      'SALIDA',
      'CONSUMO',
      String(datos.ref    || '').toUpperCase(),
      String(datos.codigo || '').toUpperCase(),
      '',
      Number(datos.kg),
      'USUARIO: ' + usuario.toUpperCase()
    ]);

    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log('mrpmpGuardarSalida ERROR: ' + e.message);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ────────────────────────────────────────────────────────────
// OBTENER LISTA DE CÓDIGOS PARA AUTOCOMPLETE DE SALIDAS
// ────────────────────────────────────────────────────────────
function mrpmpObtenerCodigosAlambre() {
  try {
    var lista = [];
    var vistos = {};

    // Desde CIRCULANTE
    try {
      var ssCir = SpreadsheetApp.openById(ID_CIRCULANTE);
      var shCir = ssCir.getSheetByName('CIRCULANTE');
      if (shCir) {
        var dataCir = shCir.getDataRange().getValues();
        for (var c = 1; c < dataCir.length; c++) {
          var mov = String(dataCir[c][18] || '').trim().toUpperCase();
          if (mov !== 'ALAMBRES') continue;
          var cod = String(dataCir[c][3] || '').trim().toUpperCase();
          if (cod && !vistos[cod]) { lista.push(cod); vistos[cod] = true; }
        }
      }
    } catch(e) {}

    // Desde INV_ALAMBRES
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shInv = ss.getSheetByName('INV_ALAMBRES');
    if (shInv) {
      var dataInv = shInv.getDataRange().getValues();
      for (var i = 1; i < dataInv.length; i++) {
        var cod = String(dataInv[i][4] || '').trim().toUpperCase();
        if (cod && !vistos[cod]) { lista.push(cod); vistos[cod] = true; }
      }
    }

    lista.sort();
    return JSON.stringify({ success: true, codigos: lista });
  } catch (e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ────────────────────────────────────────────────────────────
// CREAR HOJA INV_ALAMBRES (ejecutar una sola vez manualmente)
// ────────────────────────────────────────────────────────────
function mrpmpCrearHojaInventario() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);

  if (ss.getSheetByName('INV_ALAMBRES')) {
    Logger.log('INV_ALAMBRES ya existe — no se creó de nuevo.');
    return;
  }

  var sh = ss.insertSheet('INV_ALAMBRES');
  var headers = ['FECHA', 'TIPO', 'ORIGEN', 'REFERENCIA', 'MATERIAL', 'CALIBRE', 'CANTIDAD_KG', 'NOTAS'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#0f172a').setFontColor('white').setFontWeight('bold');
  sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 90);  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 120); sh.setColumnWidth(6, 90);
  sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 200);
  sh.getRange('A2:A1000').setNumberFormat('dd/mm/yyyy');
  sh.getRange('B2:B1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['ENTRADA','SALIDA'], true).setAllowInvalid(false).build());
  sh.getRange('C2:C1000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['COMPRA','TRANSFORMACION','CONSUMO','AJUSTE'], true).setAllowInvalid(false).build());
  sh.setFrozenRows(1);

  Logger.log('✅ Hoja INV_ALAMBRES creada correctamente.');
  try { SpreadsheetApp.getUi().alert('✅ Hoja INV_ALAMBRES creada correctamente.'); } catch(e) {}
}

// ────────────────────────────────────────────────────────────
// DETALLE DE ÓRDENES PARA UNA MÁQUINA Y DÍA ESPECÍFICO
// Carga lazy — se llama al hacer clic en una celda del plan
// ────────────────────────────────────────────────────────────
function mrpmpObtenerDetalleDia(params) {
  try {
    var maquina = String(params.maquina || '').trim().toUpperCase();
    var diaIdx  = Number(params.diaIdx  || 0);
    var serie   = String(params.serie   || 'TODAS').trim().toUpperCase();

    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd = ss.getSheetByName('ORDENES');
    var shEst = ss.getSheetByName('ESTANDARES');

    var dataEst = shEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iEM = hEst.indexOf('MAQUINA'), iEV = hEst.indexOf('VELOCIDAD');
    var iEF = hEst.indexOf('EFICIENCIA'), iET = hEst.indexOf('TURNOS');
    var iEU = hEst.indexOf('UNIDAD_VEL');
    var std = null;
    for (var ei = 1; ei < dataEst.length; ei++) {
      if (String(dataEst[ei][iEM]||'').trim().toUpperCase() === maquina) {
        std = {
          vel:    Number(dataEst[ei][iEV] || 0),
          efic:   Number(dataEst[ei][iEF] || 1),
          turnos: Number(dataEst[ei][iET] || 1),
          unid:   String(dataEst[ei][iEU] || '').toUpperCase()
        };
        break;
      }
    }
    if (!std || std.vel <= 0) return JSON.stringify({ success: false, msg: 'Sin estandar para ' + maquina });

    // Leer órdenes — MP desde Col AD
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var gc = function(n){ var i=hOrd.indexOf(n); return i>=0?i:9999; };
    var iOrden=gc('ORDEN'), iCod=gc('CODIGO'), iSec=gc('SEC'), iSol=gc('SOLICITADO');
    var iProd=gc('PRODUCIDO'), iEst=gc('ESTADO'), iSerie=gc('SERIE'), iMaq=gc('MAQUINA');
    var iPrio=gc('PRIORIDAD'), iPeso=gc('PESO'), iUnid=gc('UNIDAD'), iTipo=gc('TIPO');
    var iDiam=gc('DIAMETRO'), iLong=gc('LONGITUD'), iCuerda=gc('CUERDA');
    var iCuerpo=gc('CUERPO'), iAcero=gc('ACERO');
    var iAD=gc('ALERTA_SECUENCIA');  // ← MP desde Col AD

    var MUERTOS = ['TERMINADO','CANCELADO','SOBREPRODUCCION'];
    var mapaOrd = {};
    for (var i = 1; i < dataOrd.length; i++) {
      var f = dataOrd[i];
      if (MUERTOS.indexOf(String(f[iEst]||'').trim().toUpperCase()) >= 0) continue;
      if (String(f[iMaq]||'').trim().toUpperCase() !== maquina) continue;
      if (serie !== 'TODAS' && String(f[iSerie]||'').trim().toUpperCase() !== serie) continue;
      var ord  = String(f[iOrden]||'').trim().toUpperCase();
      var sec  = Number(f[iSec] ||9999);
      var pend = Math.max(0, Number(f[iSol]||0) - Number(f[iProd]||0));
      if (pend <= 0) continue;
      if (!mapaOrd[ord] || sec < mapaOrd[ord].sec) {
        mapaOrd[ord] = {
          orden:  ord, codigo: String(f[iCod]   ||'').trim().toUpperCase(),
          tipo:   String(f[iTipo]  ||'').trim().toUpperCase(),
          diam:   String(f[iDiam]  ||'').trim(),
          long:   String(f[iLong]  ||'').trim(),
          cuerda: String(f[iCuerda]||'').trim().toUpperCase(),
          cuerpo: String(f[iCuerpo]||'').trim().toUpperCase(),
          acero:  String(f[iAcero] ||'').trim().toUpperCase(),
          prio:   Number(f[iPrio]  ||9999), pend: pend,
          pesoU:  Number(f[iPeso]  ||0),    sec: sec,
          mp:     String(f[iAD]    ||'').trim().toUpperCase()  // ← desde Col AD
        };
      }
    }

    var cola = Object.values(mapaOrd).sort(function(a,b){ return a.prio - b.prio; });
    var hpd  = std.turnos === 3 ? 22.5 : std.turnos === 2 ? 14.5 : 7.5;

    var colaViva = cola.map(function(o) {
      var kgPH = 0;
      if (std.unid.indexOf('PZA') >= 0 || std.unid.indexOf('MIN') >= 0) {
        if (o.pesoU > 0) kgPH = std.vel * o.pesoU * 60 * std.efic;
      } else {
        kgPH = std.vel * std.efic;
      }
      return { restante: o.pend, kgPH: kgPH, meta: o };
    });

    var colaIdx = 0;
    for (var d = 0; d <= diaIdx; d++) {
      var horasRest  = hpd;
      var consumoDia = [];
      while (horasRest > 0.001 && colaIdx < colaViva.length) {
        var item      = colaViva[colaIdx];
        if (item.kgPH <= 0) { colaIdx++; continue; }
        var horasUsar = Math.min(item.restante / item.kgPH, horasRest);
        var consumir  = horasUsar * item.kgPH;
        if (d === diaIdx) consumoDia.push({ kg: consumir, meta: item.meta });
        item.restante -= consumir;
        horasRest     -= horasUsar;
        if (item.restante <= 0.001) colaIdx++;
      }
      if (d === diaIdx) {
        var grupos = {};
        consumoDia.forEach(function(c) {
          var mp = c.meta.mp || 'SIN MP';  // ← desde Col AD
          if (!grupos[mp]) grupos[mp] = [];
          grupos[mp].push({
            orden:  c.meta.orden, codigo: c.meta.codigo,
            tipo:   c.meta.tipo,  diam:   c.meta.diam,
            long:   c.meta.long,  cuerda: c.meta.cuerda,
            cuerpo: c.meta.cuerpo,acero:  c.meta.acero,
            kg:     Math.round(c.kg)
          });
        });
        return JSON.stringify({ success: true, maquina: maquina, grupos: grupos });
      }
    }
    return JSON.stringify({ success: true, maquina: maquina, grupos: {} });
  } catch(e) {
    Logger.log('mrpmpObtenerDetalleDia ERROR: ' + e.message);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

//*** OJO ESTAS LAS METI NUEVAS PARA HACER EL CAMBIO Y REEMPLACE OTRAS */
// ─────────────────────────────────────────────────────────────────────────────
// [1] NUEVA — rellenarMPMasivo()
//     Recorre ORDENES, para filas cuyo PROCESO esté en los 6 procesos MP,
//     busca la primera MP en RUTAS Col R y la escribe en Col AD (índice 29).
//     Llama desde el módulo de relleno masivo en el HTML.
// ─────────────────────────────────────────────────────────────────────────────
function rellenarMPMasivo() {
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd   = ss.getSheetByName('ORDENES');
    var shRut   = ss.getSheetByName('RUTAS');

    var PROCESOS_MP = ['ESTAMPADO','FORJA','ROSCADO','ESTIRADO','TREFILADO','COLATADO'];

    // Leer RUTAS → mapa codigo → { primera MP, todas las MP }
    var dataRut = shRut.getDataRange().getValues();
    var hRut    = dataRut[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var rCOD = hRut.indexOf('CODIGO');
    var rPROC = hRut.indexOf('PROCESO');
    var rMP  = hRut.indexOf('MP');

    // mapaMP[codigo][proceso] = "MP1,MP2,..."
    var mapaMP = {};
    for (var r = 1; r < dataRut.length; r++) {
      var cod   = String(dataRut[r][rCOD]  || '').trim().toUpperCase();
      var proc  = String(dataRut[r][rPROC] || '').trim().toUpperCase();
      var mpVal = String(dataRut[r][rMP]   || '').trim();
      if (!cod || !mpVal) continue;
      if (PROCESOS_MP.indexOf(proc) < 0) continue;
      if (!mapaMP[cod]) mapaMP[cod] = {};
      if (!mapaMP[cod][proc]) mapaMP[cod][proc] = mpVal;
    }

    // Leer ORDENES
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iCOD    = hOrd.indexOf('CODIGO');
    var iPROC   = hOrd.indexOf('PROCESO');
    var iAD     = hOrd.indexOf('ALERTA_SECUENCIA');  // Col AD = índice 29

    if (iAD < 0) {
      return JSON.stringify({ success: false, msg: 'No se encontró columna ALERTA_SECUENCIA en ORDENES' });
    }

    var actualizadas = 0;
    var sinMP = 0;

    for (var i = 1; i < dataOrd.length; i++) {
      var proc   = String(dataOrd[i][iPROC] || '').trim().toUpperCase();
      var codigo = String(dataOrd[i][iCOD]  || '').trim().toUpperCase();
      if (PROCESOS_MP.indexOf(proc) < 0) continue;

      var mpActual = String(dataOrd[i][iAD] || '').trim();
      if (mpActual) continue;  // ya tiene MP, no sobreescribir

      var mpEntry = mapaMP[codigo] && mapaMP[codigo][proc];
      if (!mpEntry) { sinMP++; continue; }

      var primera = mpEntry.split(',')[0].trim().toUpperCase();
      shOrd.getRange(i + 1, iAD + 1).setValue(primera);
      actualizadas++;
    }

    return JSON.stringify({
      success: true,
      actualizadas: actualizadas,
      sinMP: sinMP,
      msg: 'Listo. ' + actualizadas + ' filas actualizadas. ' + sinMP + ' sin MP en RUTAS.'
    });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// [2] NUEVA — obtenerOpcionesMP(idOrden)
//     Devuelve la MP actual guardada en Col AD y las opciones disponibles
//     según RUTAS Col R para ese código+proceso.
// ─────────────────────────────────────────────────────────────────────────────
function obtenerOpcionesMP(idOrden) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd = ss.getSheetByName('ORDENES');
    var shRut = ss.getSheetByName('RUTAS');

    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iID     = hOrd.indexOf('ID');
    var iCOD    = hOrd.indexOf('CODIGO');
    var iPROC   = hOrd.indexOf('PROCESO');
    var iAD     = hOrd.indexOf('ALERTA_SECUENCIA');

    var filaOrd = -1;
    var codigo  = '';
    var proceso = '';
    var mpActual = '';

    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][iID] || '').trim() === String(idOrden).trim()) {
        filaOrd  = i + 1;  // fila real (base 1)
        codigo   = String(dataOrd[i][iCOD]  || '').trim().toUpperCase();
        proceso  = String(dataOrd[i][iPROC] || '').trim().toUpperCase();
        mpActual = String(dataOrd[i][iAD]   || '').trim().toUpperCase();
        break;
      }
    }
    if (filaOrd < 0) return JSON.stringify({ success: false, msg: 'Orden no encontrada' });

    // Buscar opciones en RUTAS
    var dataRut = shRut.getDataRange().getValues();
    var hRut    = dataRut[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var rCOD    = hRut.indexOf('CODIGO');
    var rPROC   = hRut.indexOf('PROCESO');
    var rMP     = hRut.indexOf('MP');

    var opciones = [];
    for (var r = 1; r < dataRut.length; r++) {
      if (String(dataRut[r][rCOD] ||'').trim().toUpperCase() === codigo &&
          String(dataRut[r][rPROC]||'').trim().toUpperCase() === proceso) {
        var mpVal = String(dataRut[r][rMP] || '').trim();
        if (mpVal) {
          mpVal.split(',').forEach(function(m) {
            var mt = m.trim().toUpperCase();
            if (mt && opciones.indexOf(mt) < 0) opciones.push(mt);
          });
        }
        break;
      }
    }

    return JSON.stringify({
      success:   true,
      idOrden:   idOrden,
      codigo:    codigo,
      proceso:   proceso,
      mpActual:  mpActual,
      opciones:  opciones
    });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// [3] NUEVA — guardarMP(idOrden, nuevaMP)
//     Guarda la MP elegida en Col AD:ALERTA_SECUENCIA de la orden indicada.
// ─────────────────────────────────────────────────────────────────────────────
function guardarMP(idOrden, nuevaMP) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd = ss.getSheetByName('ORDENES');
    var data  = shOrd.getDataRange().getValues();
    var hOrd  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iID   = hOrd.indexOf('ID');
    var iAD   = hOrd.indexOf('ALERTA_SECUENCIA');

    if (iAD < 0) return JSON.stringify({ success: false, msg: 'Columna ALERTA_SECUENCIA no encontrada' });

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iID] || '').trim() === String(idOrden).trim()) {
        shOrd.getRange(i + 1, iAD + 1).setValue(nuevaMP.trim().toUpperCase());
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, msg: 'Orden no encontrada' });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ════════════════════════════════════════════════════════════════════════
// SINCRONIZAR SALIDAS DE PRODUCCIÓN → INV_ALAMBRES
//
// PROPÓSITO:
//   Lee PRODUCCION de los últimos 7 días (desde el 01-Mar-2026 como mínimo),
//   agrupa los KG producidos por fecha + código de MP, y graba/actualiza
//   renglones en INV_ALAMBRES con TIPO=SALIDA, ORIGEN=TRANSFORMACION.
//
// REGLAS:
//   - Solo procesos: ESTAMPADO, FORJA, ROSCADO, ESTIRADO, TREFILADO, COLATADO
//   - Solo alambres: MP cuyo código NO contiene "mm" (el "mm" identifica alambrón)
//   - Si ya existe un renglón para esa fecha+MP (ORIGEN=TRANSFORMACION), 
//     se actualiza el KG en lugar de duplicar
//   - La existencia final nunca baja de 0 (Math.max ya lo garantiza en mrpmpCalcularExistencias)
//
// LLAMADA DESDE: HTML con google.script.run.mrpmpSincronizarSalidasProduccion()
// TRIGGER FUTURO: instalable onEdit o time-based sobre hoja PRODUCCION
// ════════════════════════════════════════════════════════════════════════

function mrpmpSincronizarSalidasProduccion() {
  try {
    var ss     = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shProd = ss.getSheetByName('PRODUCCION');
    var shOrd  = ss.getSheetByName('ORDENES');
    var shRut  = ss.getSheetByName('RUTAS');
    var shInv  = ss.getSheetByName('INV_ALAMBRES');

    if (!shProd || !shOrd || !shInv) {
      return JSON.stringify({ success: false, msg: 'Faltan hojas PRODUCCION, ORDENES o INV_ALAMBRES' });
    }

    var PROCESOS_MP = ['ESTAMPADO','FORJA','ROSCADO','ESTIRADO','TREFILADO','COLATADO'];

    // Fecha mínima: 01-Mar-2026
    var FECHA_CORTE = new Date('2026-03-01T00:00:00');
    // Ventana deslizante: últimos 7 días
    var FECHA_7DIAS = new Date();
    FECHA_7DIAS.setDate(FECHA_7DIAS.getDate() - 7);
    // Usar la más reciente entre el corte fijo y 7 días atrás
    var FECHA_INICIO = FECHA_7DIAS > FECHA_CORTE ? FECHA_7DIAS : FECHA_CORTE;

    // ── 1. Mapa ORDEN → código de MP (desde columna ALERTA_SECUENCIA de ORDENES) ─
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iOrdNom = hOrd.indexOf('ORDEN');
    var iOrdAD  = hOrd.indexOf('ALERTA_SECUENCIA'); // Col AD = código MP asignado
    var iOrdCod = hOrd.indexOf('CODIGO');

    var mapaOrdenMP = {}; // { 'F.1234': { mp: '03061006S', codigo: 'TORNILLO-M6' } }
    for (var o = 1; o < dataOrd.length; o++) {
      var ord = String(dataOrd[o][iOrdNom] || '').trim().toUpperCase();
      var mp  = iOrdAD >= 0 ? String(dataOrd[o][iOrdAD] || '').trim().toUpperCase() : '';
      var cod = String(dataOrd[o][iOrdCod] || '').trim().toUpperCase();
      if (ord && mp) mapaOrdenMP[ord] = { mp: mp, codigo: cod };
    }

    // ── 2. Mapa código MP → descripción (desde RUTAS, columna MP) ───────────────
    var mapaDescMP = {}; // { '03061006S': 'ALAMBRE 06 TREFILADO' }
    if (shRut) {
      var dataRut = shRut.getDataRange().getValues();
      var hRut    = dataRut[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var iRutPro = hRut.indexOf('PROCESO');
      var iRutMP  = hRut.indexOf('MP');
      for (var r = 1; r < dataRut.length; r++) {
        var rProc = String(dataRut[r][iRutPro] || '').trim().toUpperCase();
        if (PROCESOS_MP.indexOf(rProc) < 0) continue;
        var mpCods = String(dataRut[r][iRutMP] || '').trim();
        if (!mpCods) continue;
        var mpPrim = mpCods.split(',')[0].trim().toUpperCase();
        if (mpPrim && !mapaDescMP[mpPrim]) mapaDescMP[mpPrim] = mpPrim; // placeholder
      }
    }

    // ── 3. Enriquecer descripciones desde CIRCULANTE ─────────────────────────────
    try {
      var ssCir  = SpreadsheetApp.openById(ID_CIRCULANTE);
      var shCir  = ssCir.getSheetByName('CIRCULANTE');
      if (shCir) {
        var dataCir = shCir.getDataRange().getValues();
        var hCir    = dataCir[0].map(function(h){ return String(h).toUpperCase().trim(); });
        var iCirCod = hCir.indexOf('CODIGO');      if (iCirCod < 0) iCirCod = 3;
        var iCirDes = hCir.indexOf('DESCRIPCION'); if (iCirDes < 0) iCirDes = 4;
        var iCirMov = hCir.indexOf('MOVIMIENTO');  if (iCirMov < 0) iCirMov = 18;
        for (var ci = 1; ci < dataCir.length; ci++) {
          var movC = String(dataCir[ci][iCirMov] || '').trim().toUpperCase();
          if (movC !== 'ALAMBRES') continue;
          var codC = String(dataCir[ci][iCirCod] || '').trim().toUpperCase();
          var desC = String(dataCir[ci][iCirDes] || '').trim();
          if (codC && desC) mapaDescMP[codC] = desC;
        }
      }
    } catch(ec) {
      Logger.log('mrpmpSincronizar - CIRCULANTE desc: ' + ec.message);
    }

    // ── 4. Leer PRODUCCION y acumular KG por fecha + código MP ──────────────────
    var dataProd  = shProd.getDataRange().getValues();
    var hProd     = dataProd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iProdFec  = hProd.indexOf('FECHA');
    var iProdOrd  = hProd.indexOf('ORDEN');   if (iProdOrd  < 0) iProdOrd  = hProd.indexOf('ID_ORDEN');
    var iProdKg   = hProd.indexOf('PRODUCIDO'); if (iProdKg < 0) iProdKg   = hProd.indexOf('KILOS');
    var iProdProc = hProd.indexOf('PROCESO');
    var iProdPT   = hProd.indexOf('PT');

    // acum: { 'YYYY-MM-DD|COD_MP': { fecha: Date, fechaKey: str, mp: str, kg: num } }
    var acum = {};

    for (var p = 1; p < dataProd.length; p++) {
      // ── Fecha válida y dentro del periodo ──
      var fecP = dataProd[p][iProdFec];
      if (!(fecP instanceof Date)) continue;
      var fecSolo = new Date(fecP.getFullYear(), fecP.getMonth(), fecP.getDate());
      if (fecSolo < FECHA_INICIO) continue;

      // ── Solo procesos que consumen alambre ──
      var procP = iProdProc >= 0 ? String(dataProd[p][iProdProc] || '').trim().toUpperCase() : '';
      if (PROCESOS_MP.indexOf(procP) < 0) continue;

      // ── PT = SI (parte terminada contabilizada) — si no hay columna PT, contar todo ──
      if (iProdPT >= 0) {
        var ptP = String(dataProd[p][iProdPT] || '').trim().toUpperCase();
        if (ptP !== 'SI' && ptP !== '') continue;
      }

      var kgP = Number(dataProd[p][iProdKg] || 0);
      if (kgP <= 0) continue;

      var ordP = String(dataProd[p][iProdOrd] || '').trim().toUpperCase();
      if (!ordP) continue;

      // ── Obtener MP de la orden ──
      var infoOrd = mapaOrdenMP[ordP];
      if (!infoOrd || !infoOrd.mp) continue;
      var mpCod = infoOrd.mp;

      // ── Filtrar alambrón: código con "mm" es alambrón, no alambre ──
      if (mpCod.toLowerCase().indexOf('mm') >= 0) continue;

      // ── Construir clave y acumular ──
      var mm  = fecSolo.getMonth() + 1;
      var dd  = fecSolo.getDate();
      var fechaKey = fecSolo.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
      var clave    = fechaKey + '|' + mpCod;

      if (!acum[clave]) acum[clave] = { fecha: fecSolo, fechaKey: fechaKey, mp: mpCod, kg: 0 };
      acum[clave].kg += kgP;
    }

    // ── 5. Leer INV_ALAMBRES y mapear renglones existentes (SALIDA/TRANSFORMACION) ─
    var dataInv = shInv.getDataRange().getValues();
    var mapaFilasExistentes = {}; // { 'YYYY-MM-DD|COD_MP': numeroFila_1based }

    for (var inv = 1; inv < dataInv.length; inv++) {
      var fInvRaw = dataInv[inv][0];
      if (!(fInvRaw instanceof Date)) continue;
      var tipoInv = String(dataInv[inv][1] || '').trim().toUpperCase();
      var origInv = String(dataInv[inv][2] || '').trim().toUpperCase();
      var mpInv   = String(dataInv[inv][4] || '').trim().toUpperCase();
      if (tipoInv !== 'SALIDA' || origInv !== 'TRANSFORMACION') continue;

      var fInvSolo = new Date(fInvRaw.getFullYear(), fInvRaw.getMonth(), fInvRaw.getDate());
      var mInv = fInvSolo.getMonth() + 1;
      var dInv = fInvSolo.getDate();
      var fInvKey  = fInvSolo.getFullYear() + '-' + (mInv < 10 ? '0' : '') + mInv + '-' + (dInv < 10 ? '0' : '') + dInv;
      mapaFilasExistentes[fInvKey + '|' + mpInv] = inv + 1; // +1 por encabezado
    }

    // ── 6. Insertar renglones nuevos o actualizar los existentes ────────────────
    var insertados = 0, actualizados = 0;
    var claves = Object.keys(acum);

    for (var k = 0; k < claves.length; k++) {
      var item  = acum[claves[k]];
      var kgFin = Math.round(item.kg * 100) / 100;
      var desc  = mapaDescMP[item.mp] || item.mp;
      var fecDate = new Date(item.fecha.getFullYear(), item.fecha.getMonth(), item.fecha.getDate(), 12, 0, 0);

      if (mapaFilasExistentes[claves[k]]) {
        // ── Actualizar solo la columna G (CANTIDAD_KG = col 7) ──
        shInv.getRange(mapaFilasExistentes[claves[k]], 7).setValue(kgFin);
        actualizados++;
      } else {
        // ── Insertar nuevo renglón ──
        shInv.appendRow([
          fecDate,          // A: FECHA
          'SALIDA',         // B: TIPO
          'TRANSFORMACION', // C: ORIGEN
          'PRODUCCION',     // D: REFERENCIA
          item.mp,          // E: MATERIAL  (código MP)
          desc,             // F: CALIBRE   (descripción MP)
          kgFin,            // G: CANTIDAD_KG
          'AUTO-SYNC'       // H: NOTAS
        ]);
        insertados++;
      }
    }

    SpreadsheetApp.flush();
    Logger.log('mrpmpSincronizarSalidasProduccion: ' + insertados + ' insertados, ' + actualizados + ' actualizados, ' + claves.length + ' total');
    return JSON.stringify({
      success:      true,
      insertados:   insertados,
      actualizados: actualizados,
      total:        claves.length
    });

  } catch(e) {
    Logger.log('mrpmpSincronizarSalidasProduccion ERROR: ' + e.message);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA GeneradorOrdenesHTML (MENU GENERADOR ORDENES)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function obtenerDatosGenerador() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetPed = ss.getSheetByName("PEDIDOS");
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetRut = ss.getSheetByName("RUTAS");
  var dataRut  = sheetRut.getDataRange().getValues();

  var dataPed = sheetPed.getDataRange().getValues();
  var dataOrd = sheetOrd.getDataRange().getValues();
  
  // 1. Mapeo para sumar cantidades de ORDENES (evitando duplicar procesos)
  // La clave incluye FOLIO+PARTIDA para no mezclar partidas del mismo pedido
  var planeadoMap = {};
  for(var i=1; i<dataOrd.length; i++){
    var folioPed  = String(dataOrd[i][1]); // Col B: PEDIDO
    var partidaOrd = String(dataOrd[i][2] || '1'); // Col C: PARTIDA
    var serie = dataOrd[i][4];             // Col E: SERIE
    var nOrd = dataOrd[i][5];              // Col F: ORDEN
    var mapKey = folioPed + "||" + partidaOrd; // clave única por partida
    var keyUnicaOrden = mapKey + "_" + serie + "." + nOrd;
    var estadoOrd = String(dataOrd[i][15]); // Col P: ESTADO

    // Solo sumar si NO está CANCELADO ni TERMINADO
    if(estadoOrd !== "CANCELADO" && estadoOrd !== "TERMINADO"){
      if(!planeadoMap[mapKey]) planeadoMap[mapKey] = { keys: {}, total: 0 };

      // Si esta combinación de Serie.Orden no ha sido sumada para esta partida
      if(!planeadoMap[mapKey].keys[keyUnicaOrden]){
        planeadoMap[mapKey].total += (Number(dataOrd[i][8]) || 0); // Col I: CANTIDAD
        planeadoMap[mapKey].keys[keyUnicaOrden] = true;
      }
    }
  }

  // 2. Preparar datos y actualizar Columna P en la hoja PEDIDOS
  var lista = [];
  var actualizacionesP = []; // Para escribir en batch si fuera necesario, o uno a uno
  
  for(var j=1; j<dataPed.length; j++){
    var folio = String(dataPed[j][1]);
    var partidaPed = String(dataPed[j][5] || '1'); // Col F: PARTIDA
    var mapKeyPed = folio + "||" + partidaPed;
    var sumaPlaneado = planeadoMap[mapKeyPed] ? planeadoMap[mapKeyPed].total : 0;

    // Actualizamos físicamente la Columna P (índice 15) en la hoja si es diferente
    if(dataPed[j][15] !== sumaPlaneado) {
      sheetPed.getRange(j + 1, 16).setValue(sumaPlaneado);
    }

    lista.push({
  id: String(dataPed[j][0]),
  folio: folio,
  partida: dataPed[j][5],
  fecha: dataPed[j][2] instanceof Date ? Utilities.formatDate(dataPed[j][2], Session.getScriptTimeZone(), "dd/MM/yyyy") : dataPed[j][2],
  codigo: dataPed[j][3],
  descripcion: dataPed[j][4],
  cantidad: dataPed[j][6],
  unidad: dataPed[j][7],
  estado: dataPed[j][8],
  planeado: sumaPlaneado,
  fila: j + 1,
  serie: (function() {
    for (var ri = 1; ri < dataRut.length; ri++) {
      if (String(dataRut[ri][1]).trim() === String(dataPed[j][3]).trim()) {
        return String(dataRut[ri][14] || '').trim().toUpperCase();
      }
    }
    return '';
  })()
});
  }
  return lista;
}

/**
 * PROCESA LOS CAMBIOS EN BATCH (Generar, Cancelar, Activar)
 */
function ejecutarCambiosGenerador(cambios) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetPed = ss.getSheetByName("PEDIDOS");
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetLot = ss.getSheetByName("LOTES");
  var sheetRut = ss.getSheetByName("RUTAS");

  if (!sheetPed || !sheetOrd || !sheetLot || !sheetRut) {
    return "Error: No se encontró una de las hojas (PEDIDOS, ORDENES, LOTES o RUTAS)";
  }

  var dataRut = sheetRut.getDataRange().getValues();
  var lock = LockService.getScriptLock();

  // Procesos que consumen materia prima (guardar en Col AD)
  var PROCS_MP = ['ESTAMPADO','FORJA','ROSCADO','ESTIRADO','TREFILADO','COLATADO'];
  
  try {
    lock.waitLock(25000);
    
    cambios.forEach(c => {
      // 1. CANCELAR
      if(c.accion === "CANCELAR"){
        var estadoCancelable = String(sheetPed.getRange(c.fila, 9).getValue());
        sheetPed.getRange(c.fila, 9).setValue("CANCELADO");
        var dOrd = sheetOrd.getDataRange().getValues();
        var ordenesAfectadas = 0;
        for(var o=1; o<dOrd.length; o++){
          if(String(dOrd[o][1]) === String(c.folio)){
            sheetOrd.getRange(o+1, 16).setValue("CANCELADO");
            ordenesAfectadas++;
          }
        }
        Logger.log("CANCELAR pedido " + c.folio + " (estado anterior: " + estadoCancelable + ") — órdenes canceladas: " + ordenesAfectadas);
      }
      
      // 2. ACTIVAR
      else if(c.accion === "ACTIVAR") {
        sheetPed.getRange(c.fila, 9).setValue("PLANEADO");
      }
      
      // 3. GENERAR
      else if(c.accion === "GENERAR"){
        var ruta = dataRut.filter(r => String(r[1]).trim() === String(c.codigo).trim());
        if(ruta.length === 0) throw new Error("Sin ruta para código: " + c.codigo + " (Pedido: " + c.folio + "). Verifica que el código exista en la hoja RUTAS.");

        var serie = String(ruta[0][14]); // Col O de RUTAS
        var nOrden = getSiguienteNumeroOrden(serie);
        var firstProcessId = "";
        var solicitadoBase = 0;

        for(var n=0; n < c.nOrdenes; n++){
          var nOrdenIndividual = (n === 0) ? nOrden : getSiguienteNumeroOrden(serie);
          
          ruta.forEach((paso, idx) => {
            var idProc = Utilities.getUuid().substring(0,8);
            if(idx === 0) {
              firstProcessId = idProc;
            }

            // Lógica de Conversión
            var peso = paso[15]; // Col P de RUTAS (PESO)
            var unidadPedido = String(c.unidad).toUpperCase().trim();
            var descripcionPedido = String(c.descripcion).toUpperCase();
            var especificaciones = String(paso[8]); // Col I de RUTAS
            var solicitado = c.cantPorOrden;
            
            if(["PZA", "CTO"].includes(unidadPedido)){
              solicitado = c.cantPorOrden * peso;
              if(descripcionPedido.includes("VARILLA")){
                var match = especificaciones.match(/([\d\.]+)/);
                var longitud = match ? parseFloat(match[0]) : 1;
                solicitado = c.cantPorOrden * peso * longitud;
              }
            }
            
            if(idx === 0) solicitadoBase = solicitado;

            var hoy = new Date(); hoy.setHours(0,0,0,0);

            // Determinar MP del paso (Col R de RUTAS = índice 17)
            var procPaso = String(paso[4] || '').trim().toUpperCase();
            var mpVal    = String(paso[17] || '').trim(); // Col R: MP (puede tener varias separadas por coma)
            var mpPaso   = (PROCS_MP.indexOf(procPaso) >= 0 && mpVal)
                           ? mpVal.split(',')[0].trim().toUpperCase()
                           : '';

            // Insertar en ORDENES
            sheetOrd.appendRow([
              idProc, c.folio, c.partida, hoy, serie, nOrdenIndividual, "'" + String(c.codigo).padStart(9, '0'), paso[2],
              c.cantPorOrden, c.unidad, paso[3], paso[4], String(paso[5]).split(",")[0].trim(),
              solicitado, 0, "ABIERTO",
              paso[12], // Q: PT
              paso[13], // R: VENTA
              paso[15], // S: PESO
              paso[6], "'" + String(paso[7]), "'" + String(paso[8]), paso[9], paso[10], paso[11], new Date(), // T a Z (25)
              "",       // AA=26: PRIORIDAD
              "",       // AB=27: FECHA_INICIO_PROG
              "",       // AC=28: FECHA_FIN_PROG
              mpPaso    // AD=29: ALERTA_SECUENCIA = MP
            ]);
          });

          // GENERAR LOTES
          var cantLoteBase = parseFloat(ruta[0][22]) || solicitadoBase;
          var numLotes = Math.ceil(solicitadoBase / cantLoteBase);
          for(var l=1; l<=numLotes; l++){
            var nombreLote = serie + "." + ("0000"+nOrdenIndividual).slice(-4) + "." + (l < 100 ? ("00"+l).slice(-2) : String(l));
            sheetLot.appendRow([
              Utilities.getUuid().substring(0,8), serie, firstProcessId, l,
              nombreLote, cantLoteBase, 0, "ABIERTO", new Date()
            ]);
          }
        }

        // 4. Actualizar estado del pedido
        var cellEstado = sheetPed.getRange(c.fila, 9);
        var estadoActual = cellEstado.getValue();
        if(["ABIERTO", "PLANEADO", "PARCIAL", "EN PROCESO"].includes(estadoActual)){
          cellEstado.setValue(c.nuevoEstadoPedido);
        }
      }
    });
    return "OK";
  } catch(e) { return "Error: " + e.toString(); } 
  finally { lock.releaseLock(); }
}

/**
 * OBTIENE EL DETALLE DE LAS ÓRDENES RELACIONADAS A UN PEDIDO
 */
function obtenerDetalleOrdenes(folioPedido, partidaParam) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var data = sheetOrd.getDataRange().getValues();
  var filtrarPartida = (partidaParam !== undefined && partidaParam !== null && partidaParam !== '');

  var tempMap = {}; // Para agrupar procesos por Serie.Orden

  for(var i=1; i<data.length; i++) {
    var folioEnHoja = String(data[i][1]); // Col B
    var partidaEnHoja = String(data[i][2] || '1'); // Col C: PARTIDA
    var estado = String(data[i][15]);     // Col P

    var matchPartida = !filtrarPartida || (partidaEnHoja === String(partidaParam));
    if(folioEnHoja === String(folioPedido) && matchPartida && estado !== "CANCELADO") {
      var serie = data[i][4];  // Col E
      var num = data[i][5];    // Col F
      var key = serie + "." + num;
      
      if(!tempMap[key]) {
        tempMap[key] = {
          serie: serie,
          numero: num,
          cantidad: data[i][8],   // Col I
          maquina: data[i][12],  // Col M
          solicitado: data[i][13], // Col N
          producido: data[i][14]   // Col O
        };
      }
    }
  }
  return Object.values(tempMap);
}

/**
 * CANCELA UNA ORDEN ESPECÍFICA (TODOS SUS PROCESOS) Y ACTUALIZA PLANEACIÓN EN PEDIDOS
 */
function cancelarOrdenEspecifica(folioPedido, serie, nOrden) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetPed = ss.getSheetByName("PEDIDOS");
  
  var dataOrd = sheetOrd.getDataRange().getValues();
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(15000);
    
    // 1. Cancelar en ORDENES
    for(var i=1; i<dataOrd.length; i++) {
      if(String(dataOrd[i][1]) === String(folioPedido) && 
         String(dataOrd[i][4]) === String(serie) && 
         String(dataOrd[i][5]) === String(nOrden)) {
         sheetOrd.getRange(i+1, 16).setValue("CANCELADO"); // Col P
      }
    }
    
    // 2. Recalcular la suma planeada para el pedido y actualizar columna P (indice 16)
    // Volvemos a leer datos frescos de órdenes
    var dataOrdNueva = sheetOrd.getDataRange().getValues();
    var planeadoTotal = 0;
    var keysContadas = {};
    
    for(var j=1; j<dataOrdNueva.length; j++) {
      if(String(dataOrdNueva[j][1]) === String(folioPedido) && String(dataOrdNueva[j][15]) !== "CANCELADO") {
         var key = dataOrdNueva[j][4] + "." + dataOrdNueva[j][5];
         if(!keysContadas[key]) {
           planeadoTotal += Number(dataOrdNueva[j][8]);
           keysContadas[key] = true;
         }
      }
    }
    
    // 3. Buscar la fila del pedido para actualizar la columna P (CANT_PLAN)
    var dataPed = sheetPed.getDataRange().getValues();
    for(var k=1; k<dataPed.length; k++) {
      if(String(dataPed[k][1]) === String(folioPedido)) {
        sheetPed.getRange(k+1, 16).setValue(planeadoTotal); // Columna P
        
        // Si el planeado bajó, quizá deba regresar a "PARCIAL" o "ABIERTO" si no hay nada
        var estadoActual = dataPed[k][8];
        if(estadoActual === "PLANEADO" && planeadoTotal < dataPed[k][6]) {
          sheetPed.getRange(k+1, 9).setValue("PARCIAL");
        } else if (planeadoTotal === 0) {
          sheetPed.getRange(k+1, 9).setValue("ABIERTO");
        }
        break;
      }
    }
    
    return "OK";
  } catch(e) {
    return e.toString();
  } finally {
    lock.releaseLock();
  }
}

// =================================================================================
// ================= PEDIDOS TEM (ReemplazoPedidosHTML) ============================
// =================================================================================

// 1. OBTENER PEDIDOS "TEM"
function obtenerPedidosTEM() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetPed = ss.getSheetByName("PEDIDOS");
    var data = sheetPed.getDataRange().getValues();
    
    // Indices: A=0(ID), B=1(Pedido), C=2(Fecha), D=3(Codigo), E=4(Desc), F=5(Partida), G=6(Cant), H=7(Unidad)
    var lista = [];
    var timeZone = Session.getScriptTimeZone();
    
    for(var i=1; i<data.length; i++) {
       // Validación para evitar filas vacías
       if (!data[i][0]) continue;

       var pedido = String(data[i][1]).toUpperCase().trim();
       var estado = String(data[i][8]).toUpperCase(); // Col I Estado

       // FILTRO: Contiene "TEM" y VIVO
       if(pedido.includes("TEM") && estado != "CANCELADO" && estado != "CERRADO" && estado != "TERMINADO") {
          
          // Formatear fecha para evitar bloqueo
          var fechaRaw = data[i][2];
          var fechaStr = "";
          if (fechaRaw instanceof Date) {
             fechaStr = Utilities.formatDate(fechaRaw, timeZone, "yyyy-MM-dd");
          } else {
             fechaStr = String(fechaRaw);
          }

          lista.push({
             id: data[i][0],      
             fila: i + 1,
             pedido: data[i][1],  
             fecha: fechaStr,     // <--- FECHA COMO TEXTO SEGURO
             codigo: data[i][3],  
             desc: data[i][4],    
             partida: data[i][5], 
             cant: Number(data[i][6]), 
             unidad: String(data[i][7]).toUpperCase() 
          });
       }
    }
    
    // Retornamos JSON para asegurar transmisión
    return JSON.stringify(lista);

  } catch (e) {
    // Si falla, retornamos el error para verlo en el HTML
    return JSON.stringify({ "error": e.toString() });
  }
}

// 2. GUARDAR CAMBIOS (CASCADA PEDIDOS -> ORDENES -> ENVIADO)
function guardarCambiosTEM(cambios) {
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(10000)) return "Error: Servidor ocupado.";

  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetPed = ss.getSheetByName("PEDIDOS");
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetEnv = ss.getSheetByName("ENVIADO"); // <--- Nueva referencia
    
    // --- 1. ACTUALIZAR HOJA "PEDIDOS" ---
    var dataPed = sheetPed.getDataRange().getValues();
    var mapPedFilas = {};
    for(var p=1; p<dataPed.length; p++) mapPedFilas[String(dataPed[p][0])] = p+1;

    cambios.forEach(c => {
       var fila = mapPedFilas[c.id];
       if(fila) {
          sheetPed.getRange(fila, 2).setValue(c.newPedido); // Col B: Pedido
          sheetPed.getRange(fila, 7).setValue(c.newCant);   // Col G: Cantidad
       }
    });

    // --- 2. PREPARAR MAPEO DE CAMBIOS PARA BUSQUEDA ---
    // Usaremos un mapa simple de NombreViejo -> NombreNuevo para las hojas masivas
    var mapNombres = {}; 
    var mapCambiosKey = {}; // Para la lógica de Ordenes (Pedido + Partida)
    
    cambios.forEach(c => {
       var oldPed = String(c.originalPedido).trim();
       mapNombres[oldPed] = c.newPedido;
       
       var key = oldPed + "_" + c.partida;
       mapCambiosKey[key] = c;
    });

    // --- 3. ACTUALIZAR HOJA "ORDENES" ---
    var dataOrd = sheetOrd.getDataRange().getValues();
    for(var o=1; o<dataOrd.length; o++) {
       var pedOrd = String(dataOrd[o][1]).trim(); // Col B
       var partOrd = dataOrd[o][2];               // Col C
       var keyOrd = pedOrd + "_" + partOrd;

       if(mapCambiosKey[keyOrd]) {
           var cambio = mapCambiosKey[keyOrd];
           var filaReal = o + 1;
           sheetOrd.getRange(filaReal, 2).setValue(cambio.newPedido);
           sheetOrd.getRange(filaReal, 9).setValue(cambio.newCant);
           
           var unidad = cambio.unidad;
           var peso = Number(dataOrd[o][18]) || 0; // Col S
           if(unidad.includes("PZA") || unidad.includes("CTO") || unidad.includes("PIEZA")) {
              if(peso > 0) {
                 var nuevoSolicitado = Math.round((cambio.newCant * peso) * 100) / 100;
                 sheetOrd.getRange(filaReal, 14).setValue(nuevoSolicitado);
              }
           } else {
              sheetOrd.getRange(filaReal, 14).setValue(cambio.newCant);
           }
       }
    }

    // --- 4. ACTUALIZAR HOJA "ENVIADO" (NUEVO) ---
    if (sheetEnv) {
      var dataEnv = sheetEnv.getDataRange().getValues();
      // Columnas ENVIADO: A=0, B=1, C=2, D=3, E=4, F=5 (PEDIDO)
      for (var e = 1; e < dataEnv.length; e++) {
        var pedidoEnviado = String(dataEnv[e][5]).trim(); // Columna F
        
        if (mapNombres[pedidoEnviado]) {
          // Si el pedido en esta fila de ENVIADO coincide con un TEM que estamos cambiando
          sheetEnv.getRange(e + 1, 6).setValue(mapNombres[pedidoEnviado]); 
        }
      }
    }

    SpreadsheetApp.flush();
    return "✅ Actualización Correcta: " + cambios.length + " pedidos procesados en Pedidos, Ordenes y Enviados.";

  } catch(e) {
    return "❌ Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA PlanificadorHTML (MENU PLANIFICADOR)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function actualizarAceroFamilia(id, nuevoAcero) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sh = ss.getSheetByName("ORDENES");
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(h => String(h).toUpperCase().trim());
  
  var colId = 0;
  var colSerie = headers.indexOf("SERIE"); // Col E
  var colOrden = headers.indexOf("ORDEN"); // Col F
  var colAcero = headers.indexOf("ACERO"); // Col Y

  var serieTarget = "";
  var ordenTarget = "";

  // 1. Buscamos primero la Serie y Orden de la fila que el usuario clickeó
  for(var i=1; i<data.length; i++){
    if(String(data[i][colId]) === String(id)){
      serieTarget = String(data[i][colSerie]);
      ordenTarget = String(data[i][colOrden]);
      break;
    }
  }

  // 2. Actualizamos TODAS las filas que coincidan con esa familia
  if(serieTarget && ordenTarget) {
    for(var j=1; j<data.length; j++){
      if(String(data[j][colSerie]) === serieTarget && String(data[j][colOrden]) === ordenTarget){
        sh.getRange(j+1, colAcero+1).setValue(nuevoAcero);
      }
    }
  }

  // Devolvemos el nuevo SVG para actualizar la interfaz
  return getSvgAcero(nuevoAcero);
}

function obtenerProduccionPlanificador(filtros) {
  // Filtros trae { procesos, fechaIni, fechaFin }
  var res = buscarProduccion({
    maquinas: [], // Buscamos por procesos seleccionados, no por máquinas fijas
    fechaIni: filtros.fechaIni,
    fechaFin: filtros.fechaFin
  });
  // Filtramos por procesos seleccionados
  var filtrados = res.filter(r => filtros.procesos.includes(r.proceso));
  return JSON.stringify(filtrados);
}

// Obtiene procesos y máquinas para el inicio
function obtenerCatalogosPlanificador() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("ESTANDARES");
  var data = sheet.getDataRange().getValues();
  var catalogo = {}; // { PROCESO: [MAQUINA1, MAQUINA2] }
  
  for(var i=1; i<data.length; i++){
    var p = String(data[i][2]).toUpperCase().trim(); // Proceso
    var m = String(data[i][3]).trim(); // Maquina
    if(p && m){
      if(!catalogo[p]) catalogo[p] = [];
      if(!catalogo[p].includes(m)) catalogo[p].push(m);
    }
  }
  return catalogo;
}

function obtenerOrdenesPlanificador(procesosSeleccionados) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd   = ss.getSheetByName("ORDENES");
    var shRutas = ss.getSheetByName("RUTAS");
    var data    = shOrd.getDataRange().getValues();
    var headers = data[0].map(h => String(h).toUpperCase().trim());

    var getIdx = function(n) { return headers.indexOf(n); };
    var idx = {
      ID:    getIdx("ID"),    ESTADO: getIdx("ESTADO"), PROC:   getIdx("PROCESO"),
      SOL:   getIdx("SOLICITADO"), PROD: getIdx("PRODUCIDO"), MAQ: getIdx("MAQUINA"),
      COD:   getIdx("CODIGO"),    ORD:  getIdx("ORDEN"),      PED:  getIdx("PEDIDO"),
      DESC:  getIdx("DESCRIPCION"), PRIO: getIdx("PRIORIDAD"), DIA: getIdx("DIAMETRO"),
      LONG:  getIdx("LONGITUD"),  CUERDA: getIdx("CUERDA"),   CUERPO: getIdx("CUERPO"),
      ACERO: getIdx("ACERO"),     TIPO: getIdx("TIPO"),       SERIE: getIdx("SERIE"),
      MP:    getIdx("ALERTA_SECUENCIA")   // ← Col AD
    };

    var dRutas = shRutas.getDataRange().getValues();
    var maqPermitidasMap = {};
    for (var r = 1; r < dRutas.length; r++) {
      maqPermitidasMap[String(dRutas[r][1]).trim() + "_" + String(dRutas[r][4]).trim().toUpperCase()] = String(dRutas[r][5] || "");
    }

    // Cargar INVENTARIO_EXTERNO: col A=CODIGO, B=EXISTENCIA, C=MINIMO, D=MAXIMO, buscar UNIDAD por header
    var shInvExt = ss.getSheetByName("INVENTARIO_EXTERNO");
    var invExtMap = {};
    if (shInvExt) {
      var dInvExt = shInvExt.getDataRange().getValues();
      var hInvExt = dInvExt[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var ieUNID  = hInvExt.indexOf("UNIDAD"); if (ieUNID < 0) ieUNID = hInvExt.indexOf("UNID");
      for (var ie = 1; ie < dInvExt.length; ie++) {
        var invCod = String(dInvExt[ie][0] || "").trim().toUpperCase();
        if (!invCod) continue;
        invExtMap[invCod] = {
          exist:  parseFloat(dInvExt[ie][1]) || 0,
          min:    parseFloat(dInvExt[ie][2]) || 0,
          max:    parseFloat(dInvExt[ie][3]) || 0,
          back:   parseFloat(dInvExt[ie][4]) || 0,
          unidad: ieUNID >= 0 ? String(dInvExt[ie][ieUNID] || "").trim().toUpperCase() : ""
        };
      }
    }

    // Cargar CODIGOS: col A=CODIGO, F=CODIGO_VENTA (índice 5)
    var shCodigos = ss.getSheetByName("CODIGOS");
    var codigoVentaMap = {};
    if (shCodigos) {
      var dCodigos = shCodigos.getDataRange().getValues();
      for (var cv = 1; cv < dCodigos.length; cv++) {
        var cvCod = String(dCodigos[cv][0] || "").trim().toUpperCase();
        var cvVenta = String(dCodigos[cv][5] || "").trim().toUpperCase();
        if (cvCod && cvVenta) codigoVentaMap[cvCod] = cvVenta;
      }
    }

    // Función helper para obtener datos de inventario por código de orden
    var getInvExt = function(codigoOrden) {
      var cUp = String(codigoOrden || "").trim().toUpperCase();
      var esC7 = cUp.charAt(0) === "7";
      if (esC7) {
        var cVenta = codigoVentaMap[cUp] || "";
        // Varilla con código_venta diferente: usa datos del código venta + existNeg del código negro
        if (cVenta && cVenta !== cUp) {
          var datoVenta    = invExtMap[cVenta] || null;
          var datoOriginal = invExtMap[cUp]    || null;
          if (datoVenta) {
            return {
              exist:    datoVenta.exist,
              min:      datoVenta.min,
              max:      datoVenta.max,
              back:     datoVenta.back,
              existNeg: datoOriginal ? datoOriginal.exist : 0,
              esVarilla: true,
              codigoVenta: cVenta
            };
          }
        }
        // Varilla con codigo = codigo_venta (sin código negro diferente): solo datos del propio código
        if (cVenta && cVenta === cUp) {
          var datoPropio = invExtMap[cUp] || null;
          return datoPropio
            ? { exist: datoPropio.exist, min: datoPropio.min, max: datoPropio.max, back: datoPropio.back, existNeg: null, esVarilla: true, codigoVenta: cUp }
            : { exist: null, min: null, max: null, back: 0, existNeg: null, esVarilla: false, codigoVenta: '' };
        }
        // Código 7 sin entrada en CODIGOS: busca directo
        var datoDirecto = invExtMap[cUp] || null;
        return datoDirecto
          ? { exist: datoDirecto.exist, min: datoDirecto.min, max: datoDirecto.max, back: datoDirecto.back, existNeg: datoDirecto.exist, esVarilla: false, codigoVenta: '' }
          : { exist: null, min: null, max: null, back: 0, existNeg: null, esVarilla: false, codigoVenta: '' };
      }
      var dato = invExtMap[cUp];
      return dato
        ? { exist: dato.exist, min: dato.min, max: dato.max, back: dato.back, existNeg: null, esVarilla: false, codigoVenta: '' }
        : { exist: null, min: null, max: null, back: 0, existNeg: null, esVarilla: false, codigoVenta: '' };
    };

    var planif_invTmp = null;
    var resultados = [];
    var procsFiltro = procesosSeleccionados.map(p => String(p).toUpperCase().trim());

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (procsFiltro.includes(String(row[idx.PROC]).toUpperCase().trim())) {
        var est = String(row[idx.ESTADO]).toUpperCase().trim();
        if (est !== "TERMINADO" && est !== "CANCELADO") {
          var tipoVal  = String(row[idx.TIPO]  || "ESP");
          var aceroVal = String(row[idx.ACERO] || "-");
          var mpVal    = idx.MP >= 0 ? String(row[idx.MP] || "").trim().toUpperCase() : "";
          resultados.push({
            id:                String(row[idx.ID]),
            peso:              parseFloat(row[18]) || 0,
            pedido:            String(row[idx.PED] || ""),
            orden:             String(row[idx.ORD] || ""),
            codigo:            String(row[idx.COD] || ""),
            desc:              String(row[idx.DESC] || ""),
            sol:               parseFloat(row[idx.SOL]) || 0,
            prod:              parseFloat(row[idx.PROD]) || 0,
            estado:            est,
            proceso:           String(row[idx.PROC]).toUpperCase().trim(),
            maquina:           String(row[idx.MAQ]).trim(),
            prioridad:         parseInt(row[idx.PRIO]) || 999,
            dia:               String(row[idx.DIA] || ""),
            long:              String(row[idx.LONG] || ""),
            cuerda:            String(row[idx.CUERDA] || ""),
            cuerpo:            String(row[idx.CUERPO] || ""),
            acero:             aceroVal,
            tipo:              tipoVal,
            serie:             String(row[idx.SERIE] || ""),
            mp:                mpVal,   // ← NUEVO
            iconoSVG:          obtenerIconoSVG(tipoVal),
            estadoSVG:         getSvgEstado(est),
            aceroSVG:          getSvgAcero(aceroVal),
            avance:            row[idx.SOL] > 0 ? Math.round((row[idx.PROD] / row[idx.SOL]) * 100) : 0,
            maquinasPermitidas: maqPermitidasMap[String(row[idx.COD]) + "_" + String(row[idx.PROC]).toUpperCase().trim()] || String(row[idx.MAQ]),
            invExist:    (function(){ var _d=getInvExt(row[idx.COD]); planif_invTmp=_d; return _d.exist; })(),
            invMin:      planif_invTmp ? planif_invTmp.min : null,
            invMax:      planif_invTmp ? planif_invTmp.max : null,
            invExistNeg: planif_invTmp ? (planif_invTmp.existNeg !== undefined ? planif_invTmp.existNeg : null) : null,
            invBack:      planif_invTmp ? (planif_invTmp.back !== undefined ? planif_invTmp.back : 0) : 0,
            invUnidad:    planif_invTmp ? (planif_invTmp.unidad || "") : "",
            invEsVarilla: planif_invTmp ? (planif_invTmp.esVarilla === true) : false
          });
        }
      }
    }
    return JSON.stringify(resultados);
  } catch(e) { return JSON.stringify([{ error: true, mensaje: e.toString() }]); }
}

// Función para el historial
function obtenerHistorialMaquina(maquina) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("ORDENES");
    var shRutas = ss.getSheetByName("RUTAS");
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(h => String(h).toUpperCase().trim());
    
    var fLimite = new Date(); fLimite.setDate(fLimite.getDate() - 90);
    var getIdx = function(n) { return headers.indexOf(n); };

    var idx = {
      ID: 0, ORD: getIdx("ORDEN"), COD: getIdx("CODIGO"), DESC: getIdx("DESCRIPCION"),
      EST: getIdx("ESTADO"), SOL: getIdx("SOLICITADO"), PROD: getIdx("PRODUCIDO"),
      TIPO: getIdx("TIPO"), ACERO: getIdx("ACERO"), DIA: getIdx("DIAMETRO"),
      LONG: getIdx("LONGITUD"), CUERDA: getIdx("CUERDA"), CUERPO: getIdx("CUERPO"),
      PED: getIdx("PEDIDO"), PROC: getIdx("PROCESO"), MAQ: getIdx("MAQUINA"),
      FECHA: headers.indexOf("FECHA_REGISTRO")
    };
    if(idx.FECHA == -1) idx.FECHA = headers.indexOf("FECHA");

    var dRutas = shRutas.getDataRange().getValues();
    var maqPermitidasMap = {}; 
    for(var r=1; r<dRutas.length; r++){
      maqPermitidasMap[String(dRutas[r][1]).trim() + "_" + String(dRutas[r][4]).trim().toUpperCase()] = String(dRutas[r][5] || ""); 
    }

    var res = [];
    for(var i=1; i<data.length; i++){
      var fVal = (data[i][idx.FECHA] instanceof Date) ? data[i][idx.FECHA] : new Date();
      if(String(data[i][idx.MAQ]).trim() === maquina && fVal >= fLimite){
        var t = String(data[i][idx.TIPO] || "ESP");
        var a = String(data[i][idx.ACERO] || "-");
        var s = parseFloat(data[i][idx.SOL]) || 0;
        var p = parseFloat(data[i][idx.PROD]) || 0;
        var est = String(data[i][idx.EST]).toUpperCase().trim();
        var cod = String(data[i][idx.COD]);
        var prc = String(data[i][idx.PROC]).toUpperCase().trim();

        res.push({
          id: String(data[i][idx.ID]),
          serie: String(data[i][4]),
          orden: String(data[i][idx.ORD]),
          codigo: cod,
          desc: String(data[i][idx.DESC]),
          estado: est,
          sol: s, prod: p, tipo: t, acero: a,
          dia: String(data[i][idx.DIA]),
          long: String(data[i][idx.LONG]),
          cuerda: String(data[i][idx.CUERDA]),
          cuerpo: String(data[i][idx.CUERPO]),
          pedido: String(data[i][idx.PED]),
          proceso: prc,
          maquina: maquina,
          avance: s > 0 ? Math.round((p/s)*100) : 0,
          iconoSVG: obtenerIconoSVG(t),
          estadoSVG: getSvgEstado(est),
          aceroSVG: getSvgAcero(a),
          maquinasPermitidas: maqPermitidasMap[cod + "_" + prc] || maquina
        });
      }
    }
    return JSON.stringify(res);
  } catch(e) { return JSON.stringify({error: e.toString()}); }
}

function guardarPlanificacion(listaCambios) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sh = ss.getSheetByName("ORDENES");
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });

  var colId   = 0;
  var colMaq  = headers.indexOf("MAQUINA");
  var colPrio = headers.indexOf("PRIORIDAD");
  var colProc = headers.indexOf("PROCESO");

  // 1. Aplicar cambios de máquina y prioridad en el array en memoria
  var mapaCAmbios = {};
  (listaCambios || []).forEach(function(c){ mapaCAmbios[String(c.id)] = c; });

  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][colId]);
    if (mapaCAmbios[id]) {
      data[i][colMaq]  = mapaCAmbios[id].maquinaNueva;
      data[i][colPrio] = mapaCAmbios[id].prioridadNueva;
    }
  }

  // 2. Reordenar prioridades por grupo máquina+proceso en memoria
  var grupos = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][colMaq]).trim() + "_" + String(data[i][colProc]).trim();
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push({ fila: i, prio: parseInt(data[i][colPrio]) || 999 });
  }
  Object.keys(grupos).forEach(function(key) {
    grupos[key].sort(function(a, b){ return a.prio - b.prio; });
    grupos[key].forEach(function(item, index){
      data[item.fila][colPrio] = index + 1;
    });
  });

  // 3. UNA SOLA escritura batch de toda la columna PRIORIDAD
  var prioBatch = data.slice(1).map(function(row){ return [row[colPrio]]; });
  sh.getRange(2, colPrio + 1, prioBatch.length, 1).setValues(prioBatch);

  // 4. Si hubo cambios de máquina, escribir también esa columna en batch
  if (listaCambios && listaCambios.some(function(c){ return c.maquinaNueva !== c.maquinaAnterior; })) {
    var maqBatch = data.slice(1).map(function(row){ return [row[colMaq]]; });
    sh.getRange(2, colMaq + 1, maqBatch.length, 1).setValues(maqBatch);
  }

  SpreadsheetApp.flush();
  return true;
}

function terminarPedidoColatado(folioPedido) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shPed = ss.getSheetByName("PEDIDOS");
    var shOrd = ss.getSheetByName("ORDENES");
    var dataPed = shPed.getDataRange().getValues();
    var dataOrd = shOrd.getDataRange().getValues();
    var hPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var hOrd = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdFOL = hPed.indexOf("FOLIO"); if (pdFOL === -1) pdFOL = 1;
    var pdEST = hPed.indexOf("ESTADO"); if (pdEST === -1) pdEST = 8;
    var oID   = hOrd.indexOf("ID");
    var oPED  = hOrd.indexOf("PEDIDO");
    var oEST  = hOrd.indexOf("ESTADO");
    var ordenesCanceladas = 0, ordenesTerminadas = 0;
    // 1. Pasar pedido a TERMINADO
    for (var p = 1; p < dataPed.length; p++) {
      if (String(dataPed[p][pdFOL]).trim() === String(folioPedido).trim()) {
        shPed.getRange(p+1, pdEST+1).setValue('TERMINADO');
        break;
      }
    }
    // 2. Pasar todas sus órdenes a TERMINADO excepto las CANCELADAS
    for (var o = 1; o < dataOrd.length; o++) {
      if (String(dataOrd[o][oPED]).trim() !== String(folioPedido).trim()) continue;
      var estActual = String(dataOrd[o][oEST]||'').toUpperCase().trim();
      if (estActual === 'CANCELADO') { ordenesCanceladas++; continue; }
      shOrd.getRange(o+1, oEST+1).setValue('TERMINADO');
      ordenesTerminadas++;
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, msg: ordenesTerminadas + ' orden(es) terminadas, ' + ordenesCanceladas + ' canceladas conservadas.' });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

function obtenerDetallePedidoAlerta(codigo) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shPed = ss.getSheetByName("PEDIDOS");
    var shOrd = ss.getSheetByName("ORDENES");
    var dataPed = shPed.getDataRange().getValues();
    var dataOrd = shOrd.getDataRange().getValues();
    var hPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var hOrd = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdFOL  = hPed.indexOf("FOLIO");       if (pdFOL  < 0) pdFOL  = 1;
    var pdCOD  = hPed.indexOf("CODIGO");      if (pdCOD  < 0) pdCOD  = 3;
    var pdEST  = hPed.indexOf("ESTADO");      if (pdEST  < 0) pdEST  = 8;
    var pdDESC = hPed.indexOf("DESCRIPCION"); if (pdDESC < 0) pdDESC = 4;
    var pdCAN  = hPed.indexOf("CANTIDAD");    if (pdCAN  < 0) pdCAN  = 6;
    var oPED   = hOrd.indexOf("PEDIDO");
    var oCOD   = hOrd.indexOf("CODIGO");
    var oEST   = hOrd.indexOf("ESTADO");
    var oPROC  = hOrd.indexOf("PROCESO");
    var oSOL   = hOrd.indexOf("SOLICITADO");
    var oPROD  = hOrd.indexOf("PRODUCIDO");
    var oMAQ   = hOrd.indexOf("MAQUINA");
    var oSERIE = hOrd.indexOf("SERIE");
    var oORDEN = hOrd.indexOf("ORDEN");
    // Buscar pedidos vivos con ese código
    var PED_MUERTOS = ['TERMINADO','CANCELADO','CERRADO','SOBREPRODUCCION','ENTREGADO'];
    var pedidosVivos = [];
    for (var p = 1; p < dataPed.length; p++) {
      var pCod = String(dataPed[p][pdCOD]||'').trim().toUpperCase();
      var pEst = String(dataPed[p][pdEST]||'').trim().toUpperCase();
      var pFol = String(dataPed[p][pdFOL]||'').trim();
      if (pCod !== String(codigo).trim().toUpperCase()) continue;
      if (PED_MUERTOS.indexOf(pEst) > -1) continue;
      pedidosVivos.push({
        folio:  pFol,
        estado: pEst,
        desc:   String(dataPed[p][pdDESC]||'').trim(),
        cant:   Number(dataPed[p][pdCAN]||0)
      });
    }
    // Para cada pedido vivo, obtener sus órdenes
    pedidosVivos.forEach(function(ped) {
      var ordenes = [];
      for (var o = 1; o < dataOrd.length; o++) {
        if (String(dataOrd[o][oPED]||'').trim() !== ped.folio) continue;
        ordenes.push({
          id:      String(dataOrd[o][0]||''),
          orden:   String(dataOrd[o][oSERIE]||'') + '.' + ('0000'+String(dataOrd[o][oORDEN]||'')).slice(-4),
          proceso: String(dataOrd[o][oPROC]||'').trim(),
          maquina: String(dataOrd[o][oMAQ]||'').trim(),
          sol:     Number(dataOrd[o][oSOL]||0),
          prod:    Number(dataOrd[o][oPROD]||0),
          estado:  String(dataOrd[o][oEST]||'').trim().toUpperCase()
        });
      }
      ordenes.sort(function(a,b){ return a.proceso.localeCompare(b.proceso); });
      ped.ordenes = ordenes;
    });
    return JSON.stringify({ success: true, pedidos: pedidosVivos });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

function ejecutarCambioEstadoDirecto(id, nuevoEstado) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sh = ss.getSheetByName("ORDENES");
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(h => String(h).toUpperCase().trim());
  var colId = 0; 
  var colEstado = headers.indexOf("ESTADO");

  for(var i=1; i<data.length; i++){
    if(String(data[i][colId]) === String(id)){
      sh.getRange(i+1, colEstado+1).setValue(nuevoEstado);
      break;
    }
  }
  // Devolvemos el SVG actualizado para que el HTML lo pinte sin recargar
  return getSvgEstado(nuevoEstado);
}

function actualizarMaquinasRutaPlanif(codigo, proceso, maqsString) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("RUTAS");
    var data = sh.getDataRange().getValues();
    
    // Buscamos todas las filas que coincidan con el Código y el Proceso
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(codigo).trim() && 
          String(data[i][4]).toUpperCase().trim() === String(proceso).toUpperCase().trim()) {
        
        // Actualizamos la Columna F (índice 6) que es MAQUINA
        sh.getRange(i + 1, 6).setValue(maqsString);
      }
    }
    return "Ruta Maestra Actualizada: " + maqsString;
  } catch (e) {
    throw new Error("Error actualizando ruta: " + e.message);
  }
}

//**Agregado nuevo para hacer mejoras */

function obtenerDatosCantidadOrden(id) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh = ss.getSheetByName("ORDENES");
    var data = sh.getDataRange().getValues();
    var headers = data[0].map(h => String(h).toUpperCase().trim());
    
    var getIdx = function(n) { return headers.indexOf(n); };
    var idx = {
      ID: getIdx("ID"),
      SERIE: getIdx("SERIE"),
      ORDEN: getIdx("ORDEN"),
      PROCESO: getIdx("PROCESO"),
      UNIDAD: getIdx("UNIDAD"),
      CANTIDAD: getIdx("CANTIDAD"),
      SOLICITADO: getIdx("SOLICITADO"),
      PRODUCIDO: getIdx("PRODUCIDO"),
      DESCRIPCION: getIdx("DESCRIPCION"),
      PESO: getIdx("PESO"),
      LONGITUD: getIdx("LONGITUD")
    };
    
    var serieTarget = "";
    var ordenTarget = "";
    
    // Buscar la serie y orden del ID seleccionado
    for(var i=1; i<data.length; i++){
      if(String(data[i][idx.ID]) === String(id)){
        serieTarget = String(data[i][idx.SERIE]);
        ordenTarget = String(data[i][idx.ORDEN]);
        break;
      }
    }
    
    // Obtener todos los procesos de esa orden
    var resultados = [];
    for(var j=1; j<data.length; j++){
      if(String(data[j][idx.SERIE]) === serieTarget && String(data[j][idx.ORDEN]) === ordenTarget){
        resultados.push({
          id: String(data[j][idx.ID]),
          serie: String(data[j][idx.SERIE]),
          orden: String(data[j][idx.ORDEN]),
          proceso: String(data[j][idx.PROCESO]),
          unidad: String(data[j][idx.UNIDAD] || ""),
          cantidad: parseFloat(data[j][idx.CANTIDAD]) || 0,
          solicitado: parseFloat(data[j][idx.SOLICITADO]) || 0,
          producido: parseFloat(data[j][idx.PRODUCIDO]) || 0,
          descripcion: String(data[j][idx.DESCRIPCION] || ""),
          peso: parseFloat(data[j][idx.PESO]) || 0,
          longitud: parseFloat(data[j][idx.LONGITUD]) || 0
        });
      }
    }
    
    return JSON.stringify(resultados);
  } catch(e) {
    return JSON.stringify({error: e.toString()});
  }
}

function actualizarCantidadOrden(id, nuevaCantidad) {
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh        = ss.getSheetByName("ORDENES");
    var shPedidos = ss.getSheetByName("PEDIDOS");
    var data      = sh.getDataRange().getValues();
    var headers   = data[0].map(function(h){ return String(h).toUpperCase().trim(); });

    var getIdx = function(n){ return headers.indexOf(n); };
    var idx = {
      ID:          getIdx("ID"),
      SERIE:       getIdx("SERIE"),
      ORDEN:       getIdx("ORDEN"),
      PEDIDO:      getIdx("PEDIDO"),
      UNIDAD:      getIdx("UNIDAD"),
      CANTIDAD:    getIdx("CANTIDAD"),
      SOLICITADO:  getIdx("SOLICITADO"),
      DESCRIPCION: getIdx("DESCRIPCION"),
      PESO:        getIdx("PESO"),
      LONGITUD:    getIdx("LONGITUD"),
      PRODUCIDO:   getIdx("PRODUCIDO"),
      ESTADO:      getIdx("ESTADO")
    };

    // Buscar serie, orden y pedido de la orden recibida
    var serieTarget  = "";
    var ordenTarget  = "";
    var pedidoTarget = "";

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idx.ID]) === String(id)) {
        serieTarget  = String(data[i][idx.SERIE]);
        ordenTarget  = String(data[i][idx.ORDEN]);
        pedidoTarget = String(data[i][idx.PEDIDO]);
        break;
      }
    }

    // ── Actualizar ORDENES: todos los procesos de esa serie.orden ──
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][idx.SERIE]) !== serieTarget || String(data[j][idx.ORDEN]) !== ordenTarget) continue;

      var unidad      = String(data[j][idx.UNIDAD]      || "").toUpperCase().trim();
      var descripcion = String(data[j][idx.DESCRIPCION] || "").toUpperCase();
      var peso        = parseFloat(data[j][idx.PESO])   || 0;
      var longStr     = String(data[j][idx.LONGITUD]    || "");
      var longMatch   = longStr.match(/[\d.]+/);
      var longitud    = longMatch ? parseFloat(longMatch[0]) : 0;
      var esVarilla   = descripcion.indexOf("VARILLA") > -1;

      // Col I: CANTIDAD = nuevaCantidad (en unidad del pedido, sin convertir)
      sh.getRange(j + 1, idx.CANTIDAD + 1).setValue(nuevaCantidad);

      // Col N: SOLICITADO = conversión a KG según unidad
      var nuevoSolicitado = nuevaCantidad;
      if (unidad === "KG" || unidad === "ROL") {
        nuevoSolicitado = nuevaCantidad;
      } else if (unidad === "PZA" || unidad === "CTO") {
        nuevoSolicitado = nuevaCantidad * peso;
        if (esVarilla && longitud > 0) {
          nuevoSolicitado = nuevoSolicitado * longitud;
        }
      }
      sh.getRange(j + 1, idx.SOLICITADO + 1).setValue(nuevoSolicitado);

      // Re-evaluar ESTADO si estaba TERMINADO o SOBREPRODUCCION
      if (idx.PRODUCIDO > -1 && idx.ESTADO > -1) {
        var producido    = parseFloat(data[j][idx.PRODUCIDO]) || 0;
        var estadoActual = String(data[j][idx.ESTADO] || "").toUpperCase().trim();
        var REVERTIBLES  = ["TERMINADO", "SOBREPRODUCCION"];
        if (REVERTIBLES.indexOf(estadoActual) > -1 && producido < nuevoSolicitado) {
          sh.getRange(j + 1, idx.ESTADO + 1).setValue("EN PROCESO");
        }
      }
    }

    // ── Actualizar PEDIDOS: Col G (Cant) y Col P (CANT_PLAN) ──
    // Col G = nueva cantidad en unidad del pedido (PZA, KG, etc.) — sin convertir
    // Col P = suma de CANTIDAD de órdenes no canceladas del pedido (sin duplicar por serie.orden)
    if (pedidoTarget && shPedidos) {
      // Leer ORDENES frescos (ya escribimos encima, pero getDataRange vuelve a leer de Sheets)
      // Para evitar doble lectura, recorremos data[] actualizada en memoria
      // pero como ya hicimos setValue, recalculamos con los valores nuevos
      var cantPlanTotal = 0;
      var keysContadas  = {};
      for (var k = 1; k < data.length; k++) {
        if (String(data[k][idx.PEDIDO]).trim() !== String(pedidoTarget).trim()) continue;
        if (String(data[k][idx.ESTADO] || "").toUpperCase().trim() === "CANCELADO") continue;
        var keyOrd = String(data[k][idx.SERIE]) + "." + String(data[k][idx.ORDEN]);
        if (keysContadas[keyOrd]) continue;
        keysContadas[keyOrd] = true;
        // Sumar CANTIDAD (en unidad del pedido), pero actualizar con nuevaCantidad si es la orden editada
        var estaOrden = String(data[k][idx.SERIE]) === serieTarget && String(data[k][idx.ORDEN]) === ordenTarget;
        cantPlanTotal += estaOrden ? nuevaCantidad : (parseFloat(data[k][idx.CANTIDAD]) || 0);
      }

      var dataPed    = shPedidos.getDataRange().getValues();
      var headersPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var colPlanPed = headersPed.indexOf("CANT_PLAN"); // Col P

      for (var p = 1; p < dataPed.length; p++) {
        if (String(dataPed[p][1]).trim() !== String(pedidoTarget).trim()) continue;
        // Col G (índice 6): Cant = cantPlanTotal (en unidad del pedido)
        shPedidos.getRange(p + 1, 7).setValue(cantPlanTotal);
        // Col P: CANT_PLAN = cantPlanTotal (mismo valor)
        if (colPlanPed > -1) {
          shPedidos.getRange(p + 1, colPlanPed + 1).setValue(cantPlanTotal);
        }
        // Col I (índice 8): si el pedido estaba TERMINADO y se reabrió la orden, pasar a EN PROCESO
        var colEstPed = headersPed.indexOf("ESTADO");
        if (colEstPed === -1) colEstPed = 8; // fallback Col I
        var estadoPedActual = String(dataPed[p][colEstPed] || '').toUpperCase().trim();
        if (estadoPedActual === 'TERMINADO') {
          shPedidos.getRange(p + 1, colEstPed + 1).setValue('EN PROCESO');
        }
        break;
      }
    }

    SpreadsheetApp.flush();

    // Leer los valores reales que quedaron escritos para la fila de la orden solicitada (id)
    var solFinal    = 0;
    var estadoFinal = 'EN PROCESO';
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idx.ID]) === String(id)) {
        // nuevoSolicitado se calculó arriba; releerlo del arreglo de memoria no es posible
        // porque ya hicimos setValue (data[] no se actualiza). Lo recalculamos igual:
        var _unidad   = String(data[r][idx.UNIDAD]      || '').toUpperCase().trim();
        var _desc     = String(data[r][idx.DESCRIPCION] || '').toUpperCase();
        var _peso     = parseFloat(data[r][idx.PESO])   || 0;
        var _longStr  = String(data[r][idx.LONGITUD]    || '');
        var _longM    = _longStr.match(/[\d.]+/);
        var _long     = _longM ? parseFloat(_longM[0]) : 0;
        var _esVar    = _desc.indexOf('VARILLA') > -1;
        if (_unidad === 'KG' || _unidad === 'ROL') {
          solFinal = nuevaCantidad;
        } else if (_unidad === 'PZA' || _unidad === 'CTO') {
          solFinal = nuevaCantidad * _peso;
          if (_esVar && _long > 0) solFinal = solFinal * _long;
        } else {
          solFinal = nuevaCantidad;
        }
        // Estado: si era TERMINADO/SOBREPRODUCCION y producido < solFinal → EN PROCESO, si no conservar
        var _prod    = parseFloat(data[r][idx.PRODUCIDO]) || 0;
        var _estAct  = String(data[r][idx.ESTADO] || '').toUpperCase().trim();
        var _revert  = ['TERMINADO', 'SOBREPRODUCCION'];
        if (_revert.indexOf(_estAct) > -1 && _prod < solFinal) {
          estadoFinal = 'EN PROCESO';
        } else {
          estadoFinal = _estAct || 'EN PROCESO';
        }
        break;
      }
    }
    return JSON.stringify({ success: true, sol: solFinal, estado: estadoFinal });
  } catch(e) {
    throw new Error("Error actualizando cantidad: " + e.message);
  }
}

function actualizarPrioridadesMaquina(maquina, proceso, cambiosPrio) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("ORDENES");
    var data = sh.getDataRange().getValues();
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colID   = hdr.indexOf("ID");
    var colPRIO = hdr.indexOf("PRIORIDAD");

    var mapa = {};
    (cambiosPrio || []).forEach(function(c){ mapa[String(c.id)] = c.prioridad; });

    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][colID]);
      if (mapa.hasOwnProperty(id)) {
        sh.getRange(i+1, colPRIO+1).setValue(mapa[id]);
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

function eliminarRegistroProduccion(idRegistro, filaHoja) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return JSON.stringify({ success: false, msg: "Servidor ocupado." });
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd = ss.getSheetByName("PRODUCCION");
    var dataProd  = sheetProd.getDataRange().getValues();

    // Verificar que la fila corresponde al ID esperado
    var filaIdx = filaHoja - 1; // índice base 0
    if (filaIdx < 1 || filaIdx >= dataProd.length) {
      lock.releaseLock();
      return JSON.stringify({ success: false, msg: "Fila fuera de rango: " + filaHoja });
    }
    if (String(dataProd[filaIdx][0]) !== String(idRegistro)) {
      lock.releaseLock();
      return JSON.stringify({ success: false, msg: "ID no coincide con la fila indicada." });
    }

    // Obtener el ID de la orden antes de borrar
    var idOrden = String(dataProd[filaIdx][2]);

    // Eliminar la fila de PRODUCCION
    sheetProd.deleteRow(filaHoja);
    SpreadsheetApp.flush();

    // Recalcular estado de la orden
    var nuevoEstado = recalcularEstadoOrdenMaestro(idOrden);

    lock.releaseLock();
    return JSON.stringify({ success: true, nuevoEstado: nuevoEstado || "ACTUALIZADO" });
  } catch(e) {
    try { lock.releaseLock(); } catch(x) {}
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function obtenerProdPlanif(idOrden) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shProd = ss.getSheetByName("PRODUCCION");
    var shOrd  = ss.getSheetByName("ORDENES");
    var tz     = ss.getSpreadsheetTimeZone();

    var dOrd = shOrd.getDataRange().getValues();
    var ordenInfo = null;
    for (var o = 1; o < dOrd.length; o++) {
      if (String(dOrd[o][0]) == String(idOrden)) {
        ordenInfo = {
          tipo:    String(dOrd[o][19] || ""),
          dia:     String(dOrd[o][20] || ""),
          long:    String(dOrd[o][21] || ""),
          cuerda:  String(dOrd[o][22] || ""),
          cuerpo:  String(dOrd[o][23] || ""),
          acero:   String(dOrd[o][24] || ""),
          proceso: String(dOrd[o][11] || ""),
          iconoSVG: obtenerIconoSVG(String(dOrd[o][19] || ""))
        };
        break;
      }
    }

    var dProd = shProd.getDataRange().getValues();
    var registros = [];
    for (var p = 1; p < dProd.length; p++) {
      if (String(dProd[p][2]) !== String(idOrden)) continue;
      var fVal = dProd[p][5];
      var fStr = (fVal instanceof Date)
        ? Utilities.formatDate(fVal, tz, "yyyy-MM-dd")
        : String(fVal || "");
      registros.push({
        id:        String(dProd[p][0]),
        filaHoja:  p + 1,
        fecha:     fStr,
        turno:     String(dProd[p][6]  || ""),
        lote:      String(dProd[p][3]  || ""),
        maquina:   String(dProd[p][4]  || ""),
        pesoI:     Number(dProd[p][8])  || 0,
        pesoF:     Number(dProd[p][9])  || 0,
        pesoTina:  Number(dProd[p][22]) || 0,
        producido: Number(dProd[p][10]) || 0
      });
    }

    return JSON.stringify({ success: true, ordenInfo: ordenInfo, registros: registros });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function actualizarPesoMasivoPlanif(id, nuevoPeso) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOrd = ss.getSheetByName("ORDENES");
  var shRutas = ss.getSheetByName("RUTAS");
  var dataO = shOrd.getDataRange().getValues();
  var headO = dataO[0].map(h => String(h).toUpperCase().trim());
  
  var fLimite = new Date(); fLimite.setDate(fLimite.getDate() - 90);
  
  // Índices de ORDENES
  var col = {
    ID: 0, SERIE: 4, ORDEN: 5, CODIGO: 6, DESC: 7, CANT: 8, UNI: 9, SOL: 13, PESO: 18, LONG: 21, FECHA: headO.indexOf("FECHA_REGISTRO")
  };
  if(col.FECHA == -1) col.FECHA = headO.indexOf("FECHA");

  var serieTarget = "", ordenTarget = "", codigoTarget = "";

  // 1. Obtener datos de la fila que detonó el cambio
  for(var i=1; i<dataO.length; i++){
    if(String(dataO[i][col.ID]) === String(id)){
      serieTarget = String(dataO[i][col.SERIE]);
      ordenTarget = String(dataO[i][col.ORDEN]);
      codigoTarget = String(dataO[i][col.CODIGO]);
      break;
    }
  }

  // 2. Actualizar ORDENES (Misma familia O mismo código en 90 días)
  for(var j=1; j<dataO.length; j++){
    var fRow = dataO[j][col.FECHA];
    var fVal = (fRow instanceof Date) ? fRow : new Date();
    var matchFamilia = (String(dataO[j][col.SERIE]) === serieTarget && String(dataO[j][col.ORDEN]) === ordenTarget);
    var matchCodigoHistorial = (String(dataO[j][col.CODIGO]) === codigoTarget && fVal >= fLimite);

    if(matchFamilia || matchCodigoHistorial){
      // A. Actualizar Peso
      shOrd.getRange(j+1, col.PESO + 1).setValue(nuevoPeso);
      
      // B. Recalcular Solicitado
      var unidad = String(dataO[j][col.UNI]).toUpperCase();
      var cantidad = parseFloat(dataO[j][col.CANT]) || 0;
      var descripcion = String(dataO[j][col.DESC]).toUpperCase();
      var nSolicitado = 0;

      if(unidad.includes("KG") || unidad.includes("ROL")){
        nSolicitado = cantidad; // No cambia
      } else {
        if(descripcion.includes("VARILLA")){
          var longNum = evaluarFraccionPlanif(dataO[j][col.LONG]);
          nSolicitado = cantidad * nuevoPeso * longNum;
        } else {
          nSolicitado = cantidad * nuevoPeso;
        }
      }
      shOrd.getRange(j+1, col.SOL + 1).setValue(nSolicitado);
    }
  }

  // 3. Actualizar RUTAS MAESTRAS
  var dataR = shRutas.getDataRange().getValues();
  for(var r=1; r<dataR.length; r++){
    if(String(dataR[r][1]).trim().toUpperCase() === codigoTarget.toUpperCase()){
      shRutas.getRange(r+1, 16).setValue(nuevoPeso); // Col P = 16
    }
  }

  return "OK";
}

function obtenerCandidatosTransferencia(idOrdenOrigen) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd = ss.getSheetByName("ORDENES");
    var dOrd  = shOrd.getDataRange().getValues();

    var codigoOrigen = "";
    for (var i = 1; i < dOrd.length; i++) {
      if (String(dOrd[i][0]) == String(idOrdenOrigen)) {
        codigoOrigen = String(dOrd[i][6]).trim();
        break;
      }
    }
    if (!codigoOrigen) return JSON.stringify({ list: [] });

    var cerradas = ["TERMINADO", "CANCELADO", "SOBREPRODUCCION"];
    var list = [];
    for (var j = 1; j < dOrd.length; j++) {
      if (String(dOrd[j][0]) == String(idOrdenOrigen)) continue;
      if (String(dOrd[j][6]).trim() !== codigoOrigen) continue;
      if (cerradas.indexOf(String(dOrd[j][15]).toUpperCase().trim()) !== -1) continue;
      list.push({
        id:         String(dOrd[j][0]),
        nombre:     String(dOrd[j][4]) + "." + ("0000" + Number(dOrd[j][5])).slice(-4),
        solicitado: Number(dOrd[j][13]) || 0,
        producido:  Number(dOrd[j][14]) || 0
      });
    }
    list.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });
    return JSON.stringify({ list: list });
  } catch(e) {
    return JSON.stringify({ list: [], error: e.toString() });
  }
}

function ejecutarTraspasoParcialOTotal(idRegistro, idOrdenDest, kgs) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return JSON.stringify({ success: false, msg: "Servidor ocupado." });
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shProd = ss.getSheetByName("PRODUCCION");
    var dProd  = shProd.getDataRange().getValues();

    var filaReg = -1;
    for (var p = 1; p < dProd.length; p++) {
      if (String(dProd[p][0]) == String(idRegistro)) { filaReg = p + 1; break; }
    }
    if (filaReg < 0) return JSON.stringify({ success: false, msg: "Registro no encontrado." });

    var idOrdenOrigen = String(dProd[filaReg - 1][2]);
    var prodOrig      = Number(dProd[filaReg - 1][10]) || 0;
    var kgsNum        = Number(kgs);

    if (kgsNum <= 0) return JSON.stringify({ success: false, msg: "Cantidad inválida." });

    if (kgsNum >= prodOrig) {
      shProd.getRange(filaReg, 3).setValue(idOrdenDest);
    } else {
      shProd.getRange(filaReg, 11).setValue(prodOrig - kgsNum);
      var newRow = dProd[filaReg - 1].slice();
      newRow[0]  = Utilities.getUuid();
      newRow[2]  = idOrdenDest;
      newRow[10] = kgsNum;
      shProd.appendRow(newRow);
    }
    SpreadsheetApp.flush();
    recalcularEstadoOrdenMaestro(idOrdenOrigen);
    recalcularEstadoOrdenMaestro(idOrdenDest);
    lock.releaseLock();
    return JSON.stringify({ success: true, msg: "Traspaso realizado correctamente." });
  } catch(e) {
    try { lock.releaseLock(); } catch(ex) {}
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function getSvgEstado(estado) {
  var e = String(estado).toUpperCase().trim();
  var fill = "#f5f5f5"; var stroke = "#9e9e9e"; var text = "#616161";
  if (e == "ABIERTO")         { fill = "#fffde7"; stroke = "#fbc02d"; text = "#f57f17"; } 
  else if (e == "ACTIVE")     { fill = "#e3f2fd"; stroke = "#2196f3"; text = "#1565c0"; } 
  else if (e == "EN PROCESO") { fill = "#fff3e0"; stroke = "#fb8c00"; text = "#e65100"; } 
  else if (e == "TERMINADO")  { fill = "#e8f5e9"; stroke = "#4caf50"; text = "#2e7d32"; } 
  else if (e == "SOBREPRODUCCION") { fill = "#ffebee"; stroke = "#f44336"; text = "#c62828"; } 
  else if (e == "PLANEADO")   { fill = "#f3e5f5"; stroke = "#8e24aa"; text = "#7b1fa2"; }
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 110 24" style="width:100px; height:auto;">' +
         '<rect width="110" height="24" rx="6" fill="' + fill + '" stroke="' + stroke + '" stroke-width="1.5" />' +
         '<text x="55" y="16" font-family="Arial" font-size="11" font-weight="bold" text-anchor="middle" fill="' + text + '">' + e + '</text>' +
         '</svg>';
}

// Aumentamos el tamaño del Acero para que sea legible
function getSvgAcero(acero) {
  var txt = String(acero).toUpperCase().trim();
  if(!txt || txt == "UNDEFINED") txt = "-";

  // Configuración: fill=fondo, text=color texto, stripe=color franja triangular (null=sin franja)
  var fill = "#263238";
  var textColor = "white";
  var stripe = null; // color de la franja triangular en esquina inferior derecha

  if (txt === "1004") {
    fill = "#FFFFFF"; textColor = "#000000"; stripe = "#2E7D32"; // blanco + franja verde
  } else if (txt === "1006" && !txt.includes("REC")) {
    fill = "#FFFFFF"; textColor = "#000000"; stripe = "#212121"; // blanco + franja negra
  } else if (txt === "1008" && !txt.includes("REC")) {
    fill = "#FFFFFF"; textColor = "#000000"; stripe = null; // blanco puro
  } else if (txt === "1018" && !txt.includes("REC")) {
    fill = "#2E7D32"; textColor = "#FFFFFF"; stripe = null; // verde bandera
  } else if (txt.includes("10B21")) {
    fill = "#FFFFFF"; textColor = "#000000"; stripe = "#7B1FA2"; // blanco + franja morada
  } else if (txt === "1038") {
    fill = "#FFD600"; textColor = "#000000"; stripe = null; // amarillo intenso
  } else if (txt === "1033") {
    fill = "#FFD600"; textColor = "#000000"; stripe = "#212121"; // amarillo + franja negra
  } else if (txt === "1541") {
    fill = "#FF6F00"; textColor = "#000000"; stripe = null; // naranja intenso
  } else if (txt === "4140") {
    fill = "#F06292"; textColor = "#000000"; stripe = null; // rosa intenso
  } else if (txt === "1060") {
    fill = "#0D47A1"; textColor = "#FFFFFF"; stripe = null; // azul marino
  } else if (txt === "1055") {
    fill = "#1565C0"; textColor = "#FFFFFF"; stripe = "#212121"; // azul + franja negra
  } else if (txt === "1045") {
    fill = "#C62828"; textColor = "#FFFFFF"; stripe = null; // rojo
  } else if (txt === "1006 REC") {
    fill = "#FFFFFF"; textColor = "#000000"; stripe = "#7B1FA2"; // blanco + franja morada
  } else if (txt === "1008 REC") {
    fill = "#FFFFFF"; textColor = "#000000"; stripe = "#212121"; // blanco + franja negra
  } else if (txt === "1018 REC") {
    fill = "#2E7D32"; textColor = "#FFFFFF"; stripe = "#7B1FA2"; // verde + franja morada
  } else if (txt.includes("INOX") || txt.includes("304")) {
    fill = "#00BCD4"; textColor = "#FFFFFF"; stripe = null;
  }

  var svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 90 28" style="width:85px; height:auto;">';

  // Fondo principal con borde negro
  svg += '<rect width="90" height="28" rx="6" fill="' + fill + '" stroke="#000000" stroke-width="1.5"/>';

  // Franja triangular en esquina inferior derecha (como la imagen de referencia)
  if (stripe) {
    svg += '<clipPath id="cp' + txt.replace(/[^a-zA-Z0-9]/g,'') + '">';
    svg += '<rect width="90" height="28" rx="6"/>';
    svg += '</clipPath>';
    svg += '<polygon points="55,28 90,28 90,0" fill="' + stripe + '" clip-path="url(#cp' + txt.replace(/[^a-zA-Z0-9]/g,'') + ')"/>';
  }

  // Texto
  svg += '<text x="40" y="19" font-family="Arial" font-weight="900" font-size="13" text-anchor="middle" fill="' + textColor + '">' + txt + '</text>';

  svg += '</svg>';
  return svg;
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA ProgramadorHTML (MENU PROGRAMADOR MRP)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function obtenerFamiliasMRP() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetInv = ss.getSheetByName("INVENTARIO_EXTERNO");
  var sheetCod = ss.getSheetByName("CODIGOS");
  var sheetPlan = ss.getSheetByName("PLANIFICACION_STOCK");
  
  // A. MAPEO CODIGO -> FAMILIA Y PESO
  var dataCod = sheetCod.getDataRange().getValues();
  var hCod = dataCod[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var icCod = hCod.indexOf("CODIGO"); if(icCod<0) icCod=0;
  var icFam = hCod.indexOf("FAMILY"); if(icFam<0) icFam = hCod.indexOf("FAMILIA"); if(icFam<0) icFam=8;
  var icPeso = hCod.indexOf("PESO"); 

  var mapInfo = {};
  for(var c=1; c<dataCod.length; c++) {
     var cod = String(dataCod[c][icCod]).trim();
     var fam = String(dataCod[c][icFam]||"OTROS").toUpperCase().trim();
     var peso = (icPeso > -1) ? (Number(dataCod[c][icPeso]) || 0) : 0;
     if(cod) mapInfo[cod] = { f: fam, p: peso };
  }

  // B. FAMILIAS ACTIVAS
  var dataInv = sheetInv.getDataRange().getValues();
  var familiasActivas = {};
  for(var i=1; i<dataInv.length; i++) {
     var codInv = String(dataInv[i][0]).trim();
     if(codInv && mapInfo[codInv]) familiasActivas[mapInfo[codInv].f] = true;
  }
  var listaFamilias = Object.keys(familiasActivas).sort();

  // C. CALCULAR ESTADISTICAS
  var dataPlan = sheetPlan.getDataRange().getValues();
  // Estructura Indices V131: 
  // 1=CODIGO, 3=MAX, 4=EXIS, 5=A_FABRICAR, 13=BACK, 17=ESTATUS, 20=LONGITUD, 26=FECHA
  
  var stats = {}; 
  var global = { lastUpdate:0, count:0, pendientes:0, procesados:0, sumMax:0, sumExis:0, sumBack:0, good:0, bad:0, progKg:0 };

  listaFamilias.forEach(f => {
      stats[f] = { lastUpdate:0, count:0, pendientes:0, procesados:0, sumMax:0, sumExis:0, sumBack:0, good:0, bad:0, progKg:0 };
  });
  stats["OTROS"] = { lastUpdate:0,count:0,pendientes:0,procesados:0,sumMax:0,sumExis:0,sumBack:0,good:0,bad:0,progKg:0 };

  // EXCEPCIONES: Familias que NO se multiplican por peso (Pasan directo)
  var familiasSinPeso = ["ACERO REDONDO", "BARRA PULIDA", "CLAVO CONCRETO", "CLAVO MADERA", "COLD ROLLED", "COLD ROLLED CUA", "COLD ROLLED HEX"];

  if (dataPlan.length > 1) {
      for (var p = 1; p < dataPlan.length; p++) {
          var row = dataPlan[p];
          var codP = String(row[1]).trim();
          var info = mapInfo[codP] || { f: "OTROS", p: 0 };
          var familia = info.f;
          
          if (!stats[familia]) continue; 

          var max = Number(row[3]) || 0;
          var exis = Number(row[4]) || 0;
          
          // Dato clave: A_FABRICAR (Ya viene calculado con la lógica compleja de Varilla Max-Exis-Back-Exis2 desde sincronizarMRP)
          var aFabricar = Number(row[5]) || 0; 
          
          var back = Number(row[13]) || 0; 
          var estatus = String(row[17]).toUpperCase(); 
          var fecha = row[26]; 
          var longitudTxt = String(row[20]); // Columna T (LONGITUD)

          var ts = (fecha instanceof Date) ? fecha.getTime() : 0;
          if (ts > stats[familia].lastUpdate) stats[familia].lastUpdate = ts;
          if (ts > global.lastUpdate) global.lastUpdate = ts;

          stats[familia].count++; global.count++;
          if (estatus == "PENDIENTE") { stats[familia].pendientes++; global.pendientes++; }
          else if (estatus == "PROCESADO" || estatus == "EJECUTADO") { stats[familia].procesados++; global.procesados++; }
          
          stats[familia].sumMax += max; global.sumMax += max;
          stats[familia].sumExis += exis; global.sumExis += exis;
          stats[familia].sumBack += back; global.sumBack += back;

          // --- KPI PROGRAMA (En Kilos) ---
          if (aFabricar > 0) {
             var pesoU = info.p; // Peso unitario del catalogo
             var valorKg = 0;

             if (familia.includes("VARILLA")) {
                 // REGLA VARILLA: A_FABRICAR * PESO * LONGITUD
                 // Extraemos el número de "3.00 mts"
                 var match = longitudTxt.match(/[\d\.]+/);
                 var largo = match ? parseFloat(match[0]) : 0;
                 
                 // Seguridad: si no hay largo o peso, asumimos 1 para no dar 0
                 var factorLargo = (largo > 0) ? largo : 0; 
                 var factorPeso = (pesoU > 0) ? pesoU : 0;
                 
                 valorKg = aFabricar * factorPeso * factorLargo;
             } 
             else if (familiasSinPeso.includes(familia)) {
                 // EXCEPCIONES: Pasa directo (A_FABRICAR se asume ya en la unidad correcta o Kilos directos)
                 valorKg = aFabricar; 
             } 
             else {
                 // ESTÁNDAR (Tornillos, etc): A_FABRICAR * PESO
                 valorKg = aFabricar * (pesoU > 0 ? pesoU : 1);
             }
             
             stats[familia].progKg += valorKg;
             global.progKg += valorKg;
          }

          var ratio = (max > 0) ? (exis / max) : 0;
          if (ratio >= 0.25) { stats[familia].good++; global.good++; }
          else { stats[familia].bad++; global.bad++; }
      }
  }
  
  // Redondear a Enteros
  for(var k in stats) stats[k].progKg = Math.round(stats[k].progKg);
  global.progKg = Math.round(global.progKg);

  return { lista: listaFamilias, stats: stats, global: global };
}

// MOTOR MRP V142 (COLUMNA PESO AB + IGNORAR WIP/PISO)
function sincronizarMRP(familiasSeleccionadas, ignorarWip, ignorarPiso) {
  var tInicio = new Date().getTime();
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  
  try {
    var sheetPlan = ss.getSheetByName("PLANIFICACION_STOCK");
    var sheetInv  = ss.getSheetByName("INVENTARIO_EXTERNO");
    var sheetCod  = ss.getSheetByName("CODIGOS");
    var sheetOrd  = ss.getSheetByName("ORDENES");
    var sheetPed  = ss.getSheetByName("PEDIDOS"); 
    var sheetEnv  = ss.getSheetByName("ENVIADO");
    var sheetRut  = ss.getSheetByName("RUTAS");

    var clean  = function(val) { return String(val).trim().toUpperCase(); };
    var getCol = function(h, n) { return h.map(x => String(x).toUpperCase().trim()).indexOf(n); };
    var valNum = function(n) { var v = parseFloat(n); return isNaN(v) ? 0 : v; };

    // 1. MAPA MAESTRO (CATÁLOGO DE CÓDIGOS)
    var dataCod = sheetCod.getDataRange().getValues();
    var hCod = dataCod[0];
    var icCod   = getCol(hCod, "CODIGO"); 
    var icFam   = getCol(hCod, "FAMILY"); if(icFam < 0) icFam = getCol(hCod, "FAMILIA");
    var icDesc  = getCol(hCod, "DESCRIPCION"); 
    var icUni   = getCol(hCod, "UNIDAD"); 
    var icPeso  = getCol(hCod, "PESO");
    var icVenta = getCol(hCod, "CODIGO_VENTA"); // Columna F

    var mapaMaestro = {}; 
    for(var i=1; i<dataCod.length; i++) {
       var fam = (icFam > -1) ? clean(dataCod[i][icFam]) : "OTROS";
       if (familiasSeleccionadas.includes(fam)) {
          var cFab = clean(dataCod[i][icCod]); // El 7... (Negro)
          if(cFab) {
             mapaMaestro[cFab] = {
                fam: fam,
                uni: (icUni > -1) ? clean(dataCod[i][icUni]) : "PZA",
                peso: (icPeso > -1) ? valNum(dataCod[i][icPeso]) : 0,
                desc: (icDesc > -1) ? dataCod[i][icDesc] : "",
                venta: (icVenta > -1) ? clean(dataCod[i][icVenta]) : cFab // El 1... (Venta)
             };
          }
       }
    }

    // 2. MAPA DE INVENTARIO (MAX, EXIS, BACK ALPHA)
    var dataInv = sheetInv.getDataRange().getValues();
    var hInv = dataInv[0];
    var iiCod = getCol(hInv, "CODIGO"); 
    var iiMax = getCol(hInv, "MAXIMO"); 
    var iiExi = getCol(hInv, "EXISTENCIA");
    var iiBack = getCol(hInv, "BACKORDER");
    var iiMin = getCol(hInv, "MINIMO");
    var iiDesc = getCol(hInv, "DESCRIPCION"); 

    var mapInvFull = {};
    for(var i=1; i<dataInv.length; i++) {
       var cInv = clean(dataInv[i][iiCod]);
       if(cInv) {
           mapInvFull[cInv] = {
               max: valNum(dataInv[i][iiMax]),
               min: valNum(dataInv[i][iiMin]),
               exi: valNum(dataInv[i][iiExi]),
               back: (iiBack > -1) ? valNum(dataInv[i][iiBack]) : 0,
               desc: (iiDesc > -1) ? dataInv[i][iiDesc] : ""
           };
       }
    }

    // 3. PREPARAR CÓDIGOS A PROCESAR (LÓGICA ESPECIAL VARILLA)
    var codigosProcesar = {}; 
    var inventarioCalculado = [];

    for (var cFab in mapaMaestro) {
        var info = mapaMaestro[cFab];
        var cVenta = info.venta;
        
        // REGLA ORO: Solo procesar si el código de VENTA existe en inventario externo
        if (!mapInvFull[cVenta]) continue;

        var esVarillaEspecial = (info.fam.includes("VARILLA") && cFab !== cVenta);
        
        // El inventario base (MAX, EXIS, BACK) se lee del código de VENTA (1...)
        var invVenta = mapInvFull[cVenta];
        
        // La existencia del Negro (7...) se lee para el cálculo posterior
        var exiNegro = 0;
        if (mapInvFull[cFab]) exiNegro = mapInvFull[cFab].exi;

        var dataItem = { 
          cod: cFab, 
          desc: info.desc || invVenta.desc, 
          max: invVenta.max, 
          min: invVenta.min, 
          exi: invVenta.exi, 
          backAlpha: invVenta.back,
          exiNegro: exiNegro,
          esEspecial: esVarillaEspecial
        };

        // Lógica inicial A_FABRICAR (Sugerencia bruta)
        var aFab = invVenta.max - invVenta.exi;
        // Si es varilla especial, la sugerencia bruta ya considera el back y existencia del negro para no inflar pedidos
        if (esVarillaEspecial) {
           aFab = invVenta.max - invVenta.exi - invVenta.back - exiNegro;
        }
        dataItem.aFab = Math.max(0, aFab);
        
        codigosProcesar[cFab] = true;
        inventarioCalculado.push(dataItem);
    }

    // 4. INFO PEDIDOS (MANTENIDO IGUAL)
    var dataPed = sheetPed.getDataRange().getValues();
    var hPed = dataPed[0].map(x => String(x).toUpperCase().trim());
    var ipPed = getCol(hPed, "PEDIDO");
    var ipFecha = getCol(hPed, "FECHA"); 
    var ipCant = getCol(hPed, "CANTIDAD");
    var ipEst = getCol(hPed, "ESTADO"); 
    
    var infoPedidos = {}; var mapEstadosPedidos = {}; 
    for(var p=1; p<dataPed.length; p++) {
       var nomP = clean(dataPed[p][ipPed]);
       var estP = (ipEst > -1) ? clean(dataPed[p][ipEst]) : "";
       mapEstadosPedidos[nomP] = estP;
       if((nomP.includes("QSQ") || nomP.includes("TEM")) && !["CANCELADO","CERRADO","TERMINADO"].includes(estP)) {
          var fRaw = dataPed[p][ipFecha];
          var fStr = (fRaw instanceof Date) ? Utilities.formatDate(fRaw, Session.getScriptTimeZone(), "dd/MM/yy") : String(fRaw);
          infoPedidos[nomP] = { fecha: fStr, cant: valNum(dataPed[p][ipCant]) };
       }
    }

    // 5. RUTAS (MANTENIDO IGUAL - CUIDA VARILLA)
    var dataRut = sheetRut.getDataRange().getDisplayValues(); 
    var hRut = dataRut[0].map(x => String(x).toUpperCase().trim());
    var mapaTecnico = {};
    for(var r=1; r<dataRut.length; r++) {
       var cRut = clean(dataRut[r][getCol(hRut, "CODIGO")]);
       if(codigosProcesar[cRut] && !mapaTecnico[cRut]) {
          var txtL = dataRut[r][getCol(hRut, "LONGITUD")];
          var match = String(txtL).match(/[\d\.]+/);
          mapaTecnico[cRut] = {
             tipo: dataRut[r][getCol(hRut, "TIPO")],
             dia: "'" + dataRut[r][getCol(hRut, "DIAMETRO")],
             long: "'" + txtL,
             longVal: match ? parseFloat(match[0]) : 0,
             cuerda: "'" + dataRut[r][getCol(hRut, "CUERDA")],
             cuerpo: "'" + dataRut[r][getCol(hRut, "CUERPO")],
             acero: dataRut[r][getCol(hRut, "ACERO")]
          };
       }
    }

    // 6. ÓRDENES (MANTENIDO IGUAL - CALCULA BACK PLANTA)
    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd = dataOrd[0].map(x => String(x).toUpperCase().trim());
    var ioCod = getCol(hOrd, "CODIGO"), ioSol = getCol(hOrd, "SOLICITADO"), ioProd = getCol(hOrd, "PRODUCIDO");
    var ioEst = getCol(hOrd, "ESTADO"), ioPed = getCol(hOrd, "PEDIDO"), ioSec = getCol(hOrd, "SEC");
    var ioSerie = getCol(hOrd, "SERIE"), ioNum = getCol(hOrd, "ORDEN"), ioPeso = getCol(hOrd, "PESO"), ioUni = getCol(hOrd, "UNIDAD");

    var ordenesUnicas = {};
    for(var o=1; o<dataOrd.length; o++) {
       var cOrd = clean(dataOrd[o][ioCod]);
       if(codigosProcesar[cOrd]) {
          var est = clean(dataOrd[o][ioEst]);
          if(!["CANCELADO","CERRADA"].includes(est)) {
             var key = dataOrd[o][ioSerie] + "." + dataOrd[o][ioNum];
             // Agregamos 'uni' para guardar la unidad de la orden (Col J)
              if(!ordenesUnicas[key]) ordenesUnicas[key] = { rows: [], cod: cOrd, pedido: clean(dataOrd[o][ioPed]), peso: valNum(dataOrd[o][ioPeso]), uni: clean(dataOrd[o][ioUni]) };
             ordenesUnicas[key].rows.push({ sec: valNum(dataOrd[o][ioSec]), sol: valNum(dataOrd[o][ioSol]), prod: valNum(dataOrd[o][ioProd]), est: est });
          }
       }
    }

    var mapResultados = {};
    var mapEnvWIP = {}; 

    for(var key in ordenesUnicas) {
        var ord = ordenesUnicas[key];
        var c = ord.cod;
        var tec = mapaTecnico[c] || { longVal: 0 };
        var esV = mapaMaestro[c].fam.includes("VARILLA");
        
        if(!mapResultados[c]) mapResultados[c] = { back2:0, pt:0, wipFirst:0, wipLast:0, pedsSet:new Set(), pedsMap:{} };
        
        // LÓGICA QUIRÚRGICA BACK 2 (V150)
        // LÓGICA QUIRÚRGICA BACK 2 (V151)
        var divisor = 1;
        var unidadOrden = ord.uni || "";

        // Solo calculamos divisor si el producto se gestiona por PIEZAS o CIENTOS
        if (unidadOrden.includes("PZA") || unidadOrden.includes("CTO")) {
            if (esV && tec.longVal > 0) {
                // Caso Varilla: Peso Catálogo * Longitud Real
                divisor = ord.peso * tec.longVal;
            } else {
                // Caso Estándar: Solo Peso Catálogo
                divisor = (ord.peso > 0) ? ord.peso : 1;
            }
        } else {
            // Caso KG o ROL: No se convierte nada (divisor 1)
            divisor = 1;
        }

        var minS=999, maxS=-1, rMin=null, rMax=null;
        ord.rows.forEach(r => { if(r.sec < minS){minS=r.sec; rMin=r;} if(r.sec > maxS){maxS=r.sec; rMax=r;} });

        if(rMin) {
           mapResultados[c].wipFirst += rMin.prod;
           if(!["TERMINADO","SOBREPRODUCCION"].includes(rMin.est)) {
              mapResultados[c].back2 += Math.max(0, (rMin.sol - rMin.prod) / divisor);
           }
        }
        if(rMax) {
           mapResultados[c].wipLast += rMax.prod;
           mapResultados[c].pt += (rMax.prod / divisor);
        }

        if(infoPedidos[ord.pedido]) {
            var pData = infoPedidos[ord.pedido];
            if(!mapResultados[c].pedsMap[ord.pedido]) mapResultados[c].pedsMap[ord.pedido] = { fecha:pData.fecha, cant:pData.cant, prod:0 };
            mapResultados[c].pedsMap[ord.pedido].prod += (rMin ? rMin.prod / divisor : 0);
            mapResultados[c].pedsSet.add(ord.pedido);
        }
    }

    // 7. ENVIADOS (MANTENIDO IGUAL)
    var dataEnv = sheetEnv.getDataRange().getValues();
    var hEnv = dataEnv[0].map(x => String(x).toUpperCase().trim());
    var ieCod = getCol(hEnv, "CODIGO"), ieKg = getCol(hEnv, "KILOS"), iePed = getCol(hEnv, "PEDIDO");
    var mapEnvTotal = {};

    for(var e=1; e<dataEnv.length; e++) {
       var cEnv = clean(dataEnv[e][ieCod]);
       if(codigosProcesar[cEnv]) {
          var pedE = clean(dataEnv[e][iePed]);
          if(["TERMINADO","CANCELADO"].includes(mapEstadosPedidos[pedE])) continue;

          var kilos = valNum(dataEnv[e][ieKg]);
          var info = mapaMaestro[cEnv];
          var tec = mapaTecnico[cEnv] || { longVal: 0 };
          
          // LÓGICA QUIRÚRGICA: Solo dividir si es PZA o CTO
          var div = 1;
          var uni = info.uni || "";
          if (uni.includes("PZA") || uni.includes("CTO")) {
             if(info.peso > 0) {
                div = (info.fam.includes("VARILLA") && tec.longVal > 0) ? (info.peso * tec.longVal) : info.peso;
             }
          }
          
          mapEnvTotal[cEnv] = (mapEnvTotal[cEnv] || 0) + (kilos / div);
          
          if(pedE) {
             var kEnv = cEnv + "_" + pedE;
             mapEnvWIP[kEnv] = (mapEnvWIP[kEnv] || 0) + kilos;
          }
       }
    }

    // 8. BALANCE FINAL (LÓGICA QUIRÚRGICA VARILLA)
    var nuevosRegistros = [];
    var fechaEjec = new Date();

    for (var k=0; k<inventarioCalculado.length; k++) {
    var item = inventarioCalculado[k];
    var c = item.cod; // Este es el 7...
    
    // RE-VERIFICACIÓN DE EXISTENCIA NEGRO (7...)
    // Lo buscamos directamente en el mapa de inventario original para asegurar el dato
    var datoInvNegro = mapInvFull[c] ? mapInvFull[c].exi : 0;

    var res = mapResultados[c] || { back2:0, pt:0, wipFirst:0, wipLast:0, pedsSet:new Set(), pedsMap:{} };
    var info = mapaMaestro[c];
    var tec = mapaTecnico[c] || { tipo:"", dia:"", long:"", cuerda:"", cuerpo:"", acero:"", longVal:0 };
    var env = mapEnvTotal[c] || 0;

    var kilosEnvRel = 0;
    res.pedsSet.forEach(p => { kilosEnvRel += (mapEnvWIP[c + "_" + p] || 0); });
    
    var divWip = 1;
    var uniActual = info.uni || "";
    
    // Solo aplicar divisor si la unidad es Piezas o Cientos
    if (uniActual.includes("PZA") || uniActual.includes("CTO")) {
       if(info.peso > 0) {
          divWip = (info.fam.includes("VARILLA") && tec.longVal > 0) ? (info.peso * tec.longVal) : info.peso;
       }
    }
    
    var wipPzas = Math.max(0, (res.wipFirst / divWip) - res.pt);
    var pisoPzas = Math.max(0, res.pt - env);

    var vWip = ignorarWip ? 0 : wipPzas;
    var vPiso = ignorarPiso ? 0 : pisoPzas;
    
    var balance = 0;
    if (item.esEspecial) {
       // Formula Varilla 7 vs 1
       balance = Math.round(item.aFab - res.back2 - vWip - vPiso - item.backAlpha - datoInvNegro);
    } else {
       balance = Math.round(item.aFab - res.back2 - vWip - vPiso);
    }

    var accion = "NADA"; var cantA = 0;
    var aFabr = Math.max(0, Math.round(item.aFab));
    var umbral20 = aFabr * 0.20;
    
    if (balance > 0) {
        cantA = balance;
        if (res.back2 > 0) {
            // Ya hay órdenes en planta
            accion = (balance <= umbral20) ? "CANT A MAS" : "NUEVO PEDIDO";
        } else {
            accion = "NUEVO PEDIDO";
        }
    } else if (balance < 0 && res.back2 > 0) {
        var excedente = Math.abs(balance);
        cantA = excedente;
        accion = (excedente > umbral20) ? "CANCELAR EX" : "CANT A MENOS";
    }

    var listaPeds = [];
    for(var id in res.pedsMap) { listaPeds.push({ id:id, f:res.pedsMap[id].fecha, q:res.pedsMap[id].cant, p:Math.round(res.pedsMap[id].prod) }); }

    nuevosRegistros.push([
       Utilities.getUuid(), c, item.desc, item.max, item.exi, Math.round(item.aFab),
       listaPeds.length > 0 ? JSON.stringify(listaPeds) : "-", listaPeds.length,
       Math.round(res.back2), Math.round(res.pt), Math.round(wipPzas), Math.round(env),
       Math.round(pisoPzas), item.backAlpha, accion, Math.round(cantA), "", "PENDIENTE",
       tec.tipo, tec.dia, tec.long, tec.cuerda, tec.cuerpo, tec.acero,
       item.min, 
       Math.round(datoInvNegro), // COL 26: Aseguramos que guarde la existencia del 7...
       fechaEjec, info.peso
    ]);
}

    // 9. ACTUALIZAR HOJA — REEMPLAZO SEGURO (sin clearContents para evitar pérdida de datos)
    var headers = ["ID","CODIGO","DESCRIPCION","MAXIMO","EXISTENCIA","A_FABRICAR","ID_ULTIMO_PEDIDO","PEDIDOS_ACTIVOS","BACK_PLANTA","PROD_TERMINADO","WIP","ENVIADO","STOCK_PISO","BACK_ALPHA","ACCION_SUGERIDA","CANTIDAD_ACCION","DETALLE_ACCION","ESTATUS_REVISION","TIPO","DIAMETRO","LONGITUD","CUERDA","CUERPO","ACERO","MINIMO","EXTRA_EXIS_NEGRO","FECHA_HORA","PESO"];
    var dataOld = sheetPlan.getDataRange().getValues();
    var finalData = [headers];
    if(dataOld.length > 1) {
       for(var j=1; j<dataOld.length; j++){
          var cOld = clean(dataOld[j][1]);
          if(!codigosProcesar[cOld]) {
             var r = dataOld[j]; while(r.length < 28) r.push("");
             finalData.push(r.slice(0,28));
          }
       }
    }
    var allData = finalData.concat(nuevosRegistros);
    var totalFilas = allData.length;
    var totalFilasAntes = sheetPlan.getLastRow();
    
    // Escribir los datos nuevos
    sheetPlan.getRange(1, 1, totalFilas, 28).setValues(allData);
    
    // Si había más filas antes, limpiar solo las sobrantes
    if (totalFilasAntes > totalFilas) {
       sheetPlan.getRange(totalFilas + 1, 1, totalFilasAntes - totalFilas, 28).clearContent();
    }

    return "✅ Sincronizado en " + ((new Date().getTime()-tInicio)/1000) + "s. " + nuevosRegistros.length + " códigos negros procesados.";
  } catch(e) { return "❌ ERROR: " + e.toString(); }
}

function obtenerDatosProgramador(familias) {
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetPlan = ss.getSheetByName("PLANIFICACION_STOCK");
    var sheetCod  = ss.getSheetByName("CODIGOS");
    var sheetOrd  = ss.getSheetByName("ORDENES");

    // Mapa MP por código desde ORDENES Col AD
    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(h => String(h).toUpperCase().trim());
    var oCOD    = hOrd.indexOf("CODIGO");
    var oMP     = hOrd.indexOf("ALERTA_SECUENCIA");
    var oEST    = hOrd.indexOf("ESTADO");

    var mapaMP = {};  // codigo → MP más frecuente en órdenes vivas
    for (var i = 1; i < dataOrd.length; i++) {
      var est = String(dataOrd[i][oEST]||"").toUpperCase().trim();
      if (est === "TERMINADO" || est === "CANCELADO") continue;
      var cod = String(dataOrd[i][oCOD]||"").trim().toUpperCase();
      var mp  = oMP >= 0 ? String(dataOrd[i][oMP]||"").trim().toUpperCase() : "";
      if (cod && mp && !mapaMP[cod]) mapaMP[cod] = mp;
    }

    var dataCod = sheetCod.getDataRange().getValues();
    var hCod    = dataCod[0].map(h => String(h).toUpperCase().trim());
    var icCod   = hCod.indexOf("CODIGO"); if(icCod<0) icCod=0;
    var icFamCod= hCod.indexOf("FAMILY"); if(icFamCod<0) icFamCod=hCod.indexOf("FAMILIA");
    var icPeso  = hCod.indexOf("PESO");

    var mapPesos={}, mapFam={};
    for (var c=1; c<dataCod.length; c++) {
      var rawCod=dataCod[c][icCod];
      var codStr=String(rawCod).trim(), codNum=Number(rawCod);
      mapPesos[codStr]=mapPesos[codNum]=(icPeso>-1)?(Number(dataCod[c][icPeso])||0):0;
      mapFam[codStr]=mapFam[codNum]=String(dataCod[c][icFamCod]||"OTROS").toUpperCase().trim();
    }

    var data    = sheetPlan.getDataRange().getValues();
    if (data.length<=1) return "[]";

    var headers       = data[0].map(h=>String(h).toUpperCase().trim());
    var idxPlanCod    = headers.indexOf("CODIGO");
    var idxStatus     = headers.indexOf("ESTATUS_REVISION");
    var idxExtraNegro = headers.indexOf("EXTRA_EXIS_NEGRO");

    var resultado = [];
    for (var i=1; i<data.length; i++) {
      var row    = data[i];
      var status = (idxStatus>-1)?String(row[idxStatus]).toUpperCase():"PENDIENTE";
      if (status==="PROCESADO"||status==="RECHAZADO"||status==="EJECUTADO") continue;

      var codVal  = row[idxPlanCod];
      var codItem = String(codVal).trim();
      var famItem = mapFam[codItem]||mapFam[Number(codVal)]||"SIN FAMILIA";
      if (!familias.includes(famItem)) continue;

      var obj={};
      for (var k=0; k<headers.length; k++) {
        var key=headers[k], val=row[k];
        if (val instanceof Date) val=Utilities.formatDate(val,Session.getScriptTimeZone(),"yyyy-MM-dd");
        if (val===undefined||val===null) val="";
        if (key==="DIAMETRO") val=String(val);
        obj[key]=val;
      }
      obj["EXTRA_EXIS_NEGRO"]=(idxExtraNegro>-1)?(Number(row[idxExtraNegro])||0):0;
      obj["FAMILIA"]=famItem;
      obj["MAT_PRIMA"]=mapaMP[codItem]||mapaMP[codItem.replace(/^0+/,"")]||"";  // ← NUEVO
      resultado.push(obj);
    }
    return JSON.stringify(resultado);
  } catch(e) { return JSON.stringify([{"ERROR":e.toString()}]); }
}

// 4.5 OBTENER PRODUCCION DETALLADA (Para modal WIP en ProgramadorHTML)
function obtenerProduccionDetallada(codigoBuscado) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetProd = ss.getSheetByName("PRODUCCION");
    
    var codigo = String(codigoBuscado).trim();
    
    // 1. Buscar órdenes vivas de este código
    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ioCod=hOrd.indexOf("CODIGO"), ioSerie=hOrd.indexOf("SERIE"), ioOrden=hOrd.indexOf("ORDEN");
    var ioSec=hOrd.indexOf("SEC"), ioProc=hOrd.indexOf("PROCESO"), ioEst=hOrd.indexOf("ESTADO");
    var ioSol=hOrd.indexOf("SOLICITADO"), ioPed=hOrd.indexOf("PEDIDO"), ioID=0;
    
    var ordenesMap = {};
    for (var o=1; o<dataOrd.length; o++) {
      var codOrd = String(dataOrd[o][ioCod]).replace(/^'/, "").trim();
      var est = String(dataOrd[o][ioEst]).toUpperCase();
      if (codOrd == codigo && est != "CANCELADO" && est != "CERRADA") {
        var key = String(dataOrd[o][ioSerie]) + "." + ("0000" + dataOrd[o][ioOrden]).slice(-4);
        if (!ordenesMap[key]) {
          ordenesMap[key] = { nombre: key, pedido: dataOrd[o][ioPed], procesos: [] };
        }
        ordenesMap[key].procesos.push({
          idOrden: dataOrd[o][ioID],
          sec: Number(dataOrd[o][ioSec]),
          proceso: dataOrd[o][ioProc],
          solicitado: Number(dataOrd[o][ioSol]) || 0,
          produccion: [],
          totalProd: 0
        });
      }
    }
    
    // 2. Buscar producción
    var dataProd = sheetProd.getDataRange().getValues();
    var hProd = dataProd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ipOrden=hProd.indexOf("ID_ORDEN"); if(ipOrden<0) ipOrden=hProd.indexOf("ORDEN");
    var ipFecha=hProd.indexOf("FECHA"), ipTurno=hProd.indexOf("TURNO");
    var ipLote=hProd.indexOf("LOTE"), ipMaq=hProd.indexOf("MAQUINA");
    var ipProd=hProd.indexOf("PRODUCIDO"); if(ipProd<0) ipProd=hProd.indexOf("KILOS");
    var ipProceso=hProd.indexOf("PROCESO");
    
    var tz = Session.getScriptTimeZone();
    
    // Mapear producción por ID_ORDEN
    var prodMap = {};
    for (var p=1; p<dataProd.length; p++) {
      var idOrd = String(dataProd[p][ipOrden]).trim();
      if (!prodMap[idOrd]) prodMap[idOrd] = [];
      var fRaw = dataProd[p][ipFecha];
      var fStr = (fRaw instanceof Date) ? Utilities.formatDate(fRaw, tz, "dd/MM/yy") : String(fRaw);
      prodMap[idOrd].push({
        fecha: fStr,
        turno: dataProd[p][ipTurno] || "-",
        lote: dataProd[p][ipLote] || "-",
        maquina: dataProd[p][ipMaq] || "-",
        producido: Number(dataProd[p][ipProd]) || 0
      });
    }
    
    // 3. Unir producción con órdenes
    var resultado = { ordenes: [] };
    for (var key in ordenesMap) {
      var ord = ordenesMap[key];
      ord.procesos.sort(function(a,b){ return a.sec - b.sec; });
      ord.procesos.forEach(function(proc) {
        var prods = prodMap[String(proc.idOrden)] || [];
        proc.produccion = prods;
        var total = 0;
        prods.forEach(function(r){ total += r.producido; });
        proc.totalProd = Math.round(total * 100) / 100;
      });
      resultado.ordenes.push(ord);
    }
    
    return JSON.stringify(resultado);
  } catch(e) { return JSON.stringify({ ordenes: [], error: e.toString() }); }
}

// 4. GUARDAR CAMBIOS (EDICIÓN)
function guardarCambiosProgramador(cambios) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetPlan = ss.getSheetByName("PLANIFICACION_STOCK");
  var dataPlan = sheetPlan.getDataRange().getValues();
  var hPlan = dataPlan[0];
  var idxID = hPlan.indexOf("ID");
  var idxStatus = hPlan.indexOf("ESTATUS_REVISION");
  var idxDetalle = hPlan.indexOf("DETALLE_ACCION");

  var mapFilas = {};
  for(var i=1; i<dataPlan.length; i++) mapFilas[String(dataPlan[i][idxID])] = i + 1;

  var log = [];

  for (var k=0; k<cambios.length; k++) {
     var item = cambios[k];
     var fila = mapFilas[item.ID];
     if (!fila) continue;

     // CASO 1: RECHAZADO -> Solo cambiar estatus (desaparece de vista)
     if (item.ESTATUS_REVISION == "RECHAZADO") {
        sheetPlan.getRange(fila, idxStatus + 1).setValue("RECHAZADO");
     }
     
     // CASO 2: ACEPTADO -> EJECUTAR LÓGICA
     else if (item.ESTATUS_REVISION == "ACEPTADO") {
        // LLAMAMOS A LA NUEVA FUNCIÓN CON NOMBRE DISTINTO
        var resultado;
        // CANCELAR EX: Procesar cancelaciones y ediciones desde el modal
        if (item.ACCION_SUGERIDA == "CANCELAR EX" && (item.CANCEL_PEDIDOS || item.MODIFY_PEDIDOS)) {
           resultado = procesarCancelarEx(item);
        } else {
           resultado = procesarAjusteMRP(item);
        }
        
        if (resultado.success) {
           sheetPlan.getRange(fila, idxStatus + 1).setValue("PROCESADO");
           sheetPlan.getRange(fila, idxDetalle + 1).setValue(resultado.msg);
           log.push("✅ " + item.CODIGO + ": " + resultado.msg);
        } else {
           sheetPlan.getRange(fila, idxStatus + 1).setValue("PENDIENTE"); 
           sheetPlan.getRange(fila, idxDetalle + 1).setValue("ERROR: " + resultado.msg);
           log.push("❌ " + item.CODIGO + ": " + resultado.msg);
        }
     }
  }
  
  return "Proceso Terminado:\n" + log.join("\n");
}

function procesarAjusteMRP(item) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetPed = ss.getSheetByName("PEDIDOS");
    var sheetCod = ss.getSheetByName("CODIGOS"); 
    
    var accion = item.ACCION_SUGERIDA;
    var cantidadDelta = Number(item.CANTIDAD_ACCION); 
    var pedidoTargetID = "";

    // A. SELECCIONAR PEDIDO
    if (item.TEMP_SEL_PEDIDO) { pedidoTargetID = item.TEMP_SEL_PEDIDO; } 
    else {
       var rawPed = item.ID_ULTIMO_PEDIDO;
       // Limpieza si viene como JSON string
       if (String(rawPed).startsWith("[")) {
           try { var arr = JSON.parse(rawPed); if (arr.length > 0) pedidoTargetID = arr[0].id; } catch(e) {}
       } else { pedidoTargetID = rawPed; }
    }

    // CASO NUEVO PEDIDO
    if (accion == "NUEVO PEDIDO") {
       var dataCod = sheetCod.getDataRange().getValues(); 
       var unidadEncontrada = "PZA"; 
       var targetCod = String(item.CODIGO).trim().toUpperCase();
       
       for(var c=1; c<dataCod.length; c++) { 
           if(String(dataCod[c][0]).trim().toUpperCase() == targetCod) { 
               var u = dataCod[c][2]; if(u) unidadEncontrada = String(u).toUpperCase().trim(); 
               break; 
           } 
       }
       
       var dataPed = sheetPed.getDataRange().getValues(); var maxTem = 0;
       for(var p=1; p<dataPed.length; p++) { var val = String(dataPed[p][1]).toUpperCase(); if(val.startsWith("TEM-")) { var num = parseInt(val.split("-")[1]); if(!isNaN(num) && num > maxTem) maxTem = num; } }
       var nuevoPedidoName = "TEM-" + (maxTem + 1); var hoy = new Date(); hoy.setHours(0,0,0,0);
       
       sheetPed.appendRow([ Utilities.getUuid(), nuevoPedidoName, hoy, "'" + String(item.CODIGO).padStart(9, "0"), item.DESCRIPCION, 1, cantidadDelta, unidadEncontrada, "ABIERTO" ]);
       return { success: true, msg: "Creado " + nuevoPedidoName + " (" + cantidadDelta + ")" };
    }

    if (!pedidoTargetID || pedidoTargetID == "-" || pedidoTargetID == "undefined") return { success: false, msg: "No se seleccionó un pedido válido." };
    if (accion == "CANCELAR EX") return { success: true, msg: "Operación manual requerida para cancelar." };

    // B. BUSCAR ORDENES CANDIDATAS
    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ioPed=hOrd.indexOf("PEDIDO"), ioCant=hOrd.indexOf("CANTIDAD"), ioSol=hOrd.indexOf("SOLICITADO");
    var ioUni=hOrd.indexOf("UNIDAD"), ioPeso=hOrd.indexOf("PESO"), ioEst=hOrd.indexOf("ESTADO");
    var ioSerie=hOrd.indexOf("SERIE"), ioOrden=hOrd.indexOf("ORDEN");
    var ioPart=hOrd.indexOf("PARTIDA"); if(ioPart==-1) ioPart=2;
    var ioTipo=hOrd.indexOf("TIPO"), ioLong=hOrd.indexOf("LONGITUD"), ioDesc=hOrd.indexOf("DESCRIPCION");

    // AGRUPAR FILAS POR Serie.Orden (cada orden tiene N procesos = N filas)
    var ordenesMap = {};
    var pesoUnitario = 0; var unidad = ""; var partidaObjetivo = ""; 
    var longitudValor = 0; var esVarilla = false;

    for (var i=1; i<dataOrd.length; i++) {
       var est = String(dataOrd[i][ioEst]).toUpperCase();
       if (String(dataOrd[i][ioPed]).trim() == String(pedidoTargetID).trim() && est != "CANCELADO" && est != "CERRADA") {
          var key = String(dataOrd[i][ioSerie]) + "." + String(dataOrd[i][ioOrden]);
          
          if (!ordenesMap[key]) {
             ordenesMap[key] = {
                key: key,
                cant: Number(dataOrd[i][ioCant]) || 0,
                partida: dataOrd[i][ioPart],
                serie: dataOrd[i][ioSerie],
                orden: dataOrd[i][ioOrden],
                filas: [] // Todas las filas (procesos) de esta orden
             };
          }
          ordenesMap[key].filas.push(i); // Índice de fila en la hoja

          // Tomamos datos maestros de la primera que encontremos
          if (unidad === "") {
             unidad = String(dataOrd[i][ioUni]).toUpperCase().trim();
             pesoUnitario = Number(dataOrd[i][ioPeso]) || 0;
             partidaObjetivo = dataOrd[i][ioPart];
             // Detectar VARILLA y obtener longitud para regla de oro
             var tipoVal = (ioTipo > -1) ? String(dataOrd[i][ioTipo]).toUpperCase() : "";
             var descVal = (ioDesc > -1) ? String(dataOrd[i][ioDesc]).toUpperCase() : "";
             esVarilla = (tipoVal.includes("VARILLA") || tipoVal.includes("VAR") || descVal.includes("VARILLA"));
             if (ioLong > -1) {
                var matchL = String(dataOrd[i][ioLong]).match(/[\d\.]+/);
                longitudValor = matchL ? parseFloat(matchL[0]) : 0;
             }
          }
       }
    }

    var ordenesCandidatas = Object.values(ordenesMap);
    if (ordenesCandidatas.length === 0) return { success: false, msg: "No hay órdenes vivas para ajustar." };

    // ORDENAR POR CANTIDAD DESCENDENTE (Para restar a la más grande primero)
    ordenesCandidatas.sort(function(a,b) { return b.cant - a.cant; });

    // C. LOGICA DE CASCADA (RESTA / SUMA) — ACTUALIZA TODOS LOS PROCESOS
    var remanente = cantidadDelta; 
    var esResta = (accion == "CANT A MENOS");
    var cambiosRealizados = false;

    for(var k=0; k<ordenesCandidatas.length; k++) {
        if (remanente <= 0) break;

        var cand = ordenesCandidatas[k];
        var cantActual = cand.cant;
        var nuevaCant = 0;

        if (esResta) {
            if (cantActual >= remanente) {
                nuevaCant = cantActual - remanente;
                remanente = 0;
            } else {
                nuevaCant = 0;
                remanente = remanente - cantActual; 
            }
        } else {
            // SUMA: agregamos todo a la primera (la más grande)
            nuevaCant = cantActual + remanente;
            remanente = 0;
        }

        // Aplicar cambio en TODAS LAS FILAS (procesos) de esta orden
        if (cantActual !== nuevaCant) {
            // Calcular SOLICITADO con Regla de Oro completa
            var nuevoSolicitado = nuevaCant;
            if ((unidad.includes("PZA") || unidad.includes("CTO")) && pesoUnitario > 0) {
                if (esVarilla && longitudValor > 0) {
                    nuevoSolicitado = nuevaCant * pesoUnitario * longitudValor;
                } else {
                    nuevoSolicitado = nuevaCant * pesoUnitario;
                }
            }
            nuevoSolicitado = Math.round(nuevoSolicitado * 100) / 100;

            // Actualizar CADA proceso de la orden
            for (var f=0; f<cand.filas.length; f++) {
                var fila = cand.filas[f] + 1; // +1 porque getRange es 1-based
                sheetOrd.getRange(fila, ioCant + 1).setValue(nuevaCant); 
                sheetOrd.getRange(fila, ioSol + 1).setValue(nuevoSolicitado);
            }
            
            cand.cant = nuevaCant;
            cambiosRealizados = true;
        }
    }

    if (esResta && remanente > 0) {
        return { success: false, msg: "No hay suficiente cantidad en órdenes activas para restar todo." };
    }

    // D. RECALCULAR TOTAL PEDIDO (SUMANDO TODAS LAS PARTIDAS VIVAS)
    // Usamos un mapa para no sumar duplicados de procesos (misma serie.orden)
    var mapaOrdenesUnicas = {};
    var nuevoTotalPedido = 0;

    // Sumamos las candidatas (que ya tienen los valores nuevos)
    for(var c of ordenesCandidatas) {
        var key = c.serie + "." + c.orden;
        if(!mapaOrdenesUnicas[key]) {
            mapaOrdenesUnicas[key] = true;
            nuevoTotalPedido += c.cant;
        }
    }

    // E. ACTUALIZAR CABECERA PEDIDO
    var dataPed = sheetPed.getDataRange().getValues();
    var hPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ipNom = hPed.indexOf("PEDIDO"); if(ipNom==-1)ipNom=1;
    var ipPart = hPed.indexOf("PARTIDA"); if(ipPart==-1)ipPart=5;
    var ipCant = hPed.indexOf("CANTIDAD"); if(ipCant==-1)ipCant=6;
    
    var actualizado = false;
    for (var p=1; p<dataPed.length; p++) {
       var n = String(dataPed[p][ipNom]).trim();
       var pt = dataPed[p][ipPart];
       // Coincidir Pedido y Partida
       if (n == String(pedidoTargetID).trim() && (pt == partidaObjetivo || Math.trunc(Number(pt)) == Math.trunc(Number(partidaObjetivo)))) {
          sheetPed.getRange(p+1, ipCant + 1).setValue(nuevoTotalPedido);
          actualizado = true;
          break;
       }
    }

    return { success: true, msg: "Ajustado Pedido " + pedidoTargetID + " a " + nuevoTotalPedido };

  } catch(e) { return { success: false, msg: "Error Fatal: " + e.toString() }; }
}

// FUNCION AUXILIAR PARA CANCELAR TODO LO RELACIONADO A UN PEDIDO
function ejecutarCancelacionCascada(sheetPed, sheetOrd, pedidoID, hOrd) {
    var pedidoID = String(pedidoID).trim();

    // 1. Cancelar Cabecera PEDIDOS
    var dataPed = sheetPed.getDataRange().getValues();
    var hPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ipNom = hPed.indexOf("PEDIDO"); if(ipNom==-1)ipNom=1;
    var ipEst = hPed.indexOf("ESTADO"); if(ipEst==-1)ipEst=8; // Ajustar segun tu hoja
    
    var encontrado = false;
    for(var p=1; p<dataPed.length; p++) {
        if(String(dataPed[p][ipNom]).trim() == pedidoID) {
            sheetPed.getRange(p+1, ipEst + 1).setValue("CANCELADO");
            encontrado = true;
        }
    }

    // 2. Cancelar todas las ORDENES asociadas
    var dataOrd = sheetOrd.getDataRange().getValues();
    var ioPed = hOrd.indexOf("PEDIDO");
    var ioEst = hOrd.indexOf("ESTADO");

    var conteoOrdenes = 0;
    for(var o=1; o<dataOrd.length; o++) {
        if(String(dataOrd[o][ioPed]).trim() == pedidoID) {
            // Solo cancelamos si no está ya cancelada o cerrada para evitar re-escritura incesante
            var estActual = String(dataOrd[o][ioEst]).toUpperCase();
            if(estActual != "CANCELADO" && estActual != "CERRADA") {
                sheetOrd.getRange(o+1, ioEst + 1).setValue("CANCELADO");
                conteoOrdenes++;
            }
        }
    }

    return { success: true, msg: "PEDIDO " + pedidoID + " Y " + conteoOrdenes + " PROCESOS CANCELADOS." };
}

function procesarCancelarEx(item) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetPed = ss.getSheetByName("PEDIDOS");
    
    var log = [];
    var hOrd = sheetOrd.getRange(1, 1, 1, sheetOrd.getLastColumn()).getValues()[0]
               .map(function(h){ return String(h).toUpperCase().trim(); });
    
    // A. CANCELAR PEDIDOS SELECCIONADOS (pedido + todas sus órdenes)
    var cancelados = item.CANCEL_PEDIDOS || [];
    for (var c = 0; c < cancelados.length; c++) {
        var res = ejecutarCancelacionCascada(sheetPed, sheetOrd, cancelados[c], hOrd);
        log.push(res.msg);
    }
    
    // B. MODIFICAR CANTIDADES DE PEDIDOS NO CANCELADOS
    var modificados = item.MODIFY_PEDIDOS || [];
    if (modificados.length > 0) {
        var dataPed = sheetPed.getDataRange().getValues();
        var hPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
        var ipNom = hPed.indexOf("PEDIDO"); if(ipNom==-1) ipNom=1;
        var ipCant = hPed.indexOf("CANTIDAD"); if(ipCant==-1) ipCant=6;
        
        var dataOrd2 = sheetOrd.getDataRange().getValues();
        var hOrd2 = dataOrd2[0].map(function(h){ return String(h).toUpperCase().trim(); });
        var ioPed=hOrd2.indexOf("PEDIDO"), ioCant=hOrd2.indexOf("CANTIDAD");
        var ioSol=hOrd2.indexOf("SOLICITADO"), ioEst=hOrd2.indexOf("ESTADO");
        var ioUni=hOrd2.indexOf("UNIDAD"), ioPeso=hOrd2.indexOf("PESO");
        var ioTipo=hOrd2.indexOf("TIPO"), ioLong=hOrd2.indexOf("LONGITUD");
        var ioSerie=hOrd2.indexOf("SERIE"), ioOrden=hOrd2.indexOf("ORDEN");
        var ioDesc=hOrd2.indexOf("DESCRIPCION");
        
        for (var m = 0; m < modificados.length; m++) {
            var mod = modificados[m];
            var nuevaCant = mod.newQty;
            
            // Actualizar cabecera pedido
            for (var p=1; p<dataPed.length; p++) {
                if (String(dataPed[p][ipNom]).trim() == String(mod.id).trim()) {
                    sheetPed.getRange(p+1, ipCant+1).setValue(nuevaCant);
                    break;
                }
            }
            
            // Agrupar filas de ORDENES por Serie.Orden
            var ordenesMap = {};
            for (var o=1; o<dataOrd2.length; o++) {
                var est = String(dataOrd2[o][ioEst]).toUpperCase();
                if (String(dataOrd2[o][ioPed]).trim() == String(mod.id).trim() && est != "CANCELADO" && est != "CERRADA") {
                    var key = String(dataOrd2[o][ioSerie]) + "." + String(dataOrd2[o][ioOrden]);
                    if (!ordenesMap[key]) {
                        var uni = String(dataOrd2[o][ioUni]).toUpperCase().trim();
                        var peso = Number(dataOrd2[o][ioPeso]) || 0;
                        var tipo = (ioTipo > -1) ? String(dataOrd2[o][ioTipo]).toUpperCase() : "";
                        var desc = (ioDesc > -1) ? String(dataOrd2[o][ioDesc]).toUpperCase() : "";
                        var longTxt = (ioLong > -1) ? String(dataOrd2[o][ioLong]) : "";
                        var matchL = longTxt.match(/[\d\.]+/);
                        
                        ordenesMap[key] = { 
                            filas: [], unidad: uni, peso: peso,
                            esVarilla: (tipo.includes("VARILLA") || desc.includes("VARILLA")),
                            longitud: matchL ? parseFloat(matchL[0]) : 0,
                            numOrden: Number(dataOrd2[o][ioOrden]) || 0
                        };
                    }
                    ordenesMap[key].filas.push(o);
                }
            }
            
            // Solo modificar la ÚLTIMA orden (número de orden más alto)
            var ultimaKey = null; var maxNum = -1;
            for (var key in ordenesMap) {
                if (ordenesMap[key].numOrden > maxNum) { maxNum = ordenesMap[key].numOrden; ultimaKey = key; }
            }
            
            if (ultimaKey) {
                var info = ordenesMap[ultimaKey];
                var nuevoSol = nuevaCant;
                
                if ((info.unidad.includes("PZA") || info.unidad.includes("CTO")) && info.peso > 0) {
                    nuevoSol = info.esVarilla && info.longitud > 0 
                        ? nuevaCant * info.peso * info.longitud 
                        : nuevaCant * info.peso;
                }
                nuevoSol = Math.round(nuevoSol * 100) / 100;
                
                for (var f = 0; f < info.filas.length; f++) {
                    var idx = info.filas[f];
                    sheetOrd.getRange(idx+1, ioCant+1).setValue(nuevaCant);
                    sheetOrd.getRange(idx+1, ioSol+1).setValue(nuevoSol);
                }
            }
            
            log.push("EDITADO: " + mod.id + " → " + nuevaCant);
        }
    }
    
    return { success: true, msg: log.join(" | ") };
  } catch(e) { return { success: false, msg: "Error: " + e.toString() }; }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA Programador_InterHTML (MENU PROG. INTERACTIVO)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function obtenerDatosProgramadorInter(proceso) {
  try {
    var ss         = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd   = ss.getSheetByName("ORDENES");
    var sheetEst   = ss.getSheetByName("ESTANDARES");
    var sheetRutas = ss.getSheetByName("RUTAS");

    var dataEst = sheetEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var eID=hEst.indexOf("ID"), ePROC=hEst.indexOf("PROCESO"), eMAQ=hEst.indexOf("MAQUINA");
    var eVEL=hEst.indexOf("VELOCIDAD"), eUNID=hEst.indexOf("UNIDAD_VEL");
    var eEFIC=hEst.indexOf("EFICIENCIA"), eTURN=hEst.indexOf("TURNOS");
    var eGRUPO=hEst.indexOf("GRUPO"), eFOTO=hEst.indexOf("FOTO_MAQUINA");

    var maquinas=[], nombresMaqs=[], mapaEstandar={};
    for (var i=1; i<dataEst.length; i++) {
      var proc=String(dataEst[i][ePROC]||"").toUpperCase().trim();
      if (proc!==proceso.toUpperCase().trim()) continue;
      var maqNom=String(dataEst[i][eMAQ]||"").trim();
      if (!maqNom) continue;
      var obj={id:String(dataEst[i][eID]||i),maquina:maqNom,proceso:proc,
        velocidad:Number(dataEst[i][eVEL]||0),unidadVel:String(dataEst[i][eUNID]||"").toUpperCase(),
        eficiencia:Number(dataEst[i][eEFIC]||1),turnos:Number(dataEst[i][eTURN]||1),
        grupo:String(dataEst[i][eGRUPO]||""),foto:String(dataEst[i][eFOTO]||"")};
      maquinas.push(obj); nombresMaqs.push(maqNom.toUpperCase()); mapaEstandar[maqNom.toUpperCase()]=obj;
    }

    var dataRutas=sheetRutas.getDataRange().getValues();
    var hRutas=dataRutas[0].map(function(h){return String(h).toUpperCase().trim();});
    var rCOD=hRutas.indexOf("CODIGO"), rPROC=hRutas.indexOf("PROCESO"), rMAQ=hRutas.indexOf("MAQUINA");
    var rutasMaqMap={};
    for (var r=1; r<dataRutas.length; r++) {
      var rCodVal=String(dataRutas[r][rCOD]||"").trim();
      var rProcVal=String(dataRutas[r][rPROC]||"").trim().toUpperCase();
      var rMaqVal=String(dataRutas[r][rMAQ]||"").trim();
      if (!rCodVal||!rMaqVal) continue;
      var key=rCodVal+"|"+rProcVal;
      if (!rutasMaqMap[key]) {
        rutasMaqMap[key]=rMaqVal;
      } else {
        var lista=rutasMaqMap[key].split(",").map(function(s){return s.trim();});
        rMaqVal.split(",").forEach(function(m){var mt=m.trim();if(mt&&lista.indexOf(mt)===-1) lista.push(mt);});
        rutasMaqMap[key]=lista.join(", ");
      }
    }

    var dataOrd=sheetOrd.getDataRange().getValues();
    var hOrd=dataOrd[0].map(function(h){return String(h).toUpperCase().trim();});
    function col(nombre,fallback){var i=hOrd.indexOf(nombre);return i>=0?i:fallback;}
    var oID=col("ID",0),oPED=col("PEDIDO",1),oSERIE=col("SERIE",4),oORDEN=col("ORDEN",5);
    var oCOD=col("CODIGO",6),oDESC=col("DESCRIPCION",7),oSEC=col("SEC",10);
    var oPROC=col("PROCESO",11),oMAQ=col("MAQUINA",12),oSOL=col("SOLICITADO",13);
    var oPROD=col("PRODUCIDO",14),oEST=col("ESTADO",15),oTIPO=col("TIPO",19);
    var oDIA=col("DIAMETRO",20),oLONG=col("LONGITUD",21),oCUERDA=col("CUERDA",22);
    var oCUERPO=col("CUERPO",23),oACERO=col("ACERO",24),oPRIO=col("PRIORIDAD",26);
    var oFINI=col("FECHA_INICIO_PROG",27);
    var oMP=col("ALERTA_SECUENCIA",29);  // ← Col AD

    var ESTADOS_MUERTOS=["CANCELADO","TERMINADO","SOBREPRODUCCION","CERRADO"];
    var ordenes=[];

    for (var o=1; o<dataOrd.length; o++) {
      var row=dataOrd[o];
      var estado=String(row[oEST]||"").toUpperCase().trim();
      var maqOrd=String(row[oMAQ]||"").trim();
      var procO=String(row[oPROC]||"").toUpperCase().trim();
      if (ESTADOS_MUERTOS.indexOf(estado)>-1) continue;
      if (procO!==proceso.toUpperCase().trim()) continue;
      var maqsOrden=maqOrd.split(",").map(function(s){return s.toUpperCase().trim();});
      var enProceso=maqsOrden.some(function(m){return nombresMaqs.indexOf(m)>-1;});
      if (!enProceso) continue;

      var sol=Math.max(Number(row[oSOL]||0),0);
      var prod=Math.max(Number(row[oPROD]||0),0);
      var pend=Math.max(sol-prod,0);
      var avance=sol>0?Math.round(Math.min(prod/sol,1)*100):0;
      var maqKey=maqsOrden[0]||"";
      var std=mapaEstandar[maqKey]||{};
      var vel=Number(std.velocidad||0), efic=Number(std.eficiencia||1);
      var turnos=Number(std.turnos||1), unidVel=String(std.unidadVel||"").toUpperCase();
      var hpd=turnos===3?22.5:turnos===2?14.5:7.5;
      var velReal=vel*efic, horasEst=0, diasEst=0;
      if (velReal>0&&pend>0) {
        horasEst=unidVel.indexOf("MIN")>-1?(pend/velReal)/60:(pend/velReal);
        horasEst=Math.max(Math.round(horasEst*10)/10,0);
        diasEst=Math.max(Math.round((horasEst/hpd)*100)/100,0);
      }
      var fechaIni="";
      var fRaw=row[oFINI];
      if (fRaw instanceof Date){var mm=fRaw.getMonth()+1,dd=fRaw.getDate();fechaIni=fRaw.getFullYear()+"-"+(mm<10?"0":"")+mm+"-"+(dd<10?"0":"")+dd;}

      var codigo=String(row[oCOD]||"").trim();
      var maqPermKey=codigo+"|"+proceso.toUpperCase().trim();
      ordenes.push({
        id:                String(row[oID]),
        pedido:            String(row[oPED]||""),
        serie:             String(row[oSERIE]||""),
        orden:             ("0000"+String(row[oORDEN]||"")).slice(-4),
        codigo:            codigo,
        desc:              String(row[oDESC]||""),
        maquina:           maqsOrden[0]||maqOrd,
        proceso:           procO,
        sec:               Number(row[oSEC]||0),
        estado:            estado,
        prioridad:         Number(row[oPRIO]||999),
        sol:               sol, prod: prod, pend: pend, avance: avance,
        horasEst:          horasEst, diasEst: diasEst,
        tipo:              String(row[oTIPO]||""),
        dia:               String(row[oDIA]||""),
        long:              String(row[oLONG]||""),
        cuerda:            String(row[oCUERDA]||""),
        cuerpo:            String(row[oCUERPO]||""),
        acero:             String(row[oACERO]||""),
        fechaIni:          fechaIni,
        maquinasPermitidas: rutasMaqMap[maqPermKey]||maqOrd,
        mp:                String(row[oMP]||"").trim().toUpperCase()  // ← NUEVO
      });
    }

    ordenes.sort(function(a,b){
      if(a.maquina<b.maquina) return -1;
      if(a.maquina>b.maquina) return 1;
      return a.prioridad-b.prioridad;
    });

    return JSON.stringify({success:true,maquinas:maquinas,ordenes:ordenes});
  } catch(e) {
    Logger.log("obtenerDatosProgramadorInter ERROR: "+e.message+"\n"+e.stack);
    return JSON.stringify({success:false,msg:e.message});
  }
}

// ── 1. OBTENER LISTA DE PROCESOS ──────────────────────────────────────────────────
function obtenerProcesosParaProgramador() {
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetEst = ss.getSheetByName("ESTANDARES");
    var data     = sheetEst.getDataRange().getValues();
    var headers  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var idxProc  = headers.indexOf("PROCESO");

    var procesosSet = {};
    for (var i = 1; i < data.length; i++) {
      var p = String(data[i][idxProc] || "").trim();
      if (p) procesosSet[p] = true;
    }

    var lista = Object.keys(procesosSet).sort().map(function(p){ return { nombre: p }; });
    return JSON.stringify({ success: true, procesos: lista });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 2. OBTENER DATOS DEL GANTT PARA UN PROCESO ────────────────────────────────────
// Devuelve { success, maquinas: [...], ordenes: [...] }
function obtenerDatosProgramadorInter(proceso) {
  try {
    var ss          = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd    = ss.getSheetByName("ORDENES");
    var sheetEst    = ss.getSheetByName("ESTANDARES");
    var sheetRutas  = ss.getSheetByName("RUTAS");

    // ── A. MÁQUINAS DEL PROCESO (desde ESTANDARES) ───────────────────────────
    var dataEst  = sheetEst.getDataRange().getValues();
    var hEst     = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var eID      = hEst.indexOf("ID");
    var ePROC    = hEst.indexOf("PROCESO");
    var eMAQ     = hEst.indexOf("MAQUINA");
    var eVEL     = hEst.indexOf("VELOCIDAD");
    var eUNID    = hEst.indexOf("UNIDAD_VEL");
    var eEFIC    = hEst.indexOf("EFICIENCIA");
    var eTURN    = hEst.indexOf("TURNOS");
    var eGRUPO   = hEst.indexOf("GRUPO");
    var eFOTO    = hEst.indexOf("FOTO_MAQUINA");

    var maquinas     = [];
    var nombresMaqs  = [];
    var mapaEstandar = {};

    for (var i = 1; i < dataEst.length; i++) {
      var proc = String(dataEst[i][ePROC] || "").toUpperCase().trim();
      if (proc !== proceso.toUpperCase().trim()) continue;
      var maqNom = String(dataEst[i][eMAQ] || "").trim();
      if (!maqNom) continue;

      var obj = {
        id:         String(dataEst[i][eID]    || i),
        maquina:    maqNom,
        proceso:    proc,
        velocidad:  Number(dataEst[i][eVEL]   || 0),
        unidadVel:  String(dataEst[i][eUNID]  || "").toUpperCase(),
        eficiencia: Number(dataEst[i][eEFIC]  || 1),
        turnos:     Number(dataEst[i][eTURN]  || 1),
        grupo:      String(dataEst[i][eGRUPO] || ""),
        foto:       String(dataEst[i][eFOTO]  || "")
      };
      maquinas.push(obj);
      nombresMaqs.push(maqNom.toUpperCase());
      mapaEstandar[maqNom.toUpperCase()] = obj;
    }

    // ── B. MAPA DE MÁQUINAS PERMITIDAS (RUTAS col F) ─────────────────────────
    // RUTAS: B=CODIGO, E=PROCESO, F=MAQUINA
    var dataRutas = sheetRutas.getDataRange().getValues();
    var hRutas    = dataRutas[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var rCOD  = hRutas.indexOf("CODIGO");   // col B = 1
    var rPROC = hRutas.indexOf("PROCESO");  // col E = 4
    var rMAQ  = hRutas.indexOf("MAQUINA");  // col F = 5

    var rutasMaqMap = {};
    for (var r = 1; r < dataRutas.length; r++) {
      var rCodVal  = String(dataRutas[r][rCOD]  || "").trim();
      var rProcVal = String(dataRutas[r][rPROC] || "").trim().toUpperCase();
      var rMaqVal  = String(dataRutas[r][rMAQ]  || "").trim();
      if (!rCodVal || !rMaqVal) continue;
      var key = rCodVal + "|" + rProcVal;
      if (!rutasMaqMap[key]) {
        rutasMaqMap[key] = rMaqVal;
      } else {
        var lista = rutasMaqMap[key].split(",").map(function(s){return s.trim();});
        rMaqVal.split(",").forEach(function(m){
          var mt = m.trim();
          if (mt && lista.indexOf(mt) === -1) lista.push(mt);
        });
        rutasMaqMap[key] = lista.join(", ");
      }
    }

    // ── C. ÓRDENES VIVAS DEL PROCESO ─────────────────────────────────────────
    // Columnas ORDENES (índices fijos de tu estructura):
    // A=0:ID  B=1:PEDIDO  E=4:SERIE  F=5:ORDEN  G=6:CODIGO  H=7:DESCRIPCION
    // K=10:SEC  L=11:PROCESO  M=12:MAQUINA  N=13:SOLICITADO  O=14:PRODUCIDO
    // P=15:ESTADO  T=19:TIPO  U=20:DIAMETRO  V=21:LONGITUD  W=22:CUERDA
    // X=23:CUERPO  Y=24:ACERO  AA=26:PRIORIDAD  AB=27:FECHA_INICIO_PROG

    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });

    function col(nombre, fallback) {
      var idx = hOrd.indexOf(nombre);
      return idx >= 0 ? idx : fallback;
    }
    var oID     = col("ID",               0);
    var oPED    = col("PEDIDO",            1);
    var oSERIE  = col("SERIE",             4);
    var oORDEN  = col("ORDEN",             5);
    var oCOD    = col("CODIGO",            6);
    var oDESC   = col("DESCRIPCION",       7);
    var oSEC    = col("SEC",              10);
    var oPROC   = col("PROCESO",          11);
    var oMAQ    = col("MAQUINA",          12);
    var oSOL    = col("SOLICITADO",       13);
    var oPROD   = col("PRODUCIDO",        14);
    var oEST    = col("ESTADO",           15);
    var oTIPO   = col("TIPO",             19);
    var oDIA    = col("DIAMETRO",         20);
    var oLONG   = col("LONGITUD",         21);
    var oCUERDA = col("CUERDA",           22);
    var oCUERPO = col("CUERPO",           23);
    var oACERO  = col("ACERO",            24);
    var oPRIO   = col("PRIORIDAD",        26);
    var oFINI   = col("FECHA_INICIO_PROG",27);

    var ESTADOS_MUERTOS = ["CANCELADO","TERMINADO","SOBREPRODUCCION","CERRADO"];
    var ordenes = [];

    for (var o = 1; o < dataOrd.length; o++) {
      var row    = dataOrd[o];
      var estado = String(row[oEST]  || "").toUpperCase().trim();
      var maqOrd = String(row[oMAQ]  || "").trim();
      var procO  = String(row[oPROC] || "").toUpperCase().trim();

      if (ESTADOS_MUERTOS.indexOf(estado) > -1)   continue;
      if (procO !== proceso.toUpperCase().trim())  continue;

      // La máquina debe estar en las del proceso
      var maqsOrden = maqOrd.split(",").map(function(s){return s.toUpperCase().trim();});
      var enProceso = maqsOrden.some(function(m){ return nombresMaqs.indexOf(m) > -1; });
      if (!enProceso) continue;

      var sol    = Math.max(Number(row[oSOL]  || 0), 0);
      var prod   = Math.max(Number(row[oPROD] || 0), 0);
      var pend   = Math.max(sol - prod, 0);          // FIX 6: nunca negativo
      var avance = sol > 0 ? Math.round(Math.min(prod / sol, 1) * 100) : 0;

      // Calcular horas/días con ESTANDARES
      var maqKey   = maqsOrden[0] || "";
      var std      = mapaEstandar[maqKey] || {};
      var vel      = Number(std.velocidad  || 0);
      var efic     = Number(std.eficiencia || 1);
      var turnos   = Number(std.turnos     || 1);
      var unidVel  = String(std.unidadVel  || "").toUpperCase();
      var hpd      = turnos === 3 ? 22.5 : turnos === 2 ? 14.5 : 7.5;
      var velReal  = vel * efic;
      var horasEst = 0, diasEst = 0;

      if (velReal > 0 && pend > 0) {
        horasEst = unidVel.indexOf("MIN") > -1
          ? (pend / velReal) / 60
          : (pend / velReal);
        horasEst = Math.max(Math.round(horasEst * 10) / 10, 0);
        diasEst  = Math.max(Math.round((horasEst / hpd) * 100) / 100, 0);
      }

      // Fecha inicio fija
      var fechaIni = "";
      var fRaw = row[oFINI];
      if (fRaw instanceof Date) {
        var mm = fRaw.getMonth()+1, dd = fRaw.getDate();
        fechaIni = fRaw.getFullYear()+"-"+(mm<10?"0":"")+mm+"-"+(dd<10?"0":"")+dd;
      }

      // Máquinas permitidas desde RUTAS
      var codigo = String(row[oCOD] || "").trim();
      var maqPermKey = codigo + "|" + proceso.toUpperCase().trim();
      var maquinasPermitidas = rutasMaqMap[maqPermKey] || maqOrd;

      ordenes.push({
        id:                 String(row[oID]),
        pedido:             String(row[oPED]    || ""),
        serie:              String(row[oSERIE]  || ""),
        orden:              ("0000" + String(row[oORDEN] || "")).slice(-4),
        codigo:             codigo,
        desc:               String(row[oDESC]   || ""),
        maquina:            maqsOrden[0] || maqOrd,
        proceso:            procO,
        sec:                Number(row[oSEC]    || 0),
        estado:             estado,
        prioridad:          Number(row[oPRIO]   || 999),
        sol:                sol,
        prod:               prod,
        pend:               pend,
        avance:             avance,
        horasEst:           horasEst,
        diasEst:            diasEst,
        tipo:               String(row[oTIPO]   || ""),
        dia:                String(row[oDIA]    || ""),
        long:               String(row[oLONG]   || ""),
        cuerda:             String(row[oCUERDA] || ""),
        cuerpo:             String(row[oCUERPO] || ""),
        acero:              String(row[oACERO]  || ""),
        fechaIni:           fechaIni,
        maquinasPermitidas: maquinasPermitidas
      });
    }

    ordenes.sort(function(a, b) {
      if (a.maquina < b.maquina) return -1;
      if (a.maquina > b.maquina) return  1;
      return a.prioridad - b.prioridad;
    });

    return JSON.stringify({ success: true, maquinas: maquinas, ordenes: ordenes });

  } catch(e) {
    Logger.log("obtenerDatosProgramadorInter ERROR: " + e.message + "\n" + e.stack);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 3. GUARDAR FECHA INICIO MANUAL ───────────────────────────────────────────────
function guardarFechaIniOrdenProg(idOrden, fecha) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("ORDENES");
    var data = sh.getDataRange().getValues();
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colID   = hdr.indexOf("ID");
    var colFIni = hdr.indexOf("FECHA_INICIO_PROG"); if(colFIni===-1) colFIni=27;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]) === String(idOrden)) {
        sh.getRange(i+1, colFIni+1).setValue(new Date(fecha + "T12:00:00"));
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, msg: "Orden no encontrada: " + idOrden });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 4. GUARDAR EFICIENCIA Y TURNOS DE MÁQUINA ────────────────────────────────────
function guardarEficienciaTurnos(idEstandar, eficiencia, turnos) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("ESTANDARES");
    var data = sh.getDataRange().getValues();
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colID   = hdr.indexOf("ID");
    var colEFIC = hdr.indexOf("EFICIENCIA");
    var colTURN = hdr.indexOf("TURNOS");

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]) === String(idEstandar)) {
        sh.getRange(i+1, colEFIC+1).setValue(eficiencia);
        sh.getRange(i+1, colTURN+1).setValue(turnos);
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, msg: "Estándar no encontrado: " + idEstandar });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 5. OBTENER RUTA COMPLETA DE UNA ORDEN ────────────────────────────────────────
function obtenerRutaCompletaOrden(serie, orden) {
  try {
    var ss         = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd   = ss.getSheetByName("ORDENES");
    var sheetRutas = ss.getSheetByName("RUTAS");
    var sheetEst   = ss.getSheetByName("ESTANDARES");

    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oSERIE  = hOrd.indexOf("SERIE");
    var oORDEN  = hOrd.indexOf("ORDEN");
    var oCOD    = hOrd.indexOf("CODIGO");
    var oSOL    = hOrd.indexOf("SOLICITADO");
    var oPROD   = hOrd.indexOf("PRODUCIDO");
    var oEST    = hOrd.indexOf("ESTADO");
    var oFINI   = hOrd.indexOf("FECHA_INICIO_PROG"); if(oFINI===-1) oFINI=27;

    var ordenRow = null, codigo = "";
    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][oSERIE]) === String(serie) &&
          String(dataOrd[i][oORDEN]) === String(orden)) {
        ordenRow = dataOrd[i];
        codigo   = String(dataOrd[i][oCOD] || "").trim();
        break;
      }
    }
    if (!ordenRow) return JSON.stringify({ success: false, msg: "Orden no encontrada: "+serie+"."+orden });

    // Mapa de estándares
    var dataEst = sheetEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var eMaq = hEst.indexOf("MAQUINA"), eVel = hEst.indexOf("VELOCIDAD");
    var eEfic= hEst.indexOf("EFICIENCIA"), eTurn= hEst.indexOf("TURNOS");
    var eUnid= hEst.indexOf("UNIDAD_VEL"), eProc= hEst.indexOf("PROCESO");

    var mapaEst = {};
    for (var e = 1; e < dataEst.length; e++) {
      var mNom = String(dataEst[e][eMaq] || "").toUpperCase().trim();
      if (mNom) mapaEst[mNom] = {
        vel: Number(dataEst[e][eVel]  || 0),
        efic:Number(dataEst[e][eEfic] || 1),
        turn:Number(dataEst[e][eTurn] || 1),
        unid:String(dataEst[e][eUnid] || "").toUpperCase(),
        proc:String(dataEst[e][eProc] || "")
      };
    }

    // Pasos de la ruta
    var dataRutas = sheetRutas.getDataRange().getValues();
    var hRutas    = dataRutas[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var rCOD = hRutas.indexOf("CODIGO");
    var rSEC = hRutas.indexOf("SEC");
    var rPROC= hRutas.indexOf("PROCESO");
    var rMAQ = hRutas.indexOf("MAQUINA");

    var pasos = [];
    for (var rt = 1; rt < dataRutas.length; rt++) {
      if (String(dataRutas[rt][rCOD] || "").trim() !== codigo) continue;
      var maqRuta  = String(dataRutas[rt][rMAQ] || "").trim();
      var std      = mapaEst[maqRuta.toUpperCase()] || {};
      var sol      = Math.max(Number(ordenRow[oSOL]  || 0), 0);
      var prod     = Math.max(Number(ordenRow[oPROD] || 0), 0);
      var pend     = Math.max(sol - prod, 0);
      var velReal  = Number(std.vel||0) * Number(std.efic||1);
      var hpd      = Number(std.turn||1)===3?22.5:Number(std.turn||1)===2?14.5:7.5;
      var horasEst = 0, diasEst = 0;
      if (velReal > 0 && pend > 0) {
        horasEst = (std.unid||"").indexOf("MIN")>-1?(pend/velReal)/60:(pend/velReal);
        horasEst = Math.max(Math.round(horasEst*10)/10, 0);
        diasEst  = Math.max(Math.round((horasEst/hpd)*100)/100, 0);
      }
      var fRaw = ordenRow[oFINI];
      var fechaIniStr = "";
      if (fRaw instanceof Date) {
        var fm=fRaw.getMonth()+1, fd=fRaw.getDate();
        fechaIniStr = fRaw.getFullYear()+"-"+(fm<10?"0":"")+fm+"-"+(fd<10?"0":"")+fd;
      }
      pasos.push({
        sec:      Number(dataRutas[rt][rSEC]  || 0),
        proceso:  String(dataRutas[rt][rPROC] || ""),
        maquina:  maqRuta,
        estado:   String(ordenRow[oEST] || ""),
        sol: sol, prod: prod,
        horasEst: horasEst, diasEst: diasEst,
        fechaIni: fechaIniStr
      });
    }
    pasos.sort(function(a,b){ return a.sec - b.sec; });
    return JSON.stringify({ success: true, pasos: pasos });

  } catch(e) {
    Logger.log("obtenerRutaCompletaOrden ERROR: " + e.message);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 6. CAMBIAR ESTADO DE ORDEN ────────────────────────────────────────────────────
function cambiarEstadoOrdenProg(idOrden, nuevoEstado) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("ORDENES");
    var data = sh.getDataRange().getValues();
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colID  = hdr.indexOf("ID");
    var colEST = hdr.indexOf("ESTADO");
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colID]) === String(idOrden)) {
        sh.getRange(i+1, colEST+1).setValue(nuevoEstado);
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, msg: "Orden no encontrada: " + idOrden });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}


// ── 9. REACTIVAR ÓRDENES ─────────────────────────────────────────────────────────
function reactivarOrdenes(ids, proceso) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("ORDENES");
    var data = sh.getDataRange().getValues();
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colID  = hdr.indexOf("ID");
    var colEST = hdr.indexOf("ESTADO");
    var idsSet = ids.map(String);
    var count  = 0;
    for (var i = 1; i < data.length; i++) {
      if (idsSet.indexOf(String(data[i][colID])) > -1) {
        sh.getRange(i+1, colEST+1).setValue("EN PROCESO");
        count++;
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, msg: count + " orden(es) reactivada(s)." });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 10. GUARDAR FECHAS CALCULADAS (FECHA_INICIO_PROG / FECHA_FIN_PROG) ───────────
function guardarFechasProgOrdenes(fechas) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sh   = ss.getSheetByName("ORDENES");
    var data = sh.getDataRange().getValues();
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colID   = hdr.indexOf("ID");
    var colFIni = hdr.indexOf("FECHA_INICIO_PROG"); if(colFIni===-1) colFIni=27;
    var colFFin = hdr.indexOf("FECHA_FIN_PROG");    if(colFFin===-1) colFFin=28;
    var mapa = {};
    fechas.forEach(function(f){ mapa[String(f.id)] = f; });
    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][colID]);
      if (mapa[id]) {
        sh.getRange(i+1, colFIni+1).setValue(new Date(mapa[id].fIni));
        sh.getRange(i+1, colFFin+1).setValue(new Date(mapa[id].fFin));
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 11. RECALCULAR FECHAS TODOS LOS PROCESOS (botón Actualizar Todo) ─────────────
function recalcularFechasTodosLosProcesos() {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd= ss.getSheetByName("ORDENES");
    var shEst= ss.getSheetByName("ESTANDARES");
    var dOrd = shOrd.getDataRange().getValues();
    var hOrd = dOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var dEst = shEst.getDataRange().getValues();
    var hEst = dEst[0].map(function(h){ return String(h).toUpperCase().trim(); });

    var colID   = hOrd.indexOf("ID");
    var colMAQ  = hOrd.indexOf("MAQUINA");
    var colPRIO = hOrd.indexOf("PRIORIDAD");
    var colSOL  = hOrd.indexOf("SOLICITADO");
    var colPROD = hOrd.indexOf("PRODUCIDO");
    var colEST  = hOrd.indexOf("ESTADO");
    var colVEL  = hEst.indexOf("VELOCIDAD");
    var colEFIC = hEst.indexOf("EFICIENCIA");
    var colTURN = hEst.indexOf("TURNOS");
    var colUNID = hEst.indexOf("UNIDAD_VEL");
    var colEMAQ = hEst.indexOf("MAQUINA");
    var colFIni = hOrd.indexOf("FECHA_INICIO_PROG"); if(colFIni===-1) colFIni=27;
    var colFFin = hOrd.indexOf("FECHA_FIN_PROG");    if(colFFin===-1) colFFin=28;

    // Mapa de estándares
    var mapaEst = {};
    for (var e = 1; e < dEst.length; e++) {
      var mNom = String(dEst[e][colEMAQ]||"").toUpperCase().trim();
      if (mNom) mapaEst[mNom] = {
        vel:  Number(dEst[e][colVEL]  || 0),
        efic: Number(dEst[e][colEFIC] || 1),
        turn: Number(dEst[e][colTURN] || 1),
        unid: String(dEst[e][colUNID] || "").toUpperCase()
      };
    }

    // Hora de inicio según turno
    var ahora = new Date(), h = ahora.getHours();
    var base  = new Date(ahora);
    if (h >= 8  && h < 16) base.setHours(6, 30, 0, 0);
    else if (h >= 16)      base.setHours(14, 30, 0, 0);

    var TZ = Session.getScriptTimeZone();
    var DIAS_INH = ["2026-03-16","2026-04-02","2026-04-03","2026-04-04","2026-05-01",
      "2026-09-15","2026-09-16","2026-11-02","2026-11-16","2026-12-12",
      "2026-12-24","2026-12-25","2026-12-31"];

    function esInhabil(d){ return DIAS_INH.indexOf(Utilities.formatDate(d,TZ,"yyyy-MM-dd"))>-1; }
    function esNoLaboral(d){ return d.getDay()===0||esInhabil(d); }
    function addDias(d,n){ var r=new Date(d); r.setDate(r.getDate()+n); return r; }
    function calcFin(desde,horas,hpd){
      var rest=Math.max(horas,0), cur=new Date(desde);
      while(rest>0.001){ cur=addDias(cur,1); if(!esNoLaboral(cur)) rest-=(cur.getDay()===6?hpd*0.5:hpd); }
      return cur;
    }

    // Agrupar por máquina
    var MUERTOS = ["CANCELADO","TERMINADO","SOBREPRODUCCION","CERRADO"];
    var porMaq  = {};
    for (var i = 1; i < dOrd.length; i++) {
      var estado = String(dOrd[i][colEST]||"").toUpperCase().trim();
      if (MUERTOS.indexOf(estado)>-1) continue;
      var maq = String(dOrd[i][colMAQ]||"").trim().split(",")[0].toUpperCase().trim();
      if (!maq) continue;
      var sol  = Math.max(Number(dOrd[i][colSOL] ||0),0);
      var prod = Math.max(Number(dOrd[i][colPROD]||0),0);
      var pend = Math.max(sol-prod,0);
      var std  = mapaEst[maq]||{};
      var vel  = Number(std.vel||0)*Number(std.efic||1);
      var hpd  = Number(std.turn||1)===3?22.5:Number(std.turn||1)===2?14.5:7.5;
      var horasEst = 0;
      if(vel>0&&pend>0) horasEst=Math.max((std.unid||"").indexOf("MIN")>-1?(pend/vel)/60:(pend/vel),0);
      if(!porMaq[maq]) porMaq[maq]=[];
      porMaq[maq].push({rowIdx:i, prio:Number(dOrd[i][colPRIO]||999), horas:horasEst, hpd:hpd});
    }

    var updates = [];
    Object.keys(porMaq).forEach(function(maq){
      var ords = porMaq[maq].sort(function(a,b){return a.prio-b.prio;});
      var cur  = new Date(base); while(esNoLaboral(cur)) cur=addDias(cur,1);
      ords.forEach(function(o){
        var fi=new Date(cur), ff=calcFin(fi,o.horas,o.hpd);
        cur=new Date(ff);
        updates.push({rowIdx:o.rowIdx, fi:fi, ff:ff});
      });
    });

    updates.forEach(function(u){
      shOrd.getRange(u.rowIdx+1, colFIni+1).setValue(u.fi);
      shOrd.getRange(u.rowIdx+1, colFFin+1).setValue(u.ff);
    });
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, msg: updates.length+" órdenes actualizadas." });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 12. GUARDAR MÁQUINAS PERMITIDAS EN RUTAS (col F) ─────────────────────────────
function guardarMaquinasPermitidas(idOrden, codigoOrden, maqString, proceso) {
  try {
    var ss     = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shRutas= ss.getSheetByName("RUTAS");
    var data   = shRutas.getDataRange().getValues();
    var hdr    = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var colCOD = hdr.indexOf("CODIGO");
    var colPROC= hdr.indexOf("PROCESO");
    var colMAQ = hdr.indexOf("MAQUINA"); if(colMAQ===-1) colMAQ=5;
    var count  = 0;
    for (var i = 1; i < data.length; i++) {
      var cod  = String(data[i][colCOD]  || "").trim();
      var proc = String(data[i][colPROC] || "").trim().toUpperCase();
      if (cod === String(codigoOrden).trim() &&
          (!proceso || proc === proceso.toUpperCase().trim())) {
        shRutas.getRange(i+1, colMAQ+1).setValue(maqString);
        count++;
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, msg: count+" fila(s) en RUTAS actualizadas." });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 13. OBTENER ÓRDENES INACTIVAS DE UNA MÁQUINA (últimos N días) ────────────────
function obtenerOrdenesInactivasMaquina(maquina, dias, filtroEstado) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shO  = ss.getSheetByName("ORDENES");
    var data = shO.getDataRange().getValues();

    // Índices por posición (estructura fija de ORDENES)
    var iID   = 0;   // A: ID
    var iPed  = 1;   // B: PEDIDO
    var iFech = 3;   // D: FECHA_REG
    var iSerie= 4;   // E: SERIE
    var iOrden= 5;   // F: ORDEN  (número)
    var iCod  = 6;   // G: CODIGO
    var iProc = 11;  // L: PROCESO
    var iMaq  = 12;  // M: MAQUINA
    var iSol  = 13;  // N: SOLICITADO
    var iProd = 14;  // O: PRODUCIDO
    var iEst  = 15;  // P: ESTADO  ← fuente de verdad
    var iTipo = 19;  // T: TIPO
    var iDia  = 20;  // U: DIAMETRO
    var iLon  = 21;  // V: LONGITUD
    var iPrio = 26;  // AA: PRIORIDAD

    var maqUp        = String(maquina).toUpperCase().trim();
    var soloSobrep   = (typeof filtroEstado === "string") &&
                       filtroEstado.toUpperCase().trim() === "SOBREPRODUCCION";

    // Límite de fecha sólo para el historial (no sobreproducción)
    var fechaLimite  = null;
    if (!soloSobrep) {
      fechaLimite = new Date();
      fechaLimite.setDate(fechaLimite.getDate() - (Number(dias) || 90));
    }

    var ESTADOS_MUERTOS = ["TERMINADO","CANCELADO","SOBREPRODUCCION","CERRADO"];
    var ordenes = [];

    for (var i = 1; i < data.length; i++) {
      var r      = data[i];
      var estado = String(r[iEst]).toUpperCase().trim();
      var maqRow = String(r[iMaq]).toUpperCase().trim();

      if (!estado || !maqRow) continue;
      if (maqRow !== maqUp)   continue;

      if (soloSobrep) {
        // Modo sobreproducción: sólo ESTADO = SOBREPRODUCCION, sin filtro de fecha
        if (estado !== "SOBREPRODUCCION") continue;
      } else {
        // Modo historial: sólo estados muertos dentro del rango de días
        if (ESTADOS_MUERTOS.indexOf(estado) === -1) continue;
        if (fechaLimite) {
          var fReg = r[iFech] instanceof Date ? r[iFech] : new Date(r[iFech]);
          if (isNaN(fReg.getTime()) || fReg < fechaLimite) continue;
        }
      }

      var sol  = Number(r[iSol])  || 0;
      var prod = Number(r[iProd]) || 0;

      ordenes.push({
        id:      String(r[iID]),
        serie:   String(r[iSerie]),
        orden:   ("0000" + parseInt(r[iOrden])).slice(-4),  // siempre 4 dígitos
        pedido:  String(r[iPed]  || ""),
        codigo:  String(r[iCod]  || ""),
        proceso: String(r[iProc] || ""),
        maquina: String(r[iMaq]  || ""),
        estado:  estado,                  // lee directamente de col P
        tipo:    String(r[iTipo] || ""),
        dia:     String(r[iDia]  || ""),
        long:    String(r[iLon]  || ""),
        sol:     sol,
        prod:    prod,
        avance:  sol > 0 ? Math.round((prod / sol) * 100) : 0,
        prioridad: Number(r[iPrio]) || 999
      });
    }

    // Ordenar por prioridad
    ordenes.sort(function(a,b){ return a.prioridad - b.prioridad; });

    return JSON.stringify({ success:true, ordenes:ordenes });
  } catch(e) {
    return JSON.stringify({ success:false, msg:e.toString() });
  }
}

// ── 14. calcularActualizacionEjecucion — actualiza pedido + UNA orden elegida ──
function calcularActualizacionEjecucion() {
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetInv = ss.getSheetByName("INVENTARIO_EXTERNO");
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetPed = ss.getSheetByName("PEDIDOS");

    if (!sheetInv) throw new Error("No existe la hoja INVENTARIO_EXTERNO");
    if (!sheetPed) throw new Error("No existe la hoja PEDIDOS");
    if (!sheetOrd) throw new Error("No existe la hoja ORDENES");

    var dataInv = sheetInv.getDataRange().getValues();
    var dataPed = sheetPed.getDataRange().getValues();
    var hPed    = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdID    = hPed.indexOf("ID");
    var pdCOD   = hPed.indexOf("CODIGO");
    var pdCANT  = hPed.indexOf("CANTIDAD");
    var pdUNI   = hPed.indexOf("UNIDAD");
    var pdDESC  = hPed.indexOf("DESCRIPCION");
    var pdPESO  = hPed.indexOf("PESO");
    var pdLONG  = hPed.indexOf("LONGITUD");
    var pdEST   = hPed.indexOf("ESTADO");
    var pdSOL   = hPed.indexOf("SOLICITADO");
    var pdFOLIO = hPed.indexOf("FOLIO"); if (pdFOLIO === -1) pdFOLIO = 1;

    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oID     = hOrd.indexOf("ID");
    var oSOL    = hOrd.indexOf("SOLICITADO");
    var oCANT   = hOrd.indexOf("CANTIDAD");
    var oPROD   = hOrd.indexOf("PRODUCIDO");
    var oEST    = hOrd.indexOf("ESTADO");
    var oPED    = hOrd.indexOf("PEDIDO");
    var oUNI    = hOrd.indexOf("UNIDAD");
    var oPESO   = hOrd.indexOf("PESO");
    var oLONG   = hOrd.indexOf("LONGITUD");
    var oDESC   = hOrd.indexOf("DESCRIPCION");
    var oSERIE  = hOrd.indexOf("SERIE");
    var oORDEN  = hOrd.indexOf("ORDEN");

    var MUERTOS = ["CANCELADO","TERMINADO","SOBREPRODUCCION","CERRADO","ENTREGADO"];
    var cambios = [];

    function calcSolicitado(cantidad, unidad, desc, peso, longitud) {
      var u = String(unidad).toUpperCase().trim();
      if (u === "KG" || u === "ROL") return cantidad;
      var p = Number(peso) || 0;
      var sol = cantidad * p;
      if (String(desc).toUpperCase().indexOf("VARILLA") > -1) {
        var l = parseFloat(longitud) || 0;
        sol = cantidad * p * l;
      }
      return sol;
    }

    function calcNuevoEstado(producido, nuevoSol, estadoActual) {
      if (MUERTOS.indexOf(estadoActual) > -1) return estadoActual;
      if (producido <= 0) return "ABIERTA";
      if (producido >= nuevoSol) return "TERMINADO";
      return "EN PROCESO";
    }

    for (var i = 1; i < dataInv.length; i++) {
      var rowInv    = dataInv[i];
      var ejecucion = String(rowInv[5] || "").toUpperCase().trim();
      if (ejecucion !== "AUTO") continue;

      var codigo     = String(rowInv[0] || "").trim();
      var existencia = Number(rowInv[1] || 0);
      var minimo     = Number(rowInv[2] || 0);
      var maximo     = Number(rowInv[3] || 0);
      var aFabricar  = maximo - existencia;
      if (aFabricar <= 0) continue;

      for (var p = 1; p < dataPed.length; p++) {
        var rowPed = dataPed[p];
        if (String(rowPed[pdCOD] || "").trim() !== codigo) continue;
        var estPed = String(rowPed[pdEST] || "").toUpperCase().trim();
        if (MUERTOS.indexOf(estPed) > -1) continue;

        var pedidoId    = String(rowPed[pdID]    || "");
        var pedidoFolio = String(rowPed[pdFOLIO] || "");
        var _folio = pedidoFolio.toUpperCase();
        if (_folio.indexOf("QSQ") === -1 && _folio.indexOf("TEM") === -1) continue;

        var cantAntes   = Number(rowPed[pdCANT] || 0);
        var cantDespues = aFabricar;
        if (cantDespues === cantAntes) continue;

        var unidad        = String(rowPed[pdUNI]  || "");
        var desc          = String(rowPed[pdDESC] || "");
        var peso          = Number(rowPed[pdPESO] || 0);
        var longitud      = rowPed[pdLONG];
        var solPedDespues = calcSolicitado(cantDespues, unidad, desc, peso, longitud);

        // ── Agrupar órdenes vivas por SERIE+ORDEN usando folio del pedido ──
        var gruposOrden = {};
        for (var o = 1; o < dataOrd.length; o++) {
          var rowOrd = dataOrd[o];
          if (String(rowOrd[oPED] || "").trim() !== pedidoFolio.trim()) continue;
          var estOrd = String(rowOrd[oEST] || "").toUpperCase().trim();
          if (MUERTOS.indexOf(estOrd) > -1) continue;

          var serie = String(rowOrd[oSERIE] || "");
          var orden = String(rowOrd[oORDEN] || "");
          var key   = serie + "." + orden;

          if (!gruposOrden[key]) {
            gruposOrden[key] = {
              key:        key,
              idOrden:    String(rowOrd[oID]),
              cantActual: Number(rowOrd[oCANT] || 0),
              filas:      []
            };
          }
          gruposOrden[key].filas.push(o);
        }

        var claves = Object.keys(gruposOrden);
        // Si no hay órdenes vivas, igual mostrar el cambio del pedido sin filas de orden
        var claveElegida = claves.length ? claves[0] : null;

        // ── Elegir la orden con mayor CANTIDAD actual para recibir el ajuste ──
        for (var k = 1; k < claves.length; k++) {
          if (gruposOrden[claves[k]].cantActual > gruposOrden[claveElegida].cantActual) {
            claveElegida = claves[k];
          }
        }
        var ordenElegida = gruposOrden[claveElegida];

        // Suma de CANTIDAD de las órdenes NO elegidas
        var sumaOtras = 0;
        for (var k2 = 0; k2 < claves.length; k2++) {
          if (claves[k2] !== claveElegida) sumaOtras += gruposOrden[claves[k2]].cantActual;
        }

        // Nueva CANTIDAD para la orden elegida = cantDespues - sumaOtras (nunca negativo)
        var nuevaCantOrden = Math.max(0, cantDespues - sumaOtras);

        // ── Construir lista de procesos de la orden elegida (1 entrada por fila de proceso) ──
        var ordenesDisplay = [];
        var idsVistos = {};
        for (var f = 0; ordenElegida && f < ordenElegida.filas.length; f++) {
          var fi   = ordenElegida.filas[f];
          var rowO = dataOrd[fi];
          var idO  = String(rowO[oID]);
          if (idsVistos[idO]) continue;
          idsVistos[idO] = true;

          var uOrd          = String(rowO[oUNI]  || "") || unidad;
          var dOrd          = String(rowO[oDESC] || "") || desc;
          var pOrd          = Number(rowO[oPESO] || 0) || peso;
          var lOrd          = rowO[oLONG] || longitud;
          var prod          = Math.max(Number(rowO[oPROD] || 0), 0);
          var estOrdActual  = String(rowO[oEST] || "").toUpperCase().trim();
          var nuevoSol      = calcSolicitado(nuevaCantOrden, uOrd, dOrd, pOrd, lOrd);
          var nuevoEstado   = calcNuevoEstado(prod, nuevoSol, estOrdActual);

          ordenesDisplay.push({
            idOrden:       idO,
            pedido:        pedidoId,
            cantAntes:     Number(rowO[oCANT] || 0),
            cantDespues:   nuevaCantOrden,
            solAntes:      Number(rowO[oSOL] || 0),
            solDespues:    nuevoSol,
            estadoAntes:   estOrdActual,
            estadoDespues: nuevoEstado,
            razon:         "Ajuste auto stock"
          });
        }

        cambios.push({
          planId:          codigo,
          codigo:          codigo,
          descripcion:     desc,
          existencia:      existencia,
          minimo:          minimo,
          maximo:          maximo,
          aFabricar:       aFabricar,
          accion:          "AJUSTE AUTO",
          pedidoId:        pedidoId,
          pedidoFolio:     pedidoFolio,
          cantPedAntes:    cantAntes,
          cantPedDespues:  cantDespues,
          solPedDespues:   solPedDespues,
          ordenElegidaKey: claveElegida,
          nuevaCantOrden:  nuevaCantOrden,
          ordenes:         ordenesDisplay
        });
      }
    }

    return JSON.stringify({ success: true, cambios: cambios });
  } catch(e) {
    Logger.log("calcularActualizacionEjecucion ERROR: " + e.message);
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ══════════════════════════════════════════════════════════════════
// SINCRONIZAR INVENTARIO — COLATADO
// Calcula qué códigos de línea COLATADO necesitan fabricarse
// basado en MIN - EXIST de INVENTARIO_EXTERNO
// ══════════════════════════════════════════════════════════════════
function calcularSincronizarColatado() {
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shInv   = ss.getSheetByName("INVENTARIO_EXTERNO");
    var shOrd   = ss.getSheetByName("ORDENES");
    var shPed   = ss.getSheetByName("PEDIDOS");
    var shRutas = ss.getSheetByName("RUTAS");
    var shCod   = ss.getSheetByName("CODIGOS");

    var dataInv = shInv.getDataRange().getValues();
    var dataOrd = shOrd.getDataRange().getValues();
    var dataPed = shPed.getDataRange().getValues();
    var dataRut = shRutas.getDataRange().getValues();
    var dataCod = shCod ? shCod.getDataRange().getValues() : [];

    var hOrd = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oID   = hOrd.indexOf("ID");
    var oCOD  = hOrd.indexOf("CODIGO");
    var oSOL  = hOrd.indexOf("SOLICITADO");
    var oPROD = hOrd.indexOf("PRODUCIDO");
    var oEST  = hOrd.indexOf("ESTADO");
    var oPED  = hOrd.indexOf("PEDIDO");
    var oMAQ  = hOrd.indexOf("MAQUINA");
    var oSERIE= hOrd.indexOf("SERIE");
    var oORD  = hOrd.indexOf("ORDEN");
    var oCOD2 = hOrd.indexOf("CODIGO");
    var oDESC = hOrd.indexOf("DESCRIPCION");
    var oUNI  = hOrd.indexOf("UNIDAD");
    var oPESO = hOrd.indexOf("PESO");
    var oLONG = hOrd.indexOf("LONGITUD");
    var oCANT = hOrd.indexOf("CANTIDAD");
    var oTIPO = hOrd.indexOf("TIPO");
    var oDIA  = hOrd.indexOf("DIAMETRO");
    var oCUERDA = hOrd.indexOf("CUERDA");
    var oCUERPO = hOrd.indexOf("CUERPO");
    var oACERO  = hOrd.indexOf("ACERO");
    var oPROC   = hOrd.indexOf("PROCESO");
    var oPRIO   = hOrd.indexOf("PRIORIDAD");

    var hPed = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdID   = hPed.indexOf("ID");
    var pdFOL  = hPed.indexOf("FOLIO"); if (pdFOL === -1) pdFOL = 1;
    var pdCOD  = hPed.indexOf("CODIGO");
    var pdCANT = hPed.indexOf("CANTIDAD");
    var pdUNI  = hPed.indexOf("UNIDAD");
    var pdDESC = hPed.indexOf("DESCRIPCION");
    var pdPESO = hPed.indexOf("PESO");
    var pdLONG = hPed.indexOf("LONGITUD");
    var pdEST  = hPed.indexOf("ESTADO");
    var pdSOL  = hPed.indexOf("SOLICITADO");

    // Mapa código→venta desde CODIGOS (col A=fab, col F=venta)
    var codigoVentaMap = {};
    for (var cv = 1; cv < dataCod.length; cv++) {
      var cvF = String(dataCod[cv][0]||'').trim().toUpperCase();
      var cvV = String(dataCod[cv][5]||'').trim().toUpperCase();
      if (cvF && cvV) codigoVentaMap[cvF] = cvV;
    }

    // Mapa inventario externo: código → {exist, min, max}
    var invMap = {};
    for (var i = 1; i < dataInv.length; i++) {
      var ic = String(dataInv[i][0]||'').trim().toUpperCase();
      if (ic) invMap[ic] = {
        exist: Number(dataInv[i][1])||0,
        min:   Number(dataInv[i][2])||0,
        max:   Number(dataInv[i][3])||0
      };
    }

    // Función: obtener datos inventario por código de orden (igual que obtenerOrdenesPlanificador)
    function getInv(codigoOrden) {
      var cUp = String(codigoOrden||'').trim().toUpperCase();
      if (cUp.charAt(0) === '7') {
        var cV = codigoVentaMap[cUp] || '';
        if (cV && invMap[cV]) return invMap[cV];
      }
      return invMap[cUp] || null;
    }

    // Mapa ruta: codigo → pasos (solo COLATADO)
    var rutaMap = {};
    for (var r = 1; r < dataRut.length; r++) {
      var rCod  = String(dataRut[r][1]||'').trim();
      var rProc = String(dataRut[r][4]||'').trim().toUpperCase();
      if (rProc === 'COLATADO') {
        if (!rutaMap[rCod]) rutaMap[rCod] = [];
        rutaMap[rCod].push({
          sec:    dataRut[r][3],
          proceso: rProc,
          maquinas: String(dataRut[r][5]||'').split(',').map(function(m){ return m.trim(); }),
          cantLote: parseFloat(dataRut[r][22])||0,
          pt:     dataRut[r][12],
          venta:  dataRut[r][13],
          peso:   Number(dataRut[r][15])||0,
          tipo:   String(dataRut[r][6]||''),
          diam:   String(dataRut[r][7]||''),
          long:   String(dataRut[r][8]||''),
          cuerda: String(dataRut[r][9]||''),
          cuerpo: String(dataRut[r][10]||''),
          acero:  String(dataRut[r][11]||''),
          serie:  String(dataRut[r][14]||'P')
        });
      }
    }

    // Fecha de pedido: mapa folio → fecha
    var oPEDFECHA = hPed.indexOf("FECHA"); if (oPEDFECHA === -1) oPEDFECHA = 2;
    var mapFechaPed = {};
    for (var fp = 1; fp < dataPed.length; fp++) {
      var folK = String(dataPed[fp][pdFOL]||'');
      var fVal = dataPed[fp][oPEDFECHA];
      var fStr = '';
      if (fVal instanceof Date) {
        fStr = ('0'+fVal.getDate()).slice(-2) + '/' + ('0'+(fVal.getMonth()+1)).slice(-2) + '/' + fVal.getFullYear();
      } else if (fVal) { fStr = String(fVal); }
      mapFechaPed[folK] = fStr;
    }

    // Recolectar TODOS los códigos con órdenes COLATADO (cualquier estado, para mostrar historial completo)
    var MUERTOS_SINC = ['CERRADO'];  // solo excluimos CERRADO; mostramos todos los demás
    var codigosColatado = {};  // codigo → [ordenesDetalle]

    for (var o = 1; o < dataOrd.length; o++) {
      var proc = String(dataOrd[o][oPROC]||'').trim().toUpperCase();
      if (proc !== 'COLATADO') continue;
      var est = String(dataOrd[o][oEST]||'').trim().toUpperCase();
      if (MUERTOS_SINC.indexOf(est) > -1) continue;
      var cod = String(dataOrd[o][oCOD2]||'').trim();
      if (!cod) continue;
      if (!codigosColatado[cod]) codigosColatado[cod] = [];
      var pedFolio = String(dataOrd[o][oPED]||'');
      codigosColatado[cod].push({
        id:        String(dataOrd[o][oID]),
        pedido:    pedFolio,
        fechaPed:  mapFechaPed[pedFolio] || '',
        serie:     String(dataOrd[o][oSERIE]||''),
        orden:     Number(dataOrd[o][oORD])||0,
        sol:       Number(dataOrd[o][oSOL])||0,
        prod:      Number(dataOrd[o][oPROD])||0,
        estado:    est,
        maquina:   String(dataOrd[o][oMAQ]||''),
        cant:      Number(dataOrd[o][oCANT])||0,
        unidad:    String(dataOrd[o][oUNI]||''),
        desc:      String(dataOrd[o][oDESC]||''),
        peso:      Number(dataOrd[o][oPESO])||0,
        long:      String(dataOrd[o][oLONG]||'').trim(),
        tipo:      String(dataOrd[o][oTIPO]||''),
        dia:       String(dataOrd[o][oDIA]||''),
        cuerda:    String(dataOrd[o][oCUERDA]||''),
        cuerpo:    String(dataOrd[o][oCUERPO]||''),
        acero:     String(dataOrd[o][oACERO]||''),
        prioridad: Number(dataOrd[o][oPRIO])||999
      });
    }

    var resultados = [];

    Object.keys(codigosColatado).forEach(function(cod) {
      var inv = getInv(cod);
      if (!inv) return;               // sin inventario, ignorar
      if (inv.max > 500000) return;   // máximos > 500k: ignorar siempre

      var necesario = inv.max - inv.exist;
      if (necesario <= 0) return;     // inventario completo, no necesita fabricar

      var ordenes = codigosColatado[cod];

      // ── Mapa estado de pedidos para este código ──
      var PED_MUERTOS = ['TERMINADO','CANCELADO','CERRADO'];
      var mapEstPedLocal = {};
      for (var fp2 = 1; fp2 < dataPed.length; fp2++) {
        var fol2 = String(dataPed[fp2][pdFOL]||'');
        var est2 = String(dataPed[fp2][pdEST]||'').toUpperCase().trim();
        if (fol2) mapEstPedLocal[fol2] = est2;
      }

      // Backorder = Σsol - Σprod SOLO de órdenes pertenecientes a pedidos VIVOS
      var totalSolOrdenes  = 0;
      var totalProdOrdenes = 0;
      ordenes.forEach(function(o) {
        var estPed = mapEstPedLocal[o.pedido] || '';
        if (PED_MUERTOS.indexOf(estPed) > -1) return; // excluir pedidos muertos del backorder
        totalSolOrdenes  += Number(o.sol)  || 0;
        totalProdOrdenes += Number(o.prod) || 0;
      });
      var backorder = totalSolOrdenes - totalProdOrdenes;

      // Ajuste correcto = lo que falta fabricar - lo que ya está en backorder de órdenes
      var ajuste = necesario - backorder;

      // COLATADO: si |ajuste| < 1500 → está OK, no mostrar
      if (Math.abs(ajuste) < 1500) return;

      // pedidosAgrupados: SOLO pedidos VIVOS; pero todas sus órdenes se muestran (cualquier estado)
      var pedidosAgrupados = [];
      ordenes.forEach(function(o) {
        var estPed = mapEstPedLocal[o.pedido] || '';
        if (PED_MUERTOS.indexOf(estPed) > -1) return; // excluir pedidos muertos de la lista
        var key = o.pedido;
        var existing = pedidosAgrupados.find(function(x){ return x.pedido === key; });
        var ordenKey = o.serie + '.' + ('0000'+o.orden).slice(-4);
        if (existing) {
          var yaEsta = existing.detOrdenes.find(function(d){ return d.ordenKey === ordenKey; });
          if (!yaEsta) {
            existing.detOrdenes.push({ id: o.id, ordenKey: ordenKey, sol: o.sol, prod: o.prod, estado: o.estado, maquina: o.maquina, unidad: o.unidad });
            existing.sol  += Number(o.sol)  || 0;
            existing.prod += Number(o.prod) || 0;
          }
        } else {
          pedidosAgrupados.push({
            pedido:     key,
            fechaPed:   o.fechaPed || '',
            detOrdenes: [{ id: o.id, ordenKey: ordenKey, sol: o.sol, prod: o.prod, estado: o.estado, maquina: o.maquina, unidad: o.unidad }],
            sol:        Number(o.sol)  || 0,
            prod:       Number(o.prod) || 0
          });
        }
      });

      // Buscar orden candidato con avance ≤ 15% solo entre pedidos VIVOS
      var ordenCandidato = null;
      for (var i = 0; i < ordenes.length; i++) {
        var estPedOrd = mapEstPedLocal[ordenes[i].pedido] || '';
        if (PED_MUERTOS.indexOf(estPedOrd) > -1) continue;
        var avanceOrd = ordenes[i].sol > 0 ? (ordenes[i].prod / ordenes[i].sol) : 0;
        if (avanceOrd <= 0.15) { ordenCandidato = ordenes[i]; break; }
      }

      // Datos de la primera orden (de pedido vivo) para info de especificaciones
      var ord0 = ordenes.find(function(o){
        return PED_MUERTOS.indexOf(mapEstPedLocal[o.pedido]||'') === -1;
      }) || ordenes[0];

      if (ordenCandidato) {
        // CASO A: Aumentar orden existente con el ajuste (no necesario completo)
        resultados.push({
          tipo:              'ACTUALIZAR',
          codigo:            cod,
          desc:              ord0.desc,
          maquina:           ordenCandidato.maquina,
          exist:             inv.exist,
          min:               inv.min,
          max:               inv.max,
          necesario:         necesario,
          ajuste:            ajuste,
          totalSolOrdenes:   totalSolOrdenes,
          totalProdOrdenes:  totalProdOrdenes,
          backorder:         backorder,
          pedidosAgrupados:  pedidosAgrupados,
          idOrden:           ordenCandidato.id,
          pedido:            ordenCandidato.pedido,
          serie:             ordenCandidato.serie,
          ordenNum:          ordenCandidato.orden,
          solActual:         ordenCandidato.sol,
          solNuevo:          ordenCandidato.sol + ajuste,
          cantActual:        ordenCandidato.cant,
          cantNuevo:         ordenCandidato.cant + ajuste,
          unidad:            ordenCandidato.unidad,
          peso:              ordenCandidato.peso,
          avance:            ordenCandidato.sol > 0 ? Math.round((ordenCandidato.prod/ordenCandidato.sol)*100) : 0,
          tipo_prod:         ord0.tipo,
          dia:               ord0.dia,
          long:              ord0.long,
          cuerda:            ord0.cuerda,
          cuerpo:            ord0.cuerpo,
          acero:             ord0.acero,
          seleccionado:      true
        });
      } else {
        // CASO B: Crear nuevo pedido TEM — folio único por código (maxTem se busca globalmente incluyendo los ya creados en este lote)
        // maxTem se recalcula aquí para incluir los TEM ya agregados en resultados de este ciclo
        var maxTemLocal = 0;
        for (var pt = 1; pt < dataPed.length; pt++) {
          var folT = String(dataPed[pt][pdFOL]||'');
          if (folT.toUpperCase().indexOf('TEM-') === 0) {
            var nT = parseInt(folT.replace(/[^0-9]/g,''))||0;
            if (nT > maxTemLocal) maxTemLocal = nT;
          }
        }
        // También contar los TEM ya generados en este ciclo (resultados previos tipo NUEVA_ORDEN)
        resultados.forEach(function(r) {
          if (r.tipo === 'NUEVA_ORDEN' && r.folioTEM) {
            var nR = parseInt(r.folioTEM.replace(/[^0-9]/g,''))||0;
            if (nR > maxTemLocal) maxTemLocal = nR;
          }
        });
        var folioTEM = 'TEM-' + (maxTemLocal + 1); // sin padding para TEM-9, TEM-10, TEM-11...
        var ruta = rutaMap[cod] || [];
        var serie = ruta.length > 0 ? ruta[0].serie : 'P';

        resultados.push({
          tipo:              'NUEVA_ORDEN',
          codigo:            cod,
          desc:              ord0.desc,
          maquina:           ord0.maquina,
          exist:             inv.exist,
          min:               inv.min,
          max:               inv.max,
          necesario:         necesario,
          ajuste:            ajuste,
          totalSolOrdenes:   totalSolOrdenes,
          totalProdOrdenes:  totalProdOrdenes,
          backorder:         backorder,
          pedidosAgrupados:  pedidosAgrupados,
          folioTEM:          folioTEM,
          cantNuevo:         ajuste,
          unidad:            ord0.unidad,
          peso:              ord0.peso,
          serie:             serie,
          ruta:              ruta,
          tipo_prod:         ord0.tipo,
          dia:               ord0.dia,
          long:              ord0.long,
          cuerda:            ord0.cuerda,
          cuerpo:            ord0.cuerpo,
          acero:             ord0.acero,
          ordenRef:          ord0.serie + '.' + ('0000'+ord0.orden).slice(-4),
          seleccionado:      true
        });
      }
    });

    // ── Códigos sin pedidos/órdenes vivas COLATADO pero por debajo del 75% de Exist/Max ──
    // Iterar invMap buscando códigos que tengan ruta COLATADO y no aparezcan ya en resultados
    var codsProcesados = {};
    resultados.forEach(function(r){ codsProcesados[r.codigo] = true; });

    Object.keys(invMap).forEach(function(codInv) {
      if (codsProcesados[codInv]) return;
      if (!rutaMap[codInv]) return;
      var inv2 = invMap[codInv];
      if (!inv2 || inv2.max > 500000 || inv2.max <= 0) return;
      var pctInv = inv2.exist / inv2.max;
      if (pctInv >= 0.75) return;
      var ruta2 = rutaMap[codInv];
      var ord2  = ruta2[0] || {};
      // Calcular folio TEM incremental igual que en NUEVA_ORDEN
      var maxTemSP = 0;
      for (var pts = 1; pts < dataPed.length; pts++) {
        var folTS = String(dataPed[pts][pdFOL]||'');
        if (folTS.toUpperCase().indexOf('TEM-') === 0) {
          var nTS = parseInt(folTS.replace(/[^0-9]/g,''))||0;
          if (nTS > maxTemSP) maxTemSP = nTS;
        }
      }
      resultados.forEach(function(r) {
        if (r.tipo === 'NUEVA_ORDEN' || r.tipo === 'SIN_PEDIDO') {
          var nR = parseInt((r.folioTEM||'').replace(/[^0-9]/g,''))||0;
          if (nR > maxTemSP) maxTemSP = nR;
        }
      });
      var folioTEMsp = 'TEM-' + (maxTemSP + 1);
      var necesarioSP = inv2.max - inv2.exist;
      resultados.push({
        tipo:              'SIN_PEDIDO',
        codigo:            codInv,
        desc:              '',
        maquina:           '',
        exist:             inv2.exist,
        min:               inv2.min,
        max:               inv2.max,
        necesario:         necesarioSP,
        ajuste:            necesarioSP,
        backorder:         0,
        totalSolOrdenes:   0,
        totalProdOrdenes:  0,
        pedidosAgrupados:  [],
        folioTEM:          folioTEMsp,
        cantNuevo:         necesarioSP,
        unidad:            'KG',
        peso:              ord2.peso || 0,
        serie:             ord2.serie || 'P',
        ruta:              ruta2,
        tipo_prod:         ord2.tipo || '',
        dia:               ord2.diam || '',
        long:              String(ord2.long || '').trim(),
        cuerda:            ord2.cuerda || '',
        cuerpo:            ord2.cuerpo || '',
        acero:             ord2.acero || '',
        ordenRef:          '',
        seleccionado:      false
      });
    });

    return JSON.stringify({ success: true, items: resultados });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── SIN PEDIDO + BAJO STOCK genérico (campanita PRIORIDAD) ───────────────────
// Reutiliza la misma lógica que calcularSincronizarColatado pero para cualquier proceso.
// Devuelve solo los ítems tipo SIN_PEDIDO (exist/max <= 0.75, sin pedido activo, con ruta en el proceso).
function obtenerSinPedidoBajoStock(proceso) {
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shInv   = ss.getSheetByName("INVENTARIO_EXTERNO");
    var shOrd   = ss.getSheetByName("ORDENES");
    var shPed   = ss.getSheetByName("PEDIDOS");
    var shRutas = ss.getSheetByName("RUTAS");
    var shCod   = ss.getSheetByName("CODIGOS");
    var shEst   = ss.getSheetByName("ESTANDARES");

    if (!shInv || !shOrd || !shPed || !shRutas) return JSON.stringify({ success: false, msg: "Hojas no encontradas", items: [] });

    var PROC = String(proceso).toUpperCase().trim();
    var PED_MUERTOS = ['TERMINADO','CANCELADO','CERRADO','SOBREPRODUCCION','ENTREGADO'];
    var ORD_MUERTOS = ['CERRADO'];

    // ── 1. Máquinas del proceso (ESTANDARES) ──
    var dataEst = shEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ePROC   = hEst.indexOf("PROCESO");
    var eMAQ    = hEst.indexOf("MAQUINA");
    var maqsDelProceso = [];
    for (var e = 1; e < dataEst.length; e++) {
      if (String(dataEst[e][ePROC]||"").toUpperCase().trim() === PROC) {
        var mn = String(dataEst[e][eMAQ]||"").trim().toUpperCase();
        if (mn) maqsDelProceso.push(mn);
      }
    }

    // ── 2. Mapa código→venta (CODIGOS col A fab, col F venta) ──
    var dataCod = shCod ? shCod.getDataRange().getValues() : [];
    var codigoVentaMap = {};
    for (var cv = 1; cv < dataCod.length; cv++) {
      var cvF = String(dataCod[cv][0]||'').trim().toUpperCase();
      var cvV = String(dataCod[cv][5]||'').trim().toUpperCase();
      if (cvF && cvV) codigoVentaMap[cvF] = cvV;
    }

    // ── 3. Mapa inventario: código → {exist, min, max, back, existNeg} ──
    // Col A=CODIGO, B=EXIST, C=MIN, D=MAX, E=BACKORDER(Galva)
    var dataInv = shInv.getDataRange().getValues();
    var invMap = {};
    for (var i = 1; i < dataInv.length; i++) {
      var ic = String(dataInv[i][0]||'').trim().toUpperCase();
      if (ic) invMap[ic] = {
        exist: Number(dataInv[i][1])||0,
        min:   Number(dataInv[i][2])||0,
        max:   Number(dataInv[i][3])||0,
        back:  Number(dataInv[i][4])||0   // col E = BACKORDER = Galva
      };
    }
    // getInv: para código negro (no empieza en 7) busca si tiene codigoVenta asociado con max>0
    // Devuelve {exist, min, max, back, existNeg, codigoVenta, esVarilla}
    var esProcesoVarilla = (PROC === 'ROSCADO');

    function getInv(cod) {
      var cUp = String(cod||'').trim().toUpperCase();
      var cVenta = codigoVentaMap[cUp] || '';
      // Lógica especial de varilla SOLO para proceso ROSCADO
      if (esProcesoVarilla && cUp.charAt(0) === '7' && cVenta && invMap[cVenta] && invMap[cVenta].max > 0) {
        var invNegro = invMap[cUp]    || { exist: 0, back: 0 };
        var invVenta = invMap[cVenta];
        return {
          exist:       invVenta.exist,
          min:         invVenta.min,
          max:         invVenta.max,
          back:        invVenta.back,
          existNeg:    invNegro.exist,
          codigoVenta: cVenta,
          esVarilla:   true
        };
      }
      // Para todos los procesos: código 7 busca su venta si existe
      if (cUp.charAt(0) === '7') {
        if (cVenta && invMap[cVenta]) return Object.assign({}, invMap[cVenta], { existNeg: null, codigoVenta: cVenta, esVarilla: false });
        return invMap[cUp] ? Object.assign({}, invMap[cUp], { existNeg: null, codigoVenta: '', esVarilla: false }) : null;
      }
      // Código normal (no empieza en 7)
      return invMap[cUp] ? Object.assign({}, invMap[cUp], { existNeg: null, codigoVenta: '', esVarilla: false }) : null;
    }

    // ── 4. Mapa ruta: código → TODOS los pasos (igual que calcularSincronizarColatado) ──
    var dataRut = shRutas.getDataRange().getValues();
    var rutaMap    = {};  // codigo → [pasos completos]
    var rutaInfoMap = {}; // codigo → info visual (primer paso del PROC)
    for (var r = 1; r < dataRut.length; r++) {
      var rCod  = String(dataRut[r][1]||'').trim();
      var rProc = String(dataRut[r][4]||'').trim().toUpperCase();
      if (!rCod) continue;
      // Guardar todos los pasos de la ruta completa
      if (!rutaMap[rCod]) rutaMap[rCod] = [];
      rutaMap[rCod].push({
        sec:              Number(dataRut[r][3])||0,
        proceso:          rProc,
        maquinas:         String(dataRut[r][5]||'').split(',').map(function(m){ return m.trim(); }),
        maquinaSeleccionada: String(dataRut[r][5]||'').split(',')[0].trim(),
        cantLote:         parseFloat(dataRut[r][22])||0,
        pt:               dataRut[r][12],
        venta:            dataRut[r][13],
        peso:             Number(dataRut[r][15])||0,
        tipo:             String(dataRut[r][6]||''),
        diam:             String(dataRut[r][7]||''),
        long:             String(dataRut[r][8]||''),
        cuerda:           String(dataRut[r][9]||''),
        cuerpo:           String(dataRut[r][10]||''),
        acero:            String(dataRut[r][11]||''),
        serie:            String(dataRut[r][14]||'P'),
        unidad:           String(dataRut[r][16]||'KG'),
        mp:               String(dataRut[r][29]||'')
      });
      // Info visual: solo del proceso activo (primer paso encontrado)
      if (rProc === PROC && !rutaInfoMap[rCod]) {
        rutaInfoMap[rCod] = {
          tipo: String(dataRut[r][6]||''), dia: String(dataRut[r][7]||''),
          long: String(dataRut[r][8]||''), cuerda: String(dataRut[r][9]||''),
          cuerpo: String(dataRut[r][10]||''), acero: String(dataRut[r][11]||''),
          maquinas: String(dataRut[r][5]||'').split(',').map(function(m){ return m.trim(); }),
          serie: String(dataRut[r][14]||'P'),
          unidad: String(dataRut[r][16]||'KG'),
          cantLote: parseFloat(dataRut[r][22])||0,
          peso: Number(dataRut[r][15])||0
        };
      }
    }

    // ── 5. Estado de pedidos: código → tiene pedido activo? ──
    var dataPed = shPed.getDataRange().getValues();
    var hPed    = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdFOL   = hPed.indexOf("FOLIO");   if (pdFOL  < 0) pdFOL  = 1;
    var pdCOD   = hPed.indexOf("CODIGO");  if (pdCOD  < 0) pdCOD  = 3;
    var pdEST   = hPed.indexOf("ESTADO");  if (pdEST  < 0) pdEST  = 8;
    var pdDESC  = hPed.indexOf("DESCRIPCION"); if (pdDESC < 0) pdDESC = 4;

    var codigosConPedidoVivo = {};
    var descPorCodigo = {};
    for (var p = 1; p < dataPed.length; p++) {
      var pCod = String(dataPed[p][pdCOD]||'').trim().toUpperCase();
      var pEst = String(dataPed[p][pdEST]||'').trim().toUpperCase();
      if (!pCod) continue;
      if (!descPorCodigo[pCod]) descPorCodigo[pCod] = String(dataPed[p][pdDESC]||'').trim();
      if (PED_MUERTOS.indexOf(pEst) === -1) codigosConPedidoVivo[pCod] = true;
    }

    // ── 6. Leer TODAS las órdenes para análisis de proceso crítico ──
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oCOD    = hOrd.indexOf("CODIGO");
    var oEST    = hOrd.indexOf("ESTADO");
    var oPROC   = hOrd.indexOf("PROCESO");
    var oMAQ    = hOrd.indexOf("MAQUINA");
    var oDESC   = hOrd.indexOf("DESCRIPCION");
    var oPED    = hOrd.indexOf("PEDIDO");
    var oSEC    = hOrd.indexOf("SEC");

    var ESTADOS_MUERTOS_ORD = ['TERMINADO','CANCELADO','SOBREPRODUCCION','CERRADO'];

    var codigosConOrdenViva = {};

    // Mapa: codigo → { pedidos: Set, ordenesPorPedidoYProc: {pedido: {proc: [estados]}} }
    var mapaOrdenesPorCodigo = {};
    for (var o = 1; o < dataOrd.length; o++) {
      var cod  = String(dataOrd[o][oCOD] ||'').trim().toUpperCase();
      var est  = String(dataOrd[o][oEST] ||'').toUpperCase().trim();
      var proc = String(dataOrd[o][oPROC]||'').toUpperCase().trim();
      var ped  = String(dataOrd[o][oPED] ||'').trim();
      var sec  = Number(dataOrd[o][oSEC] ||0);
      if (!cod) continue;
      if (!mapaOrdenesPorCodigo[cod]) mapaOrdenesPorCodigo[cod] = {};
      if (!mapaOrdenesPorCodigo[cod][ped]) mapaOrdenesPorCodigo[cod][ped] = {};
      if (!mapaOrdenesPorCodigo[cod][ped][proc]) mapaOrdenesPorCodigo[cod][ped][proc] = { estados: [], secMin: 999 };
      mapaOrdenesPorCodigo[cod][ped][proc].estados.push(est);
      if (sec < mapaOrdenesPorCodigo[cod][ped][proc].secMin) mapaOrdenesPorCodigo[cod][ped][proc].secMin = sec;
      // Para el proceso activo: registrar órdenes vivas
      if (proc === PROC) {
        var maqsOrden = String(dataOrd[o][oMAQ]||'').toUpperCase().split(',').map(function(m){return m.trim();});
        var tieneMAQ  = maqsOrden.some(function(m){ return maqsDelProceso.indexOf(m) > -1; });
        if (tieneMAQ && ESTADOS_MUERTOS_ORD.indexOf(est) === -1) {
          codigosConOrdenViva[cod] = true;
          if (!descPorCodigo[cod]) descPorCodigo[cod] = String(dataOrd[o][oDESC]||'').trim();
        }
      }
      if (!descPorCodigo[cod] && dataOrd[o][oDESC]) descPorCodigo[cod] = String(dataOrd[o][oDESC]||'').trim();
    }

    // ── Determinar proceso crítico por código ──
    // Varilla (ROSCADO): proceso crítico = ROSCADO
    // Resto: proceso crítico = el de menor SEC en RUTAS
    function getProcesoCritico(cod) {
      if (esProcesoVarilla) return 'ROSCADO';
      var pasos = rutaMap[cod] || [];
      if (pasos.length === 0) return PROC;
      var minSec = pasos[0].sec, minProc = pasos[0].proceso;
      pasos.forEach(function(p){ if(p.sec < minSec){ minSec = p.sec; minProc = p.proceso; } });
      return minProc.toUpperCase().trim();
    }

    // ── Detectar códigos con pedido vivo pero proceso crítico completamente muerto ──
    var codigosAlerta = {}; // codigo → true si pedido vivo pero proceso crítico terminado
    Object.keys(codigosConPedidoVivo).forEach(function(cod) {
      var pedidos = mapaOrdenesPorCodigo[cod];
      if (!pedidos) return;
      var procCrit = getProcesoCritico(cod);
      // Verificar si TODOS los pedidos vivos tienen el proceso crítico terminado/cancelado
      var todosCriticosTerminados = true;
      Object.keys(pedidos).forEach(function(ped) {
        // Solo considerar pedidos vivos (ya sabemos cod tiene pedido vivo)
        var procData = pedidos[ped][procCrit];
        if (!procData) { todosCriticosTerminados = false; return; }
        var algunaViva = procData.estados.some(function(e){
          return ESTADOS_MUERTOS_ORD.indexOf(e) === -1;
        });
        if (algunaViva) todosCriticosTerminados = false;
      });
      if (todosCriticosTerminados) codigosAlerta[cod] = true;
    });

    // ── 7. Construir resultado: bajo stock sin pedido vivo, + alertas ──
    var resultados = [];
    var resultadosAlerta = [];
    Object.keys(rutaInfoMap).forEach(function(cod) {
      var codU = cod.trim().toUpperCase();
      // Caso alerta: pedido vivo pero proceso crítico terminado
      if (codigosConPedidoVivo[codU] && codigosAlerta[codU]) {
        var inv2 = getInv(codU); if (!inv2 || inv2.max <= 0 || inv2.max > 500000) return;
        var pct2, pedir2;
        if (inv2.esVarilla) {
          var g2=inv2.back||0, en2=inv2.existNeg||0, tot2=inv2.exist+g2+en2;
          pct2=inv2.max>0?tot2/inv2.max:0; pedir2=Math.max(0,Math.round(inv2.max-inv2.exist-g2-en2));
        } else { pct2=inv2.exist/inv2.max; pedir2=Math.max(0,Math.round(inv2.max-inv2.exist)); }
        if (pct2 >= 0.75) return;
        var rInfo2=rutaInfoMap[cod]||{}, ruta2=rutaMap[cod]||[];
        resultadosAlerta.push({
          codigo:codU, desc:descPorCodigo[codU]||codU,
          exist:inv2.exist, min:inv2.min, max:inv2.max,
          back:inv2.back||0, existNeg:inv2.esVarilla?(inv2.existNeg||0):null,
          esVarilla:inv2.esVarilla||false, codigoVenta:inv2.codigoVenta||'',
          pct:Math.round(pct2*100), pedirCalc:pedir2,
          tipo_prod:rInfo2.tipo||'', dia:rInfo2.dia||'', long:rInfo2.long||'',
          cuerda:rInfo2.cuerda||'', cuerpo:rInfo2.cuerpo||'', acero:rInfo2.acero||'',
          maquinas:rInfo2.maquinas||[], serie:rInfo2.serie||'P',
          unidad:rInfo2.unidad||'KG', cantLote:rInfo2.cantLote||0, peso:rInfo2.peso||0,
          ruta:ruta2, alerta:true
        });
        return;
      }
      if (codigosConPedidoVivo[codU]) return;
      var inv = getInv(codU);
      if (!inv || inv.max <= 0 || inv.max > 500000) return;

      // Calcular pct y pedir según tipo
      var pct, pedir;
      if (inv.esVarilla) {
        // Varilla: (exist_venta + galva + exis_neg) / max_venta
        var galva    = inv.back    || 0;
        var existNeg = inv.existNeg|| 0;
        var total    = inv.exist + galva + existNeg;
        pct   = inv.max > 0 ? total / inv.max : 0;
        pedir = Math.max(0, Math.round(inv.max - inv.exist - galva - existNeg));
      } else {
        pct   = inv.exist / inv.max;
        pedir = Math.max(0, Math.round(inv.max - inv.exist));
      }
      if (pct >= 0.75) return;

      var rInfo = rutaInfoMap[cod] || {};
      var rutaCompleta = rutaMap[cod] || [];
      resultados.push({
        codigo:      codU,
        desc:        descPorCodigo[codU] || codU,
        exist:       inv.exist,
        min:         inv.min,
        max:         inv.max,
        back:        inv.back    || 0,
        existNeg:    inv.esVarilla ? (inv.existNeg || 0) : null,
        esVarilla:   inv.esVarilla || false,
        codigoVenta: inv.codigoVenta || '',
        pct:         Math.round(pct * 100),
        pedirCalc:   pedir,
        tipo_prod:   rInfo.tipo    || '',
        dia:         rInfo.dia     || '',
        long:        rInfo.long    || '',
        cuerda:      rInfo.cuerda  || '',
        cuerpo:      rInfo.cuerpo  || '',
        acero:       rInfo.acero   || '',
        maquinas:    rInfo.maquinas || [],
        serie:       rInfo.serie   || 'P',
        unidad:      rInfo.unidad  || 'KG',
        cantLote:    rInfo.cantLote|| 0,
        peso:        rInfo.peso    || 0,
        ruta:        rutaCompleta
      });
    });


    resultados.sort(function(a, b){ return a.pct - b.pct; });
    resultadosAlerta.sort(function(a, b){ return a.pct - b.pct; });
    return JSON.stringify({ success: true, items: resultadosAlerta.concat(resultados) });
  } catch(e) {
    Logger.log("obtenerSinPedidoBajoStock ERROR: " + e.message);
    return JSON.stringify({ success: false, msg: e.message, items: [] });
  }
}

// ── Aplicar Sincronizar Colatado ──────────────────────────────────────────────
function aplicarSincronizarColatado(items) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return JSON.stringify({ success: false, msg: 'Servidor ocupado.' });
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd   = ss.getSheetByName("ORDENES");
    var shPed   = ss.getSheetByName("PEDIDOS");
    var shLot   = ss.getSheetByName("LOTES");

    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oID     = hOrd.indexOf("ID");
    var oSOL    = hOrd.indexOf("SOLICITADO");
    var oCANT   = hOrd.indexOf("CANTIDAD");
    var oSERIE  = hOrd.indexOf("SERIE");
    var oORDN   = hOrd.indexOf("ORDEN");

    var dataPed = shPed.getDataRange().getValues();
    var hPed    = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdFOL   = hPed.indexOf("FOLIO"); if (pdFOL === -1) pdFOL = 1;
    var pdCANT  = hPed.indexOf("CANTIDAD");
    var pdSOL   = hPed.indexOf("SOLICITADO");

    var log = [];
    var hoySoloFecha = new Date(); hoySoloFecha.setHours(0,0,0,0);

    items.forEach(function(item) {
      if (item.tipo === 'ACTUALIZAR') {
        // Usar idOrden del item (puede ser el seleccionado en el dropdown)
        var idOrdenTarget = item.idOrden || '';
        var serieTarget = item.serie || '';
        var ordenNumTarget = Number(item.ordenNum)||0;
        // A. Actualizar SOLICITADO y CANTIDAD: buscar por idOrden primero, si no por serie+ordenNum
        for (var o = 1; o < dataOrd.length; o++) {
          var matchId    = idOrdenTarget && String(dataOrd[o][oID]) === String(idOrdenTarget);
          var matchSerie = !idOrdenTarget && String(dataOrd[o][oSERIE]) === serieTarget && Number(dataOrd[o][oORDN]) === ordenNumTarget;
          if (matchId || matchSerie) {
            shOrd.getRange(o+1, oSOL+1).setValue(item.solNuevo);
            shOrd.getRange(o+1, oCANT+1).setValue(item.cantNuevo);
          }
        }
        // B. Actualizar CANTIDAD en PEDIDOS sumando el ajuste (no el necesario completo)
        var ajusteVal = Number(item.ajuste || item.necesario || 0);
        for (var p = 1; p < dataPed.length; p++) {
          if (String(dataPed[p][pdFOL]) === String(item.pedido)) {
            var cantPedActual = Number(dataPed[p][pdCANT])||0;
            shPed.getRange(p+1, pdCANT+1).setValue(cantPedActual + ajusteVal);
            break;
          }
        }
        log.push('ACTUALIZADO: ' + item.codigo + ' ajuste:' + ajusteVal);

      } else if (item.tipo === 'NUEVA_ORDEN') {
        // A. Crear pedido TEM en PEDIDOS
        shPed.appendRow([
          Utilities.getUuid().substring(0,8),
          item.folioTEM,
          hoySoloFecha,
          "'" + String(item.codigo).padStart(9,'0'),
          item.desc,
          1,
          item.cantNuevo,
          item.unidad || 'KG',
          'PLANEADO'
        ]);

        // B. Crear órdenes en ORDENES (una por paso de la ruta)
        var ruta = item.ruta || [];
        if (ruta.length === 0) {
          log.push('SIN RUTA: ' + item.codigo + ' — no se creó orden');
          return;
        }
        var nOrden = getSiguienteNumeroOrden(item.serie);
        ruta.forEach(function(paso, idx) {
          var idProc = Utilities.getUuid().substring(0,8);
          var sol = item.cantNuevo;
          var uni = String(item.unidad||'KG').toUpperCase();
          if (uni === 'PZA' || uni === 'CTO') {
            sol = item.cantNuevo * (paso.peso || item.peso || 0);
          }
          shOrd.appendRow([
            idProc,
            item.folioTEM,
            1,
            hoySoloFecha,
            item.serie,
            nOrden,
            "'" + String(item.codigo).padStart(9,'0'),
            item.desc,
            item.cantNuevo,
            item.unidad || 'KG',
            paso.sec,
            paso.proceso,
            paso.maquinas[0] || '',
            sol,
            0,
            'ABIERTO',
            paso.pt,
            paso.venta,
            paso.peso || item.peso,
            paso.tipo || item.tipo_prod,
            "'" + String(paso.diam || item.dia),
            "'" + String(paso.long || ''),
            paso.cuerda || item.cuerda,
            paso.cuerpo || item.cuerpo,
            paso.acero  || item.acero,
            new Date(),
            '', '', '', ''
          ]);
        });

        // C. Crear LOTES para el primer proceso
        var paso0 = ruta[0];
        var cantLoteBase = paso0.cantLote || item.cantNuevo;
        var nLotes = Math.max(1, Math.ceil(item.cantNuevo / cantLoteBase));
        var dataLotesActual = shLot.getDataRange().getValues();
        // Buscar último consecutivo de esta serie+orden
        var maxConsec = 0;
        for (var l = 1; l < dataLotesActual.length; l++) {
          var nomL = String(dataLotesActual[l][4]||'');
          var partes = nomL.split('.');
          if (partes[0] === item.serie && Number(partes[1]) === nOrden) {
            var cs = Number(partes[2])||0;
            if (cs > maxConsec) maxConsec = cs;
          }
        }
        // Buscar idOrdenPrimerProc (ya recién insertada — re-leer ORDENES)
        var dataOrdFresh = shOrd.getDataRange().getValues();
        var idRef = '';
        var minSecTem = Infinity;
        for (var ox = 1; ox < dataOrdFresh.length; ox++) {
          if (String(dataOrdFresh[ox][4]) === item.serie && Number(dataOrdFresh[ox][5]) === nOrden) {
            var secTem = Number(dataOrdFresh[ox][10]);
            if (secTem < minSecTem) { minSecTem = secTem; idRef = String(dataOrdFresh[ox][0]); }
          }
        }
        var pesoLote = paso0.cantLote || item.cantNuevo;
        var nuevosLotes = [];
        for (var lt = 1; lt <= nLotes; lt++) {
          var consec = maxConsec + lt;
          var numOrdStr  = ('0000'+nOrden).slice(-4);
          var consecStr  = consec < 100 ? ('00'+consec).slice(-2) : String(consec);
          var nombreLote = item.serie + '.' + numOrdStr + '.' + consecStr;
          var cantEste   = (lt < nLotes) ? pesoLote : (item.cantNuevo - pesoLote*(nLotes-1));
          nuevosLotes.push([
            Utilities.getUuid(), item.serie, idRef, consec,
            nombreLote, cantEste, 0, 'ABIERTO', hoySoloFecha,
            '', '', '', '', '', 'NADA'
          ]);
        }
        if (nuevosLotes.length > 0) {
          shLot.getRange(shLot.getLastRow()+1, 1, nuevosLotes.length, nuevosLotes[0].length).setValues(nuevosLotes);
        }
        log.push('NUEVA ORDEN TEM: ' + item.folioTEM + ' — ' + item.codigo);
      }
    });

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, msg: log.join(' | '), count: log.length });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

// ── 15. aplicarActualizacionEjecucion — aplica cambios en PEDIDO + ORDEN elegida ──
function aplicarActualizacionEjecucion(seleccionados) {
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetPed = ss.getSheetByName("PEDIDOS");
    var sheetOrd = ss.getSheetByName("ORDENES");

    if (!sheetPed) throw new Error("No existe la hoja PEDIDOS");
    if (!sheetOrd) throw new Error("No existe la hoja ORDENES");

    var dataPed = sheetPed.getDataRange().getValues();
    var hPed    = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var pdID    = hPed.indexOf("ID");
    var pdCANT  = hPed.indexOf("CANTIDAD");
    var pdSOL   = hPed.indexOf("SOLICITADO");

    var dataOrd = sheetOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oID     = hOrd.indexOf("ID");
    var oCANT   = hOrd.indexOf("CANTIDAD");
    var oSOL    = hOrd.indexOf("SOLICITADO");
    var oEST    = hOrd.indexOf("ESTADO");
    var oSERIE  = hOrd.indexOf("SERIE");
    var oORDEN  = hOrd.indexOf("ORDEN");
    var oUNI    = hOrd.indexOf("UNIDAD");
    var oDESC   = hOrd.indexOf("DESCRIPCION");
    var oPESO   = hOrd.indexOf("PESO");
    var oLONG   = hOrd.indexOf("LONGITUD");
    var oPROD   = hOrd.indexOf("PRODUCIDO");

    var MUERTOS = ["CANCELADO","TERMINADO","SOBREPRODUCCION","CERRADO","ENTREGADO"];

    function calcSolicitado(cantidad, unidad, desc, peso, longitud) {
      var u = String(unidad).toUpperCase().trim();
      if (u === "KG" || u === "ROL") return cantidad;
      var p = Number(peso) || 0;
      var sol = cantidad * p;
      if (String(desc).toUpperCase().indexOf("VARILLA") > -1) {
        var l = parseFloat(longitud) || 0;
        sol = cantidad * p * l;
      }
      return sol;
    }

    function calcNuevoEstado(producido, nuevoSol, estadoActual) {
      if (MUERTOS.indexOf(estadoActual) > -1) return estadoActual;
      if (producido <= 0) return "ABIERTA";
      if (producido >= nuevoSol) return "TERMINADO";
      return "EN PROCESO";
    }

    var countPed = 0;
    var countOrd = 0;

    (seleccionados || []).forEach(function(item) {

      // 1. Actualizar CANTIDAD y SOLICITADO en PEDIDOS
      for (var p = 1; p < dataPed.length; p++) {
        if (String(dataPed[p][pdID]) === String(item.pedidoId)) {
          if (pdCANT > -1) sheetPed.getRange(p+1, pdCANT+1).setValue(item.cantPedDespues);
          if (pdSOL  > -1) sheetPed.getRange(p+1, pdSOL+1 ).setValue(item.solPedDespues || item.cantPedDespues);
          countPed++;
          break;
        }
      }

      // 2. Actualizar TODOS LOS PROCESOS de la orden elegida (misma SERIE+ORDEN)
      // La clave viene en item.ordenElegidaKey = "SERIE.ORDEN"
      // Separamos en las dos partes
      var partes      = String(item.ordenElegidaKey || "").split(".");
      var serieEleg   = partes[0] || "";
      var ordenEleg   = partes[1] || "";
      var nuevaCant   = Number(item.nuevaCantOrden || 0);

      if (serieEleg && ordenEleg) {
        for (var o = 1; o < dataOrd.length; o++) {
          var rowO     = dataOrd[o];
          var serieO   = String(rowO[oSERIE] || "");
          var ordenO   = String(rowO[oORDEN] || "");
          if (serieO !== serieEleg || ordenO !== ordenEleg) continue;

          var estActual = String(rowO[oEST] || "").toUpperCase().trim();
          if (MUERTOS.indexOf(estActual) > -1) continue;

          // Calcular nuevo SOLICITADO con regla de oro
          var uOrd     = String(rowO[oUNI]  || "");
          var dOrd     = String(rowO[oDESC] || "");
          var pOrd     = Number(rowO[oPESO] || 0);
          var lOrd     = rowO[oLONG];
          var prod     = Math.max(Number(rowO[oPROD] || 0), 0);
          var nuevoSol = calcSolicitado(nuevaCant, uOrd, dOrd, pOrd, lOrd);
          var nuevoEst = calcNuevoEstado(prod, nuevoSol, estActual);

          // Escribir CANTIDAD, SOLICITADO y ESTADO
          if (oCANT > -1) sheetOrd.getRange(o+1, oCANT+1).setValue(nuevaCant);
          if (oSOL  > -1) sheetOrd.getRange(o+1, oSOL+1 ).setValue(Math.max(nuevoSol, 0));
          if (oEST  > -1) sheetOrd.getRange(o+1, oEST+1 ).setValue(nuevoEst);
          countOrd++;
        }
      }
    });

    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, msg: countPed + " pedido(s) y " + countOrd + " proceso(s) de orden actualizados." });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.message });
  }
}

function ajustarOrdenSobreproduccion(id, nuevoEstado, nuevaCantidad, procesoActual) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success:false, msg:"Servidor ocupado, intenta de nuevo." };
  try {
    var ss  = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shO = ss.getSheetByName("ORDENES");
    var shP = ss.getSheetByName("PEDIDOS");

    var dataO   = shO.getDataRange().getValues();
    var headers = dataO[0].map(function(h){ return String(h).toUpperCase().trim(); });

    var iID   = headers.indexOf("ID");
    var iEst  = headers.indexOf("ESTADO");
    var iSerie= headers.indexOf("SERIE");
    var iOrden= headers.indexOf("ORDEN");
    var iPed  = headers.indexOf("PEDIDO");
    var iUni  = headers.indexOf("UNIDAD");
    var iCant = headers.indexOf("CANTIDAD");
    var iSol  = headers.indexOf("SOLICITADO");
    var iDesc = headers.indexOf("DESCRIPCION");
    var iPeso = headers.indexOf("PESO");
    var iLon  = headers.indexOf("LONGITUD");

    // 1. Localizar la orden raíz para obtener SERIE + ORDEN + PEDIDO + CANTIDAD original
    var serieT="", ordenT="", pedidoT="", cantOriginal=0;
    for (var i=1; i<dataO.length; i++) {
      if (String(dataO[i][iID]) === String(id)) {
        serieT     = String(dataO[i][iSerie]);
        ordenT     = String(dataO[i][iOrden]);
        pedidoT    = String(dataO[i][iPed]);
        cantOriginal = parseFloat(dataO[i][iCant]) || 0;
        break;
      }
    }
    if (!serieT) return { success:false, msg:"Orden no encontrada: "+id };

    var diferencia = nuevaCantidad - cantOriginal; // puede ser 0

    // 2. Actualizar TODOS los procesos de esa orden (ESTADO + CANTIDAD + SOLICITADO)
    for (var j=1; j<dataO.length; j++) {
      if (String(dataO[j][iSerie])===serieT && String(dataO[j][iOrden])===ordenT) {
        // Estado
        shO.getRange(j+1, iEst+1).setValue(nuevoEstado);

        if (diferencia !== 0) {
          // Cantidad
          shO.getRange(j+1, iCant+1).setValue(nuevaCantidad);

          // SOLICITADO con regla de oro (igual que actualizarCantidadOrden)
          var uni  = String(dataO[j][iUni]  || "").toUpperCase();
          var desc = String(dataO[j][iDesc] || "").toUpperCase();
          var peso = parseFloat(dataO[j][iPeso]) || 0;
          var lon  = parseFloat(dataO[j][iLon])  || 0;
          var nuevoSol = nuevaCantidad;
          if (uni === "PZA" || uni === "CTO") {
            nuevoSol = nuevaCantidad * peso;
            if (desc.indexOf("VARILLA") > -1) nuevoSol *= lon;
          }
          shO.getRange(j+1, iSol+1).setValue(nuevoSol);
        }
      }
    }

    // 3. Ajustar CANTIDAD del PEDIDO si hubo cambio
    if (diferencia !== 0 && pedidoT) {
      var dataPed    = shP.getDataRange().getValues();
      var hPed       = dataPed[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var colIdPed   = hPed.indexOf("ID");
      var colCantPed = hPed.indexOf("CANTIDAD");
      if (colIdPed>-1 && colCantPed>-1) {
        for (var k=1; k<dataPed.length; k++) {
          if (String(dataPed[k][colIdPed]) === pedidoT) {
            var cantPedActual = parseFloat(dataPed[k][colCantPed]) || 0;
            shP.getRange(k+1, colCantPed+1).setValue(cantPedActual + diferencia);
            break;
          }
        }
      }
    }

    SpreadsheetApp.flush();
    lock.releaseLock();
    return { success:true, msg:"OK" };

  } catch(e) {
    try { lock.releaseLock(); } catch(x){}
    return { success:false, msg:e.toString() };
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA WipHTML (MENU WIP)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

function obtenerDatosWIP(familiasSeleccionadas) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetCod = ss.getSheetByName("CODIGOS");
  var sheetLotes = ss.getSheetByName("LOTES");

  var dataOrd = sheetOrd.getDataRange().getValues();
  var infoProcesos = {}; 
  var rutasPorOrdenBase = {}; 
  
  for(var o=1; o<dataOrd.length; o++) {
      var idUnico = String(dataOrd[o][0]).trim();
      var ordenBase = String(dataOrd[o][5]).trim();
      var info = {
          id: idUnico, ordenBase: ordenBase, codigo: String(dataOrd[o][6]).trim(),
          sec: Number(dataOrd[o][10]), proceso: String(dataOrd[o][11]).toUpperCase().trim(),
          descFormateada: dataOrd[o][19] + " DE " + dataOrd[o][20] + " X " + dataOrd[o][21],
          isLast: false
      };
      infoProcesos[idUnico] = info;
      if(!rutasPorOrdenBase[ordenBase]) rutasPorOrdenBase[ordenBase] = [];
      rutasPorOrdenBase[ordenBase].push(info);
  }
  for(var rb in rutasPorOrdenBase) {
    rutasPorOrdenBase[rb].sort((a,b) => a.sec - b.sec);
    if(rutasPorOrdenBase[rb].length > 0) rutasPorOrdenBase[rb][rutasPorOrdenBase[rb].length - 1].isLast = true;
  }

  var dataCod = sheetCod.getDataRange().getValues();
  var mapFamily = {};
  for(var c=1; c<dataCod.length; c++) {
      mapFamily[String(dataCod[c][0]).trim()] = String(dataCod[c][8] || "OTROS").toUpperCase().trim();
  }

  var dataProd = sheetProd.getDataRange().getValues();
  var prodAcum = {}; 
  for(var p=1; p<dataProd.length; p++) {
      var key = String(dataProd[p][3]).trim() + "|" + String(dataProd[p][2]).trim();
      prodAcum[key] = (prodAcum[key] || 0) + (Number(dataProd[p][10]) || 0);
  }

  var dataLotes = sheetLotes.getDataRange().getValues();
  var resultado = {};
  var mapaCodigosFiltro = {}; 

  for(var l=1; l<dataLotes.length; l++) {
      var loteName = String(dataLotes[l][4]).trim();
      var estadoLote = String(dataLotes[l][7]).toUpperCase().trim();
      var idReferencia = String(dataLotes[l][2]).trim(); 
      var calidadLote = String(dataLotes[l][22] || "NADA").toUpperCase().trim();

      if(!loteName || estadoLote === "CONCLUIDO" || !infoProcesos[idReferencia]) continue;

      var ordenBase = infoProcesos[idReferencia].ordenBase;
      var ruta = rutasPorOrdenBase[ordenBase];
      var codProducto = infoProcesos[idReferencia].codigo;
      var familia = mapFamily[codProducto] || "OTROS";

      if(familiasSeleccionadas.length > 0 && !familiasSeleccionadas.includes(familia)) continue;
      mapaCodigosFiltro[codProducto] = infoProcesos[idReferencia].descFormateada;

      var lineaProduccion = ruta.map(paso => ({
         ...paso,
         kg: prodAcum[loteName + "|" + paso.id] || 0
      }));

      for(var i=0; i < lineaProduccion.length; i++) {
          var actual = lineaProduccion[i];
          
          // FILTRO MANDATORIO: SI NO TIENE KILOS, NO SE MUESTRA (INCLUSO EN LAVADO/FINAL)
          if (actual.kg <= 0.5) continue;

          // REGLA CASCADA GLOBAL 60%
          var saltadoPorPosterior = false;
          for (var j = i + 1; j < lineaProduccion.length; j++) {
             if (lineaProduccion[j].kg >= (actual.kg * 0.60)) {
                saltadoPorPosterior = true;
                break;
             }
          }
          if (saltadoPorPosterior) continue;

          var siguiente = (i + 1 < lineaProduccion.length) ? lineaProduccion[i+1] : null;
          var anterior = (i > 0) ? lineaProduccion[i-1] : null;

          if(!resultado[actual.proceso]) resultado[actual.proceso] = { sec: actual.sec, lotes: [] };

          // A. REMANENTE
          if (siguiente && siguiente.kg > 0.5) {
             var remanente = actual.kg - siguiente.kg;
             if (remanente > 0.5) {
                resultado[actual.proceso].lotes.push({
                   lote: loteName, producto: codProducto, descFull: actual.descFormateada,
                   kilos: remanente, kgTotalRef: actual.kg,
                   tipo: "REMANENTE", calidad: calidadLote, isLast: false
                });
             }
          } 
          // B. PROCESANDO / NORMAL
          else {
             var kgReferencia = anterior ? anterior.kg : actual.kg;
             resultado[actual.proceso].lotes.push({
                lote: loteName, producto: codProducto, descFull: actual.descFormateada,
                kilos: actual.kg, kgTotalRef: kgReferencia,
                tipo: (actual.kg / kgReferencia < 0.60 && !actual.isLast) ? "PROCESANDO" : "NORMAL",
                calidad: calidadLote, isLast: actual.isLast
             });
          }
      }
  }

  return { 
    wip: resultado, familias: [...new Set(Object.values(mapFamily))].sort(),
    metadatosCodigos: mapaCodigosFiltro 
  };
}

// NUEVO: MARCAR LOTE COMO CONCLUIDO
function marcarLoteConcluido(nombreLote) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("LOTES");
  var data = sheet.getDataRange().getValues();
  
  // Indices LOTES (A=0): E:LOTE(4), H:ESTADO(7)
  for(var i=1; i<data.length; i++) {
     if(String(data[i][4]).trim() == String(nombreLote).trim()) {
         sheet.getRange(i+1, 8).setValue("CONCLUIDO"); // Col H
         return "OK";
     }
  }
  return "Lote no encontrado";
}

function enviarImagenTelegram(base64Data, caption, chatID_Destino, threadID) { // <--- Agregado threadID
  // --- CONFIGURACIÓN DE ENVÍO ---
  var BOT_TOKEN = "7947767393:AAFmZUcSTnV5gvP6u_UsBcSHlz-0s9x1kSQ";
  var CHAT_ID_EFICIENCIA = "-1003690353557"; 

  // --- LÓGICA DE DESTINO ---
  var finalChatId = chatID_Destino || CHAT_ID_EFICIENCIA;
  
  // --- PREPARACIÓN DE LA IMAGEN ---
  var data = base64Data.split(",")[1]; 
  var decoded = Utilities.base64Decode(data);
  var blob = Utilities.newBlob(decoded, "image/png", "Reporte.png");
  var url = "https://api.telegram.org/bot" + BOT_TOKEN + "/sendPhoto";
  
  // --- PREPARACIÓN DEL ENVÍO ---
  var payload = { 
    "chat_id": String(finalChatId), 
    "photo": blob, 
    "caption": caption || "Reporte",
    "parse_mode": "Markdown" 
  };

  // --- AJUSTE QUIRÚRGICO: Si hay threadID, lo metemos al envío ---
  if (threadID) {
    payload["message_thread_id"] = String(threadID);
  }

  var options = { "method": "post", "payload": payload, "muteHttpExceptions": true };

  try {
    var response = UrlFetchApp.fetch(url, options);
    return "✅ Enviado";
  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA NuevaCapturaHTML (MENU NUEVA CAPTURA)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

// 1. CARGA AL ABRIR LA APP (Para ver Procesos y Operadores)
function obtenerDatosInicialesNvo() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOps = ss.getSheetByName("OPERADORES");
    var sheetMenu = ss.getSheetByName("MENU_OPERATIVO");
    var sheetStd = ss.getSheetByName("ESTANDARES");

    var procesos = sheetMenu.getDataRange().getValues().slice(1)
      .filter(r => String(r[1]).toUpperCase().trim() == "PROCESO")
      .map(r => ({ nombre: r[0], img: r[2] }));

    var maquinas = sheetStd.getDataRange().getValues().slice(1).map(r => ({
      proc: String(r[2]).toUpperCase().trim(),
      maq: String(r[3]).toUpperCase().trim(),
      grupo: String(r[9]).toUpperCase().trim()
    }));

    var operadores = sheetOps.getDataRange().getValues().slice(1).map(r => ({
      id: r[0],
      nombre: r[1],
      procs: String(r[2]).toUpperCase()
    }));

    return { procesos: procesos, maquinas: maquinas, operadores: operadores };
  } catch (e) {
    throw new Error("Error en carga inicial: " + e.toString());
  }
}

// 2. CARGA MASIVA DE UN PROCESO (La que evita que se trabe el Tool CN)
function obtenerTodoElProcesoNvo(nombreProceso) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetOrd  = ss.getSheetByName("ORDENES");
    var sheetProd = ss.getSheetByName("PRODUCCION");
    var sheetLotes= ss.getSheetByName("LOTES");
    var sheetStd  = ss.getSheetByName("ESTANDARES");

    // 1. Máquinas y grupos del proceso
    var dataStd = sheetStd.getDataRange().getValues();
    var maquinasProceso = [], nombresMaquinas = [];
    dataStd.slice(1).forEach(function(r) {
      if (String(r[2]).toUpperCase().trim() == nombreProceso.toUpperCase().trim()) {
        var m = String(r[3]).toUpperCase().trim();
        maquinasProceso.push({ maq: m, grupo: String(r[9]) });
        nombresMaquinas.push(m);
      }
    });

    // 2. IDs de órdenes con producción en los últimos 2 días
    //    (para mostrar aunque estén TERMINADAS/SOBREPRODUCCION)
    var hoy = new Date(); hoy.setHours(0,0,0,0);
    var dosAtras = new Date(hoy); dosAtras.setDate(dosAtras.getDate() - 2);
    var lastProdRow = sheetProd.getLastRow();
    var ordenesConProdReciente = {};
    if (lastProdRow > 1) {
      // Leer solo las últimas 3000 filas de producción para no saturar
      var startProdRow = Math.max(2, lastProdRow - 2999);
      var numProdRows  = lastProdRow - startProdRow + 1;
      var dataProdFiltro = sheetProd.getRange(startProdRow, 1, numProdRows, 6).getValues(); // cols A-F
      for (var pf = 0; pf < dataProdFiltro.length; pf++) {
        var fProd = dataProdFiltro[pf][5]; // Col F: FECHA
        if (!(fProd instanceof Date)) continue;
        var fProdD = new Date(fProd); fProdD.setHours(0,0,0,0);
        if (fProdD >= dosAtras) {
          var idOrdProd = String(dataProdFiltro[pf][2]); // Col C: ID_ORDEN
          if (idOrdProd) ordenesConProdReciente[idOrdProd] = true;
        }
      }
    }

    // 3. Órdenes: vivas + TERMINADAS/SOBREPRODUCCION con prod reciente
    var rawOrd = sheetOrd.getDataRange().getValues();
    var EXCLUIDOS = ["CANCELADO"]; // Solo excluir CANCELADO — TERMINADO/SOBREPRODUCCION se evalúan por prod reciente
    var ordenesLean = [];
    for (var i = 1; i < rawOrd.length; i++) {
      var r = rawOrd[i];
      var estado   = String(r[15]).toUpperCase().trim();
      var maqOrden = String(r[12]).toUpperCase().trim();
      var idOrden  = String(r[0]);

      if (EXCLUIDOS.indexOf(estado) > -1) continue; // Siempre excluir CANCELADO

      var pertenece = nombresMaquinas.some(function(nm){ return maqOrden.includes(nm); });
      if (!pertenece) continue;

      // Si está TERMINADA o SOBREPRODUCCION, solo mostrar si tuvo prod en últimos 2 días
      var estaTerminada = (estado === "TERMINADO" || estado === "SOBREPRODUCCION");
      if (estaTerminada && !ordenesConProdReciente[idOrden]) continue;

      ordenesLean.push({
        id:        idOrden,
        serie:     String(r[4]),
        num:       ("0000" + parseInt(r[5])).slice(-4), // parseInt evita "83.0"
        tipo:      String(r[19]),
        medidas:   r[20] + " x " + r[21],
        acero:     String(r[24]),
        sol:       Number(r[13]) || 0,
        prod:      Number(r[14]) || 0,
        prioridad: Number(r[26]) || 999,
        maq:       maqOrden,
        estado:    estado, // útil para mostrar badge en el HTML si se quiere
        codigo:    String(r[6]).trim().toUpperCase()
      });
    }

    // 4. Historial de producción (todas las filas)
    var dataProd = lastProdRow > 1
      ? sheetProd.getRange(2, 1, lastProdRow - 1, 15).getValues()
      : [];
    var historial = {}, ultimasPorMaquina = {};

    for (var p = dataProd.length - 1; p >= 0; p--) {
      var oID    = String(dataProd[p][2]); if (!oID || oID === "0") continue;
      var maqC   = String(dataProd[p][4]).toUpperCase();
      var loteStr= String(dataProd[p][3]).trim();

      if (!ultimasPorMaquina[maqC]) ultimasPorMaquina[maqC] = oID;
      if (!historial[oID]) historial[oID] = { maxLote: 0, loteWeights: {}, lastFullRecord: null };

      if (historial[oID].lastFullRecord === null) {
        var f = dataProd[p][5];
        historial[oID].lastFullRecord = {
          fecha:   (f instanceof Date) ? Utilities.formatDate(f, "GMT-6", "dd/MM/yyyy") : String(f),
          turno:   dataProd[p][6],
          maquina: maqC,
          lote:    loteStr,
          pi:      Number(dataProd[p][8])  || 0,
          pf:      Number(dataProd[p][9])  || 0,
          tina:    Number(dataProd[p][9]) - Number(dataProd[p][8]) - Number(dataProd[p][10]),
          prod:    Number(dataProd[p][10]) || 0,
          sello:   String(dataProd[p][14])
        };
      }

      if (loteStr.includes('.')) {
        if (historial[oID].loteWeights[loteStr] === undefined)
          historial[oID].loteWeights[loteStr] = Number(dataProd[p][9]);
        var nLote = parseInt(loteStr.split('.').pop());
        if (!isNaN(nLote) && nLote > historial[oID].maxLote) historial[oID].maxLote = nLote;
      }
      // También intentar extraer número de lote del formato sin punto
      if (historial[oID].maxLote === 0) {
        var nLoteAlt = parseInt(loteStr);
        if (!isNaN(nLoteAlt) && nLoteAlt > historial[oID].maxLote)
          historial[oID].maxLote = nLoteAlt;
      }
    }

    // 5. Lotes existentes — lee TODA la tabla sin límite
    var lotesExistentes = _getLotesExistentes(sheetLotes);

    return JSON.stringify({
      success:           true,
      maquinas:          maquinasProceso,
      ordenes:           ordenesLean,
      historial:         historial,
      ultimasPorMaquina: ultimasPorMaquina,
      lotesExistentes:   lotesExistentes
    });

  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

// _getLotesExistentes — FIX: MAX_LOTES = 99999 para cubrir tablas grandes
function _getLotesExistentes(sheetLotes) {
  var lastRow = sheetLotes.getLastRow();
  if (lastRow < 2) return [];
  var MAX_LOTES = 99999; // sin límite práctico
  var startRow = Math.max(2, lastRow - MAX_LOTES + 1);
  var numRows  = lastRow - startRow + 1;
  return sheetLotes
    .getRange(startRow, 5, numRows, 1)
    .getValues()
    .map(function(r){ return String(r[0]).trim(); })
    .filter(function(s){ return s !== ""; });
}


// 3. BUSCADOR DE ÓRDENES TERMINADAS (Historial 90 días)
function buscarTerminadasNvo(proceso, texto) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var data = ss.getSheetByName("ORDENES").getRange(1,1,ss.getSheetByName("ORDENES").getLastRow(), 27).getDisplayValues();
  var res = [];
  var limite = new Date(); limite.setDate(limite.getDate() - 90);
  for(var i=1; i<data.length; i++) {
    var est = data[i][15].toUpperCase();
    var f = new Date(data[i][3]);
    var nombre = data[i][4] + "." + ("0000" + data[i][5]).slice(-4);
    if((est == "TERMINADO" || est == "SOBREPRODUCCION") && nombre.includes(texto) && f >= limite && data[i][11].toUpperCase().includes(proceso.toUpperCase())) {
      res.push({
        id: String(data[i][0]), 
        serie: String(data[i][4]), 
        nombre: nombre, 
        tipo: data[i][19],
        medidas: data[i][20] + " x " + data[i][21], 
        acero: data[i][24],
        sol: Number(data[i][13].replace(/,/g,'')), 
        prod: Number(data[i][14].replace(/,/g,'')),
        prioridad: 999, 
        estado: est, 
        maquina: data[i][12]
      });
    }
  }
  return res;
}

function guardarProduccionCompleta(payload) {
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(15000)) return { success: false, msg: "Servidor ocupado." };
  
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd = ss.getSheetByName("PRODUCCION");
    var sheetLotes = ss.getSheetByName("LOTES");
    var sheetOrd = ss.getSheetByName("ORDENES");
    var sheetPed = ss.getSheetByName("PEDIDOS");
    var sheetEnv = ss.getSheetByName("ENVIADO");

    var registros = payload.registros; 
    var ordenesAfectadas = new Set();
    var lotesParaActualizar = new Set();

    // --- A. CREACIÓN DE LOTES FALTANTES ---
    var dataLotes = sheetLotes.getDataRange().getValues();
    var maxIdLote = 0;
    for(var i=1; i<dataLotes.length; i++) {
      var valId = parseInt(dataLotes[i][0]);
      if(!isNaN(valId) && valId > maxIdLote) maxIdLote = valId;
    }

    registros.forEach(r => {
      if (r.crearLote) {
        // Doble verificación: ¿Ya existe el lote en la hoja?
        var existeYa = false;
        var filaReferencia = -1;
        for(var l=dataLotes.length-1; l>=1; l--) {
          if (String(dataLotes[l][4]) == r.loteFull) { existeYa = true; break; }
          // Buscamos el último lote de esta misma orden para copiar SOLICITADO y QR base
          if (filaReferencia == -1 && String(dataLotes[l][2]) == r.ordenID) {
            filaReferencia = l;
          }
        }

        if (!existeYa) {
          maxIdLote++;
          var numLoteFormato = Number(r.numLote) < 100 ? ("00" + r.numLote).slice(-2) : String(r.numLote);
          var qrOriginal = (filaReferencia > -1) ? String(dataLotes[filaReferencia][10]) : "";
          var nuevoQr = qrOriginal !== "" ? qrOriginal.substring(0, qrOriginal.length - 2) + numLoteFormato : "";
          
          var nuevaFilaLote = [
            maxIdLote,                                      // A: ID
            r.serie,                                        // B: SERIE
            r.ordenID,                                      // C: ORDEN
            parseInt(r.numLote),                            // D: CONSECUTIVO
            r.loteFull,                                     // E: LOTE
            (filaReferencia > -1) ? dataLotes[filaReferencia][5] : 0, // F: SOLICITADO
            0,                                              // G: PRODUCIDO
            "ABIERTO",                                      // H: ESTADO
            new Date(),                                     // I: FECHA_REG
            "",                                             // J: SELLO
            nuevoQr                                         // K: QR_CODE
          ];
          sheetLotes.appendRow(nuevaFilaLote);
          // Actualizamos memoria local para evitar duplicar IDs en el mismo proceso
          dataLotes.push(nuevaFilaLote); 
        }
      }
    });

    // --- B. INSERTAR EN PRODUCCION ---
    var nuevasFilasProd = [];
    registros.forEach(r => {
       var horas = (r.turno == "2") ? 7.0 : (r.turno == "3" ? 8.0 : 7.5);
       var rowProd = [
          Utilities.getUuid(), r.serie, r.ordenID, r.loteFull, r.maquina,    // A-E
          new Date(r.fecha + "T12:00:00"), r.turno, r.op1_id, r.pesoI, r.pesoF, // F-J
          r.producido, "", "", horas, r.sello || "", r.comentarios || "",    // K-P
          r.numLote, r.op1_txt, r.op2_id, r.op2_txt, "", "", r.pesoTina, "", // Q-X (U queda vacía)
          r.usuario || "",          // Y: USER
          new Date(),               // Z: FECHA_REG
          "",                       // AA: (vacío)
          r.proceso || ""           // AB: PROCESO
       ];
       nuevasFilasProd.push(rowProd);
       ordenesAfectadas.add(r.ordenID);
       lotesParaActualizar.add(r.loteFull);
    });

    if(nuevasFilasProd.length > 0) {
       sheetProd.getRange(sheetProd.getLastRow()+1, 1, nuevasFilasProd.length, nuevasFilasProd[0].length).setValues(nuevasFilasProd);
    }
    SpreadsheetApp.flush();

    // === NUEVO: VINCULAR UUIDs PARA RASTREO ===
    // Esto mapea el ID único de cada fila recién guardada al objeto de memoria
    registros.forEach((r, idx) => r.idProdInterno = nuevasFilasProd[idx][0]);

    // --- C. ACTUALIZAR SUMATORIA EN LOTES ---
    var dataLotesActualizado = sheetLotes.getDataRange().getValues();
    var dataProdActualizado = sheetProd.getDataRange().getValues();
    
    lotesParaActualizar.forEach(lKey => {
       var sumaLote = 0;
       var filaLoteHoja = -1;
       var solLote = 0;
       for(var p=1; p<dataProdActualizado.length; p++) {
          if(String(dataProdActualizado[p][3]) == lKey) sumaLote += (Number(dataProdActualizado[p][10]) || 0); 
       }
       for(var l=1; l<dataLotesActualizado.length; l++) {
          if(String(dataLotesActualizado[l][4]) == lKey) { 
             filaLoteHoja = l+1; 
             solLote = Number(dataLotesActualizado[l][5]); 
             break; 
          }
       }
       if(filaLoteHoja > -1) {
           sheetLotes.getRange(filaLoteHoja, 7).setValue(sumaLote); 
           var nvoEst = (sumaLote >= solLote) ? "CERRADO" : (sumaLote > 0 ? "EN PROCESO" : "ABIERTO");
           sheetLotes.getRange(filaLoteHoja, 8).setValue(nvoEst); 
       }
    });

    // --- D. ACTUALIZAR ORDENES ---
    var dataOrd = sheetOrd.getDataRange().getValues();
    var pedidosAfectados = new Set();

    ordenesAfectadas.forEach(ordID => {
       var filaOrd = -1; var solOrd = 0; var idPedido = ""; var partida = ""; var indexMemoria = -1;
       
       for(var o=1; o<dataOrd.length; o++) {
          if(String(dataOrd[o][0]) == String(ordID)) {
             filaOrd = o+1; indexMemoria = o;
             solOrd = Number(dataOrd[o][13]); 
             idPedido = dataOrd[o][1]; 
             partida = dataOrd[o][2]; 
             break;
          }
       }
       
       if(filaOrd > -1) {
           var sumaOrd = 0;
           for(var p=1; p<dataProdActualizado.length; p++) {
              if(String(dataProdActualizado[p][2]) == String(ordID)) sumaOrd += (Number(dataProdActualizado[p][10]) || 0);
           }
           
           var estadoAnterior = String(dataOrd[indexMemoria][15]).toUpperCase();
           sheetOrd.getRange(filaOrd, 15).setValue(sumaOrd); 
           dataOrd[indexMemoria][14] = sumaOrd; 

           var ratio = (solOrd > 0) ? (sumaOrd / solOrd) : 0;
           var stOrd = "ABIERTO";
           if (sumaOrd === 0) stOrd = "ACTIVE";
           else if (ratio > 0 && ratio < 1.05) stOrd = "EN PROCESO";
           else if (ratio >= 1.05 && ratio < 1.10) stOrd = "TERMINADO";
           else if (ratio >= 1.10) stOrd = "SOBREPRODUCCION";

           sheetOrd.getRange(filaOrd, 16).setValue(stOrd); 

           // --- BLOQUE QUIRÚRGICO DE NOTIFICACIÓN ---
           var estadosNotificar = ["EN PROCESO", "TERMINADO", "SOBREPRODUCCION", "ACTIVE"];
           if (stOrd !== estadoAnterior && estadosNotificar.includes(stOrd)) {
              
              // FIX: Buscar el registro dentro del payload original para esta orden
              var regOrigen = registros.find(function(x) { return x.ordenID == ordID; });
              var maqCapturada = regOrigen ? regOrigen.maquina : "N/A";
              var prodCapturada = regOrigen ? regOrigen.producido : 0;
              
              var procNom = String(dataOrd[indexMemoria][11]); 
              var secNom = String(dataOrd[indexMemoria][10]);
              
              enviarMensajeOrdenTelegram(dataOrd[indexMemoria], sumaOrd, stOrd, procNom, secNom, maqCapturada, prodCapturada); 
           }
           // ------------------------------------------

           if (stOrd == "TERMINADO" || stOrd == "SOBREPRODUCCION") {
               var serieC = String(dataOrd[indexMemoria][4]);
               var numC = String(dataOrd[indexMemoria][5]);
               var secC = Number(dataOrd[indexMemoria][10]);
               var maxSec = 0;
               for(var x=1; x<dataOrd.length; x++){
                 if(String(dataOrd[x][4])==serieC && String(dataOrd[x][5])==numC){
                   var s = Number(dataOrd[x][10]); if(s > maxSec) maxSec = s;
                 }
               }
               if(secC == maxSec) {
                 for(var y=1; y<dataOrd.length; y++){
                   if(String(dataOrd[y][4])==serieC && String(dataOrd[y][5])==numC){
                     var estY = String(dataOrd[y][15]).toUpperCase();
                     if(estY != "SOBREPRODUCCION" && estY != "TERMINADO" && estY != "CANCELADO"){
                        sheetOrd.getRange(y+1, 16).setValue("TERMINADO");
                     }
                   }
                 }
               }
           }
           if(idPedido) pedidosAfectados.add(idPedido + "|" + partida);
       }
    });

    // --- E. ACTUALIZAR PEDIDOS ---
    var dataPed = sheetPed.getDataRange().getValues();
    pedidosAfectados.forEach(key => {
       var pts = key.split("|");
       var pID = pts[0]; var pPart = pts[1];
       for(var p=1; p<dataPed.length; p++) {
          if(String(dataPed[p][1]) == pID && String(dataPed[p][5]) == pPart) {
             var estActual = String(dataPed[p][8]).toUpperCase();
             if(estActual != "CANCELADO" && estActual != "CERRADO" && estActual != "TERMINADO") {
                // Aquí podrías recalcular si es PARA ENVIAR o EN PROCESO
                sheetPed.getRange(p+1, 9).setValue("EN PROCESO");
             }
             break;
          }
       }
    });

    // === NUEVO: DISPARAR VALIDACIÓN DE ACEROS ===
    // Solo se ejecuta una vez que todo lo demás se guardó correctamente
    validarAcerosProduccion(registros);

    return { success: true, msg: "Producción y Lotes registrados correctamente." };

  } catch(e) {
    return { success: false, msg: "Error: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// --- FUNCIÓN PARA BUSCAR ÓRDENES TERMINADAS EN EL PASADO (60 DÍAS) ---
// XXX VERIFICAR SI HAY QUE ELIMINARSE 
function obtenerOrdenesTerminadasPasado(maquina, textoBusqueda, nombreProceso) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetRut = ss.getSheetByName("RUTAS");
  
  // A. OBTENER RUTAS (Misma lógica de tu función original)
  var dataRut = sheetRut.getDataRange().getValues();
  var mapMaquinasValidas = {}; 
  for(var r=1; r<dataRut.length; r++) {
     var key = String(dataRut[r][1]).trim() + "|" + String(dataRut[r][4]).toUpperCase().trim() + "|" + String(dataRut[r][3]).trim(); 
     if(dataRut[r][5]) mapMaquinasValidas[key] = String(dataRut[r][5]).trim();
  }

  // B. FILTRAR ÓRDENES
  var dataOrdRaw = sheetOrd.getDataRange().getValues();
  var displayDataOrd = sheetOrd.getDataRange().getDisplayValues();
  var ordenesEncontradas = [];
  var limiteFecha = new Date();
  limiteFecha.setDate(limiteFecha.getDate() - 60);

  for(var i=1; i<dataOrdRaw.length; i++) {
     var est = String(dataOrdRaw[i][15]).toUpperCase(); // Col P
     var fOrden = new Date(dataOrdRaw[i][3]); // Col D
     var nombreFull = displayDataOrd[i][4] + "." + ("0000" + displayDataOrd[i][5]).slice(-4);
     var maqRow = String(dataOrdRaw[i][12]).toUpperCase();

     // CONDICIONES: 
     // 1. Estado TERMINADO o SOBREPRODUCCION
     // 2. Que contenga el texto buscado
     // 3. Que pertenezca a la máquina actual (o sea válida para el proceso)
     // 4. Que sea de máximo 60 días atrás
     if((est == "TERMINADO" || est == "SOBREPRODUCCION") && 
        nombreFull.includes(textoBusqueda) && 
        maqRow.includes(maquina.toUpperCase()) &&
        fOrden >= limiteFecha) {

         var oid = String(dataOrdRaw[i][0]);
         ordenesEncontradas.push({
            id: oid,
            serie: displayDataOrd[i][4],
            nombre: nombreFull,
            tipo: displayDataOrd[i][19],
            medidas: displayDataOrd[i][20] + " x " + displayDataOrd[i][21],
            detalles: displayDataOrd[i][23] + " / " + displayDataOrd[i][22],
            acero: displayDataOrd[i][24],
            sol: Number(dataOrdRaw[i][13])||0,
            prod: Number(dataOrdRaw[i][14])||0,
            pendiente: (Number(dataOrdRaw[i][13])||0) - (Number(dataOrdRaw[i][14])||0),
            avance: (Number(dataOrdRaw[i][13])>0) ? (Number(dataOrdRaw[i][14])/Number(dataOrdRaw[i][13]))*100 : 0,
            prioridad: Number(dataOrdRaw[i][26]) || 999,
            esUltima: false,
            maquinasOpciones: mapMaquinasValidas[String(dataOrdRaw[i][6]) + "|" + String(dataOrdRaw[i][11]).toUpperCase() + "|" + String(dataOrdRaw[i][10])] || maqRow
         });
     }
  }
  return ordenesEncontradas;
}


// ================Sección para enviar notificación al telegram con las capturas de CapturaProduccionHTML====================================
// 1. OBTENER DATOS PARA PANTALLA DETALLE (MÓVIL)
// XXX VERIFICAR SI HAY QUE ELIMINAR
function obtenerDatosPantallaNotif(id) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sOrd = ss.getSheetByName("ORDENES");
  var sInv = ss.getSheetByName("INVENTARIO_EXTERNO");
  var sProd = ss.getSheetByName("PRODUCCION");

  var dOrd = sOrd.getDataRange().getValues();
  var principal = dOrd.find(r => String(r[0]) == id);
  if(!principal) return { error: "No encontrada" };

  var serieP = principal[4];
  var numeroP = principal[5];
  var codigoP = principal[6];

  // Inventario
  var inv = { ex: 0, min: 0, max: 0, bo: 0 };
  var dInv = sInv.getDataRange().getValues();
  var rInv = dInv.find(r => String(r[0]) == codigoP);
  if(rInv) inv = { ex: rInv[1], min: rInv[2], max: rInv[3], bo: rInv[4] };

  // Procesos de la Orden Principal
  var procesos = dOrd.filter(o => o[4] == serieP && o[5] == numeroP).map(o => {
    return { id: o[0], sec: o[10], proc: o[11], sol: o[13], prod: o[14], est: o[15] };
  });

  // Ordenes Vivas Relacionadas (Agrupadas)
  var vivasMap = {};
  dOrd.forEach(o => {
    var key = o[4] + "." + o[5];
    if(o[6] == codigoP && key !== (serieP + "." + numeroP) && ["ABIERTO","ACTIVE","EN PROCESO"].includes(o[15])) {
      if(!vivasMap[key]) {
        vivasMap[key] = { 
        id: o[0], 
        ord: key, 
        cant: o[17], 
        sol: o[13],       // Añadido: Solicitado en kg
        prod: o[14],      // Añadido: Producido en kg
        procesos: [] 
      };
    }
      vivasMap[key].procesos.push(o[11]);
    }
  });

  // Historial de producción de esta orden
  var historial = [];
  var dProd = sProd.getDataRange().getValues();
  for(var i = dProd.length-1; i>=1; i--) {
    if(String(dProd[i][2]) == id) {
      historial.push({ id: dProd[i][0], fecha: Utilities.formatDate(dProd[i][5], "GMT-6", "dd/MM/yy"), maq: dProd[i][4], neto: dProd[i][10], lote: dProd[i][3] });
      if(historial.length > 10) break;
    }
  }

  return { 
    inv: inv, 
    principal: { id: id, ord: serieP + "." + numeroP, cant: principal[17], desc: principal[19]+" "+principal[20], un: principal[9] },
    procesos: procesos, 
    vivas: Object.values(vivasMap), 
    historial: historial 
  };
}

// 2. FUNCIÓN MAESTRA: GUARDADO MASIVO (CANTIDAD, ESTADOS, HISTORIAL)
// XXX VERIFICAR SI HAY QUE ELIMINAR
function procesarGranCambioDetalle(payload) {
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(15000)) return { success: false, msg: "Servidor ocupado." };
  
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sOrd = ss.getSheetByName("ORDENES");
    var sProd = ss.getSheetByName("PRODUCCION");
    var dOrd = sOrd.getDataRange().getValues();
    var dProd = sProd.getDataRange().getValues();
    
    var ordenesAfectadas = new Set();
    ordenesAfectadas.add(payload.idPrincipal);

    // A. ACTUALIZAR ÓRDENES (CANTIDAD GLOBAL O ESTADOS)
    payload.cambiosOrdenes.forEach(c => {
      var ref = dOrd.find(r => String(r[0]) == c.id);
      if(ref) {
        var serie = ref[4]; var numero = ref[5];
        for(var i=1; i<dOrd.length; i++) {
          if(String(dOrd[i][4]) == String(serie) && String(dOrd[i][5]) == String(numero)) {
            var fila = i + 1;
            if(c.est) {
              sOrd.getRange(fila, 16).setValue(c.est);
            } else if(c.cant !== undefined) {
              var nCant = Number(c.cant);
              var unidad = String(dOrd[i][9]).toUpperCase();
              var tipo = String(dOrd[i][19]).toUpperCase();
              var peso = Number(dOrd[i][18]) || 0;
              var nSolicitado = nCant;

              if(unidad === "PZA" || unidad === "CTO") {
                if(tipo.includes("VARILLA")) {
                  var lonRaw = String(dOrd[i][21]);
                  var m = lonRaw.match(/[\d\.]+/);
                  var lonNum = m ? parseFloat(m[0]) : 1;
                  nSolicitado = nCant * peso * lonNum;
                } else { nSolicitado = nCant * peso; }
              }
              sOrd.getRange(fila, 9).setValue(nCant);       // Col I
              sOrd.getRange(fila, 14).setValue(nSolicitado); // Col N
              if(Number(dOrd[i][14]) < nSolicitado) sOrd.getRange(fila, 16).setValue("EN PROCESO");
            }
          }
        }
      }
    });

    // B. ACTUALIZAR HISTORIAL (EDICIÓN Y BORRADO)
    var registrosABorrar = payload.cambiosHistorial.filter(h => h.borrar).map(h => {
       for(var i=1; i<dProd.length; i++) { if(String(dProd[i][0]) == h.id) return i+1; }
    }).filter(f => f != null).sort((a,b) => b-a);
    registrosABorrar.forEach(f => sProd.deleteRow(f));

    payload.cambiosHistorial.filter(h => !h.borrar && h.neto).forEach(h => {
      for(var i=1; i<dProd.length; i++) {
        if(String(dProd[i][0]) == h.id) {
          var fila = i + 1;
          sProd.getRange(fila, 11).setValue(Number(h.neto));
          var nPF = Number(dProd[i][8]) + Number(dProd[i][22] || 0) + Number(h.neto);
          sProd.getRange(fila, 10).setValue(nPF);
          break;
        }
      }
    });

    SpreadsheetApp.flush();
    ordenesAfectadas.forEach(id => recalcularEstadoOrden(sOrd, sProd, id));
    return { success: true, msg: "Cambios aplicados correctamente." };

  } catch(e) { return { success: false, msg: e.toString() }; } 
  finally { lock.releaseLock(); }
}

// 3. NOTIFICACIÓN QUIRÚRGICA DE TELEGRAM
function enviarMensajeOrdenTelegram(filaOrd, producidoTotal, estado, procesoNombre, secuencia, maquina, netoMovimiento) {
  var webAppUrl = ScriptApp.getService().getUrl();
  var ordenNombre = filaOrd[4] + "." + ("0000" + filaOrd[5]).slice(-4);
  
  // Forzar a String para evitar conversiones a fecha
  var tipo = String(filaOrd[19]);
  var diam = String(filaOrd[20]); // Fix fecha
  var long = String(filaOrd[21]); // Fix fecha
  var productoFull = tipo + " " + diam + " X " + long;

  var solicitado = filaOrd[13];
  
  var msj = "🚀 *CAMBIO DE ESTADO: " + estado + "*\n\n";
  msj += "🏭 *Proceso:* [" + secuencia + "] " + procesoNombre + "\n";
  msj += "🖥 *Máquina:* " + (maquina || "N/A") + "\n";
  msj += "📥 *Captura:* +" + (netoMovimiento || 0).toLocaleString() + " kg\n\n";
  msj += "📦 *Orden:* " + ordenNombre + "\n";
  msj += "🛠 *Item:* " + productoFull + "\n";
  msj += "📊 *Progreso:* " + Math.round(producidoTotal).toLocaleString() + " / " + Math.round(solicitado).toLocaleString() + " kg\n";

  var urlDetalle = webAppUrl + "?v=notif_orden&id=" + filaOrd[0];

  var payload = {
    "chat_id": CHAT_ID_AVISOS,
    "text": msj,
    "parse_mode": "Markdown",
    "reply_markup": JSON.stringify({
      "inline_keyboard": [[{ "text": "🔍 Gestionar Orden / Ver Detalles", "url": urlDetalle }]]
    })
  };
  
  UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN_TELEGRAM + "/sendMessage", {
    "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true
  });
}

// 4. MOTOR DE ESTADOS (DISPARA TELEGRAM)
function recalcularEstadoOrden(sheetOrd, sheetProd, idOrden) {
  if(!sheetOrd) sheetOrd = SpreadsheetApp.openById(ID_HOJA_CALCULO).getSheetByName("ORDENES");
  if(!sheetProd) sheetProd = SpreadsheetApp.openById(ID_HOJA_CALCULO).getSheetByName("PRODUCCION");

  var dataProd = sheetProd.getDataRange().getValues();
  var dataOrd = sheetOrd.getDataRange().getValues();
  
  var suma = 0;
  for(var i=1; i<dataProd.length; i++) {
    if(String(dataProd[i][2]) == String(idOrden)) suma += (Number(dataProd[i][10]) || 0);
  }
  
  for(var j=1; j<dataOrd.length; j++) {
    if(String(dataOrd[j][0]) == String(idOrden)) {
       var sol = Number(dataOrd[j][13]); 
       var estAnterior = String(dataOrd[j][15]).toUpperCase();
       
       // CAPTURA QUIRÚRGICA
       var procNom = String(dataOrd[j][11]); // Col L
       var secNom = String(dataOrd[j][10]);  // Col K

       if(estAnterior == "CANCELADO") return;

       var nuevoEst = "ACTIVE"; 
       if(suma > 0) {
          var ratio = (sol > 0) ? (suma / sol) : 0;
          if(ratio < 1.05) nuevoEst = "EN PROCESO";
          else if (ratio < 1.10) nuevoEst = "TERMINADO";
          else nuevoEst = "SOBREPRODUCCION";
       }

       sheetOrd.getRange(j+1, 15).setValue(suma);      
       sheetOrd.getRange(j+1, 16).setValue(nuevoEst);

       var estadosInteres = ["EN PROCESO", "TERMINADO", "SOBREPRODUCCION", "ACTIVE"];
       if (nuevoEst !== estAnterior && estadosInteres.includes(nuevoEst)) {
         // Llamada con los 5 parámetros
         enviarMensajeOrdenTelegram(dataOrd[j], suma, nuevoEst, procNom, secNom);
       }
       break;
    }
  }
}

// =============

function actualizarSelloDesdeTelegram(chatId, msgIdOrig, nSello) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var data = sheetProd.getDataRange().getValues();
  var f = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][23]) == String(msgIdOrig)) { f = i + 1; break; }
  }

  if (f > -1) {
    var row = data[f-1];
    sheetProd.getRange(f, 15).setValue(nSello); // Actualizar Sello
    borrarMensajeTelegram(chatId, msgIdOrig);  // Borrar mensaje viejo
    sheetProd.getRange(f, 24).setValue("");    // Limpiar ID_MSG
    SpreadsheetApp.flush();

    // Re-validar leyendo TODA la fila para que el mensaje no salga incompleto
    var reg = {
      idProdInterno: row[0],
      ordenID: row[2],
      loteFull: row[3],
      maquina: row[4],
      fecha: Utilities.formatDate(new Date(row[5]), "GMT-6", "dd/MM/yy"),
      turno: row[6],
      producido: row[10],
      sello: nSello
    };
    
    validarAcerosProduccion([reg]);
    
    SpreadsheetApp.flush();
    if (sheetProd.getRange(f, 24).getValue() == "") {
       enviarTextoTelegram_Interno(chatId, "✅ *¡Cambio Exitoso!*\nEl sello `" + nSello + "` es correcto.", THREAD_ID_SELLOS);
    }
  } else {
    enviarTextoTelegram_Interno(chatId, "🙅🏻‍♂️ Ya no se puede modificar este registro.", THREAD_ID_SELLOS);
  }
}

function validarAcerosProduccion(registrosNuevos, esRevalidacion) {
  var ssProd = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var ssMP = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheetProd = ssProd.getSheetByName("PRODUCCION");
  var sheetOrd = ssProd.getSheetByName("ORDENES");
  var sheetMP = ssMP.getSheetByName("ENTRADAS_MP");
  
  var dMP = sheetMP.getDataRange().getValues();
  var dOrd = sheetOrd.getDataRange().getValues();

  registrosNuevos.forEach(reg => {
    var infoO = dOrd.find(o => String(o[0]) == reg.ordenID);
    if (!infoO) return;
    
    var proceso = String(infoO[11]).toUpperCase();
    if (!["FORJA", "PUNTEADO", "ROLADO TORN"].includes(proceso)) return;
    
    var selloIngresado = String(reg.sello).trim();
    var infoM = dMP.find(m => String(m[8]).trim() == selloIngresado);
    var aceroOrden = String(infoO[24]); 
    var productoDesc = infoO[19] + " " + infoO[20] + " X " + infoO[21];
    
    var msjTecnico = "";
    var msjAmigable = "";
    var lanzarAlerta = false;

    if (!infoM) {
      lanzarAlerta = true;
      msjTecnico = "🚫🙅‍♂️ *SELLO NO EXISTE* 🙅‍♂️🚫\n" +
            "🎯 *Lote:* " + reg.loteFull + "\n" +
            "🎰 *Máq:* " + reg.maquina + "\n" +
            "🔩 *Prod:* " + productoDesc + "\n" +
            "💠 *Producción:* " + reg.producido + " kg\n" +
            "#️⃣ *Sello reg:* `" + selloIngresado + "`\n" +
            "❌ El sello no existe en la base de datos de MP.\n" +
            "La orden requiere Acero: *" + aceroOrden + "*";
      
      if (esRevalidacion) {
        msjAmigable = "😰 *El sello nuevo que registraste sigue estando mal*\n\n" +
                      "❌ El sello `" + selloIngresado + "` *NO EXISTE* en la base de datos de MP.\n\n" +
                      "Por favor corrige nuevamente 😬";
      }
    } else {
      var aceroMP = String(infoM[4]);
      var selloMP = String(infoM[8]);
      
      if (!fuzzySteelMatch(aceroOrden, aceroMP)) {
        lanzarAlerta = true;
        msjTecnico = "⚠️📛 *ACERO INCORRECTO* 📛⚠️\n" +
              "🎯 *Lote:* " + reg.loteFull + "\n" +
              "🎰 *Máq:* " + reg.maquina + "\n" +
              "🔩 *Prod:* " + productoDesc + "\n" +
              "💠 *Producción:* " + reg.producido + " kg\n" +
              "#️⃣ *Sello reg:* `" + selloIngresado + "` (" + aceroOrden + ")\n" +
              "🈴 *Base MP:* Sello `" + selloMP + "` (" + aceroMP + ")";
        
        if (esRevalidacion) {
          msjAmigable = "😰 *El sello nuevo que registraste sigue estando mal*\n\n" +
                        "El sello `" + selloIngresado + "` corresponde a un acero *" + aceroMP + "*\n" +
                        "pero la orden requiere un acero *" + aceroOrden + "*\n\n" +
                        "Por favor corrige nuevamente 😬";
        }
      }
    }

    if (lanzarAlerta) {
      // SIEMPRE enviar al TOPIC específico
      var res = enviarTextoTelegram_Interno(CHAT_ID_CALIDAD, msjTecnico, THREAD_ID_SELLOS);
      
      if (res && res.ok) {
        var dP = sheetProd.getDataRange().getValues();
        for(var i = dP.length - 1; i >= 1; i--) {
          if (String(dP[i][0]) == String(reg.idProdInterno)) {
            sheetProd.getRange(i + 1, 24).setValue(res.result.message_id);
            break;
          }
        }
        
        // Si es re-validación, enviar mensaje amigable AL MISMO TOPIC
        if (esRevalidacion && msjAmigable) {
          enviarTextoTelegram_Interno(CHAT_ID_CALIDAD, msjAmigable, THREAD_ID_SELLOS);
        }
      }
    }
  });
}

function enviarTextoTelegram_Interno(chatId, texto, threadId) {
  var url = "https://api.telegram.org/bot" + TOKEN + "/sendMessage";
  var payload = { "chat_id": String(chatId), "text": texto, "parse_mode": "Markdown" };
  if (threadId) { payload["message_thread_id"] = String(threadId); }
  
  try {
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options);
    var resText = response.getContentText();
    return resText ? JSON.parse(resText) : { ok: false };
  } catch (e) {
    return { ok: false, error: e.toString() }; 
  }
}


function fuzzySteelMatch(steelA, steelB) {
  // Extraer todos los dígitos consecutivos más largos
  var extraerNumeros = function(s) {
    var texto = String(s).toUpperCase().trim();
    
    // Buscar secuencias de 4 dígitos (con o sin letras intercaladas)
    // Ejemplo: "10B21" → "1021", "TER \ 10B21 CHQ AK" → "1021"
    var matches = texto.match(/\d+/g);
    
    if (!matches) return null;
    
    // Buscar el número más largo o el primero de 4 dígitos
    for (var i = 0; i < matches.length; i++) {
      if (matches[i].length >= 4) {
        return matches[i].substring(0, 4);
      }
    }
    
    // Si no hay de 4 dígitos, concatenar los que haya
    var concatenado = matches.join('');
    return concatenado.length >= 4 ? concatenado.substring(0, 4) : null;
  };
  
  var numA = extraerNumeros(steelA);
  var numB = extraerNumeros(steelB);
  
  // Si no se pudieron extraer números, no hay coincidencia
  if (!numA || !numB) return false;
  
  // Si los 4 dígitos son iguales, es correcto
  if (numA === numB) return true;
  
  // Grupo especial: 1006, 1008, 1010 son intercambiables
  var grupoEspecial = ["1006", "1008", "1010"];
  if (grupoEspecial.includes(numA) && grupoEspecial.includes(numB)) {
    return true;
  }
  
  return false;
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA TablerosupervisorHTML (MENU TABLERO SUPERVISOR)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

// 1. CARGA INICIAL: Solo procesos (igual que NuevaCapturaHTML)
function obtenerProcesosTablero() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetMenu = ss.getSheetByName("MENU_OPERATIVO");
  
  var procesos = sheetMenu.getDataRange().getValues().slice(1)
    .filter(r => String(r[1]).toUpperCase().trim() == "PROCESO")
    .map(r => ({ nombre: String(r[0]).toUpperCase().trim(), img: r[2] }));

  return { procesos: procesos };
}

// 2. CARGA MASIVA DE UN PROCESO: Órdenes + Lotes + Rutas + Docs (TODO EN UNA LLAMADA)
function obtenerOrdenesTableroLight(nombreProceso) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd   = ss.getSheetByName("ORDENES");
  var sheetPed   = ss.getSheetByName("PEDIDOS");
  var sheetLotes = ss.getSheetByName("LOTES");
  var sheetEst   = ss.getSheetByName("ESTANDARES");

  var dataOrd   = sheetOrd.getDataRange().getValues();
  var dataPed   = sheetPed.getDataRange().getValues();
  var dataLotes = sheetLotes.getDataRange().getValues();
  var dataEst   = sheetEst.getDataRange().getValues();

  // A. Mapa pedidos vivos
  var pedidosVivos = {};
  for (var p = 1; p < dataPed.length; p++) {
    var estPed = String(dataPed[p][8]).toUpperCase().trim();
    if (estPed !== "TERMINADO" && estPed !== "CANCELADO" && estPed !== "CERRADO") {
      pedidosVivos[String(dataPed[p][1]).trim().toUpperCase()] = true;
    }
  }

  // B. Mapa máquina→grupo desde ESTANDARES (Col D idx3=MAQUINA, Col J idx9=GRUPO)
  var maqGrupoMap  = {};
  var grupoMaqsMap = {};
  for (var s = 1; s < dataEst.length; s++) {
    var maqEst   = String(dataEst[s][3]).trim().toUpperCase();
    var grupoEst = String(dataEst[s][9]).trim().toUpperCase();
    if (!maqEst || !grupoEst) continue;
    maqGrupoMap[maqEst] = grupoEst;
    if (!grupoMaqsMap[grupoEst]) grupoMaqsMap[grupoEst] = [];
    if (grupoMaqsMap[grupoEst].indexOf(maqEst) === -1) grupoMaqsMap[grupoEst].push(maqEst);
  }

  // C. Filtrar órdenes — solo VIVAS del proceso
  var ordenes      = [];
  var serieOrdenMap = {};
  // Mapa extendido: TODAS las órdenes del proceso independiente de estado/pedido.
  // Necesario para que el modo Históricas también encuentre sus lotes en el cache.
  var serieOrdenMapParaLotes = {};

  for (var i = 1; i < dataOrd.length; i++) {
    var r      = dataOrd[i];
    var proc   = String(r[11]).toUpperCase().trim();
    var estOrd = String(r[15]).toUpperCase().trim();
    var pedido = String(r[1]).trim();

    if (proc !== nombreProceso.toUpperCase()) continue;

    var serie    = String(r[4]);
    var numOrd   = Number(r[5]);
    var serieOrd = serie + "." + ("0000" + numOrd).slice(-4);

    // Registrar en mapa extendido SIN filtros de estado ni pedido
    if (!serieOrdenMapParaLotes[serieOrd]) serieOrdenMapParaLotes[serieOrd] = true;

    // Solo órdenes VIVAS — cerradas van al botón Histórico
    if (estOrd === "CANCELADO" || estOrd === "TERMINADO" || estOrd === "SOBREPRODUCCION") continue;
    // No filtrar por estado del pedido padre: mostrar orden viva aunque el pedido esté cerrado

    if (!serieOrdenMap[serieOrd]) {
      serieOrdenMap[serieOrd] = { idPrimerProc: String(r[0]) };
    }

    // Primera máquina antes del primer coma
    var maqUnica   = String(r[12]).trim().split(",")[0].trim().toUpperCase();
    var grupoOrden = maqGrupoMap[maqUnica] || maqUnica;

    ordenes.push({
      id:           String(r[0]),
      pedido:       pedido,
      serie:        serie,
      orden:        numOrd,
      codigo:       String(r[6]).trim(),
      descripcion:  String(r[7]),
      cantidad:     Number(r[8]),
      unidad:       String(r[9]),
      sec:          Number(r[10]),
      proceso:      proc,
      maquina:      maqUnica,
      grupo:        grupoOrden,
      sol:          Number(r[13]) || 0,
      prod:         Number(r[14]) || 0,
      estado:       String(r[15]),
      tipo:         String(r[19]),
      diametro:     String(r[20]),
      longitud:     String(r[21]),
      cuerda:       String(r[22]),
      cuerpo:       String(r[23]),
      acero:        String(r[24]),
      loteIni:      r[34] || 1,
      loteFin:      r[35] || 1,
      prioridad:    Number(r[26]) || 999,
      idPrimerProc: String(r[0])
    });
  }

  ordenes.sort(function(a, b) { return a.prioridad - b.prioridad; });
  ordenes.forEach(function(o) {
    var so = o.serie + "." + ("0000" + o.orden).slice(-4);
    if (serieOrdenMap[so]) o.idPrimerProc = serieOrdenMap[so].idPrimerProc;
  });

  // D. Lotes agrupados por Serie.Orden
  // FIX: se usa serieOrdenMapParaLotes (todas las órdenes del proceso, no solo vivas)
  // para que el modo Históricas también muestre sus lotes correctamente.
  var lotes     = {};
  var configMap = {};
  for (var k = 1; k < dataLotes.length; k++) {
    var nombreLote   = String(dataLotes[k][4]).trim();
    var partes       = nombreLote.split(".");
    if (partes.length < 2) continue;
    var serieOrdLote = partes[0] + "." + partes[1];
    if (!serieOrdenMapParaLotes[serieOrdLote]) continue;
    if (!lotes[serieOrdLote]) lotes[serieOrdLote] = [];
    var consec = Number(dataLotes[k][3]) || 0;
    lotes[serieOrdLote].push({
      lote:   nombreLote,
      peso:   Number(dataLotes[k][5]) || 0,
      estado: String(dataLotes[k][7]),
      consec: consec
    });
    if (!configMap[serieOrdLote]) configMap[serieOrdLote] = { maxConsec: 0, pesoDefault: 0 };
    if (consec > configMap[serieOrdLote].maxConsec) configMap[serieOrdLote].maxConsec = consec;
  }

  // E. Peso default desde RUTAS (Col W idx22 = CANT_LOTE)
  var sheetRut = ss.getSheetByName("RUTAS");
  var dataRut  = sheetRut.getDataRange().getValues();
  var codigosUsados = {};
  ordenes.forEach(function(o) { codigosUsados[o.codigo] = true; });
  for (var rr = 1; rr < dataRut.length; rr++) {
    var codR  = String(dataRut[rr][1]).trim();
    var pesoD = Number(dataRut[rr][22]) || 0;
    if (!codigosUsados[codR] || pesoD <= 0) continue;
    ordenes.forEach(function(o) {
      if (o.codigo !== codR) return;
      var so = o.serie + "." + ("0000" + o.orden).slice(-4);
      if (configMap[so] && configMap[so].pesoDefault === 0) configMap[so].pesoDefault = pesoD;
    });
  }

  return { ordenes: ordenes, lotes: lotes, config: configMap, grupoMaqsMap: grupoMaqsMap };
}

function obtenerRutaYDocsOrden(codigo, procesoActual, idOrden) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetRut = ss.getSheetByName("RUTAS");
  var sheetOrd = ss.getSheetByName("ORDENES");
  var dataRut  = sheetRut.getDataRange().getValues();
  var dataOrd  = sheetOrd.getDataRange().getValues();

  // Construir ruta del código
  var ruta = {};
  for (var r = 1; r < dataRut.length; r++) {
    if (String(dataRut[r][1]).trim() !== codigo) continue;
    var sec = Number(dataRut[r][3]);
    var maqDisp = String(dataRut[r][5]).split(",").map(function(m){ return m.trim(); }).filter(Boolean);
    if (!ruta[sec]) {
      ruta[sec] = {
        sec: sec,
        proceso: String(dataRut[r][4]),
        maquinasDisponibles: maqDisp,
        maquinaActual: maqDisp[0] || ""
      };
    }
  }

  // Sobreescribir máquina actual con lo que tienen las ORDENES
  for (var i = 1; i < dataOrd.length; i++) {
    if (String(dataOrd[i][6]).trim() !== codigo) continue;
    var secOrd = Number(dataOrd[i][10]);
    var maqOrd = String(dataOrd[i][12]);
    if (ruta[secOrd] && maqOrd) ruta[secOrd].maquinaActual = maqOrd;
  }

  var rutaArr = Object.values(ruta).sort(function(a,b){ return a.sec - b.sec; });

  // Documentación técnica
  var docs = {};
  try {
    var d = obtenerDocsTecnicos(codigo);
    if (d && d.datos) docs = d.datos;
  } catch(e) {}

  return { ruta: rutaArr, docs: docs };
}

// 3. HISTÓRICAS: Pedidos de últimos 120 días, cualquier estado
function obtenerHistoricasTablero(nombreProceso) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var dataOrd = sheetOrd.getDataRange().getValues();
  
  var limite = new Date();
  limite.setDate(limite.getDate() - 120);
  var tz = Session.getScriptTimeZone();
  
  var ordenes = [];
  for(var i=1; i<dataOrd.length; i++) {
    var r = dataOrd[i];
    var proc = String(r[11]).toUpperCase().trim();
    if(proc !== nombreProceso.toUpperCase()) continue;
    
    var fecha = r[3];
    if(fecha instanceof Date && fecha < limite) continue;
    
    var serie = String(r[4]);
    var numOrd = Number(r[5]);
    
    ordenes.push({
      id: String(r[0]),
      pedido: String(r[1]).trim(),
      serie: serie,
      orden: numOrd,
      codigo: String(r[6]).trim(),
      descripcion: String(r[7]),
      cantidad: Number(r[8]),
      unidad: String(r[9]),
      sec: Number(r[10]),
      proceso: proc,
      maquina: String(r[12]),
      sol: Number(r[13]) || 0,
      prod: Number(r[14]) || 0,
      estado: String(r[15]),
      tipo: String(r[19]),
      diametro: String(r[20]),
      longitud: String(r[21]),
      cuerda: String(r[22]),
      cuerpo: String(r[23]),
      acero: String(r[24]),
      loteIni: r[34] || 1,
      loteFin: r[35] || 1,
      idPrimerProc: String(r[0])
    });
  }
  return ordenes;
}

// 4. GENERAR LOTES (Referencia por Serie.Orden, ID como texto)
function generarLotesTablero(payload) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetLotes = ss.getSheetByName("LOTES");
  var sheetOrd   = ss.getSheetByName("ORDENES");

  var serieOrden    = payload.serieOrden;
  var serie         = payload.serie;
  var ordenNum      = Number(payload.ordenNum);
  var consecInicial = Number(payload.consecInicial);
  var cantGenerar   = Number(payload.cantidad);
  var pesoLote      = Number(payload.peso);

  // FIX: buscar el primer proceso (menor SEC) de este serie+orden para guardar el ID
  // correcto en col C de LOTES. Una orden tiene N procesos; el ID correcto es el de SEC=1.
  var dataOrd = sheetOrd.getDataRange().getValues();
  var idOrdenRef = payload.idOrdenPrimerProc || "";
  var minSec = Infinity;
  for (var i = 1; i < dataOrd.length; i++) {
    if (String(dataOrd[i][4]).trim() === String(serie).trim()
        && Number(dataOrd[i][5]) === ordenNum) {
      var sec = Number(dataOrd[i][10]);
      if (sec < minSec) {
        minSec = sec;
        idOrdenRef = String(dataOrd[i][0]);
      }
    }
  }

  var nuevos = [];
  var lotesCreados = [];
  var fecha = new Date();

  for (var c = 1; c <= cantGenerar; c++) {
    var consec     = consecInicial + c;
    var numOrdStr  = ("0000" + ordenNum).slice(-4);
    var consecStr  = consec < 100 ? ("00" + consec).slice(-2) : String(consec);
    var nombreLote = serie + "." + numOrdStr + "." + consecStr;

    // FIX 3: UUID como texto — único garantizado, sin pérdida de precisión
    var idTexto = Utilities.getUuid();

    var fila = [idTexto, serie, idOrdenRef, consec, nombreLote, pesoLote, 0, "ABIERTO", fecha, "", "", "", "", "", "NADA"];
    nuevos.push(fila);
    lotesCreados.push({ lote: nombreLote, peso: pesoLote, estado: "ABIERTO", consec: consec });
  }

  if (nuevos.length > 0) {
    sheetLotes.getRange(sheetLotes.getLastRow() + 1, 1, nuevos.length, nuevos[0].length).setValues(nuevos);
  }

  return { nuevosLotes: lotesCreados };
}

// 5. CAMBIAR MÁQUINA DE UNA ORDEN
function cambiarMaquinaOrden(idOrden, secObjetivo, nuevaMaquina, procesoActual, codigo) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var dataOrd = sheetOrd.getDataRange().getValues();
  
  // Buscar la orden por proceso actual (mismo codigo, mismo procesoActual)
  for(var i=1; i<dataOrd.length; i++) {
    if(String(dataOrd[i][0]) === String(idOrden)) {
      sheetOrd.getRange(i+1, 13).setValue(nuevaMaquina); // Columna M = MAQUINA
      break;
    }
  }
  return true;
}

// 6. URL para ReporteProduccion
function obtenerUrlReporteProduccion(idOrden, ini, fin) {
  guardarRangoImpresion(idOrden, ini, fin);
  var url = "https://script.google.com/macros/s/AKfycbyUE_d2XeLxNMzx94J40FmpLnhCbvYumvMVewc5nQNPYO4ezSnZKNsAg55gtSau_uN9/exec";
  url += "?tipo=REPORTE_SOLO&idOrden=" + idOrden + "&inicio=" + ini + "&fin=" + fin;
  return url;
}

//////////***Estas no se si quitarlas ¿Eliminarlas? */
// 3. GENERAR LOTES
function generarNuevosLotes(payload) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetLotes = ss.getSheetByName("LOTES");
  var sheetOrd = ss.getSheetByName("ORDENES"); 
  
  var idOrden = payload.idOrden;
  var cantGenerar = Number(payload.cantidad);
  var pesoLote = Number(payload.peso);
  var serie = payload.serie;
  var ordenNum = payload.ordenNum;
  var consecInicial = Number(payload.consecInicial);

  var dataLotes = sheetLotes.getDataRange().getValues();
  var maxID = 0;
  for(var i=1; i<dataLotes.length; i++) {
     var val = Number(dataLotes[i][0]);
     if(val > maxID) maxID = val;
  }

  var nuevos = [];
  var fecha = new Date();
  
  for(var c=1; c<=cantGenerar; c++) {
      maxID++;
      var consec = consecInicial + c;
      var numOrdStr = ("0000" + ordenNum).slice(-4);
      var consecStr = consec < 100 ? ("00" + consec).slice(-2) : String(consec);
      var nombreLote = serie + "." + numOrdStr + "." + consecStr;
      
      var fila = [ maxID, serie, idOrden, consec, nombreLote, pesoLote, 0, "ABIERTO", fecha, "", "", "", "", "", "NADA" ];
      nuevos.push(fila);
  }

  if(nuevos.length > 0) sheetLotes.getRange(sheetLotes.getLastRow()+1, 1, nuevos.length, nuevos[0].length).setValues(nuevos);

  // Guardar temporales
  var dataOrd = sheetOrd.getDataRange().getValues();
  for(var o=1; o<dataOrd.length; o++) {
      if(String(dataOrd[o][0]) == String(idOrden)) {
          sheetOrd.getRange(o+1, 31).setValue(cantGenerar); // AE
          sheetOrd.getRange(o+1, 32).setValue(pesoLote);    // AF
          break;
      }
  }
  return "✅ Se generaron " + cantGenerar + " lotes.";
}

// 4. GUARDAR RANGO IMPRESION Y ACTUALIZAR ESTATUS A "IMPRESO"
function guardarRangoImpresion(idOrden, ini, fin) {
   var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
   var sheetOrd = ss.getSheetByName("ORDENES");
   var sheetLotes = ss.getSheetByName("LOTES");

   // A. GUARDAR RANGO EN ORDEN — también construir prefijo de lote para búsqueda por nombre
   var dataOrd = sheetOrd.getDataRange().getValues();
   var prefijoBusqueda = "";
   for(var o=1; o<dataOrd.length; o++) {
      if(String(dataOrd[o][0]) == String(idOrden)) {
          sheetOrd.getRange(o+1, 35).setValue(ini);
          sheetOrd.getRange(o+1, 36).setValue(fin);
          // FIX: construir prefijo "SERIE.ORDEN." para buscar lotes por nombre, no por col C
          var serieOrd = String(dataOrd[o][4]).trim();
          var numOrd   = ("0000" + Number(dataOrd[o][5])).slice(-4);
          prefijoBusqueda = serieOrd + "." + numOrd + ".";
          break;
      }
   }

   // B. ACTUALIZAR ESTADO DE LOTES (ABIERTO -> IMPRESO)
   // FIX: se busca por prefijo de nombre del lote (ej. "P.0024.") en lugar de col C (ID orden-proceso)
   // Esto garantiza que funcione sin importar desde qué proceso se generaron los lotes.
   if (!prefijoBusqueda) return true;
   var dataLotes = sheetLotes.getDataRange().getValues();
   for(var k=1; k<dataLotes.length; k++) {
       var nomLote = String(dataLotes[k][4]).trim();
       if (nomLote.indexOf(prefijoBusqueda) !== 0) continue; // no pertenece a esta orden
       var consec = Number(dataLotes[k][3]);
       var estadoActual = String(dataLotes[k][7]).toUpperCase();
       if(consec >= ini && consec <= fin) {
           // REGLA: Sólo cambiar si es ABIERTO
           if(estadoActual == "ABIERTO") {
               sheetLotes.getRange(k+1, 8).setValue("IMPRESO");
           }
       }
   }
   return true;
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA EditorProduccionHTML (MENU EDITOR PRODUCCION)
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

// ── 1. CATÁLOGOS ──────────────────────────────────────────────────────────────
function obtenerCatalogosProd() {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetMenu = ss.getSheetByName("MENU_OPERATIVO");
  var sheetEst  = ss.getSheetByName("ESTANDARES");
  var sheetOps  = ss.getSheetByName("OPERADORES");

  var dataMenu = sheetMenu.getDataRange().getValues();
  var procesos = [];
  for (var i = 1; i < dataMenu.length; i++) {
    if (String(dataMenu[i][1]).toUpperCase().trim() === "PROCESO")
      procesos.push({ nombre: String(dataMenu[i][0]).toUpperCase().trim(), img: dataMenu[i][2] });
  }

  var dataEst = sheetEst.getDataRange().getValues();
  var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iProcE  = hEst.indexOf("PROCESO"); if (iProcE < 0) iProcE = 2;
  var iMaqE   = -1;
  for (var k = 0; k < hEst.length; k++) {
    if (hEst[k].includes("MAQUINA") && !hEst[k].includes("FOTO")) { iMaqE = k; break; }
  }
  if (iMaqE < 0) iMaqE = 3;

  var maquinasMap = {};
  for (var k = 1; k < dataEst.length; k++) {
    var p = String(dataEst[k][iProcE]).toUpperCase().trim();
    var m = String(dataEst[k][iMaqE]).toUpperCase().trim();
    if (p && m) {
      if (!maquinasMap[p]) maquinasMap[p] = [];
      if (maquinasMap[p].indexOf(m) < 0) maquinasMap[p].push(m);
    }
  }
  for (var key in maquinasMap) maquinasMap[key].sort();

  var dataOps   = sheetOps.getDataRange().getValues();
  var operadores = [];
  for (var j = 1; j < dataOps.length; j++) {
    if (dataOps[j][0])
      operadores.push({ id: String(dataOps[j][0]), nombre: String(dataOps[j][1]) });
  }
  operadores.sort(function(a, b){ return a.nombre.localeCompare(b.nombre); });

  return { procesos: procesos, maquinas: maquinasMap, operadores: operadores };
}


// ── 2. BUSCAR PRODUCCIÓN — SOLO LEE PRODUCCION, NADA MÁS ─────────────────────
//
//  Lee únicamente la hoja PRODUCCION, acotada al rango de fechas exacto.
//  ORDENES y LOTES se cargan por separado solo cuando se necesitan.
//
function buscarProduccionCompleta(filtros) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");

  // Fechas de filtro como objetos Date
  var pi = filtros.fechaIni.split("-");
  var pf = filtros.fechaFin.split("-");
  var fechaIni = new Date(+pi[0], +pi[1]-1, +pi[2],  0,  0,  0);
  var fechaFin = new Date(+pf[0], +pf[1]-1, +pf[2], 23, 59, 59);

  // PASO 1: Leer solo la columna FECHA para ubicar el rango de filas
  var headerRow = sheetProd.getRange(1, 1, 1, sheetProd.getLastColumn()).getValues()[0];
  var hProd     = headerRow.map(function(h){ return String(h).toUpperCase().trim(); });
  var colFECHA  = hProd.indexOf("FECHA") + 1;
  if (colFECHA < 1) colFECHA = 6;

  var totalFilas = sheetProd.getLastRow();
  if (totalFilas < 2) return { registros: [], dicLotes: {} };

  var soloFechas = sheetProd.getRange(2, colFECHA, totalFilas - 1, 1).getValues();

  var filaInicioRel = -1, filaFinRel = -1;
  for (var f = 0; f < soloFechas.length; f++) {
    var fv = soloFechas[f][0];
    if (!(fv instanceof Date)) continue;
    if (fv >= fechaIni && fv <= fechaFin) {
      if (filaInicioRel === -1) filaInicioRel = f;
      filaFinRel = f;
    }
  }

  if (filaInicioRel === -1) return { registros: [], dicLotes: {} };

  // PASO 2: Leer SOLO las filas del rango encontrado
  var filaReal    = filaInicioRel + 2;
  var numFilas    = filaFinRel - filaInicioRel + 1;
  var dataProd    = sheetProd.getRange(filaReal, 1, numFilas, sheetProd.getLastColumn()).getValues();

  var IDX = {
    ID:       hProd.indexOf("ID"),
    LOTE:     hProd.indexOf("LOTE"),
    MAQ:      hProd.indexOf("MAQUINA"),
    FECHA:    hProd.indexOf("FECHA"),
    TURNO:    hProd.indexOf("TURNO"),
    OPER_ID:  hProd.indexOf("OPERADOR"),
    OPER_TXT: hProd.indexOf("NOMBRE_OPERADOR_TXT"),
    PESO_I:   hProd.indexOf("PESO_I"),
    PESO_F:   hProd.indexOf("PESO_F"),
    PESO_T:   hProd.indexOf("PESO_TINA"),
    PROD:     hProd.indexOf("PRODUCIDO"),
    HORAS:    hProd.indexOf("HORAS"),
    SELLO:    hProd.indexOf("SELLO"),
    COM:      hProd.indexOf("COMENTARIO"),
    ORDEN:    hProd.indexOf("ORDEN")
  };

  // PASO 3: Filtrar por máquina
  var setMaqs = {};
  filtros.maquinas.forEach(function(m){ setMaqs[m.toUpperCase()] = true; });

  var resultados = [];
  for (var i = 0; i < dataProd.length; i++) {
    var fVal = dataProd[i][IDX.FECHA];
    if (!(fVal instanceof Date) || fVal < fechaIni || fVal > fechaFin) continue;
    if (!setMaqs[String(dataProd[i][IDX.MAQ]).toUpperCase().trim()]) continue;

    var fd   = fVal;
    var fStr = fd.getFullYear()
      + "-" + (fd.getMonth()+1 < 10 ? "0" : "") + (fd.getMonth()+1)
      + "-" + (fd.getDate()    < 10 ? "0" : "") +  fd.getDate();

    resultados.push({
      id:          String(dataProd[i][IDX.ID]),
      lote:        String(dataProd[i][IDX.LOTE]),
      maquina:     String(dataProd[i][IDX.MAQ]).trim(),
      fecha:       fStr,
      turno:       dataProd[i][IDX.TURNO],
      operadorId:  dataProd[i][IDX.OPER_ID],
      operadorTxt: dataProd[i][IDX.OPER_TXT],
      pesoI:       Number(dataProd[i][IDX.PESO_I])  || 0,
      pesoF:       Number(dataProd[i][IDX.PESO_F])  || 0,
      pesoTina:    Number(dataProd[i][IDX.PESO_T])  || 0,
      producido:   Number(dataProd[i][IDX.PROD])    || 0,
      horas:       Number(dataProd[i][IDX.HORAS])   || 0,
      sello:       dataProd[i][IDX.SELLO],
      comentario:  dataProd[i][IDX.COM],
      ordenRef:    String(dataProd[i][IDX.ORDEN]),
      // Campos de orden: vacíos — se llenan con obtenerInfoOrden si se necesitan
      proceso: "", tipo: "", dia: "", long: "", cuerpo: "", cuerda: "", acero: ""
    });
  }

  // Ordenar: más reciente primero
  resultados.sort(function(a, b){
    return (b.fecha + String(b.turno)).localeCompare(a.fecha + String(a.turno));
  });

  // dicLotes vacío — se carga lazy por obtenerLotesDeOrden
  return { registros: resultados, dicLotes: {} };
}

// ── Historial FORJA: producción enriquecida con datos de ORDENES ──
function buscarHistorialForja(fechaIni, fechaFin) {
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shProd    = ss.getSheetByName('PRODUCCION');
    var shOrd     = ss.getSheetByName('ORDENES');
    if (!shProd) return { registros: [] };

    var pi = fechaIni.split('-'); var pf = fechaFin.split('-');
    var dIni = new Date(+pi[0],+pi[1]-1,+pi[2], 0, 0, 0);
    var dFin = new Date(+pf[0],+pf[1]-1,+pf[2],23,59,59);

    // ── Leer ORDENES una sola vez → mapa ORDEN → datos producto ──
    var mapaOrden = {};
    if (shOrd) {
      var dataOrd  = shOrd.getDataRange().getValues();
      var hOrd     = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var oORD  = hOrd.indexOf('ORDEN');
      var oTIPO = hOrd.indexOf('TIPO');
      var oDIA  = hOrd.indexOf('DIAMETRO');
      var oLONG = hOrd.indexOf('LONGITUD');
      var oCPO  = hOrd.indexOf('CUERPO');
      var oCDA  = hOrd.indexOf('CUERDA');
      var oACE  = hOrd.indexOf('ACERO');
      var oCOD  = hOrd.indexOf('CODIGO');
      for (var oi = 1; oi < dataOrd.length; oi++) {
        var ordKey = String(dataOrd[oi][oORD] || '').trim();
        if (!ordKey || mapaOrden[ordKey]) continue;
        mapaOrden[ordKey] = {
          tipo:    oTIPO>-1 ? String(dataOrd[oi][oTIPO]||'').trim() : '',
          dia:     oDIA >-1 ? String(dataOrd[oi][oDIA] ||'').trim() : '',
          long:    oLONG>-1 ? String(dataOrd[oi][oLONG]||'').trim() : '',
          cuerpo:  oCPO >-1 ? String(dataOrd[oi][oCPO] ||'').trim() : '',
          cuerda:  oCDA >-1 ? String(dataOrd[oi][oCDA] ||'').trim() : '',
          acero:   oACE >-1 ? String(dataOrd[oi][oACE] ||'').trim() : '',
          codigo:  oCOD >-1 ? String(dataOrd[oi][oCOD] ||'').trim() : ''
        };
      }
    }

    // ── Leer PRODUCCION ──
    var headerRow = shProd.getRange(1,1,1,shProd.getLastColumn()).getValues()[0];
    var hProd = headerRow.map(function(h){ return String(h).toUpperCase().trim(); });
    var IDX = {
      ID:      hProd.indexOf('ID'),
      LOTE:    hProd.indexOf('LOTE'),
      MAQ:     hProd.indexOf('MAQUINA'),
      FECHA:   hProd.indexOf('FECHA'),
      TURNO:   hProd.indexOf('TURNO'),
      OPER_TXT:hProd.indexOf('NOMBRE_OPERADOR_TXT'),
      PESO_I:  hProd.indexOf('PESO_I'),
      PESO_F:  hProd.indexOf('PESO_F'),
      PESO_T:  hProd.indexOf('PESO_TINA'),
      PROD:    hProd.indexOf('PRODUCIDO'),
      HORAS:   hProd.indexOf('HORAS'),
      SELLO:   hProd.indexOf('SELLO'),
      COM:     hProd.indexOf('COMENTARIO'),
      ORDEN:   hProd.indexOf('ORDEN')
    };
    var colFECHA = (IDX.FECHA >= 0 ? IDX.FECHA : 5);
    var totalFilas = shProd.getLastRow();
    if (totalFilas < 2) return { registros: [] };

    // Ubicar rango por fecha
    var soloFechas = shProd.getRange(2, colFECHA+1, totalFilas-1, 1).getValues();
    var fIni = -1, fFin = -1;
    for (var f = 0; f < soloFechas.length; f++) {
      var fv = soloFechas[f][0];
      if (!(fv instanceof Date)) continue;
      if (fv >= dIni && fv <= dFin) { if (fIni===-1) fIni=f; fFin=f; }
    }
    if (fIni === -1) return { registros: [] };

    var dataProd = shProd.getRange(fIni+2, 1, fFin-fIni+1, shProd.getLastColumn()).getValues();
    var tz = Session.getScriptTimeZone();
    var resultados = [];

    for (var i = 0; i < dataProd.length; i++) {
      var fVal = dataProd[i][IDX.FECHA];
      if (!(fVal instanceof Date) || fVal < dIni || fVal > dFin) continue;
      var fStr = Utilities.formatDate(fVal, tz, 'yyyy-MM-dd');
      var ordenRef = String(dataProd[i][IDX.ORDEN] || '').trim();
      var inf = mapaOrden[ordenRef] || { tipo:'', dia:'', long:'', cuerpo:'', cuerda:'', acero:'', codigo:'' };

      resultados.push({
        id:          String(dataProd[i][IDX.ID]),
        fecha:       fStr,
        turno:       dataProd[i][IDX.TURNO],
        maquina:     String(dataProd[i][IDX.MAQ] || '').trim(),
        lote:        String(dataProd[i][IDX.LOTE] || ''),
        tipo:        inf.tipo,
        diametro:    inf.dia,
        longitud:    inf.long,
        cuerpo:      inf.cuerpo,
        cuerda:      inf.cuerda,
        acero:       inf.acero,
        operadorTxt: String(dataProd[i][IDX.OPER_TXT] || ''),
        pesoI:       Number(dataProd[i][IDX.PESO_I])  || 0,
        pesoF:       Number(dataProd[i][IDX.PESO_F])  || 0,
        pesoTina:    Number(dataProd[i][IDX.PESO_T])  || 0,
        producido:   Number(dataProd[i][IDX.PROD])    || 0,
        horas:       Number(dataProd[i][IDX.HORAS])   || 0,
        sello:       String(dataProd[i][IDX.SELLO]    || ''),
        comentario:  String(dataProd[i][IDX.COM]      || ''),
        ordenRef:    ordenRef
      });
    }

    resultados.sort(function(a,b){
      return (b.fecha + String(b.turno)).localeCompare(a.fecha + String(a.turno));
    });
    return { registros: resultados };
  } catch(e) {
    return { registros: [], error: e.message };
  }
}

//3. Lee solo col ID para ubicar la fila, luego lee solo esa fila completa
function obtenerInfoOrden(idOrden) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOrd = ss.getSheetByName("ORDENES");
  var total = shOrd.getLastRow();
  if (total < 2) return {};

  // Paso 1: leer solo la columna ID (col A = 1) para encontrar la fila
  var soloIds = shOrd.getRange(2, 1, total - 1, 1).getValues();
  var filaEncontrada = -1;
  for (var i = 0; i < soloIds.length; i++) {
    if (String(soloIds[i][0]) === String(idOrden)) { filaEncontrada = i + 2; break; }
  }
  if (filaEncontrada === -1) return {};

  // Paso 2: leer solo esa fila + headers
  var headers = shOrd.getRange(1, 1, 1, shOrd.getLastColumn()).getValues()[0]
    .map(function(h){ return String(h).toUpperCase().trim(); });
  var fila = shOrd.getRange(filaEncontrada, 1, 1, shOrd.getLastColumn()).getValues()[0];

  var g = function(col){ var i = headers.indexOf(col); return i > -1 ? String(fila[i]) : ""; };
  return {
    tipo:   g("TIPO"),
    dia:    g("DIAMETRO"),
    long:   g("LONGITUD"),
    cuerpo: g("CUERPO"),
    cuerda: g("CUERDA")
  };
}

// ── 4. OBTENER LOTES DE UNA ORDEN (lazy, solo al abrir el combo) ─────────────
function obtenerLotesDeOrden(idOrden) {
  var ss         = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetLotes = ss.getSheetByName("LOTES");
  var totalFilas = sheetLotes.getLastRow();
  if (totalFilas < 2) return {};

  // Solo leer columna ID_ORDEN (C=3) para ubicar filas rápido
  var colIdOrden = 3;
  var colNomLote = 5;
  var soloIds    = sheetLotes.getRange(2, colIdOrden, totalFilas - 1, 1).getValues();

  var filasMatch = [];
  for (var i = 0; i < soloIds.length; i++) {
    if (String(soloIds[i][0]) === String(idOrden)) filasMatch.push(i + 2);
  }
  if (filasMatch.length === 0) return {};

  var lotes = [];
  filasMatch.forEach(function(fila) {
    var nom = String(sheetLotes.getRange(fila, colNomLote).getValue());
    if (nom && lotes.indexOf(nom) < 0) lotes.push(nom);
  });
  lotes.sort();

  var resultado = {};
  resultado[idOrden] = lotes;
  return resultado;
}

function obtenerLotesDeOrdenDetalle(idOrden) {
  // Busca lotes por Serie.Orden derivado del nombre del lote (igual que Tablero Supervisor)
  // idOrden puede ser el ID de cualquier orden del grupo — se busca su Serie.Orden en ORDENES
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd   = ss.getSheetByName('ORDENES');
    var shLotes = ss.getSheetByName('LOTES');
    if (!shLotes || shLotes.getLastRow() < 2) return [];

    // Obtener Serie y Orden de la orden para construir el prefijo P.XXXX
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(c){ return String(c).toUpperCase().trim(); });
    var iID    = hOrd.indexOf('ID');     if (iID    < 0) iID    = 0;
    var iSerie = hOrd.indexOf('SERIE');  if (iSerie < 0) iSerie = 7;
    var iOrden = hOrd.indexOf('ORDEN');  if (iOrden < 0) iOrden = 8;
    var serie = '', orden = '';
    for (var r = 1; r < dataOrd.length; r++) {
      if (String(dataOrd[r][iID]).trim() === String(idOrden).trim()) {
        serie = String(dataOrd[r][iSerie]||'').trim();
        orden = String(dataOrd[r][iOrden]||'').trim();
        break;
      }
    }
    if (!serie || !orden) return [];
    // Prefijo del lote: Serie.OrdenPadded — ej: P.0469
    var prefijo = serie + '.' + ('0000' + orden).slice(-4);

    // Leer lotes y filtrar por prefijo en nombre
    var data = shLotes.getDataRange().getValues();
    var h    = data[0].map(function(c){ return String(c).toUpperCase().trim(); });
    var iNom = h.indexOf('NOMBRE'); if (iNom < 0) iNom = 4;
    var iPes = h.indexOf('PESO');   if (iPes < 0) iPes = 5;
    var iEst = h.indexOf('ESTADO'); if (iEst < 0) iEst = 7;
    var lotes = [];
    for (var i = 1; i < data.length; i++) {
      var nom = String(data[i][iNom]||'').trim();
      if (!nom) continue;
      var partes = nom.split('.');
      if (partes.length < 2) continue;
      var serieOrd = partes[0] + '.' + partes[1];
      if (serieOrd !== prefijo) continue;
      lotes.push({
        lote:   nom,
        estado: String(data[i][iEst]||'ABIERTO').trim().toUpperCase(),
        peso:   Number(data[i][iPes])||0
      });
    }
    lotes.sort(function(a,b){ return a.lote.localeCompare(b.lote); });
    return lotes;
  } catch(e) {
    Logger.log('obtenerLotesDeOrdenDetalle ERROR: ' + e.message);
    return [];
  }
}

function obtenerMaquinasDeOrden(ordenRef) {
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd    = ss.getSheetByName("ORDENES");
    var shEst    = ss.getSheetByName("ESTANDARES");

    // 1. Buscar el proceso en ORDENES por ID (col A)
    var dataOrd = shOrd.getDataRange().getValues();
    var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iID     = hOrd.indexOf("ID");     if (iID  < 0) iID  = 0;
    var iProc   = hOrd.indexOf("PROCESO"); if (iProc < 0) iProc = 11;

    var proceso = "";
    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][iID]).trim() === String(ordenRef).trim()) {
        proceso = String(dataOrd[i][iProc]).toUpperCase().trim();
        break;
      }
    }
    if (!proceso) return JSON.stringify({ success: false, maquinas: [] });

    // 2. Buscar todas las máquinas de ese proceso en ESTANDARES (col C = PROCESO)
    var dataEst = shEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iProcE  = hEst.indexOf("PROCESO"); if (iProcE < 0) iProcE = 2;
    var iMaqE   = -1;
    for (var k = 0; k < hEst.length; k++) {
      if (hEst[k].includes("MAQUINA") && !hEst[k].includes("FOTO")) { iMaqE = k; break; }
    }
    if (iMaqE < 0) iMaqE = 3;

    var maquinas = [];
    for (var j = 1; j < dataEst.length; j++) {
      var p = String(dataEst[j][iProcE]).toUpperCase().trim();
      var m = String(dataEst[j][iMaqE]).toUpperCase().trim();
      if (p === proceso && m && maquinas.indexOf(m) < 0) maquinas.push(m);
    }
    maquinas.sort();

    return JSON.stringify({ success: true, proceso: proceso, maquinas: maquinas });
  } catch(e) {
    return JSON.stringify({ success: false, maquinas: [], error: e.message });
  }
}

// ── 5. GUARDAR CAMBIOS ────────────────────────────────────────────────────────
// NOTA: Si ya tienes guardarCambiosProduccion en tu script, NO la copies.
function guardarCambiosProduccion(cambios) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet    = ss.getSheetByName("PRODUCCION");
  var sheetOrd = ss.getSheetByName("ORDENES");

  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var getIdx  = function(n){ return headers.indexOf(n); };

  var IDX = {
    ID:       getIdx("ID"),       LOTE:     getIdx("LOTE"),
    MAQ:      getIdx("MAQUINA"),  FECHA:    getIdx("FECHA"),
    TURNO:    getIdx("TURNO"),    OPER_ID:  getIdx("OPERADOR"),
    OPER_TXT: getIdx("NOMBRE_OPERADOR_TXT"),
    PESO_I:   getIdx("PESO_I"),   PESO_F:   getIdx("PESO_F"),
    PESO_T:   getIdx("PESO_TINA"),PROD:     getIdx("PRODUCIDO"),
    HORAS:    getIdx("HORAS"),    SELLO:    getIdx("SELLO"),
    COM:      getIdx("COMENTARIO"), CONSEC:  16
  };

  var mapaFilas = {};
  for (var i = 1; i < data.length; i++)
    mapaFilas[String(data[i][IDX.ID])] = i + 1;

  var actualizaciones = [], borrados = [], ordenesAfectadas = {};
  cambios.forEach(function(row) {
    var fila = mapaFilas[row.id];
    if (!fila) return;
    if (row.borrar === true) borrados.push(fila);
    else                     actualizaciones.push({ fila: fila, data: row });
    if (row.ordenRef) ordenesAfectadas[row.ordenRef] = true;
  });

  actualizaciones.forEach(function(item) {
    var f = item.fila, r = item.data;
    if (r.lote !== undefined) {
      sheet.getRange(f, IDX.LOTE + 1).setValue(r.lote);
      var partes = String(r.lote).split(".");
      var num = parseInt(partes[partes.length - 1], 10);
      if (!isNaN(num)) sheet.getRange(f, IDX.CONSEC + 1).setValue(num);
    }
    if (r.maquina    !== undefined) sheet.getRange(f, IDX.MAQ      + 1).setValue(r.maquina);
    if (r.turno      !== undefined) sheet.getRange(f, IDX.TURNO    + 1).setValue(r.turno);
    if (r.operadorId !== undefined) sheet.getRange(f, IDX.OPER_ID  + 1).setValue(r.operadorId);
    if (r.operadorTxt!== undefined) sheet.getRange(f, IDX.OPER_TXT + 1).setValue(r.operadorTxt);
    if (r.pesoI      !== undefined) sheet.getRange(f, IDX.PESO_I   + 1).setValue(Number(r.pesoI)    || 0);
    if (r.pesoF      !== undefined) sheet.getRange(f, IDX.PESO_F   + 1).setValue(Number(r.pesoF)    || 0);
    if (r.pesoTina   !== undefined) sheet.getRange(f, IDX.PESO_T   + 1).setValue(Number(r.pesoTina) || 0);
    if (r.producido  !== undefined) sheet.getRange(f, IDX.PROD     + 1).setValue(Number(r.producido)|| 0);
    if (r.horas      !== undefined) sheet.getRange(f, IDX.HORAS    + 1).setValue(Number(r.horas)    || 0);
    if (r.sello      !== undefined) sheet.getRange(f, IDX.SELLO    + 1).setValue(r.sello);
    if (r.comentario !== undefined) sheet.getRange(f, IDX.COM      + 1).setValue(r.comentario);
    if (r.fecha !== undefined && r.fecha !== "") {
      var p = r.fecha.split("-");
      sheet.getRange(f, IDX.FECHA + 1).setValue(new Date(+p[0], +p[1]-1, +p[2], 12, 0, 0));
    }
  });

  borrados.sort(function(a, b){ return b - a; });
  borrados.forEach(function(f){ sheet.deleteRow(f); });
  SpreadsheetApp.flush();

  // Recalcular producido y estado de órdenes afectadas
  var dataOrdActual  = sheetOrd.getDataRange().getValues();
  var hOrd = dataOrdActual[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iID  = hOrd.indexOf("ID"), iPROD = hOrd.indexOf("PRODUCIDO");
  var iSOL = hOrd.indexOf("SOLICITADO"), iEST = hOrd.indexOf("ESTADO");

  var dataProdActual = sheet.getDataRange().getValues();
  var hP    = dataProdActual[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iPOrd = hP.indexOf("ORDEN"), iPProd = hP.indexOf("PRODUCIDO");

  var sumProd = {};
  for (var p = 1; p < dataProdActual.length; p++) {
    var idO = String(dataProdActual[p][iPOrd]);
    sumProd[idO] = (sumProd[idO] || 0) + (Number(dataProdActual[p][iPProd]) || 0);
  }
  for (var o = 1; o < dataOrdActual.length; o++) {
    var idOrden = String(dataOrdActual[o][iID]);
    if (!ordenesAfectadas[idOrden]) continue;
    var totalProd  = sumProd[idOrden] || 0;
    var solicitado = Number(dataOrdActual[o][iSOL]) || 0;
    sheetOrd.getRange(o + 1, iPROD + 1).setValue(totalProd);
    var est = String(dataOrdActual[o][iEST]).toUpperCase();
    if (est !== "CANCELADO") {
      sheetOrd.getRange(o + 1, iEST + 1).setValue(
        totalProd >= solicitado
          ? (totalProd > solicitado ? "SOBREPRODUCCION" : "TERMINADO")
          : (totalProd > 0 ? "EN PROCESO" : "ABIERTO")
      );
    }
  }
  SpreadsheetApp.flush();
  return "✅ Guardado: " + actualizaciones.length + " modificados, " + borrados.length + " eliminados.";
}


// ── 6. EXPORTAR EXCEL ─────────────────────────────────────────────────────────
function generarReporteKioscoExcel(filtros) {
  var ss     = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var tz     = ss.getSpreadsheetTimeZone();
  var shProd = ss.getSheetByName("PRODUCCION");
  var shOrd  = ss.getSheetByName("ORDENES");
  var shOps  = ss.getSheetByName("OPERADORES");

  var dOps = shOps.getDataRange().getValues();
  var mapaNomina = {};
  for (var i = 1; i < dOps.length; i++) mapaNomina[String(dOps[i][0])] = dOps[i][6];

  var dOrd = shOrd.getDataRange().getValues();
  var hOrd = dOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var mapaOrd = {};
  for (var j = 1; j < dOrd.length; j++) mapaOrd[String(dOrd[j][0])] = dOrd[j];

  var pi = filtros.fechaIni.split("-"), pf = filtros.fechaFin.split("-");
  var fechaIni = new Date(+pi[0], +pi[1]-1, +pi[2],  0,  0,  0);
  var fechaFin = new Date(+pf[0], +pf[1]-1, +pf[2], 23, 59, 59);
  var setProcs = {};
  filtros.procesos.forEach(function(p){ setProcs[p.toUpperCase()] = true; });

  // Mismo truco: acotar por fecha primero
  var headerRow = shProd.getRange(1, 1, 1, shProd.getLastColumn()).getValues()[0];
  var hProd     = headerRow.map(function(h){ return String(h).toUpperCase().trim(); });
  var colFECHA  = hProd.indexOf("FECHA") + 1; if (colFECHA < 1) colFECHA = 6;
  var totalFilas = shProd.getLastRow();
  var soloFechas = shProd.getRange(2, colFECHA, totalFilas - 1, 1).getValues();

  var filaIni = -1, filaFin = -1;
  for (var f = 0; f < soloFechas.length; f++) {
    var fv = soloFechas[f][0];
    if (!(fv instanceof Date)) continue;
    if (fv >= fechaIni && fv <= fechaFin) {
      if (filaIni === -1) filaIni = f;
      filaFin = f;
    }
  }
  if (filaIni === -1) return [["Sin datos para el rango seleccionado"]];

  var dProd = shProd.getRange(filaIni + 2, 1, filaFin - filaIni + 1, shProd.getLastColumn()).getValues();
  var colP  = {
    FECHA: hProd.indexOf("FECHA"), OPER_ID: hProd.indexOf("OPERADOR"),
    OPER_TXT: hProd.indexOf("NOMBRE_OPERADOR_TXT"), TURNO: hProd.indexOf("TURNO"),
    TINA: hProd.indexOf("SELECTOR_NUMERO_LOTE"), SELLO: hProd.indexOf("SELLO"),
    HRS: hProd.indexOf("HORAS"), MAQ: hProd.indexOf("MAQUINA"),
    PROD: hProd.indexOf("PRODUCIDO"), ORDEN: hProd.indexOf("ORDEN")
  };

  var csvData = [["DIA","AÑO","FECHA","N_NOMINA","OPERADOR","SUPERVISOR","TURNO",
                  "OP","TINA","SELLO","HRS","MAQUINA","PEDIDO","CODIGO","TIPO",
                  "DIAMETRO","LONGITUD","CUERDA","CUERPO","ACERO","PESO","PRODUCIDO","PZA"]];

  for (var k = 0; k < dProd.length; k++) {
    var fVal = dProd[k][colP.FECHA];
    if (!(fVal instanceof Date) || fVal < fechaIni || fVal > fechaFin) continue;
    var rowO = mapaOrd[String(dProd[k][colP.ORDEN])];
    if (!rowO) continue;
    var proc = String(rowO[hOrd.indexOf("PROCESO")] || "").toUpperCase().trim();
    if (Object.keys(setProcs).length > 0 && !setProcs[proc]) continue;
    var peso = parseFloat(rowO[18]) || 0;
    var prod = parseFloat(dProd[k][colP.PROD]) || 0;
    csvData.push([
      ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"][fVal.getDay()],
      fVal.getFullYear() + " " + Utilities.formatDate(fVal, tz, "w"),
      Utilities.formatDate(fVal, tz, "dd/MM/yyyy"),
      Number(mapaNomina[String(dProd[k][colP.OPER_ID])]) || 0,
      dProd[k][colP.OPER_TXT], "", dProd[k][colP.TURNO],
      String(rowO[4]) + "." + ("0000" + Number(rowO[5])).slice(-4) + (proc === "FORJA" ? ".1" : proc === "PUNTEADO" ? ".25" : proc === "ROLADO TORN" ? ".3" : proc === "LAVADO" ? ".4" : ""),
      Math.floor(Number(dProd[k][colP.TINA]) || 0),
      dProd[k][colP.SELLO], dProd[k][colP.HRS], dProd[k][colP.MAQ],
      rowO[1], rowO[6], rowO[19],
      '="' + String(rowO[20]) + '"', '="' + String(rowO[21]) + '"',
      rowO[22], rowO[23], rowO[24],
      peso, prod, peso > 0 ? prod / peso : 0
    ]);
  }
  return csvData;
}

// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
// FUNCIONES GS NECESARIAS PARA SECCION DE INGENIERIA
// * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +- SECCION NUEVA RUTA +-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// 1. NUEVA RUTA: CARGAR CATALOGOS (MAQUINAS + FAMILIAS)
function obtenerCatalogosParaNuevaRuta() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  
  // A. OBTENER MAQUINAS (ESTANDARES)
  var sheetEst = ss.getSheetByName("ESTANDARES");
  var dataEst = sheetEst.getDataRange().getValues();
  
  var catalogo = {}; 
  var headers = dataEst[0].map(function(h){return String(h).toUpperCase()});
  var idxProc = headers.indexOf("PROCESO"); if(idxProc < 0) idxProc = 2;
  var idxMaq = headers.indexOf("MAQUINA"); if(idxMaq < 0) idxMaq = 3;

  for(var k=1; k<dataEst.length; k++) {
     var p = String(dataEst[k][idxProc]).toUpperCase().trim();
     var m = String(dataEst[k][idxMaq]).trim();
     
     if(p && m) {
       if(!catalogo[p]) catalogo[p] = [];
       if(catalogo[p].indexOf(m) === -1) catalogo[p].push(m);
     }
  }

  // B. OBTENER FAMILIAS UNICAS (RUTAS - COLUMNA X)
  var sheetRutas = ss.getSheetByName("RUTAS");
  var dataRutas = sheetRutas.getDataRange().getValues();
  var familiasSet = new Set();
  
  // Columna X es el índice 23 (Base 0: A=0... X=23)
  for(var r=1; r<dataRutas.length; r++) {
     var f = String(dataRutas[r][23]).trim(); 
     if(f) familiasSet.add(f);
  }
  
  // Convertir Set a Array ordenado
  var listaFamilias = Array.from(familiasSet).sort();

  return { maquinas: catalogo, familias: listaFamilias };
}

// 2. NUEVA RUTA: BUSCAR SI YA EXISTE (INCLUYE FAMILIA)
function buscarRutaPorCodigo(codigoInput) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetRutas = ss.getSheetByName("RUTAS");
  var dataRutas = sheetRutas.getDataRange().getValues();
  
  var codigoTarget = String(codigoInput).trim().toUpperCase();
  var rutaEncontrada = [];
  var existe = false;
  
  for (var i = 1; i < dataRutas.length; i++) {
     if (String(dataRutas[i][1]).trim().toUpperCase() === codigoTarget) { // Col B
        existe = true;
        var r = dataRutas[i];
        rutaEncontrada.push({
           DESCRIPCION: r[2], SEC: r[3], PROCESO: r[4], MAQUINA: r[5],
           TIPO: r[6], DIAMETRO: r[7], LONGITUD: r[8], CUERDA: r[9], CUERPO: r[10], ACERO: r[11],
           PT: r[12], VENTA: r[13], SERIE: r[14], PESO: r[15], UNIDAD: r[16],
           MP: r[17], ESP_1: r[18], ESP_2: r[19], ESP_3: r[20],
           CANT_LOTE: r[22], 
           FAMILIA: r[23], // Col X (Indice 23)
           GENERA_INT: r[24]
        });
     }
  }
  
  rutaEncontrada.sort(function(a,b) { return Number(a.SEC) - Number(b.SEC); });

  return { existe: existe, ruta: rutaEncontrada };
}

// 3. GUARDAR NUEVA RUTA MAESTRA (SOPORTE COL X Y Z)
function guardarNuevaRutaCompleta(payload) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetRutas = ss.getSheetByName("RUTAS");
  
  var codigoObjetivo = String(payload.codigo).trim().toUpperCase();
  var nuevaRuta = payload.ruta; 
  
  // 1. LIMPIEZA PREVIA (Borrar filas existentes de este código)
  var dataR = sheetRutas.getDataRange().getValues();
  for (var i = dataR.length - 1; i >= 1; i--) {
    if (String(dataR[i][1]).trim().toUpperCase() == codigoObjetivo) { 
      sheetRutas.deleteRow(i + 1);
    }
  }
  
  // 2. GENERAR ID NUEVO
  var nextIdRutas = obtenerSiguienteIdNumerico(sheetRutas, 0); 

  // 3. MAPEO EXACTO DE COLUMNAS (A=0 ... Z=25)
  var filasAInsertar = nuevaRuta.map(function(step) {
     // Aumentamos el tamaño del array a 26 para llegar hasta la columna Z
     var row = new Array(26).fill(""); 
     
     row[0] = nextIdRutas++;        // A: ID
     row[1] = codigoObjetivo;       // B: CODIGO
     row[2] = step.DESCRIPCION;     // C: DESCRIPCION
     row[3] = step.SEC;             // D: SEC
     row[4] = step.PROCESO;         // E: PROCESO
     row[5] = step.MAQUINA;         // F: MAQUINA
     row[6] = step.TIPO;            // G: TIPO
     row[7] = step.DIAMETRO;        // H: DIAMETRO
     row[8] = step.LONGITUD;        // I: LONGITUD
     row[9] = step.CUERDA;          // J: CUERDA
     row[10] = step.CUERPO;         // K: CUERPO
     row[11] = step.ACERO;          // L: ACERO
     row[12] = step.PT;             // M: PT
     row[13] = step.VENTA;          // N: VENTA
     
     row[14] = step.SERIE;          // O: SERIE
     row[15] = step.PESO;           // P: PESO
     row[16] = step.UNIDAD;         // Q: UNIDAD
     row[17] = step.MP;             // R: MP
     row[18] = step.ESP_1;          // S: ESP_1
     row[19] = step.ESP_2;          // T: ESP_2
     row[20] = step.ESP_3;          // U: ESP_3
     
     // V (21) vacío
     
     row[22] = step.CANT_LOTE;      // W: CANT_LOTE
     
     // --- NUEVOS CAMPOS ---
     row[23] = step.FAMILIA;        // X: FAMILIA (Indice 23)
     row[24] = step.GENERA_INT;     // Y: GENERA_INT (Indice 24)
     row[25] = codigoObjetivo;      // Z: CODIGO_VENTA (Indice 25 - Igual al código)
     
     return row;
  });
  
  if(filasAInsertar.length > 0) {
     sheetRutas.getRange(sheetRutas.getLastRow()+1, 1, filasAInsertar.length, filasAInsertar[0].length)
               .setValues(filasAInsertar);
  }
  
  return "✅ Ruta Maestra guardada/actualizada correctamente.";
}

// EN TU ARCHIVO Code.gs
function obtenerUltimoIdRuta() {
  // Simplemente devuelve un timestamp o busca el maximo real si tienes logica
  // Aquí devolvemos un aleatorio basado en tiempo para garantizar unicidad rapida
  return Math.floor(Math.random() * 9000) + 1000;
}

// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +- SECCION EDITOR RUTAS +-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

// A. OBTENER DATOS (LECTURA MEJORADA PARA MAQUINAS)
function obtenerDatosEditor(idOrden) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd   = ss.getSheetByName("ORDENES");
  var sheetEst   = ss.getSheetByName("ESTANDARES");
  var sheetRutas = ss.getSheetByName("RUTAS");

  var dataOrd    = sheetOrd.getDataRange().getValues();
  var headersOrd = dataOrd[0];

  var getIdx = function(name) {
    var clean = name.toUpperCase().trim();
    for (var k = 0; k < headersOrd.length; k++) {
      if (String(headersOrd[k]).toUpperCase().trim() === clean) return k;
    }
    for (var k = 0; k < headersOrd.length; k++) {
      if (String(headersOrd[k]).toUpperCase().trim().indexOf(clean) > -1) return k;
    }
    return -1;
  };

  var idxID     = getIdx("ID");
  var idxSerie  = getIdx("SERIE");
  // Tu hoja usa "ORDEN" para el consecutivo (Col F)
  var idxConsec = getIdx("ORDEN");
  if (idxConsec == -1) idxConsec = getIdx("CONSECUTIVO");
  var idxCodigo = getIdx("CODIGO");

  Logger.log("idOrden recibido: " + JSON.stringify(idOrden) + " tipo: " + typeof idOrden);
  Logger.log("Indices → ID:" + idxID + " SERIE:" + idxSerie + " ORDEN:" + idxConsec + " CODIGO:" + idxCodigo);

  var targetSerie  = "";
  var targetConsec = "";
  var targetCodigo = "";

  // Detectar modo: objeto = búsqueda manual, string = ID desde Planificador
  var busqSerie  = "";
  var busqConsec = "";
  var modoObjeto = false;

  if (idOrden !== null && idOrden !== undefined && typeof idOrden === "object") {
    modoObjeto = true;
    busqSerie  = String(idOrden.serie  || "").trim().toUpperCase();
    busqConsec = String(idOrden.consecutivo || "").trim();
  } else if (typeof idOrden === "string" && idOrden.indexOf("{") === 0) {
    try {
      var parsed = JSON.parse(idOrden);
      modoObjeto = true;
      busqSerie  = String(parsed.serie  || "").trim().toUpperCase();
      busqConsec = String(parsed.consecutivo || "").trim();
    } catch(e) {}
  }

  Logger.log("modoObjeto=" + modoObjeto + " busqSerie='" + busqSerie + "' busqConsec='" + busqConsec + "'");

  if (modoObjeto) {
    var busqConsecNum = parseFloat(busqConsec);

    for (var i = 1; i < dataOrd.length; i++) {
      if (!dataOrd[i][idxSerie] && !dataOrd[i][idxConsec]) continue;

      var filaSerie     = String(dataOrd[i][idxSerie]).trim().toUpperCase();
      var filaConsecRaw = dataOrd[i][idxConsec];
      var filaConsecNum = parseFloat(filaConsecRaw);
      var filaConsecStr = String(filaConsecRaw).trim();

      var coincideSerie  = (busqSerie === "" || filaSerie === busqSerie);
      var coincideConsec = (!isNaN(busqConsecNum) && !isNaN(filaConsecNum) && busqConsecNum === filaConsecNum)
                        || (filaConsecStr === busqConsec);

      if (coincideSerie && coincideConsec) {
        targetSerie  = String(dataOrd[i][idxSerie]).trim();
        targetConsec = filaConsecStr;
        targetCodigo = dataOrd[i][idxCodigo];
        Logger.log("Encontrado fila " + i + ": " + targetSerie + "." + targetConsec + " codigo=" + targetCodigo);
        break;
      }
    }

  } else {
    var idBuscar = String(idOrden).trim();
    for (var i = 1; i < dataOrd.length; i++) {
      if (String(dataOrd[i][idxID]).trim() == idBuscar) {
        targetSerie  = String(dataOrd[i][idxSerie]).trim();
        targetConsec = String(dataOrd[i][idxConsec]).trim();
        targetCodigo = dataOrd[i][idxCodigo];
        break;
      }
    }
  }

  if (!targetCodigo) {
    // Log diagnóstico: muestra las primeras 5 filas reales
    for (var d = 1; d <= Math.min(5, dataOrd.length - 1); d++) {
      Logger.log("Muestra fila " + d + ": SERIE='" + dataOrd[d][idxSerie] + "' ORDEN='" + dataOrd[d][idxConsec] + "' CODIGO='" + dataOrd[d][idxCodigo] + "'");
    }
    return { error: "Orden no encontrada. Verifica el formato SERIE.CONSECUTIVO (ej: F.8848)" };
  }

  var nombreFormateado = targetSerie + "." + targetConsec;

  // Leer máquinas de RUTAS
  var mapaMaquinasRutas = {};
  var dataRutas = sheetRutas.getDataRange().getValues();
  for (var r = 1; r < dataRutas.length; r++) {
    if (String(dataRutas[r][1]) == String(targetCodigo)) {
      mapaMaquinasRutas[String(dataRutas[r][3])] = dataRutas[r][5];
    }
  }

  // Obtener ruta de la orden
  var rutaActual = [];
  var cols = ["ID","SEC","PROCESO","MAQUINA","DESCRIPCION","PT","VENTA",
              "PESO","TIPO","DIAMETRO","LONGITUD","CUERDA","CUERPO","ACERO","CANTIDAD","UNIDAD"];
  var indices = {};
  cols.forEach(function(c) { indices[c] = getIdx(c); });

  for (var i = 1; i < dataOrd.length; i++) {
    var fs = String(dataOrd[i][idxSerie]).trim();
    var fc = String(dataOrd[i][idxConsec]).trim();
    if (fs === targetSerie && fc === targetConsec) {
      var filaObj = {};
      cols.forEach(function(c) { filaObj[c] = (indices[c] > -1) ? dataOrd[i][indices[c]] : ""; });
      if (mapaMaquinasRutas[String(filaObj["SEC"])]) {
        filaObj["MAQUINA"] = mapaMaquinasRutas[String(filaObj["SEC"])];
      }
      rutaActual.push(filaObj);
    }
  }
  rutaActual.sort(function(a,b){ return Number(a.SEC) - Number(b.SEC); });

  // Catálogo ESTANDARES
  var dataEst = sheetEst.getDataRange().getValues();
  var headersEst = dataEst[0];
  var catalogo = {};
  var idxProcE = -1, idxMaqE = -1;
  for (var h = 0; h < headersEst.length; h++) {
    var dh = String(headersEst[h]).toUpperCase().trim();
    if (dh === "PROCESO") idxProcE = h;
    if (dh === "MAQUINA") idxMaqE = h;
  }
  if (idxProcE == -1) idxProcE = 2;
  if (idxMaqE  == -1) idxMaqE  = 3;
  for (var k = 1; k < dataEst.length; k++) {
    var p = String(dataEst[k][idxProcE]).toUpperCase().trim();
    var m = String(dataEst[k][idxMaqE]).trim();
    if (p && m) {
      if (!catalogo[p]) catalogo[p] = [];
      if (catalogo[p].indexOf(m) === -1) catalogo[p].push(m);
    }
  }

  return {
    codigo: targetCodigo, nombreDisplay: nombreFormateado,
    serie: targetSerie, consecutivo: targetConsec,
    ruta: rutaActual, catalogo: catalogo
  };
}

// ===============
// B. GUARDAR CAMBIOS RUTA (V12 - SOPORTE COLUMNAS X, Y, Z + ACTUALIZACIÓN ORDENES)
// ===============
function guardarCambiosRuta(payload) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetRutas = ss.getSheetByName("RUTAS");
  
  var codigoObjetivo = String(payload.codigo).trim().toUpperCase();
  var nuevaRuta = payload.ruta; 
  var serieObjetivo = payload.serie || "GEN"; 
  
  // --- 1. RESCATAR DATOS EXISTENTES (PRESERVACIÓN) ---
  var dataR = sheetRutas.getDataRange().getValues();
  
  // Variables para guardar lo que ya existe en BD si la edición viene de una vista parcial
  var datosPreservados = {
      cantLote: 0,      // Col W (22)
      familia: "",      // Col X (23)
      generaInt: "",    // Col Y (24)
      codVenta: ""      // Col Z (25)
  };

  for (var i = 1; i < dataR.length; i++) {
    if (String(dataR[i][1]).trim().toUpperCase() == codigoObjetivo) {
       // Guardamos los datos de la primera coincidencia
       datosPreservados.cantLote = dataR[i][22] || 0;
       datosPreservados.familia = dataR[i][23] || "";
       datosPreservados.generaInt = dataR[i][24] || "";
       datosPreservados.codVenta = dataR[i][25] || "";
       break;
    }
  }

  // --- 2. BORRAR FILAS VIEJAS EN RUTAS ---
  for (var i = dataR.length - 1; i >= 1; i--) {
    if (String(dataR[i][1]).trim().toUpperCase() == codigoObjetivo) { 
      sheetRutas.deleteRow(i + 1);
    }
  }
  
  var nextIdRutas = obtenerSiguienteIdNumerico(sheetRutas, 0);

  // --- 3. INSERTAR NUEVAS FILAS (CON LOGICA DE FUSIÓN) ---
  var nuevasFilasRutas = nuevaRuta.map(function(step) {
     // Aumentamos el array a 26 espacios (0 a 25) para llegar a la Columna Z
     var row = new Array(26).fill(""); 

     row[0] = nextIdRutas++; 
     row[1] = codigoObjetivo;
     row[2] = step.DESCRIPCION;
     row[3] = step.SEC;
     row[4] = step.PROCESO;
     row[5] = step.MAQUINA;
     row[6] = step.TIPO;
     row[7] = step.DIAMETRO;
     row[8] = step.LONGITUD;
     row[9] = step.CUERDA;
     row[10] = step.CUERPO;
     row[11] = step.ACERO;
     row[12] = step.PT;
     row[13] = step.VENTA;
     row[14] = serieObjetivo;
     row[15] = step.PESO;
     row[16] = step.UNIDAD || "PZA";
     row[17] = step.MP || "";           // Col R
     row[18] = step.ESP_1 || "";        // Col S
     row[19] = step.ESP_2 || "";        // Col T
     row[20] = step.ESP_3 || "";        // Col U
     
     // Col V (21) vacío

     // PRIORIDAD: Dato del Payload (Nuevo) > Dato Preservado (Viejo)
     row[22] = (step.CANT_LOTE !== undefined) ? step.CANT_LOTE : datosPreservados.cantLote;
     row[23] = (step.FAMILIA !== undefined) ? step.FAMILIA : datosPreservados.familia;
     row[24] = (step.GENERA_INT !== undefined) ? step.GENERA_INT : datosPreservados.generaInt;
     
     // Si no hay código de venta definido, usamos el código de producción
     var cv = (step.CODIGO_VENTA !== undefined) ? step.CODIGO_VENTA : datosPreservados.codVenta;
     row[25] = cv || codigoObjetivo;

     return row;
  });
  
  if(nuevasFilasRutas.length > 0) {
     sheetRutas.getRange(sheetRutas.getLastRow()+1, 1, nuevasFilasRutas.length, nuevasFilasRutas[0].length)
               .setValues(nuevasFilasRutas);
  }
  
  // --- 4. ACTUALIZAR ORDENES (CORREGIDO) ---
  var fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() - 60); 
  var timeLimite = fechaLimite.getTime();
  
  var dataOrd = sheetOrd.getDataRange().getValues();
  var headersOrd = dataOrd[0];
  
  var getIdx = function(name) { 
    var clean = name.toUpperCase().trim();
    for(var k=0; k<headersOrd.length; k++) {
       var h = String(headersOrd[k]).toUpperCase().trim();
       if(h === clean) return k;
       if(h.indexOf(clean) === 0 && clean.length > 3) return k; // Match parcial al inicio
    }
    return -1;
  };

  // ÍNDICES QUIRÚRGICOS
  var IDX = {
     ID: getIdx("ID"), 
     ESTADO: getIdx("ESTADO"), 
     FECHA: getIdx("FECHA"), // Busca FECHA_REGISTRO o FECHA
     CODIGO: getIdx("CODIGO"), 
     ORDEN: getIdx("ORDEN"), 
     SEC: getIdx("SEC"),
     PROCESO: getIdx("PROCESO"), // <--- ESTE FALTABA Y CAUSABA EL ERROR
     CANT: getIdx("CANTIDAD"), 
     UNI: getIdx("UNIDAD"), 
     SOL: getIdx("SOLICITADO"),
     PESO: getIdx("PESO"), 
     PROD: getIdx("PRODUCIDO"), 
     SERIE: getIdx("SERIE"),       
     CONSEC: getIdx("CONSECUTIVO") 
  };
  
  // Validaciones críticas para evitar el error de columna 0 o NaN
  if (IDX.PROCESO === -1) IDX.PROCESO = 11; // Fallback Col L si no encuentra encabezado
  if (IDX.ORDEN === -1) IDX.ORDEN = 5;      // Fallback Col F
  if (IDX.SEC === -1) IDX.SEC = 3;         // Fallback Col D 

  var ordenesMap = {}; 
  var listaNombresAfectados = []; 

  for(var i=1; i<dataOrd.length; i++) {
     var cod = String(dataOrd[i][IDX.CODIGO]);
     var fechaRaw = dataOrd[i][IDX.FECHA];
     
     // 1. Validar Código
     if (cod != String(codigoObjetivo)) continue;

     // 2. VALIDAR FECHA ESTRICTA (60 DÍAS)
     var fechaFila = null;
     if (fechaRaw instanceof Date) {
        fechaFila = fechaRaw;
     } else if (typeof fechaRaw === 'string' && fechaRaw.length > 5) {
        var partes = fechaRaw.split(' ')[0].split('/'); 
        if (partes.length === 3) fechaFila = new Date(partes[2], partes[1]-1, partes[0]);
     }

     if (!fechaFila || fechaFila.getTime() < timeLimite) continue;

     var nomOrden = "";
     var s = (IDX.SERIE > -1) ? dataOrd[i][IDX.SERIE] : "";
     var c = (IDX.CONSEC > -1) ? dataOrd[i][IDX.CONSEC] : "";
     if(s && c) nomOrden = s + "." + ("0000" + c).slice(-4);
     else nomOrden = dataOrd[i][IDX.ORDEN]; 

     if(!ordenesMap[nomOrden]) {
        ordenesMap[nomOrden] = [];
        listaNombresAfectados.push(nomOrden);
     }
     ordenesMap[nomOrden].push(i + 1); 
  }
  
  // --- 4.1 PROCESAR FAMILIAS DE ORDENES (LÓGICA QUIRÚRGICA PARA PRESERVAR IDS) ---
  for (var nombreOrd in ordenesMap) {
     var filasExistentesNums = ordenesMap[nombreOrd];
     
     // Mapeamos lo que hay actualmente en la hoja para esta orden
     var pasosEnHoja = filasExistentesNums.map(function(filaNum) {
        return {
           fila: filaNum,
           id: sheetOrd.getRange(filaNum, IDX.ID + 1).getValue(),
           proceso: String(sheetOrd.getRange(filaNum, IDX.PROCESO + 1).getValue()).trim().toUpperCase(),
           utilizado: false 
        };
     });

     // Tomamos una fila de referencia para los datos generales (Cliente, Cantidad, etc.)
     var filaReferencia = pasosEnHoja[0].fila;
     var datosCabecera = sheetOrd.getRange(filaReferencia, 1, 1, headersOrd.length).getValues()[0];

     // Comparamos la NUEVA ruta contra lo que ya existe
     nuevaRuta.forEach(function(stepNuevo) {
        var nombreBusqueda = String(stepNuevo.PROCESO).trim().toUpperCase();
        
        // BUSCAMOS EL PROCESO POR NOMBRE:
        // ¿Ya existía este proceso en esta orden?
        var match = pasosEnHoja.find(function(p) {
           return !p.utilizado && p.proceso === nombreBusqueda;
        });

        if (match) {
           // SI EXISTE: Actualizamos esa misma fila (mantiene su ID original)
           actualizarFilaOrden(sheetOrd, match.fila, stepNuevo, headersOrd, datosCabecera, IDX);
           match.utilizado = true;
        } else {
           // NO EXISTE: Es un proceso nuevo (ej. ROLADO), insertamos fila con ID nuevo
           insertarFilaOrden(sheetOrd, datosCabecera, stepNuevo, headersOrd, IDX);
        }
     });

     // ELIMINACIÓN: Si en la hoja había pasos que ya no están en la nueva ruta, los borramos
     // Se hace en reversa para no arruinar los índices de fila al borrar
     pasosEnHoja.reverse().forEach(function(p) {
        if (!p.utilizado) {
           sheetOrd.deleteRow(p.fila);
        }
     });
  }
  
  // ORDENAMIENTO FINAL: Para que visualmente se vea bien en la hoja
  if (IDX.ORDEN > -1 && sheetOrd.getLastRow() > 1) {
      sheetOrd.getRange(2, 1, sheetOrd.getLastRow() - 1, sheetOrd.getLastColumn())
              .sort([
                {column: IDX.ORDEN + 1, ascending: true},
                {column: IDX.SEC + 1, ascending: true}
              ]);
  }

  if (listaNombresAfectados.length > 0) {
     return "✅ Ingeniería actualizada.\nÓrdenes afectadas (" + listaNombresAfectados.length + "): " + listaNombresAfectados.join(", ");
  } else {
     return "✅ Ruta maestra actualizada (Datos X, Y, Z preservados).\n⚠️ No se encontraron órdenes vivas recientes.";
  }
}

// Helper: Actualiza fila con protección de columna
function actualizarFilaOrden(sheet, fila, datos, headers, datosFilaCompleta, IDX) {
   var safeSet = function(idx, val) {
      if(idx !== undefined && idx > -1) sheet.getRange(fila, idx + 1).setValue(val);
   };

   safeSet(IDX.SEC, datos.SEC);
   safeSet(IDX.PROCESO, datos.PROCESO);
   
   // Máquina limpia (Solo la primera)
   if(IDX.MAQUINA === undefined) IDX.MAQUINA = headers.indexOf("MAQUINA");
   var maqLimpia = (datos.MAQUINA) ? String(datos.MAQUINA).split(",")[0].trim() : "";
   safeSet(IDX.MAQUINA, maqLimpia);

   safeSet(IDX.PESO, datos.PESO);
   safeSet(IDX.TIPO, headers.indexOf("TIPO"));
   safeSet(IDX.DIAMETRO, headers.indexOf("DIAMETRO"));
   safeSet(IDX.LONGITUD, headers.indexOf("LONGITUD"));
   safeSet(IDX.CUERDA, headers.indexOf("CUERDA"));
   safeSet(IDX.CUERPO, headers.indexOf("CUERPO"));
   safeSet(IDX.ACERO, headers.indexOf("ACERO"));
   safeSet(IDX.PT, headers.indexOf("PT"));
   safeSet(IDX.VENTA, headers.indexOf("VENTA"));

   // RECALCULO DE SOLICITADO Y ESTADO
   if (IDX.CANT > -1 && IDX.SOL > -1 && IDX.PESO > -1) {
       var cantidad = Number(datosFilaCompleta[IDX.CANT]) || 0;
       var unidad = String(datosFilaCompleta[IDX.UNI] || "PZA").toUpperCase();
       var peso = parseFloat(datos.PESO) || 0;
       
       var esDirecto = (unidad.includes("KG") || unidad.includes("KILOS") || unidad.includes("ROL"));
       var nuevoSolicitado = esDirecto ? cantidad : (cantidad * peso);
       nuevoSolicitado = Math.round(nuevoSolicitado * 100) / 100;
       
       sheet.getRange(fila, IDX.SOL + 1).setValue(nuevoSolicitado);

       // Recálculo de Estado
       if (IDX.PROD > -1 && IDX.ESTADO > -1) {
           var producidoReal = Number(sheet.getRange(fila, IDX.PROD + 1).getValue()) || 0;
           var estadoActual = String(sheet.getRange(fila, IDX.ESTADO + 1).getValue());
           
           if (estadoActual !== "CANCELADO" && estadoActual !== "CERRADA") {
               var ratio = (nuevoSolicitado > 0) ? (producidoReal / nuevoSolicitado) : 0;
               var nuevoEstado = "ABIERTO";
               if (producidoReal > 0) {
                  if (ratio >= 1.05) nuevoEstado = "TERMINADO";
                  else nuevoEstado = "EN PROCESO";
               }
               sheet.getRange(fila, IDX.ESTADO + 1).setValue(nuevoEstado);
           }
       }
   }
}

// Helper: Inserta nueva con protección
function insertarFilaOrden(sheet, datosBase, datosNuevos, headers, IDX) {
   var nextId = obtenerSiguienteIdNumerico(sheet, IDX.ID);
   var nuevaFila = [...datosBase]; 
   
   var safeSet = function(idx, val) {
      if(idx !== undefined && idx > -1) nuevaFila[idx] = val;
   };
   
   safeSet(IDX.ID, nextId); 
   safeSet(IDX.SEC, datosNuevos.SEC);
   safeSet(IDX.PROCESO, datosNuevos.PROCESO);
   
   var maq = (datosNuevos.MAQUINA) ? String(datosNuevos.MAQUINA).split(",")[0].trim() : "";
   safeSet(headers.indexOf("MAQUINA"), maq);

   safeSet(IDX.PESO, datosNuevos.PESO);
   safeSet(IDX.PT, headers.indexOf("PT"));
   safeSet(IDX.VENTA, headers.indexOf("VENTA"));
   
   if (IDX.CANT > -1 && IDX.SOL > -1) {
       var cantidad = Number(nuevaFila[IDX.CANT]) || 0;
       var peso = parseFloat(datosNuevos.PESO) || 0;
       var unidad = String(nuevaFila[IDX.UNI] || "PZA").toUpperCase();
       var esDirecto = (unidad.includes("KG") || unidad.includes("KILOS") || unidad.includes("ROL"));
       nuevaFila[IDX.SOL] = esDirecto ? cantidad : (cantidad * peso);
   }

   if(IDX.PROD > -1) nuevaFila[IDX.PROD] = 0;
   if(IDX.ESTADO > -1) nuevaFila[IDX.ESTADO] = "ABIERTO";
   
   sheet.appendRow(nuevaFila);
}

// HELPER: OBTENER SIGUIENTE ID NUMÉRICO
function obtenerSiguienteIdNumerico(sheet, colIndexID) {
  var data = sheet.getDataRange().getValues();
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    var val = Number(data[i][colIndexID]);
    if (!isNaN(val) && val > maxId) {
      maxId = val;
    }
  }
  return maxId + 1;
}

// HELPER: Convierte texto de fracción (ej. "5 1/2") a número decimal
function evaluarFraccionPlanif(texto) {
  if (!texto) return 0;
  var s = String(texto).trim();
  var partes = s.split(' ');
  var valor = 0;
  partes.forEach(p => {
    if (p.includes('/')) {
      var f = p.split('/');
      valor += (Number(f[0]) / Number(f[1]));
    } else {
      valor += Number(p);
    }
  });
  return isNaN(valor) ? 0 : valor;
}

// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +- GESTION DOCUMENTAL +-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

// Sube un archivo a Google Drive en la carpeta de documentación técnica
// Llamado por gdoc_subirUnArchivo()
function subirArchivoDrive(data, nombre, mimeType) {
  try {
    // Carpeta destino — ajusta el ID a tu carpeta de documentación
    var FOLDER_ID = "18znDdtnJ-u1jeyKb2o9IXYsLfUJsSc8Y";
    var folder = DriveApp.getFolderById(FOLDER_ID);

    // data viene como base64 desde FileReader.readAsDataURL → "data:tipo;base64,XXXX"
    var base64 = data.split(",")[1];
    var decoded = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, nombre);
    folder.createFile(decoded);
    return { success: true };
  } catch(e) {
    return { success: false, error: String(e) };
  }
}

// Marca un documento como obsoleto renombrándolo con prefijo [OBSOLETO] en Drive
// Llamado por gdoc_confirmarObsolescencia()
function marcarDocumentoObsoleto(url) {
  try {
    // Extraer ID del archivo desde la URL de Drive
    var match = url.match(/[-\w]{25,}/);
    if (!match) return "Error: No se pudo extraer el ID del archivo desde la URL.";
    var fileId = match[0];
    var file   = DriveApp.getFileById(fileId);
    var nombre = file.getName();
    if (nombre.indexOf("[OBSOLETO]") < 0) {
      file.setName("[OBSOLETO] " + nombre);
    }
    return "✅ Documento marcado como obsoleto: " + nombre;
  } catch(e) {
    return "Error: " + e;
  }
}

// Lanza el indexador remoto después de subir archivos
// El indexador debe ser un Web App separado que actualiza BIBLIOTECA_DIGITAL
// Si no tienes indexador, este stub evita que el frontend quede bloqueado
function ejecutarIndexadorRemoto() {
  try {
    // Si tienes una URL de indexador, descomenta y ajusta:
    // var URL_INDEXADOR = "https://script.google.com/macros/s/TU_INDEXADOR/exec";
    // UrlFetchApp.fetch(URL_INDEXADOR, { method: "get", muteHttpExceptions: true });
    return true; // Devuelve true para que el frontend continúe
  } catch(e) {
    return true; // Aunque falle, no bloqueamos al usuario
  }
}

function obtenerDocsTecnicos(codigo) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("BIBLIOTECA_DIGITAL");
  var data = sheet.getDataRange().getValues();
  
  var codigoObjetivo = String(codigo).trim().toUpperCase();
  var resultados = {};
  var hayDocs = false;

  // Barremos toda la biblioteca
  for(var i=1; i<data.length; i++) {
     var rowCod = String(data[i][1]).trim().toUpperCase(); // Col B: CODIGO
     
     if(rowCod === codigoObjetivo) {
        var proceso = String(data[i][2]).trim().toUpperCase(); // Col C: PROCESO
        var tipo = String(data[i][3]).trim().toUpperCase();    // Col D: TIPO
        var nombre = data[i][4];                               // Col E: NOMBRE
        var url = data[i][5];                                  // Col F: URL

        // Si el tipo está vacío, lo ponemos como OTROS
        if(tipo === "") tipo = "OTROS DOCUMENTOS";
        
        // Agrupar por TIPO (Ficha, Herramental, etc.)
        if(!resultados[tipo]) resultados[tipo] = [];
        
        resultados[tipo].push({
           nombre: nombre,
           proceso: proceso, // Guardamos el proceso solo como referencia visual
           url: url
        });
        hayDocs = true;
     }
  }
  
  // Obtenemos descripción del código para el encabezado
  var desc = "";
  var sheetCod = ss.getSheetByName("CODIGOS");
  var dataCod = sheetCod.getDataRange().getValues();
  for(var c=1; c<dataCod.length; c++) {
     if(String(dataCod[c][0]).trim().toUpperCase() === codigoObjetivo) {
        // Asumiendo Col B (indice 1) es Descripción, ajusta si es diferente
        desc = dataCod[c][1]; 
        break;
     }
  }

  return { existe: hayDocs, datos: resultados, descripcion: desc, codigo: codigoObjetivo };
}


// =================================================================================
// INDEXADOR AUTOMÁTICO DE DOCUMENTOS (FIX: CEROS A LA IZQUIERDA)
// =================================================================================
function indexarBibliotecaDigital() {
  // CONFIGURACIÓN: ID de tu carpeta "SISTEMA_DOCUMENTAL"
  var ID_CARPETA_RAIZ = "18znDdtnJ-u1jeyKb2o9IXYsLfUJsSc8Y"; 

  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("BIBLIOTECA_DIGITAL");
  
  if(!sheet) {
    sheet = ss.insertSheet("BIBLIOTECA_DIGITAL");
    sheet.appendRow(["ID", "CODIGO", "PROCESO", "TIPO", "NOMBRE", "URL", "FECHA_ACT"]);
  }
  
  var carpeta = DriveApp.getFolderById(ID_CARPETA_RAIZ);
  var archivos = carpeta.getFiles();
  
  var baseDatos = [];
  var fechaHoy = new Date();
  
  while (archivos.hasNext()) {
    var file = archivos.next();
    var nombreCompleto = file.getName();
    
    // IGNORAR SI ES OBSOLETO
    if (nombreCompleto.toUpperCase().indexOf("_OBSOLETO") > -1) {
       continue; 
    }

    var nombreSinExt = nombreCompleto.toUpperCase().replace(".PDF","");
    var url = file.getUrl();
    
    // LOGICA DE NOMENCLATURA
    var partes = nombreSinExt.split("_");
    
    // 1. El primer elemento SIEMPRE es el CÓDIGO (Forzamos String)
    var codigo = String(partes[0]).trim(); 
    
    var proceso = "";
    var tipo = "OTROS DOCUMENTOS";
    
    // 2. DETECCIÓN TIPO
    if (nombreSinExt.indexOf("FICHATECNICA") > -1) {
       tipo = "FICHA_TECNICA";
    } 
    else if (nombreSinExt.indexOf("FICHAFAB") > -1) {
       tipo = "FICHA_FAB";
    } 
    else if (nombreSinExt.indexOf("INSPECCION") > -1) {
       tipo = "INSPECCION";
       if (partes.length >= 3) proceso = partes[1].trim(); 
    } 
    else if (nombreSinExt.indexOf("HERRAMENTAL") > -1) {
       tipo = "HERRAMENTAL";
       if (partes.length >= 3) proceso = partes[1].trim();
    }
    else if (partes.length >= 2) {
       proceso = partes[1].trim();
    }

    baseDatos.push([
      Utilities.getUuid(), 
      codigo, 
      proceso, 
      tipo, 
      nombreCompleto, 
      url, 
      fechaHoy
    ]);
  }
  
  // 3. ESCRIBIR EN LA HOJA
  sheet.clearContents();
  sheet.appendRow(["ID", "CODIGO", "PROCESO", "TIPO", "NOMBRE", "URL", "FECHA_ACT"]);
  
  if(baseDatos.length > 0) {
    // --- CORRECCIÓN CLAVE ---
    // Forzamos formato TEXTO PLANO (@) en la columna B (Código) 
    // para que Google Sheets no borre los ceros a la izquierda (0100...)
    sheet.getRange(2, 2, baseDatos.length, 1).setNumberFormat("@");
    
    // Pegar datos
    sheet.getRange(2, 1, baseDatos.length, 7).setValues(baseDatos);
  }
  
  Logger.log("✅ Indexación completada. Se encontraron " + baseDatos.length + " documentos activos.");
  return "Indexación completada: " + baseDatos.length + " docs activos.";
}

// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +- RESPALDO DIARIO A LAS 2AM +-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
// +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

function obtenerDatosPlanProduccion(procesoInput) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var data     = sheetOrd.getDataRange().getValues();
  var headers  = data[0];
  var getIdx   = function(name) { return headers.indexOf(name); };
  var idx = {
    orden:  getIdx("ORDEN"),   codigo: getIdx("CODIGO"),
    tipo:   getIdx("TIPO"),    dia:    getIdx("DIAMETRO"), long: getIdx("LONGITUD"),
    cuerda: getIdx("CUERDA"),  cuerpo: getIdx("CUERPO"),  acero: getIdx("ACERO"),
    sol:    getIdx("SOLICITADO"), prod: getIdx("PRODUCIDO"), est: getIdx("ESTADO"),
    maq:    getIdx("MAQUINA"), proc:   getIdx("PROCESO"),  prio: getIdx("PRIORIDAD"),
    inicio: getIdx("FECHA_INICIO_PROG"), nota: getIdx("NOTA")
  };

  var EXCLUIR = ["CANCELADO", "SOBREPRODUCCION", "TERMINADO"];
  var gruposMaquinas = {};

  for (var i = 1; i < data.length; i++) {
    var row     = data[i];
    var proceso = String(row[idx.proc]).trim();
    if (proceso.toUpperCase() !== String(procesoInput).trim().toUpperCase()) continue;

    var estado = String(row[idx.est]).trim().toUpperCase();
    if (EXCLUIR.indexOf(estado) > -1) continue;

    var maquina = String(row[idx.maq]).trim();
    if (!maquina) continue;

    if (!gruposMaquinas[maquina]) gruposMaquinas[maquina] = [];

    var sol      = Number(row[idx.sol]) || 0;
    var prod     = Number(row[idx.prod]) || 0;
    var pct      = sol > 0 ? (prod / sol) * 100 : 0;
    var prioRaw  = row[idx.prio];
    var tienePrio = (prioRaw !== "" && prioRaw !== null && prioRaw !== undefined && !isNaN(Number(prioRaw)));

    var fechaInicioStr = "";
    if (row[idx.inicio] instanceof Date) {
      var fd = row[idx.inicio];
      fechaInicioStr = (fd.getDate() < 10 ? "0" : "") + fd.getDate() + "/"
                     + (fd.getMonth()+1 < 10 ? "0" : "") + (fd.getMonth()+1);
    }

    gruposMaquinas[maquina].push({
      orden:       "P." + ("0000" + row[idx.orden]).slice(-4),
      codigo:      String(row[idx.codigo] || ""),
      producto:    String(row[idx.tipo]   || ""),
      diametro:    String(row[idx.dia]    || ""),
      longitud:    String(row[idx.long]   || ""),
      detalle:     (String(row[idx.cuerda]||"") + " " + String(row[idx.cuerpo]||"")).trim(),
      acero:       String(row[idx.acero]  || ""),
      sol:         sol, prod: prod, avance: pct, restan: sol - prod,
      fechaInicio: fechaInicioStr,
      nota:        idx.nota > -1 ? String(row[idx.nota] || "") : "",
      prioridad:   tienePrio ? Number(prioRaw) : null
    });
  }

  var listaFinal = [];
  Object.keys(gruposMaquinas).sort().forEach(function(maq) {
    var conPrio = gruposMaquinas[maq].filter(function(o){ return o.prioridad !== null; });
    var sinPrio = gruposMaquinas[maq].filter(function(o){ return o.prioridad === null; });
    conPrio.sort(function(a,b){ return a.prioridad - b.prioridad; });
    listaFinal.push({ nombreMaquina: maq, ordenes: conPrio.concat(sinPrio) });
  });

  return JSON.stringify(listaFinal);
}

function obtenerDatosReporteTurnoTsup(procesoInput) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd  = ss.getSheetByName("ORDENES");
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var sheetEst  = ss.getSheetByName("ESTANDARES");
  var dataOrd   = sheetOrd.getDataRange().getValues();
  var dataProd  = sheetProd.getDataRange().getValues();
  var dataEst   = sheetEst ? sheetEst.getDataRange().getValues() : [];
  var headersOrd = dataOrd[0];
  var getIdxO    = function(n) { return headersOrd.indexOf(n); };
  var idxO = {
    id: 0, orden: getIdxO("ORDEN"), serie: getIdxO("SERIE"),
    pedido: getIdxO("PEDIDO"),
    codigo: getIdxO("CODIGO"), tipo: getIdxO("TIPO"),
    dia: getIdxO("DIAMETRO"), long: getIdxO("LONGITUD"),
    cuerda: getIdxO("CUERDA"), cuerpo: getIdxO("CUERPO"),
    acero: getIdxO("ACERO"), sol: getIdxO("SOLICITADO"),
    prod: getIdxO("PRODUCIDO"), est: getIdxO("ESTADO"),
    maq: getIdxO("MAQUINA"), proc: getIdxO("PROCESO"),
    prio: getIdxO("PRIORIDAD")
  };
  // Mapa máquina → grupo desde ESTANDARES (col D=idx3:MAQUINA, col J=idx9:GRUPO)
  var maqGrupoMap = {};
  for (var s = 1; s < dataEst.length; s++) {
    var maqEst   = String(dataEst[s][3]).trim().toUpperCase();
    var grupoEst = String(dataEst[s][9]).trim().toUpperCase();
    if (maqEst && grupoEst) maqGrupoMap[maqEst] = grupoEst;
  }
  var headersProd = dataProd[0];
  var findCol = function(candidates) {
    for (var i = 0; i < candidates.length; i++) {
      var idx = headersProd.indexOf(candidates[i]);
      if (idx > -1) return idx;
    }
    return -1;
  };
  var idxP = {
    ordenRef: 2,
    loteTxt:  findCol(["LOTE","ID_LOTE"]) > -1 ? findCol(["LOTE","ID_LOTE"]) : 4,
    pIni:     findCol(["PESO_I","PESO_INICIAL","P_INICIAL","PESO INICIAL"]),
    pFin:     findCol(["PESO_F","PESO_FINAL","P_FINAL","PESO FINAL"]),
    fecha:    findCol(["FECHA","FECHA_REGISTRO","TIMESTAMP"])
  };
  var ultimosDatos = {};
  for (var p = 1; p < dataProd.length; p++) {
    var idOrd    = String(dataProd[p][idxP.ordenRef]);
    var rawFecha = (idxP.fecha > -1) ? dataProd[p][idxP.fecha] : new Date();
    var fechaRow = (rawFecha instanceof Date) ? rawFecha : new Date(rawFecha);
    if (!ultimosDatos[idOrd] || fechaRow > ultimosDatos[idOrd].fecha) {
      ultimosDatos[idOrd] = {
        lote: dataProd[p][idxP.loteTxt],
        pIni: idxP.pIni > -1 ? dataProd[p][idxP.pIni] : "",
        pFin: idxP.pFin > -1 ? dataProd[p][idxP.pFin] : "",
        fecha: fechaRow
      };
    }
  }
  // FORM.PROD solo muestra ACTIVE y EN PROCESO
  var PERMITIDOS = ["ACTIVE", "EN PROCESO", "ABIERTO"];
  var listaMaquinas = {};
  for (var i = 1; i < dataOrd.length; i++) {
    var row     = dataOrd[i];
    var proceso = String(row[idxO.proc]).trim();
    var estado  = String(row[idxO.est]).trim().toUpperCase();
    if (proceso.toUpperCase() !== String(procesoInput).trim().toUpperCase()) continue;
    if (PERMITIDOS.indexOf(estado) < 0) continue;
    var maquina = String(row[idxO.maq]).trim();
    if (!maquina) continue;
    if (!listaMaquinas[maquina]) {
      var grupoMaquina = maqGrupoMap[maquina.toUpperCase()] || '';
      listaMaquinas[maquina] = { nombre: maquina, grupo: grupoMaquina, ordenes: [] };
    }
    var sol  = Number(row[idxO.sol]) || 0;
    var prod = Number(row[idxO.prod]) || 0;
    var pct  = sol > 0 ? (prod / sol) * 100 : 0;
    var numOrden    = row[idxO.orden];
    var nombreOrden = String(row[idxO.serie] || "P") + "." + ("0000" + numOrden).slice(-4);
    var last = ultimosDatos[String(row[0])] || { lote: "-", pIni: "", pFin: "" };
    listaMaquinas[maquina].ordenes.push({
      ordenFull: nombreOrden,
      pedido:    String(row[idxO.pedido] || ""),
      producto:  String(row[idxO.tipo] || ""),
      medidas:   String(row[idxO.dia] || "") + " x " + String(row[idxO.long] || ""),
      detalle:   (String(row[idxO.cuerda]||"") + " " + String(row[idxO.cuerpo]||"")).trim(),
      cuerda:    String(row[idxO.cuerda] || ""),
      cuerpo:    String(row[idxO.cuerpo] || ""),
      acero:     String(row[idxO.acero] || ""),
      sol: sol, prod: prod, avance: pct, restan: sol - prod,
      prioridad: idxO.prio >= 0 ? (Number(row[idxO.prio]) || 999) : 999,
      ultimoLote: last.lote,
      ultimoPIni: last.pIni,
      ultimoPFin: last.pFin
    });
  }
  // Ordenar órdenes de cada máquina por prioridad ascendente
  Object.keys(listaMaquinas).forEach(function(k) {
    listaMaquinas[k].ordenes.sort(function(a, b) {
      return (a.prioridad || 999) - (b.prioridad || 999);
    });
  });
  return JSON.stringify(
    Object.keys(listaMaquinas).sort().map(function(k){ return listaMaquinas[k]; })
  );
}

// ─── MIS REGISTROS DE PRODUCCIÓN ──────────────────────────────────────────────
function obtenerMisRegistros(nombreUsuario) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var sheetOrd  = ss.getSheetByName("ORDENES");

  // ── 1. Leer ORDENES una sola vez → mapa ordenID → producto ──
  var dataOrd = sheetOrd.getDataRange().getValues();
  var hOrd    = dataOrd[0].map(function(x){ return String(x).toUpperCase().trim(); });
  var oTIPO   = hOrd.indexOf("TIPO");
  var oDIA    = hOrd.indexOf("DIAMETRO");
  var oLON    = hOrd.indexOf("LONGITUD");
  var oCUERPO = hOrd.indexOf("CUERPO");
  var oCUERDA = hOrd.indexOf("CUERDA");
  var oACERO  = hOrd.indexOf("ACERO");
  var mapaProducto = {};
  for (var o = 1; o < dataOrd.length; o++) {
    var or     = dataOrd[o];
    var ordId  = String(or[0]);
    var tipo   = String(or[oTIPO]   || "").trim();
    var dia    = String(or[oDIA]    || "").trim();
    var lon    = String(or[oLON]    || "").trim();
    var cuerpo = oCUERPO >= 0 ? String(or[oCUERPO] || "").trim() : "";
    var cuerda = oCUERDA >= 0 ? String(or[oCUERDA] || "").trim() : "";
    var acero  = oACERO  >= 0 ? String(or[oACERO]  || "").trim() : "";
    var partes = [tipo];
    if (dia || lon) partes.push(dia + " X " + lon);
    if (cuerpo) partes.push(cuerpo);
    if (cuerda) partes.push(cuerda);
    if (acero)  partes.push("Acero " + acero);
    mapaProducto[ordId] = partes.filter(function(p){ return p !== ""; }).join(" ");
  }

  // ── 2. Leer PRODUCCION ──
  var data = sheetProd.getDataRange().getValues();
  var h    = data[0].map(function(x){ return String(x).toUpperCase().trim(); });

  var IDX = {
    ID:       h.indexOf("ID"),
    SERIE:    h.indexOf("SERIE"),
    ORDEN:    h.indexOf("ORDEN"),
    LOTE:     h.indexOf("LOTE"),
    MAQ:      h.indexOf("MAQUINA"),
    FECHA:    h.indexOf("FECHA"),
    TURNO:    h.indexOf("TURNO"),
    OP_TXT:   h.indexOf("NOMBRE_OPERADOR_TXT"),
    PESO_I:   h.indexOf("PESO_I"),
    PESO_F:   h.indexOf("PESO_F"),
    PESO_T:   h.indexOf("PESO_TINA"),
    PROD:     h.indexOf("PRODUCIDO"),
    SELLO:    h.indexOf("SELLO"),
    COM:      h.indexOf("COMENTARIO"),
    USER:     h.indexOf("USER"),
    PROCESO:  h.indexOf("PROCESO"),
    FECHA_REG:h.indexOf("FECHA_REG"),
    CAMBIOS:  h.indexOf("CAMBIOS")
  };

  // ── Mapa ordenID → proceso (fallback cuando PRODUCCION no tiene col PROCESO) ──
  var mapaProceso = {};
  var iProcOrd = hOrd.indexOf("PROCESO");
  if (iProcOrd >= 0) {
    for (var o2 = 1; o2 < dataOrd.length; o2++) {
      var idO = String(dataOrd[o2][0]);
      var prO = String(dataOrd[o2][iProcOrd] || "").trim().toUpperCase();
      if (idO && prO && !mapaProceso[idO]) mapaProceso[idO] = prO;
    }
  }

  // Función para obtener proceso de una fila: primero de PRODUCCION, luego de ORDENES
  function getProceso(row) {
    if (IDX.PROCESO >= 0) {
      var p = String(row[IDX.PROCESO] || "").trim().toUpperCase();
      if (p) return p;
    }
    var ordRef = String(row[IDX.ORDEN] || "");
    return mapaProceso[ordRef] || "";
  }

  // ── Helper: convertir cualquier valor a Date de forma robusta ──
  // Col Z FECHA_REG puede venir como Date nativo o como string "17/3/2026 17:22:43"
  function toDate(val) {
    if (!val || val === "") return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
    // Intentar parsear string — soporta dd/MM/yyyy y yyyy-MM-dd
    var s = String(val).trim();
    // Formato dd/MM/yyyy HH:mm:ss o dd/MM/yyyy
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      var d = new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
      return isNaN(d.getTime()) ? null : d;
    }
    // Formato yyyy-MM-dd
    var d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
  }

  // ── 3. Límite: hace 2 días a las 00:00:00 ──
  var hoy       = new Date();
  var limiteMin = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate() - 2, 0, 0, 0);

  // ── 4. Primera pasada: encontrar procesos del usuario en el rango ──
  var procesosUsuario = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // FECHA_REG debe existir y estar dentro del rango (col Z)
    var fechaReg = toDate(row[IDX.FECHA_REG]);
    if (!fechaReg || fechaReg < limiteMin) continue;
    var user = String(row[IDX.USER] || "").trim().toUpperCase();
    if (user !== nombreUsuario.toUpperCase().trim()) continue;
    var proc = getProceso(row);
    if (proc) procesosUsuario[proc] = true;
  }

  // ── 5. Segunda pasada: traer TODOS los registros de esos procesos en el rango ──
  var resultados = [];
  for (var i = 1; i < data.length; i++) {
    var row      = data[i];
    var fechaReg = toDate(row[IDX.FECHA_REG]);
    if (!fechaReg || fechaReg < limiteMin) continue;  // sin FECHA_REG → omitir
    var proc = getProceso(row);
    if (!procesosUsuario[proc]) continue;

    var fVal = row[IDX.FECHA];
    var fStr = "";
    if (fVal instanceof Date && !isNaN(fVal.getTime())) {
      fStr = fVal.getFullYear()
           + "-" + (fVal.getMonth()+1 < 10 ? "0" : "") + (fVal.getMonth()+1)
           + "-" + (fVal.getDate()    < 10 ? "0" : "") +  fVal.getDate();
    }

    var ordId = String(row[IDX.ORDEN] || "");
    resultados.push({
      id:          String(row[IDX.ID]    || ""),
      fila:        i + 1,
      serie:       String(row[IDX.SERIE] || ""),
      ordenRef:    ordId,
      lote:        String(row[IDX.LOTE]  || ""),
      maquina:     String(row[IDX.MAQ]   || ""),
      producto:    mapaProducto[ordId]   || "",
      fecha:       fStr,
      turno:       String(row[IDX.TURNO] || ""),
      operadorTxt: String(row[IDX.OP_TXT]|| ""),
      pesoI:       Number(row[IDX.PESO_I])  || 0,
      pesoF:       Number(row[IDX.PESO_F])  || 0,
      pesoTina:    Number(row[IDX.PESO_T])  || 0,
      producido:   Number(row[IDX.PROD])    || 0,
      sello:       String(row[IDX.SELLO]   || ""),
      comentario:  String(row[IDX.COM]     || ""),
      registrador: String(row[IDX.USER]    || ""),
      proceso:     getProceso(row),
      cambios:     String(row[IDX.CAMBIOS] || "")
    });
  }

  resultados.sort(function(a,b){ return b.fecha.localeCompare(a.fecha) || b.turno - a.turno; });
  return JSON.stringify(resultados);
}

function guardarCambioMiRegistro(payload) {
  try {
    var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var sheetProd = ss.getSheetByName("PRODUCCION");
    var data      = sheetProd.getDataRange().getValues();
    var h         = data[0].map(function(x){ return String(x).toUpperCase().trim(); });

    var mapaColumnas = {
      fecha:    h.indexOf("FECHA")    + 1,
      turno:    h.indexOf("TURNO")    + 1,
      maquina:  h.indexOf("MAQUINA")  + 1,
      pesoI:    h.indexOf("PESO_I")   + 1,
      pesoF:    h.indexOf("PESO_F")   + 1,
      pesoTina: h.indexOf("PESO_TINA")+ 1,
      producido:h.indexOf("PRODUCIDO")+ 1
    };
    var colCAMBIOS = h.indexOf("CAMBIOS") + 1;

    var colNum = mapaColumnas[payload.campo];
    if (!colNum || colNum < 1) return JSON.stringify({ success: false, msg: "Campo no válido: " + payload.campo });

    var filaVerif = payload.fila - 1;
    if (filaVerif < 1 || filaVerif >= data.length)
      return JSON.stringify({ success: false, msg: "Fila fuera de rango" });
    if (String(data[filaVerif][0]) !== String(payload.id))
      return JSON.stringify({ success: false, msg: "ID no coincide" });

    var valorAnterior = data[filaVerif][colNum - 1];
    var valor = payload.valorNuevo;

    if (payload.campo === "fecha") {
      valor = new Date(valor + "T12:00:00");
      valorAnterior = (valorAnterior instanceof Date)
        ? valorAnterior.getFullYear() + "-" + (valorAnterior.getMonth()+1) + "-" + valorAnterior.getDate()
        : valorAnterior;
    } else if (["pesoI","pesoF","pesoTina","producido"].indexOf(payload.campo) >= 0) {
      valor = Number(valor) || 0;
      valorAnterior = Number(valorAnterior) || 0;
    }
    sheetProd.getRange(payload.fila, colNum).setValue(valor);

    // Si cambió un peso, recalcular producido = pesoF - pesoTina - pesoI y guardarlo
    if (["pesoI", "pesoF", "pesoTina"].indexOf(payload.campo) >= 0) {
      var colPI = mapaColumnas.pesoI    - 1;
      var colPF = mapaColumnas.pesoF    - 1;
      var colPT = mapaColumnas.pesoTina - 1;
      var colPROD = mapaColumnas.producido - 1;
      // Leer la fila actualizada (ya escribimos el nuevo valor arriba)
      var filaActual = sheetProd.getRange(payload.fila, 1, 1, sheetProd.getLastColumn()).getValues()[0];
      var newPI   = Number(filaActual[colPI])   || 0;
      var newPF   = Number(filaActual[colPF])   || 0;
      var newPT   = Number(filaActual[colPT])   || 0;
      var newProd = Math.round(newPF - newPT - newPI);
      if (newProd < 0) newProd = 0;
      sheetProd.getRange(payload.fila, mapaColumnas.producido).setValue(newProd);
      // Recalcular estado de la orden maestra
      recalcularEstadoOrdenMaestro(String(data[filaVerif][2]));
    }

    if (payload.campo === "producido") {
      recalcularEstadoOrdenMaestro(String(data[filaVerif][2]));
    }

    if (colCAMBIOS > 0) {
      var entradaAnterior = String(data[filaVerif][colCAMBIOS - 1] || "").trim();
      var etiquetas = { fecha:"FECHA", turno:"TURNO", maquina:"MAQUINA", pesoI:"PESO_I", pesoF:"PESO_F", pesoTina:"PESO_TINA", producido:"PRODUCIDO" };
      var nuevaEntrada = payload.nombreUsuario + " cambió " + (etiquetas[payload.campo] || payload.campo)
                       + " de " + valorAnterior + " a " + payload.valorNuevo;
      var historico = entradaAnterior ? entradaAnterior + " | " + nuevaEntrada : nuevaEntrada;
      sheetProd.getRange(payload.fila, colCAMBIOS).setValue(historico);
    }

    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, msg: e.toString() });
  }
}

function obtenerPedidosPendientes(incluirQSQ) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shPed    = ss.getSheetByName("PEDIDOS");
  var shOrd    = ss.getSheetByName("ORDENES");
  var dataPed  = shPed.getDataRange().getValues();
  var dataOrd  = shOrd.getDataRange().getValues();

  var hPed = dataPed[0].map(function(x){ return String(x).toUpperCase().trim(); });
  var hOrd = dataOrd[0].map(function(x){ return String(x).toUpperCase().trim(); });

  var iPFolio = hPed.indexOf("PEDIDO");      if(iPFolio<0) iPFolio=1;
  var iPFecha = hPed.indexOf("FECHA");       if(iPFecha<0) iPFecha=2;
  var iPCod   = hPed.indexOf("CODIGO");      if(iPCod<0)   iPCod=3;
  var iPDesc  = hPed.indexOf("DESCRIPCION"); if(iPDesc<0)  iPDesc=4;
  var iPCant  = hPed.indexOf("CANTIDAD");    if(iPCant<0)  iPCant=6;
  var iPEst   = hPed.indexOf("ESTADO");      if(iPEst<0)   iPEst=8;

  var iOPed   = hOrd.indexOf("PEDIDO");      if(iOPed<0)   iOPed=8;
  var iOProc  = hOrd.indexOf("PROCESO");     if(iOProc<0)  iOProc=11;
  var iOMaq   = hOrd.indexOf("MAQUINA");     if(iOMaq<0)   iOMaq=12;
  var iOProd  = hOrd.indexOf("PRODUCIDO");   if(iOProd<0)  iOProd=14;
  var iOEst   = hOrd.indexOf("ESTADO");      if(iOEst<0)   iOEst=15;
  var iOSol   = hOrd.indexOf("SOLICITADO");  if(iOSol<0)   iOSol=13;
  var iOSec   = hOrd.indexOf("SEC");         if(iOSec<0)   iOSec=10;

  var mapaOrdenes = {};
  for (var i = 1; i < dataOrd.length; i++) {
    var oRow   = dataOrd[i];
    var oFolio = String(oRow[iOPed]  || "").trim();
    var oEst   = String(oRow[iOEst]  || "").toUpperCase().trim();
    if (!oFolio) continue;
    if (!mapaOrdenes[oFolio]) mapaOrdenes[oFolio] = [];
    mapaOrdenes[oFolio].push({
      proceso:    String(oRow[iOProc] || "").trim(),
      maquina:    String(oRow[iOMaq]  || "").trim(),
      producido:  Number(oRow[iOProd] || 0),
      solicitado: Number(oRow[iOSol]  || 0),
      estado:     oEst,
      sec:        Number(oRow[iOSec]  || 0)
    });
  }

  var hoy      = new Date();
  var limite60 = new Date(hoy.getTime() - 60 * 24 * 3600 * 1000);
  var EXCLUIR  = ["CANCELADO", "TERMINADO"];
  var SERIES_BASE = ["ZEQ-", "ZRR-", "MAQ-", "MPR-"];

  var resultados = [];

  for (var i = 1; i < dataPed.length; i++) {
    var row    = dataPed[i];
    var folio  = String(row[iPFolio] || "").trim();
    var estado = String(row[iPEst]   || "").toUpperCase().trim();
    if (!folio) continue;
    if (EXCLUIR.indexOf(estado) > -1) continue;

    var fechaRaw = row[iPFecha];
    var fecha    = (fechaRaw instanceof Date) ? fechaRaw : new Date(fechaRaw);
    var fechaStr = "";
    if (fecha instanceof Date && !isNaN(fecha)) {
      fechaStr = (fecha.getDate()<10?"0":"")+fecha.getDate()+"/"
               + (fecha.getMonth()+1<10?"0":"")+(fecha.getMonth()+1)+"/"
               + fecha.getFullYear();
    }

    var esSerie = SERIES_BASE.some(function(s){ return folio.indexOf(s) === 0; });
    var esQSQ   = folio.indexOf("QSQ-") === 0
               && fecha instanceof Date && !isNaN(fecha)
               && fecha <= limite60;

    if (!esSerie && !(incluirQSQ && esQSQ)) continue;

    var ordenes = (mapaOrdenes[folio] || []).slice();
    ordenes.sort(function(a,b){ return a.sec - b.sec; });

    var CERRADOS = ["CANCELADO","TERMINADO","SOBREPRODUCCION"];
    var abiertas = ordenes.filter(function(o){ return CERRADOS.indexOf(o.estado) < 0; });
    var serie    = folio.split("-")[0] + "-";

    resultados.push({
      folio:           folio,
      fecha:           fechaStr,
      fechaMs:         fecha instanceof Date ? fecha.getTime() : 0,
      codigo:          String(row[iPCod]  || "").trim(),
      descripcion:     String(row[iPDesc] || "").trim(),
      cantidad:        Number(row[iPCant] || 0),
      estado:          estado,
      serie:           serie,
      numOrdAbiertas:  abiertas.length,
      ordenesAbiertas: abiertas.map(function(o){ return o.proceso + (o.maquina ? " / " + o.maquina : ""); }).join(", "),
      ruta: ordenes.map(function(o){
        return { proceso: o.proceso, maquina: o.maquina, producido: o.producido, solicitado: o.solicitado, estado: o.estado };
      })
    });
  }

  resultados.sort(function(a,b){ return b.fechaMs - a.fechaMs; });
  return JSON.stringify(resultados);
}

function respaldarSheet() {
  var origen = "1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk";
  var carpetaNombre = "Respaldos-Sheets-Prod";
  var fecha = Utilities.formatDate(new Date(), "America/Mexico_City", "yyyy-MM-dd");
  var nombreCopia = "Prod-" + fecha;
  
  var carpetas = DriveApp.getFoldersByName(carpetaNombre);
  var carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(carpetaNombre);

  DriveApp.getFileById(origen).makeCopy(nombreCopia, carpeta);
}

function obtenerDatosEficienciaParaCaptura(registros) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetEst = ss.getSheetByName("ESTANDARES");
  var dataOrd  = sheetOrd.getDataRange().getValues();
  var dataEst  = sheetEst.getDataRange().getValues();

  var cleanNum = function(val) {
    if (val == null || val === "") return 0;
    if (typeof val === 'number') return val;
    var num = parseFloat(String(val).replace(/,/g,'').trim());
    return isNaN(num) ? 0 : num;
  };

  // --- ESTANDARES: mapa MAQUINA -> velocidad ---
  var hStd = dataEst[0];
  var iSMaq = hStd.indexOf("MAQUINA");
  var iSVel = hStd.indexOf("VELOCIDAD");
  var mapStd = {};
  for (var s = 1; s < dataEst.length; s++) {
    var maqS = String(dataEst[s][iSMaq >= 0 ? iSMaq : 3] || "").toUpperCase().trim();
    var velS  = cleanNum(dataEst[s][iSVel >= 0 ? iSVel : 4]);
    if (maqS && !mapStd[maqS]) mapStd[maqS] = velS;
  }

  // --- ORDENES: mapa ID -> { tipo, desc, proceso, peso, diam, longitud, cuerpo, cuerda, acero } ---
  var hOrd = dataOrd[0];
  var iODesc    = hOrd.indexOf("DESCRIPCION");
  var iOTipo    = hOrd.indexOf("TIPO");
  var iOProceso = hOrd.indexOf("PROCESO");
  var iOPeso    = hOrd.indexOf("PESO");
  var iODiam    = hOrd.indexOf("DIAMETRO");
  var iOLong    = hOrd.indexOf("LONGITUD");
  var iOCuerpo  = hOrd.indexOf("CUERPO");
  var iOCuerda  = hOrd.indexOf("CUERDA");
  var iOAcero   = hOrd.indexOf("ACERO");
  var mapOrd = {};
  for (var o = 1; o < dataOrd.length; o++) {
    var ro = dataOrd[o];
    mapOrd[String(ro[0])] = {
      desc:     String(ro[iODesc    >= 0 ? iODesc    : 0] || "").trim(),
      tipo:     String(ro[iOTipo    >= 0 ? iOTipo    : 0] || "").trim(),
      proceso:  String(ro[iOProceso >= 0 ? iOProceso : 0] || "").toUpperCase().trim(),
      peso:     cleanNum(ro[iOPeso  >= 0 ? iOPeso    : 0]),
      diam:     String(ro[iODiam    >= 0 ? iODiam    : 0] || "").trim(),
      longitud: String(ro[iOLong    >= 0 ? iOLong    : 0] || "").trim(),
      cuerpo:   String(ro[iOCuerpo  >= 0 ? iOCuerpo  : 0] || "").trim(),
      cuerda:   String(ro[iOCuerda  >= 0 ? iOCuerda  : 0] || "").trim(),
      acero:    String(ro[iOAcero   >= 0 ? iOAcero   : 0] || "").trim()
    };
  }

  // --- MODO DE CÁLCULO según proceso (igual que obtenerDatosEficiencia) ---
  var getModo = function(proc) {
    var p = proc.toUpperCase();
    if (p.indexOf("COLATADO")   >= 0) return "ROLLOS";
    if (p.indexOf("ROSCADO")    >= 0 || p.indexOf("ESTIRADO") >= 0 ||
        p.indexOf("ENDEREZADO") >= 0 || p.indexOf("CORTE")    >= 0 ||
        p.indexOf("PULIDO")     >= 0 || p.indexOf("TREFILADO")>= 0) return "METROS";
    if (p.indexOf("FORJA")      >= 0 || p.indexOf("ESTAMPADO")>= 0 ||
        p.indexOf("PUNTEADO")   >= 0 || p.indexOf("ROLADO")   >= 0) return "PIEZAS";
    if (p.indexOf("LAVADO")     >= 0) return "KILOS_MIN";
    return "KILOS";
  };

  // --- AGRUPAR registros: proceso -> maquina+operador ---
  var reporte = {};

  registros.forEach(function(r) {
    var proc  = (r.proceso || "SIN PROCESO").toUpperCase().trim();
    var maq   = String(r.maquina  || "").toUpperCase().trim();
    var op    = String(r.op1_txt  || "");
    var prod  = cleanNum(r.producido);
    var horas = r.turno == "2" ? 7.0 : r.turno == "3" ? 8.0 : 7.5;
    var ordId = String(r.ordenID || "");
    var fecha = String(r.fecha   || "");

    var info    = mapOrd[ordId] || { desc:"", tipo:"", proceso: proc, peso:0 };
    var modo    = getModo(proc);
    var std     = mapStd[maq] || 0;
    var pesoUnit = info.peso;

    // Real para eficiencia (igual que obtenerDatosEficiencia)
    var realParaEfic = 0;
    var kilosReales  = prod;
    if (modo === "ROLLOS") {
      realParaEfic = prod;
    } else if (modo === "PIEZAS" || modo === "METROS") {
      realParaEfic = (pesoUnit > 0) ? prod / pesoUnit : prod;
    } else {
      realParaEfic = prod;
    }

    // Producto para mostrar en tabla
    var producto = (info.tipo + " " + info.desc).trim() || ordId;

    if (!reporte[proc]) reporte[proc] = { nombre: proc, modo: modo, grupos: {} };

    var gKey = maq + "||" + op;
    if (!reporte[proc].grupos[gKey]) {
      reporte[proc].grupos[gKey] = {
        maquina: maq, operador: op,
        totalKilos: 0, sumaReal: 0, sumaTeorico: 0, sumaTeoricoKg: 0,
        turnosProcesados: {},
        detalles: []
      };
    }
    var grp = reporte[proc].grupos[gKey];

    // Anti-duplicado horas: una sola vez por fecha+turno por máquina
    var llaveTurno = fecha + "_" + r.turno;
    if (!grp.turnosProcesados[llaveTurno]) {
      var teorico = 0;
      if (std > 0) {
        if (modo === "ROLLOS" || modo === "PIEZAS" || modo === "METROS" || modo === "KILOS_MIN") {
          teorico = std * 60 * horas;
        } else {
          teorico = std * horas;
        }
      }
      grp.sumaTeorico   += teorico;
      // sumaTeoricoKg: convertir a kg según modo
      if (modo === "PIEZAS" || modo === "METROS") {
        grp.sumaTeoricoKg += (pesoUnit > 0) ? teorico * pesoUnit : 0;
      } else {
        // KILOS, KILOS_MIN y ROLLOS: el teórico ya está en kg
        grp.sumaTeoricoKg += teorico;
      }
      grp.turnosProcesados[llaveTurno] = true;
    }

    grp.totalKilos += kilosReales;
    grp.sumaReal   += realParaEfic;
    grp.detalles.push({
      lote:     r.loteFull  || "",
      producto: producto,
      cantidad: prod,
      tipo:     info.tipo     || "",
      diam:     info.diam     || "",
      longitud: info.longitud || "",
      cuerpo:   info.cuerpo   || "",
      cuerda:   info.cuerda   || "",
      acero:    info.acero    || ""
    });
  });

  // --- CALCULAR EFICIENCIAS FINALES y armar estructura para el frontend ---
  var porProceso = {};

  Object.keys(reporte).forEach(function(proc) {
    var r = reporte[proc];
    var grupos = [];
    var totalKilosProc  = 0;
    var sumaRealGlobal  = 0;
    var sumaTeoGlobal   = 0;

    // ── Para COLATADO: ajustar teórico por operador dividiendo horas ──
    // Si un operador aparece en N máquinas del mismo turno, sus horas se dividen entre N
    if (r.modo === "ROLLOS") {
      // Contar cuántas máquinas tiene cada operador por turno
      var opTurnoMaqCount = {}; // "op||turno" -> count de máquinas
      Object.keys(r.grupos).forEach(function(gKey) {
        var grp = r.grupos[gKey];
        Object.keys(grp.turnosProcesados).forEach(function(llaveTurno) {
          var opKey = grp.operador + "||" + llaveTurno;
          opTurnoMaqCount[opKey] = (opTurnoMaqCount[opKey] || 0) + 1;
        });
      });
      // Recalcular sumaTeorico y sumaTeoricoKg dividiendo por N máquinas del operador
      Object.keys(r.grupos).forEach(function(gKey) {
        var grp = r.grupos[gKey];
        var nuevoTeorico   = 0;
        var nuevoTeoricoKg = 0;
        var std = mapStd[grp.maquina] || 0;
        Object.keys(grp.turnosProcesados).forEach(function(llaveTurno) {
          var partes = llaveTurno.split("_");
          var turnoNum = partes[partes.length - 1];
          var horas = turnoNum == "2" ? 7.0 : turnoNum == "3" ? 8.0 : 7.5;
          var opKey = grp.operador + "||" + llaveTurno;
          var nMaq  = opTurnoMaqCount[opKey] || 1;
          var horasEfectivas = horas / nMaq;
          var teo = std > 0 ? std * 60 * horasEfectivas : 0;
          nuevoTeorico   += teo;
          nuevoTeoricoKg += teo; // ROLLOS: teórico ya en rollos
        });
        grp.sumaTeorico    = nuevoTeorico;
        grp.sumaTeoricoKg  = nuevoTeoricoKg;
      });
    }

    Object.keys(r.grupos).forEach(function(gKey) {
      var grp    = r.grupos[gKey];
      var effGrp = grp.sumaTeorico > 0 ? (grp.sumaReal / grp.sumaTeorico) * 100
                 : grp.sumaReal   > 0  ? 100 : 0;

      grupos.push({
        maquina:      grp.maquina,
        operador:     grp.operador,
        total:        grp.totalKilos,
        eficiencia:   effGrp,
        sumaTeorico:  grp.sumaTeorico,
        sumaTeoricoKg: grp.sumaTeoricoKg,
        detalles:     grp.detalles
      });

      totalKilosProc += grp.totalKilos;
      sumaRealGlobal += grp.sumaReal;
      sumaTeoGlobal  += grp.sumaTeorico;
    });

    grupos.sort(function(a, b){ return a.maquina.localeCompare(b.maquina); });

    var effTotal = sumaTeoGlobal > 0 ? (sumaRealGlobal / sumaTeoGlobal) * 100
                 : sumaRealGlobal > 0 ? 100 : 0;

    porProceso[proc] = {
      nombre:          proc,
      totalKilos:      totalKilosProc,
      eficienciaTotal: effTotal,
      grupos:          grupos
    };
  });

  return JSON.stringify(porProceso);
}

// ── Devuelve registros reales de PRODUCCION para fecha+turno+proceso ──
// Usado por el botón de prueba del reporte de eficiencia
function obtenerRegistrosPorFechaTurnoProceso(fecha, turno, proceso) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var sheetOrd  = ss.getSheetByName("ORDENES");

  var dataOrd  = sheetOrd.getDataRange().getValues();
  var hOrd     = dataOrd[0].map(function(x){ return String(x).toUpperCase().trim(); });
  var iProcOrd = hOrd.indexOf("PROCESO");
  var iUser    = hOrd.indexOf("USER");

  var dataProd = sheetProd.getDataRange().getValues();
  var h        = dataProd[0].map(function(x){ return String(x).toUpperCase().trim(); });

  var IDX = {
    ID:      h.indexOf("ID"),
    ORDEN:   h.indexOf("ORDEN"),
    LOTE:    h.indexOf("LOTE"),
    MAQ:     h.indexOf("MAQUINA"),
    FECHA:   h.indexOf("FECHA"),
    TURNO:   h.indexOf("TURNO"),
    OP_TXT:  h.indexOf("NOMBRE_OPERADOR_TXT"),
    PROD:    h.indexOf("PRODUCIDO"),
    USER:    h.indexOf("USER"),
    PROCESO: h.indexOf("PROCESO")
  };

  // Mapa ordenID → proceso
  var mapaProceso = {};
  if (iProcOrd >= 0) {
    for (var o = 1; o < dataOrd.length; o++) {
      var idO = String(dataOrd[o][0]);
      var prO = String(dataOrd[o][iProcOrd] || "").trim().toUpperCase();
      if (idO && prO) mapaProceso[idO] = prO;
    }
  }

  function getProceso(row) {
    if (IDX.PROCESO >= 0) {
      var p = String(row[IDX.PROCESO] || "").trim().toUpperCase();
      if (p) return p;
    }
    return mapaProceso[String(row[IDX.ORDEN] || "")] || "";
  }

  var procesoUp = String(proceso || "").toUpperCase().trim();
  var turnoStr  = String(turno || "").trim();

  var resultados = [];
  for (var i = 1; i < dataProd.length; i++) {
    var row = dataProd[i];

    // Filtro proceso
    var proc = getProceso(row);
    if (proc !== procesoUp) continue;

    // Filtro turno
    if (String(row[IDX.TURNO] || "").trim() !== turnoStr) continue;

    // Filtro fecha
    var fVal = row[IDX.FECHA];
    var fStr = "";
    if (fVal instanceof Date && !isNaN(fVal.getTime())) {
      fStr = fVal.getFullYear()
           + "-" + (fVal.getMonth()+1 < 10 ? "0" : "") + (fVal.getMonth()+1)
           + "-" + (fVal.getDate()    < 10 ? "0" : "") +  fVal.getDate();
    }
    if (fStr !== fecha) continue;

    resultados.push({
      proceso:   proc,
      maquina:   String(row[IDX.MAQ]    || ""),
      op1_txt:   String(row[IDX.OP_TXT] || ""),
      producido: Number(row[IDX.PROD])  || 0,
      turno:     turnoStr,
      fecha:     fStr,
      ordenID:   String(row[IDX.ORDEN]  || ""),
      loteFull:  String(row[IDX.LOTE]   || ""),
      usuario:   String(row[IDX.USER]   || "")
    });
  }

  return JSON.stringify(resultados);
}

// ── Verifica irregularidades de lotes y pesos para el reporte de eficiencia ──
// Recibe: registros = [{loteFull, maquina, fecha, ...}], fechaStr = "YYYY-MM-DD"
// Devuelve JSON: { sobreEficiencia: [{maquina, efic}], lotesIncompletos: [{lote, maquina, falta}], pesosIncorrectos: [{lote, maquina, detalle}] }
function verificarIrregularidadesReporte(registros, fechaStr) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var dataProd  = sheetProd.getDataRange().getValues();
  var h         = dataProd[0].map(function(x){ return String(x).toUpperCase().trim(); });

  var IDX = {
    LOTE:   h.indexOf("LOTE"),
    MAQ:    h.indexOf("MAQUINA"),
    FECHA:  h.indexOf("FECHA"),
    TURNO:  h.indexOf("TURNO"),
    PESO_I: h.indexOf("PESO_I"),
    PESO_F: h.indexOf("PESO_F"),
    PROD:   h.indexOf("PRODUCIDO"),
    PROCESO:h.indexOf("PROCESO"),
    ORDEN:  h.indexOf("ORDEN")
  };

  // ── Parsear fecha objetivo y calcular ventana histórica (5 días hábiles atrás) ──
  function parseYMD(s) {
    var p = String(s).split("-");
    return new Date(Number(p[0]), Number(p[1])-1, Number(p[2]), 0, 0, 0);
  }
  function toYMD(d) {
    if (!(d instanceof Date) || isNaN(d.getTime())) return "";
    return d.getFullYear()
      + "-" + (d.getMonth()+1 < 10 ? "0" : "") + (d.getMonth()+1)
      + "-" + (d.getDate()    < 10 ? "0" : "") +  d.getDate();
  }
  function diasHabilesAtras(fecha, n) {
    var d = new Date(fecha.getTime());
    var contados = 0;
    while (contados < n) {
      d.setDate(d.getDate() - 1);
      var dow = d.getDay();
      if (dow !== 0 && dow !== 6) contados++;
    }
    return d;
  }

  var fechaObj      = parseYMD(fechaStr);
  var fechaMinPesos = diasHabilesAtras(fechaObj, 5);   // 5 días hábiles: para verificar pesos
  var fechaMinLotes = diasHabilesAtras(fechaObj, 20);  // 20 días hábiles: para continuidad de lotes

  // histLotes: clave lote+proceso, últimos 5 días hábiles (pesos)
  // hoyLotes:  clave lote+proceso, solo hoy (pesos)
  // todosDias: clave lote+proceso, últimos 20 días hábiles + hoy (continuidad de lotes)
  var histLotes = {};
  var hoyLotes  = {};
  var todosDias = {};

  // ── Mapa ordenID → proceso (para filas de PRODUCCION sin col PROCESO) ──
  var sheetOrdV = SpreadsheetApp.openById(ID_HOJA_CALCULO).getSheetByName("ORDENES");
  var dataOrdV  = sheetOrdV.getDataRange().getValues();
  var hOrdV     = dataOrdV[0].map(function(x){ return String(x).toUpperCase().trim(); });
  var iOrdProcV = hOrdV.indexOf("PROCESO");
  var mapaProcesoV = {};
  if (iOrdProcV >= 0) {
    for (var o = 1; o < dataOrdV.length; o++) {
      var idO = String(dataOrdV[o][0]);
      var prO = String(dataOrdV[o][iOrdProcV] || "").toUpperCase().trim();
      if (idO && prO) mapaProcesoV[idO] = prO;
    }
  }

  for (var i = 1; i < dataProd.length; i++) {
    var row  = dataProd[i];
    var fVal = row[IDX.FECHA];
    if (!(fVal instanceof Date) || isNaN(fVal.getTime())) continue;
    var fDate = fVal;
    var fYMD  = toYMD(fDate);
    if (fDate < fechaMinLotes && fYMD !== fechaStr) continue; // fuera de ventana máxima
    var lote = String(row[IDX.LOTE] || "").trim();
    if (!lote) continue;
    // Resolver proceso: primero col PROCESO de la fila, luego mapa ORDENES
    var procFila = IDX.PROCESO >= 0 ? String(row[IDX.PROCESO] || "").toUpperCase().trim() : "";
    if (!procFila && IDX.ORDEN >= 0) {
      procFila = mapaProcesoV[String(row[IDX.ORDEN] || "")] || "";
    }
    var clave = lote + "||" + procFila;

    if (fYMD === fechaStr) {
      todosDias[clave] = true;
      if (!hoyLotes[clave]) hoyLotes[clave] = [];
      hoyLotes[clave].push({
        pesoI: Number(row[IDX.PESO_I]) || 0,
        pesoF: Number(row[IDX.PESO_F]) || 0
      });
    } else if (fDate < fechaObj) {
      todosDias[clave] = true;
      if (fDate >= fechaMinPesos) {
        if (!histLotes[clave]) histLotes[clave] = [];
        histLotes[clave].push({
          pesoI: Number(row[IDX.PESO_I]) || 0,
          pesoF: Number(row[IDX.PESO_F]) || 0
        });
      }
    }
  }

  // ── Helper: parsear lote "X.YYYY.ZZ" → {serie:"X.YYYY", consec:ZZ} ──
  function parseLote(lote) {
    var m = lote.match(/^(.+)\.(\d+)$/);
    if (!m) return null;
    return { serie: m[1], consec: parseInt(m[2], 10) };
  }

  var procesoReporte = registros.length > 0
    ? String(registros[0].proceso || "").toUpperCase().trim() : "";

  // ── 1. Continuidad de lotes ──
  // Regla: si existe lote X.YY.NN en el reporte, el lote X.YY.(NN-1) debe
  // existir en todosDias del mismo proceso. Si NN=1 y no existe el .00 → OK
  // (es el primer lote de la serie). Solo reportar cuando hay evidencia de
  // que la serie está activa (existe NN-2 o NN-1 es 1).
  var lotesIncompletos = [];
  var lotesVistos = {};

  registros.forEach(function(r) {
    var lote = String(r.loteFull || "").trim();
    var maq  = String(r.maquina  || "");
    if (!lote || lotesVistos[lote]) return;
    lotesVistos[lote] = true;

    var parsed = parseLote(lote);
    if (!parsed || parsed.consec <= 0) return;

    var consecAnterior = parsed.consec - 1;
    // Formato con ceros: mínimo 2 dígitos
    var pad = function(n){ return n < 10 ? "0" + n : String(n); };
    var loteAnt  = parsed.serie + "." + pad(consecAnterior);
    var claveAnt = loteAnt + "||" + procesoReporte;

    if (todosDias[claveAnt]) return; // existe → OK

    if (consecAnterior >= 1) {
      var loteAnt2  = parsed.serie + "." + pad(consecAnterior - 1);
      var claveAnt2 = loteAnt2 + "||" + procesoReporte;
      if (todosDias[claveAnt2] || consecAnterior === 1) {
        lotesIncompletos.push({ lote: lote, maquina: maq, falta: loteAnt });
      }
    }
  });

  // ── 2. Continuidad de pesos ──
  // Regla: PESO_I = 0 → inicio limpio, OK.
  //        PESO_I > 0 → debe coincidir con algún PESO_F previo del mismo lote+proceso.
  var pesosIncorrectos = [];
  var lotesVistosP = {};

  registros.forEach(function(r) {
    var lote = String(r.loteFull || "").trim();
    var maq  = String(r.maquina  || "");
    var proc = String(r.proceso  || "").toUpperCase().trim();
    if (!lote || lotesVistosP[lote]) return;
    lotesVistosP[lote] = true;

    var clave   = lote + "||" + proc;
    var hoyRegs = hoyLotes[clave];
    if (!hoyRegs || hoyRegs.length === 0) return;

    var pesoIHoy = hoyRegs.reduce(function(min, x){ return x.pesoI < min ? x.pesoI : min; }, hoyRegs[0].pesoI);
    if (pesoIHoy === 0) return; // inicio limpio → OK

    var histRegs = histLotes[clave];
    if (!histRegs || histRegs.length === 0) {
      pesosIncorrectos.push({
        lote: lote, maquina: maq,
        detalle: "P.Inicial=" + pesoIHoy + " pero no hay registro previo del lote (debería empezar en 0)"
      });
      return;
    }

    var coincide = histRegs.some(function(x){ return Math.abs(x.pesoF - pesoIHoy) < 1; });
    if (!coincide) {
      var prevStr = histRegs.map(function(x){ return "PI:" + x.pesoI + " / PF:" + x.pesoF; }).join(" | ");
      pesosIncorrectos.push({
        lote: lote, maquina: maq,
        detalle: "P.Inicial=" + pesoIHoy + " no coincide con P.Final previo (" + prevStr + ")"
      });
    }
  });

  return JSON.stringify({
    lotesIncompletos: lotesIncompletos,
    pesosIncorrectos: pesosIncorrectos
  });
}

function obtenerDatosProgramadorForja() {
  try {
    var ss    = SpreadsheetApp.openById('1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk');
    var shOrd = ss.getSheetByName('ORDENES');
    var shPed = ss.getSheetByName('PEDIDOS');
    var shInv = ss.getSheetByName('INVENTARIO_EXTERNO');
    var shRut = ss.getSheetByName('RUTAS');
    if (!shOrd || !shPed) return JSON.stringify({ error: 'Hojas no encontradas' });

    var dataOrd = shOrd.getDataRange().getValues();
    var dataPed = shPed.getDataRange().getValues();
    var dataInv = shInv ? shInv.getDataRange().getValues() : [];
    var dataRut = shRut ? shRut.getDataRange().getValues() : [];
    var tz      = Session.getScriptTimeZone();

    // ORDENES índices: [1]=PEDIDO [2]=PARTIDA [4]=SERIE [5]=ORDEN [6]=CODIGO
    // [11]=PROCESO [12]=MAQUINA [14]=PRODUCIDO [15]=ESTADO [18]=PESO [19]=TIPO
    // [20]=DIAMETRO [21]=LONGITUD [22]=CUERDA [23]=CUERPO [24]=ACERO [30]=MERMA [38]=NOTE
    // PEDIDOS índices: [1]=PEDIDO [2]=FECHA [3]=CODIGO [4]=DESCRIPCION
    // [5]=PARTIDA [6]=CANTIDAD [8]=ESTADO

    var PROC_OP_MAP = {
      'FORJA':10,'PUNTEADO':25,'ROLADO TORN':30,'LAVADO':40,'TEMPLE Y REVENIDO':50
    };

    // Normaliza nombre de pedido: quita espacios, corrige Z- → ZEQ-
    function normPed(p) {
      var s = String(p||'').trim().replace(/\s+/g,' ');
      if (/^Z-\d/.test(s)) s = 'ZEQ-' + s.substring(2);
      return s;
    }

    // ── 1. Mapa INVENTARIO_EXTERNO (excluye código negro) ──
    var mapaInv = {};
    if (dataInv.length > 1) {
      var hInv = dataInv[0];
      var iiCod=0,iiExt=1,iiMin=2,iiMax=3,iiBack=-1;
      for (var h=0;h<hInv.length;h++) {
        var hN=String(hInv[h]).toUpperCase().trim();
        if(hN==='CODIGO')     iiCod=h;
        if(hN==='EXISTENCIA') iiExt=h;
        if(hN==='MINIMO')     iiMin=h;
        if(hN==='MAXIMO')     iiMax=h;
        if(hN==='BACKORDER')  iiBack=h;
      }
      for (var iv=1;iv<dataInv.length;iv++) {
        var cInv=String(dataInv[iv][iiCod]||'').trim();
        if(!cInv||cInv.charAt(0)==='7') continue;
        mapaInv[cInv]={
          existencia:Number(dataInv[iv][iiExt])||0,
          minimo:    Number(dataInv[iv][iiMin])||0,
          maximo:    Number(dataInv[iv][iiMax])||0,
          backorder: iiBack>-1?Number(dataInv[iv][iiBack])||0:0
        };
      }
    }

    // ── 2. Mapa RUTAS: código → datos de fila FORJA + máquinas permitidas por proceso ──
    var mapaRutas    = {};  // código → { maquina, desc, tipo, ... , maquinasPorProceso:{PROC:'m1, m2'} }
    var codigosForja = {};
    if (dataRut.length > 1) {
      var hRut  = dataRut[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var rCOD  = hRut.indexOf('CODIGO');
      var rPROC = hRut.indexOf('PROCESO');
      var rMAQ  = hRut.indexOf('MAQUINA');
      var rDESC = hRut.indexOf('DESCRIPCION');
      var rTIPO = hRut.indexOf('TIPO');
      var rDIA  = hRut.indexOf('DIAMETRO');
      var rLONG = hRut.indexOf('LONGITUD');
      var rCDA  = hRut.indexOf('CUERDA');
      var rCPO  = hRut.indexOf('CUERPO');
      var rACE  = hRut.indexOf('ACERO');
      var rPESO = hRut.indexOf('PESO');
      for (var r=1;r<dataRut.length;r++) {
        var rCodV  = String(dataRut[r][rCOD] ||'').trim();
        var rProcV = String(dataRut[r][rPROC]||'').trim().toUpperCase();
        if(!rCodV||rCodV.charAt(0)==='7') continue;
        if(rProcV==='FORJA') codigosForja[rCodV]=true;
        // Acumular máquinas permitidas por proceso
        var maqRaw = rMAQ>-1 ? String(dataRut[r][rMAQ]||'').trim() : '';
        if(!mapaRutas[rCodV]) {
          mapaRutas[rCodV] = {
            maquina:           maqRaw.split(',')[0].trim(),
            desc:              rDESC>-1?String(dataRut[r][rDESC]||'').trim():'',
            tipo:              rTIPO>-1?String(dataRut[r][rTIPO]||'').trim():'',
            diametro:          rDIA >-1?String(dataRut[r][rDIA] ||'').trim():'',
            longitud:          rLONG>-1?String(dataRut[r][rLONG]||'').trim():'',
            cuerda:            rCDA >-1?String(dataRut[r][rCDA] ||'').trim():'',
            cuerpo:            rCPO >-1?String(dataRut[r][rCPO] ||'').trim():'',
            acero:             rACE >-1?String(dataRut[r][rACE] ||'').trim():'',
            peso:              rPESO>-1?Number(dataRut[r][rPESO])||0:0,
            _forjaSet:         rProcV==='FORJA',
            maquinasPorProceso:{}
          };
        } else if (!mapaRutas[rCodV]._forjaSet && rProcV==='FORJA') {
          // Promover a fila FORJA si aún no se había hecho
          mapaRutas[rCodV].maquina   = maqRaw.split(',')[0].trim();
          mapaRutas[rCodV]._forjaSet = true;
        }
        // Acumular máquinas del proceso en maquinasPorProceso
        if(maqRaw) {
          var mpp = mapaRutas[rCodV].maquinasPorProceso;
          if(!mpp[rProcV]) mpp[rProcV] = [];
          maqRaw.split(',').forEach(function(m){
            var mt=m.trim();
            if(mt && mpp[rProcV].indexOf(mt)<0) mpp[rProcV].push(mt);
          });
        }
      }
    }

    // ── 3. Mapa PEDIDOS: vivos y último por código ──
    var mapPedVivo  = {};  // nomPedNorm + '_' + partida → info
    var mapUltimoPed = {}; // codigo → info del pedido más reciente

    for(var p=1;p<dataPed.length;p++) {
      var rp = dataPed[p];
      if(!rp[1]||!rp[3]) continue;
      var nomPed  = normPed(rp[1]);
      var codPed  = String(rp[3]||'').trim();
      if(codPed.charAt(0)==='7') continue;
      var estPed  = String(rp[8]||'').trim().toUpperCase();
      var fechaPed= rp[2] instanceof Date ? rp[2] : null;
      var fechaStr= fechaPed
        ? Utilities.formatDate(fechaPed,tz,'dd/MM/yyyy')
        : (rp[2]?String(rp[2]).trim():'');

      // Último pedido por código (cualquier estado)
      if(!mapUltimoPed[codPed]||
        (fechaPed&&(!mapUltimoPed[codPed].fechaRaw||fechaPed>mapUltimoPed[codPed].fechaRaw))) {
        mapUltimoPed[codPed]={
          pedidoRaw:   nomPed,
          fecha:       fechaStr,
          fechaRaw:    fechaPed,
          descripcion: String(rp[4]||'').trim(),
          partida:     String(rp[5]||'').trim(),
          cantidad:    Number(rp[6])||0
        };
      }

      // Pedidos vivos
      if(estPed!=='TERMINADO'&&estPed!=='CANCELADO') {
        var keyP = nomPed + '_' + Math.trunc(rp[5]);
        mapPedVivo[keyP]={
          pedidoRaw:   nomPed,
          fecha:       fechaStr,
          codigo:      codPed,
          descripcion: String(rp[4]||'').trim(),
          partida:     String(rp[5]||'').trim(),
          cantidad:    Number(rp[6])||0
        };
      }
    }

    // ── 4. Recorrer ORDENES SERIE=F con pedido vivo ──
    var resultado = [];
    var vistos    = {};
    var codigosConPedidoVivo = {};

    for(var i=1;i<dataOrd.length;i++) {
      var ro = dataOrd[i];
      if(String(ro[4]||'').trim().toUpperCase()!=='F') continue;

      var ordenNum  = String(ro[5]||'').trim();
      var nomPedOrd = normPed(ro[1]);  // ← normalizar igual que en mapPedVivo
      var partida   = Math.trunc(ro[2]);
      var proceso   = String(ro[11]||'').trim().toUpperCase();
      if(!ordenNum||!nomPedOrd) continue;

      // Buscar en pedidos vivos: exacto primero, luego solo por nombre
      var keyExacto = nomPedOrd + '_' + partida;
      var pedInfo   = mapPedVivo[keyExacto];
      if(!pedInfo) {
        // Fallback: cualquier partida del mismo pedido
        var prefixBusq = nomPedOrd + '_';
        for(var kk in mapPedVivo) {
          if(kk.indexOf(prefixBusq)===0){ pedInfo=mapPedVivo[kk]; break; }
        }
      }
      if(!pedInfo) continue;

      var operacion = PROC_OP_MAP[proceso]||0;
      var opInt     = ordenNum + '.' + operacion;
      if(vistos[opInt]) continue;
      vistos[opInt] = true;

      var codOrden = String(ro[6]||pedInfo.codigo||'').trim();
      codigosConPedidoVivo[codOrden] = true;

      // Separar tipo y número del pedido normalizado
      var pedRaw = pedInfo.pedidoRaw;
      var guion  = pedRaw.indexOf('-');
      var tipoPed= guion>-1 ? pedRaw.substring(0,guion)  : pedRaw.substring(0,3);
      var numPed = guion>-1 ? pedRaw.substring(guion+1)  : pedRaw;

      var invData = mapaInv[pedInfo.codigo] || null;
      var rut     = mapaRutas[pedInfo.codigo] || {};

      resultado.push({
        id_orden:      String(ro[0]||'').trim(),
        orden:         ordenNum,
        num_orden:     String(ro[5]||'').trim(),
        nom_pedido:    normPed(ro[1]),
        operacion:     operacion,
        op_int:        opInt,
        fecha:         pedInfo.fecha,
        tipo:          tipoPed,
        pedido:        numPed,
        partida:       pedInfo.partida,
        codigo:        pedInfo.codigo,
        cantidad:      pedInfo.cantidad,
        backorder:     invData ? invData.backorder  : null,
        minimo:        invData ? invData.minimo     : null,
        maximo:        invData ? invData.maximo     : null,
        existencia:    invData ? invData.existencia : null,
        descripcion:   pedInfo.descripcion,
        maquina:            String(ro[12]||'').trim() || rut.maquina || '',
        maquinasPermitidas: (rut.maquinasPorProceso && rut.maquinasPorProceso[proceso])
                              ? rut.maquinasPorProceso[proceso].join(', ')
                              : (String(ro[12]||'').trim() || rut.maquina || ''),
        desc_tipo:     String(ro[19]||'').trim() || rut.tipo     || '',
        diametro:      String(ro[20]||'').trim() || rut.diametro || '',
        longitud:      String(ro[21]||'').trim() || rut.longitud || '',
        cuerda:        String(ro[22]||'').trim() || rut.cuerda   || '',
        cuerpo:        String(ro[23]||'').trim() || rut.cuerpo   || '',
        acero:         String(ro[24]||'').trim() || rut.acero    || '',
        peso_pza:      Number(ro[18]) || rut.peso || 0,
        producido:     Number(ro[14]) || 0,
        peso_merma:    Number(ro.length > 30 ? ro[30]||0 : 0),
        mp_especif:    String(ro.length>29 ? ro[29]||'' : '').trim(),
        observaciones: String(ro.length>38 ? ro[38]||'' : '').trim(),
        fila_hoja:     i + 1,
        sin_pedido_vivo: false
      });
    }

    // ── 5. Códigos FORJA sin pedido vivo (solo los que están en INVENTARIO_EXTERNO) ──
    // Regla: sin pedido vivo → DEBE existir en mapaInv (INVENTARIO_EXTERNO, no empieza con 7)
    Object.keys(codigosForja).forEach(function(cod) {
      if(codigosConPedidoVivo[cod]) return;
      var inv    = mapaInv[cod] || null;
      if(!inv) return;  // ← excluir si no aparece en INVENTARIO_EXTERNO
      var ultPed = mapUltimoPed[cod]||null;
      var rut    = mapaRutas[cod]||{};

      var pedTipo='',pedNum='SIN PEDIDO',pedFecha='',pedDesc='',pedPart='',pedCant=0;
      if(ultPed) {
        var pr = ultPed.pedidoRaw;
        var g  = pr.indexOf('-');
        pedTipo  = g>-1?pr.substring(0,g):pr.substring(0,3);
        pedNum   = g>-1?pr.substring(g+1):pr;
        pedFecha = ultPed.fecha;
        pedDesc  = ultPed.descripcion || rut.desc || '';
        pedPart  = ultPed.partida;
        pedCant  = ultPed.cantidad;
      }

      resultado.push({
        id_orden:       '',
        orden:          '—',
        operacion:      10,
        op_int:         '—',
        fecha:          pedFecha,
        tipo:           pedTipo,
        pedido:         pedNum,
        partida:        pedPart,
        codigo:         cod,
        cantidad:       pedCant,
        backorder:      inv ? inv.backorder  : null,
        minimo:         inv ? inv.minimo     : null,
        maximo:         inv ? inv.maximo     : null,
        existencia:     inv ? inv.existencia : null,
        descripcion:    pedDesc || rut.desc || '',
        maquina:        rut.maquina  ||'',
        desc_tipo:      rut.tipo     ||'',
        diametro:       rut.diametro ||'',
        longitud:       rut.longitud ||'',
        cuerda:         rut.cuerda   ||'',
        cuerpo:         rut.cuerpo   ||'',
        acero:          rut.acero    ||'',
        peso_pza:       rut.peso     ||0,
        producido:      0,
        peso_merma:     0,
        mp_especif:     '',
        observaciones:  '',
        fila_hoja:      -1,
        sin_pedido_vivo: true
      });
    });

    // Ordenar: vivos primero por número de orden, sin pedido al final por código
    resultado.sort(function(a,b) {
      if(a.sin_pedido_vivo!==b.sin_pedido_vivo) return a.sin_pedido_vivo?1:-1;
      var na=parseInt(a.orden)||0, nb=parseInt(b.orden)||0;
      if(na!==nb) return na-nb;
      return (a.operacion||0)-(b.operacion||0);
    });

    return JSON.stringify(resultado);
  } catch(e) {
    return JSON.stringify({error:e.message+'|'+e.stack});
  }
}

// ── Guarda un campo desde la vista Forja ──
// Campos por fila individual: maquina, acero, peso_pza, observaciones
// Campos por orden completa (todos los procesos): desc_tipo, diametro, longitud, cuerda, cuerpo, mp_especif
// Campos especiales: fecha → actualiza PEDIDOS col B (índice 2)
function fprogGuardarCampo(payload) {
  try {
    var ss    = SpreadsheetApp.openById('1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk');
    var shOrd = ss.getSheetByName('ORDENES');
    if (!shOrd) return JSON.stringify({ ok:false, msg:'Hoja ORDENES no encontrada' });

    var fila      = parseInt(payload.fila_hoja);
    var campo     = String(payload.campo||'').trim();
    var valor     = payload.valor;
    var numOrden  = String(payload.num_orden||'').trim();
    var nomPedido = String(payload.nom_pedido||'').trim();
    var partida   = String(payload.partida||'').trim();

    // ── Campos individuales por fila ──
    var CAMPO_COL_INDIVIDUAL = {
      'maquina':13, 'acero':25, 'peso_pza':19, 'observaciones':39, 'cantidad':9
    };

    // ── Campos que se actualizan en TODOS los procesos de la misma orden ──
    var CAMPO_COL_ORDEN = {
      'desc_tipo':20, 'diametro':21, 'longitud':22, 'cuerda':23, 'cuerpo':24, 'mp_especif':30
    };

    // ── FECHA: actualiza PEDIDOS ──
    if (campo === 'fecha') {
      if (!nomPedido) return JSON.stringify({ ok:false, msg:'Falta nom_pedido para actualizar fecha' });
      var shPed = ss.getSheetByName('PEDIDOS');
      if (!shPed) return JSON.stringify({ ok:false, msg:'Hoja PEDIDOS no encontrada' });
      // Parsear fecha dd/mm/aaaa → objeto Date
      var partes = String(valor).trim().split('/');
      if (partes.length !== 3) return JSON.stringify({ ok:false, msg:'Formato de fecha inválido, use dd/mm/aaaa' });
      var d = parseInt(partes[0]), m = parseInt(partes[1])-1, y = parseInt(partes[2]);
      if (isNaN(d)||isNaN(m)||isNaN(y)) return JSON.stringify({ ok:false, msg:'Fecha inválida' });
      var fechaDate = new Date(y, m, d);
      var dataPed = shPed.getDataRange().getValues();
      var actualizados = 0;
      for (var p = 1; p < dataPed.length; p++) {
        var nomFila    = String(dataPed[p][1]||'').trim();
        var partidaFila= String(dataPed[p][5]||'').trim();
        // Normalizar: Z- → ZEQ-
        if (/^Z-\d/.test(nomFila)) nomFila = 'ZEQ-' + nomFila.substring(2);
        if (nomFila === nomPedido && (partida === '' || partidaFila === partida)) {
          shPed.getRange(p+1, 3).setValue(fechaDate); // col C (índice 2 base-0 = col 3 base-1)
          actualizados++;
        }
      }
      if (actualizados === 0) return JSON.stringify({ ok:false, msg:'No se encontró el pedido: ' + nomPedido });
      return JSON.stringify({ ok:true, actualizados:actualizados });
    }

    // ── Campos individuales ──
    if (CAMPO_COL_INDIVIDUAL[campo]) {
      if (!fila || fila < 2) return JSON.stringify({ ok:false, msg:'Fila inválida' });
      shOrd.getRange(fila, CAMPO_COL_INDIVIDUAL[campo]).setValue(valor);
      return JSON.stringify({ ok:true });
    }

    // ── Campos por orden completa ──
    if (CAMPO_COL_ORDEN[campo]) {
      if (!numOrden) return JSON.stringify({ ok:false, msg:'Falta num_orden para actualizar por orden' });
      var colIdx  = CAMPO_COL_ORDEN[campo]; // base-1
      var dataOrd = shOrd.getDataRange().getValues();
      var hOrd    = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var iOrden  = hOrd.indexOf('ORDEN'); // col de número de orden
      if (iOrden < 0) iOrden = 5; // fallback índice 5 (col F)
      var filasActualizadas = [];
      for (var i = 1; i < dataOrd.length; i++) {
        if (String(dataOrd[i][iOrden]||'').trim() === numOrden) {
          filasActualizadas.push(i + 1);
        }
      }
      if (filasActualizadas.length === 0) return JSON.stringify({ ok:false, msg:'No se encontraron filas para orden: ' + numOrden });
      filasActualizadas.forEach(function(f) {
        shOrd.getRange(f, colIdx).setValue(valor);
      });
      return JSON.stringify({ ok:true, filasActualizadas: filasActualizadas.length });
    }

    return JSON.stringify({ ok:false, msg:'Campo no soportado: ' + campo });
  } catch(e) {
    return JSON.stringify({ ok:false, msg:e.message });
  }
}

// ── Genérica: máquinas permitidas en RUTAS para cualquier código + proceso ──
function obtenerMaquinasPorCodigoYProceso(codigo, proceso) {
  try {
    var ss   = SpreadsheetApp.openById('1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk');
    var shRut = ss.getSheetByName('RUTAS');
    if (!shRut) return JSON.stringify({ maquinas: [] });
    var data = shRut.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ maquinas: [] });
    var hdr   = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iCod  = hdr.indexOf('CODIGO');
    var iProc = hdr.indexOf('PROCESO');
    var iMaq  = hdr.indexOf('MAQUINA');
    if (iCod < 0 || iProc < 0 || iMaq < 0) return JSON.stringify({ maquinas: [] });
    var procUp = String(proceso).trim().toUpperCase();
    var maquinas = [];
    for (var i = 1; i < data.length; i++) {
      var cod  = String(data[i][iCod]  || '').trim();
      var proc = String(data[i][iProc] || '').trim().toUpperCase();
      var maq  = String(data[i][iMaq]  || '').trim();
      if (cod !== codigo || proc !== procUp || !maq) continue;
      maq.split(',').forEach(function(m) {
        var mt = m.trim();
        if (mt && maquinas.indexOf(mt) < 0) maquinas.push(mt);
      });
    }
    return JSON.stringify({ maquinas: maquinas });
  } catch(e) {
    return JSON.stringify({ maquinas: [], error: e.message });
  }
}

// ── Obtiene máquinas permitidas en RUTAS para un código con proceso FORJA ──
function fprogObtenerMaquinasForja(codigo) {
  try {
    var ss   = SpreadsheetApp.openById('1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk');
    var shRut = ss.getSheetByName('RUTAS');
    if (!shRut) return JSON.stringify({ maquinas: [] });

    var data = shRut.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ maquinas: [] });

    var hdr   = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iCod  = hdr.indexOf('CODIGO');
    var iProc = hdr.indexOf('PROCESO');
    var iMaq  = hdr.indexOf('MAQUINA');
    if (iCod < 0 || iProc < 0 || iMaq < 0) return JSON.stringify({ maquinas: [] });

    var maquinas = [];
    for (var i = 1; i < data.length; i++) {
      var cod  = String(data[i][iCod]  || '').trim();
      var proc = String(data[i][iProc] || '').trim().toUpperCase();
      var maq  = String(data[i][iMaq]  || '').trim();
      if (cod !== codigo || proc !== 'FORJA' || !maq) continue;
      // La columna MAQUINA puede tener múltiples separadas por coma
      maq.split(',').forEach(function(m) {
        var mt = m.trim();
        if (mt && maquinas.indexOf(mt) < 0) maquinas.push(mt);
      });
    }
    return JSON.stringify({ maquinas: maquinas });
  } catch(e) {
    return JSON.stringify({ maquinas: [], error: e.message });
  }
}

// ── Plan Diario FORJA: grupos y máquinas desde ESTANDARES ──
function fpd_obtenerGruposMaquinas() {
  try {
    var ss  = SpreadsheetApp.openById('1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk');
    var shE = ss.getSheetByName('ESTANDARES');
    if (!shE) return JSON.stringify({ grupos: {} });
    var data = shE.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ grupos: {} });
    var hdr  = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iProc  = hdr.indexOf('PROCESO');   // col C = índice 2
    var iMaq   = hdr.indexOf('MAQUINA');   // col D = índice 3
    var iGrupo = hdr.indexOf('GRUPO');     // col J = índice 9
    if (iMaq < 0 || iGrupo < 0) return JSON.stringify({ grupos: {} });
    // grupos = { nombreGrupo: [maq1, maq2, ...] }
    var grupos = {};
    var ordenGrupos = [];
    for (var i = 1; i < data.length; i++) {
      var maq   = String(data[i][iMaq]   || '').trim();
      var grupo = String(data[i][iGrupo] || '').trim();
      var proc  = iProc >= 0 ? String(data[i][iProc] || '').trim().toUpperCase() : '';
      if (!maq || !grupo) continue;
      // Solo procesos que FORJA maneja
      var procsFORJA = ['FORJA','PUNTEADO','ROLADO TORN','LAVADO','TEMPLE Y REVENIDO'];
      if (proc && procsFORJA.indexOf(proc) < 0) continue;
      if (!grupos[grupo]) { grupos[grupo] = []; ordenGrupos.push(grupo); }
      if (grupos[grupo].indexOf(maq) < 0) grupos[grupo].push(maq);
    }
    return JSON.stringify({ grupos: grupos, orden: ordenGrupos });
  } catch(e) {
    return JSON.stringify({ grupos: {}, error: e.message });
  }
}

// ══════════════════════════════════════════════════════════════════
// TRACKING — PLAN COLATADO
// Devuelve ordenes COLATADO agrupadas por máquina con fechas calculadas
// en ROL, con turnos por día y saltando fin de semana (Sáb 22:00 → Lun 06:30)
// ══════════════════════════════════════════════════════════════════
function trkGetPlanColatado() {
  try {
    var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shOrd    = ss.getSheetByName('ORDENES');
    var shEst    = ss.getSheetByName('ESTANDARES');
    var dataOrd  = shOrd.getDataRange().getValues();
    var dataEst  = shEst.getDataRange().getValues();
    var tz       = ss.getSpreadsheetTimeZone();

    // ── Encabezados ORDENES ──
    var hOrd = dataOrd[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var oID   = hOrd.indexOf('ID');
    var oPED  = hOrd.indexOf('PEDIDO');
    var oORD  = hOrd.indexOf('ORDEN');
    var oSERIE= hOrd.indexOf('SERIE');
    var oCOD  = hOrd.indexOf('CODIGO');
    var oDESC = hOrd.indexOf('DESCRIPCION');
    var oSOL  = hOrd.indexOf('SOLICITADO');
    var oPROD = hOrd.indexOf('PRODUCIDO');
    var oMAQ  = hOrd.indexOf('MAQUINA');
    var oPROC = hOrd.indexOf('PROCESO');
    var oEST  = hOrd.indexOf('ESTADO');
    var oPRIO = hOrd.indexOf('PRIORIDAD');
    var oFI     = hOrd.indexOf('FECHA_INICIO_PROG');
    var oFF     = hOrd.indexOf('FECHA_FIN_PROG');
    var oTIPO   = hOrd.indexOf('TIPO');
    var oDIA    = hOrd.indexOf('DIAMETRO');
    var oLONG   = hOrd.indexOf('LONGITUD');
    var oCUERDA = hOrd.indexOf('CUERDA');
    var oCUERPO = hOrd.indexOf('CUERPO');

    // ── Encabezados ESTANDARES ──
    var hEst  = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var ePROC  = hEst.indexOf('PROCESO');
    var eMAQ   = hEst.indexOf('MAQUINA');
    var eVEL   = hEst.indexOf('VELOCIDAD');
    var eEFIC  = hEst.indexOf('EFICIENCIA');
    var eTURNOS= hEst.indexOf('TURNOS');

    // Mapa maquina → { vel, efic, hrsXdia } solo para COLATADO
    // TURNOS: 1=7.5h/día, 2=14.5h/día, 3=22.5h/día
    var mapaEst = {};
    for (var e = 1; e < dataEst.length; e++) {
      var proc = String(dataEst[e][ePROC] || '').toUpperCase().trim();
      if (proc !== 'COLATADO') continue;
      var maq  = String(dataEst[e][eMAQ]  || '').trim().toUpperCase();
      if (!maq) continue;
      var turnos = parseInt(dataEst[e][eTURNOS]) || 1;
      var hrsXdia = turnos === 3 ? 22.5 : turnos === 2 ? 14.5 : 7.5;
      mapaEst[maq] = {
        vel:     parseFloat(dataEst[e][eVEL])  || 0,
        efic:    parseFloat(dataEst[e][eEFIC]) || 1,
        hrsXdia: hrsXdia
      };
    }

    // ── Fecha inicial según hora de ejecución ──
    var ahora = new Date();
    var hora  = ahora.getHours() + ahora.getMinutes() / 60;
    var fechaBase;
    if (hora < 15) {
      fechaBase = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate(), 6, 30, 0);
    } else if (hora < 24) {
      fechaBase = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate(), 14, 30, 0);
    } else {
      var ayer = new Date(ahora);
      ayer.setDate(ayer.getDate() - 1);
      fechaBase = new Date(ayer.getFullYear(), ayer.getMonth(), ayer.getDate(), 22, 0, 0);
    }

    // ── Helper: avanzar un Date saltando el fin de semana (Sáb 22:00 → Lun 06:30) ──
    function saltarFinSemana(d) {
      var result = new Date(d);
      var maxIter = 100;
      while (maxIter-- > 0) {
        var dow  = result.getDay();
        var hh   = result.getHours() + result.getMinutes() / 60 + result.getSeconds() / 3600;

        var bloqueado = false;
        if (dow === 6 && hh >= 22) bloqueado = true;
        if (dow === 0) bloqueado = true;
        if (dow === 1 && hh < 6.5) bloqueado = true;

        if (!bloqueado) break;

        var lunesTarget = new Date(result);
        var diasHastaLunes = (8 - dow) % 7;
        if (dow === 1) diasHastaLunes = 0;
        if (diasHastaLunes === 0 && hh < 6.5) {
          lunesTarget.setHours(6, 30, 0, 0);
        } else {
          lunesTarget.setDate(lunesTarget.getDate() + diasHastaLunes);
          lunesTarget.setHours(6, 30, 0, 0);
        }
        result = lunesTarget;
      }
      return result;
    }

    // ── Helper: sumar horas de máquina a una fecha, convirtiendo a tiempo de reloj ──
    // Lógica: horas_maquina → días_maquina → días_reloj → horas_reloj
    // Cada día de máquina = hrsXdia horas trabajadas = 24h de reloj
    // Por tanto: horas_reloj = horas_maquina * (24 / hrsXdia)
    // Luego se suman esas horas_reloj al cursor, saltando fin de semana bloque a bloque.
    // ── Helper: sumar horas de máquina a una fecha, respetando fin de semana ──
    // Trabaja siempre en horas de MÁQUINA. Para avanzar el reloj convierte el bloque
    // disponible de reloj a horas de máquina: maq_disponible = reloj_disponible * (hrsXdia/24)
    // Al mover el cursor en el tiempo usa horas de reloj: reloj_avanzado = maq_consumidas * (24/hrsXdia)
    function sumarHorasLaborables(fechaInicio, horasMaquina, hrsXdia) {
      if (horasMaquina <= 0) return new Date(fechaInicio);
      var MS_HR      = 3600000;
      var factorMaq  = hrsXdia > 0 ? hrsXdia / 24 : 1; // horas maquina por hora reloj
      var cursor     = saltarFinSemana(new Date(fechaInicio));
      var restMaq    = horasMaquina; // horas de máquina pendientes

      var iter = 0;
      while (restMaq > 0.0001 && iter < 500) {
        iter++;
        var d  = cursor.getDay();
        var hh = cursor.getHours() + cursor.getMinutes() / 60 + cursor.getSeconds() / 3600;

        // Seguridad: si cursor cayó en zona bloqueada, saltar
        if (d === 6 && hh >= 22) { cursor = saltarFinSemana(cursor); continue; }
        if (d === 0)              { cursor = saltarFinSemana(cursor); continue; }
        if (d === 1 && hh < 6.5) { cursor = saltarFinSemana(cursor); continue; }

        // Calcular horas de RELOJ disponibles hasta el próximo bloqueo (Sáb 22:00)
        var sabTarget = new Date(cursor);
        var diasHastaSab = d === 6 ? 0 : (6 - d + 7) % 7 || 7;
        sabTarget.setDate(sabTarget.getDate() + diasHastaSab);
        sabTarget.setHours(22, 0, 0, 0);
        var horasRelojHastaSab = (sabTarget - cursor) / MS_HR; // horas reloj hasta corte

        // Convertir ese bloque de reloj a horas de máquina equivalentes
        var horasMaqEnBloque = horasRelojHastaSab * factorMaq;

        if (restMaq <= horasMaqEnBloque) {
          // Las horas restantes caben en este bloque antes del fin de semana
          // Convertir horas de máquina restantes a horas de reloj para mover el cursor
          var horasRelojAvanzar = restMaq / factorMaq;
          cursor = new Date(cursor.getTime() + horasRelojAvanzar * MS_HR);
          restMaq = 0;
        } else {
          // Consumir todo el bloque disponible y saltar al lunes 06:30
          restMaq -= horasMaqEnBloque;
          // Calcular lunes siguiente
          var lunesTarget = new Date(cursor);
          var diasHastaLun = (8 - d) % 7;
          if (diasHastaLun === 0) diasHastaLun = 7;
          lunesTarget.setDate(lunesTarget.getDate() + diasHastaLun);
          lunesTarget.setHours(6, 30, 0, 0);
          cursor = lunesTarget;
        }
      }
      return cursor;
    }

    // ── Recoger filas de COLATADO activas ──
    var filas = [];
    for (var i = 1; i < dataOrd.length; i++) {
      var row  = dataOrd[i];
      var proc = String(row[oPROC] || '').toUpperCase().trim();
      if (proc !== 'COLATADO') continue;
      var est  = String(row[oEST]  || '').toUpperCase().trim();
      if (est === 'TERMINADO' || est === 'CANCELADO') continue;

      var sol  = parseFloat(row[oSOL])  || 0;
      var prod = parseFloat(row[oPROD]) || 0;
      var prio = parseInt(row[oPRIO])   || 999;
      var maq  = String(row[oMAQ] || '').trim().toUpperCase();

      // COLATADO: SOLICITADO y PRODUCIDO ya están en ROLLOS directamente
      var pendiente = Math.max(sol - prod, 0);

      // Calcular horas: velocidad en ESTANDARES está en rol/min para COLATADO
      // horas = pendiente(rol) / (vel(rol/min) * efic * 60)
      var horas = 0;
      var std   = mapaEst[maq];
      if (std && std.vel > 0 && pendiente > 0) {
        horas = pendiente / (std.vel * std.efic * 60);
      }

      filas.push({
        rowIdx:     i,
        id:         String(row[oID]   || ''),
        pedido:     String(row[oPED]  || ''),
        orden:      String(row[oORD]  || ''),
        serie:      String(row[oSERIE]|| ''),
        codigo:     String(row[oCOD]  || ''),
        desc:       String(row[oDESC] || ''),
        sol:        sol,
        prod:       prod,
        pendiente:  pendiente,
        maq:        maq,
        prio:       prio,
        est:        est,
        horas:      horas,
        tipo:       String(row[oTIPO] || ''),
        dia:        String(row[oDIA]  || ''),
        long:       String(row[oLONG] || ''),
        cuerda:     String(row[oCUERDA] || ''),
        cuerpo:     String(row[oCUERPO] || '')
      });
    }

    // ── Agrupar por máquina y ordenar por prioridad ──
    var grupos = {};
    filas.forEach(function(f) {
      if (!grupos[f.maq]) grupos[f.maq] = [];
      grupos[f.maq].push(f);
    });
    Object.keys(grupos).forEach(function(maq) {
      grupos[maq].sort(function(a, b) { return a.prio - b.prio; });
    });

    // ── Calcular fechas en cascada por máquina, respetando turnos y fin de semana ──
    var resultado = [];

    Object.keys(grupos).sort().forEach(function(maq) {
      var lista   = grupos[maq];
      var std     = mapaEst[maq] || { vel: 0, efic: 1, hrsXdia: 7.5 };
      var cursor  = saltarFinSemana(new Date(fechaBase));

      var ordenesConFechas = [];
      lista.forEach(function(f, idxF) {
        var fi = new Date(cursor);
        var ff;
        if (f.pendiente <= 0) {
          ff = new Date(cursor);
        } else {
          ff = sumarHorasLaborables(fi, f.horas, std.hrsXdia);
          // Comparar longitud de esta orden con la SIGUIENTE para calcular tiempo de cambio
          var siguiente = lista[idxF + 1];
          var horasMuertas = 2; // default: misma longitud
          if (siguiente) {
            var cambioLong = String(f.long).trim() !== String(siguiente.long).trim();
            horasMuertas = cambioLong ? 6 : 2;
          }
          cursor = new Date(ff.getTime() + horasMuertas * 3600000);
          cursor = saltarFinSemana(cursor);
        }

        var fiStr = Utilities.formatDate(fi, tz, 'dd/MM/yyyy HH:mm');
        var ffStr = Utilities.formatDate(ff, tz, 'dd/MM/yyyy HH:mm');
        if (oFI >= 0) shOrd.getRange(f.rowIdx + 1, oFI + 1).setValue(fiStr);
        if (oFF >= 0) shOrd.getRange(f.rowIdx + 1, oFF + 1).setValue(ffStr);

        ordenesConFechas.push({
          id:          f.id,
          pedido:      f.pedido,
          orden:       f.orden,
          serie:       f.serie,
          codigo:      f.codigo,
          desc:        f.desc,
          sol:         Math.round(f.sol),
          prod:        Math.round(f.prod),
          pendiente:   Math.round(f.pendiente),
          prio:        f.prio,
          est:         f.est,
          horas:       Math.round(f.horas * 100) / 100,
          fechaInicio: fiStr,
          fechaFin:    ffStr,
          tipo:        f.tipo,
          dia:         f.dia,
          long:        f.long,
          cuerda:      f.cuerda,
          cuerpo:      f.cuerpo
        });
      });

      resultado.push({ maquina: maq, ordenes: ordenesConFechas });
    });

    return JSON.stringify({ success: true, grupos: resultado });
  } catch(err) {
    return JSON.stringify({ success: false, msg: err.message });
  }
}

// ══════════════════════════════════════════════════════════════════
// TRACKING — MATERIAL ENVIADO (tabla ENVIADO completa filtrable)
// Columnas: B=SEM, C=FECHA, F=PEDIDO, G=CODIGO, H=DESCRIPCION,
//           I=FAMILIA, J=KILOS, K=PIEZAS, L=COMENTARIOS, M=ENVIO,
//           P=CHOFER, Q=URL
// ══════════════════════════════════════════════════════════════════
function trkGetMaterialEnviado(anio, mes) {
  try {
    var ss     = SpreadsheetApp.openById(ID_HOJA_CALCULO);
    var shEnv  = ss.getSheetByName('ENVIADO');
    if (!shEnv) return JSON.stringify({ success: false, msg: 'Hoja ENVIADO no encontrada' });
    var data   = shEnv.getDataRange().getValues();
    var tz     = ss.getSpreadsheetTimeZone();
    // Si no se pasa mes/año usar el mes actual
    var hoy    = new Date();
    var filtAnio = anio ? parseInt(anio) : hoy.getFullYear();
    var filtMes  = mes  ? parseInt(mes)  : hoy.getMonth() + 1; // 1-12
    var filas  = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var fechaRaw = r[2];
      if (!fechaRaw) continue;
      var fechaObj = (fechaRaw instanceof Date) ? fechaRaw : new Date(fechaRaw);
      if (isNaN(fechaObj)) continue;
      // Filtrar por mes/año
      if (fechaObj.getFullYear() !== filtAnio || (fechaObj.getMonth() + 1) !== filtMes) continue;
      var fechaTxt = Utilities.formatDate(fechaObj, tz, 'dd/MM/yyyy');
      filas.push({
        sem:          String(r[1]  || ''),
        fecha:        fechaTxt,
        pedido:       String(r[5]  || ''),
        codigo:       String(r[6]  || ''),
        descripcion:  String(r[7]  || ''),
        familia:      String(r[8]  || ''),
        kilos:        parseFloat(r[9])  || 0,
        piezas:       parseFloat(r[10]) || 0,
        comentarios:  String(r[11] || ''),
        envio:        String(r[12] || ''),
        chofer:       String(r[15] || ''),
        url:          String(r[16] || '')
      });
    }
    return JSON.stringify({ success: true, filas: filas, anio: filtAnio, mes: filtMes });
  } catch(err) {
    return JSON.stringify({ success: false, msg: err.message });
  }
}
