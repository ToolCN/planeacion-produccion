import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
content = open('C:/Users/rcdga/planeacion-produccion/Mod_Planificador.html', 'r', encoding='utf-8').read()
changes = []

# ═══════════════════════════════════════════════════════════════
# FIX 1A — CSS: agregar clase planif-fullscreen
# ═══════════════════════════════════════════════════════════════
old_css = "#planif-panel {\n  font-family: 'Roboto', sans-serif;\n  background: #f0f2f5;\n  margin: 0; padding: 0;\n  height: calc(100vh - 0px);\n  overflow: hidden;\n  display: flex;\n  flex-direction: column;\n}"
new_css = "#planif-panel {\n  font-family: 'Roboto', sans-serif;\n  background: #f0f2f5;\n  margin: 0; padding: 0;\n  height: calc(100vh - 0px);\n  overflow: hidden;\n  display: flex;\n  flex-direction: column;\n}\n#planif-panel.planif-fullscreen {\n  position: fixed !important;\n  inset: 0;\n  z-index: 1100;\n  height: 100vh !important;\n  width: 100vw !important;\n}"
if old_css in content:
    content = content.replace(old_css, new_css, 1)
    changes.append('FIX1A: CSS planif-fullscreen added')
else:
    print('WARNING FIX1A: CSS not found')

# ═══════════════════════════════════════════════════════════════
# FIX 1B — planifIrAStep: agregar/quitar clase fullscreen
# ═══════════════════════════════════════════════════════════════
old_step = (
    "function planifIrAStep(n) {\n"
    "  if (n === 1 && planif_tablaCambiosPendientes) {\n"
    "    var nCambios = 0;\n"
    "    Object.keys(planif_tablaOrdenLocal).forEach(function(k) { nCambios += planif_tablaOrdenLocal[k].length; });\n"
    "    sConfirm(\n"
    "      'Tienes cambios de prioridad sin guardar. Si sales perderás los cambios. ¿Deseas continuar?',\n"
    "      'Cambios pendientes (' + nCambios + ' órdenes)',\n"
    "      'Sí, salir'\n"
    "    ).then(function(r) {\n"
    "      if (!r.isConfirmed) return;\n"
    "      planif_tablaCambiosPendientes = false;\n"
    "      document.querySelectorAll('#planif-panel .step-view').forEach(function(s) { s.classList.remove('active'); });\n"
    "      document.getElementById('planif-step' + n).classList.add('active');\n"
    "    });\n"
    "    return;\n"
    "  }\n"
    "  document.querySelectorAll('#planif-panel .step-view').forEach(function(s) { s.classList.remove('active'); });\n"
    "  document.getElementById('planif-step' + n).classList.add('active');\n"
    "}"
)
new_step = (
    "function planifIrAStep(n) {\n"
    "  var panel = document.getElementById('planif-panel');\n"
    "  function _irAStep(n) {\n"
    "    document.querySelectorAll('#planif-panel .step-view').forEach(function(s) { s.classList.remove('active'); });\n"
    "    document.getElementById('planif-step' + n).classList.add('active');\n"
    "    if (panel) {\n"
    "      if (n === 2) { panel.classList.add('planif-fullscreen'); }\n"
    "      else         { panel.classList.remove('planif-fullscreen'); }\n"
    "    }\n"
    "  }\n"
    "  if (n === 1 && planif_tablaCambiosPendientes) {\n"
    "    var nCambios = 0;\n"
    "    Object.keys(planif_tablaOrdenLocal).forEach(function(k) { nCambios += planif_tablaOrdenLocal[k].length; });\n"
    "    sConfirm(\n"
    "      'Tienes cambios de prioridad sin guardar. Si sales perderás los cambios. ¿Deseas continuar?',\n"
    "      'Cambios pendientes (' + nCambios + ' ordenes)',\n"
    "      'Si, salir'\n"
    "    ).then(function(r) {\n"
    "      if (!r.isConfirmed) return;\n"
    "      planif_tablaCambiosPendientes = false;\n"
    "      _irAStep(1);\n"
    "    });\n"
    "    return;\n"
    "  }\n"
    "  _irAStep(n);\n"
    "}"
)
if old_step in content:
    content = content.replace(old_step, new_step, 1)
    changes.append('FIX1B: planifIrAStep fullscreen logic')
else:
    print('WARNING FIX1B: planifIrAStep not found')
    # debug: find closest match
    idx = content.find('function planifIrAStep')
    print('  Found at:', idx)
    print('  Actual:', repr(content[idx:idx+200]))

# ═══════════════════════════════════════════════════════════════
# FIX 2A — Header: dar ID al div de botones (pt-undoArea-)
# ═══════════════════════════════════════════════════════════════
# The div with display:flex;gap:6px in the render function
old_btn_div = "'<div style=\"display:flex;gap:6px;\">';"
new_btn_div = "'<div id=\"pt-undoArea-' + maqSafe + '\" style=\"display:flex;gap:6px;\">';"
if old_btn_div in content:
    content = content.replace(old_btn_div, new_btn_div, 1)
    changes.append('FIX2A: pt-undoArea ID added')
else:
    print('WARNING FIX2A: header btn div not found')
    # debug
    idx = content.find('display:flex;gap:6px')
    print('  Found at:', idx, repr(content[idx-50:idx+80]))

# ═══════════════════════════════════════════════════════════════
# FIX 2B — Agregar planifActualizarBotonUndo() y llamarla tras drag
# ═══════════════════════════════════════════════════════════════
old_tab_upd = (
    "function planifTablaActualizarOrdenLocal(maqKey, maqSafe) {\n"
    "  var tbody = document.getElementById('pt-body-' + maqSafe);\n"
    "  if (!tbody) return;\n"
    "  planifTablaGuardarUndo(maqKey);\n"
    "  planif_tablaOrdenLocal[maqKey] = Array.from(tbody.querySelectorAll('tr')).map(function(tr){ return tr.dataset.id; });\n"
    "}"
)
new_tab_upd = (
    "function planifActualizarBotonUndo(maqKey, maqSafe) {\n"
    "  var area = document.getElementById('pt-undoArea-' + maqSafe);\n"
    "  if (!area) return;\n"
    "  var stack = planif_undoStack[maqKey];\n"
    "  var n = stack ? stack.length : 0;\n"
    "  area.innerHTML = n > 0\n"
    "    ? '<button onclick=\"planifTablaUndo(\\'' + maqKey + '\\',\\'' + maqSafe + '\\')\" style=\"background:#ef6c00;color:white;border:none;padding:4px 12px;border-radius:4px;cursor:pointer;font-size:12px;font-weight:700;\" title=\"Deshacer ultimo cambio de orden\">&#8617; Deshacer (' + n + ')</button>'\n"
    "    : '';\n"
    "}\n"
    "\n"
    "function planifTablaActualizarOrdenLocal(maqKey, maqSafe) {\n"
    "  var tbody = document.getElementById('pt-body-' + maqSafe);\n"
    "  if (!tbody) return;\n"
    "  planifTablaGuardarUndo(maqKey);\n"
    "  planif_tablaOrdenLocal[maqKey] = Array.from(tbody.querySelectorAll('tr')).map(function(tr){ return tr.dataset.id; });\n"
    "  planifActualizarBotonUndo(maqKey, maqSafe);\n"
    "}"
)
if old_tab_upd in content:
    content = content.replace(old_tab_upd, new_tab_upd, 1)
    changes.append('FIX2B: planifActualizarBotonUndo added + called after drag')
else:
    print('WARNING FIX2B: planifTablaActualizarOrdenLocal not found')

# ═══════════════════════════════════════════════════════════════
# FIX 2C — planifTablaMoverAPosicion también actualiza botón
# ═══════════════════════════════════════════════════════════════
old_mover = (
    "  planifTablaGuardarUndo(maqKey);\n"
    "  ids.splice(idxActual, 1);\n"
    "  ids.splice(nuevaPos - 1, 0, id);\n"
    "  planif_tablaOrdenLocal[maqKey] = ids;"
)
new_mover = (
    "  planifTablaGuardarUndo(maqKey);\n"
    "  ids.splice(idxActual, 1);\n"
    "  ids.splice(nuevaPos - 1, 0, id);\n"
    "  planif_tablaOrdenLocal[maqKey] = ids;\n"
    "  planifActualizarBotonUndo(maqKey, maqSafe);"
)
if old_mover in content:
    content = content.replace(old_mover, new_mover, 1)
    changes.append('FIX2C: planifTablaMoverAPosicion calls ActualizarBotonUndo')
else:
    print('WARNING FIX2C: planifTablaMoverAPosicion pattern not found')
    idx = content.find('planifTablaGuardarUndo(maqKey)')
    while idx != -1:
        print('  GuardarUndo at:', idx, repr(content[idx:idx+120]))
        idx = content.find('planifTablaGuardarUndo(maqKey)', idx+1)

print()
print('Changes applied:', changes)
open('C:/Users/rcdga/planeacion-produccion/Mod_Planificador.html', 'w', encoding='utf-8').write(content)
print('Saved OK')
