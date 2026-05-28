# -*- coding: utf-8 -*-
path = r'C:\Users\rcdga\planeacion-produccion\Mod_Planificador.html'
content = open(path, encoding='utf-8').read()

# ── CAMBIO 4: HTML planif-totales-global (en string JS con \n y \" literales) ──
old4 = (
    'id=\\"planif-totales-global\\" style=\\"display:none;background:#1a237e;color:white;padding:6px 24px;flex-shrink:0;display:none;justify-content:flex-end;align-items:center;gap:24px;font-size:13px;font-weight:900;\\">\\n'
    '      <span>TOTAL GLOBAL:</span>\\n'
    '      <span>Solicitado: <span id=\\"ptg-sol\\" style=\\"color:#90caf9;\\">0</span></span>\\n'
    '      <span>Producido: <span id=\\"ptg-prod\\" style=\\"color:#ce93d8;\\">0</span></span>\\n'
    '      <span>Restan: <span id=\\"ptg-rest\\" style=\\"color:#ffcc02;\\">0</span></span>\\n'
    '    </div>'
)
new4 = (
    'id=\\"planif-totales-global\\" style=\\"display:none;background:#1a237e;color:white;padding:6px 24px;flex-shrink:0;display:none;justify-content:flex-end;align-items:center;gap:24px;font-size:13px;font-weight:900;\\">\\n'
    '      <span>TOTAL GLOBAL:</span>\\n'
    '      <span>Solicitado: <span id=\\"ptg-sol\\" style=\\"color:#90caf9;\\">0</span></span>\\n'
    '      <span>Producido: <span id=\\"ptg-prod\\" style=\\"color:#ce93d8;\\">0</span></span>\\n'
    '      <span>Restan: <span id=\\"ptg-rest\\" style=\\"color:#ffcc02;\\">0</span></span>\\n'
    '      <span style=\\"margin-left:auto;border-left:1px solid rgba(255,255,255,0.3);padding-left:24px;\\">Total a fabricar (kg): <span id=\\"ptg-fab\\" style=\\"color:#4ade80;\\">0</span></span>\\n'
    '    </div>'
)
print('C4:', 'OK' if old4 in content else 'NO ENCONTRADO')
content = content.replace(old4, new4, 1)

# ── CAMBIO 5a: planifRenderBoardTabla (sin IIFE de filtro, líneas ~4902) ──
# Busca el bloque dentro del if(tgDiv)
old5a = (
    "    var elS = document.getElementById('ptg-sol');  if (elS) elS.textContent = planifFormatN(gSol);\n"
    "    var elP = document.getElementById('ptg-prod'); if (elP) elP.textContent = planifFormatN(gProd);\n"
    "    var elR = document.getElementById('ptg-rest'); if (elR) elR.textContent = planifFormatN(gRest);\n"
    "  }\n"
    "\n"
    "  // Inicializar Sortable en cada tbody"
)
new5a = (
    "    var elS = document.getElementById('ptg-sol');  if (elS) elS.textContent = planifFormatN(gSol);\n"
    "    var elP = document.getElementById('ptg-prod'); if (elP) elP.textContent = planifFormatN(gProd);\n"
    "    var elR = document.getElementById('ptg-rest'); if (elR) elR.textContent = planifFormatN(gRest);\n"
    "    (function() {\n"
    "      var gFab = gRest;\n"
    "      planif_sinPedidoData.forEach(function(it) {\n"
    "        var _back4 = it.back || 0;\n"
    "        var _exNeg4 = (it.existNeg !== null && it.existNeg !== undefined) ? (it.existNeg || 0) : 0;\n"
    "        var _pedirSP = it.esVarilla\n"
    "          ? Math.round(it.max - it.exist - _back4 - _exNeg4) - (it.totalRestanActivo || 0)\n"
    "          : Math.round(it.max - it.exist) - (it.totalRestanActivo || 0);\n"
    "        if (_pedirSP > 0) {\n"
    "          if (it.esVarilla) {\n"
    "            var _pvP4 = parseFloat(it.peso) || 0;\n"
    "            var _pvL4 = (function(s){ var m = s.match(/[\\d.]+/); return m ? parseFloat(m[0]) : 0; })(String(it.long||''));\n"
    "            var _pvF4 = (_pvP4 > 0 && _pvL4 > 0) ? _pvP4 * _pvL4 : 0;\n"
    "            gFab += _pvF4 > 0 ? Math.round(_pedirSP * _pvF4) : _pedirSP;\n"
    "          } else {\n"
    "            gFab += _pedirSP;\n"
    "          }\n"
    "        }\n"
    "      });\n"
    "      var elF = document.getElementById('ptg-fab'); if (elF) elF.textContent = planifFormatN(Math.max(0, gFab));\n"
    "    })();\n"
    "  }\n"
    "\n"
    "  // Inicializar Sortable en cada tbody"
)
print('C5a:', 'OK' if old5a in content else 'NO ENCONTRADO')
content = content.replace(old5a, new5a, 1)

# ── CAMBIO 5b: planifFiltrarTablero (dentro del IIFE + Bloquear/desbloquear) ──
old5b = (
    "    var elS = document.getElementById('ptg-sol');  if (elS) elS.textContent = planifFormatN(gSol);\n"
    "    var elP = document.getElementById('ptg-prod'); if (elP) elP.textContent = planifFormatN(gProd);\n"
    "    var elR = document.getElementById('ptg-rest'); if (elR) elR.textContent = planifFormatN(gRest);\n"
    "  })();\n"
    "  // Bloquear/desbloquear handles de drag según filtro activo"
)
new5b = (
    "    var elS = document.getElementById('ptg-sol');  if (elS) elS.textContent = planifFormatN(gSol);\n"
    "    var elP = document.getElementById('ptg-prod'); if (elP) elP.textContent = planifFormatN(gProd);\n"
    "    var elR = document.getElementById('ptg-rest'); if (elR) elR.textContent = planifFormatN(gRest);\n"
    "    // Total a fabricar = Restan (ya en kg) + Pedir de SP (convertido a kg si varilla)\n"
    "    (function() {\n"
    "      var gFab = gRest;\n"
    "      planif_sinPedidoData.forEach(function(it) {\n"
    "        var _back4 = it.back || 0;\n"
    "        var _exNeg4 = (it.existNeg !== null && it.existNeg !== undefined) ? (it.existNeg || 0) : 0;\n"
    "        var _pedirSP = it.esVarilla\n"
    "          ? Math.round(it.max - it.exist - _back4 - _exNeg4) - (it.totalRestanActivo || 0)\n"
    "          : Math.round(it.max - it.exist) - (it.totalRestanActivo || 0);\n"
    "        if (_pedirSP > 0) {\n"
    "          if (it.esVarilla) {\n"
    "            var _pvP4 = parseFloat(it.peso) || 0;\n"
    "            var _pvL4 = (function(s){ var m = s.match(/[\\d.]+/); return m ? parseFloat(m[0]) : 0; })(String(it.long||''));\n"
    "            var _pvF4 = (_pvP4 > 0 && _pvL4 > 0) ? _pvP4 * _pvL4 : 0;\n"
    "            gFab += _pvF4 > 0 ? Math.round(_pedirSP * _pvF4) : _pedirSP;\n"
    "          } else {\n"
    "            gFab += _pedirSP;\n"
    "          }\n"
    "        }\n"
    "      });\n"
    "      var elF = document.getElementById('ptg-fab'); if (elF) elF.textContent = planifFormatN(Math.max(0, gFab));\n"
    "    })();\n"
    "  })();\n"
    "  // Bloquear/desbloquear handles de drag según filtro activo"
)
print('C5b:', 'OK' if old5b in content else 'NO ENCONTRADO')
content = content.replace(old5b, new5b, 1)

open(path, 'w', encoding='utf-8').write(content)
print('GUARDADO')
