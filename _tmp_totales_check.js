var planif_seleccionProcesos=[], planif_catalogo={}, planif_maqsVisibles=[], planif_tablaOrdenLocal={}, planif_ordenesVivas=[], planif_modoFiltro98=false, planif_modoExcedentes=false, window={};
function planifGFOrdenVisible(){};
function planifFormatN(x){return x;};
Calcular y mostrar totales globales
  var gSol = 0, gProd = 0;
  planif_seleccionProcesos.forEach(function(proc) {
    var procUpper = String(proc).trim().toUpperCase();
    (planif_catalogo[proc] || []).forEach(function(maq) {
      if (planif_maqsVisibles.indexOf(maq) < 0) return;
      var maqUpper = String(maq).trim().toUpperCase();
      var maqKey = maqUpper + '||' + procUpper;
      (planif_tablaOrdenLocal[maqKey] || []).forEach(function(id) {
        var o = planif_ordenesVivas.find(function(x){ return x.id == id; });
        if (!o) return;
        if (planif_modoFiltro98 && o.avance < 98) return;
        if (planif_modoExcedentes && !(window._planifExcMap && window._planifExcMap[String(o.codigo||'').toUpperCase()])) return;
        if (typeof planifGFOrdenVisible === 'function' && !planifGFOrdenVisible(o)) return;
        gSol  += Number(o.sol)  || 0;
        gProd += Number(o.prod) || 0;
      });
    });
  });
  var gRest = Math.max(0, gSol - gProd);
  var tgDiv = document.getElementById('planif-totales-global');
  if (tgDiv) {
    tgDiv.style.display = 'flex';
    var elS = document.getElementById('ptg-sol');  if (elS) elS.textContent = planifFormatN(gSol);
    var elP = document.getElementById('ptg-prod'); if (elP) elP.textContent = planifFormatN(gProd);
    var elR = document.getElementById('ptg-rest'); if (elR) elR.textContent = planifFormatN(gRest);
  }

  