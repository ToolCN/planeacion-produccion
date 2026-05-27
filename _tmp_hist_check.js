function planifCargarHistorialBg() {
  // Precarga el historial de TODAS las máquinas visibles en background
  planif_seleccionProcesos.forEach(function(proc) {
    (planif_catalogo[proc] || []).forEach(function(maq) {
      if (planif_maqsVisibles.indexOf(maq) < 0) return;
      if (planif_histCache[maq] || planif_histCargando[maq]) return;
      planif_histCargando[maq] = true;
      google.script.run
        .withSuccessHandler(function(maqRef) {
          return function(resRaw) {
            try { planif_histCache[maqRef] = JSON.parse(resRaw); }
            catch(e) { planif_histCache[maqRef] = []; }
            planif_histCargando[maqRef] = false;
          };
        }(maq))
        .withFailureHandler(function(maqRef) {
          return function() { planif_histCargando[maqRef] = false; };
        }(maq))
        .obtenerHistorialMaquina(maq);
    });
  });
}

