function planifVerHistorial(maq) {
  planif_maqSyncActiva = maq;

  function planifMostrarHistorial(datos) {
    planif_dataHistTemp = datos;
    var html = '<table style="width:100%; border-collapse:collapse; font-size:11px;">';
    html += '<thead style="background:#eee; position:sticky; top:0;"><tr>';
    html += '<th style="padding:8px; width:40px;">SEL</th><th style="padding:8px; width:80px;">ORDEN</th>';
    html += '<th style="padding:8px; width:80px;">CODIGO</th><th style="padding:8px; width:200px;">DESCRIPCION</th>';
    html += '<th style="padding:8px; width:60px;">DIAM.</th><th style="padding:8px; width:70px;">LONG.</th>';
    html += '<th style="padding:8px; width:60px;">CUERDA</th><th style="padding:8px; width:60px;">CUERPO</th>';
    html += '<th style="padding:8px; width:60px;">ACERO</th>';
    html += '<th style="padding:8px; width:60px;">ESTADO</th><th style="padding:8px; width:70px;">PROD</th>';
    html += '</tr></thead><tbody>';
    datos.forEach(function(o) {
      html += '<tr class="planif-hist-row">';
      html += '<td style="padding:8px;"><input type="checkbox" class="planif-chk-hist" value="' + o.id + '"></td>';
      html += '<td style="padding:8px;">' + o.serie + '.' + String(o.orden).padStart(4,'0') + '</td>';
      html += '<td style="padding:8px;">' + (o.codigo||'') + '</td>';
      html += '<td style="padding:8px; font-size:10px;">' + (o.desc || '') + '</td>';
      html += '<td style="padding:8px;">' + (o.dia||'') + '</td>';
      html += '<td style="padding:8px;">' + (o.long||'') + '</td>';
      html += '<td style="padding:8px;">' + (o.cuerda||'') + '</td>';
      html += '<td style="padding:8px;">' + (o.cuerpo||'') + '</td>';
      html += '<td style="padding:8px;">' + (o.acero||'') + '</td>';
      html += '<td style="padding:8px;">' + o.estadoSVG + '</td>';
      html += '<td style="padding:8px;">' + planifFormatN(o.prod) + '</td>';
      html += '</tr>';
    });
    document.getElementById('planif-histContent').innerHTML = html + "</tbody></table>";
    document.getElementById('planif-modalHistorial').style.display = 'flex';
    hideLoader();
  }

  // Si ya está en cache → mostrar inmediato
  if (planif_histCache[maq]) {
    planifMostrarHistorial(planif_histCache[maq]);
    return;
  }
  // Si está cargando en background → esperar
  if (planif_histCargando[maq]) {
    showLoader("Cargando Historial...");
    var _wait = setInterval(function() {
      if (planif_histCache[maq]) {
        clearInterval(_wait);
        planifMostrarHistorial(planif_histCache[maq]);
      }
    }, 200);
    return;
  }
  // No hay cache → cargar normalmente
  showLoader("Cargando Historial...");
  google.script.run.withSuccessHandler(function(resRaw) {
    try { planif_histCache[maq] = JSON.parse(resRaw); }
    catch(e) { planif_histCache[maq] = []; }
    planifMostrarHistorial(planif_histCache[maq]);
  }).obtenerHistorialMaquina(maq);
}

