function planifAbrirModalEst(ev, id) {
  ev.preventDefault();
  ev.stopPropagation();
  planif_estModalId = id;
  var o = planif_ordenesVivas.find(function(x){ return x.id == id; });
  var estadoActual = o ? String(o.estado||'').toUpperCase() : '';
  var tit = document.getElementById('pme-titulo');
  var btnR = document.getElementById('pme-btnRegresar');
  if (tit) tit.textContent = 'Estado: ' + (estadoActual||'ACTIVE');
  var esEstadoEspecial = estadoActual === 'FALTA MP' || estadoActual === 'FALTA ESPEC.' || estadoActual === 'CANCELADO';
  if (btnR) btnR.style.display = esEstadoEspecial ? 'block' : 'none';
  var modal = document.getElementById('planif-modal-est');
  if (!modal) return;
  // Posicionar junto al botón
  var rect = ev.currentTarget.getBoundingClientRect();
  var left = rect.right + 4;
  var top  = rect.top;
  if (left + 210 > window.innerWidth) left = rect.left - 214;
  if (top + 180 > window.innerHeight) top = window.innerHeight - 185;
  modal.style.left = left + 'px';
  modal.style.top  = top + 'px';
  modal.style.display = 'block';
  // Cerrar al click fuera
  setTimeout(function() {
    document.addEventListener('click', planifCerrarModalEstFuera, { once: true });
  }, 10);
}

function planifCerrarModalEstFuera() {
  planifCerrarModalEst();
}

function planifCerrarModalEst() {
  var modal = document.getElementById('planif-modal-est');
  if (modal) modal.style.display = 'none';
  document.removeEventListener('click', planifCerrarModalEstFuera);
}

var planif_estNuevoEstado = null;
var planif_estPedido = null;
var planif_estPartida = null;

function planifAplicarEstadoRapido(nuevoEst) {
  planifCerrarModalEst();
  var id = planif_estModalId;
  var o = planif_ordenesVivas.find(function(x){ return x.id == id; });
  if (!o) return;
  planif_estNuevoEstado = nuevoEst;
  planif_estPedido  = o.pedido  || o.serie || '';
  planif_estPartida = o.partida || '';

  // Si va a FALTA MP o FALTA ESPEC → pedir nota
  if (nuevoEst === 'FALTA MP' || nuevoEst === 'FALTA ESPEC.') {
    var titulo = document.getElementById('planif-nota-titulo');
    if (titulo) titulo.textContent = 'Motivo para ' + nuevoEst;
    var txt = document.getElementById('planif-nota-txt');
    if (txt) txt.value = '';
    var mn = document.getElementById('planif-modal-nota');
    if (mn) mn.classList.add('abierto');
    return;
  }
  // CANCELADO o ACTIVE → aplicar directo sin nota
  planifEjecutarCambioEstadoOrden(id, nuevoEst, '');
}

function planifCerrarModalNota() {
  var mn = document.getElementById('planif-modal-nota');
  if (mn) mn.classList.remove('abierto');
}

function planifConfirmarNota() {
  var txt = document.getElementById('planif-nota-txt');
  var nota = txt ? txt.value.trim() : '';
  planifCerrarModalNota();
  planifEjecutarCambioEstadoOrden(planif_estModalId, planif_estNuevoEstado, nota);
}

function planifEjecutarCambioEstadoOrden(id, nuevoEst, nota) {
  var o = planif_ordenesVivas.find(function(x){ return x.id == id; });
  if (!o) return;
  var btn = document.getElementById('pt-estbtn-' + id);
  if (btn) { btn.textContent = '...'; btn.disabled = true; }

  // Cambiar en memoria todos los procesos de la misma orden (serie+nOrden)
  var serieOrden = String(o.serie||'') + String(o.orden||'');
  planif_ordenesVivas.forEach(function(ov) {
    if (String(ov.serie||'') + String(ov.orden||'') === serieOrden) {
      ov.estado = nuevoEst;
    }
  });

  showLoader('Cambiando estado...');
  google.script.run
    .withSuccessHandler(function() {
      hideLoader();
      if (planif_vistaTablaActiva) planifRenderBoardTabla();
      else planifRenderBoard();
    })
    .withFailureHandler(function(e) {
      hideLoader();
      // Revertir en memoria
      planif_ordenesVivas.forEach(function(ov) {
        if (String(ov.serie||'') + String(ov.orden||'') === serieOrden) {
          ov.estado = o.estado;
        }
      });
      if (btn) { btn.disabled = false; }
      sAlert('Error: ' + e);
    })
    .planifCambiarEstadoOrdenCompleta(id, nuevoEst, nota);
}

// ══ FIN MODAL ESTADO RÁPIDO ════════════════════════════════

