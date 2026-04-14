---
# ESTADO DEL PROYECTO — PCP Tool CN
**Última actualización:** 2026-04-14
**Deploy activo:** @480
**URL producción:** https://script.google.com/macros/s/AKfycbyDxTSHPygL1Ba93kHFisiInmZZf6zHM63jOHfqPGv45bdbOJ1Av2lg-DigAGjH50Co1w/exec

---

## ARQUITECTURA

### Plataforma
- Google Apps Script + Google Sheets como base de datos
- Frontend: HTML/CSS/JS servido por GAS vía `doGet()`
- Deploy con clasp desde `C:\Users\rcdga\planeacion-produccion`

### Archivos principales
| Archivo | Descripción |
|---|---|
| `Codigo_PCP_GS.js` | Backend GAS — todas las funciones del servidor |
| `MainApp_PlaneacionHTML.html` | Frontend principal — núcleo + todos los módulos concatenados |
| `Mod_*.html` | Módulos individuales — fuente de verdad para edición |
| `rebuild_main.js` | Script Node.js que reconstruye MainApp desde los módulos |
| `package.json` | Scripts: `npm run deploy` = rebuild + clasp push |

### Sistema de módulos
El MainApp se genera automáticamente concatenando los 23 módulos.
**Flujo obligatorio para cualquier cambio:**
1. Editar el `Mod_*.html` correspondiente
2. Ejecutar `npm run deploy`
3. Hacer commit y push a main

**NUNCA editar MainApp_PlaneacionHTML.html directamente** — se sobreescribe con cada deploy.

---

## MÓDULOS (23 total)

| Módulo | Líneas | Descripción |
|---|---|---|
| Mod_Dashboard.html | 351 | KPIs de pedidos y órdenes |
| Mod_Metricas.html | 1,223 | Tablero de métricas por semana/mes |
| Mod_Validador.html | 920 | Auditoría y validación de pedidos |
| Mod_Tracking.html | 99 | Tracking rápido |
| Mod_NuevoPedido.html | 834 | Captura de pedidos nuevos |
| Mod_PedidoINT.html | 205 | Pedidos internos |
| Mod_Generador.html | 777 | Generador de órdenes de producción |
| Mod_Planificador.html | 7,748 | Planificador FORJA — PROGRAMADOR, PLAN DIARIO, CAPTURA |
| Mod_Programador.html | 892 | Programador MRP |
| Mod_ProgInteractivo.html | 1,579 | Programador interactivo |
| Mod_Wip.html | 329 | WIP piso |
| Mod_Captura.html | 1,535 | Nueva captura de producción |
| Mod_EditorProd.html | 878 | Editor de producción |
| Mod_NuevaRuta.html | 464 | Creación de rutas |
| Mod_EditorRutas.html | 497 | Editor de rutas existentes |
| Mod_TableroSupervisor.html | 1,345 | Tablero del supervisor por proceso |
| Mod_GestionDocs.html | 194 | Gestión de documentos técnicos |
| Mod_Usuarios.html | 393 | Gestión de usuarios y permisos |
| Mod_TrackingModule.html | 774 | Módulo tracking completo |
| Mod_MrpMp.html | 958 | MRP materia prima |
| Mod_SalidaAlambres.html | 210 | Salida de alambres |
| Mod_MisRegistros.html | 566 | Registros del usuario actual |
| Mod_PedidosPendientes.html | 256 | Pedidos pendientes |

---

## BASE DE DATOS (Google Sheets)
**ID hoja principal:** `1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk`

Hojas clave: ORDENES, PEDIDOS, PRODUCCION, LOTES, RUTAS, ESTANDARES, INVENTARIO_EXTERNO, CODIGOS, USUARIOS

---

## CONVENCIONES DE CÓDIGO

- Prefijos por módulo: `dash_`, `tm_`, `val_`, `planif_`, `fprog_`, `fpd_`, `tsup_`, `np_`, `pint_`, `gen_`, `wip_`, `cap_`, `edp_`, `nr_`, `er_`, `mrpmp_`, `sal_`, `misreg_`, `pp_`
- CSS siempre scoped bajo `#[prefijo]-panel`
- Modales: SweetAlert2 en toda la app
- No usar template literals (backticks) — GAS no los soporta en strings de innerHTML/textContent
- No usar `window.onload` en módulos inyectados
- Loader global: `showLoader()` / `hideLoader()`

---

## RESTRICCIONES TÉCNICAS IMPORTANTES

- **GAS CSP bloquea:** ExcelJS CDN → usar SheetJS (XLSX)
- **Template literals:** prohibidos en strings asignados a innerHTML/textContent
- **`window.onload`:** no usar en módulos — sobreescribe el login
- **`require()`:** no disponible en GAS — nunca subir scripts Node.js al proyecto
- **`.claspignore`:** excluye rebuild_main.js, fix_backticks.js, wrap_modules.js, *.py, node_modules/

---

## HISTORIAL DE DECISIONES CLAVE

| Fecha | Decisión |
|---|---|
| 2026-04-14 | Separación de MainApp (25,302 líneas) en 23 módulos Mod_*.html |
| 2026-04-14 | Sistema de build con rebuild_main.js + npm run deploy |
| 2026-04-14 | Conversión de template literals a strings normales (GAS no los soporta) |
| 2026-04-14 | Marcador `// ════ MÓDULOS — rebuild_main.js ════ //` en MainApp para rebuild |

---
