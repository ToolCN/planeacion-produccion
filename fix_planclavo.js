const fs = require('fs');
let content = fs.readFileSync('MainApp_PlaneacionHTML.html', 'utf8');

const busca = `trkAbrirMaterialEnviado()" style="display:flex;align-items:center;gap:8px;background:rgba(255,107,43,0.08);border:1px solid rgba(255,107,43,0.3);border-radius:12px;color:#ff6b2b;`;

const nuevo = `trkAbrirPlanClavo()" style="display:flex;align-items:center;gap:8px;background:rgba(168,85,247,0.08);border:1px solid rgba(168,85,247,0.3);border-radius:12px;color:#a855f7;padding:12px 24px;font-family:Rajdhani,sans-serif;font-size:14px;font-weight:700;letter-spacing:1.5px;cursor:pointer;transition:all .2s;text-transform:uppercase;" onmouseover="this.style.background=\\'rgba(168,85,247,0.18)\\'" onmouseout="this.style.background=\\'rgba(168,85,247,0.08)\\'">\\n      🔨 PLAN CLAVO\\n    </button>\\n    <button onclick="trkAbrirMaterialEnviado()" style="display:flex;align-items:center;gap:8px;background:rgba(255,107,43,0.08);border:1px solid rgba(255,107,43,0.3);border-radius:12px;color:#ff6b2b;`;

if (content.includes(busca)) {
  content = content.replace(busca, nuevo);
  fs.writeFileSync('MainApp_PlaneacionHTML.html', content, 'utf8');
  console.log('OK - botón insertado');
} else {
  console.log('NO ENCONTRADO');
}
