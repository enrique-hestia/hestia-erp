/* patch: mover debug DESPUÉS de root.innerHTML=h.join('') */
const fs = require('fs');
const path = 'G:/Mi unidad/ERP/hestia-fertility-dashboard.html';
let c = fs.readFileSync(path, 'utf8');

// 1. Quitar el debug que está ANTES del innerHTML (en _aeEnsureStyles area)
const OLD_EARLY_DEBUG = `  _aeEnsureStyles();\r\n  // ---- DEBUG INLINE (quitar después de diagnostico) ----\r\n  (function(){\r\n    var _dEl=document.getElementById('ai-raw-debug');\r\n    if(!_dEl){_dEl=document.createElement('pre');_dEl.id='ai-raw-debug';}\r\n    var _m0=d.meses&&d.meses[0]||{}, _m5=d.meses&&d.meses[5]||{};\r\n    _dEl.style.cssText='background:#111;color:#0f0;font-size:10px;padding:8px;border-radius:6px;margin-bottom:10px;overflow:auto;max-height:120px;white-space:pre-wrap';\r\n    _dEl.innerHTML='ytdActual='+d.ytdActual+'  meses.length='+(d.meses||[]).length+'\\n'\r\n      +'meses[0]='+JSON.stringify(_m0)+'\\n'\r\n      +'meses[5]='+JSON.stringify(_m5)+'\\n'\r\n      +'meses[-1]='+JSON.stringify((d.meses||[]).slice(-1)[0]||{});\r\n    var _rt=document.getElementById('ai-root');\r\n    if(_rt&&!document.getElementById('ai-raw-debug'))_rt.prepend(_dEl);\r\n  })();\r\n  // ---- FIN DEBUG ----`;
const NEW_EARLY = `  _aeEnsureStyles();`;

if (c.indexOf(OLD_EARLY_DEBUG) < 0) { console.error('OLD EARLY DEBUG NO ENCONTRADO'); process.exit(1); }
c = c.replace(OLD_EARLY_DEBUG, NEW_EARLY);
console.log('✓ Debug temprano eliminado');

// 2. Insertar debug DESPUÉS de root.innerHTML=h.join('');
const OLD_AFTER_ROOT = "root.innerHTML=h.join('');\r\n  try { if(window.lucide) lucide.createIcons(); } catch(e){}\r\n  try {";
const NEW_AFTER_ROOT = `root.innerHTML=h.join('');\r\n  try { if(window.lucide) lucide.createIcons(); } catch(e){}\r\n  // ---- DEBUG RAW (después de innerHTML para sobrevivir) ----\r\n  (function(){\r\n    var _dEl=document.getElementById('ai-raw-debug');\r\n    if(!_dEl){_dEl=document.createElement('pre');_dEl.id='ai-raw-debug';}\r\n    var _m0=d.meses&&d.meses[0]||{}, _m5=d.meses&&d.meses[5]||{};\r\n    _dEl.style.cssText='background:#111;color:#0f0;font-size:10px;padding:8px;border-radius:6px;margin-bottom:10px;overflow:auto;max-height:120px;white-space:pre-wrap';\r\n    _dEl.textContent='ytdActual='+d.ytdActual+'  meses.length='+(d.meses||[]).length+'\\nmeses[0]='+JSON.stringify(_m0)+'\\nmeses[5]='+JSON.stringify(_m5)+'\\nmeses[-1]='+JSON.stringify((d.meses||[]).slice(-1)[0]||{});\r\n    var _rt=document.getElementById('ai-root');\r\n    if(_rt&&!document.getElementById('ai-raw-debug'))_rt.prepend(_dEl);\r\n    else if(_rt&&_dEl.parentNode!==_rt)_rt.prepend(_dEl);\r\n  })();\r\n  // ---- FIN DEBUG ----\r\n  try {`;

if (c.indexOf(OLD_AFTER_ROOT) < 0) { console.error('OLD_AFTER_ROOT NO ENCONTRADO'); process.exit(1); }
c = c.replace(OLD_AFTER_ROOT, NEW_AFTER_ROOT);
console.log('✓ Debug movido a después del innerHTML');

fs.writeFileSync(path, c, 'utf8');
console.log('Tamaño final:', c.length);
