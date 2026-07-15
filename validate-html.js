// Valida la sintaxis del JS embebido en el HTML antes de hacer push.
//
// Uso:  node validate-html.js [ruta.html]      (default: el dashboard del repo)
//
// Por que existe: el HTML es un solo archivo de ~1.4MB con todo el JS inline. Un
// error de sintaxis lo rompe COMPLETO y en silencio hasta que se abre en el navegador.
//
// Notas de implementacion (2026-07-15):
//  - Antes la ruta estaba HARDCODEADA y se ignoraba el argv -> validar otro archivo
//    (ej. un baseline de git) validaba en realidad el del repo y daba falsos "iguales".
//  - Antes usaba new Function() para TODO, lo que reventaba en <script type="module">
//    ("Cannot use import statement outside a module"): un FALSO POSITIVO permanente que
//    hacia que el validador SIEMPRE dijera "NO HACER PUSH". Un validador que siempre
//    falla entrena a ignorarlo. Ahora cada script se checa con el parser real de V8
//    (`node --check`) usando la extension correcta: .mjs si es module, .cjs si es clasico.
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFileSync } = require('child_process');

const file = process.argv[2] || path.join(__dirname, 'hestia-fertility-dashboard.html');
if (!fs.existsSync(file)) { console.error('No existe: ' + file); process.exit(2); }
const c = fs.readFileSync(file, 'utf8');

const scripts = [];
let pos = 0;
while (true) {
  const s = c.indexOf('<script', pos);
  if (s < 0) break;
  const gt = c.indexOf('>', s);
  if (gt < 0) break;
  const e = c.indexOf('</script>', gt);
  if (e < 0) break;
  const tag = c.substring(s, gt + 1);
  if (!/\ssrc\s*=/.test(tag)) {
    scripts.push({
      idx: scripts.length,
      isModule: /type\s*=\s*["']module["']/.test(tag),
      line: c.substring(0, s).split('\n').length,
      content: c.substring(gt + 1, e)
    });
  }
  pos = e + 1;
}

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'valhtml-'));
let ok = true;
for (const sc of scripts) {
  const f = path.join(tmp, 's' + sc.idx + (sc.isModule ? '.mjs' : '.cjs'));
  fs.writeFileSync(f, sc.content, 'utf8');
  const kind = sc.isModule ? 'module' : 'script';
  try {
    execFileSync(process.execPath, ['--check', f], { stdio: 'pipe' });
    console.log('Script ' + sc.idx + ' (' + kind + ', linea ~' + sc.line + ') OK');
  } catch (err) {
    const msg = String(err.stderr || err.message).split('\n').filter(Boolean).slice(0, 4).join('\n  ');
    console.error('Script ' + sc.idx + ' (' + kind + ', linea ~' + sc.line + ') ERROR:\n  ' + msg);
    ok = false;
  }
}
fs.rmSync(tmp, { recursive: true, force: true });

console.log('');
console.log(path.basename(file) + ': ' + scripts.length + ' scripts inline revisados.');
if (!ok) { console.error('NO HACER PUSH - hay errores de sintaxis'); process.exit(1); }
console.log('JS valido - seguro hacer push');
