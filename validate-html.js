// Validar sintaxis JS del HTML antes de hacer push
const fs = require('fs');
const c = fs.readFileSync('G:/Mi unidad/ERP/hestia-fertility-dashboard.html', 'utf8');
const scripts = [];
let pos = 0;
while(true) {
  const s = c.indexOf('<script', pos);
  if(s<0) break;
  const e = c.indexOf('</script>', s);
  if(e<0) break;
  const tag = c.substring(s, c.indexOf('>',s)+1);
  if(!tag.includes('src')) scripts.push({ content: c.substring(c.indexOf('>',s)+1, e), idx: scripts.length });
  pos = e+1;
}
let ok = true;
scripts.forEach(function(sc) {
  try { new Function(sc.content); console.log('Script',sc.idx,'✓'); }
  catch(e) {
    console.error('Script',sc.idx,'ERROR:', e.message);
    ok = false;
  }
});
if(!ok) { console.error('\n❌ NO HACER PUSH — hay errores de sintaxis'); process.exit(1); }
else console.log('\n✅ JS válido — seguro hacer push');
