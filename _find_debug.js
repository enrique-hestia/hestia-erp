const fs = require('fs');
const c = fs.readFileSync('G:/Mi unidad/ERP/hestia-fertility-dashboard.html', 'utf8');
const fn = c.indexOf('function _renderAnalisisIngresos');
const pos = c.indexOf('root.innerHTML', fn);
console.log('root.innerHTML at line:', c.substring(0,pos).split('\n').length);
console.log('context:', JSON.stringify(c.substring(pos, pos+200)));
// Find "aiChart" canvas creation (inside the innerHTML string)
const canvas = c.indexOf('id="aiChart"', fn);
console.log('aiChart canvas at line:', c.substring(0,canvas).split('\n').length);
// Find "var ctx=document.getElementById" which is AFTER innerHTML
const ctxPos = c.indexOf("var ctx=document.getElementById('aiChart')", fn);
console.log('var ctx line:', c.substring(0,ctxPos).split('\n').length);
console.log('ctx context:', JSON.stringify(c.substring(ctxPos-100, ctxPos+50)));
