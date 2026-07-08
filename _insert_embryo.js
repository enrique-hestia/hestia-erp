const fs = require('fs');
let c = fs.readFileSync('G:/Mi unidad/ERP/hestia-fertility-dashboard.html', 'utf8');

// ── 1. Replace the spinner with the embryo SVG loader ──────────────────────
const OLD_LOADER = '<div id="ec-loading" style="display:none;text-align:center;padding:40px;color:var(--color-text-muted);font-size:14px">\n      <div style="width:32px;height:32px;border:3px solid var(--color-divider);border-top-color:var(--color-primary);border-radius:50%;animation:spin 0.8s linear infinite;margin:0 auto 14px"></div>\n      Consultando historial…\n    </div>';

const NEW_LOADER = `<div id="ec-loading" style="display:none;text-align:center;padding:48px 40px">
      <svg class="embryo-loader" viewBox="0 0 80 80" width="88" height="88" style="display:block;margin:0 auto">
        <!-- Zona pelúcida (membrana exterior) -->
        <circle cx="40" cy="40" r="37" fill="rgba(196,106,122,0.05)" stroke="rgba(196,106,122,0.18)" stroke-width="1"/>

        <!-- Etapa 1: Óvulo -->
        <g class="emb-s1">
          <circle cx="40" cy="40" r="21" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.8"/>
          <ellipse cx="33" cy="33" rx="6" ry="4" fill="rgba(255,255,255,0.28)" transform="rotate(-30 33 33)"/>
          <circle cx="40" cy="40" r="5" fill="rgba(196,106,122,0.45)" stroke="#c46a7a" stroke-width="1"/>
        </g>

        <!-- Etapa 2: 2 células -->
        <g class="emb-s2">
          <circle cx="29" cy="40" r="13.5" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.5"/>
          <circle cx="51" cy="40" r="13.5" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.5"/>
          <ellipse cx="24" cy="35" rx="4" ry="2.5" fill="rgba(255,255,255,0.25)" transform="rotate(-20 24 35)"/>
          <ellipse cx="46" cy="35" rx="4" ry="2.5" fill="rgba(255,255,255,0.25)" transform="rotate(-20 46 35)"/>
          <circle cx="29" cy="40" r="3" fill="rgba(196,106,122,0.5)"/>
          <circle cx="51" cy="40" r="3" fill="rgba(196,106,122,0.5)"/>
        </g>

        <!-- Etapa 3: 4 células -->
        <g class="emb-s3">
          <circle cx="29" cy="29" r="11" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.4"/>
          <circle cx="51" cy="29" r="11" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.4"/>
          <circle cx="29" cy="51" r="11" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.4"/>
          <circle cx="51" cy="51" r="11" fill="rgba(196,106,122,0.22)" stroke="#c46a7a" stroke-width="1.4"/>
          <ellipse cx="24" cy="24" rx="3.5" ry="2" fill="rgba(255,255,255,0.25)" transform="rotate(-30 24 24)"/>
          <ellipse cx="46" cy="24" rx="3.5" ry="2" fill="rgba(255,255,255,0.25)" transform="rotate(-30 46 24)"/>
          <circle cx="29" cy="29" r="2.5" fill="rgba(196,106,122,0.5)"/>
          <circle cx="51" cy="29" r="2.5" fill="rgba(196,106,122,0.5)"/>
          <circle cx="29" cy="51" r="2.5" fill="rgba(196,106,122,0.5)"/>
          <circle cx="51" cy="51" r="2.5" fill="rgba(196,106,122,0.5)"/>
        </g>

        <!-- Etapa 4: Mórula (8 células) -->
        <g class="emb-s4">
          <circle cx="40" cy="22" r="8.5" fill="rgba(196,106,122,0.25)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="55" cy="28" r="8.5" fill="rgba(196,106,122,0.20)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="59" cy="43" r="8.5" fill="rgba(196,106,122,0.25)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="51" cy="57" r="8.5" fill="rgba(196,106,122,0.20)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="36" cy="60" r="8.5" fill="rgba(196,106,122,0.25)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="22" cy="54" r="8.5" fill="rgba(196,106,122,0.20)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="18" cy="38" r="8.5" fill="rgba(196,106,122,0.25)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="26" cy="24" r="8.5" fill="rgba(196,106,122,0.20)" stroke="#c46a7a" stroke-width="1.2"/>
          <ellipse cx="37" cy="19" rx="3" ry="1.8" fill="rgba(255,255,255,0.22)" transform="rotate(-20 37 19)"/>
          <ellipse cx="52" cy="25" rx="3" ry="1.8" fill="rgba(255,255,255,0.22)" transform="rotate(-20 52 25)"/>
        </g>

        <!-- Etapa 5: Blastocisto -->
        <g class="emb-s5">
          <!-- Células trofoblasto (anillo exterior) -->
          <circle cx="40" cy="17" r="7" fill="rgba(196,106,122,0.32)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="55" cy="22" r="7" fill="rgba(196,106,122,0.28)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="63" cy="36" r="7" fill="rgba(196,106,122,0.32)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="61" cy="52" r="7" fill="rgba(196,106,122,0.28)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="49" cy="63" r="7" fill="rgba(196,106,122,0.32)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="32" cy="64" r="7" fill="rgba(196,106,122,0.28)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="19" cy="54" r="7" fill="rgba(196,106,122,0.32)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="16" cy="38" r="7" fill="rgba(196,106,122,0.28)" stroke="#c46a7a" stroke-width="1.1"/>
          <circle cx="24" cy="22" r="7" fill="rgba(196,106,122,0.32)" stroke="#c46a7a" stroke-width="1.1"/>
          <!-- Blastocele (cavidad fluida) -->
          <circle cx="40" cy="40" r="19" fill="rgba(253,245,246,0.55)" stroke="rgba(196,106,122,0.12)" stroke-width="0.5"/>
          <!-- Masa celular interna (ICM) -->
          <circle cx="49" cy="31" r="7"  fill="rgba(196,106,122,0.45)" stroke="#c46a7a" stroke-width="1.2"/>
          <circle cx="55" cy="38" r="4.5" fill="rgba(196,106,122,0.38)" stroke="#c46a7a" stroke-width="1"/>
          <ellipse cx="46" cy="28" rx="2.5" ry="1.5" fill="rgba(255,255,255,0.30)" transform="rotate(-25 46 28)"/>
          <!-- Pulso final -->
          <circle cx="40" cy="40" r="37" fill="none" stroke="#c46a7a" stroke-width="1" class="emb-pulse"/>
        </g>
      </svg>
      <p class="emb-label" style="font-size:13px;color:var(--color-text-muted);margin-top:16px;letter-spacing:.02em">Consultando historial…</p>
    </div>`;

c = c.replace(OLD_LOADER, NEW_LOADER);

// ── 2. Add CSS keyframes in the <style> block ─────────────────────────────
const STYLE_INSERT_BEFORE = '[data-theme="dark"]';
if(c.indexOf(STYLE_INSERT_BEFORE) < 0){ console.error('style marker not found'); process.exit(1); }

const EMBRYO_CSS = `
/* ── Embryo loader ───────────────────────────────────────────── */
.embryo-loader { overflow: visible; }
.emb-s1, .emb-s2, .emb-s3, .emb-s4, .emb-s5 { opacity: 0; }
.emb-s1 { animation: emb1 5s ease-in-out infinite; }
.emb-s2 { animation: emb2 5s ease-in-out infinite; }
.emb-s3 { animation: emb3 5s ease-in-out infinite; }
.emb-s4 { animation: emb4 5s ease-in-out infinite; }
.emb-s5 { animation: emb5 5s ease-in-out infinite; }
.emb-pulse { opacity: 0; animation: emb-pulse 5s ease-in-out infinite; }
.emb-label { animation: emb-fade 5s ease-in-out infinite; }

@keyframes emb1 {
  0%    { opacity: 0; transform: scale(0.6); }
  4%    { opacity: 1; transform: scale(1); }
  16%   { opacity: 1; transform: scale(1); }
  22%   { opacity: 0; transform: scale(0.85); }
  100%  { opacity: 0; }
}
@keyframes emb2 {
  0%,18%  { opacity: 0; }
  22%     { opacity: 1; }
  36%     { opacity: 1; }
  42%     { opacity: 0; }
  100%    { opacity: 0; }
}
@keyframes emb3 {
  0%,38%  { opacity: 0; }
  42%     { opacity: 1; }
  56%     { opacity: 1; }
  62%     { opacity: 0; }
  100%    { opacity: 0; }
}
@keyframes emb4 {
  0%,58%  { opacity: 0; }
  62%     { opacity: 1; }
  74%     { opacity: 1; }
  80%     { opacity: 0; }
  100%    { opacity: 0; }
}
@keyframes emb5 {
  0%,76%  { opacity: 0; }
  82%     { opacity: 1; transform: scale(1); }
  90%     { opacity: 1; transform: scale(1.03); }
  96%     { opacity: 1; transform: scale(1); }
  100%    { opacity: 0; }
}
@keyframes emb-pulse {
  0%,76%   { opacity: 0; transform: scale(1); }
  85%      { opacity: 0; transform: scale(1); }
  90%      { opacity: 0.5; transform: scale(1.18); }
  96%      { opacity: 0; transform: scale(1.4); }
  100%     { opacity: 0; }
}
@keyframes emb-fade {
  0%,76%  { opacity: 0.4; }
  84%     { opacity: 1; }
  96%     { opacity: 1; }
  100%    { opacity: 0.4; }
}

`;

c = c.replace(STYLE_INSERT_BEFORE, EMBRYO_CSS + STYLE_INSERT_BEFORE);

fs.writeFileSync('G:/Mi unidad/ERP/hestia-fertility-dashboard.html', c, 'utf8');
console.log('Embryo loader inserted OK');
