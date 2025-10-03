// ====== Sécuriser la config depuis le HTML (APP_CONFIG) ======
(function ensureConfig(){
  if (typeof window.GAS_URL === 'undefined') {
    window.GAS_URL = (window.APP_CONFIG && window.APP_CONFIG.GAS_URL) || '';
  }
  if (typeof window.SHEET_ID === 'undefined') {
    window.SHEET_ID = (window.APP_CONFIG && window.APP_CONFIG.SHEET_ID) || '';
  }
  // Validation basique
  if (!/^https?:\/\//i.test(window.GAS_URL)) {
    console.error('GAS_URL invalide : renseigne une URL /exec complète de ta Web App Apps Script.');
  }
})();

// ====== Helpers ======
const $ = (id) => document.getElementById(id);

function jsonp(url, params = {}) {
  return new Promise((resolve, reject) => {
    const cbName = 'cb_' + Math.random().toString(36).slice(2);
    params.callback = cbName;
    const qs = Object.entries(params)
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
      .join('&');
    const src = url + (url.includes('?') ? '&' : '?') + qs;
    const s = document.createElement('script');
    s.src = src; s.async = true;
    s.onerror = () => { delete window[cbName]; reject(new Error('JSONP load error')); };
    window[cbName] = (data) => { delete window[cbName]; document.body.removeChild(s); resolve(data); };
    document.body.appendChild(s);
  });
}

// ====== Download button state ======
function setDownloadWaiting(){
  const a = $('downloadSuivi'), st = $('dlStatus');
  if (!a) return;
  a.classList.remove('primary'); a.classList.add('wait');
  a.textContent = 'Please wait for download…';
  a.href = '#'; if (st) st.textContent = '';
}
function setDownloadReady(info){
  const a = $('downloadSuivi'), st = $('dlStatus');
  if (!a) return;
  a.classList.remove('wait'); a.classList.add('primary');
  a.textContent = 'Download file';
  a.href = info.url; if (st) st.textContent = info.name || '';
}

// Poll le dernier fichier “suivi”
async function pollLatestSuivi(gasUrl, { attempts=24, intervalMs=5000 } = {}){
  setDownloadWaiting();
  for (let i=0; i<attempts; i++){
    try{
      const res = await jsonp(gasUrl, { action:'latestFile', type:'suivi' });
      if (res?.ok && res.latest?.found && res.latest?.url){
        setDownloadReady(res.latest);
        return;
      }
    }catch(_){}
    await new Promise(r => setTimeout(r, intervalMs));
  }
  // après les tentatives → reste en attente
}

// ====== ML parsing / rendering ======
function parseML(data){
  const headers = data.headers || [];
  const rows = data.rows || [];
  if (!headers.length) return { days: [], teams: [], matrix: [] };

  const dayCols = [], days = [];
  for (let i=1; i<headers.length; i++){
    const h = String(headers[i] || '');
    const m = h.match(/nombre colonne carton\s+(\d{4}-\d{2}-\d{2})$/i);
    if (m){ dayCols.push(i); days.push(m[1]); }
  }

  const teams = [], matrix = [];
  for (const r of rows){
    const team = String(r[0] || '').trim();
    if (!team) continue;
    teams.push(team);
    const vals = [];
    for (let ci=0; ci<dayCols.length; ci++){
      const idx = dayCols[ci];
      let v = r[idx]; if (v == null || v === '') v = 0;
      const n = typeof v === 'number' ? v : Number(String(v).replace(/\s/g,'').replace(',', '.'));
      vals.push(isFinite(n) ? n : 0);
    }
    matrix.push(vals);
  }
  return { days, teams, matrix };
}

function updateKpis({days, teams, matrix}){
  const kTotal = matrix.flat().reduce((a,b)=>a+(b||0),0);
  const kTeams = teams.length;
  const lastDay = days[days.length-1] || '—';
  let lastTotal = 0;
  if (days.length){
    const c = days.length-1;
    for (let r=0;r<matrix.length;r++) lastTotal += (matrix[r][c]||0);
  }

  $('kpiTotal').textContent   = kTotal.toFixed(2);
  $('kpiDays').textContent    = `${days.length} jours`;
  $('kpiTeams').textContent   = kTeams;
  $('kpiLastDay').textContent = lastDay;
  $('kpiLastTotal').textContent = `Total jour : ${lastTotal.toFixed(2)}`;
}

function renderTable({days, teams, matrix}){
  const wrap = $('tableWrap');
  if (!wrap) return;
  if (!days.length) { wrap.innerHTML = '<p><em>Pas de colonnes “nombre colonne carton …” trouvées.</em></p>'; return; }

  let html = '<table><thead><tr><th>Contact</th>';
  for (const d of days) html += `<th>${d}</th>`;
  html += '<th>Total</th></tr></thead><tbody>';

  for (let i=0;i<teams.length;i++){
    const t = teams[i], row = matrix[i] || [];
    const sum = row.reduce((a,b)=>a+(b||0),0);
    html += `<tr><td>${t}</td>`;
    for (const v of row) html += `<td>${v.toFixed(2)}</td>`;
    html += `<td><strong>${sum.toFixed(2)}</strong></td></tr>`;
  }

  const colTotals = new Array(days.length).fill(0);
  for (let c=0;c<days.length;c++){
    for (let r=0;r<matrix.length;r++){ colTotals[c] += (matrix[r][c] || 0); }
  }
  const grandTotal = colTotals.reduce((a,b)=>a+b,0);

  html += `<tfoot><tr><th>Total</th>`;
  for (const v of colTotals) html += `<th>${v.toFixed(2)}</th>`;
  html += `<th>${grandTotal.toFixed(2)}</th></tr></tfoot></tbody></table>`;
  wrap.innerHTML = html;
}

let chartInst = null;
function renderChart({days, teams, matrix}){
  const ctx = $('chart');
  if (!ctx || !days.length) return;

  const base = ['#F07A24','#404040','#8C8C8C','#BFBFBF','#FFB37A','#737373','#D9D9D9','#595959','#FFC699','#A6A6A6'];
  const ds = teams.map((t,i)=>({
    label:t, data:matrix[i]||[],
    backgroundColor: base[i % base.length],
    borderColor:'#ffffff', borderWidth:1
  }));

  if (chartInst) chartInst.destroy();
  chartInst = new Chart(ctx, {
    type:'bar',
    data:{ labels:days, datasets:ds },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'bottom' } },
      scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true } }
    }
  });
}

// ====== Weather embed (iframe + fallback) ======
function initWeatherEmbed(){
  const url = 'https://www.ecowitt.net/home/share?authorize=6H4BCA';
  const frame = document.getElementById('wxFrame');
  const msg = document.getElementById('wxMsg');
  if (!frame) return;

  if (msg) msg.textContent = 'Chargement de la station…';
  frame.src = url;

  frame.addEventListener('load', () => {
    frame.style.display = 'block';
    if (msg) msg.textContent = ' ';
  });

  setTimeout(() => {
    if (frame.style.display === 'none') {
      if (msg) msg.textContent = 'Impossible d’intégrer la page (bloquée par le site). Utilise le bouton “Ouvrir dans un nouvel onglet”.';
    }
  }, 3000);
}

// ====== Boot ======
document.addEventListener('DOMContentLoaded', async ()=>{
  // Charger ML
  try{
    const res = await jsonp(window.GAS_URL, { action:'ml', sheetId: window.SHEET_ID });
    if (res?.ok){
      const parsed = parseML(res.data || {});
      updateKpis(parsed); renderTable(parsed); renderChart(parsed);
    }
  }catch(e){ /* silencieux */ }

  // Bouton téléchargement : poll jusqu’à disponibilité
  if (/^https?:\/\//i.test(window.GAS_URL)) {
    pollLatestSuivi(window.GAS_URL, { attempts:24, intervalMs:5000 });
  }

  // Init météo
  initWeatherEmbed();
});
