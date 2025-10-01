// ====== Config par défaut ======
const DEFAULT_GAS_URL = 'https://script.google.com/macros/s/AKfycbwO0P3Yo5kw9PPriJPXzUMipBrzlGTR_r-Ff6OyEUnsNu-I9q-rESbBq7l2m6KLA3RJ/exec'; // <-- colle ton /exec
const DEFAULT_SHEET_ID = '1AptbV2NbY0WQZpe_Xt1K2iVlDpgKADElamKQCg3GcXQ';


// ========= HELPERS =========
const $ = (id) => document.getElementById(id);

function jsonp(url, params={}){
  return new Promise((resolve, reject)=>{
    const cbName = 'cb_' + Math.random().toString(36).slice(2);
    params.callback = cbName;
    const qs = Object.entries(params).map(([k,v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&');
    const src = url + (url.includes('?') ? '&' : '?') + qs;
    const s = document.createElement('script');
    s.src = src; s.async = true;
    s.onerror = () => { delete window[cbName]; reject(new Error('JSONP load error')); };
    window[cbName] = (data) => { delete window[cbName]; document.body.removeChild(s); resolve(data); };
    document.body.appendChild(s);
  });
}

function setBtnState({text, href, enabled}){
  const a = $('downloadExtract'); if (!a) return;
  a.textContent = text || 'Download extract';
  a.href = href || '#';
  if (enabled) a.removeAttribute('disabled'); else a.setAttribute('disabled','');
}

// ========= LOGIQUE =========
async function fetchLatestExtract(gasUrl){
  // 1) essaie “extraction”
  const tryKeyword = async (kw) => {
    try {
      const res = await jsonp(gasUrl, { action:'latestFile', type: kw });
      if (res?.ok && res.latest?.found && res.latest?.url) return res.latest;
      return null;
    } catch { return null; }
  };

  let info = await tryKeyword('extraction');
  // 2) tolérance orthographe & variantes (si nécessaire)
  if (!info) info = await tryKeyword('extract');
  if (!info) info = await tryKeyword('inventaire');
  if (!info) info = await tryKeyword('inventory');

  // 3) dernier recours: rien trouvé
  return info;
}

async function updateExtractLink(){
  const gasUrl = ($('gasUrl')?.value || '').trim();
  const statusEl = $('status');
  const metaEl = $('meta');

  if (!gasUrl) {
    setBtnState({ text:'Renseigne l’URL Apps Script', enabled:false });
    if (statusEl) statusEl.textContent = 'Saisis l’URL /exec de la Web App puis clique Actualiser.';
    return;
  }

  setBtnState({ text:'Recherche…', enabled:false, href:'#' });
  if (statusEl) statusEl.textContent = 'Recherche du dernier extract dans le dossier Drive…';
  if (metaEl) metaEl.textContent = '';

  const info = await fetchLatestExtract(gasUrl);

  if (info) {
    setBtnState({ text:`Download extract (${info.name})`, enabled:true, href:info.url });
    if (statusEl) statusEl.textContent = 'Fichier trouvé.';
    if (metaEl) metaEl.textContent = `Taille: ${(info.size/1024/1024).toFixed(2)} MB • Créé le: ${new Date(info.created).toLocaleString()}`;
  } else {
    setBtnState({ text:'No extract found', enabled:false, href:'#' });
    if (statusEl) statusEl.textContent =
      'Aucun fichier “extraction” détecté. Vérifie que ton flux PWA envoie bien saveExtraction=true et que les fichiers contiennent « extraction » dans le nom.';
  }
}

// ========= BOOT =========
document.addEventListener('DOMContentLoaded', ()=>{
  const urlInput = $('gasUrl');
  const saved = localStorage.getItem('GAS_URL_EXTRACT');
  urlInput.value = saved || DEFAULT_GAS_URL;

  $('refreshBtn').addEventListener('click', ()=>{
    localStorage.setItem('GAS_URL_EXTRACT', urlInput.value.trim());
    updateExtractLink();
  });

  // auto au chargement
  updateExtractLink();
});
