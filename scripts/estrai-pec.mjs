// Estrazione delle PEC dei Comuni italiani dagli Open Data di IndicePA (AgID).
// Versione per GitHub Actions (Node 20+): stessa logica di pec-comuni-italiani.user.js.
// Uso:  node scripts/estrai-pec.mjs [cartella-di-uscita]
// Esito: <cartella>/PEC_Comuni_italiani.xlsx  e  <cartella>/riepilogo.md

import fs from 'node:fs';
import path from 'node:path';
import XLSX from 'xlsx';

const ORIGINE = process.env.IPA_ORIGINE || 'https://indicepa.gov.it';
const USCITA = process.argv[2] || 'dati';

// ---------------- risorse Open Data IndicePA (stessi ID dello userscript) ----------------
const RISORSE = {
  enti:       { id: 'd09adf99-dc10-4349-8c53-27b1e5aa97b6', nome: 'Enti' },
  domicili:   { id: '0226284e-27e5-4fdd-8bf9-4deba9f4d239', nome: 'Domicili digitali (ente e AOO)' },
  domiciliUO: { id: 'c306c72c-77ab-4f37-8a3a-6f7888979003', nome: 'Domicili digitali delle Unità Organizzative' },
  uo:         { id: 'b0aa1f6c-f135-4c8a-b416-396fed4e1a5d', nome: 'Unità Organizzative' },
};
const AMMINISTRAZIONI_TXT = '/ipa-dati/dataset/502ff370-1b2c-4310-94c7-f39ceb7500e3/resource/3ed63523-ff9c-41f6-a6fe-980f3d9e501f/download/amministrazioni.txt';
const urlDump = id => `${ORIGINE}/ipa-dati/datastore/dump/${id}?format=json`;

// ---------------- utilità ----------------
const log = t => console.log('[PEC Comuni]', t);
const pulisci = s => (s == null ? '' : String(s)).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
const isPec = tipo => /pec/i.test(pulisci(tipo));
const isIndirizzo = s => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(pulisci(s));
const campo = (o, ...nomi) => {
  for (const n of nomi) {
    if (o[n] != null) return pulisci(o[n]);
    const k = Object.keys(o).find(x => x.toLowerCase() === n.toLowerCase());
    if (k && o[k] != null) return pulisci(o[k]);
  }
  return '';
};

async function scaricaTesto(url, etichetta, tentativi = 3) {
  for (let t = 1; ; t++) {
    try {
      const r = await fetch(url, { headers: { 'User-Agent': 'pec-comuni-italiani (GitHub Actions)' } });
      if (!r.ok) throw new Error(`risposta ${r.status}`);
      const testo = await r.text();
      log(`«${etichetta}»: ${(testo.length / 1e6).toFixed(1)} MB scaricati`);
      return testo;
    } catch (e) {
      if (t >= tentativi) throw new Error(`${etichetta}: ${e.message}`);
      log(`«${etichetta}»: ${e.message}, nuovo tentativo tra ${10 * t} s`);
      await new Promise(r => setTimeout(r, 10000 * t));
    }
  }
}

async function leggiRisorsa(chiave) {
  const ris = RISORSE[chiave];
  const j = JSON.parse(await scaricaTesto(urlDump(ris.id), ris.nome));
  const campi = (j.fields || []).map(f => f.id);
  const righe = (j.records || []).map(rec => Array.isArray(rec) ? Object.fromEntries(campi.map((c, i) => [c, rec[i]])) : rec);
  log(`«${ris.nome}»: ${righe.length} righe`);
  return righe;
}

async function leggiAmministrazioni() {
  try {
    const righe = (await scaricaTesto(ORIGINE + AMMINISTRAZIONI_TXT, 'Amministrazioni')).split(/\r?\n/).filter(Boolean);
    const sep = righe[0].includes('\t') ? '\t' : (righe[0].includes(';') ? ';' : ',');
    const intest = righe[0].split(sep).map(pulisci);
    const i = re => intest.findIndex(h => re.test(h));
    const [iCod, iCom, iPro, iReg] = [i(/^cod_amm$/i), i(/^comune$/i), i(/^provincia$/i), i(/^regione$/i)];
    const mappa = new Map();
    if (iCod < 0) return mappa;
    for (const r of righe.slice(1)) {
      const c = r.split(sep);
      mappa.set(pulisci(c[iCod]).toLowerCase(), { comune: pulisci(c[iCom]), provincia: pulisci(c[iPro]), regione: pulisci(c[iReg]) });
    }
    return mappa;
  } catch (e) {
    log('amministrazioni.txt non letto: ' + e.message);
    return new Map();
  }
}

// ---------------- elaborazione ----------------
async function main() {
  const t0 = Date.now();

  // 1) Enti: solo i Comuni
  const comuni = new Map();
  for (const e of await leggiRisorsa('enti')) {
    const ipa = campo(e, 'Codice_IPA'), den = campo(e, 'Denominazione_ente');
    if (!(campo(e, 'Codice_Categoria') === 'L6' && (/^c_[a-z]\d{3}$/i.test(ipa) || /^comune di /i.test(den)))) continue;
    comuni.set(ipa.toLowerCase(), {
      ipa, denominazione: den, cf: campo(e, 'Codice_fiscale_ente'),
      istat: campo(e, 'Codice_comune_ISTAT'), catastale: campo(e, 'Codice_catastale_comune'),
      sito: campo(e, 'Sito_istituzionale'), comune: den.replace(/^comune di\s+/i, ''),
      provincia: '', regione: '', domicilio: '', pec: [],
      mailEnte: [1, 2, 3, 4, 5].map(n => ({ m: campo(e, 'Mail' + n), t: campo(e, 'Tipo_Mail' + n) })),
    });
  }
  if (!comuni.size) throw new Error('nessun comune trovato nel dataset Enti (il formato potrebbe essere cambiato)');
  log(`Comuni trovati: ${comuni.size}`);

  // 2) provincia e regione
  const amm = await leggiAmministrazioni();
  for (const c of comuni.values()) {
    const a = amm.get(c.ipa.toLowerCase());
    if (a) { c.provincia = a.provincia; c.regione = a.regione; if (a.comune) c.comune = a.comune; }
  }

  // 3) PEC senza doppioni
  const visti = new Set();
  const aggiungi = (c, r) => {
    const pec = pulisci(r.pec).toLowerCase();
    if (!isIndirizzo(pec)) return;
    const k = c.ipa + '|' + pec;
    if (visti.has(k)) return;
    visti.add(k);
    c.pec.push({ ...r, pec });
  };

  // 3a) domicili ente e AOO
  for (const d of await leggiRisorsa('domicili')) {
    const c = comuni.get(campo(d, 'Codice_IPA').toLowerCase());
    if (!c) continue;
    const aoo = campo(d, 'Codice_uni_aoo');
    const r = { livello: aoo ? 'AOO – domicilio digitale' : 'Ente – domicilio digitale', ufficio: aoo ? campo(d, 'Descrizione_aoo') : 'Ente', codiceUfficio: aoo,
      pec: campo(d, 'Domicilio_digitale'), tipo: campo(d, 'Tipo_domicilio_digitale') || 'PEC', data: campo(d, 'Data_pubblicazione'), fonte: RISORSE.domicili.nome };
    if (!aoo && !c.domicilio) c.domicilio = pulisci(r.pec).toLowerCase();
    aggiungi(c, r);
  }
  // domicilio principale registrato sull'AOO
  for (const c of comuni.values()) {
    if (c.domicilio) continue;
    const aoo = c.pec.filter(p => p.livello === 'AOO – domicilio digitale');
    const p = aoo.find(x => /domicilio digitale/i.test(x.ufficio)) || aoo[0];
    if (p) c.domicilio = p.pec;
  }
  // 3b) domicili UO
  for (const d of await leggiRisorsa('domiciliUO')) {
    const c = comuni.get(campo(d, 'Codice_IPA').toLowerCase());
    if (!c) continue;
    aggiungi(c, { livello: 'UO – domicilio digitale', ufficio: campo(d, 'Descrizione_uo', 'Descrizione_ufficio'), codiceUfficio: campo(d, 'Codice_uni_uo'),
      pec: campo(d, 'Domicilio_digitale'), tipo: campo(d, 'Tipo_domicilio_digitale') || 'PEC', data: campo(d, 'Data_pubblicazione'), fonte: RISORSE.domiciliUO.nome });
  }
  // 3c) PEC della scheda ente
  for (const c of comuni.values()) {
    c.mailEnte.filter(x => isPec(x.t)).forEach(x => aggiungi(c, { livello: 'Ente – casella PEC', ufficio: 'Ente', codiceUfficio: '', pec: x.m, tipo: 'PEC', data: '', fonte: RISORSE.enti.nome }));
    if (!c.domicilio) { const p = c.mailEnte.find(x => isPec(x.t) && isIndirizzo(x.m)); if (p) c.domicilio = p.m.toLowerCase(); }
  }
  // 3d) PEC degli uffici
  for (const u of await leggiRisorsa('uo')) {
    const c = comuni.get(campo(u, 'Codice_IPA').toLowerCase());
    if (!c) continue;
    for (const n of [1, 2, 3]) {
      if (!isPec(campo(u, 'Tipo_Mail' + n))) continue;
      aggiungi(c, { livello: 'UO – casella PEC', ufficio: campo(u, 'Descrizione_uo'), codiceUfficio: campo(u, 'Codice_uni_uo'),
        pec: campo(u, 'Mail' + n), tipo: 'PEC', data: '', fonte: RISORSE.uo.nome });
    }
  }

  // 4) Excel
  const elenco = [...comuni.values()].sort((a, b) =>
    (a.regione || '~').localeCompare(b.regione || '~', 'it') || (a.provincia || '~').localeCompare(b.provincia || '~', 'it') || a.comune.localeCompare(b.comune, 'it'));
  const dataIt = new Date().toLocaleDateString('it-IT', { timeZone: 'Europe/Rome' });
  const shComuni = [['Comune', 'Provincia', 'Regione', 'Codice ISTAT', 'Codice catastale', 'Codice IPA', 'Codice fiscale', 'Domicilio digitale principale', 'N. indirizzi PEC', 'Sito istituzionale'],
    ...elenco.map(c => [c.comune, c.provincia, c.regione, c.istat, c.catastale, c.ipa, c.cf, c.domicilio, c.pec.length, c.sito])];
  const shPec = [['Comune', 'Provincia', 'Regione', 'Codice IPA', 'Livello', 'Ufficio', 'Codice ufficio', 'Indirizzo PEC', 'Tipo', 'Data pubblicazione', 'Fonte (dataset IndicePA)']];
  elenco.forEach(c => c.pec.forEach(p => shPec.push([c.comune, c.provincia, c.regione, c.ipa, p.livello, p.ufficio, p.codiceUfficio, p.pec, p.tipo, p.data, p.fonte])));
  const LIVELLI = ['Ente – domicilio digitale', 'AOO – domicilio digitale', 'UO – domicilio digitale', 'Ente – casella PEC', 'UO – casella PEC'];
  const conta = l => shPec.filter(r => r[4] === l).length;
  const senzaDomicilio = elenco.filter(c => !c.domicilio).length;
  const senzaRegione = elenco.filter(c => !c.regione).length;
  const shRiep = [
    ['PEC dei Comuni italiani – estrazione da IndicePA'],
    ['Data estrazione', dataIt],
    ['Fonte', 'IndicePA – Indice dei domicili digitali della PA (AgID), Open Data: ' + ORIGINE + '/ipa-dati/'],
    [],
    ['Comuni', elenco.length],
    ['Comuni senza domicilio digitale principale', senzaDomicilio],
    ['Comuni senza provincia/regione (non trovati in amministrazioni.txt)', senzaRegione],
    ['Indirizzi PEC totali (senza doppioni)', shPec.length - 1],
    ...LIVELLI.map(l => ['  ' + l, conta(l)]),
    [],
    ['Dataset usati'],
    ...Object.values(RISORSE).map(r => [r.nome, urlDump(r.id)]),
    ['Amministrazioni (provincia e regione)', ORIGINE + AMMINISTRAZIONI_TXT],
    [],
    ['Nota', 'Per le comunicazioni con valore legale va usato il domicilio digitale dell\'ente (art. 6-ter del D.Lgs. 82/2005, CAD). I dati sono pubblici: riportare la fonte e usarli per comunicazioni istituzionali, non per invii massivi indesiderati.'],
  ];
  const wb = XLSX.utils.book_new();
  const foglio = (righe, nome, larghezze) => {
    const ws = XLSX.utils.aoa_to_sheet(righe);
    ws['!cols'] = larghezze.map(w => ({ wch: w }));
    if (righe.length > 1 && nome !== 'Riepilogo') ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: righe.length - 1, c: righe[0].length - 1 } }) };
    XLSX.utils.book_append_sheet(wb, ws, nome);
  };
  foglio(shComuni, 'Comuni', [28, 9, 22, 12, 10, 10, 13, 40, 9, 40]);
  foglio(shPec, 'PEC', [26, 9, 20, 10, 24, 40, 12, 42, 6, 12, 36]);
  foglio(shRiep, 'Riepilogo', [62, 90]);

  fs.mkdirSync(USCITA, { recursive: true });
  const fileXlsx = path.join(USCITA, 'PEC_Comuni_italiani.xlsx');
  XLSX.writeFile(wb, fileXlsx);

  // riepilogo leggibile su GitHub
  const md = [
    '# Ultima estrazione', '',
    `- Data: **${dataIt}**`,
    `- Comuni: **${elenco.length}** (senza domicilio digitale principale: ${senzaDomicilio}; senza provincia/regione: ${senzaRegione})`,
    `- Indirizzi PEC (senza doppioni): **${shPec.length - 1}**`,
    ...LIVELLI.map(l => `  - ${l}: ${conta(l)}`),
    '', 'Fonte: IndicePA (AgID), Open Data.', '',
  ].join('\n');
  fs.writeFileSync(path.join(USCITA, 'riepilogo.md'), md);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, md);
  log(`Fatto in ${Math.round((Date.now() - t0) / 1000)} s: ${elenco.length} comuni, ${shPec.length - 1} PEC → ${fileXlsx}`);
}

main().catch(e => { console.error('[PEC Comuni] ERRORE:', e.message); process.exit(1); });
