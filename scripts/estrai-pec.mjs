// Estrazione delle PEC dei Comuni e delle Scuole italiane dagli Open Data di IndicePA (AgID).
// Versione per GitHub Actions (Node 20+): stessa logica di pec-comuni-italiani.user.js.
// Uso:  node scripts/estrai-pec.mjs [cartella-di-uscita]
// Esito: <cartella>/PEC_Comuni_italiani.xlsx, <cartella>/PEC_Scuole_italiane.xlsx e <cartella>/riepilogo.md

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
const log = t => console.log('[PEC]', t);
const pulisci = s => (s == null ? '' : String(s)).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
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

// ---------------- categorie di enti (stessa definizione dello userscript) ----------------
const CATEGORIE = [
  {
    chiave: 'comuni', titolo: 'PEC dei Comuni italiani', nomeFile: 'PEC_Comuni_italiani.xlsx', foglio: 'Comuni', etichetta: 'Comuni',
    filtro: e => campo(e, 'Codice_Categoria') === 'L6' && (/^c_[a-z]\d{3}$/i.test(campo(e, 'Codice_IPA')) || /^comune di /i.test(campo(e, 'Denominazione_ente'))),
    nome: (e, a) => (a && a.comune) || campo(e, 'Denominazione_ente').replace(/^comune di\s+/i, ''),
    intestazione: ['Comune', 'Provincia', 'Regione', 'Codice ISTAT', 'Codice catastale', 'Codice IPA', 'Codice fiscale', 'Domicilio digitale principale', 'N. indirizzi PEC', 'Sito istituzionale'],
    riga: s => [s.nome, s.provincia, s.regione, s.istat, s.catastale, s.ipa, s.cf, s.domicilio, s.pec.length, s.sito],
    larghezze: [28, 9, 22, 12, 10, 10, 13, 40, 9, 40],
  },
  {
    chiave: 'scuole', titolo: 'PEC delle Scuole statali italiane', nomeFile: 'PEC_Scuole_italiane.xlsx', foglio: 'Scuole', etichetta: 'Scuole',
    // istituti scolastici statali: hanno il codice meccanografico del Ministero (Codice_MIUR)
    filtro: e => campo(e, 'Codice_MIUR') !== '',
    nome: (e) => campo(e, 'Denominazione_ente'),
    intestazione: ['Istituto', 'Codice meccanografico', 'Comune', 'Provincia', 'Regione', 'Codice ISTAT comune', 'Codice IPA', 'Codice fiscale', 'Domicilio digitale principale', 'N. indirizzi PEC', 'Sito istituzionale'],
    riga: s => [s.nome, s.miur, s.comune, s.provincia, s.regione, s.istat, s.ipa, s.cf, s.domicilio, s.pec.length, s.sito],
    larghezze: [45, 14, 24, 9, 22, 12, 16, 13, 40, 9, 40],
  },
];
const LIVELLI = ['Ente – domicilio digitale', 'AOO – domicilio digitale', 'UO – domicilio digitale', 'Ente – casella PEC', 'UO – casella PEC'];

// ---------------- download ----------------
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

// ---------------- elaborazione di una categoria ----------------
function elabora(cat, d) {
  // 1) enti della categoria
  const enti = new Map();
  for (const e of d.enti) {
    if (!cat.filtro(e)) continue;
    const ipa = campo(e, 'Codice_IPA');
    const a = d.amm.get(ipa.toLowerCase());
    enti.set(ipa.toLowerCase(), {
      ipa, nome: cat.nome(e, a), cf: campo(e, 'Codice_fiscale_ente'), miur: campo(e, 'Codice_MIUR'),
      istat: campo(e, 'Codice_comune_ISTAT'), catastale: campo(e, 'Codice_catastale_comune'), sito: campo(e, 'Sito_istituzionale'),
      comune: a ? a.comune : '', provincia: a ? a.provincia : '', regione: a ? a.regione : '',
      domicilio: '', pec: [],
      mailEnte: [1, 2, 3, 4, 5].map(n => ({ m: campo(e, 'Mail' + n), t: campo(e, 'Tipo_Mail' + n) })),
    });
  }
  if (!enti.size) throw new Error(`nessun ente trovato per «${cat.etichetta}» (il formato del dataset Enti potrebbe essere cambiato)`);
  log(`${cat.etichetta}: ${enti.size}`);

  // 2) PEC senza doppioni
  const visti = new Set();
  const aggiungi = (s, r) => {
    const pec = pulisci(r.pec).toLowerCase();
    if (!isIndirizzo(pec)) return;
    const k = s.ipa + '|' + pec;
    if (visti.has(k)) return;
    visti.add(k);
    s.pec.push({ ...r, pec });
  };
  for (const x of d.domicili) {           // domicili ente e AOO
    const s = enti.get(campo(x, 'Codice_IPA').toLowerCase());
    if (!s) continue;
    const aoo = campo(x, 'Codice_uni_aoo');
    const r = { livello: aoo ? LIVELLI[1] : LIVELLI[0], ufficio: aoo ? campo(x, 'Descrizione_aoo') : 'Ente', codiceUfficio: aoo,
      pec: campo(x, 'Domicilio_digitale'), tipo: campo(x, 'Tipo_domicilio_digitale') || 'PEC', data: campo(x, 'Data_pubblicazione'), fonte: RISORSE.domicili.nome };
    if (!aoo && !s.domicilio) s.domicilio = pulisci(r.pec).toLowerCase();
    aggiungi(s, r);
  }
  for (const s of enti.values()) {        // domicilio principale registrato sull'AOO
    if (s.domicilio) continue;
    const aoo = s.pec.filter(p => p.livello === LIVELLI[1]);
    const p = aoo.find(x => /domicilio digitale/i.test(x.ufficio)) || aoo[0];
    if (p) s.domicilio = p.pec;
  }
  for (const x of d.domiciliUO) {         // domicili UO
    const s = enti.get(campo(x, 'Codice_IPA').toLowerCase());
    if (!s) continue;
    aggiungi(s, { livello: LIVELLI[2], ufficio: campo(x, 'Descrizione_uo', 'Descrizione_ufficio'), codiceUfficio: campo(x, 'Codice_uni_uo'),
      pec: campo(x, 'Domicilio_digitale'), tipo: campo(x, 'Tipo_domicilio_digitale') || 'PEC', data: campo(x, 'Data_pubblicazione'), fonte: RISORSE.domiciliUO.nome });
  }
  for (const s of enti.values()) {        // PEC della scheda ente
    s.mailEnte.filter(x => isPec(x.t)).forEach(x => aggiungi(s, { livello: LIVELLI[3], ufficio: 'Ente', codiceUfficio: '', pec: x.m, tipo: 'PEC', data: '', fonte: RISORSE.enti.nome }));
    if (!s.domicilio) { const p = s.mailEnte.find(x => isPec(x.t) && isIndirizzo(x.m)); if (p) s.domicilio = p.m.toLowerCase(); }
  }
  for (const u of d.uo) {                 // PEC degli uffici
    const s = enti.get(campo(u, 'Codice_IPA').toLowerCase());
    if (!s) continue;
    for (const n of [1, 2, 3]) {
      if (!isPec(campo(u, 'Tipo_Mail' + n))) continue;
      aggiungi(s, { livello: LIVELLI[4], ufficio: campo(u, 'Descrizione_uo'), codiceUfficio: campo(u, 'Codice_uni_uo'),
        pec: campo(u, 'Mail' + n), tipo: 'PEC', data: '', fonte: RISORSE.uo.nome });
    }
  }

  // 3) Excel
  const elenco = [...enti.values()].sort((a, b) =>
    (a.regione || '~').localeCompare(b.regione || '~', 'it') || (a.provincia || '~').localeCompare(b.provincia || '~', 'it')
    || (a.comune || '').localeCompare(b.comune || '', 'it') || a.nome.localeCompare(b.nome, 'it'));
  const dataIt = new Date().toLocaleDateString('it-IT', { timeZone: 'Europe/Rome' });
  const shEnti = [cat.intestazione, ...elenco.map(cat.riga)];
  const shPec = [[cat.etichetta === 'Comuni' ? 'Comune' : 'Istituto', 'Provincia', 'Regione', 'Codice IPA', 'Livello', 'Ufficio', 'Codice ufficio', 'Indirizzo PEC', 'Tipo', 'Data pubblicazione', 'Fonte (dataset IndicePA)']];
  elenco.forEach(s => s.pec.forEach(p => shPec.push([s.nome, s.provincia, s.regione, s.ipa, p.livello, p.ufficio, p.codiceUfficio, p.pec, p.tipo, p.data, p.fonte])));
  const conta = l => shPec.filter(r => r[4] === l).length;
  const senzaDomicilio = elenco.filter(s => !s.domicilio).length;
  const senzaRegione = elenco.filter(s => !s.regione).length;
  const shRiep = [
    [cat.titolo + ' – estrazione da IndicePA'],
    ['Data estrazione', dataIt],
    ['Fonte', 'IndicePA – Indice dei domicili digitali della PA (AgID), Open Data: ' + ORIGINE + '/ipa-dati/'],
    [],
    [cat.etichetta, elenco.length],
    [cat.etichetta + ' senza domicilio digitale principale', senzaDomicilio],
    [cat.etichetta + ' senza provincia/regione (non trovati in amministrazioni.txt)', senzaRegione],
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
  foglio(shEnti, cat.foglio, cat.larghezze);
  foglio(shPec, 'PEC', [36, 9, 20, 16, 24, 40, 12, 42, 6, 12, 36]);
  foglio(shRiep, 'Riepilogo', [62, 90]);
  const file = path.join(USCITA, cat.nomeFile);
  XLSX.writeFile(wb, file);
  log(`${cat.etichetta}: ${elenco.length}, PEC ${shPec.length - 1} → ${file}`);

  return [
    `## ${cat.titolo}`, '',
    `- File: [\`${cat.nomeFile}\`](${cat.nomeFile})`,
    `- ${cat.etichetta}: **${elenco.length}** (senza domicilio digitale principale: ${senzaDomicilio}; senza provincia/regione: ${senzaRegione})`,
    `- Indirizzi PEC (senza doppioni): **${shPec.length - 1}**`,
    ...LIVELLI.map(l => `  - ${l}: ${conta(l)}`), '',
  ].join('\n');
}

// ---------------- avvio ----------------
async function main() {
  const t0 = Date.now();
  const dati = {
    enti: await leggiRisorsa('enti'),
    amm: await leggiAmministrazioni(),
    domicili: await leggiRisorsa('domicili'),
    domiciliUO: await leggiRisorsa('domiciliUO'),
    uo: await leggiRisorsa('uo'),
  };
  fs.mkdirSync(USCITA, { recursive: true });
  const parti = CATEGORIE.map(c => elabora(c, dati));
  const dataIt = new Date().toLocaleDateString('it-IT', { timeZone: 'Europe/Rome' });
  const md = ['# Ultima estrazione', '', `Data: **${dataIt}** – Fonte: IndicePA (AgID), Open Data.`, '', ...parti].join('\n');
  fs.writeFileSync(path.join(USCITA, 'riepilogo.md'), md);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, md);
  log(`Fatto in ${Math.round((Date.now() - t0) / 1000)} s`);
}

main().catch(e => { console.error('[PEC] ERRORE:', e.message); process.exit(1); });
