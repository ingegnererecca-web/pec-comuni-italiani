// ==UserScript==
// @name         PEC dei Comuni italiani – estrazione da IndicePA
// @namespace    https://github.com/ingegnererecca-web/pec-comuni-italiani
// @version      1.0.0
// @description  Scarica gli Open Data ufficiali di IndicePA (AgID) e produce un file Excel con le PEC di tutti i Comuni italiani: domicilio digitale dell'ente, dei protocolli (AOO) e degli uffici (UO).
// @author       ingegnererecca-web
// @match        https://indicepa.gov.it/*
// @match        https://www.indicepa.gov.it/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @updateURL    https://raw.githubusercontent.com/ingegnererecca-web/pec-comuni-italiani/main/pec-comuni-italiani.user.js
// @downloadURL  https://raw.githubusercontent.com/ingegnererecca-web/pec-comuni-italiani/main/pec-comuni-italiani.user.js
// @grant        none
// @run-at       document-idle
// @noframes
// ==/UserScript==

/*
  USO
  1. Apri https://indicepa.gov.it (qualsiasi pagina del sito).
  2. In basso a destra premi "Estrai PEC dei Comuni".
  3. Lo script scarica dagli Open Data di IndicePA i dataset ufficiali (Enti, Domicili digitali,
     Domicili digitali delle UO, Unità Organizzative, Amministrazioni), tiene solo i Comuni e
     scarica il file "PEC_Comuni_italiani_<data>.xlsx" con tre fogli:
       - Comuni:    una riga per comune, con il domicilio digitale principale;
       - PEC:       una riga per ogni indirizzo PEC (ente, protocollo/AOO, uffici/UO);
       - Riepilogo: fonti, data, conteggi.
  Lo script legge soltanto dati pubblici: non invia nulla e non scrive su IndicePA.
  I dati sono di IndicePA (AgID): riportarne sempre la fonte.
*/

(function () {
  'use strict';

  // ---------------- risorse Open Data IndicePA ----------------
  const RISORSE = {
    enti:          { id: 'd09adf99-dc10-4349-8c53-27b1e5aa97b6', nome: 'Enti' },
    domicili:      { id: '0226284e-27e5-4fdd-8bf9-4deba9f4d239', nome: 'Domicili digitali (ente e AOO)' },
    domiciliUO:    { id: 'c306c72c-77ab-4f37-8a3a-6f7888979003', nome: 'Domicili digitali delle Unità Organizzative' },
    uo:            { id: 'b0aa1f6c-f135-4c8a-b416-396fed4e1a5d', nome: 'Unità Organizzative' },
  };
  const AMMINISTRAZIONI_TXT = '/ipa-dati/dataset/502ff370-1b2c-4310-94c7-f39ceb7500e3/resource/3ed63523-ff9c-41f6-a6fe-980f3d9e501f/download/amministrazioni.txt';
  const urlDump = id => `${location.origin}/ipa-dati/datastore/dump/${id}?format=json`;

  // ---------------- interfaccia ----------------
  const box = document.createElement('div');
  box.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;background:#fff;border:1px solid #0b4f8a;border-radius:8px;padding:10px 12px;font:13px Arial,sans-serif;box-shadow:0 2px 10px rgba(0,0,0,.25);max-width:340px;color:#222';
  const btn = document.createElement('button');
  btn.textContent = 'Estrai PEC dei Comuni';
  btn.style.cssText = 'background:#0b4f8a;color:#fff;border:0;border-radius:5px;padding:8px 12px;cursor:pointer;font-weight:bold';
  const btnFile = document.createElement('button');
  btnFile.textContent = 'Scarica di nuovo il file Excel';
  btnFile.style.cssText = 'margin-top:6px;background:#1f4e79;color:#fff;border:0;border-radius:5px;padding:6px 10px;cursor:pointer;display:none';
  const stato = document.createElement('div');
  stato.style.cssText = 'margin-top:6px;white-space:pre-line';
  box.append(btn, document.createElement('br'), btnFile, stato);
  document.body.appendChild(box);
  const log = t => { stato.textContent = t; console.log('[PEC Comuni]', t); };

  let fileExcel = null;
  function scarica(f) {
    const url = URL.createObjectURL(f.blob);
    const a = document.createElement('a');
    a.href = url; a.download = f.nome; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  btnFile.addEventListener('click', () => fileExcel && scarica(fileExcel));
  btn.addEventListener('click', () => avvia().catch(e => { log('Errore: ' + e.message); btn.disabled = false; }));

  // ---------------- utilità ----------------
  const pulisci = s => (s == null ? '' : String(s)).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  const isPec = tipo => /pec/i.test(pulisci(tipo));
  const isIndirizzo = s => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(pulisci(s));

  // scarica un testo mostrando l'avanzamento in MB
  async function scaricaTesto(url, etichetta) {
    const r = await fetch(url, { credentials: 'omit' });
    if (!r.ok) throw new Error(`${etichetta}: risposta ${r.status}`);
    if (!r.body || !r.body.getReader) return r.text();
    const reader = r.body.getReader(), parti = [];
    let byte = 0, ultimo = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      parti.push(value); byte += value.length;
      if (byte - ultimo > 2e6) { ultimo = byte; log(`Scarico «${etichetta}»… ${(byte / 1e6).toFixed(0)} MB`); }
    }
    const tutto = new Uint8Array(byte); let off = 0;
    for (const p of parti) { tutto.set(p, off); off += p.length; }
    return new TextDecoder('utf-8').decode(tutto);
  }

  // dump JSON di CKAN: { fields: [{id}], records: [[...]] } -> array di oggetti
  async function leggiRisorsa(chiave) {
    const ris = RISORSE[chiave];
    log(`Scarico «${ris.nome}»…`);
    const testo = await scaricaTesto(urlDump(ris.id), ris.nome);
    const j = JSON.parse(testo);
    const campi = (j.fields || []).map(f => f.id);
    const righe = (j.records || []).map(rec => Array.isArray(rec) ? Object.fromEntries(campi.map((c, i) => [c, rec[i]])) : rec);
    log(`«${ris.nome}»: ${righe.length} righe.`);
    return righe;
  }

  // accesso ai campi senza dipendere da maiuscole/minuscole
  const campo = (o, ...nomi) => {
    for (const n of nomi) {
      if (o[n] != null) return pulisci(o[n]);
      const k = Object.keys(o).find(x => x.toLowerCase() === n.toLowerCase());
      if (k && o[k] != null) return pulisci(o[k]);
    }
    return '';
  };

  // amministrazioni.txt: file di testo con separatore (di norma tabulazione) e riga di intestazione
  async function leggiAmministrazioni() {
    try {
      const testo = await scaricaTesto(location.origin + AMMINISTRAZIONI_TXT, 'Amministrazioni');
      const righe = testo.split(/\r?\n/).filter(Boolean);
      const sep = righe[0].includes('\t') ? '\t' : (righe[0].includes(';') ? ';' : ',');
      const intest = righe[0].split(sep).map(pulisci);
      const iCod = intest.findIndex(h => /^cod_amm$/i.test(h));
      const iCom = intest.findIndex(h => /^comune$/i.test(h));
      const iPro = intest.findIndex(h => /^provincia$/i.test(h));
      const iReg = intest.findIndex(h => /^regione$/i.test(h));
      const mappa = new Map();
      if (iCod < 0) return mappa;
      for (const r of righe.slice(1)) {
        const c = r.split(sep);
        mappa.set(pulisci(c[iCod]).toLowerCase(), { comune: pulisci(c[iCom]), provincia: pulisci(c[iPro]), regione: pulisci(c[iReg]) });
      }
      return mappa;
    } catch (e) {
      console.warn('[PEC Comuni] amministrazioni.txt non letto:', e);
      return new Map();
    }
  }

  // ---------------- elaborazione ----------------
  async function avvia() {
    btn.disabled = true; btnFile.style.display = 'none';
    const t0 = Date.now();

    // 1) Enti: si tengono solo i Comuni
    const enti = await leggiRisorsa('enti');
    const comuni = new Map();   // codice IPA (minuscolo) -> scheda
    for (const e of enti) {
      const ipa = campo(e, 'Codice_IPA');
      const den = campo(e, 'Denominazione_ente');
      const cat = campo(e, 'Codice_Categoria');
      const eComune = cat === 'L6' && (/^c_[a-z]\d{3}$/i.test(ipa) || /^comune di /i.test(den));
      if (!eComune) continue;
      comuni.set(ipa.toLowerCase(), {
        ipa, denominazione: den, cf: campo(e, 'Codice_fiscale_ente'),
        istat: campo(e, 'Codice_comune_ISTAT'), catastale: campo(e, 'Codice_catastale_comune'),
        sito: campo(e, 'Sito_istituzionale'), comune: den.replace(/^comune di\s+/i, ''),
        provincia: '', regione: '', domicilio: '', pec: [],
        mailEnte: [1, 2, 3, 4, 5].map(n => ({ m: campo(e, 'Mail' + n), t: campo(e, 'Tipo_Mail' + n) })),
      });
    }
    if (!comuni.size) throw new Error('nessun comune trovato nel dataset Enti (il formato potrebbe essere cambiato)');
    log(`Comuni trovati: ${comuni.size}.`);

    // 2) provincia e regione
    const amm = await leggiAmministrazioni();
    for (const c of comuni.values()) {
      const a = amm.get(c.ipa.toLowerCase());
      if (a) { c.provincia = a.provincia; c.regione = a.regione; if (a.comune) c.comune = a.comune; }
    }

    // 3) raccolta delle PEC, senza doppioni (per comune e indirizzo)
    const visti = new Set();
    const aggiungi = (c, r) => {
      const pec = pulisci(r.pec).toLowerCase();
      if (!isIndirizzo(pec)) return;
      const k = c.ipa + '|' + pec;
      if (visti.has(k)) return;
      visti.add(k);
      c.pec.push({ ...r, pec });
    };

    // 3a) domicili digitali dell'ente e dei protocolli (AOO)
    for (const d of await leggiRisorsa('domicili')) {
      const c = comuni.get(campo(d, 'Codice_IPA').toLowerCase());
      if (!c) continue;
      const aoo = campo(d, 'Codice_uni_aoo');
      const r = { livello: aoo ? 'AOO – domicilio digitale' : 'Ente – domicilio digitale', ufficio: aoo ? campo(d, 'Descrizione_aoo') : 'Ente', codiceUfficio: aoo,
        pec: campo(d, 'Domicilio_digitale'), tipo: campo(d, 'Tipo_domicilio_digitale') || 'PEC', data: campo(d, 'Data_pubblicazione'), fonte: RISORSE.domicili.nome };
      if (!aoo && !c.domicilio) c.domicilio = pulisci(r.pec).toLowerCase();
      aggiungi(c, r);
    }

    // domicilio principale: in IPA spesso è registrato sull'AOO ("Domicilio digitale del Comune di …")
    for (const c of comuni.values()) {
      if (c.domicilio) continue;
      const aoo = c.pec.filter(p => p.livello === 'AOO – domicilio digitale');
      const p = aoo.find(x => /domicilio digitale/i.test(x.ufficio)) || aoo[0];
      if (p) c.domicilio = p.pec;
    }

    // 3b) domicili digitali degli uffici (UO)
    for (const d of await leggiRisorsa('domiciliUO')) {
      const c = comuni.get(campo(d, 'Codice_IPA').toLowerCase());
      if (!c) continue;
      aggiungi(c, { livello: 'UO – domicilio digitale', ufficio: campo(d, 'Descrizione_uo', 'Descrizione_ufficio'), codiceUfficio: campo(d, 'Codice_uni_uo'),
        pec: campo(d, 'Domicilio_digitale'), tipo: campo(d, 'Tipo_domicilio_digitale') || 'PEC', data: campo(d, 'Data_pubblicazione'), fonte: RISORSE.domiciliUO.nome });
    }

    // 3c) caselle PEC indicate nella scheda dell'ente
    for (const c of comuni.values()) {
      c.mailEnte.filter(x => isPec(x.t)).forEach(x => aggiungi(c, { livello: 'Ente – casella PEC', ufficio: 'Ente', codiceUfficio: '', pec: x.m, tipo: 'PEC', data: '', fonte: RISORSE.enti.nome }));
      if (!c.domicilio) { const p = c.mailEnte.find(x => isPec(x.t) && isIndirizzo(x.m)); if (p) c.domicilio = p.m.toLowerCase(); }
    }

    // 3d) caselle PEC dei singoli uffici (UO)
    for (const u of await leggiRisorsa('uo')) {
      const c = comuni.get(campo(u, 'Codice_IPA').toLowerCase());
      if (!c) continue;
      for (const n of [1, 2, 3]) {
        if (!isPec(campo(u, 'Tipo_Mail' + n))) continue;
        aggiungi(c, { livello: 'UO – casella PEC', ufficio: campo(u, 'Descrizione_uo'), codiceUfficio: campo(u, 'Codice_uni_uo'),
          pec: campo(u, 'Mail' + n), tipo: 'PEC', data: '', fonte: RISORSE.uo.nome });
      }
    }

    // 4) file Excel
    log('Creo il file Excel…');
    const elenco = [...comuni.values()].sort((a, b) =>
      (a.regione || '~').localeCompare(b.regione || '~', 'it') || (a.provincia || '~').localeCompare(b.provincia || '~', 'it') || a.comune.localeCompare(b.comune, 'it'));
    const oggi = new Date();
    const dataIt = oggi.toLocaleDateString('it-IT');

    const shComuni = [['Comune', 'Provincia', 'Regione', 'Codice ISTAT', 'Codice catastale', 'Codice IPA', 'Codice fiscale', 'Domicilio digitale principale', 'N. indirizzi PEC', 'Sito istituzionale'],
      ...elenco.map(c => [c.comune, c.provincia, c.regione, c.istat, c.catastale, c.ipa, c.cf, c.domicilio, c.pec.length, c.sito])];
    const shPec = [['Comune', 'Provincia', 'Regione', 'Codice IPA', 'Livello', 'Ufficio', 'Codice ufficio', 'Indirizzo PEC', 'Tipo', 'Data pubblicazione', 'Fonte (dataset IndicePA)']];
    elenco.forEach(c => c.pec.forEach(p => shPec.push([c.comune, c.provincia, c.regione, c.ipa, p.livello, p.ufficio, p.codiceUfficio, p.pec, p.tipo, p.data, p.fonte])));

    const conta = l => shPec.filter(r => r[4] === l).length;
    const senzaDomicilio = elenco.filter(c => !c.domicilio).length;
    const shRiep = [
      ['PEC dei Comuni italiani – estrazione da IndicePA'],
      ['Data estrazione', dataIt],
      ['Fonte', 'IndicePA – Indice dei domicili digitali della PA (AgID), Open Data: ' + location.origin + '/ipa-dati/'],
      [],
      ['Comuni', elenco.length],
      ['Comuni senza domicilio digitale principale', senzaDomicilio],
      ['Comuni senza provincia/regione (non trovati in amministrazioni.txt)', elenco.filter(c => !c.regione).length],
      ['Indirizzi PEC totali (senza doppioni)', shPec.length - 1],
      ...['Ente – domicilio digitale', 'AOO – domicilio digitale', 'UO – domicilio digitale', 'Ente – casella PEC', 'UO – casella PEC'].map(l => ['  ' + l, conta(l)]),
      [],
      ['Dataset usati'],
      ...Object.values(RISORSE).map(r => [r.nome, urlDump(r.id)]),
      ['Amministrazioni (provincia e regione)', location.origin + AMMINISTRAZIONI_TXT],
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
    const dati = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    fileExcel = { blob: new Blob([dati], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      nome: `PEC_Comuni_italiani_${oggi.toISOString().slice(0, 10)}.xlsx` };
    scarica(fileExcel);
    btnFile.style.display = 'inline-block';
    btn.disabled = false;
    log(`Fatto in ${Math.round((Date.now() - t0) / 1000)} s.\nComuni: ${elenco.length} (senza domicilio digitale: ${senzaDomicilio}).\nIndirizzi PEC: ${shPec.length - 1}.\nFile Excel scaricato.`);
  }
})();
