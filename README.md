# PEC dei Comuni italiani – userscript Tampermonkey

Userscript per Tampermonkey che produce un file Excel con le **PEC di tutti i Comuni italiani**, prese dagli **Open Data ufficiali di IndicePA** (Indice dei domicili digitali della Pubblica Amministrazione, gestito da AgID).

Per ogni comune lo script raccoglie:
- il **domicilio digitale** dell'ente (l'indirizzo da usare per le comunicazioni con valore legale);
- i domicili digitali dei **protocolli** (AOO) e degli **uffici** (UO);
- le altre caselle **PEC** indicate nelle schede dell'ente e degli uffici.

I doppioni vengono eliminati: ogni indirizzo compare una sola volta per comune.

## Senza estensione: il file Excel pronto su GitHub

Non serve installare nulla. Il file viene creato sui server di GitHub, solo quando lo chiedi tu.

1. Apri la scheda **[Actions](../../actions/workflows/estrai-pec.yml)** del repository.
2. Premi **«Run workflow»** e poi il pulsante verde **«Run workflow»** (ramo `main`).
3. Dopo qualche minuto l'esecuzione diventa verde.
4. Scarica l'Excel da **[`dati/PEC_Comuni_italiani.xlsx`](dati/PEC_Comuni_italiani.xlsx)**: premi il pulsante di download in alto a destra. I numeri dell'ultima estrazione sono in [`dati/riepilogo.md`](dati/riepilogo.md).

Il file si trova anche tra gli «Artifacts» di ogni esecuzione, conservati per 30 giorni. Il pulsante «Run workflow» compare solo a chi ha i permessi di scrittura sul repository.

## Con Tampermonkey (nel browser)

### Installazione

1. Installa l'estensione [Tampermonkey](https://www.tampermonkey.net/) in Chrome.
2. Abilita l'esecuzione degli script: `chrome://extensions` → Tampermonkey → **Dettagli** → attiva **«Consenti script utente»** (Chrome 138 e successivi).
3. Apri il file [`pec-comuni-italiani.user.js`](pec-comuni-italiani.user.js) in versione *Raw*: Tampermonkey propone l'installazione. Conferma.

Gli aggiornamenti arrivano da soli: Tampermonkey controlla periodicamente questo repository.

### Uso

1. Apri [indicepa.gov.it](https://indicepa.gov.it).
2. In basso a destra premi **«Estrai PEC dei Comuni»**.
3. Lo script scarica i dataset (alcune decine di MB, il riquadro mostra l'avanzamento) e poi il file `PEC_Comuni_italiani_<data>.xlsx`.

Se il download non parte, usa il pulsante **«Scarica di nuovo il file Excel»**.

## Contenuto del file Excel

| Foglio | Contenuto |
|---|---|
| **Comuni** | Una riga per comune: provincia, regione, codici ISTAT, catastale, IPA e fiscale, domicilio digitale principale, numero di PEC, sito |
| **PEC** | Una riga per indirizzo PEC: comune, livello (ente / AOO / UO), ufficio, indirizzo, data di pubblicazione, dataset di provenienza |
| **Riepilogo** | Data, fonti con i link ai dataset, conteggi |

Tutti i fogli hanno i filtri attivi: per esempio puoi filtrare per regione o provincia.

## Fonti

Dataset Open Data di IndicePA usati dallo script:
- Enti
- Domicili digitali (ente e AOO)
- Domicili digitali delle Unità Organizzative
- Unità Organizzative
- Amministrazioni (`amministrazioni.txt`, per provincia e regione)

Portale: <https://indicepa.gov.it/ipa-dati/>. I dati sono di AgID / IndicePA: riportane sempre la fonte.

## Uso corretto dei dati

Gli indirizzi sono pubblicati per legge (art. 6-ter del D.Lgs. 82/2005, Codice dell'Amministrazione Digitale) per le comunicazioni con le pubbliche amministrazioni. Usali per comunicazioni istituzionali e non per invii massivi indesiderati.

## Note tecniche

- Le due versioni usano la stessa logica: `pec-comuni-italiani.user.js` (browser) e `scripts/estrai-pec.mjs` (Node, usata da GitHub Actions con `.github/workflows/estrai-pec.yml`).
- Lo script legge solo dati pubblici: non invia nulla e non scrive su IndicePA.
- Un ente è considerato «Comune» se ha la categoria IPA `L6` e un codice IPA del tipo `c_xxxx` oppure una denominazione che inizia con «Comune di». Unioni e consorzi di comuni sono esclusi.
- Se IndicePA cambia i dataset o i nomi delle colonne, lo script può segnalare un errore nel riquadro: in quel caso apri una *issue*.
