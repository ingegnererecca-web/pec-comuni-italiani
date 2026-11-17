# Istruzioni per Claude – progetto «PEC dei Comuni italiani»

File di sincronizzazione per chi collabora al progetto con la propria istanza di Claude. Leggilo prima di modificare il codice e aggiornalo quando cambia qualcosa di rilevante.

## Scopo
Scaricare gli Open Data di IndicePA e produrre un Excel con le PEC di tutti i Comuni italiani (ente, AOO, UO). Due versioni con la stessa logica:
- userscript Tampermonkey (`pec-comuni-italiani.user.js`), eseguito su `indicepa.gov.it`;
- script Node (`scripts/estrai-pec.mjs`), eseguito da GitHub Actions **solo a mano** (`workflow_dispatch`), che salva l'Excel in `dati/`.

## Struttura del repository
| File | Ruolo |
|---|---|
| `pec-comuni-italiani.user.js` | Userscript Tampermonkey |
| `scripts/estrai-pec.mjs` | Stessa logica in Node 20 (fetch globale, pacchetto `xlsx`) |
| `.github/workflows/estrai-pec.yml` | Workflow manuale: esegue lo script Node, salva `dati/` con un commit del bot, allega l'Excel come artifact |
| `dati/PEC_Comuni_italiani.xlsx`, `dati/riepilogo.md` | Risultato dell'ultima estrazione (generati, non modificare a mano) |
| `README.md` | Istruzioni per gli utenti |
| `CLAUDE.md` | Questo file: contesto e regole per le istanze di Claude |
| `GRAPH.md` | Grafo del codice: flusso, chi chiama chi, indice delle funzioni con numeri di riga, strutture dati |

## Come funziona lo script
1. Pulsante fisso in basso a destra su ogni pagina di `indicepa.gov.it`.
2. `leggiRisorsa()` scarica i dump JSON CKAN `/ipa-dati/datastore/dump/<resource_id>?format=json` (formato `{fields:[{id}], records:[[…]]}`) con la stessa origine della pagina, senza permessi aggiuntivi.
3. Risorse (costante `RISORSE`):
   - Enti `d09adf99-dc10-4349-8c53-27b1e5aa97b6`
   - Domicili digitali (ente e AOO) `0226284e-27e5-4fdd-8bf9-4deba9f4d239`
   - Domicili digitali delle UO `c306c72c-77ab-4f37-8a3a-6f7888979003`
   - Unità Organizzative `b0aa1f6c-f135-4c8a-b416-396fed4e1a5d`
   - `amministrazioni.txt` (provincia e regione, unito per `cod_amm` = `Codice_IPA`)
4. Filtro Comuni: `Codice_Categoria === 'L6'` e (`Codice_IPA` tipo `c_a267` oppure denominazione «Comune di …»).
5. Raccolta PEC con priorità: domicilio ente → domicilio AOO → domicilio UO → Mail PEC dell'ente → Mail PEC delle UO. Doppioni eliminati per coppia (codice IPA, indirizzo).
6. Domicilio principale: domicilio a livello ente; se manca, AOO la cui descrizione contiene «domicilio digitale», poi la prima AOO, poi la prima PEC dell'ente.
7. Excel con SheetJS (`xlsx@0.18.5` via `@require`): fogli Comuni, PEC, Riepilogo.

## Regole del progetto
- Lingua: italiano (interfaccia, commenti, documentazione).
- Solo dati pubblici e solo lettura: lo script non deve mai inviare dati o scrivere su IndicePA.
- Nessun dato personale nel repository oltre al nome dell'account (`ingegnererecca-web`).
- Repository: https://github.com/ingegnererecca-web/pec-comuni-italiani (pubblico, ramo `main`).
- A ogni modifica del codice aumenta `@version` (gli utenti si aggiornano tramite `@updateURL`).
- Accesso ai campi sempre con `campo(oggetto, 'Nome')` (indipendente da maiuscole/minuscole): IndicePA può cambiare l'intestazione delle colonne.
- Prima di pubblicare: `node --check pec-comuni-italiani.user.js` e `node --check scripts/estrai-pec.mjs`.
- **Le due versioni vanno tenute allineate**: una modifica alla logica (filtri, campi, fogli Excel) va fatta in entrambi i file nello stesso commit.
- Il workflow parte solo a mano: non aggiungere `schedule` o altri trigger senza che l'utente lo chieda.
- Per orientarsi nel codice leggi prima `GRAPH.md`; se aggiungi, sposti o rinomini funzioni, aggiorna `GRAPH.md` nello stesso commit.

## Stato
- Workflow GitHub Actions manuale aggiunto (17/11/2026); prima esecuzione da verificare.
- v1.0.0: prima versione. La logica è stata provata con dati di esempio che riproducono le colonne dei dataset; la prima esecuzione sui dati reali va verificata (conteggio comuni atteso circa 7.900).

## Da fare
- Verificare la prima esecuzione sui dati reali.
