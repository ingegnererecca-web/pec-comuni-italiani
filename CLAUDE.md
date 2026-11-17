# Istruzioni per Claude – progetto «PEC dei Comuni e delle Scuole italiane»

File di sincronizzazione per chi collabora al progetto con la propria istanza di Claude. Leggilo prima di modificare il codice e aggiornalo quando cambia qualcosa di rilevante.

## Scopo
Scaricare gli Open Data di IndicePA e produrre due Excel: PEC di tutti i Comuni e di tutte le Scuole statali italiane (ente, AOO, UO). Le categorie sono definite in `CATEGORIE`. Due versioni con la stessa logica:
- userscript Tampermonkey (`pec-comuni-italiani.user.js`), eseguito su `indicepa.gov.it`;
- script Node (`scripts/estrai-pec.mjs`), eseguito da GitHub Actions **solo a mano** (`workflow_dispatch`), che salva gli Excel in `dati/`.

## Struttura del repository
| File | Ruolo |
|---|---|
| `pec-comuni-italiani.user.js` | Userscript Tampermonkey |
| `scripts/estrai-pec.mjs` | Stessa logica in Node 20 (fetch globale, pacchetto `xlsx`) |
| `.github/workflows/estrai-pec.yml` | Workflow manuale: esegue lo script Node, salva `dati/` con un commit del bot, allega l'Excel come artifact |
| `dati/PEC_Comuni_italiani.xlsx`, `dati/PEC_Scuole_italiane.xlsx`, `dati/riepilogo.md` | Risultato dell'ultima estrazione (generati, non modificare a mano) |
| `README.md` | Istruzioni per gli utenti |
| `CLAUDE.md` | Questo file: contesto e regole per le istanze di Claude |
| `GRAPH.md` | Grafo del codice: flusso, chi chiama chi, indice delle funzioni con numeri di riga, strutture dati |

## Come funziona lo script
Dettagli e numeri di riga in `GRAPH.md`. In sintesi:
1. `avvia()` (userscript) / `main()` (Node) scaricano **una sola volta** i dati: Enti, `amministrazioni.txt`, Domicili digitali (ente/AOO), Domicili digitali UO, Unità Organizzative (ID in `RISORSE`, dump JSON CKAN `/ipa-dati/datastore/dump/<id>?format=json`).
2. Per ogni voce di `CATEGORIE` si chiama `elabora(cat, dati)`:
   - **Comuni**: `Codice_Categoria === 'L6'` e (`Codice_IPA` tipo `c_a267` oppure «Comune di …»);
   - **Scuole**: `Codice_MIUR` presente (istituti statali).
3. Raccolta PEC con priorità: domicilio ente → domicilio AOO → domicilio UO → Mail PEC dell'ente → Mail PEC delle UO. Doppioni eliminati per coppia (codice IPA, indirizzo).
4. Domicilio principale: livello ente; se manca, l'AOO con «domicilio digitale» nella descrizione, poi la prima AOO, poi la prima PEC dell'ente.
5. Excel con SheetJS (`xlsx@0.18.5`): foglio dell'ente (Comuni/Scuole), PEC, Riepilogo.

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
- v1.1.0 (17/11/2026): aggiunte le Scuole statali (`Codice_MIUR`), file `PEC_Scuole_italiane.xlsx`; codice unificato con `CATEGORIE` + `elabora()`. Prima estrazione delle scuole da verificare.
- Prima estrazione reale dei Comuni (17/11/2026, workflow): 7.893 comuni, tutti con domicilio digitale, 19.167 PEC. 0 domicili a livello UO: da verificare se è un dato reale o un nome di colonna diverso.
- v1.0.0: prima versione (solo Comuni) e workflow GitHub Actions manuale.

## Da fare
- Verificare la prima estrazione delle Scuole (atteso: circa 8.000 istituti statali).
- Verificare i domicili digitali delle UO (colonne del dataset).
