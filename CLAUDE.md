# Word (ex Parola)

## Cos'è
Elaboratore di testi nel browser, stile Word: schede Home, Inserisci, Layout, Revisione, Visualizza (+ Tabella contestuale), gruppi con etichette, ricerca dei comandi (⌘J), riquadro Struttura con i titoli, modalità Concentrazione, barra riducibile (doppio clic su una scheda), tema chiaro/scuro/auto.
Salva da solo nel browser (localStorage), scarica .docx (libreria docx), apre .docx (mammoth), .html, .txt e .md. Esporta anche PDF (stampa), .html, .md, .txt.
Si usa anche come app per Mac (cartella `mac/`: `./mac/build.sh` crea e installa `/Applications/Word.app`, Apple Silicon, e toglie la vecchia Parola.app).
Fino a novembre 2026 si chiamava «Parola»; dal 22/11/2026 anche cartella, repository e indirizzo online si chiamano `word`.

## Dove sta online
Repository: https://github.com/nemesiMark/word (ramo `main`, pubblico).
Online con GitHub Pages: https://nemesimark.github.io/word/ (si aggiorna a ogni push su main).
Compare tra i lavori del sito Nemesi Web (~/Desktop/Sites/nemesiWeb) come «Word».

## Come si prova in locale
- Aprire `index.html` nel browser (doppio clic), oppure `python3 -m http.server 8000` nella cartella.
- Prove automatiche (128 controlli, Chrome senza finestra): `node test/prove.mjs`. Servono Google Chrome e `puppeteer-core` (preso da `~/.claude/skills/sito-film/scripts/node_modules`, oppure da `PUPPETEER_DIR`). Vanno lanciate prima di ogni commit che tocca `app.js`, `docx-export.js`, `index.html` o `style.css`: se falliscono, non si fa il commit.

## Struttura
- `index.html` — interfaccia (icone SVG in testa, nastro, menu, finestre `<dialog>`)
- `style.css` — aspetto; colori come variabili su `:root`, tema scuro con `prefers-color-scheme` e `[data-theme]`
- `app.js` — logica; `window.__word` espone alcune funzioni solo per le prove
- `docx-export.js` — conversione in .docx (`window.WordDocx`)
- `icon.svg` — icona (pagina blu a strisce + tessera con la W davanti, stile Office); è anche la favicon e la sorgente dell'icona Mac
- `vendor/` — librerie in copia locale (docx, mammoth)
- `test/prove.mjs` — prove automatiche
- `mac/` — involucro Swift/WebKit per l'app Mac (`mac/build/` escluso da git); `icon.swift` trasforma `icon.svg` in PNG

## Note tecniche
- Archivio nel browser: `word.documento` e `word.preferenze`; se mancano si leggono i vecchi `parola.*` (nessun dato perso col cambio nome).
- App Mac: identificativo `it.nemesi.word` (prima `it.parola.editor`); la prima installazione copia `~/Library/WebKit/it.parola.editor` in `it.nemesi.word`. Ponte pagina↔app: `window.webkit.messageHandlers.word` (messaggi `save`, `print`, `title`, `theme`) e `window.word` (`toast`, `command`, `openBase64`, `saved`).
- Ricerca: le evidenziazioni usano la CSS Custom Highlight API (dove manca, la ricerca funziona lo stesso senza evidenziare tutti i risultati).

## Regole di lavoro
- Testi in italiano.
- Commit solo su richiesta; «g» = prove automatiche, poi commit e push su `origin main`.
- Dopo modifiche all'editor, ricostruire l'app Mac con `./mac/build.sh`.

## In sospeso
- Aggiornare nome e immagine nella scheda del sito Nemesi Web (lì è ancora «Parola»).
