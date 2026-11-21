# Parola

## Cos'è
Editor di testi nel browser (stile elaboratore di testi): schede Home, Inserisci, Layout, Revisione, Visualizza; salva da solo nel browser e scarica .docx (libreria docx), apre .docx (mammoth).
Si usa anche come app per Mac (cartella `mac/`: `./mac/build.sh` crea e installa Parola.app, Apple Silicon).

## Dove sta online
Repository: https://github.com/nemesiMark/parola (ramo `main`, pubblico).
Online con GitHub Pages: https://nemesimark.github.io/parola/ (si aggiorna a ogni push su main).
Compare tra i lavori del sito Nemesi Web (~/Desktop/Sites/site-nemesi).

## Come si prova in locale
Basta aprire `index.html` nel browser (doppio clic), oppure `python3 -m http.server 8000` nella cartella.

## Struttura
- `index.html`, `style.css`, `app.js` — interfaccia e logica
- `vendor/` — librerie in copia locale
- `mac/` — involucro Swift/WebKit per l'app Mac (`mac/build/` escluso da git)

## Regole di lavoro
- Testi in italiano.
- Commit solo su richiesta; «g» = commit e push su `origin main`.
