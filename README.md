# CheckOutputPBI

Strumento CLI per confrontare due file e generare un report Excel con le differenze.

## Formati supportati

- tabellari: `csv`, `xlsx`, `xlsm`
- testuali/documentali: `txt`, `pdf`, `docx`
- altri file testuali leggibili come UTF-8

> Nota: i file `.doc` legacy non hanno un parser affidabile incluso nel progetto; convertili in `.docx` o `.pdf` prima del confronto.

## Installazione

```bash
python -m pip install -r requirements.txt
```

Per migliorare le prestazioni del confronto visivo PDF, il progetto usa anche `numpy` (installato tramite `requirements.txt`).

## Utilizzo

```bash
python compare_files.py file1.csv file2.xlsx -o report.xlsx
```

Per i file tabellari puoi passare una chiave record composta da più campi usando indici 1-based nel formato richiesto `1+5`:

```bash
python compare_files.py prima.csv seconda.csv -k 1+5 -o differenze.xlsx
```

Nei file Excel (`.xlsx`/`.xlsm`) vengono considerati anche workbook con più fogli: il confronto viene eseguito per nome foglio.

Per workbook multi-sheet puoi anche ripetere `--key` e associare una chiave a uno sheet specifico indicando il suo numero 1-based:

```bash
python compare_files.py prima.xlsx seconda.xlsx --key 1:1 --key 2:2+4 -o differenze.xlsx
```

Se passi una chiave senza prefisso (`--key 1+5`), quella viene usata come default per tutti gli sheet che non hanno una chiave dedicata.

## Output

Il file Excel prodotto contiene:

- `Summary`: riepilogo del confronto
- `Differences`: dettaglio delle differenze con evidenziazione colore

Per confronti Excel multi-sheet, il report aggiunge anche un foglio dedicato per ogni foglio sorgente confrontato, così puoi verificare subito differenze e assenza di differenze per ciascun tab.

Per confronti testuali, il dettaglio include anche la posizione del punto differente nei due file e due colonne di contesto con il testo immediatamente precedente e successivo. In questo modo, per input `.pdf` e `.docx` è più semplice individuare il passaggio corretto senza affidarsi solo al conteggio di righe o paragrafi.

Quando nei `.pdf` o `.docx` vengono intercettate tabelle, la posizione riporta anche il titolo della tabella e le coordinate della cella (`riga` e `colonna`) dove è stata rilevata la differenza.

Se `file2` è `.pdf` o `.docx`, oltre al report tabellare viene generato automaticamente anche un PDF (`<nome_output>_file2_highlight.pdf`) basato sul contenuto di `file2`, con evidenziazione gialla delle differenze rispetto a `file1`.

Per `file2` in formato `.pdf`, il PDF evidenziato mantiene lo stesso layout dell'originale: viene riutilizzato direttamente il file `file2` e vengono aggiunte solo le evidenziazioni.

Per i `.pdf`, il parser prova anche un'estrazione in modalità "layout" per preservare meglio l'allineamento visivo del contenuto: questo aiuta a riconoscere correttamente tabelle che, in estrazione testuale semplice, verrebbero appiattite come testo normale.

Se due `.pdf` hanno differenze visive molto piccole (per esempio grafici, immagini o testo vettoriale che non viene intercettato bene dall'estrazione testuale), il PDF evidenziato aggiunge anche riquadri rossi nelle aree della pagina che risultano cambiate.

Se il dettaglio supera il limite di righe supportato da Excel, il report viene salvato automaticamente in formato `.csv` nello stesso percorso richiesto, sostituendo l'estensione del file di output.

## Benchmark prestazioni PDF

Per misurare rapidamente il tempo di confronto PDF (small/medium/large e scenario visual-only):

```bash
python benchmarks/benchmark_pdf_compare.py --iterations 3
```

Lo script stampa tempi medi separati per confronto (`auto_compare`) e generazione PDF evidenziato.
