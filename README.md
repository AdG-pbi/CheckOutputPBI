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

## Utilizzo

```bash
python compare_files.py file1.csv file2.xlsx -o report.xlsx
```

Per i file tabellari puoi passare una chiave record composta da più campi usando indici 1-based nel formato richiesto `1+5`:

```bash
python compare_files.py prima.csv seconda.csv -k 1+5 -o differenze.xlsx
```

Nei file Excel (`.xlsx`/`.xlsm`) vengono considerati anche workbook con più fogli: il confronto viene eseguito per nome foglio.

## Output

Il file Excel prodotto contiene:

- `Summary`: riepilogo del confronto
- `Differences`: dettaglio delle differenze con evidenziazione colore

Per confronti testuali, il dettaglio include anche la posizione del punto differente nei due file e due colonne di contesto con il testo immediatamente precedente e successivo. In questo modo, per input `.pdf` e `.docx` è più semplice individuare il passaggio corretto senza affidarsi solo al conteggio di righe o paragrafi.

Quando nei `.pdf` o `.docx` vengono intercettate tabelle, la posizione riporta anche il titolo della tabella e le coordinate della cella (`riga` e `colonna`) dove è stata rilevata la differenza.

Se `file2` è `.pdf` o `.docx`, oltre al report tabellare viene generato automaticamente anche un PDF (`<nome_output>_file2_highlight.pdf`) basato sul contenuto di `file2`, con evidenziazione gialla dei caratteri/frasi/numeri che differiscono rispetto a `file1`.

Per i `.pdf`, il parser prova anche un'estrazione in modalità "layout" per preservare meglio l'allineamento visivo del contenuto: questo aiuta a riconoscere correttamente tabelle che, in estrazione testuale semplice, verrebbero appiattite come testo normale.

Se il dettaglio supera il limite di righe supportato da Excel, il report viene salvato automaticamente in formato `.csv` nello stesso percorso richiesto, sostituendo l'estensione del file di output.
