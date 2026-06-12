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

## Output

Il file Excel prodotto contiene:

- `Summary`: riepilogo del confronto
- `Differences`: dettaglio delle differenze con evidenziazione colore
