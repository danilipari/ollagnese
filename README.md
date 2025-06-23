# Script di Compilazione Questionari

Questo script automatizza la compilazione di questionari utilizzando Ollama, un modello di linguaggio locale. Il sistema legge domande da un file Excel, invia ciascuna domanda a Ollama, elabora le risposte e le salva in un nuovo file Excel.

> **IMPORTANTE**: Si consiglia di utilizzare la versione corretta dello script `script_cq.py` che risolve diversi problemi di compatibilità e aggiunge funzionalità di gestione errori.

## Requisiti

- Python 3.x
- Ollama installato e in esecuzione localmente con il modello scelto nel py
- Un file Excel di input con domande nel formato corretto

## Dipendenze Python

Lo script richiede le seguenti librerie Python:
- pandas
- requests
- openpyxl

## Installazione

1. Crea un ambiente virtuale:
```bash
python3 -m venv ollama_env
source ollama_env/bin/activate  # Su Windows: ollama_env\Scripts\activate
```

2. Installa le dipendenze necessarie:
```bash
pip install pandas requests openpyxl
```

## Configurazione

Il script utilizza le seguenti impostazioni configurabili all'inizio del file:

```python
INPUT_FILE = "database_questionari_inglese.xlsx"  # Nome del file Excel di input
OUTPUT_FILE = "questionari_con_risposte.xlsx"     # Nome del file Excel di output
OLLAMA_MODEL = "llama3" # esempio                           # Modello Ollama da utilizzare
OLLAMA_URL = "http://localhost:11434/api/generate" # URL API di Ollama
AUTOSAVE_INTERVAL = 50                            # Frequenza di salvataggio automatico
MOTIVAZIONE_INTERVAL = 100                        # Frequenza generazione motivazione nell'excel
```

## Formato del File di Input

Il file Excel di input deve contenere almeno una colonna chiamata "Prompt" con le domande da inviare a Ollama. Il script creerà automaticamente le colonne "Risposta numerica" e "Motivazione" se non esistono già.

## Esecuzione

Per eseguire lo script:

```bash
source ollama_env/bin/activate  # Attiva l'ambiente virtuale
python script_cq.py  # Versione migliorata consigliata
```

Alternativamente, puoi usare la versione originale (non consigliata):
```bash
python script_compilazione_questionari.py
```

## Funzionamento

Lo script:
1. Legge il file Excel di input
2. Per ogni riga con un prompt:
   - Invia il prompt a Ollama
   - Estrae dal testo di risposta un valore numerico e una motivazione
   - Salva i risultati nel dataframe
3. Salva automaticamente il file ogni N righe (definito da AUTOSAVE_INTERVAL)
4. Al termine, salva il file finale con le risposte

## Output

Durante l'esecuzione, lo script mostra:
- Progresso dell'elaborazione riga per riga
- Notifiche di salvataggio automatico
- Eventuali errori di comunicazione con Ollama

Il risultato finale è un file Excel contenente tutte le domande originali più due colonne aggiuntive:
- "Risposta numerica": il valore numerico estratto dalla risposta
- "Motivazione": il testo che spiega il ragionamento dietro la risposta
