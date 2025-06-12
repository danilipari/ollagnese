import pandas as pd
import requests
import time
import warnings
import os

# Ignora i warning di pandas sulla compatibilità dei tipi
warnings.filterwarnings('ignore', category=FutureWarning)


INPUT_FILE = "database_questionari_inglese.xlsx"
OUTPUT_FILE = "questionari_con_risposte.xlsx"
OLLAMA_MODEL = "llama3:latest" # "llama3.1:latest"
OLLAMA_URL = "http://localhost:11434/api/generate"
AUTOSAVE_INTERVAL = 10
PROMPT_TEMPLATE = """
{prompt}

Rispondi nel formato seguente:
Risposta: [inserisci un numero]
Motivazione: [inserisci la tua spiegazione]
"""

def interroga_ollama(prompt):
    # Formatta il prompt per ottenere risposte strutturate
    prompt_formattato = PROMPT_TEMPLATE.format(prompt=prompt.strip())
    
    payload = {
        "model": OLLAMA_MODEL,
        "prompt": prompt_formattato,
        "stream": False
    }

    try:
        response = requests.post(OLLAMA_URL, json=payload)
        response.raise_for_status()
        testo = response.json().get("response", "")
        
        numero = ""
        motivazione = ""
        
        lines = testo.splitlines()
        parsing_motivazione = False
        motivazione_list = []
        
        for line in lines:
            line = line.strip()
            if line.lower().startswith("risposta:"):
                numero = line.split(":", 1)[-1].strip()
            elif "motivazione:" in line.lower():
                parsing_motivazione = True
                parts = line.lower().split("motivazione:", 1)
                if len(parts) > 1 and parts[1].strip():
                    motivazione_list.append(parts[1].strip())
                continue
            elif parsing_motivazione and line:  # Ignora linee vuote
                motivazione_list.append(line)
        
        motivazione = " ".join(motivazione_list).strip()
        
        if not motivazione and "motivazione:" in testo.lower():
            parts = testo.lower().split("motivazione:", 1)
            if len(parts) > 1:
                motivazione = parts[1].strip()
        
        if not motivazione and numero:
            parts = testo.lower().split("risposta:", 1)
            if len(parts) > 1:
                text_after_risposta = parts[1]
                if numero in text_after_risposta:
                    text_after_risposta = text_after_risposta.replace(numero, "", 1)
                motivazione = text_after_risposta.strip()
        
        return numero, motivazione

    except Exception as e:
        print(f"[ERRORE] Chiamata a Ollama fallita:\n{e}")
        return "", ""


df = pd.read_excel(INPUT_FILE)

df["Risposta numerica"] = df["Risposta numerica"].astype('object') if "Risposta numerica" in df.columns else pd.Series(dtype='object')
df["Motivazione"] = df["Motivazione"].astype('object') if "Motivazione" in df.columns else pd.Series(dtype='object')

if os.path.exists(OUTPUT_FILE):
    print(f"File di output {OUTPUT_FILE} trovato. Verifica se ci sono risposte già elaborate...")
    try:
        risposte_esistenti = pd.read_excel(OUTPUT_FILE)
        risposte_df = risposte_esistenti.copy()
        print(f"Caricate {len(risposte_df)} risposte già elaborate.")
        
        for idx, row in risposte_df.iterrows():
            mask = df["Prompt"] == row["Prompt"]
            if mask.any():
                i = mask.idxmax()
                df.loc[i, "Risposta numerica"] = row["Risposta numerica"]
                df.loc[i, "Motivazione"] = row["Motivazione"]
        
    except Exception as e:
        print(f"Errore nel caricamento del file esistente: {e}")
        risposte_df = pd.DataFrame(columns=df.columns)
        print("Creato DataFrame vuoto per le risposte")
else:
    risposte_df = pd.DataFrame(columns=df.columns)
    print("Creato DataFrame vuoto per le risposte")

try:
    for i, row in df.iterrows():
        if pd.notna(row["Risposta numerica"]) and pd.notna(row["Motivazione"]):
            continue

        prompt = str(row["Prompt"])
        if not prompt.strip():
            print(f"[{i}] Prompt vuoto, salto.")
            continue

        print(f"[{i}] Invio prompt a Ollama...")

        numero, motivazione = interroga_ollama(prompt)
        update_df = pd.DataFrame({
            "Risposta numerica": [numero],
            "Motivazione": [motivazione]
        }, index=[i])

        for col in update_df.columns:
            df.loc[i, col] = update_df.loc[i, col]
            
        nuova_riga = pd.Series(index=df.columns)
        
        for col in df.columns:
            nuova_riga[col] = df.loc[i, col]
        
        nuova_riga["Risposta numerica"] = numero
        nuova_riga["Motivazione"] = motivazione
        
        risposte_df.loc[len(risposte_df)] = nuova_riga
        print(f"[LOG] Aggiunta riga al DataFrame delle risposte (totale: {len(risposte_df)})")

        if i % AUTOSAVE_INTERVAL == 0 and i > 0:
            if len(risposte_df) > 0:
                risposte_df.to_excel(OUTPUT_FILE, index=False)
                print(f"[AUTO-SAVE] File salvato a riga {i} con {len(risposte_df)} risposte elaborate")
            else:
                print("[WARNING] Dataframe delle risposte vuoto, nessun salvataggio effettuato")

        time.sleep(1)

except KeyboardInterrupt:
    print("\n\n[INTERRUZIONE] Script interrotto dall'utente. Salvataggio dati...")
    if len(risposte_df) > 0:
        risposte_df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ File salvato come: {OUTPUT_FILE} con {len(risposte_df)} risposte")
    else:
        print("❌ Nessuna risposta da salvare")
    exit(0)
except Exception as e:
    print(f"\n\n[ERRORE] Si è verificato un errore: {e}")
    print("Tentativo di salvataggio dati...")
    try:
        if len(risposte_df) > 0:
            risposte_df.to_excel(OUTPUT_FILE, index=False)
            print(f"✅ File salvato come: {OUTPUT_FILE} con {len(risposte_df)} risposte")
        else:
            print("❌ Nessuna risposta da salvare")
    except Exception as save_error:
        print(f"❌ Impossibile salvare il file: {save_error}")
    exit(1)

# === SALVATAGGIO FINALE ===
if len(risposte_df) > 0:
    risposte_df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ File finale salvato come: {OUTPUT_FILE} con {len(risposte_df)} risposte")
else:
    print("❌ Nessuna risposta da salvare")
