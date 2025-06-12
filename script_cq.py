import pandas as pd
import requests
import time
import warnings

# Ignora i warning di pandas sulla compatibilità dei tipi
warnings.filterwarnings('ignore', category=FutureWarning)


INPUT_FILE = "database_questionari_inglese.xlsx"
OUTPUT_FILE = "questionari_con_risposte.xlsx"
OLLAMA_MODEL_3 = "llama3:latest" # Modello 3 a 70B # https://ollama.com/library/llama3 (4.7gb)
OLLAMA_MODEL_31 = "llama3.1:latest" # Modello 3.1 a 8B # https://ollama.com/library/llama3.1 (4.9gb)
OLLAMA_URL = "http://localhost:11434/api/generate"
AUTOSAVE_INTERVAL = 10  

def interroga_ollama(prompt):
    payload = {
        "model": OLLAMA_MODEL_31,
        "prompt": prompt.strip(),
        "stream": False
    }

    try:
        response = requests.post(OLLAMA_URL, json=payload)
        response.raise_for_status()
        testo = response.json().get("response", "")

        numero = ""
        motivazione = []
        parsing_motivazione = False

        for line in testo.splitlines():
            line = line.strip()
            if line.lower().startswith("risposta:"):
                numero = line.split(":", 1)[-1].strip()
            elif line.lower().startswith("motivazione:"):
                parsing_motivazione = True
                continue
            elif parsing_motivazione and line:
                motivazione.append(line)

        return numero, " ".join(motivazione).strip()

    except Exception as e:
        print(f"[ERRORE] Chiamata a Ollama fallita:\n{e}")
        return "", ""


# Leggi il file Excel
df = pd.read_excel(INPUT_FILE)

# Verifica i tipi di colonna attuali
print("Tipi di colonne originali:", df.dtypes)

# Converti o crea le colonne con il tipo corretto
# Indipendentemente da se esistono già o no, convertiamo al tipo object
df["Risposta numerica"] = df["Risposta numerica"].astype('object') if "Risposta numerica" in df.columns else pd.Series(dtype='object')
df["Motivazione"] = df["Motivazione"].astype('object') if "Motivazione" in df.columns else pd.Series(dtype='object')

# Verifica i tipi dopo la conversione
print("Tipi di colonne dopo conversione:", df.dtypes)

# CICLO PRINCIPALE
try:
    for i, row in df.iterrows():
        if pd.notna(row["Risposta numerica"]) and pd.notna(row["Motivazione"]):
            continue  # Riga già compilata

        prompt = str(row["Prompt"])
        if not prompt.strip():
            print(f"[{i}] Prompt vuoto, salto.")
            continue

        print(f"[{i}] Invio prompt a Ollama...")

        numero, motivazione = interroga_ollama(prompt)
        
        # Aggiorniamo il dataframe con un approccio diverso per evitare warning di tipo
        # Creiamo un nuovo DataFrame con una sola riga per poi aggiornare l'originale
        update_df = pd.DataFrame({
            "Risposta numerica": [numero],
            "Motivazione": [motivazione]
        }, index=[i])
        
        # Aggiorniamo il dataframe originale
        for col in update_df.columns:
            df.loc[i, col] = update_df.loc[i, col]

        # Autosalvataggio ogni N righe
        if i % AUTOSAVE_INTERVAL == 0 and i > 0:
            df.to_excel(OUTPUT_FILE, index=False)
            print(f"[AUTO-SAVE] File salvato a riga {i}")

        time.sleep(1)

except KeyboardInterrupt:
    print("\n\n[INTERRUZIONE] Script interrotto dall'utente. Salvataggio dati...")
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ File salvato come: {OUTPUT_FILE}")
    exit(0)
except Exception as e:
    print(f"\n\n[ERRORE] Si è verificato un errore: {e}")
    print("Tentativo di salvataggio dati...")
    try:
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ File salvato come: {OUTPUT_FILE}")
    except:
        print("❌ Impossibile salvare il file.")
    exit(1)

# === SALVATAGGIO FINALE ===
df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ File finale salvato come: {OUTPUT_FILE}")
