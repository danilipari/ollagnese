import pandas as pd
import requests
import time


INPUT_FILE = "database_questionari.xlsx"
OUTPUT_FILE = "questionari_con_risposte.xlsx"
OLLAMA_MODEL = "llama3"
OLLAMA_URL = "http://localhost:11434/api/generate"
AUTOSAVE_INTERVAL = 10  

def interroga_ollama(prompt):
    payload = {
        "model": OLLAMA_MODEL,
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


df = pd.read_excel(INPUT_FILE)

# Crea le colonne se mancanti
if "Risposta numerica" not in df.columns:
    df["Risposta numerica"] = ""
if "Motivazione" not in df.columns:
    df["Motivazione"] = ""

# CICLO PRINCIPALE

for i, row in df.iterrows():
    if pd.notna(row["Risposta numerica"]) and pd.notna(row["Motivazione"]):
        continue  # Riga già compilata

    prompt = str(row["Prompt"])
    if not prompt.strip():
        print(f"[{i}] Prompt vuoto, salto.")
        continue

    print(f"[{i}] Invio prompt a Ollama...")

    numero, motivazione = interroga_ollama(prompt)

    df.at[i, "Risposta numerica"] = numero
    df.at[i, "Motivazione"] = motivazione

    # Autosalvataggio ogni N righe
    if i % AUTOSAVE_INTERVAL == 0 and i > 0:
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"[AUTO-SAVE] File salvato a riga {i}")

    time.sleep(1)

# === SALVATAGGIO FINALE ===
df.to_excel(OUTPUT_FILE, index=False)
print(f"✅ File finale salvato come: {OUTPUT_FILE}")
