import pandas as pd
import requests
import time
import warnings
import os

warnings.filterwarnings('ignore', category=FutureWarning)


# INPUT_FILE = "database_questionari_inglese.xlsx"
# OUTPUT_FILE = "questionari_con_risposte.xlsx"
INPUT_FILE = "database_questionariFR.xlsx"
OUTPUT_FILE = "questionari_con_risposteFR.xlsx"
OLLAMA_MODEL = "gemma3:12b"
OLLAMA_URL = "http://localhost:11434/api/generate"
# AUTOSAVE_INTERVAL = 50
# MOTIVAZIONE_INTERVAL = 100
AUTOSAVE_INTERVAL = 25
MOTIVAZIONE_INTERVAL = 50

PROMPT_TEMPLATE = """
{prompt}

Please respond with only a number.
Answer: [insert a number]
"""

PROMPT_TEMPLATE_WITH_REASONING = """
{prompt}

Please respond in the following format:
Answer: [insert a number]
Reasoning: [insert your explanation]
"""

def interroga_ollama(prompt, richiedi_motivazione=False):
    template = PROMPT_TEMPLATE_WITH_REASONING if richiedi_motivazione else PROMPT_TEMPLATE
    prompt_formattato = template.format(prompt=prompt.strip())
    
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
        
        if richiedi_motivazione:
            parsing_motivazione = False
            motivazione_list = []
            
            for line in lines:
                line = line.strip()
                if line.lower().startswith("answer:"):
                    numero = line.split(":", 1)[-1].strip()
                elif "reasoning:" in line.lower():
                    parsing_motivazione = True
                    parts = line.lower().split("reasoning:", 1)
                    if len(parts) > 1 and parts[1].strip():
                        motivazione_list.append(parts[1].strip())
                    continue
                elif parsing_motivazione and line:
                    motivazione_list.append(line)
            
            motivazione = " ".join(motivazione_list).strip()
            
            # Fallback se non trova reasoning
            if not motivazione and "reasoning:" in testo.lower():
                parts = testo.lower().split("reasoning:", 1)
                if len(parts) > 1:
                    motivazione = parts[1].strip()
        else:
            # Parsing veloce solo per il numero
            for line in lines:
                line = line.strip()
                if line.lower().startswith("answer:"):
                    numero = line.split(":", 1)[-1].strip()
                    break
                elif line.isdigit():
                    numero = line
                    break
        
        # Estrazione numero con regex
        if not numero:
            import re
            numeri = re.findall(r'\d+', testo)
            if numeri:
                numero = numeri[0]
        
        return numero, motivazione

    except Exception as e:
        print(f"[ERROR] Ollama call failed:\n{e}")
        return "", ""


df = pd.read_excel(INPUT_FILE)

# Assicuriamoci che le colonne esistano
if "Risposta numerica" not in df.columns:
    df["Risposta numerica"] = ""
if "Motivazione" not in df.columns:
    df["Motivazione"] = ""

df["Risposta numerica"] = df["Risposta numerica"].astype('object')
df["Motivazione"] = df["Motivazione"].astype('object')

print(f"[DEBUG] DataFrame columns: {list(df.columns)}")
print(f"[DEBUG] DataFrame shape: {df.shape}")

if os.path.exists(OUTPUT_FILE):
    print(f"Output file {OUTPUT_FILE} found. Checking if there are already processed responses...")
    try:
        risposte_esistenti = pd.read_excel(OUTPUT_FILE)
        risposte_df = risposte_esistenti.copy()
        print(f"Loaded {len(risposte_df)} already processed responses.")
        
        for idx, row in risposte_df.iterrows():
            mask = df["Prompt"] == row["Prompt"]
            if mask.any():
                i = mask.idxmax()
                df.loc[i, "Risposta numerica"] = row["Risposta numerica"]
                df.loc[i, "Motivazione"] = row["Motivazione"]
        
    except Exception as e:
        print(f"Error loading existing file: {e}")
        risposte_df = pd.DataFrame(columns=df.columns)
        print("Created empty DataFrame for responses")
else:
    risposte_df = pd.DataFrame(columns=df.columns)
    print("Created empty DataFrame for responses")

try:
    for i, row in df.iterrows():
        # Debug: mostra lo stato della riga
        print(f"[DEBUG] Row {i}: Risposta numerica = {row.get('Risposta numerica', 'N/A')}")
        
        # Skip temporaneo - controlla solo se esiste già la risposta numerica
        if pd.notna(row["Risposta numerica"]) and str(row["Risposta numerica"]).strip() != "":
            print(f"[{i}] Already has response, skipping.")
            continue

        prompt = str(row["Prompt"])
        if not prompt.strip():
            print(f"[{i}] Empty prompt, skipping.")
            continue

        print(f"[{i}] Sending prompt to Ollama...")

        # Determina se richiedere la motivazione ogni MOTIVAZIONE_INTERVAL righe
        richiedi_motivazione = (i % MOTIVAZIONE_INTERVAL == 0)
        if richiedi_motivazione:
            print(f"[{i}] Richiesta motivazione per test a campione (ogni {MOTIVAZIONE_INTERVAL} righe)")
        
        numero, motivazione = interroga_ollama(prompt, richiedi_motivazione)
        print(f"[DEBUG] Got response: numero='{numero}', motivazione='{motivazione}'")
        
        if not numero:
            print(f"[WARNING] No number received for row {i}")
            numero = ""  # Assicuriamoci che sia una stringa
        
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
        print(f"[LOG] Added row to responses DataFrame (total: {len(risposte_df)})")

        # Autosave ogni AUTOSAVE_INTERVAL righe
        if i % AUTOSAVE_INTERVAL == 0 and i > 0:
            if len(risposte_df) > 0:
                risposte_df.to_excel(OUTPUT_FILE, index=False)
                print(f"[AUTO-SAVE] File saved at row {i} with {len(risposte_df)} processed responses")
            else:
                print("[WARNING] Responses DataFrame is empty, no save performed")

        time.sleep(1)

except KeyboardInterrupt:
    print("\n\n[INTERRUPTION] Script interrupted by user. Saving data...")
    if len(risposte_df) > 0:
        risposte_df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ File saved as: {OUTPUT_FILE} with {len(risposte_df)} responses")
    else:
        print("❌ No responses to save")
    exit(0)
except Exception as e:
    print(f"\n\n[ERROR] An error occurred: {e}")
    print("Attempting to save data...")
    try:
        if len(risposte_df) > 0:
            risposte_df.to_excel(OUTPUT_FILE, index=False)
            print(f"✅ File saved as: {OUTPUT_FILE} with {len(risposte_df)} responses")
        else:
            print("❌ No responses to save")
    except Exception as save_error:
        print(f"❌ Unable to save file: {save_error}")
    exit(1)

# === FINAL SAVE ===
if len(risposte_df) > 0:
    risposte_df.to_excel(OUTPUT_FILE, index=False)
    print(f"✅ Final file saved as: {OUTPUT_FILE} with {len(risposte_df)} responses")
else:
    print("❌ No responses to save")
