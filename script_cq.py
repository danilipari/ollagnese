import pandas as pd
import requests
import time
import warnings
import os

# Ignore pandas warnings about type compatibility
warnings.filterwarnings('ignore', category=FutureWarning)


INPUT_FILE = "database_questionari_inglese.xlsx"
OUTPUT_FILE = "questionari_con_risposte.xlsx"
# OLLAMA_MODEL = "llama3:latest" # "llama3.1:latest"
# OLLAMA_MODEL = "gemma3:12b"
OLLAMA_MODEL = "chevalblanc/gpt-4o-mini:latest"
OLLAMA_URL = "http://localhost:11434/api/generate"
AUTOSAVE_INTERVAL = 10
PROMPT_TEMPLATE = """
{prompt}

Please respond in the following format:
Answer: [insert a number]
Reasoning: [insert your explanation]
"""

def interroga_ollama(prompt):
    # Format the prompt to get structured responses
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
            if line.lower().startswith("answer:"):
                numero = line.split(":", 1)[-1].strip()
            elif "reasoning:" in line.lower():
                parsing_motivazione = True
                parts = line.lower().split("reasoning:", 1)
                if len(parts) > 1 and parts[1].strip():
                    motivazione_list.append(parts[1].strip())
                continue
            elif parsing_motivazione and line:  # Skip empty lines
                motivazione_list.append(line)
        
        motivazione = " ".join(motivazione_list).strip()
        
        if not motivazione and "reasoning:" in testo.lower():
            parts = testo.lower().split("reasoning:", 1)
            if len(parts) > 1:
                motivazione = parts[1].strip()
        
        if not motivazione and numero:
            parts = testo.lower().split("answer:", 1)
            if len(parts) > 1:
                text_after_answer = parts[1]
                if numero in text_after_answer:
                    text_after_answer = text_after_answer.replace(numero, "", 1)
                motivazione = text_after_answer.strip()
        
        return numero, motivazione

    except Exception as e:
        print(f"[ERROR] Ollama call failed:\n{e}")
        return "", ""


df = pd.read_excel(INPUT_FILE)

df["Risposta numerica"] = df["Risposta numerica"].astype('object') if "Risposta numerica" in df.columns else pd.Series(dtype='object')
df["Motivazione"] = df["Motivazione"].astype('object') if "Motivazione" in df.columns else pd.Series(dtype='object')

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
        if pd.notna(row["Risposta numerica"]) and pd.notna(row["Motivazione"]):
            continue

        prompt = str(row["Prompt"])
        if not prompt.strip():
            print(f"[{i}] Empty prompt, skipping.")
            continue

        print(f"[{i}] Sending prompt to Ollama...")

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
        print(f"[LOG] Added row to responses DataFrame (total: {len(risposte_df)})")

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
