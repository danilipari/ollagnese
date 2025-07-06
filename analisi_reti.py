#!/usr/bin/env python3
# Analisi reti sociali - Matrici di adiacenza
# Genera matrici pesate e binarie da un file Excel di risposte SNA

import pandas as pd
import numpy as np

def crea_matrici(df, domande, nodes, is_binaria=False, soglia=2.5):
    """
    Crea le matrici di adiacenza pesata e binaria per un sottoinsieme di domande.
    - df: DataFrame con colonne ['from','to','question','score']
    - domande: lista di codici domanda da includere
    - nodes: lista di tutti i nodi (profili)
    - is_binaria: se True, restituisce la matrice binaria identica alla pesata
    - soglia: soglia per la binarizzazione (>= soglia → 1)
    """
    # Filtra risposte per le domande di rete
    df_rete = df[df["question"].isin(domande)]
    # Calcola media dei punteggi per coppia (from → to)
    media_pesi = (
        df_rete
        .groupby(["from", "to"])["score"]
        .mean()
        .reset_index()
    )

    # Matrice pesata inizializzata a zero (float)
    matrice_pesata = pd.DataFrame(0.0, index=nodes, columns=nodes)
    for _, row in media_pesi.iterrows():
        matrice_pesata.at[row["from"], row["to"]] = row["score"]

    # Imposta la diagonale a NaN
    np.fill_diagonal(matrice_pesata.values, np.nan)

    # Matrice binaria: prima booleana, poi cast a float
    matrice_binaria = (matrice_pesata >= soglia) if not is_binaria else matrice_pesata.copy()
    # Cast a float per supportare NaN
    matrice_binaria = matrice_binaria.astype(float)
    # Imposta la diagonale a NaN anche nella binaria
    np.fill_diagonal(matrice_binaria.values, np.nan)

    return matrice_pesata, matrice_binaria

def main():
    # ─── 1. Impostazioni ───────────────────────────────────────────
    # input_file  = "questionari_con_risposte_matrici.xlsx"
    # output_file = "matrici_adiacenza_reti.xlsx"
    input_file  = "questionari_con_risposte_matriciFR.xlsx"
    output_file = "matrici_adiacenza_retiFR.xlsx"

    # ─── 2. Carica e pulisci i dati ───────────────────────────────
    df = pd.read_excel(input_file, sheet_name="Sheet1")

    df = df.rename(columns={
        "Cod profilo rispondente": "from",
        "Codice domanda":         "question",
        "Cod profilo valutato":   "to",
        "Risposta numerica":      "score"
    })
    df["from"] = df["from"].str.lower()
    df["to"]   = df["to"].str.lower()

    # ─── 3. Definizione delle reti e nodi ─────────────────────────
    domande_reti = {
        "CC": [f"CC_0{i}" for i in range(1, 5)],
        "EC": [f"EC_0{i}" for i in range(1, 6)],
        "CM": [f"CM_0{i}" for i in range(1, 3)],
        "FR": ["FR_01"]
    }
    nodes = sorted(df["from"].unique())

    # ─── 4. Costruzione delle matrici ─────────────────────────────
    matrici = {}
    for rete, domande in domande_reti.items():
        is_binaria = (rete == "FR")
        pesata, binaria = crea_matrici(df, domande, nodes, is_binaria)
        matrici[rete] = {"pesata": pesata, "binaria": binaria}

    # ─── 5. Esportazione in Excel ──────────────────────────────────
    with pd.ExcelWriter(output_file) as writer:
        for rete, mats in matrici.items():
            mats["pesata"].to_excel(writer, sheet_name=f"{rete}_pesata")
            mats["binaria"].to_excel(writer, sheet_name=f"{rete}_binaria")

    print(f"Matrici esportate in '{output_file}'")

if __name__ == "__main__":
    main()