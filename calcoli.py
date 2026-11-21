"""
calcoli.py - Tutta la logica del drink cost, senza interfaccia.
La usano sia app.py (Streamlit) sia, in futuro, altri programmi.
"""
import math
import pandas as pd
from openpyxl import load_workbook

COL_MAG = ["Ingrediente", "Prezzo bottiglia (€)", "Formato (cl)"]
COL_CARTA = ["Drink", "Prezzo in carta (€)", "Extra ghiaccio/garnish (€)"]
COL_RIC = ["Drink", "Ingrediente", "cl"]


# ==============================
# LETTURA E SCRITTURA EXCEL
# ==============================
def carica_excel(percorso):
    """Legge i tre fogli e restituisce (magazzino, carta, ricette) come tabelle pandas."""
    magazzino = pd.read_excel(percorso, sheet_name="Magazzino")[COL_MAG]
    carta = pd.read_excel(percorso, sheet_name="Carta")[COL_CARTA]
    ricette = pd.read_excel(percorso, sheet_name="Ricette")[COL_RIC]
    return pulisci(magazzino, carta, ricette)


def salva_excel(percorso, magazzino, carta, ricette):
    """Riscrive i dati nell'Excel mantenendo menu a tendina, colori e formule."""
    wb = load_workbook(percorso)
    _scrivi_foglio(wb["Magazzino"], magazzino[COL_MAG], max_righe=200)
    _scrivi_foglio(wb["Carta"], carta[COL_CARTA], max_righe=200)
    _scrivi_foglio(wb["Ricette"], ricette[COL_RIC], max_righe=500)
    wb.save(percorso)


def _scrivi_foglio(ws, tabella, max_righe):
    if len(tabella) > max_righe - 1:
        raise ValueError(f"Troppe righe nel foglio {ws.title} (max {max_righe - 1})")
    # svuoto le vecchie righe (solo le colonne dei dati, le formule restano)
    for r in range(2, max_righe + 1):
        for c in range(1, len(tabella.columns) + 1):
            ws.cell(row=r, column=c).value = None
    # scrivo le nuove
    for r, riga in enumerate(tabella.itertuples(index=False), start=2):
        for c, valore in enumerate(riga, start=1):
            ws.cell(row=r, column=c).value = None if pd.isna(valore) else valore


# ==============================
# PULIZIA E CONTROLLI
# ==============================
def pulisci(magazzino, carta, ricette):
    """Toglie righe vuote e spazi nei nomi (errori tipici di chi scrive a mano)."""
    magazzino = magazzino.dropna(subset=["Ingrediente"]).copy()
    carta = carta.dropna(subset=["Drink"]).copy()
    ricette = ricette.dropna(subset=["Drink", "Ingrediente"]).copy()
    magazzino["Ingrediente"] = magazzino["Ingrediente"].astype(str).str.strip()
    carta["Drink"] = carta["Drink"].astype(str).str.strip()
    ricette["Drink"] = ricette["Drink"].astype(str).str.strip()
    ricette["Ingrediente"] = ricette["Ingrediente"].astype(str).str.strip()
    # tolgo le righe rimaste con il nome vuoto
    magazzino = magazzino[magazzino["Ingrediente"] != ""]
    carta = carta[carta["Drink"] != ""]
    ricette = ricette[(ricette["Drink"] != "") & (ricette["Ingrediente"] != "")]
    return (magazzino.reset_index(drop=True), carta.reset_index(drop=True),
            ricette.reset_index(drop=True))


def controlla_errori(magazzino, carta, ricette):
    """Restituisce una lista di frasi che spiegano cosa correggere."""
    errori = []
    ingredienti = set(magazzino["Ingrediente"])
    drink_carta = set(carta["Drink"])
    drink_ricette = set(ricette["Drink"])

    for nome, n in magazzino["Ingrediente"].value_counts().items():
        if n > 1:
            errori.append(f"'{nome}' compare {n} volte nel Magazzino")
    for nome, n in carta["Drink"].value_counts().items():
        if n > 1:
            errori.append(f"'{nome}' compare {n} volte in Carta")
    for _, r in magazzino.iterrows():
        if pd.isna(r["Prezzo bottiglia (€)"]) or pd.isna(r["Formato (cl)"]) or r["Formato (cl)"] <= 0:
            errori.append(f"'{r['Ingrediente']}': manca il prezzo o il formato in Magazzino")
    for _, r in carta.iterrows():
        if pd.isna(r["Prezzo in carta (€)"]) or r["Prezzo in carta (€)"] <= 0:
            errori.append(f"'{r['Drink']}': manca il prezzo in Carta")
    for _, r in ricette.iterrows():
        if pd.isna(r["cl"]) or r["cl"] <= 0:
            errori.append(f"Ricetta '{r['Drink']}': mancano i cl di {r['Ingrediente']}")
    for ing in sorted(set(ricette["Ingrediente"]) - ingredienti):
        errori.append(f"L'ingrediente '{ing}' è nelle ricette ma non nel Magazzino")
    for d in sorted(drink_carta - drink_ricette):
        errori.append(f"Il drink '{d}' è in Carta ma non ha una ricetta")
    for d in sorted(drink_ricette - drink_carta):
        errori.append(f"Il drink '{d}' ha una ricetta ma non è in Carta")
    return errori


# ==============================
# CALCOLI
# ==============================
def prezzo_consigliato(costo, iva, target):
    """Prezzo in carta (IVA inclusa) per stare al target, arrotondato ai 50 cent."""
    prezzo_lordo = costo / (target / 100) * (1 + iva)
    return math.ceil(prezzo_lordo * 2) / 2


def calcola(magazzino, carta, ricette, iva=0.22, sfrido=0.05, target=25):
    """
    Calcola tutto. Restituisce due tabelle:
      - riepilogo: una riga per drink
      - dettaglio: una riga per ogni ingrediente di ogni drink
    Da chiamare solo se controlla_errori() non ha trovato problemi.
    """
    prezzo_al_cl = dict(zip(magazzino["Ingrediente"],
                            magazzino["Prezzo bottiglia (€)"] / magazzino["Formato (cl)"]))

    dettaglio = ricette.copy()
    dettaglio["Costo (€)"] = [prezzo_al_cl[i] * cl * (1 + sfrido)
                              for i, cl in zip(ricette["Ingrediente"], ricette["cl"])]

    righe = []
    for _, d in carta.iterrows():
        extra = 0 if pd.isna(d["Extra ghiaccio/garnish (€)"]) else d["Extra ghiaccio/garnish (€)"]
        costo = dettaglio.loc[dettaglio["Drink"] == d["Drink"], "Costo (€)"].sum() + extra
        prezzo_netto = d["Prezzo in carta (€)"] / (1 + iva)
        drink_cost = costo / prezzo_netto * 100
        righe.append({
            "Drink": d["Drink"],
            "Costo (€)": round(costo, 2),
            "Prezzo carta (€)": d["Prezzo in carta (€)"],
            "Prezzo netto IVA (€)": round(prezzo_netto, 2),
            "Drink cost (%)": round(drink_cost, 1),
            "Margine (€)": round(prezzo_netto - costo, 2),
            "Prezzo consigliato (€)": prezzo_consigliato(costo, iva, target),
            "OK": drink_cost <= target,
        })
    riepilogo = pd.DataFrame(righe).sort_values("Margine (€)", ascending=False)
    return riepilogo.reset_index(drop=True), dettaglio
