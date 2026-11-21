import math
import pandas as pd

# ==============================
# IMPOSTAZIONI DEL LOCALE
# ==============================
FILE_EXCEL = "bar.xlsx"
IVA = 0.22               # i prezzi in carta sono IVA inclusa
SFRIDO = 0.05            # 5% di prodotto perso
DRINK_COST_TARGET = 25   # percentuale massima accettabile (calcolata sul prezzo senza IVA)


# ==============================
# LETTURA DEL FILE EXCEL
# ==============================
def carica_dati(percorso):
    """Legge i tre fogli dell'Excel e restituisce magazzino, carta e ricette."""
    magazzino = pd.read_excel(percorso, sheet_name="Magazzino")
    carta = pd.read_excel(percorso, sheet_name="Carta")
    ricette = pd.read_excel(percorso, sheet_name="Ricette")

    # Tolgo le righe vuote e gli spazi di troppo nei nomi (errore tipico di chi scrive a mano)
    magazzino = magazzino.dropna(subset=["Ingrediente"])
    carta = carta.dropna(subset=["Drink"])
    ricette = ricette.dropna(subset=["Drink", "Ingrediente"])
    for tabella, colonne in ((magazzino, ["Ingrediente"]), (carta, ["Drink"]),
                             (ricette, ["Drink", "Ingrediente"])):
        for col in colonne:
            tabella[col] = tabella[col].astype(str).str.strip()

    # Trasformo il magazzino in un dizionario: {"Gin": prezzo al cl, ...}
    prezzo_al_cl = {}
    for _, riga in magazzino.iterrows():
        prezzo_al_cl[riga["Ingrediente"]] = riga["Prezzo bottiglia (€)"] / riga["Formato (cl)"]

    return prezzo_al_cl, carta, ricette


def controlla_errori(prezzo_al_cl, carta, ricette):
    """Segnala ingredienti o drink scritti male prima di fare i conti."""
    errori = []
    for ing in ricette["Ingrediente"].unique():
        if ing not in prezzo_al_cl:
            errori.append(f"L'ingrediente '{ing}' è nelle ricette ma non nel Magazzino")
    for drink in carta["Drink"]:
        if drink not in ricette["Drink"].values:
            errori.append(f"Il drink '{drink}' è in Carta ma non ha una ricetta")
    for drink in ricette["Drink"].unique():
        if drink not in carta["Drink"].values:
            errori.append(f"Il drink '{drink}' ha una ricetta ma non è in Carta")
    return errori


# ==============================
# CALCOLI
# ==============================
def prezzo_consigliato(costo):
    """Prezzo in carta (IVA inclusa) per stare al target, arrotondato ai 50 cent."""
    prezzo_lordo = costo / (DRINK_COST_TARGET / 100) * (1 + IVA)
    return math.ceil(prezzo_lordo * 2) / 2


def analizza_drink(drink, prezzo_al_cl, carta, ricette):
    info = carta[carta["Drink"] == drink].iloc[0]
    extra = info["Extra ghiaccio/garnish (€)"]
    if pd.isna(extra):
        extra = 0
    prezzo_carta = info["Prezzo in carta (€)"]

    print(f"\n--- {drink} ---")
    costo_totale = extra
    for _, riga in ricette[ricette["Drink"] == drink].iterrows():
        costo = prezzo_al_cl[riga["Ingrediente"]] * riga["cl"] * (1 + SFRIDO)
        costo_totale += costo
        print(f"  {riga['Ingrediente']:<16} {riga['cl']:>4g} cl   {costo:5.2f} €")

    prezzo_netto = prezzo_carta / (1 + IVA)
    drink_cost = costo_totale / prezzo_netto * 100
    margine = prezzo_netto - costo_totale

    print(f"  {'Extra':<16}           {extra:5.2f} €")
    print(f"  Costo totale:      {costo_totale:.2f} €")
    print(f"  Prezzo netto IVA:  {prezzo_netto:.2f} €  (in carta {prezzo_carta:.2f} €)")
    print(f"  Drink cost:        {drink_cost:.1f} %")
    print(f"  Margine:           {margine:.2f} €")
    if drink_cost > DRINK_COST_TARGET:
        print(f"  ⚠️  Troppo alto! Prezzo consigliato: {prezzo_consigliato(costo_totale):.2f} €")
    else:
        print("  ✅ Drink cost OK")

    return {"Drink": drink, "Costo (€)": round(costo_totale, 2),
            "Prezzo carta (€)": prezzo_carta, "Drink cost (%)": round(drink_cost, 1),
            "Margine (€)": round(margine, 2), "Prezzo consigliato (€)": prezzo_consigliato(costo_totale)}


# ==============================
# PROGRAMMA PRINCIPALE
# ==============================
def main():
    try:
        prezzo_al_cl, carta, ricette = carica_dati(FILE_EXCEL)
    except FileNotFoundError:
        print(f"Non trovo '{FILE_EXCEL}': mettilo nella stessa cartella di questo programma.")
        return

    errori = controlla_errori(prezzo_al_cl, carta, ricette)
    if errori:
        print("Correggi questi problemi nell'Excel e riprova:")
        for e in errori:
            print("  ❌", e)
        return

    risultati = [analizza_drink(d, prezzo_al_cl, carta, ricette)
                 for d in carta["Drink"]]

    report = pd.DataFrame(risultati).sort_values("Margine (€)", ascending=False)
    print("\n========== RIEPILOGO (dal più redditizio) ==========")
    print(report.to_string(index=False))

    report.to_excel("report_drink_cost.xlsx", index=False)
    print("\n📄 Report salvato in 'report_drink_cost.xlsx'")


main()
