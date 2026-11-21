"""
app.py - Interfaccia grafica del drink cost.
Avvio (dal terminale di VS Code, nella cartella del progetto):
    streamlit run app.py
"""
import pandas as pd
import streamlit as st

import calcoli

FILE_EXCEL = "bar.xlsx"

st.set_page_config(page_title="Drink Cost", page_icon="🍸", layout="wide")
st.title("🍸 Drink Cost del locale")


# ==============================
# CARICAMENTO (una volta sola, poi i dati restano in memoria)
# ==============================
def carica_in_memoria():
    try:
        mag, carta, ric = calcoli.carica_excel(FILE_EXCEL)
    except FileNotFoundError:
        st.error(f"Non trovo '{FILE_EXCEL}'. Mettilo nella stessa cartella di app.py.")
        st.stop()
    st.session_state["mag"] = mag
    st.session_state["carta"] = carta
    st.session_state["ric"] = ric
    # cambio la "versione" così le tabelle modificabili ripartono da zero
    st.session_state["versione"] = st.session_state.get("versione", 0) + 1


if "mag" not in st.session_state:
    carica_in_memoria()
v = st.session_state["versione"]


# ==============================
# BARRA LATERALE: impostazioni e salvataggio
# ==============================
with st.sidebar:
    st.header("⚙️ Impostazioni")
    iva = st.number_input("IVA (%)", 0.0, 30.0, 22.0, step=1.0) / 100
    sfrido = st.slider("Sfrido (%)", 0, 20, 5) / 100
    target = st.slider("Drink cost massimo (%)", 10, 50, 25)
    st.caption("Il drink cost è calcolato sul prezzo **senza IVA**.")
    st.divider()
    salva = st.button("💾 Salva nell'Excel", type="primary")
    ricarica = st.button("↩️ Annulla modifiche (ricarica Excel)")

if ricarica:
    carica_in_memoria()
    st.rerun()


tab_analisi, tab_mag, tab_ricette = st.tabs(["📊 Analisi", "📦 Magazzino", "🍸 Carta e ricette"])

# ==============================
# TAB MAGAZZINO
# ==============================
with tab_mag:
    st.subheader("Magazzino")
    st.caption("Clicca su una cella per modificarla. Per aggiungere una bottiglia usa l'ultima riga vuota; "
               "per eliminarla selezionala a sinistra e premi Canc.")
    mag = st.data_editor(
        st.session_state["mag"], key=f"ed_mag_{v}", num_rows="dynamic", hide_index=True,
        column_config={
            "Ingrediente": st.column_config.TextColumn(required=True),
            "Prezzo bottiglia (€)": st.column_config.NumberColumn(min_value=0.0, format="%.2f €"),
            "Formato (cl)": st.column_config.NumberColumn(min_value=1, step=1),
        },
    )

# ==============================
# TAB CARTA E RICETTE
# ==============================
with tab_ricette:
    col_sx, col_dx = st.columns([2, 3])
    with col_sx:
        st.subheader("Carta")
        carta = st.data_editor(
            st.session_state["carta"], key=f"ed_carta_{v}", num_rows="dynamic", hide_index=True,
            column_config={
                "Drink": st.column_config.TextColumn(required=True),
                "Prezzo in carta (€)": st.column_config.NumberColumn(min_value=0.0, format="%.2f €"),
                "Extra ghiaccio/garnish (€)": st.column_config.NumberColumn(min_value=0.0, format="%.2f €"),
            },
        )
    with col_dx:
        st.subheader("Ricette")
        st.caption("Drink e ingrediente si scelgono dal menu: prima aggiungi il drink in Carta "
                   "e la bottiglia in Magazzino.")
        nomi_drink = sorted(carta["Drink"].dropna().astype(str).str.strip().unique())
        nomi_ing = sorted(mag["Ingrediente"].dropna().astype(str).str.strip().unique())
        ric = st.data_editor(
            st.session_state["ric"], key=f"ed_ric_{v}", num_rows="dynamic", hide_index=True,
            column_config={
                "Drink": st.column_config.SelectboxColumn(options=nomi_drink, required=True),
                "Ingrediente": st.column_config.SelectboxColumn(options=nomi_ing, required=True),
                "cl": st.column_config.NumberColumn(min_value=0.0, step=0.5, format="%.1f"),
            },
        )

# pulisco i dati modificati (righe vuote, spazi) e controllo gli errori
mag, carta, ric = calcoli.pulisci(mag, carta, ric)
errori = calcoli.controlla_errori(mag, carta, ric)

if salva:
    try:
        calcoli.salva_excel(FILE_EXCEL, mag, carta, ric)
        st.sidebar.success("Salvato! ✅")
    except PermissionError:
        st.sidebar.error("Non riesco a salvare: chiudi bar.xlsx in Excel e riprova.")

# ==============================
# TAB ANALISI
# ==============================
with tab_analisi:
    if errori:
        st.error("Correggi questi problemi per vedere l'analisi:")
        for e in errori:
            st.write("❌", e)
        st.stop()
    if carta.empty:
        st.info("Aggiungi almeno un drink in Carta.")
        st.stop()

    riepilogo, dettaglio = calcoli.calcola(mag, carta, ric, iva, sfrido, target)

    # numeri principali in alto
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Drink in carta", len(riepilogo))
    c2.metric("Drink cost medio", f"{riepilogo['Drink cost (%)'].mean():.1f} %")
    c3.metric("Margine medio", f"{riepilogo['Margine (€)'].mean():.2f} €")
    c4.metric("Sopra il target", int((~riepilogo["OK"]).sum()))

    # tabella colorata: rosso se sopra il target
    def colora(riga):
        colore = "" if riga["OK"] else "background-color: rgba(220, 50, 50, 0.25)"
        return [colore] * len(riga)

    st.subheader("Riepilogo (dal più redditizio)")
    st.dataframe(
        riepilogo.style.apply(colora, axis=1).format({
            "Costo (€)": "{:.2f} €", "Prezzo carta (€)": "{:.2f} €",
            "Prezzo netto IVA (€)": "{:.2f} €", "Drink cost (%)": "{:.1f} %",
            "Margine (€)": "{:.2f} €", "Prezzo consigliato (€)": "{:.2f} €",
        }),
        hide_index=True, column_config={"OK": st.column_config.CheckboxColumn("Sotto target")},
    )

    st.subheader("Margine per drink (€)")
    st.bar_chart(riepilogo, x="Drink", y="Margine (€)", horizontal=True)

    # dettaglio + simulatore di prezzo
    st.divider()
    st.subheader("🔍 Dettaglio e simulatore")
    scelto = st.selectbox("Scegli un drink", riepilogo["Drink"])
    riga = riepilogo[riepilogo["Drink"] == scelto].iloc[0]

    sx, dx = st.columns(2)
    with sx:
        st.dataframe(dettaglio[dettaglio["Drink"] == scelto][["Ingrediente", "cl", "Costo (€)"]]
                     .style.format({"cl": "{:g}", "Costo (€)": "{:.2f} €"}), hide_index=True)
        st.write(f"Costo totale (con extra e sfrido): **{riga['Costo (€)']:.2f} €**")
    with dx:
        nuovo_prezzo = st.slider("Se lo vendessi a... (€, IVA inclusa)", 3.0, 25.0,
                                 float(riga["Prezzo carta (€)"]), step=0.5)
        costo = riga["Costo (€)"]
        netto_attuale = riga["Prezzo carta (€)"] / (1 + iva)
        netto = nuovo_prezzo / (1 + iva)
        dc_attuale, dc = costo / netto_attuale * 100, costo / netto * 100
        st.metric("Drink cost", f"{dc:.1f} %",
                  delta=f"{dc - dc_attuale:+.1f} punti", delta_color="inverse")
        st.metric("Margine per drink", f"{netto - costo:.2f} €",
                  delta=f"{netto - netto_attuale:+.2f} €")
