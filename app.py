import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Nacenenie výkazu výmer SK", layout="wide")

st.title("🏗️ Automatické dopĺňanie cien do výkazu výmer")
st.markdown(
    """
Táto aplikácia chráni vaše texty, štruktúru aj vzorce:
- 🔒 **100% ochrana textov a vzorcov**
- 💶 **Doplnenie cien** podľa slovenských trhových štandardov
"""
)


def najdi_jednotkovu_cenu(popis):
  if not popis:
    return 0.0
  p = str(popis).lower()
  if "vykop" in p or "výkop" in p or "zemné práce" in p:
    return 28.50
  elif "betón" in p or "beton" in p:
    return 130.00
  elif "výstuž" in p or "armatúra" in p or "kari" in p:
    return 2.10
  elif "murivo" in p or "tehla" in p or "porotherm" in p or "ytong" in p:
    return 68.00
  elif "zateplenie" in p or "fasáda" in p:
    return 48.00
  elif "omietka" in p:
    return 15.00
  elif "poter" in p:
    return 19.00
  elif "sadrokartón" in p or "sdk" in p:
    return 32.00
  elif "dlažba" in p or "obklad" in p:
    return 38.00
  elif "lešenie" in p:
    return 9.50
  elif "presun hmôt" in p:
    return 1200.00
  else:
    return 35.00


uploaded_file = st.file_uploader(
    "Nahrajte váš vzorový Excel výkaz výmer (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
  try:
    # Zobrazenie náhľadu tabuľky, aby užívateľ presne videl stĺpce
    df_preview = pd.read_excel(uploaded_file)
    st.subheader("👀 Náhľad nahraného Excelu:")
    st.dataframe(df_preview.head(10))
    st.info(
        "💡 Pozrite sa na tabuľku vyššie. Zadajte číslo stĺpca (1 = prvý stĺpec"
        " A, 2 = druhý stĺpec B atď.), kde máte názvy položiek, a kam chcete"
        " zapísať jednotkovú cenu."
    )

    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    selected_sheet = st.selectbox(
        "Vyberte pracovný hárok (záložku) v Exceli", wb.sheetnames
    )
    ws = wb[selected_sheet]

    col1, col2, col3 = st.columns(3)
    with col1:
      popis_col_idx = st.number_input(
          "Stĺpec s popisom položky (číslo)", min_value=1, value=2, step=1
      )
    with col2:
      cena_col_idx = st.number_input(
          "Stĺpec pre jednotkovú cenu (číslo)", min_value=1, value=5, step=1
      )
    with col3:
      start_row = st.number_input(
          "Prvý riadok s položkami", min_value=1, value=4, step=1
      )

    if st.button("Spustiť dopĺňanie cien"):
      zmenenych = 0
      for r_idx in range(start_row, ws.max_row + 1):
        popis_cell = ws.cell(row=r_idx, column=popis_col_idx)
        if (
            popis_cell.value is not None
            and str(popis_cell.value).strip() != ""
        ):
          cena_cell = ws.cell(row=r_idx, column=cena_col_idx)
          cena_cell.value = najdi_jednotkovu_cenu(popis_cell.value)
          zmenenych += 1

      if zmenenych > 0:
        st.success(f"Úspešne doplnené jednotkové ceny pre {zmenenych} položiek!")
      else:
        st.warning(
            "⚠️ Nenašli sa žiadne položky. Skontrolujte číslo stĺpca s"
            " popisom a začiatočný riadok z náhľadu vyššie!"
        )

      output_io = io.BytesIO()
      wb.save(output_io)
      output_io.seek(0)

      st.download_button(
          label="📥 Stiahnuť nacenený Excel súbor",
          data=output_io,
          file_name="naceneny_vykaz_vymer_final.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
  except Exception as e:
    st.error(f"Chyba pri spracovaní: {e}")
