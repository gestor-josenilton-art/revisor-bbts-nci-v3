import streamlit as st
import pandas as pd
import io
import openpyxl

st.set_page_config(page_title="Agente Revisor de Planilhas BBTS x NCI")
st.title("📊 Agente Revisor de Planilhas BBTS x NCI")
st.markdown("Compare planilhas fiscais e identifique divergências de forma automatizada.")

uploaded_bbts = st.file_uploader("Envie a planilha BBTS (.xlsx ou .csv)", type=["xlsx", "csv"], key="bbts")
uploaded_nci = st.file_uploader("Envie a planilha NCI (.xlsx ou .csv)", type=["xlsx", "csv"], key="nci")

def load_file(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file, encoding='utf-8', sep=None, engine='python')
    else:
        return pd.read_excel(file)

if uploaded_bbts and uploaded_nci:
    df_bbts = load_file(uploaded_bbts)
    df_nci = load_file(uploaded_nci)

    for df in [df_bbts, df_nci]:
        df.columns = df.columns.str.strip()
        df['Número NF-e'] = df['Número NF-e'].astype(str).str.strip()
        df['CFOP'] = df['CFOP'].astype(str).str.strip()

    merged = pd.merge(df_bbts, df_nci, on='Número NF-e', how='outer', suffixes=('_BBTS', '_NCI'), indicator=True)

    exclusivas_bbts = merged[merged['_merge'] == 'left_only']
    exclusivas_nci = merged[merged['_merge'] == 'right_only']
    cfop_diferente = merged[(merged['_merge'] == 'both') & (merged['CFOP_BBTS'] != merged['CFOP_NCI'])]

    st.subheader("🔍 Notas Exclusivas - BBTS")
    st.dataframe(exclusivas_bbts[['Número NF-e', 'CFOP_BBTS']])

    st.subheader("🔍 Notas Exclusivas - NCI")
    st.dataframe(exclusivas_nci[['Número NF-e', 'CFOP_NCI']])

    st.subheader("⚠️ CFOPs Divergentes")
    st.dataframe(cfop_diferente[['Número NF-e', 'CFOP_BBTS', 'CFOP_NCI']])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        exclusivas_bbts.to_excel(writer, index=False, sheet_name='Exclusivas BBTS')
        exclusivas_nci.to_excel(writer, index=False, sheet_name='Exclusivas NCI')
        cfop_diferente.to_excel(writer, index=False, sheet_name='CFOP Divergente')

    st.download_button(
        label="📥 Baixar relatório Excel",
        data=output.getvalue(),
        file_name="relatorio_divergencias.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("🧠 Resumo Automático")
    resumo = f"Foram identificadas {len(exclusivas_bbts)} NF-es exclusivas da BBTS, {len(exclusivas_nci)} da NCI, e {len(cfop_diferente)} divergências de CFOP."
    st.markdown(resumo)

    refinar = st.radio("Deseja aplicar algum filtro adicional?", ["Não", "Sim"])
    if refinar == "Sim":
        filtro_cfop = st.text_input("Informe um CFOP para filtrar as divergências (ex: 2154)")
        if filtro_cfop:
            filtrado = cfop_diferente[(cfop_diferente['CFOP_BBTS'] == filtro_cfop) | (cfop_diferente['CFOP_NCI'] == filtro_cfop)]
            st.dataframe(filtrado[['Número NF-e', 'CFOP_BBTS', 'CFOP_NCI']])
else:
    st.info("Envie as duas planilhas para iniciar a análise.")