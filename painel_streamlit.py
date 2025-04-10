import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from io import BytesIO

# Configuração da página Streamlit
st.set_page_config(page_title="Painel de Pedidos Universal", layout="wide")
st.title("📊 Painel Interativo de Pedidos")

# Upload da planilha
uploaded_file = st.file_uploader("Faça upload da planilha anual de pedidos:", type=["xlsx"])

if uploaded_file:
    df = pd.ExcelFile(uploaded_file)
    all_sheets = df.sheet_names

    meses = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
             'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']

    consolidated_data = []

    for sheet in all_sheets:
        temp_df = df.parse(sheet)
        if not temp_df.empty and temp_df.shape[1] > 1:
            meses_df = temp_df[temp_df.iloc[:, 0].isin(meses)]

            if not meses_df.empty:
                regioes = temp_df.columns[1:]

                melted_df = meses_df.melt(id_vars=meses_df.columns[0], value_vars=regioes,
                                          var_name='Região', value_name='Valor')

                melted_df['Pessoa'] = sheet
                melted_df.rename(columns={melted_df.columns[0]: 'Mês'}, inplace=True)

                consolidated_data.append(melted_df)

    final_df = pd.concat(consolidated_data, ignore_index=True)

    import re

    def clean_value(value):
        if isinstance(value, str):
            value = re.sub(r"[^\d,\.]", "", value)
            if value.count('.') > 1:
                parts = value.split('.')
                value = ''.join(parts[:-1]) + '.' + parts[-1]
            if ',' in value:
                value = value.replace('.', '').replace(',', '.')
            try:
                return float(value)
            except:
                return 0.0
        try:
            return float(value)
        except:
            return 0.0

    final_df['Valor'] = final_df['Valor'].apply(clean_value)

    # Filtros interativos
    st.sidebar.header("Filtros")

    pessoas = st.sidebar.multiselect("Pessoa", options=final_df['Pessoa'].unique(), default=final_df['Pessoa'].unique())
    regioes = st.sidebar.multiselect("Região", options=final_df['Região'].unique(), default=final_df['Região'].unique())
    meses_selecionados = st.sidebar.multiselect("Mês", options=final_df['Mês'].unique(), default=final_df['Mês'].unique())

    filtered_df = final_df[
        (final_df['Pessoa'].isin(pessoas)) &
        (final_df['Região'].isin(regioes)) &
        (final_df['Mês'].isin(meses_selecionados))
    ]

    # Exibir tabela de dados filtrados
    st.subheader("Base Consolidada Filtrada")
    st.dataframe(filtered_df)

    # Gráficos
    st.subheader("Total por Pessoa")
    total_pessoa = filtered_df.groupby('Pessoa')['Valor'].sum().sort_values(ascending=False)
    fig1, ax1 = plt.subplots()
    total_pessoa.plot(kind='bar', ax=ax1, color='skyblue')
    ax1.set_ylabel('Valor Total (R$)')
    ax1.set_xlabel('Pessoa')
    st.pyplot(fig1)

    st.subheader("Total por Região")
    total_regiao = filtered_df.groupby('Região')['Valor'].sum().sort_values(ascending=False)
    fig2, ax2 = plt.subplots()
    total_regiao.plot(kind='bar', ax=ax2, color='orange')
    ax2.set_ylabel('Valor Total (R$)')
    ax2.set_xlabel('Região')
    st.pyplot(fig2)

    st.subheader("Evolução Mensal Geral")
    evolucao_mensal = filtered_df.groupby('Mês')['Valor'].sum().reindex(meses)
    fig3, ax3 = plt.subplots()
    evolucao_mensal.plot(kind='line', marker='o', ax=ax3, color='green')
    ax3.set_ylabel('Valor Total (R$)')
    ax3.set_xlabel('Mês')
    ax3.grid(True)
    st.pyplot(fig3)

    # Exportação para CSV
    st.sidebar.markdown("### Exportar Dados")
    csv = filtered_df.to_csv(index=False).encode('utf-8')
    st.sidebar.download_button("Baixar CSV Consolidado", data=csv, file_name='consolidado_pedidos.csv', mime='text/csv')

    # Exportação dos gráficos para PNG
    def export_plot(fig):
        buf = BytesIO()
        fig.savefig(buf, format="png")
        buf.seek(0)
        return buf

    st.sidebar.markdown("### Exportar Gráficos")
    st.sidebar.download_button("Baixar Gráfico Pessoa", data=export_plot(fig1), file_name="grafico_pessoa.png", mime="image/png")
    st.sidebar.download_button("Baixar Gráfico Região", data=export_plot(fig2), file_name="grafico_regiao.png", mime="image/png")
    st.sidebar.download_button("Baixar Gráfico Mensal", data=export_plot(fig3), file_name="grafico_mensal.png", mime="image/png")

    st.success("Painel atualizado com sucesso! 🚀")

else:
    st.warning("Por favor, faça o upload da planilha anual de pedidos para começar.")
