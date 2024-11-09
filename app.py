import streamlit as st
import pandas as pd
import datetime
from calculos import processar_dados, salvar_em_excel
from graficos import grafico_ganhos_por_aluno, grafico_distribuicao_ganhos
import os

# Configuração da página
st.set_page_config(page_title="Relatório de Aulas", layout="wide", initial_sidebar_state="expanded")

# Estilização adicional para botões
st.markdown(
    """
    <style>
    .stButton>button {
        background-color: #1F77B4;
        color: #FFFFFF;
        border-radius: 8px;
        height: 40px;
        width: 100%;
        font-size: 16px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Função para limpar arquivos .xlsx antigos
def limpar_arquivos_xlsx():
    try:
        arquivos_relatorio = [f for f in os.listdir() if f.endswith(".xlsx")]
        for arquivo in arquivos_relatorio:
            os.remove(arquivo)
        st.sidebar.info("Arquivos de relatório antigos removidos com sucesso.")
    except Exception as e:
        st.sidebar.error(f"Erro ao remover arquivos de relatório antigos: {e}")

# Exibir título e descrição
st.markdown("<h1 style='text-align: center;'>Relatório de Aulas</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Este relatório apresenta uma análise detalhada das aulas ministradas.</p>", unsafe_allow_html=True)

# Limpar arquivos ao iniciar
limpar_arquivos_xlsx()

# Exibir parâmetros de filtragem
st.sidebar.header("Parâmetros de Filtragem")

hoje = datetime.date.today()
primeiro_dia_mes_anterior = (hoje.replace(day=1) - datetime.timedelta(days=1)).replace(day=1)
ultimo_dia_mes_anterior = hoje.replace(day=1) - datetime.timedelta(days=1)

# Inputs de data
data_inicio = st.sidebar.date_input("Data de Início", value=primeiro_dia_mes_anterior)
data_fim = st.sidebar.date_input("Data de Fim", value=ultimo_dia_mes_anterior)

# Validar as datas
if data_inicio > data_fim:
    st.error("A data de início não pode ser posterior à data de fim.")
    st.stop()

# Chamada da função processar_dados com os argumentos corretos
if st.sidebar.button("📊 Gerar Relatório"):
    with st.spinner("Processando dados..."):
        try:
            saida, resumo_ano, resumo_aluno = processar_dados(data_inicio, data_fim)
        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
            st.stop()

        if saida.empty:
            st.error("Os dados retornados estão vazios.")
            st.stop()

        st.session_state["saida"] = saida.reset_index(drop=True)
        st.session_state["resumo_aluno"] = resumo_aluno

if "saida" in st.session_state and "resumo_aluno" in st.session_state:
    st.dataframe(st.session_state["saida"], use_container_width=True)

    st.subheader("Gráficos de Análise")
    grafico_ganhos_por_aluno(st.session_state["resumo_aluno"])
    grafico_distribuicao_ganhos(st.session_state["resumo_aluno"])

    filename = f'relatorio_{pd.Timestamp.now().date()}.xlsx'
    salvar_em_excel(st.session_state["saida"], filename=filename)
    st.session_state["filename"] = filename if os.path.exists(filename) else None

    if st.session_state["filename"]:
        st.sidebar.subheader("Download do Relatório")
        with open(st.session_state["filename"], "rb") as file:
            st.sidebar.download_button(
                label="📥 Baixar Relatório em Excel",
                data=file,
                file_name=st.session_state["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
