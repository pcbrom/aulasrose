import plotly.express as px
import streamlit as st

def grafico_ganhos_por_aluno(resumo_aluno):
    fig = px.bar(
        resumo_aluno,
        x='Nome do aluno',
        y='Valor da aula',
        title="Ganhos por Aluno",
        labels={'Valor da aula': 'Ganhos (R$)', 'Nome do aluno': 'Aluno'},
        color='Valor da aula',
        color_continuous_scale='Blues'
    )
    fig.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig)

def grafico_distribuicao_ganhos(resumo_aluno):
    fig = px.pie(
        resumo_aluno,
        names='Nome do aluno',
        values='Valor da aula',
        title="Distribuição dos Ganhos por Aluno",
        hole=0.3
    )
    fig.update_traces(textinfo='percent+label')
    st.plotly_chart(fig)
