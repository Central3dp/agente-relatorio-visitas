import streamlit as st
import pandas as pd
from io import BytesIO
from fpdf import FPDF
from datetime import datetime
import os

st.set_page_config(page_title="Agente de Relatórios", layout="centered")
st.title("📊 Gerador Inteligente de Relatórios de Visitas")

class PDF(FPDF):
    def __init__(self):
        super().__init__()
        font_path = os.path.dirname(__file__)
        self.add_font("DejaVu", "", os.path.join(font_path, "DejaVuSans.ttf"), uni=True)
        self.add_font("DejaVu", "B", os.path.join(font_path, "DejaVuSans-Bold.ttf"), uni=True)
        self.set_font("DejaVu", size=12)

uploaded_file = st.file_uploader("📤 Envie a planilha Excel (.xlsx) com dados de visitas")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df = df.dropna(subset=["Data", "Nome", "Resumo da visita", "Tipo de visita"])
    df["Data"] = pd.to_datetime(df["Data"], errors='coerce')
    df = df.dropna(subset=["Data"])
    df["Semana"] = df["Data"].dt.isocalendar().week - pd.to_datetime("2025-03-01").isocalendar().week + 1

    tabela = pd.pivot_table(df, values="Nome", index="Semana", columns="Tipo de visita", aggfunc="count", fill_value=0)
    tabela["Total"] = tabela.sum(axis=1)

    def formatar_tabela(tabela):
        linhas = [
            "───────────────────────────────────────────────────────────────",
            "| Semana    | Efetiva | Frustrada | Pesquisa | Total de visitas |",
            "|-----------|---------|-----------|----------|------------------|",
        ]
        for semana, linha in tabela.iterrows():
            linhas.append(f"| Semana {semana:<2}  |   {linha.get('Efetiva', 0):<5}  |   {linha.get('Frustrada', 0):<7}  |   {linha.get('Pesquisa', 0):<6}  |        {int(linha['Total']):<8}     |")
        linhas.append("───────────────────────────────────────────────────────────────")
        return "\n".join(linhas)

    tabela_texto = formatar_tabela(tabela)
    resumos = df["Resumo da visita"].str.lower().dropna()
    te
