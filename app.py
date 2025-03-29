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
        self.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
        self.add_font("DejaVu", "B", "DejaVuSans-Bold.ttf", uni=True)
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
    temas = {
        "Dentistas ausentes": "não estava|não encontrado|ausente|fechada",
        "Agradecimentos": "agradeceu|agradecimento",
        "Reclamações": "reclamaç|problema",
        "Dúvidas": "dúvida",
        "Exames": "exame",
    }
    temas_resumos = {}
    for tema, pattern in temas.items():
        encontrados = resumos[resumos.str.contains(pattern, na=False)]
        temas_resumos[tema] = encontrados.sample(min(3, len(encontrados))).tolist()

    pdf = PDF()
    pdf.add_page()
    pdf.set_font("DejaVu", style="B", size=14)
    pdf.cell(0, 10, "Análise das Visitas - Relatório Inteligente", ln=True)

    pdf.set_font("DejaVu", size=11)
    pdf.multi_cell(0, 10, f"Data do Relatório: {datetime.today().strftime('%d/%m/%Y')}")

    pdf.set_font("DejaVu", style="B", size=12)
    pdf.cell(0, 10, "Resumo Quantitativo das Visitas", ln=True)
    pdf.set_font("DejaVu", size=9)
    pdf.multi_cell(0, 5, tabela_texto)

    pdf.set_font("DejaVu", style="B", size=12)
    pdf.cell(0, 10, "Resumo Geral", ln=True)
    pdf.set_font("DejaVu", size=11)
    pdf.multi_cell(0, 10, "Este relatório apresenta uma análise criteriosa das visitas realizadas durante o mês, agrupadas por temas.")

    for tema, exemplos in temas_resumos.items():
        pdf.set_font("DejaVu", style="B", size=12)
        pdf.cell(0, 10, tema, ln=True)
        pdf.set_font("DejaVu", size=11)
        for exemplo in exemplos:
            pdf.multi_cell(0, 8, f"• {exemplo.strip()}")
        if tema == "Dentistas ausentes":
            pdf.set_text_color(200, 0, 0)
            pdf.set_font("DejaVu", style="B", size=11)
            pdf.multi_cell(0, 10, "Sugestão: CONFIRMAR PRÉVIAMENTE as visitas para evitar frustrações.")
            pdf.set_text_color(0, 0, 0)

    buffer = BytesIO()
    pdf.output(buffer)
    buffer.seek(0)

    st.success("✅ Relatório gerado com sucesso!")
    st.download_button(label="📥 Baixar PDF do Relatório", data=buffer, file_name="relatorio_visitas.pdf", mime="application/pdf")
