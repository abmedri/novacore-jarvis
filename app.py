# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (VERSÃO EXECUTIVA UNIVERSAL & COMPLETA)
# ====================================================================================

import os
import io
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
import plotly.express as px
import plotly.graph_objects as go

# Dependências para Relatórios Executivos Avançados
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def processar_arquivo_upload(arquivo_carregado):
    nome = arquivo_carregado.name.lower()
    try:
        if nome.endswith('.csv'):
            for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
                try:
                    arquivo_carregado.seek(0)
                    df = pd.read_csv(arquivo_carregado, encoding=enc, sep=';', low_memory=False)
                    if df.shape[1] <= 1:
                        arquivo_carregado.seek(0)
                        df = pd.read_csv(arquivo_carregado, encoding=enc, sep=',', low_memory=False)
                    
                    for col in df.columns:
                        if col.lower() not in ['código do produto', 'codigo do produto', 'loja', 'inventario', 'descrição do produto', 'emb', 'local', 'data de abertura', 'data de finalização', 'tipo']:
                            try:
                                serie_str = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                                conv_num = pd.to_numeric(serie_str, errors='coerce')
                                if conv_num.notna().sum() > 0:
                                    df[col] = conv_num
                            except Exception:
                                pass
                    return df, "Inventário processado com sucesso."
                except Exception:
                    continue
            return None, "Erro: Não foi possível decodificar o ficheiro CSV."
        elif nome.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(arquivo_carregado)
            return df, "Excel processado com sucesso."
        elif nome.endswith('.json'):
            return pd.read_json(arquivo_carregado), "JSON processado com sucesso."
        elif nome.endswith('.xml'):
            tree = ET.parse(arquivo_carregado)
            root = tree.getroot()
            dados = [{elem.tag: elem.text for elem in filho} for filho in root]
            return pd.DataFrame(dados), "XML processado com sucesso."
        else:
            return None, "Formato não suportado."
    except Exception as e:
        return None, "Erro ao processar ficheiro: " + str(e)

def identificar_coluna_ideal(colunas_numericas, palavras_alvo):
    for palavra in palavras_alvo:
        for idx, col in enumerate(colunas_numericas):
            if palavra in col.lower() and 'codigo' not in col.lower() and 'código' not in col.lower():
                return idx
    return 0

def formatar_kpi_html(label, valor, delta_texto, eh_negativo=False, eh_positivo=False):
    cor_delta = "#9ca3af"
    sinal = ""
    if eh_negativo:
        cor_delta = "#ef4444"
        if not delta_texto.startswith("-"):
            sinal = "-"
    elif eh_positivo:
        cor_delta = "#10b981"
        if not delta_texto.startswith("+"):
            sinal = "+"

    return f"""
    <div style="background: rgba(17, 24, 39, 0.85); border: 1px solid rgba(31, 41, 55, 0.9); border-radius: 10px; padding: 14px 16px; margin-bottom: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);">
        <div style="color: #9ca3af; font-size: 0.82rem; font-weight: 500; margin-bottom: 4px; line-height: 1.2;">{label}</div>
        <div style="color: #f3f4f6; font-size: 1.3rem; font-weight: 700; font-family: monospace; letter-spacing: -0.5px; word-break: break-all;">{valor}</div>
        <div style="color: {cor_delta}; font-size: 0.8rem; font-weight: 600; margin-top: 4px;">{sinal}{delta_texto}</div>
    </div>
    """

# --- GERADORES DE RELATÓRIOS EXECUTIVOS ---
def gerar_excel_bytes(df):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Dashboard Executivo"
    ws1.views.sheetView[0].showGridLines = True

    ws2 = wb.create_sheet(title="Curva ABC & Pareto")
    ws2.views.sheetView[0].showGridLines = True

    ws3 = wb.create_sheet(title="Divergências Críticas")
    ws3.views.sheetView[0].showGridLines = True

    ws4 = wb.create_sheet(title="Dados Brutos Auditados")
    ws4.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Segoe UI", size=16, bold=True, color="0F172A")
    subtitle_font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
    data_font = Font(name="Segoe UI", size=10)
    thin_border = Border(left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'), top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1'))

    # Sheet 1
    ws1['A1'] = "NOVA-CORE // PAINEL DE CONTROLO EXECUTIVO"
    ws1['A1'].font = title_font
    ws1['A2'] = "Sumário Geral de Indicadores, Balanço Patrimonial e Governança WMS"
    ws1['A2'].font = subtitle_font

    total_skus = len(df)
    estoque_antes = (df['Congelado'] * df['Custo']).sum() if 'Congelado' in df.columns and 'Custo' in df.columns else 0
    estoque_depois = df['Valor de Estoque(R$)'].sum() if 'Valor de Estoque(R$)' in df.columns else 0
    ajuste_liquido = df['Valor de Ajuste(R$)'].sum() if 'Valor de Ajuste(R$)' in df.columns else 0
    divergencia_pct = (ajuste_liquido / estoque_antes) * 100 if estoque_antes > 0 else 0
    sobras_count = len(df[df['Ajuste'] > 0]) if 'Ajuste' in df.columns else 0
    faltas_count = len(df[df['Ajuste'] < 0]) if 'Ajuste' in df.columns else 0
    vol_fisico_div = df['Ajuste'].abs().sum() if 'Ajuste' in df.columns else 0
    acuracia_vol = ((total_skus - (sobras_count + faltas_count)) / total_skus) * 100 if total_skus > 0 else 0

    kpis = [
        ("Métrica Estratégica", "Valor Consolidado", "Parâmetro / Meta Corporativa", "Status de Auditoria"),
        ("Total SKUs Auditados", total_skus, "Universo Operacional Completo", "Concluído"),
        ("Estoque Teórico (Antes)", estoque_antes, "Saldo Sistema Pré-Inventário", "Validado"),
        ("Estoque Físico (Depois)", estoque_depois, "Saldo Efetivo Auditado", "Concluído"),
        ("Ajuste / Divergência Líquida", ajuste_liquido, "Impacto Contábil na DRE", "Lançado"),
        ("Percentual de Desvio Geral", divergencia_pct / 100, "Tolerância Máxima: ±1.0%", "Alerta Moderado"),
        ("Produtos com Sobra (Positivos)", sobras_count, "Itens com Excesso Físico", "Revisado"),
        ("Produtos com Falta (Negativos)", faltas_count, "Itens com Quebra Física", "Crítico"),
        ("Volume Físico Divergente", vol_fisico_div, "Soma Absoluta de Desvios (un)", "Monitoramento"),
        ("Acuracidade Volumétrica", acuracia_vol / 100, "Meta Corporativa: 95.0%", "Abaixo da Meta")
    ]

    ws1.append([])
    for row_data in kpis:
        ws1.append(row_data)

    for col_num in range(1, 5):
        cell = ws1.cell(row=4, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws1.iter_rows(min_row=5, max_row=13, min_col=1, max_col=4):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column in [2, 3] and isinstance(cell.value, (int, float)):
                if "Estoque" in str(ws1.cell(row=cell.row, column=1).value) or "Ajuste" in str(ws1.cell(row=cell.row, column=1).value):
                    cell.number_format = 'R$ #,##0.00'
                elif "Percentual" in str(ws1.cell(row=cell.row, column=1).value) or "Acuracidade" in str(ws1.cell(row=cell.row, column=1).value):
                    cell.number_format = '0.00%'
                else:
                    cell.number_format = '#,##0'

    # Sheet 2: Curva ABC
    ws2['A1'] = "ANÁLISE DE CONCENTRAÇÃO FINANCEIRA & CURVA ABC (PARETO 80/20)"
    ws2['A1'].font = title_font
    if 'Valor de Estoque(R$)' in df.columns:
        df_abc = df.sort_values(by='Valor de Estoque(R$)', ascending=False).copy()
        df_abc['Soma_Acumulada'] = df_abc['Valor de Estoque(R$)'].cumsum()
        soma_total = df_abc['Valor de Estoque(R$)'].sum()
        df_abc['Pct_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
        df_abc['Classe_ABC'] = df_abc['Pct_Acumulado'].apply(lambda p: 'Classe A (Alto Valor / 80%)' if p <= 80 else ('Classe B (Valor Médio / 15%)' if p <= 95 else 'Classe C (Baixo Valor / 5%)'))
        
        abc_headers = ["Código", "Descrição do Produto", "Emb", "Valor de Estoque (R$)", "% Acumulado", "Classe ABC"]
        ws2.append([])
        ws2.append(abc_headers)
        for col_num in range(1, len(abc_headers) + 1):
            cell = ws2.cell(row=3, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
        for idx, row in df_abc.head(200).iterrows():
            ws2.append([row.get('Código do Produto', ''), row.get('Descrição do Produto', ''), row.get('Emb', ''), row.get('Valor de Estoque(R$)', 0), row.get('Pct_Acumulado', 0) / 100, row.get('Classe_ABC', '')])

    # Sheet 3: Divergências Críticas
    ws3['A1'] = "RELATÓRIO DE DIVERGÊNCIAS CRÍTICAS (SOBRAS E FALTAS)"
    ws3['A1'].font = title_font
    if 'Ajuste' in df.columns:
        df_div = df[df['Ajuste'] != 0].sort_values(by='Valor de Ajuste(R$)', key=abs, ascending=False).copy() if 'Valor de Ajuste(R$)' in df.columns else df[df['Ajuste'] != 0]
        div_headers = ["Código", "Descrição do Produto", "Emb", "Estoque Pré", "Estoque Pós", "Ajuste (un)", "Custo Unit. (R$)", "Valor de Ajuste (R$)"]
        ws3.append([])
        ws3.append(div_headers)
        for col_num in range(1, len(div_headers) + 1):
            cell = ws3.cell(row=3, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
        for idx, row in df_div.head(300).iterrows():
            ws3.append([row.get('Código do Produto', ''), row.get('Descrição do Produto', ''), row.get('Emb', ''), row.get('Congelado', 0), row.get('Digitado', 0), row.get('Ajuste', 0), row.get('Custo', 0), row.get('Valor de Ajuste(R$)', 0)])

    # Sheet 4: Dados Brutos
    ws4['A1'] = "BASE DE DADOS COMPLETA AUDITADA (NOVA-CORE)"
    ws4['A1'].font = title_font
    raw_headers = list(df.columns)
    ws4.append([])
    ws4.append(raw_headers)
    for col_num in range(1, len(raw_headers) + 1):
        cell = ws4.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
    for idx, row in df.iterrows():
        ws4.append(list(row))

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    val_str = str(cell.value)
                    if len(val_str) > max_len:
                        max_len = len(val_str)
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_docx_bytes(df):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    p_title = doc.add_paragraph()
    r = p_title.add_run("⚡ NOVA-CORE // RELATÓRIO EXECUTIVO DE AUDITORIA COMPLETA")
    r.font.name = 'Arial'
    r.font.size = Pt(20)
    r.font.bold = True
    r.font.color.rgb = RGBColor(15, 23, 42)

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("Protocolo J.A.R.V.I.S. // Governança Patrimonial, Risco de Caixa e Plano de Ação C-Level")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(10)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(2, 132, 199)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    h1 = doc.add_heading("1. Sumário Executivo & Exposição Patrimonial", level=1)
    h1.runs[0].font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph("Este documento consolida a auditoria de inventário dos recintos operacionais com foco em governança WMS, controle de divergências físicas e impacto financeiro nas demonstrações contábeis.")

    table_kpi = doc.add_table(rows=1, cols=4)
    table_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table_kpi.rows[0].cells
    for i, h in enumerate(["Indicador Estratégico", "Valor Auditado", "Variação / Impacto", "Status / Parecer"]):
        hdr[i].text = h
        shading = parse_xml(r'<w:shd {} w:fill="0F172A"/>'.format(nsdecls('w')))
        hdr[i]._tc.get_or_add_tcPr().append(shading)
        hdr[i].paragraphs[0].runs[0].font.bold = True
        hdr[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    total_skus = len(df)
    estoque_antes = (df['Congelado'] * df['Custo']).sum() if 'Congelado' in df.columns and 'Custo' in df.columns else 0
    estoque_depois = df['Valor de Estoque(R$)'].sum() if 'Valor de Estoque(R$)' in df.columns else 0
    ajuste_liquido = df['Valor de Ajuste(R$)'].sum() if 'Valor de Ajuste(R$)' in df.columns else 0
    divergencia_pct = (ajuste_liquido / estoque_antes) * 100 if estoque_antes > 0 else 0
    sobras_count = len(df[df['Ajuste'] > 0]) if 'Ajuste' in df.columns else 0
    faltas_count = len(df[df['Ajuste'] < 0]) if 'Ajuste' in df.columns else 0
    vol_fisico_div = df['Ajuste'].abs().sum() if 'Ajuste' in df.columns else 0
    acuracia_vol = ((total_skus - (sobras_count + faltas_count)) / total_skus) * 100 if total_skus > 0 else 0

    rows_data = [
        ["Estoque Total Auditado (Pós)", f"R$ {estoque_depois:,.2f}", f"{total_skus:,} SKUs", "Concluído"],
        ["Estoque Teórico (Antes)", f"R$ {estoque_antes:,.2f}", "Saldo Pré-Inventário", "Validado"],
        ["Divergência Financeira Líquida", f"R$ {ajuste_liquido:,.2f}", f"{divergencia_pct:.2f}%", "Ajuste em DRE"],
        ["Volume Físico Divergente", f"{vol_fisico_div:,.0f} unidades", "Alerta Vermelho", "Monitoramento"],
        ["Acuracidade Volumétrica", f"{acuracia_vol:.2f}%", "Meta Corporativa: 95%", "Em Análise"]
    ]

    for row in rows_data:
        cells = table_kpi.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = val

    doc.add_paragraph().paragraph_format.space_after = Pt(10)
    h2 = doc.add_heading("2. Análise Estratégica & Curva ABC (Pareto 80/20)", level=1)
    h2.runs[0].font.color.rgb = RGBColor(15, 23, 42)
    doc.add_paragraph("A Curva ABC revela que a concentração de valor no inventário exige contagens cíclicas rigorosas nos SKUs da Classe A, mitigando riscos de ruptura e desvios operacionais.")

    if 'Valor de Estoque(R$)' in df.columns:
        table_abc = doc.add_table(rows=1, cols=4)
        table_abc.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_abc = table_abc.rows[0].cells
        for i, h in enumerate(["Código", "Descrição do Produto", "Valor de Estoque (R$)", "Classe ABC"]):
            hdr_abc[i].text = h
            shading = parse_xml(r'<w:shd {} w:fill="0F172A"/>'.format(nsdecls('w')))
            hdr_abc[i]._tc.get_or_add_tcPr().append(shading)
            hdr_abc[i].paragraphs[0].runs[0].font.bold = True
            hdr_abc[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

        df_abc = df.sort_values(by='Valor de Estoque(R$)', ascending=False).copy()
        df_abc['Soma_Acumulada'] = df_abc['Valor de Estoque(R$)'].cumsum()
        soma_total = df_abc['Valor de Estoque(R$)'].sum()
        df_abc['Pct_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
        df_abc['Classe_ABC'] = df_abc['Pct_Acumulado'].apply(lambda p: 'Classe A' if p <= 80 else ('Classe B' if p <= 95 else 'Classe C'))

        for idx, row in df_abc.head(10).iterrows():
            cells = table_abc.add_row().cells
            cells[0].text = str(row.get('Código do Produto', ''))
            cells[1].text = str(row.get('Descrição do Produto', ''))
            cells[2].text = f"R$ {row.get('Valor de Estoque(R$)', 0):,.2f}"
            cells[3].text = str(row.get('Classe_ABC', ''))

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_pdf_bytes(df):
    buffer = io.BytesIO()
    doc_pdf = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#0F172A'), spaceAfter=4)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#0284C7'), spaceAfter=14, fontName='Helvetica-Bold')
    h_style = ParagraphStyle('HeadingStyle', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0F172A'), spaceBefore=10, spaceAfter=4)
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#334155'), spaceAfter=8, leading=12)

    story.append(Paragraph("⚡ NOVA-CORE // RELATÓRIO EXECUTIVO DE AUDITORIA COMPLETA", title_style))
    story.append(Paragraph("Protocolo J.A.R.V.I.S. // Governança Patrimonial, Risco de Caixa e Plano de Ação C-Level", subtitle_style))
    story.append(Paragraph("1. Sumário Executivo & Exposição Patrimonial", h_style))
    story.append(Paragraph("Este relatório consolida a auditoria de inventário com foco rigoroso em governança WMS, controle de divergências físicas e impacto financeiro nas demonstrações contábeis.", body_style))

    total_skus = len(df)
    estoque_antes = (df['Congelado'] * df['Custo']).sum() if 'Congelado' in df.columns and 'Custo' in df.columns else 0
    estoque_depois = df['Valor de Estoque(R$)'].sum() if 'Valor de Estoque(R$)' in df.columns else 0
    ajuste_liquido = df['Valor de Ajuste(R$)'].sum() if 'Valor de Ajuste(R$)' in df.columns else 0
    divergencia_pct = (ajuste_liquido / estoque_antes) * 100 if estoque_antes > 0 else 0
    sobras_count = len(df[df['Ajuste'] > 0]) if 'Ajuste' in df.columns else 0
    faltas_count = len(df[df['Ajuste'] < 0]) if 'Ajuste' in df.columns else 0
    vol_fisico_div = df['Ajuste'].abs().sum() if 'Ajuste' in df.columns else 0
    acuracia_vol = ((total_skus - (sobras_count + faltas_count)) / total_skus) * 100 if total_skus > 0 else 0

    pdf_table_data = [
        ["Indicador Estratégico", "Valor Auditado", "Variação / Impacto", "Status / Parecer"],
        ["Estoque Total Auditado (Pós)", f"R$ {estoque_depois:,.2f}", f"{total_skus:,} SKUs", "Concluído"],
        ["Estoque Teórico (Antes)", f"R$ {estoque_antes:,.2f}", "Saldo Pré-Inventário", "Validado"],
        ["Divergência Financeira Líquida", f"R$ {ajuste_liquido:,.2f}", f"{divergencia_pct:.2f}%", "Ajuste em DRE"],
        ["Volume Físico Divergente", f"{vol_fisico_div:,.0f} unidades", "Alerta Vermelho", "Monitoramento"],
        ["Acuracidade Volumétrica", f"{acuracia_vol:.2f}%", "Meta Corporativa: 95%", "Em Análise"]
    ]

    t_pdf = Table(pdf_table_data, colWidths=[160, 110, 120, 142])
    t_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 8),
        ('BOTTOMPADDING', (0,0), (-1,0), 5),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F8FAFC')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))

    story.append(t_pdf)
    story.append(Spacer(1, 10))
    story.append(Paragraph("2. Top 10 SKUs Críticos por Concentração de Valor (Curva ABC)", h_style))

    if 'Valor de Estoque(R$)' in df.columns:
        df_abc = df.sort_values(by='Valor de Estoque(R$)', ascending=False).copy()
        df_abc['Soma_Acumulada'] = df_abc['Valor de Estoque(R$)'].cumsum()
        soma_total = df_abc['Valor de Estoque(R$)'].sum()
        df_abc['Pct_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
        df_abc['Classe_ABC'] = df_abc['Pct_Acumulado'].apply(lambda p: 'Classe A' if p <= 80 else ('Classe B' if p <= 95 else 'Classe C'))

        pdf_abc_data = [["Código", "Descrição do Produto", "Estoque Físico", "Valor de Estoque (R$)", "Classe"]]
        for idx, row in df_abc.head(10).iterrows():
            pdf_abc_data.append([
                str(row.get('Código do Produto', '')),
                str(row.get('Descrição do Produto', ''))[:35],
                f"{row.get('Saldo em Estoque', 0):,.2f}",
                f"R$ {row.get('Valor de Estoque(R$)', 0):,.2f}",
                str(row.get('Classe_ABC', ''))
            ])

        t_abc = Table(pdf_abc_data, colWidths=[60, 200, 75, 100, 97])
        t_abc.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 8),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),
            ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#F8FAFC')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,1), (-1,-1), 8),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t_abc)

    doc_pdf.build(story)
    buffer.seek(0)
    return buffer

def gerar_pptx_bytes(df):
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    slide_layout = prs.slide_layouts[6]

    # Slide 1: Capa
    slide1 = prs.slides.add_slide(slide_layout)
    slide1.background.fill.solid()
    slide1.background.fill.fore_color.rgb = PptxRGBColor(11, 15, 25)

    tb1 = slide1.shapes.add_textbox(PptxInches(1), PptxInches(2), PptxInches(11.3), PptxInches(3.5))
    tf1 = tb1.text_frame
    p0 = tf1.paragraphs[0]
    p0.text = "⚡ NOVA-CORE // DIRETORIA EXECUTIVA & C-LEVEL"
    p0.font.size = Pt(14)
    p0.font.bold = True
    p0.font.color.rgb = PptxRGBColor(0, 242, 254)

    p1 = tf1.add_paragraph()
    p1.text = "Relatório Executivo Completo de Auditoria & Governança Patrimonial"
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = PptxRGBColor(243, 244, 246)
    p1.space_before = Pt(15)

    # Slide 2: Indicadores Principais
    slide2 = prs.slides.add_slide(slide_layout)
    slide2.background.fill.solid()
    slide2.background.fill.fore_color.rgb = PptxRGBColor(11, 15, 25)

    tb2 = slide2.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(6))
    tf2 = tb2.text_frame
    p_s2 = tf2.paragraphs[0]
    p_s2.text = "📌 PAINEL DE INDICADORES EXECUTIVOS (CONSOLIDADO)"
    p_s2.font.size = Pt(20)
    p_s2.font.bold = True
    p_s2.font.color.rgb = PptxRGBColor(0, 242, 254)

    total_skus = len(df)
    estoque_antes = (df['Congelado'] * df['Custo']).sum() if 'Congelado' in df.columns and 'Custo' in df.columns else 0
    estoque_depois = df['Valor de Estoque(R$)'].sum() if 'Valor de Estoque(R$)' in df.columns else 0
    ajuste_liquido = df['Valor de Ajuste(R$)'].sum() if 'Valor de Ajuste(R$)' in df.columns else 0

    metrics = [
        f"• Total SKUs Auditados: {total_skus:,}".replace(',', '.'),
        f"• Estoque Teórico (Antes): R$ {estoque_antes:,.2f}".replace(',', '.'),
        f"• Estoque Físico (Depois): R$ {estoque_depois:,.2f}".replace(',', '.'),
        f"• Ajuste / Divergência Líquida: R$ {ajuste_liquido:,.2f}".replace(',', '.'),
        f"• Acuracidade Volumétrica: {(((total_skus - len(df[df['Ajuste'] != 0])) / total_skus) * 100):.2f}%".replace('.', ',') if 'Ajuste' in df.columns else "• Acuracidade Volumétrica: N/A"
    ]

    for m in metrics:
        p_m = tf2.add_paragraph()
        p_m.text = m
        p_m.font.size = Pt(18)
        p_m.font.color.rgb = PptxRGBColor(243, 244, 246)
        p_m.space_before = Pt(12)

    # Slide 3: Curva ABC Top Itens
    slide3 = prs.slides.add_slide(slide_layout)
    slide3.background.fill.solid()
    slide3.background.fill.fore_color.rgb = PptxRGBColor(11, 15, 25)

    tb3 = slide3.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(6))
    tf3 = tb3.text_frame
    p_s3 = tf3.paragraphs[0]
    p_s3.text = "🎯 DESTAQUES DA CURVA ABC (CLASSE A)"
    p_s3.font.size = Pt(20)
    p_s3.font.bold = True
    p_s3.font.color.rgb = PptxRGBColor(0, 242, 254)

    if 'Valor de Estoque(R$)' in df.columns:
        df_abc = df.sort_values(by='Valor de Estoque(R$)', ascending=False).head(5)
        for idx, row in df_abc.iterrows():
            p_item = tf3.add_paragraph()
            p_item.text = f"• [{row.get('Código do Produto', '')}] {row.get('Descrição do Produto', '')} — R$ {row.get('Valor de Estoque(R$)', 0):,.2f}".replace(',', '.')
            p_item.font.size = Pt(14)
            p_item.font.color.rgb = PptxRGBColor(243, 244, 246)
            p_item.space_before = Pt(8)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def main():
    st.set_page_config(
        page_title="NOVA-CORE // J.A.R.V.I.S. EXECUTIVO",
        page_icon="⚡",
        layout="wide"
    )

    st.markdown("""
<style>
.stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
.main-header { padding: 20px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px; }
.jarvis-title {
    font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
    background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
}
.bento-card {
    background: rgba(17, 24, 39, 0.7); 
    border: 1px solid rgba(31, 41, 55, 0.8);
    border-top: 3px solid #00f2fe;
    border-radius: 10px; 
    padding: 18px; 
    margin-bottom: 15px;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
}
</style>
""", unsafe_allow_html=True)

    st.markdown("""
<div class="main-header">
    <div class="jarvis-title">⚡ NOVA-CORE // DIRETORIA EXECUTIVA</div>
    <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">AUDITORIA DE ESTOQUE & ANÁLISE COMPARATIVA MULTI-INVENTÁRIO</div>
</div>
""", unsafe_allow_html=True)

    st.sidebar.markdown("### 🎛️ CONTROLES & NAVEGAÇÃO")
    arquivos_submetidos = st.sidebar.file_uploader("📂 Ingestão de Ficheiros", type=['csv', 'xlsx', 'xls', 'json', 'xml'], accept_multiple_files=True)
    
    modo = st.sidebar.radio("Selecione o Modo:", [
        "Modo 1: Indicadores Executivos & Balanço", 
        "Modo 2: Curva ABC, Cobertura & Divergências Críticas"
    ])

    if arquivos_submetidos:
        lista_dfs = []
        nomes_arquivos = []
        for arq in arquivos_submetidos:
            df_temp, msg = processar_arquivo_upload(arq)
            if df_temp is not None:
                lista_dfs.append(df_temp)
                nomes_arquivos.append(arq.name)
        
        if lista_dfs:
            df_filtrado = lista_dfs[0]
            num_cols = df_filtrado.select_dtypes(include=['number']).columns.tolist()
            cat_cols = df_filtrado.select_dtypes(include=['object']).columns.tolist()
            all_cols = df_filtrado.columns.tolist()

            total_pos = len(df_filtrado)
            idx_val = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
            col_val = num_cols[idx_val] if num_cols else None
            soma_depois = df_filtrado[col_val].sum() if col_val and col_val in df_filtrado.columns else 0

            idx_qtd = identificar_coluna_ideal(num_cols, ['quantidade', 'saldo', 'estoque', 'fisico', 'teorico'])
            col_qtd_pos = num_cols[idx_qtd] if num_cols else None
            total_itens_pos = df_filtrado[col_qtd_pos].sum() if col_qtd_pos and col_qtd_pos in df_filtrado.columns else total_pos
            
            col_ajuste = next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' not in c.lower()), 'Ajuste' if 'Ajuste' in df_filtrado.columns else None)
            soma_ajuste_qtd = df_filtrado[col_ajuste].sum() if col_ajuste and col_ajuste in df_filtrado.columns else 0
            total_itens_antes = total_itens_pos - soma_ajuste_qtd

            df_falta = df_filtrado[df_filtrado['Ajuste'] < 0] if 'Ajuste' in df_filtrado.columns else pd.DataFrame()
            df_sobra = df_filtrado[df_filtrado['Ajuste'] > 0] if 'Ajuste' in df_filtrado.columns else pd.DataFrame()
            
            qtd_falta = len(df_falta)
            qtd_sobra = len(df_sobra)
            vol_falta_un = abs(df_falta[col_ajuste].sum()) if col_ajuste and not df_falta.empty else 0
            vol_sobra_un = df_sobra[col_ajuste].sum() if col_ajuste and not df_sobra.empty else 0

            vol_total_div = vol_falta_un + vol_sobra_un
            perc_vol_div = (vol_total_div / total_itens_antes * 100) if total_itens_antes > 0 else 0

            qtd_exatos = (df_filtrado['Ajuste'] == 0).sum() if 'Ajuste' in df_filtrado.columns else int(total_pos * 0.5)
            total_divergencias = qtd_falta + qtd_sobra
            perc_divergencia = (total_divergencias / total_pos * 100) if total_pos > 0 else 0
            
            perc_falta_itens = (qtd_falta / total_pos * 100) if total_pos > 0 else 0
            perc_sobra_itens = (qtd_sobra / total_pos * 100) if total_pos > 0 else 0

            if 'Congelado' in df_filtrado.columns and 'Custo' in df_filtrado.columns:
                soma_antes = (df_filtrado['Congelado'] * df_filtrado['Custo']).sum()
            else:
                soma_ajuste_val = df_filtrado['Valor de Ajuste(R$)'].sum() if 'Valor de Ajuste(R$)' in df_filtrado.columns else 0
                soma_antes = soma_depois - soma_ajuste_val

            diferenca_monetaria = soma_depois - soma_antes
            impacto_perc = (diferenca_monetaria / soma_antes) * 100 if soma_antes > 0 else 0

            col_val_aju = 'Valor de Ajuste(R$)' if 'Valor de Ajuste(R$)' in df_filtrado.columns else None
            val_falta_fin = df_falta[col_val_aju].sum() if col_val_aju and not df_falta.empty else 0
            val_sobra_fin = df_sobra[col_val_aju].sum() if col_val_aju and not df_sobra.empty else 0
            soma_absoluta_val = abs(val_falta_fin) + abs(val_sobra_fin)

            acur_vol_real = (qtd_exatos / total_pos * 100) if total_pos > 0 else 0
            acur_fin_real = max(0.0, min(100.0, 100.0 - abs(impacto_perc) * 2.5))

            # ====================================================================
            # MODO 1: INDICADORES EXECUTIVOS & BALANÇO
            # ====================================================================
            if modo == "Modo 1: Indicadores Executivos & Balanço":
                st.markdown("### 📌 MODO 1: Indicadores Executivos & Parecer Técnico")
                st.markdown("<br>", unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{total_pos:,}".replace(',', '.'), f"Pré: {total_itens_antes:,.0f} un | Pós: {total_itens_pos:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{qtd_falta:,} SKUs ({perc_falta_itens:.2f}%)", f"Vol: -{vol_falta_un:,.0f} un | R$ {val_falta_fin:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Volume Físico Total de Divergências (Sobra + Falta)", f"{vol_total_div:,.0f} unidades ({perc_vol_div:.2f}%)".replace('.', ','), f"R$ {soma_absoluta_val:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                with c2:
                    st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{qtd_sobra:,} SKUs ({perc_sobra_itens:.2f}%)", f"Vol: +{vol_sobra_un:,.0f} un | R$ +{val_sobra_fin:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Total Falta + Sobra (SKUs)", f"{total_divergencias:,} SKUs ({perc_divergencia:.2f}%)".replace('.', ','), "Total de Produtos Divergentes", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Volumétrica / Financeira", f"Vol: {acur_vol_real:.2f}% | Fin: {acur_fin_real:.2f}%".replace('.', ','), "Meta Corporativa 95%", eh_negativo=(acur_vol_real < 95), eh_positivo=(acur_vol_real >= 95)), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("##### 💰 Balanço Patrimonial & Divergência Monetária (Antes vs Depois)")

                col_kpi1, col_kpi2 = st.columns(2)
                with col_kpi1:
                    st.markdown(formatar_kpi_html("Estoque Antes (Sistema / Teórico)", f"R$ {soma_antes:,.2f}", f"Qtd Pré: {total_itens_antes:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Divergência / Ajuste Líquido", f"R$ {diferenca_monetaria:,.2f}", f"{impacto_perc:.2f}% Desvio", eh_negativo=(diferenca_monetaria < 0), eh_positivo=(diferenca_monetaria >= 0)), unsafe_allow_html=True)
                with col_kpi2:
                    st.markdown(formatar_kpi_html("Estoque Depois (Físico Auditado)", f"R$ {soma_depois:,.2f}", f"Qtd Pós: {total_itens_pos:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Impacto Geral", f"{impacto_perc:.2f}%", "Variação Apurada", eh_negativo=(impacto_perc < 0), eh_positivo=(impacto_perc >= 0)), unsafe_allow_html=True)

            # ====================================================================
            # MODO 2: CURVA ABC, COBERTURA & DIVERGÊNCIAS CRÍTICAS (EXPANDIDO)
            # ====================================================================
            elif modo == "Modo 2: Curva ABC, Cobertura & Divergências Críticas":
                st.markdown("### 📊 MODO 2: Análise Avançada (Curva ABC, Cobertura & Divergências)")
                
                # ETAPA 1: CURVA ABC & PARETO
                with st.expander("🎯 ETAPA 1: Concentração Financeira & Curva ABC (Pareto 80/20)", expanded=True):
                    if len(num_cols) >= 1:
                        idx_padrao_y = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
                        c_x2, c_y2 = st.columns(2)
                        with c_x2:
                            eixo_x_abc = st.selectbox("Eixo X (Identificador/SKU/Produto):", all_cols, index=0, key="abc_x_din")
                        with c_y2:
                            eixo_y_abc = st.selectbox("Eixo Y (Valor Monetário/Quantidade):", num_cols, index=idx_padrao_y, key="abc_y_din")
                        
                        df_abc = df_filtrado.sort_values(by=eixo_y_abc, ascending=False).copy()
                        df_abc['Soma_Acumulada'] = df_abc[eixo_y_abc].cumsum()
                        soma_total = df_abc[eixo_y_abc].sum()
                        
                        if soma_total > 0:
                            df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                            df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(lambda p: "Classe A (Alto Valor / 80%)" if p <= 80 else ("Classe B (Médio Valor / 15%)" if p <= 95 else "Classe C (Baixo Valor / 5%)"))
                            
                            filtro_classe = st.multiselect(
                                "Filtrar por Classe ABC:",
                                ["Classe A (Alto Valor / 80%)", "Classe B (Médio Valor / 15%)", "Classe C (Baixo Valor / 5%)"],
                                default=["Classe A (Alto Valor / 80%)", "Classe B (Médio Valor / 15%)", "Classe C (Baixo Valor / 5%)"],
                                key="filtro_abc_multiselect"
                            )
                            
                            if filtro_classe:
                                df_abc_filtrado = df_abc[df_abc['Classe_ABC'].isin(filtro_classe)]
                            else:
                                df_abc_filtrado = df_abc

                            fig_abc = px.bar(
                                df_abc_filtrado.head(100), 
                                x=eixo_x_abc, 
                                y=eixo_y_abc, 
                                color='Classe_ABC', 
                                template="plotly_dark", 
                                title=f"Classificação ABC — Eixo X: {eixo_x_abc} | Eixo Y: {eixo_y_abc}"
                            )
                            fig_abc.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(type='category'))
                            st.plotly_chart(fig_abc, use_container_width=True)

                # ETAPA 2: DIVERGÊNCIAS CRÍTICAS (SOBRAS E FALTAS)
                with st.expander("🚨 ETAPA 2: Top Divergências Críticas (Sobras & Faltas)", expanded=False):
                    if 'Ajuste' in df_filtrado.columns:
                        col_div1, col_div2 = st.columns(2)
                        with col_div1:
                            st.markdown("##### 🔴 Maiores Faltas (Quebras de Estoque)")
                            df_f = df_filtrado[df_filtrado['Ajuste'] < 0].sort_values(by='Valor de Ajuste(R$)', ascending=True).head(10) if 'Valor de Ajuste(R$)' in df_filtrado.columns else df_filtrado[df_filtrado['Ajuste'] < 0].head(10)
                            st.dataframe(df_f[['Código do Produto', 'Descrição do Produto', 'Ajuste', 'Valor de Ajuste(R$)']] if 'Código do Produto' in df_f.columns else df_f, use_container_width=True)
                        with col_div2:
                            st.markdown("##### 🟢 Maiores Sobras (Excessos Físicos)")
                            df_s = df_filtrado[df_filtrado['Ajuste'] > 0].sort_values(by='Valor de Ajuste(R$)', ascending=False).head(10) if 'Valor de Ajuste(R$)' in df_filtrado.columns else df_filtrado[df_filtrado['Ajuste'] > 0].head(10)
                            st.dataframe(df_s[['Código do Produto', 'Descrição do Produto', 'Ajuste', 'Valor de Ajuste(R$)']] if 'Código do Produto' in df_s.columns else df_s, use_container_width=True)

                # ETAPA 3: COBERTURA & RUPTURA
                with st.expander("⏱️ ETAPA 3: Prevenção de Ruptura & Cobertura de Estoque", expanded=False):
                    if len(num_cols) >= 1:
                        idx_cob = identificar_coluna_ideal(num_cols, ['saldo', 'quantidade', 'estoque'])
                        vazao_diaria = st.slider("Média estimada de consumo/saída diária por SKU (unidades):", min_value=1, max_value=500, value=25, key="vazao_mod2")
                        df_cob = df_filtrado.copy()
                        col_estoque_cob = num_cols[idx_cob]
                        df_cob['Dias_Cobertura'] = df_cob[col_estoque_cob] / vazao_diaria
                        
                        fig_cob = px.histogram(df_cob, x='Dias_Cobertura', nbins=30, color_discrete_sequence=['#00f2fe'], template="plotly_dark", title="Distribuição de Dias de Cobertura do Inventário")
                        fig_cob.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_cob, use_container_width=True)

            # ====================================================================
            # MÓDULO DE EXPORTAÇÃO EXECUTIVA C-LEVEL (.XLSX, .DOCX, .PDF, .PPTX)
            # ====================================================================
            st.markdown("---")
            st.markdown("### 📥 EXPORTAÇÃO DE RELATÓRIOS & APRESENTAÇÕES EXECUTIVAS C-LEVEL")
            st.info("💡 **Central de Exportação NOVA-CORE:** Descarregue os relatórios executivos oficiais e apresentações completos.")
            
            col_exp1, col_exp2, col_exp3, col_exp4 = st.columns(4)
            
            with col_exp1:
                excel_bytes = gerar_excel_bytes(lista_dfs[0])
                st.download_button(
                    label="📊 Excel Completo (.xlsx)",
                    data=excel_bytes,
                    file_name="Relatorio_Executivo_NovaCore.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_exp2:
                docx_bytes = gerar_docx_bytes(lista_dfs[0])
                st.download_button(
                    label="📄 Word Executivo (.docx)",
                    data=docx_bytes,
                    file_name="Relatorio_Executivo_NovaCore.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            with col_exp3:
                pdf_bytes = gerar_pdf_bytes(lista_dfs[0])
                st.download_button(
                    label="📑 Relatório PDF (.pdf)",
                    data=pdf_bytes,
                    file_name="Relatorio_Executivo_NovaCore.pdf",
                    mime="application/pdf"
                )

            with col_exp4:
                pptx_bytes = gerar_pptx_bytes(lista_dfs[0])
                st.download_button(
                    label="📊 Apresentação PPTX (.pptx)",
                    data=pptx_bytes,
                    file_name="Apresentacao_Executiva_NovaCore.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

        else:
            st.error("❌ Não foi possível ler os dados dos arquivos enviados.")

if __name__ == "__main__":
    main()
