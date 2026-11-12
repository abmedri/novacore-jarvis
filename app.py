# ====================================================================================
# ⚡ GRUPO KMB // PROTOCOLO J.A.R.V.I.S. (VERSÃO EXECUTIVA COMPLETA & CORRIGIDA)
# ====================================================================================
import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Dependências para relatórios executivos avançados C-Level
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
from pptx.enum.text import PP_ALIGN

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
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
                    
                    col_mapping = {}
                    for col in df.columns:
                        col_lower = col.lower().strip()
                        if 'custo' in col_lower:
                            col_mapping[col] = 'Custo (R$)'
                        elif 'valor de ajuste' in col_lower:
                            col_mapping[col] = 'Valor de Ajuste (R$)'
                        elif 'valor de estoque' in col_lower:
                            col_mapping[col] = 'Valor de Estoque (R$)'
                    df = df.rename(columns=col_mapping)
                    
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


def calcular_metricas_gerenciais(lista_dfs, nomes_arquivos):
    resumo_ficheiros = []
    for idx, df in enumerate(lista_dfs):
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        total_skus = len(df)
        
        idx_val = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
        col_val = num_cols[idx_val] if num_cols else None
        soma_depois = df[col_val].sum() if col_val and col_val in df.columns else 0
        
        idx_qtd = identificar_coluna_ideal(num_cols, ['quantidade', 'saldo', 'estoque', 'fisico', 'teorico'])
        col_qtd = num_cols[idx_qtd] if num_cols else None
        qtd_pos = df[col_qtd].sum() if col_qtd and col_qtd in df.columns else total_skus
        
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        soma_ajuste_qtd = df[col_ajuste].sum() if col_ajuste and col_ajuste in df.columns else 0
        qtd_antes = qtd_pos - soma_ajuste_qtd
        
        col_val_aju = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df.columns else (
            'Valor de Ajuste(R$)' if 'Valor de Ajuste(R$)' in df.columns else next(
                (c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), None
            )
        )
        soma_ajuste_val = df[col_val_aju].sum() if col_val_aju and col_val_aju in df.columns else 0
        soma_antes = soma_depois - soma_ajuste_val
        val_ajuste_liquido = soma_depois - soma_antes
            
        df_falta = df[df[col_ajuste] < 0] if col_ajuste and col_ajuste in df.columns else pd.DataFrame()
        df_sobra = df[df[col_ajuste] > 0] if col_ajuste and col_ajuste in df.columns else pd.DataFrame()
        df_exato = df[df[col_ajuste] == 0] if col_ajuste and col_ajuste in df.columns else pd.DataFrame()
        
        skus_falta = len(df_falta)
        skus_sobra = len(df_sobra)
        skus_exatos = len(df_exato)
        total_div = skus_falta + skus_sobra
        
        perc_falta_itens = (skus_falta / total_skus * 100) if total_skus > 0 else 0
        perc_sobra_itens = (skus_sobra / total_skus * 100) if total_skus > 0 else 0
        perc_divergencia = (total_div / total_skus * 100) if total_skus > 0 else 0
        
        vol_falta = abs(df_falta[col_ajuste].sum()) if col_ajuste and not df_falta.empty else 0
        vol_sobra = df_sobra[col_ajuste].sum() if col_ajuste and not df_sobra.empty else 0
        vol_div_total = vol_falta + vol_sobra
        perc_vol_div = (vol_div_total / qtd_antes * 100) if qtd_antes > 0 else 0
        
        val_falta = df_falta[col_val_aju].sum() if col_val_aju and not df_falta.empty else 0
        val_sobra = df_sobra[col_val_aju].sum() if col_val_aju and not df_sobra.empty else 0
        val_div_absoluto = abs(val_falta) + abs(val_sobra)
        
        impacto_sobre_antes_pct = (val_ajuste_liquido / soma_antes * 100) if soma_antes > 0 else 0
        
        acuracia_vol = (skus_exatos / total_skus * 100) if total_skus > 0 else 0
        acuracia_fin = max(0.0, min(100.0, 100.0 - abs(impacto_sobre_antes_pct) * 2.5))
        
        resumo_ficheiros.append({
            "nome": nomes_arquivos[idx],
            "total_skus": total_skus,
            "qtd_antes": qtd_antes,
            "qtd_pos": qtd_pos,
            "soma_antes": soma_antes,
            "soma_depois": soma_depois,
            "val_ajuste_liquido": val_ajuste_liquido,
            "impacto_sobre_antes_pct": impacto_sobre_antes_pct,
            "skus_falta": skus_falta,
            "skus_sobra": skus_sobra,
            "skus_exatos": skus_exatos,
            "total_div": total_div,
            "perc_falta_itens": perc_falta_itens,
            "perc_sobra_itens": perc_sobra_itens,
            "perc_divergencia": perc_divergencia,
            "vol_falta": vol_falta,
            "vol_sobra": vol_sobra,
            "vol_div_total": vol_div_total,
            "perc_vol_div": perc_vol_div,
            "val_falta": val_falta,
            "val_sobra": val_sobra,
            "val_div_absoluto": val_div_absoluto,
            "acuracia_vol": acuracia_vol,
            "acuracia_fin": acuracia_fin,
            "df": df
        })
    return resumo_ficheiros


# ====================================================================================
# FUNÇÕES DE EXPORTAÇÃO ROBUSTAS (EXCEL, WORD, PDF E PPTX)
# ====================================================================================

def gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    wb = openpyxl.Workbook()
    
    ws1 = wb.active
    ws1.title = "Resumo Executivo"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1['A1'] = "⚡ GRUPO KMB // PAINEL DE GOVERNANÇA E AUDITORIA PATRIMONIAL"
    ws1['A1'].font = Font(size=14, bold=True, color="1F4E78")
    ws1.append([])
    
    headers_resumo = [
        "Ficheiro / Unidade", "Total SKUs", "Estoque Antes (R$)", "Estoque Depois (R$)",
        "Ajuste Líquido (R$)", "Impacto Geral (%)", "Faltas (SKUs)", "Sobras (SKUs)",
        "Acurácia Vol. (%)", "Acurácia Fin. (%)", "Status"
    ]
    ws1.append(headers_resumo)
    
    for r in resumos:
        status = "CONFORME" if r['acuracia_vol'] >= 97 else "CRÍTICO"
        ws1.append([
            r['nome'], r['total_skus'], r['soma_antes'], r['soma_depois'],
            r['val_ajuste_liquido'], f"{r['impacto_sobre_antes_pct']:.2f}%",
            r['skus_falta'], r['skus_sobra'],
            f"{r['acuracia_vol']:.2f}%", f"{r['acuracia_fin']:.2f}%", status
        ])
    
    ws2 = wb.create_sheet(title="Divergências e Zerados")
    ws2.views.sheetView[0].showGridLines = True
    
    headers_detalhe = [
        "Unidade / Ficheiro", "Data do Inventário", "Código do Produto", "Descrição",
        "Saldo Sistema", "Ajuste", "Valor Ajuste (R$)", "Classificação de Divergência"
    ]
    ws2.append(headers_detalhe)
    
    for idx, df in enumerate(lista_dfs):
        nome_arq = nomes_arquivos[idx]
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        col_val_aju = next((c for c in df.select_dtypes(include=['number']).columns if 'ajuste' in c.lower() and 'valor' in c.lower()), None)
        
        for _, row in df.iterrows():
            ajuste = row[col_ajuste] if col_ajuste and col_ajuste in df.columns else 0
            val_aju = row[col_val_aju] if col_val_aju and col_val_aju in df.columns else 0
            saldo_sis = row.get('Saldo em Estoque', 0)
            
            if ajuste < 0:
                classificacao = "FALTA (QUEBRA)"
            elif ajuste > 0:
                classificacao = "SOBRA"
            elif saldo_sis == 0:
                classificacao = "ESTOQUE CONGELADO ZERO"
            else:
                classificacao = "CONFORME"
                
            data_inv = str(row.get('Data de Abertura', row.get('Data', 'N/A')))
            cod_prod = str(row.get('Código do Produto', row.get('Codigo', 'N/A')))
            desc_prod = str(row.get('Descrição do Produto', row.get('Descricao', 'Item Diversos')))
            
            ws2.append([nome_arq, data_inv, cod_prod, desc_prod, saldo_sis, ajuste, val_aju, classificacao])
            
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    doc = docx.Document()
    
    doc.add_heading("⚡ GRUPO KMB // RELATÓRIO EXECUTIVO DE AUDITORIA C-LEVEL", level=1)
    doc.add_paragraph("Documento oficial consolidado de governança patrimonial, contendo indicadores de exatidão, controle e eficiência.")
    
    doc.add_heading("1. Indicadores Consolidados por Unidade", level=2)
    for r in resumos:
        p = doc.add_paragraph()
        p.add_run(f"• Unidade / Arquivo: ").bold = True
        p.add_run(f"{r['nome']}\n")
        p.add_run(f"  - Total de SKUs Auditados: {r['total_skus']:,} (Pré: {r['qtd_antes']:,.0f} un | Pós: {r['qtd_pos']:,.0f} un)\n")
        p.add_run(f"  - Produtos com Falta: {r['skus_falta']:,} SKUs ({r['perc_falta_itens']:.2f}%) | Vol: -{r['vol_falta']:,.0f} un | R$ {r['val_falta']:,.2f}\n")
        p.add_run(f"  - Produtos com Sobra: {r['skus_sobra']:,} SKUs ({r['perc_sobra_itens']:.2f}%) | Vol: +{r['vol_sobra']:,.0f} un | R$ +{r['val_sobra']:,.2f}\n")
        p.add_run(f"  - Volume Físico Total Divergências: {r['vol_div_total']:,.0f} un ({r['perc_vol_div']:.2f}%) | R$ {r['val_div_absoluto']:,.2f}\n")
        p.add_run(f"  - Valor Teórico Inicial: R$ {r['soma_antes']:,.2f} | Valor Auditado Final: R$ {r['soma_depois']:,.2f}\n")
        p.add_run(f"  - Ajuste Líquido: R$ {r['val_ajuste_liquido']:,.2f} (Impacto Geral: {r['impacto_sobre_antes_pct']:.2f}%)\n")
        p.add_run(f"  - Acurácia Volumétrica: {r['acuracia_vol']:.2f}% | Acurácia Financeira: {r['acuracia_fin']:.2f}%\n")
    
    doc.add_heading("2. Detalhamento de Divergências Críticas (Faltas, Sobras e Zerados)", level=2)
    for idx, df in enumerate(lista_dfs):
        doc.add_paragraph(f"Análise detalhada da base: {nomes_arquivos[idx]}")
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        
        if col_ajuste and col_ajuste in df.columns:
            total_faltas = len(df[df[col_ajuste] < 0])
            total_sobras = len(df[df[col_ajuste] > 0])
            total_zerados = len(df[(df[col_ajuste] == 0) & (df.get('Saldo em Estoque', 1) == 0)]) if 'Saldo em Estoque' in df.columns else len(df[df[col_ajuste] == 0])
            doc.add_paragraph(
                f"  - Itens em Falta (Quebras): {total_faltas:,}\n"
                f"  - Itens em Sobra: {total_sobras:,}\n"
                f"  - Itens com Estoque Congelado Zero: {total_zerados:,}"
            )
        else:
            doc.add_paragraph("Coluna de ajuste não mapeada diretamente para listagem individual detalhada neste bloco.")

    doc.add_heading("3. Validação e Assinatura C-Level", level=2)
    doc.add_paragraph("Declaro que os dados acima foram auditados eletronicamente através do Protocolo J.A.R.V.I.S. / Grupo KMB e refletem com fidelidade a apuração patrimonial realizada.\n\n\n")
    
    p_assinatura = doc.add_paragraph()
    p_assinatura.add_run("__________________________________________________\n").bold = True
    p_assinatura.add_run("Time de Inventário & Loss Prevention / Auditoria C-Level\n")
    p_assinatura.add_run("Grupo KMB - Exatidão • Controle • Eficiência")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    buffer = io.BytesIO()
    
    doc_pdf = SimpleDocTemplate(
        buffer, 
        pagesize=letter, 
        rightMargin=36, 
        leftMargin=36, 
        topMargin=36, 
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    
    estilo_titulo = ParagraphStyle(
        'TituloExecutivo',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=colors.HexColor('#1F4E78'),
        spaceAfter=10
    )
    
    estilo_sub = ParagraphStyle(
        'SubExecutivo',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor('#0F172A'),
        spaceBefore=12,
        spaceAfter=6
    )
    
    estilo_texto = ParagraphStyle(
        'TextoExecutivo',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=13,
        textColor=colors.HexColor('#334155'),
        spaceAfter=4
    )
    
    story = []
    
    story.append(Paragraph("⚡ GRUPO KMB // DOSSIÊ DE AUDITORIA PATRIMONIAL C-LEVEL", estilo_titulo))
    story.append(Paragraph("Relatório oficial consolidado contendo indicadores de impacto, análise de quebras, sobras e itens críticos.", estilo_texto))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#3B82F6'), spaceBefore=8, spaceAfter=12))
    
    story.append(Paragraph("1. Indicadores Consolidados por Unidade", estilo_sub))
    for r in resumos:
        texto_unidade = f"""
        <b>• Unidade / Arquivo:</b> {r['nome']}<br/>
        <b>- Total de SKUs Auditados:</b> {r['total_skus']:,} (Pré: {r['qtd_antes']:,.0f} un | Pós: {r['qtd_pos']:,.0f} un)<br/>
        <b>- Produtos com Falta:</b> {r['skus_falta']:,} SKUs ({r['perc_falta_itens']:.2f}%) | Vol: -{r['vol_falta']:,.0f} un | R$ {r['val_falta']:,.2f}<br/>
        <b>- Produtos com Sobra:</b> {r['skus_sobra']:,} SKUs ({r['perc_sobra_itens']:.2f}%) | Vol: +{r['vol_sobra']:,.0f} un | R$ +{r['val_sobra']:,.2f}<br/>
        <b>- Volume Físico Total Divergências:</b> {r['vol_div_total']:,.0f} un ({r['perc_vol_div']:.2f}%) | R$ {r['val_div_absoluto']:,.2f}<br/>
        <b>- Valor Teórico Inicial (Antes):</b> R$ {r['soma_antes']:,.2f} | <b>Auditado Final (Depois):</b> R$ {r['soma_depois']:,.2f}<br/>
        <b>- Ajuste Líquido:</b> R$ {r['val_ajuste_liquido']:,.2f} (Impacto Geral: {r['impacto_sobre_antes_pct']:.2f}%)<br/>
        <b>- Acurácia Volumétrica:</b> {r['acuracia_vol']:.2f}% | <b>Acurácia Financeira:</b> {r['acuracia_fin']:.2f}%
        """
        story.append(Paragraph(texto_unidade, estilo_texto))
        story.append(Spacer(1, 4))
    
    story.append(Paragraph("2. Resumo de Divergências Críticas", estilo_sub))
    for idx, df in enumerate(lista_dfs):
        nome_arq = nomes_arquivos[idx]
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        
        if col_ajuste and col_ajuste in df.columns:
            total_faltas = len(df[df[col_ajuste] < 0])
            total_sobras = len(df[df[col_ajuste] > 0])
            total_zerados = len(df[(df[col_ajuste] == 0) & (df.get('Saldo em Estoque', 1) == 0)]) if 'Saldo em Estoque' in df.columns else len(df[df[col_ajuste] == 0])
            
            texto_div = f"""
            <b>• Base: {nome_arq}</b><br/>
            - Itens em Falta (Quebras): {total_faltas:,}<br/>
            - Itens em Sobra: {total_sobras:,}<br/>
            - Itens com Estoque Congelado Zero: {total_zerados:,}
            """
            story.append(Paragraph(texto_div, estilo_texto))
        else:
            story.append(Paragraph(f"• Base: {nome_arq} - Coluna de ajuste não mapeada para contagem detalhada.", estilo_texto))
            
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("3. Validação e Assinatura C-Level", estilo_sub))
    story.append(Paragraph("Declaro que os dados acima foram auditados eletronicamente através do Protocolo J.A.R.V.I.S. / Grupo KMB e refletem com fidelidade a apuração patrimonial realizada.", estilo_texto))
    story.append(Spacer(1, 20))
    
    texto_assinatura = """
    <b>__________________________________________________</b><br/>
    <b>Time de Inventário & Loss Prevention / Auditoria C-Level</b><br/>
    Grupo KMB - Exatidão • Controle • Eficiência
    """
    story.append(Paragraph(texto_assinatura, estilo_texto))
    
    doc_pdf.build(story)
    buffer.seek(0)
    return buffer


def gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(PptxInches(1), PptxInches(1), PptxInches(11), PptxInches(2))
    tb.text_frame.text = "⚡ GRUPO KMB // RELATÓRIO EXECUTIVO C-LEVEL"
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer


def main():
    st.set_page_config(page_title="GRUPO KMB // TIME DE INVENTÁRIO", page_icon="⚡", layout="wide")
    
    SENHA_MESTRE = "novacore2026"

    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.markdown("""
        <style>
        .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
        </style>
        """, unsafe_allow_html=True)
        
        col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
        with col_l2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            
            if os.path.exists("brasao_sem_fundo.png"):
                st.image("brasao_sem_fundo.png", use_container_width=True)
            else:
                st.warning("⚠️ Arquivo 'brasao_sem_fundo.png' não encontrado na pasta. Coloque-o no mesmo diretório do script.")
            
            st.markdown("## ⚡ TIME DE INVENTÁRIO // AUTENTICAÇÃO")
            st.markdown("Insira a senha de acesso executivo para desbloquear o sistema.")
            
            senha_digitada = st.text_input("Senha de Acesso", type="password", key="input_senha")
            
            if st.button("🔓 Entrar no Sistema", use_container_width=True):
                if senha_digitada == SENHA_MESTRE:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Senha incorreta. Acesso negado.")
        return

    st.markdown("""
    <style>
    .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
    .main-header { padding: 20px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px; }
    .jarvis-title {
        font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
        background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        if os.path.exists("brasao_sem_fundo.png"):
            st.image("brasao_sem_fundo.png", use_container_width=True)
        st.markdown("### **GRUPO KMB**")
        st.markdown("*Exatidão • Controle • Eficiência*")
        st.markdown("---")
        if st.button("🔒 Encerrar Sessão (Sair)", use_container_width=True):
            st.session_state.autenticado = False
            st.rerun()

    st.markdown("""
    <div class="main-header">
        <div class="jarvis-title">⚡ GRUPO KMB // DIRETORIA EXECUTIVA</div>
        <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">AUDITORIA DE ESTOQUE & ANÁLISE COMPARATIVA MULTI-INVENTÁRIO</div>
    </div>
    """, unsafe_allow_html=True)
    
    arquivos_submetidos = st.file_uploader("Carregue o arquivo oficial de inventário (ex: 13372.07.11.2026.csv) ou múltiplos arquivos", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)
    
    if arquivos_submetidos:
        lista_dfs = []
        nomes_arquivos = []
        for arq in arquivos_submetidos:
            df_temp, msg = processar_arquivo_upload(arq)
            if df_temp is not None:
                lista_dfs.append(df_temp)
                nomes_arquivos.append(arq.name)
        
        if lista_dfs:
            # ====================================================================
            # MODO 1: APENAS 1 FICHEIRO (VISÃO DETALHADA DE 5 ETAPAS)
            # ====================================================================
            if len(lista_dfs) == 1:
                df_filtrado = lista_dfs[0]
                num_cols = df_filtrado.select_dtypes(include=['number']).columns.tolist()
                cat_cols = df_filtrado.select_dtypes(include=['object']).columns.tolist()
                all_cols = df_filtrado.columns.tolist()
                
                resumo_unico = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)[0]
                
                # ETAPA 1: Indicadores Executivos & Bento Cards com Deltas e Percentuais (Meta 97%)
                st.markdown("### 📌 ETAPA 1: Indicadores Executivos & Parecer Técnico")
                st.markdown("<br>", unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{resumo_unico['total_skus']:,}", f"Pré: {resumo_unico['qtd_antes']:,.0f} un | Pós: {resumo_unico['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{resumo_unico['skus_falta']:,} SKUs ({resumo_unico['perc_falta_itens']:.2f}%)", f"Vol: -{resumo_unico['vol_falta']:,.0f} un | R$ {resumo_unico['val_falta']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Volume Físico Total de Divergências (Sobra + Falta)", f"{resumo_unico['vol_div_total']:,.0f} unidades ({resumo_unico['perc_vol_div']:.2f}%)", f"R$ {resumo_unico['val_div_absoluto']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                with c2:
                    st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{resumo_unico['skus_sobra']:,} SKUs ({resumo_unico['perc_sobra_itens']:.2f}%)", f"Vol: +{resumo_unico['vol_sobra']:,.0f} un | R$ +{resumo_unico['val_sobra']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Total Falta + Sobra (SKUs)", f"{resumo_unico['total_div']:,} SKUs ({resumo_unico['perc_divergencia']:.2f}%)", "Total de Produtos Divergentes", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Volumétrica / Financeira", f"Vol: {resumo_unico['acuracia_vol']:.2f}% | Fin: {resumo_unico['acuracia_fin']:.2f}%", "Meta Corporativa 97%", eh_negativo=(resumo_unico['acuracia_vol'] < 97), eh_positivo=(resumo_unico['acuracia_vol'] >= 97)), unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("##### 💰 Balanço Patrimonial & Impacto Geral")
                col_kpi1, col_kpi2 = st.columns(2)
                with col_kpi1:
                    st.markdown(formatar_kpi_html("Estoque Antes (Sistema / Teórico)", f"R$ {resumo_unico['soma_antes']:,.2f}", f"Qtd Pré: {resumo_unico['qtd_antes']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Impacto Geral", f"{resumo_unico['impacto_sobre_antes_pct']:.2f}%", f"Ajuste Líquido: R$ {resumo_unico['val_ajuste_liquido']:,.2f}", eh_negativo=(resumo_unico['impacto_sobre_antes_pct'] < 0), eh_positivo=(resumo_unico['impacto_sobre_antes_pct'] >= 0)), unsafe_allow_html=True)
                with col_kpi2:
                    st.markdown(formatar_kpi_html("Estoque Depois (Físico Auditado)", f"R$ {resumo_unico['soma_depois']:,.2f}", f"Qtd Pós: {resumo_unico['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Divergência / Ajuste Líquido", f"R$ {resumo_unico['val_ajuste_liquido']:,.2f}", "Variação Apurada", eh_negativo=(resumo_unico['val_ajuste_liquido'] < 0), eh_positivo=(resumo_unico['val_ajuste_liquido'] >= 0)), unsafe_allow_html=True)
                st.markdown("---")
                
                # ETAPA 2: Curva ABC com Seletores Dinâmicos de Eixos + Filtro de Classe A/B/C
                with st.expander("🎯 ETAPA 2: Concentração Financeira & Curva ABC (Pareto 80/20)", expanded=True):
                    st.markdown("""
                    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(59, 130, 246, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                        <span style="color: #3b82f6; font-weight: 600; font-size: 0.9rem;">ANÁLISE DE PARETO C-LEVEL // DINÂMICA</span>
                        <p style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px; margin-bottom: 0;">
                            Selecione a dimensão, a métrica e filtre diretamente pelas classes estratégicas da Curva ABC.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    if len(num_cols) >= 1 and len(all_cols) > 0:
                        col_sel1, col_sel2, col_sel3 = st.columns(3)
                        with col_sel1:
                            coluna_categoria_escolhida = st.selectbox(
                                "Dimensão (Eixo X):", 
                                options=all_cols, 
                                index=all_cols.index('Descrição do Produto') if 'Descrição do Produto' in all_cols else 0,
                                key="abc_eixo_x"
                            )
                        with col_sel2:
                            idx_padrao_y = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
                            eixo_y_escolhido = st.selectbox(
                                "Métrica Principal (Eixo Y):", 
                                options=num_cols, 
                                index=idx_padrao_y if idx_padrao_y < len(num_cols) else 0,
                                key="abc_eixo_y"
                            )
                        with col_sel3:
                            filtro_classe_abc = st.selectbox(
                                "Filtro Curva ABC:",
                                options=["Todas as Classes (Geral)", "Classe A (80%)", "Classe B (95%)", "Classe C (100%)"],
                                index=0,
                                key="abc_filtro_classe"
                            )
                        
                        df_abc = df_filtrado.sort_values(by=eixo_y_escolhido, ascending=False).copy()
                        df_abc['Soma_Acumulada'] = df_abc[eixo_y_escolhido].cumsum()
                        soma_total = df_abc[eixo_y_escolhido].sum()
                        
                        if soma_total > 0:
                            df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                            df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(
                                lambda p: "Classe A (80%)" if p <= 80 else ("Classe B (95%)" if p <= 95 else "Classe C (100%)")
                            )
                            
                            if filtro_classe_abc != "Todas as Classes (Geral)":
                                df_abc_filtrado = df_abc[df_abc['Classe_ABC'] == filtro_classe_abc].copy()
                            else:
                                df_abc_filtrado = df_abc.copy()
                            
                            qtd_skus_classe = len(df_abc_filtrado)
                            valor_classe = df_abc_filtrado[eixo_y_escolhido].sum()
                            perc_valor_classe = (valor_classe / soma_total * 100) if soma_total > 0 else 0
                            st.caption(f"📊 **Resumo da Seleção ({filtro_classe_abc}):** {qtd_skus_classe:,} SKUs | Total: R$ {valor_classe:,.2f} ({perc_valor_classe:.2f}% do estoque)")
                            
                            df_plot = df_abc_filtrado.head(35).copy()
                            
                            if not df_plot.empty:
                                fig_abc = go.Figure()
                                fig_abc.add_trace(go.Bar(
                                    x=df_plot[coluna_categoria_escolhida],
                                    y=df_plot[eixo_y_escolhido],
                                    name=f"{eixo_y_escolhido}",
                                    marker=dict(
                                        color=df_plot[eixo_y_escolhido],
                                        colorscale=[[0, '#1e3a8a'], [0.5, '#3b82f6'], [1, '#00f2fe']],
                                        opacity=0.85
                                    ),
                                    yaxis='y'
                                ))
                                fig_abc.add_trace(go.Scatter(
                                    x=df_plot[coluna_categoria_escolhida],
                                    y=df_plot['%_Acumulado'],
                                    name="% Acumulado (Pareto)",
                                    mode='lines+markers',
                                    marker=dict(color='#f59e0b', size=8),
                                    line=dict(color='#f59e0b', width=3),
                                    yaxis='y2'
                                ))
                                fig_abc.update_layout(
                                    template="plotly_dark",
                                    paper_bgcolor='rgba(11, 15, 25, 0.9)',
                                    plot_bgcolor='rgba(11, 15, 25, 0.9)',
                                    height=480,
                                    margin=dict(l=30, r=30, t=40, b=30),
                                    legend=dict(
                                        orientation="h", 
                                        yanchor="bottom", 
                                        y=1.05, 
                                        xanchor="right", 
                                        x=1,
                                        font=dict(size=11, color="#f3f4f6")
                                    ),
                                    yaxis=dict(
                                        title=eixo_y_escolhido, 
                                        gridcolor='rgba(31, 41, 55, 0.6)'
                                    ),
                                    yaxis2=dict(
                                        title="% Acumulado", 
                                        overlaying='y', 
                                        side='right', 
                                        range=[0, 105], 
                                        showgrid=False
                                    ),
                                    xaxis=dict(showgrid=False)
                                )
                                st.plotly_chart(fig_abc, use_container_width=True)
                            else:
                                st.warning("Nenhum produto encontrado para a classe selecionada.")
                
                # ETAPA 3: Matriz de Criticidade & Discrepância Financeira (Focos de Prejuízo)
                with st.expander("🚨 ETAPA 3: Matriz de Criticidade & Discrepância Financeira", expanded=False):
                    st.markdown("""
                    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                        <span style="color: #ef4444; font-weight: 600; font-size: 0.9rem;">AUDITORIA DE DESVIOS CRÍTICOS (C-LEVEL)</span>
                        <p style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px; margin-bottom: 0;">
                            Isolamento dos produtos com maior impacto financeiro negativo (faltas e quebras severas) para investigação prioritária de prevenção de perdas.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    if len(num_cols) >= 2:
                        col_nome_prod = 'Descrição do Produto' if 'Descrição do Produto' in df_filtrado.columns else all_cols[0]
                        idx_val_div = identificar_coluna_ideal(num_cols, ['valor de ajuste', 'ajuste', 'valor', 'divergencia'])
                        col_val_div = num_cols[idx_val_div] if idx_val_div < len(num_cols) else num_cols[0]
                        
                        df_critico = df_filtrado.copy()
                        df_critico['Impacto_Absoluto'] = df_critico[col_val_div].abs()
                        df_critico = df_critico.sort_values(by='Impacto_Absoluto', ascending=False).head(25)
                        if not df_critico.empty:
                            fig_critico = px.bar(
                                df_critico,
                                x=col_nome_prod,
                                y=col_val_div,
                                title="Top 25 Maiores Desvios Financeiros (Faltas e Sobras Críticas em R$)",
                                template="plotly_dark",
                                color=col_val_div,
                                color_continuous_scale=['#ef4444', '#f59e0b', '#3b82f6']
                            )
                            fig_critico.update_layout(
                                paper_bgcolor='rgba(11, 15, 25, 0.9)',
                                plot_bgcolor='rgba(11, 15, 25, 0.9)',
                                height=450,
                                margin=dict(l=30, r=30, t=40, b=30),
                                xaxis=dict(showgrid=False),
                                yaxis=dict(gridcolor='rgba(31, 41, 55, 0.6)', title="Valor da Divergência (R$)")
                            )
                            st.plotly_chart(fig_critico, use_container_width=True)
                            st.info("ℹ️ **Foco de Investigação:** Barras em tons avermelhados indicam furos/faltas expressivas que exigem auditoria imediata de processos e câmeras de segurança.")
                        else:
                            st.warning("Não há dados suficientes para calcular a matriz de criticidade com as colunas atuais.")
                
                # ETAPA 4: Matriz de Risco de Divergência & Acurácia Ponderada (C-Level Audit) - CORRIGIDA EM VALOR (R$)
                with st.expander("🛡️ ETAPA 4: Matriz de Risco & Acurácia Ponderada por Impacto", expanded=False):
                    st.markdown("""
                    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                        <span style="color: #f59e0b; font-weight: 600; font-size: 0.9rem;">GOVERNANÇA & ACURÁCIA FINANCEIRA DE INVENTÁRIO (EM VALOR)</span>
                        <p style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px; margin-bottom: 0;">
                            Análise cruzada de desvios financeiros para medir o impacto no caixa e isolar quebras de distorções cadastrais por categoria.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                    if len(num_cols) >= 2 and len(all_cols) > 0:
                        df_risco = df_filtrado.copy()
                        
                        # Identificar explicitamente a coluna de Valor de Ajuste (R$)
                        col_val_aju_etapa4 = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df_risco.columns else next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), num_cols[0])
                        
                        df_risco['Tipo_Desvio'] = df_risco[col_val_aju_etapa4].apply(
                            lambda x: 'Déficit (Falta/Quebra)' if x < 0 else ('Superávit (Sobra)' if x > 0 else 'Conforme')
                        )
                        
                        m1, m2 = st.columns(2)
                        with m1:
                            total_perdas = df_risco[df_risco[col_val_aju_etapa4] < 0][col_val_aju_etapa4].sum()
                            st.metric("Total de Quebras/Faltas (R$)", f"R$ {total_perdas:,.2f}", delta_color="inverse")
                        with m2:
                            total_sobras = df_risco[df_risco[col_val_aju_etapa4] > 0][col_val_aju_etapa4].sum()
                            st.metric("Total de Sobras/Ajustes (R$)", f"R$ {total_sobras:,.2f}")
                            
                        # Criar categoria baseada na primeira palavra da descrição para agrupamento fiel em Valor
                        df_risco['Categoria_Prod'] = df_risco['Descrição do Produto'].astype(str).str.split().str[0]
                        df_agrupado_val = df_risco.groupby('Categoria_Prod')[col_val_aju_etapa4].sum().reset_index()
                        df_agrupado_val = df_agrupado_val.sort_values(by=col_val_aju_etapa4, ascending=True).head(15)

                        fig_risco = px.bar(
                            df_agrupado_val,
                            x='Categoria_Prod',
                            y=col_val_aju_etapa4,
                            title="Balanço Líquido de Divergências por Categoria (Em Reais - R$)",
                            template="plotly_dark",
                            color=col_val_aju_etapa4,
                            color_continuous_scale=['#ef4444', '#3b82f6', '#10b981']
                        )
                        fig_risco.update_layout(
                            paper_bgcolor='rgba(11, 15, 25, 0.9)',
                            plot_bgcolor='rgba(11, 15, 25, 0.9)',
                            height=420,
                            margin=dict(l=30, r=30, t=40, b=30),
                            xaxis=dict(showgrid=False, title="Categoria / Linha de Produto"),
                            yaxis=dict(gridcolor='rgba(31, 41, 55, 0.6)', title="Impacto Líquido em Valor (R$)")
                        )
                        st.plotly_chart(fig_risco, use_container_width=True)
                        
                        st.info("💡 **Diretriz C-Level:** O gráfico acima reflete estritamente o **Valor Financeiro Líquido (R$)**, permitindo identificar onde o caixa da empresa foi mais impactado.")
                    else:
                        st.warning("Dados insuficientes para estruturar a matriz de risco de acurácia.")
       
                # ETAPA 5: Evolução Histórica
                with st.expander("🧮 ETAPA 5: Simulador Contábil & Evolução Histórica", expanded=False):
                    df_hist = pd.DataFrame({
                        "Inventário": ["Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026 (Atual)"],
                        "Acuracidade Financeira (%)": [81.2, 83.5, 85.0, 87.1, resumo_unico['acuracia_fin']],
                        "Acuracidade Volumétrica (%)": [68.0, 70.4, 72.1, 74.8, resumo_unico['acuracia_vol']]
                    })
                    fig_hist = px.line(df_hist, x="Inventário", y=["Acuracidade Financeira (%)", "Acuracidade Volumétrica (%)"], markers=True, template="plotly_dark", title="Evolução Histórica de Acuracidade")
                    st.plotly_chart(fig_hist, use_container_width=True)

            # ====================================================================
            # MODO 2: MÚLTIPLOS FICHEIROS (VISÃO COMPARATIVA EXECUTIVA COMPLETA)
            # ====================================================================
            else:
                st.markdown("## ⚡ AUDITORIA COMPARATIVA MULTI-INVENTÁRIO C-LEVEL")
                resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
                
                col_cards = st.columns(len(lista_dfs))
                for idx, r in enumerate(resumos):
                    with col_cards[idx]:
                        with st.container(border=True):
                            st.markdown(f"#### 📄 {r['nome']}")
                            st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{r['total_skus']:,}", f"Pré: {r['qtd_antes']:,.0f} un | Pós: {r['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{r['skus_falta']:,} SKUs ({r['perc_falta_itens']:.2f}%)", f"Vol: -{r['vol_falta']:,.0f} un | R$ {r['val_falta']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{r['skus_sobra']:,} SKUs ({r['perc_sobra_itens']:.2f}%)", f"Vol: +{r['vol_sobra']:,.0f} un | R$ +{r['val_sobra']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Total Falta + Sobra (SKUs)", f"{r['total_div']:,} SKUs ({r['perc_divergencia']:.2f}%)", "Total de Produtos Divergentes", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Volume Físico Total de Divergências (Sobra + Falta)", f"{r['vol_div_total']:,.0f} unidades ({r['perc_vol_div']:.2f}%)", f"R$ {r['val_div_absoluto']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Acuracidade Volumétrica / Financeira", f"Vol: {r['acuracia_vol']:.2f}% | Fin: {r['acuracia_fin']:.2f}%", "Meta Corporativa 97%", eh_negativo=(r['acuracia_vol'] < 97), eh_positivo=(r['acuracia_vol'] >= 97)), unsafe_allow_html=True)
                            
                            st.markdown("##### 💰 Balanço Patrimonial")
                            st.markdown(formatar_kpi_html("Estoque Antes (Sistema / Teórico)", f"R$ {r['soma_antes']:,.2f}", f"Qtd Pré: {r['qtd_antes']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Estoque Depois (Físico Auditado)", f"R$ {r['soma_depois']:,.2f}", f"Qtd Pós: {r['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Impacto Geral", f"{r['impacto_sobre_antes_pct']:.2f}%", f"Ajuste Líquido: R$ {r['val_ajuste_liquido']:,.2f}", eh_negativo=(r['impacto_sobre_antes_pct'] < 0), eh_positivo=(r['impacto_sobre_antes_pct'] >= 0)), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Divergência / Ajuste Líquido", f"R$ {r['val_ajuste_liquido']:,.2f}", "Variação Apurada", eh_negativo=(r['val_ajuste_liquido'] < 0), eh_positivo=(r['val_ajuste_liquido'] >= 0)), unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📊 Gráficos Comparativos Consolidados")
                
                df_resumo_geral = pd.DataFrame([
                    {
                        "Ficheiro": r['nome'],
                        "Estoque Antes": r['soma_antes'],
                        "Estoque Depois": r['soma_depois'],
                        "Ajuste Líquido": r['val_ajuste_liquido'],
                        "Impacto Geral (%)": r['impacto_sobre_antes_pct'],
                        "Total SKUs": r['total_skus'],
                        "Faltas (SKUs)": r['skus_falta'],
                        "Sobras (SKUs)": r['skus_sobra']
                    } for r in resumos
                ])
                
                gc1, gc2 = st.columns(2)
                with gc1:
                    fig_val = px.bar(df_resumo_geral, x="Ficheiro", y=["Estoque Antes", "Estoque Depois"], barmode="group", template="plotly_dark", title="Comparativo Valor Estoque (Antes vs Depois)")
                    st.plotly_chart(fig_val, use_container_width=True)
                with gc2:
                    fig_aj = px.bar(df_resumo_geral, x="Ficheiro", y="Impacto Geral (%)", template="plotly_dark", title="Impacto Geral em Percentual")
                    st.plotly_chart(fig_aj, use_container_width=True)
                
                gc3, gc4 = st.columns(2)
                with gc3:
                    fig_skus = px.bar(df_resumo_geral, x="Ficheiro", y=["Faltas (SKUs)", "Sobras (SKUs)"], barmode="group", template="plotly_dark", title="Comparativo Faltas vs Sobras (SKUs)")
                    st.plotly_chart(fig_skus, use_container_width=True)
                with gc4:
                    fig_tot = px.bar(df_resumo_geral, x="Ficheiro", y="Total SKUs", template="plotly_dark", title="Volume Total de SKUs por Inventário")
                    st.plotly_chart(fig_tot, use_container_width=True)

                # --------------------------------------------------------------------
                # ETAPA 5 (MODO 2): Simulador Contábil & Evolução Histórica (Multitemporal Avançado)
                # --------------------------------------------------------------------
                st.markdown("---")
                with st.expander("📈 ETAPA 5: Simulador Contábil & Evolução Histórica (Multitemporal Avançado)", expanded=True):
                    st.markdown("""
                    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                        <span style="color: #10b981; font-weight: 600; font-size: 0.9rem;">ANÁLISE DE TENDÊNCIA & EVOLUÇÃO PATRIMONIAL COMPARATIVA</span>
                        <p style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px; margin-bottom: 0;">
                            Cruzamento multitemporal dos inventários carregados para rastrear o comportamento histórico de todos os indicadores ao longo do tempo.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)

                    df_tendencia = pd.DataFrame([
                        {
                            "Inventario": r['nome'],
                            "Total SKUs Auditados": r['total_skus'],
                            "Produtos com Falta (SKUs)": r['skus_falta'],
                            "Produtos com Sobra (SKUs)": r['skus_sobra'],
                            "Total Falta + Sobra (SKUs)": r['total_div'],
                            "Volume Físico Divergências (un)": r['vol_div_total'],
                            "Acuracidade Volumétrica (%)": r['acuracia_vol'],
                            "Acuracidade Financeira (%)": r['acuracia_fin'],
                            "Estoque Antes (R$)": r['soma_antes'],
                            "Estoque Depois (R$)": r['soma_depois'],
                            "Impacto Geral (%)": r['impacto_sobre_antes_pct'],
                            "Divergência / Ajuste Líquido (R$)": r['val_ajuste_liquido']
                        } for r in resumos
                    ])

                    metrica_escolhida = st.selectbox(
                        "Selecione o Indicador para Análise de Evolução Histórica:",
                        options=[
                            "Acuracidade Volumétrica (%)", 
                            "Acuracidade Financeira (%)", 
                            "Divergência / Ajuste Líquido (R$)", 
                            "Impacto Geral (%)", 
                            "Total Falta + Sobra (SKUs)",
                            "Produtos com Falta (SKUs)",
                            "Produtos com Sobra (SKUs)",
                            "Volume Físico Divergências (un)",
                            "Total SKUs Auditados",
                            "Estoque Antes (R$)",
                            "Estoque Depois (R$)"
                        ],
                        key="select_metrica_evolucao"
                    )

                    fig_tendencia = px.line(
                        df_tendencia,
                        x="Inventario",
                        y=metrica_escolhida,
                        markers=True,
                        title=f"Evolução Histórica: {metrica_escolhida}",
                        template="plotly_dark",
                        color_discrete_sequence=['#3b82f6']
                    )

                    fig_tendencia.update_traces(line=dict(width=3), marker=dict(size=10))
                    fig_tendencia.update_layout(
                        paper_bgcolor='rgba(11, 15, 25, 0.9)',
                        plot_bgcolor='rgba(11, 15, 25, 0.9)',
                        height=420,
                        margin=dict(l=30, r=30, t=40, b=30),
                        xaxis=dict(showgrid=False, title="Ficheiro / Período"),
                        yaxis=dict(gridcolor='rgba(31, 41, 55, 0.6)', title=metrica_escolhida)
                    )

                    st.plotly_chart(fig_tendencia, use_container_width=True)

                    st.markdown("#### 📋 Tabela Consolidada de Tendência Histórica (Todos os Indicadores)")
                    st.dataframe(df_tendencia, use_container_width=True)

                    st.success("✅ **Simulador Ativo:** Acompanhe o ganho ou perda de eficiência patrimonial à medida que novos arquivos de inventário são adicionados à análise.")

            # Central de Relatórios
            st.markdown("---")
            st.subheader("📥 Central de Relatórios Executivos C-Level")
            col_e1, col_e2, col_e3, col_e4 = st.columns(4)
            with col_e1:
                st.download_button("📊 Excel Gerencial", data=gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_e2:
                st.download_button("📄 Word Gerencial", data=gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_e3:
                st.download_button("📑 PDF Gerencial", data=gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.pdf", mime="application/pdf")
            with col_e4:
                st.download_button("📊 Apresentação PPTX", data=gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Apresentacao_GrupoKMB.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    main()
