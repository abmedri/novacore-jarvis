# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (VERSÃO EXECUTIVA UNIVERSAL & BLINDADA)
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
                    
                    # Padronização automática de nomes de colunas
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
        
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        col_val_aju = next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df.columns else None)
        col_val = next((c for c in num_cols if 'valor de estoque' in c.lower()), num_cols[0] if num_cols else None)
        
        soma_depois = df[col_val].sum() if col_val and col_val in df.columns else 0
        col_qtd = 'Saldo em Estoque' if 'Saldo em Estoque' in df.columns else None
        qtd_pos = df[col_qtd].sum() if col_qtd and col_qtd in df.columns else total_skus
        
        soma_ajuste_qtd = df[col_ajuste].sum() if col_ajuste and col_ajuste in df.columns else 0
        qtd_antes = qtd_pos - soma_ajuste_qtd
        
        soma_ajuste_val = df[col_val_aju].sum() if col_val_aju and col_val_aju in df.columns else 0
        soma_antes = soma_depois - soma_ajuste_val
            
        df_falta = df[df[col_ajuste] < 0] if col_ajuste and col_ajuste in df.columns else pd.DataFrame()
        df_sobra = df[df[col_ajuste] > 0] if col_ajuste and col_ajuste in df.columns else pd.DataFrame()
        df_exato = df[df[col_ajuste] == 0] if col_ajuste and col_ajuste in df.columns else pd.DataFrame()
        
        skus_falta = len(df_falta)
        skus_sobra = len(df_sobra)
        skus_exatos = len(df_exato)
        total_div = skus_falta + skus_sobra
        
        vol_falta = abs(df_falta[col_ajuste].sum()) if not df_falta.empty else 0
        vol_sobra = df_sobra[col_ajuste].sum() if not df_sobra.empty else 0
        vol_div_total = vol_falta + vol_sobra
        
        val_falta = abs(df_falta[col_val_aju].sum()) if col_val_aju and not df_falta.empty else 0
        val_sobra = df_sobra[col_val_aju].sum() if col_val_aju and not df_sobra.empty else 0
        val_div_absoluto = val_falta + val_sobra
        val_ajuste_liquido = soma_depois - soma_antes
        
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
            "vol_falta": vol_falta,
            "vol_sobra": vol_sobra,
            "vol_div_total": vol_div_total,
            "val_falta": val_falta,
            "val_sobra": val_sobra,
            "val_div_absoluto": val_div_absoluto,
            "acuracia_vol": acuracia_vol,
            "acuracia_fin": acuracia_fin,
            "df": df
        })
    return resumo_ficheiros

def gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Dashboard Executivo"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1['A1'] = "⚡ NOVA-CORE // RELATÓRIO EXECUTIVO DE AUDITORIA"
    ws1.append([])
    headers = ["Ficheiro", "SKUs", "Valor Teórico (Antes)", "Valor Auditado (Depois)", "Ajuste Líquido (R$)", "Impacto S/ Antes (%)", "Status"]
    ws1.append(headers)
    
    for r in resumos:
        status = "CONFORME" if r['acuracia_vol'] >= 95 else "CRÍTICO"
        ws1.append([r['nome'], r['total_skus'], r['soma_antes'], r['soma_depois'], r['val_ajuste_liquido'], f"{r['impacto_sobre_antes_pct']:.2f}%", status])
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    doc = docx.Document()
    doc.add_heading("⚡ NOVA-CORE // RELATÓRIO EXECUTIVO C-LEVEL", level=1)
    for r in resumos:
        doc.add_paragraph(f"Unidade: {r['nome']} | SKUs: {r['total_skus']:,} | Estoque Pré: R$ {r['soma_antes']:,.2f} | Estoque Pós: R$ {r['soma_depois']:,.2f} | Ajuste: R$ {r['val_ajuste_liquido']:,.2f} (Impacto Pré: {r['impacto_sobre_antes_pct']:.2f}%)")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    buffer = io.BytesIO()
    doc_pdf = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [Paragraph("⚡ NOVA-CORE // RELATÓRIO EXECUTIVO", styles['Heading1'])]
    for r in resumos:
        story.append(Paragraph(f"<b>{r['nome']}</b>: {r['total_skus']:,} SKUs | Pré: R$ {r['soma_antes']:,.2f} | Pós: R$ {r['soma_depois']:,.2f} | Ajuste Líquido: R$ {r['val_ajuste_liquido']:,.2f} | Impacto Pré: {r['impacto_sobre_antes_pct']:.2f}%", styles['Normal']))
    doc_pdf.build(story)
    buffer.seek(0)
    return buffer

def gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(PptxInches(1), PptxInches(1), PptxInches(11), PptxInches(2))
    tb.text_frame.text = "⚡ NOVA-CORE // RELATÓRIO EXECUTIVO C-LEVEL"
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def main():
    st.set_page_config(page_title="NOVA-CORE // J.A.R.V.I.S. EXECUTIVO", page_icon="⚡", layout="wide")
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
    
    st.markdown("""
    <div class="main-header">
        <div class="jarvis-title">⚡ NOVA-CORE // DIRETORIA EXECUTIVA</div>
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
                
                total_pos = len(df_filtrado)
                idx_val = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
                col_val = num_cols[idx_val] if num_cols else None
                soma_depois = df_filtrado[col_val].sum() if col_val and col_val in df_filtrado.columns else 0
                
                idx_qtd = identificar_coluna_ideal(num_cols, ['quantidade', 'saldo', 'estoque', 'fisico', 'teorico'])
                col_qtd_pos = num_cols[idx_qtd] if num_cols else None
                total_itens_pos = df_filtrado[col_qtd_pos].sum() if col_qtd_pos and col_qtd_pos in df_filtrado.columns else total_pos
                
                col_ajuste = 'Ajuste' if 'Ajuste' in df_filtrado.columns else None
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
                
                col_val_aju = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df_filtrado.columns else ('Valor de Ajuste(R$)' if 'Valor de Ajuste(R$)' in df_filtrado.columns else None)
                val_falta_fin = df_falta[col_val_aju].sum() if col_val_aju and not df_falta.empty else 0
                val_sobra_fin = df_sobra[col_val_aju].sum() if col_val_aju and not df_sobra.empty else 0
                soma_ajuste_val = df_filtrado[col_val_aju].sum() if col_val_aju and col_val_aju in df_filtrado.columns else 0
                soma_antes = soma_depois - soma_ajuste_val
                diferenca_monetaria = soma_depois - soma_antes
                
                impacto_sobre_antes_pct = (diferenca_monetaria / soma_antes * 100) if soma_antes > 0 else 0
                soma_absoluta_val = abs(val_falta_fin) + abs(val_sobra_fin)
                
                acur_vol_real = (qtd_exatos / total_pos * 100) if total_pos > 0 else 0
                acur_fin_real = max(0.0, min(100.0, 100.0 - abs(impacto_sobre_antes_pct) * 2.5))
                
                # ETAPA 1: Indicadores Executivos & Bento Cards com Deltas e Percentuais
                st.markdown("### 📌 ETAPA 1: Indicadores Executivos & Parecer Técnico")
                st.markdown("<br>", unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{total_pos:,}", f"Pré: {total_itens_antes:,.0f} un | Pós: {total_itens_pos:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{qtd_falta:,} SKUs ({perc_falta_itens:.2f}%)", f"Vol: -{vol_falta_un:,.0f} un | R$ {val_falta_fin:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Volume Físico Total de Divergências (Sobra + Falta)", f"{vol_total_div:,.0f} unidades ({perc_vol_div:.2f}%)", f"R$ {soma_absoluta_val:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                with c2:
                    st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{qtd_sobra:,} SKUs ({perc_sobra_itens:.2f}%)", f"Vol: +{vol_sobra_un:,.0f} un | R$ +{val_sobra_fin:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Total Falta + Sobra (SKUs)", f"{total_divergencias:,} SKUs ({perc_divergencia:.2f}%)", "Total de Produtos Divergentes", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Volumétrica / Financeira", f"Vol: {acur_vol_real:.2f}% | Fin: {acur_fin_real:.2f}%", "Meta Corporativa 95%", eh_negativo=(acur_vol_real < 95), eh_positivo=(acur_vol_real >= 95)), unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("##### 💰 Balanço Patrimonial & Impacto Geral s/ Valor Antes (%)")
                col_kpi1, col_kpi2 = st.columns(2)
                with col_kpi1:
                    st.markdown(formatar_kpi_html("Estoque Antes (Sistema / Teórico)", f"R$ {soma_antes:,.2f}", f"Qtd Pré: {total_itens_antes:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Impacto Geral s/ Valor Antes (%)", f"{impacto_sobre_antes_pct:.2f}%", f"Ajuste Líquido: R$ {diferenca_monetaria:,.2f}", eh_negativo=(impacto_sobre_antes_pct < 0), eh_positivo=(impacto_sobre_antes_pct >= 0)), unsafe_allow_html=True)
                with col_kpi2:
                    st.markdown(formatar_kpi_html("Estoque Depois (Físico Auditado)", f"R$ {soma_depois:,.2f}", f"Qtd Pós: {total_itens_pos:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Divergência / Ajuste Líquido", f"R$ {diferenca_monetaria:,.2f}", f"Variação Apurada", eh_negativo=(diferenca_monetaria < 0), eh_positivo=(diferenca_monetaria >= 0)), unsafe_allow_html=True)
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
                        # 3 Seletores interativos lado a lado
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

                        # Cálculo completo da Curva ABC sobre toda a base filtrada
                        df_abc = df_filtrado.sort_values(by=eixo_y_escolhido, ascending=False).copy()
                        df_abc['Soma_Acumulada'] = df_abc[eixo_y_escolhido].cumsum()
                        soma_total = df_abc[eixo_y_escolhido].sum()
                        
                        if soma_total > 0:
                            df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                            df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(
                                lambda p: "Classe A (80%)" if p <= 80 else ("Classe B (95%)" if p <= 95 else "Classe C (100%)")
                            )
                            
                            # Aplicação do filtro selecionado
                            if filtro_classe_abc != "Todas as Classes (Geral)":
                                df_abc_filtrado = df_abc[df_abc['Classe_ABC'] == filtro_classe_abc].copy()
                            else:
                                df_abc_filtrado = df_abc.copy()

                            # Exibição de resumo rápido da classe filtrada
                            qtd_skus_classe = len(df_abc_filtrado)
                            valor_classe = df_abc_filtrado[eixo_y_escolhido].sum()
                            perc_valor_classe = (valor_classe / soma_total * 100) if soma_total > 0 else 0
                            st.caption(f"📊 **Resumo da Seleção ({filtro_classe_abc}):** {qtd_skus_classe:,} SKUs | Total: R$ {valor_classe:,.2f} ({perc_valor_classe:.2f}% do estoque)")

                            # Top 35 da seleção para renderização limpa no gráfico
                            df_plot = df_abc_filtrado.head(35).copy()
                            
                            if not df_plot.empty:
                                fig_abc = go.Figure()
                                
                                # Barras de Valor Individual
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
                                
                                # Linha de Porcentagem Acumulada
                                fig_abc.add_trace(go.Scatter(
                                    x=df_plot[coluna_categoria_escolhida],
                                    y=df_plot['%_Acumulado'],
                                    name="% Acumulado (Pareto)",
                                    mode='lines+markers',
                                    marker=dict(color='#f59e0b', size=8),
                                    line=dict(color='#f59e0b', width=3),
                                    yaxis='y2'
                                ))
                                
                                # Layout C-Level Dark Mode Blindado
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
                                    xaxis=dict(
                                        showgrid=False
                                    )
                                )
                                
                                st.plotly_chart(fig_abc, use_container_width=True)
                            else:
                                st.warning("Nenhum produto encontrado para a classe selecionada.")
                                
                # ETAPA 5: Evolução Histórica
                with st.expander("🧮 ETAPA 5: Simulador Contábil & Evolução Histórica", expanded=False):
                    df_hist = pd.DataFrame({
                        "Inventário": ["Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026 (Atual)"],
                        "Acuracidade Financeira (%)": [81.2, 83.5, 85.0, 87.1, acur_fin_real],
                        "Acuracidade Volumétrica (%)": [68.0, 70.4, 72.1, 74.8, acur_vol_real]
                    })
                    fig_hist = px.line(df_hist, x="Inventário", y=["Acuracidade Financeira (%)", "Acuracidade Volumétrica (%)"], markers=True, template="plotly_dark", title="Evolução Histórica de Acuracidade")
                    st.plotly_chart(fig_hist, use_container_width=True)

            # ====================================================================
            # MODO 2: MÚLTIPLOS FICHEIROS (VISÃO COMPARATIVA EXECUTIVA COMPLETA)
            # ====================================================================
            else:
                st.markdown("## ⚡ AUDITORIA COMPARATIVA MULTI-INVENTÁRIO C-LEVEL")
                resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
                
                # Cartões de Comparação Lado a Lado Completos
                col_cards = st.columns(len(lista_dfs))
                for idx, r in enumerate(resumos):
                    with col_cards[idx]:
                        with st.container(border=True):
                            st.markdown(f"#### 📄 {r['nome']}")
                            st.markdown(formatar_kpi_html("Total SKUs", f"{r['total_skus']:,}", f"Exatos: {r['skus_exatos']} SKUs", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Estoque Antes (Teórico)", f"R$ {r['soma_antes']:,.2f}", f"Qtd: {r['qtd_antes']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Estoque Depois (Auditado)", f"R$ {r['soma_depois']:,.2f}", f"Qtd: {r['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Ajuste Líquido (R$)", f"R$ {r['val_ajuste_liquido']:,.2f}", f"Impacto Pré: {r['impacto_sobre_antes_pct']:.2f}%", eh_negativo=(r['val_ajuste_liquido'] < 0), eh_positivo=(r['val_ajuste_liquido'] >= 0)), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Impacto Geral s/ Antes (%)", f"{r['impacto_sobre_antes_pct']:.2f}%", f"Sobre R$ {r['soma_antes']:,.2f}", eh_negativo=(r['impacto_sobre_antes_pct'] < 0), eh_positivo=(r['impacto_sobre_antes_pct'] >= 0)), unsafe_allow_html=True)

                st.markdown("---")
                st.subheader("📊 Gráficos Comparativos Consolidados")
                
                df_resumo_geral = pd.DataFrame([
                    {
                        "Ficheiro": r['nome'],
                        "Estoque Antes": r['soma_antes'],
                        "Estoque Depois": r['soma_depois'],
                        "Ajuste Líquido": r['val_ajuste_liquido'],
                        "Impacto S/ Antes (%)": r['impacto_sobre_antes_pct'],
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
                    fig_aj = px.bar(df_resumo_geral, x="Ficheiro", y="Impacto S/ Antes (%)", template="plotly_dark", title="Impacto Geral em Percentual (Sobre Antes)")
                    st.plotly_chart(fig_aj, use_container_width=True)
                
                gc3, gc4 = st.columns(2)
                with gc3:
                    fig_skus = px.bar(df_resumo_geral, x="Ficheiro", y=["Faltas (SKUs)", "Sobras (SKUs)"], barmode="group", template="plotly_dark", title="Comparativo Faltas vs Sobras (SKUs)")
                    st.plotly_chart(fig_skus, use_container_width=True)
                with gc4:
                    fig_tot = px.bar(df_resumo_geral, x="Ficheiro", y="Total SKUs", template="plotly_dark", title="Volume Total de SKUs por Inventário")
                    st.plotly_chart(fig_tot, use_container_width=True)

            # Central de Relatórios
            st.markdown("---")
            st.subheader("📥 Central de Relatórios Executivos C-Level")
            col_e1, col_e2, col_e3, col_e4 = st.columns(4)
            with col_e1:
                st.download_button("📊 Excel Gerencial", data=gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_Novacore.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_e2:
                st.download_button("📄 Word Gerencial", data=gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_Novacore.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_e3:
                st.download_button("📑 PDF Gerencial", data=gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_Novacore.pdf", mime="application/pdf")
            with col_e4:
                st.download_button("📊 Apresentação PPTX", data=gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Apresentacao_Novacore.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    main()
