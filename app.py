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
        impacto_sobre_pos_pct = (val_ajuste_liquido / soma_depois * 100) if soma_depois > 0 else 0
        
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
            "impacto_sobre_pos_pct": impacto_sobre_pos_pct,
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


# ====================================================================================
# FUNÇÕES DE EXPORTAÇÃO ROBUSTAS (EXCEL, WORD, PDF E PPTX)
# ====================================================================================

def gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    wb = openpyxl.Workbook()
    
    # Aba 1: Indicadores e Resumo Executivo
    ws1 = wb.active
    ws1.title = "Resumo Executivo"
    ws1.views.sheetView[0].showGridLines = True
    
    ws1['A1'] = "⚡ NOVA-CORE // PAINEL DE GOVERNANÇA E AUDITORIA PATRIMONIAL"
    ws1['A1'].font = Font(size=14, bold=True, color="1F4E78")
    ws1.append([])
    
    headers_resumo = ["Ficheiro / Unidade", "Total SKUs", "Estoque Antes (R$)", "Estoque Depois (R$)", "Ajuste Líquido (R$)", "Impacto S/ Antes (%)", "Acurácia Vol. (%)", "Status"]
    ws1.append(headers_resumo)
    
    for r in resumos:
        status = "CONFORME" if r['acuracia_vol'] >= 95 else "CRÍTICO"
        ws1.append([r['nome'], r['total_skus'], r['soma_antes'], r['soma_depois'], r['val_ajuste_liquido'], f"{r['impacto_sobre_antes_pct']:.2f}%", f"{r['acuracia_vol']:.2f}%", status])
    
    # Aba 2: Detalhamento Completo (Faltas, Sobras, Zerados e Data)
    ws2 = wb.create_sheet(title="Divergências e Zerados")
    ws2.views.sheetView[0].showGridLines = True
    
    headers_detalhe = ["Unidade / Ficheiro", "Data do Inventário", "Código do Produto", "Descrição", "Saldo Sistema", "Ajuste", "Valor Ajuste (R$)", "Classificação de Divergência"]
    ws2.append(headers_detalhe)
    
    for idx, df in enumerate(lista_dfs):
        nome_arq = nomes_arquivos[idx]
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        col_val_aju = next((c for c in df.select_dtypes(include=['number']).columns if 'ajuste' in c.lower() and 'valor' in c.lower()), None)
        
        for _, row in df.iterrows():
            ajuste = row[col_ajuste] if col_ajuste and col_ajuste in df.columns else 0
            val_aju = row[col_val_aju] if col_val_aju and col_val_aju in df.columns else 0
            
            if ajuste < 0:
                classificacao = "FALTA (QUEBRA)"
            elif ajuste > 0:
                classificacao = "SOBRA"
            else:
                classificacao = "CONFORME / ZERO"
                
            data_inv = str(row.get('Data de Abertura', row.get('Data', 'N/A')))
            cod_prod = str(row.get('Código do Produto', row.get('Codigo', 'N/A')))
            desc_prod = str(row.get('Descrição do Produto', row.get('Descricao', 'Item Diversos')))
            saldo_sis = row.get('Saldo em Estoque', 0)
            
            ws2.append([nome_arq, data_inv, cod_prod, desc_prod, saldo_sis, ajuste, val_aju, classificacao])
            
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    doc = docx.Document()
    
    doc.add_heading("⚡ NOVA-CORE // RELATÓRIO EXECUTIVO DE AUDITORIA C-LEVEL", level=1)
    doc.add_paragraph("Documento oficial consolidado de governança patrimonial, contendo indicadores de impacto, análise de quebras, sobras e itens críticos.")
    
    doc.add_heading("1. Indicadores Consolidados por Unidade", level=2)
    for r in resumos:
        p = doc.add_paragraph()
        p.add_run(f"• Unidade / Arquivo: ").bold = True
        p.add_run(f"{r['nome']}\n")
        p.add_run(f"  - Total de SKUs Auditados: {r['total_skus']:,}\n")
        p.add_run(f"  - Valor Teórico Inicial: R$ {r['soma_antes']:,.2f}\n")
        p.add_run(f"  - Valor Auditado Final: R$ {r['soma_depois']:,.2f}\n")
        p.add_run(f"  - Ajuste Líquido: R$ {r['val_ajuste_liquido']:,.2f} (Impacto: {r['impacto_sobre_antes_pct']:.2f}%)\n")
        p.add_run(f"  - Acurácia Volumétrica: {r['acuracia_vol']:.2f}% | Acurácia Financeira: {r['acuracia_fin']:.2f}%\n")
    
    doc.add_heading("2. Detalhamento de Divergências Críticas (Faltas, Sobras e Zerados)", level=2)
    for idx, df in enumerate(lista_dfs):
        doc.add_paragraph(f"Análise detalhada da base: {nomes_arquivos[idx]}")
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        
        if col_ajuste and col_ajuste in df.columns:
            total_faltas = len(df[df[col_ajuste] < 0])
            total_sobras = len(df[df[col_ajuste] > 0])
            total_zerados = len(df[df[col_ajuste] == 0])
            doc.add_paragraph(f"  - Itens em Falta (Quebras): {total_faltas:,}\n  - Itens em Sobra: {total_sobras:,}\n  - Itens Conformes / Estoque Congelado Zero: {total_zerados:,}")
        else:
            doc.add_paragraph("Coluna de ajuste não mapeada diretamente para listagem individual detalhada neste bloco.")

    doc.add_heading("3. Validação e Assinatura C-Level", level=2)
    doc.add_paragraph("Declaro que os dados acima foram auditados eletronicamente através do Protocolo J.A.R.V.I.S. / NovaCore e refletem com fidelidade a apuração patrimonial realizada.\n\n\n")
    
    p_assinatura = doc.add_paragraph()
    p_assinatura.add_run("__________________________________________________\n").bold = True
    p_assinatura.add_run("Diretoria de Operações & Loss Prevention / Auditoria C-Level\n")
    p_assinatura.add_run("NovaCore Analytics & Governance System")

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
    
    story.append(Paragraph("⚡ NOVA-CORE // DOSSIÊ DE AUDITORIA PATRIMONIAL C-LEVEL", estilo_titulo))
    story.append(Paragraph("Relatório oficial consolidado contendo indicadores de impacto, análise de quebras, sobras e itens críticos.", estilo_texto))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#3B82F6'), spaceBefore=8, spaceAfter=12))
    
    story.append(Paragraph("1. Indicadores Consolidados por Unidade", estilo_sub))
    for r in resumos:
        texto_unidade = f"""
        <b>• Unidade / Arquivo:</b> {r['nome']}<br/>
        <b>- Total de SKUs Auditados:</b> {r['total_skus']:,}<br/>
        <b>- Valor Teórico Inicial (Antes):</b> R$ {r['soma_antes']:,.2f}<br/>
        <b>- Valor Auditado Final (Depois):</b> R$ {r['soma_depois']:,.2f}<br/>
        <b>- Ajuste Líquido:</b> R$ {r['val_ajuste_liquido']:,.2f} (Impacto: {r['impacto_sobre_antes_pct']:.2f}%)<br/>
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
            total_zerados = len(df[df[col_ajuste] == 0])
            
            texto_div = f"""
            <b>• Base: {nome_arq}</b><br/>
            - Itens em Falta (Quebras): {total_faltas:,}<br/>
            - Itens em Sobra: {total_sobras:,}<br/>
            - Itens Conformes / Estoque Congelado Zero: {total_zerados:,}
            """
            story.append(Paragraph(texto_div, estilo_texto))
        else:
            story.append(Paragraph(f"• Base: {nome_arq} - Coluna de ajuste não mapeada para contagem detalhada.", estilo_texto))
            
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("3. Validação e Assinatura C-Level", estilo_sub))
    story.append(Paragraph("Declaro que os dados acima foram auditados eletronicamente através do Protocolo J.A.R.V.I.S. / NovaCore e refletem com fidelidade a apuração patrimonial realizada.", estilo_texto))
    story.append(Spacer(1, 20))
    
    texto_assinatura = """
    <b>__________________________________________________</b><br/>
    <b>Diretoria de Operações & Loss Prevention / Auditoria C-Level</b><br/>
    NovaCore Analytics & Governance System
    """
    story.append(Paragraph(texto_assinatura, estilo_texto))
    
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
                st.markdown("### 🔍 Visão de Auditoria Pontual (Inventário Único)")
                df_filtrado = lista_dfs[0]
                
                # --- ETAPA 5 (MODO 1): Inspeção Analítica Master & Central de Exportação C-Level ---
                with st.expander("📂 ETAPA 5: Inspeção Analítica & Central de Exportação Executiva", expanded=True):
                    st.markdown("""
                    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(59, 130, 246, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                        <span style="color: #3b82f6; font-weight: 600; font-size: 0.9rem;">CONSOLIDAÇÃO & RELATÓRIOS MULTIFORMATO</span>
                        <p style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px; margin-bottom: 0;">
                            Inspeção linha a linha da base de inventário auditada e exportação instantânea dos relatórios executivos para a diretoria.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)

                    st.dataframe(df_filtrado, use_container_width=True, height=350)

                    st.markdown("---")
                    st.markdown("### 📥 Central de Download de Relatórios C-Level")

                    col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)

                    with col_dl1:
                        excel_buffer = gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos)
                        st.download_button(
                            label="📥 Baixar Excel Executivo",
                            data=excel_buffer,
                            file_name="Relatorio_Gerencial_Inventario.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

                    with col_dl2:
                        docx_buffer = gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos)
                        st.download_button(
                            label="📄 Baixar Relatório Word",
                            data=docx_buffer,
                            file_name="Relatorio_Gerencial_Inventario.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )

                    with col_dl3:
                        pdf_buffer = gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos)
                        st.download_button(
                            label="📊 Baixar Relatório PDF",
                            data=pdf_buffer,
                            file_name="Relatorio_Gerencial_Inventario.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

                    with col_dl4:
                        pptx_buffer = gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos)
                        st.download_button(
                            label="📑 Baixar Apresentação PPT",
                            data=pptx_buffer,
                            file_name="Apresentacao_Executiva_Inventario.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )

            # ====================================================================
            # MODO 2: MÚLTIPLOS FICHEIROS (VISÃO COMPARATIVA & MULTITEMPORAL)
            # ====================================================================
            else:
                st.markdown("### 📊 Visão Executiva Comparativa (Múltiplos Inventários)")
                
                num_cols = lista_dfs[0].select_dtypes(include=['number']).columns.tolist()

                # --------------------------------------------------------------------
                # ETAPA 5 (MODO 2): Simulador Contábil & Evolução Histórica (Posicionado antes dos relatórios)
                # --------------------------------------------------------------------
                with st.expander("📈 ETAPA 5: Simulador Contábil & Evolução Histórica (Multitemporal)", expanded=False):
                    st.markdown("""
                    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(16, 185, 129, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 16px;">
                        <span style="color: #10b981; font-weight: 600; font-size: 0.9rem;">ANÁLISE DE TENDÊNCIA & EVOLUÇÃO PATRIMONIAL</span>
                        <p style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px; margin-bottom: 0;">
                            Comparação temporal cruzando os múltiplos inventários carregados para rastrear o comportamento histórico da acurácia e desvios.
                        </p>
                    </div>
                    """, unsafe_allow_html=True)

                    if len(lista_dfs) >= 2 and len(num_cols) >= 1:
                        col_evolu_y = num_cols[0]
                        
                        df_historico_lista = []
                        for nome_arq, df_temp in zip(nomes_arquivos, lista_dfs):
                            df_copia = df_temp.copy()
                            df_copia['Periodo_Inventario'] = nome_arq
                            df_historico_lista.append(df_copia)
                        
                        df_multitemporal = pd.concat(df_historico_lista, ignore_index=True)

                        fig_evolucao = px.box(
                            df_multitemporal,
                            x='Periodo_Inventario',
                            y=col_evolu_y,
                            title=f"Evolução Histórica da Métrica: {col_evolu_y}",
                            template="plotly_dark",
                            color='Periodo_Inventario'
                        )

                        fig_evolucao.update_layout(
                            paper_bgcolor='rgba(11, 15, 25, 0.9)',
                            plot_bgcolor='rgba(11, 15, 25, 0.9)',
                            height=450,
                            margin=dict(l=30, r=30, t=40, b=30),
                            showlegend=False
                        )

                        st.plotly_chart(fig_evolucao, use_container_width=True)
                        st.success("✅ **Análise Temporal Ativa:** Comparando o comportamento e a dispersão dos valores entre os inventários históricos carregados.")
                    else:
                        st.warning("⚠️ O Modo 2 requer pelo menos 2 arquivos carregados para gerar a linha do tempo e o gráfico de evolução histórica comparativa.")

                # --------------------------------------------------------------------
                # CENTRAL DE RELATÓRIOS GERENCIAIS / EXPORTAÇÃO (Fim do Modo 2)
                # --------------------------------------------------------------------
                st.markdown("---")
                st.markdown("### 📥 Central de Download de Relatórios Consolidados (C-Level)")

                col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)

                with col_dl1:
                    excel_buffer = gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos)
                    st.download_button("📥 Excel Executivo", excel_buffer, "Relatorio_Consolidado_Inventarios.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

                with col_dl2:
                    docx_buffer = gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos)
                    st.download_button("📄 Relatório Word", docx_buffer, "Relatorio_Consolidado_Inventarios.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

                with col_dl3:
                    pdf_buffer = gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos)
                    st.download_button("📊 Relatório PDF", pdf_buffer, "Relatorio_Consolidado_Inventarios.pdf", mime="application/pdf", use_container_width=True)

                with col_dl4:
                    pptx_buffer = gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos)
                    st.download_button("📑 Apresentação PPT", pptx_buffer, "Apresentacao_Consolidada_Inventarios.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)

if __name__ == "__main__":
    main()
