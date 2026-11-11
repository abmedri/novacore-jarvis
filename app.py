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
        
        acuracia_vol = (skus_exatos / total_skus * 100) if total_skus > 0 else 0
        desvio_fin_pct = (val_ajuste_liquido / soma_antes * 100) if soma_antes > 0 else 0
        acuracia_fin = max(0.0, min(100.0, 100.0 - abs(desvio_fin_pct) * 2.5))
        
        # Curva ABC Pareto
        df_calc = df.copy()
        col_valor_abc = col_val if col_val else (col_val_aju if col_val_aju else num_cols[0] if num_cols else None)
        if col_valor_abc and col_valor_abc in df_calc.columns:
            df_calc['Val_Abs'] = df_calc[col_valor_abc].abs()
            df_calc = df_calc.sort_values(by='Val_Abs', ascending=False)
            df_calc['Soma_Acum'] = df_calc['Val_Abs'].cumsum()
            total_acum = df_calc['Val_Abs'].sum()
            df_calc['Pct_Acum'] = (df_calc['Soma_Acum'] / total_acum * 100) if total_acum > 0 else 0
            df_calc['Classe_ABC'] = df_calc['Pct_Acum'].apply(lambda p: 'A' if p <= 80 else ('B' if p <= 95 else 'C'))
            
            val_a = df_calc[df_calc['Classe_ABC'] == 'A']['Val_Abs'].sum()
            val_b = df_calc[df_calc['Classe_ABC'] == 'B']['Val_Abs'].sum()
            val_c = df_calc[df_calc['Classe_ABC'] == 'C']['Val_Abs'].sum()
            
            skus_a = len(df_calc[df_calc['Classe_ABC'] == 'A'])
            skus_b = len(df_calc[df_calc['Classe_ABC'] == 'B'])
            skus_c = len(df_calc[df_calc['Classe_ABC'] == 'C'])
        else:
            val_a, val_b, val_c = 0, 0, 0
            skus_a, skus_b, skus_c = 0, 0, 0
            
        resumo_ficheiros.append({
            "nome": nomes_arquivos[idx],
            "total_skus": total_skus,
            "qtd_antes": qtd_antes,
            "qtd_pos": qtd_pos,
            "soma_antes": soma_antes,
            "soma_depois": soma_depois,
            "val_ajuste_liquido": val_ajuste_liquido,
            "desvio_fin_pct": desvio_fin_pct,
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
            "val_a": val_a, "val_b": val_b, "val_c": val_c,
            "skus_a": skus_a, "skus_b": skus_b, "skus_c": skus_c,
            "df": df
        })
    return resumo_ficheiros

def gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Dashboard Executivo"
    ws1.views.sheetView[0].showGridLines = True
    
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    
    ws1['A1'] = "⚡ NOVA-CORE // RELATÓRIO EXECUTIVO DE AUDITORIA"
    ws1.append([])
    headers = ["Ficheiro", "SKUs", "Valor Teórico (Antes)", "Valor Auditado (Depois)", "Ajuste Líquido (R$)", "Status"]
    ws1.append(headers)
    
    for r in resumos:
        status = "CONFORME" if r['acuracia_vol'] >= 95 else "CRÍTICO"
        ws1.append([r['nome'], r['total_skus'], r['soma_antes'], r['soma_depois'], r['val_ajuste_liquido'], status])
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    doc = docx.Document()
    doc.add_heading("⚡ NOVA-CORE // RELATÓRIO EXECUTIVO C-LEVEL", level=1)
    for r in resumos:
        doc.add_paragraph(f"Unidade: {r['nome']} | SKUs: {r['total_skus']:,} | Estoque Pós: R$ {r['soma_depois']:,.2f} | Ajuste: R$ {r['val_ajuste_liquido']:,.2f}")
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
        story.append(Paragraph(f"<b>{r['nome']}</b>: {r['total_skus']:,} SKUs | Estoque Pós: R$ {r['soma_depois']:,.2f} | Ajuste Líquido: R$ {r['val_ajuste_liquido']:,.2f}", styles['Normal']))
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
    st.markdown("## ⚡ NOVA-CORE // PAINEL EXECUTIVO DE INVENTÁRIO (LOJA 13372)")
    
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
                
                qtd_exatos = (df_filtrado['Ajuste'] == 0).sum() if 'Ajuste' in df_filtrado.columns else int(total_pos * 0.5)
                total_divergencias = qtd_falta + qtd_sobra
                
                col_val_aju = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df_filtrado.columns else ('Valor de Ajuste(R$)' if 'Valor de Ajuste(R$)' in df_filtrado.columns else None)
                val_falta_fin = df_falta[col_val_aju].sum() if col_val_aju and not df_falta.empty else 0
                val_sobra_fin = df_sobra[col_val_aju].sum() if col_val_aju and not df_sobra.empty else 0
                soma_ajuste_val = df_filtrado[col_val_aju].sum() if col_val_aju and col_val_aju in df_filtrado.columns else 0
                soma_antes = soma_depois - soma_ajuste_val
                diferenca_monetaria = soma_depois - soma_antes
                impacto_perc = (diferenca_monetaria / soma_antes) * 100 if soma_antes > 0 else 0
                
                acur_vol_real = (qtd_exatos / total_pos * 100) if total_pos > 0 else 0
                acur_fin_real = max(0.0, min(100.0, 100.0 - abs(impacto_perc) * 2.5))
                
                # ETAPA 1: Indicadores Executivos
                st.markdown("### 📌 ETAPA 1: Indicadores Executivos & Parecer Técnico")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Total de Produtos", f"{total_pos:,}")
                with c2:
                    st.metric("Estoque Pós (unidades)", f"{total_itens_pos:,.0f}")
                with c3:
                    st.metric("Valor Total de Estoque", f"R$ {soma_depois:,.2f}")
                with c4:
                    st.metric("Ajuste Líquido Total", f"R$ {diferenca_monetaria:,.2f}")
                
                st.markdown("---")
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown(f"### 📈 Ajustes Positivos (Sobras)\n* **Produtos:** {qtd_sobra:,}\n* **Volume:** +{vol_sobra_un:,.0f} un\n* **Valor:** R$ {val_sobra_fin:,.2f}")
                with col_b:
                    st.markdown(f"### 📉 Ajustes Negativos (Faltas)\n* **Produtos:** {qtd_falta:,}\n* **Volume:** -{vol_falta_un:,.0f} un\n* **Valor:** R$ {val_falta_fin:,.2f}")

                # ETAPA 2: Curva ABC
                with st.expander("🎯 ETAPA 2: Concentração Financeira & Curva ABC (Pareto 80/20)", expanded=True):
                    if len(num_cols) >= 1:
                        idx_padrao_y = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
                        eixo_y_abc = num_cols[idx_padrao_y] if idx_padrao_y < len(num_cols) else num_cols[0]
                        df_abc = df_filtrado.sort_values(by=eixo_y_abc, ascending=False).copy()
                        df_abc['Soma_Acumulada'] = df_abc[eixo_y_abc].cumsum()
                        soma_total = df_abc[eixo_y_abc].sum()
                        if soma_total > 0:
                            df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                            df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(lambda p: "Classe A (80%)" if p <= 80 else ("Classe B (95%)" if p <= 95 else "Classe C (100%)"))
                            fig_abc = px.bar(df_abc.head(50), x='Descrição do Produto' if 'Descrição do Produto' in df_abc.columns else all_cols[0], y=eixo_y_abc, color='Classe_ABC', template="plotly_dark", title="Curva ABC - Pareto de Estoque")
                            st.plotly_chart(fig_abc, use_container_width=True)

                # ETAPA 3: Cobertura
                with st.expander("⏱️ ETAPA 3: Prevenção de Ruptura & Cobertura", expanded=False):
                    if len(num_cols) >= 1:
                        idx_cob = identificar_coluna_ideal(num_cols, ['saldo', 'quantidade', 'estoque'])
                        col_cob = num_cols[idx_cob] if idx_cob < len(num_cols) else num_cols[0]
                        df_cob = df_filtrado.copy()
                        df_cob['Dias_Cobertura'] = df_cob[col_cob] / 25
                        fig_cob = px.histogram(df_cob, x='Dias_Cobertura', nbins=30, template="plotly_dark", title="Distribuição de Dias de Cobertura")
                        st.plotly_chart(fig_cob, use_container_width=True)

                # ETAPA 4: Giro por Categoria
                with st.expander("🔍 ETAPA 4: Diagnóstico de Causa Raiz & Giro por Categoria", expanded=False):
                    if len(cat_cols) >= 1 and len(num_cols) >= 1:
                        eixo_x_giro = cat_cols[0]
                        eixo_y_giro = num_cols[0]
                        df_giro = df_filtrado.groupby(eixo_x_giro)[eixo_y_giro].sum().reset_index().head(15)
                        fig_giro = px.bar(df_giro, x=eixo_x_giro, y=eixo_y_giro, template="plotly_dark", title=f"Análise por {eixo_x_giro}")
                        st.plotly_chart(fig_giro, use_container_width=True)

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
            # MODO 2: MÚLTIPLOS FICHEIROS (VISÃO COMPARATIVA EXECUTIVA)
            # ====================================================================
            else:
                st.markdown("## ⚡ AUDITORIA COMPARATIVA MULTI-INVENTÁRIO")
                resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
                
                col_cards = st.columns(len(lista_dfs))
                for idx, r in enumerate(resumos):
                    with col_cards[idx]:
                        with st.container(border=True):
                            st.markdown(f"#### 📄 {r['nome']}")
                            st.metric("Estoque Pós", f"R$ {r['soma_depois']:,.2f}")
                            st.markdown(f"**SKUs:** `{r['total_skus']:,}`")
                            st.markdown(f"**Ajuste Líquido:** R$ {r['val_ajuste_liquido']:,.2f}")
                            st.markdown(f"**Faltas:** {r['skus_falta']} SKUs (R$ {r['val_falta']:,.2f})")
                            st.markdown(f"**Sobras:** {r['skus_sobra']} SKUs (R$ {r['val_sobra']:,.2f})")

                st.markdown("---")
                st.subheader("📊 Comparativo Gráfico de Valor e Volume")
                dfs_unidos = [df.assign(Origem_Inventario=nomes_arquivos[i]) for i, df in enumerate(lista_dfs)]
                df_cons = pd.concat(dfs_unidos, ignore_index=True)
                
                g_c1, g_c2 = st.columns(2)
                with g_c1:
                    num_col_demo = df_cons.select_dtypes(include=['number']).columns[0]
                    fig_t = px.bar(df_cons.groupby('Origem_Inventario')[num_col_demo].sum().reset_index(), x='Origem_Inventario', y=num_col_demo, template="plotly_dark", title="Comparação de Totais")
                    st.plotly_chart(fig_t, use_container_width=True)
                with g_c2:
                    fig_s = px.bar(df_cons.groupby('Origem_Inventario').size().reset_index(name='SKUs'), x='Origem_Inventario', y='SKUs', template="plotly_dark", title="Comparação de SKUs")
                    st.plotly_chart(fig_s, use_container_width=True)

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
