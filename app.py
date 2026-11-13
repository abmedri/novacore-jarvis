# ====================================================================================
# ⚡ GRUPO KMB // PROTOCOLO J.A.R.V.I.S. (VERSÃO OFICIAL COM EXPANSÃO AUTOMÁTICA DE LOCAIS)
# ====================================================================================
import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import docx
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.util import Inches as PptxInches

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
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
                    return df, "Processado com sucesso."
                except Exception:
                    continue
            return None, "Erro ao decodificar CSV."
        elif nome.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(arquivo_carregado)
            return df, "Processado com sucesso."
        else:
            return None, "Formato não suportado."
    except Exception as e:
        return None, str(e)


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
        
        col_val_aju = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df.columns else next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), None)
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
            "nome": nomes_arquivos[idx], "total_skus": total_skus, "qtd_antes": qtd_antes,
            "qtd_pos": qtd_pos, "soma_antes": soma_antes, "soma_depois": soma_depois,
            "val_ajuste_liquido": val_ajuste_liquido, "impacto_sobre_antes_pct": impacto_sobre_antes_pct,
            "skus_falta": skus_falta, "skus_sobra": skus_sobra, "skus_exatos": skus_exatos,
            "total_div": total_div, "perc_falta_itens": perc_falta_itens, "perc_sobra_itens": perc_sobra_itens,
            "perc_divergencia": perc_divergencia, "vol_falta": vol_falta, "vol_sobra": vol_sobra,
            "vol_div_total": vol_div_total, "perc_vol_div": perc_vol_div, "val_falta": val_falta,
            "val_sobra": val_sobra, "val_div_absoluto": val_div_absoluto, "acuracia_vol": acuracia_vol,
            "acuracia_fin": acuracia_fin, "df": df
        })
    return resumo_ficheiros


# ====================================================================================
# GERADOR DA PLANILHA DE CONFERÊNCIA COM EXPANSÃO AUTOMÁTICA DE LOCAIS
# ====================================================================================
def gerar_planilha_conferencia_kmb(df_localizacao, df_estoque, numero_loja_input=13016):
    wb = openpyxl.Workbook()
    
    fonte_padrao = Font(name="Arial", size=10)
    fonte_negrito = Font(name="Arial", size=10, bold=True)
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )
    
    # 1. Aba LOJA
    ws_loja = wb.active
    ws_loja.title = "LOJA"
    ws_loja.views.sheetView[0].showGridLines = True
    
    tabela_lojas = [
        [4684, 'MINAS', 'CATANDUVA'], [4685, 'MIRASSOL', 'MIRASSOL'], [4701, 'FARIA', 'RIO PRETO'],
        [4794, 'GARDEN', 'CATANDUVA'], [5036, 'GENERAL', 'RIO PRETO'], [5037, 'ATA CENTRO', 'ARAÇATUBA'],
        [5783, 'TANABI', 'TANABI'], [6119, 'PRAÇA', 'RIO PRETO'], [11021, 'ELDORADO', 'RIO PRETO'],
        [12524, 'FERNANDÓPOLIS QUIOSQUE', 'FERNANDÓPOLIS'], [12760, 'MUFFATO', 'RIO PRETO'], [12853, 'CARREFOUR', 'RIO PRETO'],
        [12901, 'PLAZA', 'RIO PRETO'], [13015, 'RONDON', 'ARAÇATUBA'], [13016, 'POTIRENDABA', 'POTIRENDABA'],
        [13017, 'NOVA GRANADA', 'NOVA GRANADA'], [13018, 'JALES', 'JALES'], [13019, 'FERNANDÓPOLIS', 'FERNANDÓPOLIS'],
        [13123, 'TABAPUÃ', 'TABAPUÃ'], [13372, 'VD RIO PRETO', 'RIO PRETO'], [13904, 'PROENÇA', 'CATANDUVA'],
        [13922, 'CIDADE NORTE', 'RIO PRETO'], [14252, 'AMIGÃO', 'ARAÇATUBA'], [14268, 'ATACADÃO', 'CATANDUVA'],
        [14378, 'GUAPIAÇU', 'GUAPIAÇU'], [14736, 'MUFFATO ARAÇATUBA', 'ARAÇATUBA'], [17315, 'IGUATEMI', 'RIO PRETO'],
        [17612, 'DAMHA', 'RIO PRETO'], [18591, 'PRAÇA NOVA', 'ARAÇATUBA'], [20385, 'CONTAINER VETORASSO', 'RIO PRETO'],
        [22373, 'SAUDADES', 'BIRIGUI'], [22374, 'GUARARAPES', 'GUARARAPES'], [22375, 'AMIGÃO MIRAGAIA', 'BIRIGUI'],
        [22401, 'ANDRADINA SHOPPING', 'ANDRADINA'], [22402, 'CASTILHO', 'CASTILHO'], [22403, 'BANDEIRANTE', 'BIRIGUI'],
        [22404, 'BARÃO', 'BIRIGUI'], [22405, 'ANDRADINA CENTRO', 'ANDRADINA'], [22406, 'MIRANDÓPOLIS', 'MIRANDÓPOLIS'],
        [22491, 'VD ARAÇATUBA', 'ARAÇATUBA'], [22511, 'VD CATANDUVA', 'CATANDUVA'], [22512, 'VD JALES', 'JALES'],
        [23015, 'VD BIRIGUI', 'BIRIGUI'], [23505, 'SANTA FÉ DO SUL', 'SANTA FÉ DO SUL'], [23514, 'PEREIRA BARRETO', 'PEREIRA BARRETO'],
        [23515, 'GENERAL SALGADO', 'GENERAL SALGADO'], [23518, 'ILHA SOLTEIRA', 'ILHA SOLTEIRA'], [23525, 'AURIFLAMA', 'AURIFLAMA'],
        [23526, "PALMEIRA D'OESTE", "PALMEIRA D'OESTE"], [23527, 'ILHA SOLTEIRA QUIOSQUE', 'ILHA SOLTEIRA'],
        [23528, 'VD FERNANDÓPOLIS', 'FERNANDÓPOLIS'], [23529, 'SANTA FÉ DO SUL QUIOSQUE', 'SANTA FÉ DO SUL'],
        [23530, 'PEREIRA BARRETO QUIOSQUE', 'PEREIRA BARRETO'], [24756, 'BADY BASSITT', 'BADY BASSITT']
    ]
    for r_loja in tabela_lojas:
        ws_loja.append(r_loja)

    # 2. Aba CompTar
    ws_comp = wb.create_sheet(title="CompTar")
    ws_comp.views.sheetView[0].showGridLines = True
    
    headers_comp = ["Local", "Codigo", "Descricao", "Qtd"]
    ws_comp.append(headers_comp)
    for col_num in range(1, 5):
        cell = ws_comp.cell(row=1, column=col_num)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    col_loc = next((c for c in df_localizacao.columns if 'local' in c.lower()), df_localizacao.columns[2])
    col_cod_loc = next((c for c in df_localizacao.columns if 'codigo' in c.lower() or 'cod' in c.lower()), df_localizacao.columns[3])
    col_desc_loc = next((c for c in df_localizacao.columns if 'desc' in c.lower()), df_localizacao.columns[4])
    col_qtd_loc = next((c for c in df_localizacao.columns if 'qtd' in c.lower() or 'quantidade' in c.lower()), df_localizacao.columns[5])
    
    locais_por_produto = {}
    max_locais_encontrados = 12
    
    for _, row in df_localizacao.iterrows():
        loc = str(row.get(col_loc, ''))
        cod = str(row.get(col_cod_loc, ''))
        desc = str(row.get(col_desc_loc, ''))
        qtd = row.get(col_qtd_loc, 0)
        
        ws_comp.append([loc, cod, desc, qtd])
        
        if cod not in locais_por_produto:
            locais_por_produto[cod] = []
        locais_por_produto[cod].append((loc, qtd))
        if len(locais_por_produto[cod]) > max_locais_encontrados:
            max_locais_encontrados = len(locais_por_produto[cod])
    
    for row in ws_comp.iter_rows(min_row=2, max_row=ws_comp.max_row, min_col=1, max_col=4):
        for cell in row:
            cell.font = fonte_padrao
            cell.border = borda_fina

    # 3. Aba Saldo Loja (Expansão automática se houver mais de 12 locais)
    ws_saldo = wb.create_sheet(title="Saldo Loja")
    ws_saldo.views.sheetView[0].showGridLines = True
    ws_saldo.page_setup.orientation = ws_saldo.ORIENTATION_LANDSCAPE
    ws_saldo.page_setup.paperSize = ws_saldo.PAPERSIZE_A4
    
    ws_saldo['D1'] = "DATA:"
    ws_saldo['D1'].font = fonte_negrito
    ws_saldo['E1'] = pd.Timestamp.now().strftime('%d/%m/%Y')
    ws_saldo['E1'].font = fonte_padrao
    
    ws_saldo['D2'] = "LOJA:"
    ws_saldo['D2'].font = fonte_negrito
    ws_saldo['E2'] = int(numero_loja_input)
    ws_saldo['E2'].font = fonte_padrao
    
    ws_saldo['G2'] = '=IF(_xlfn.XLOOKUP(E2,LOJA!A:A,LOJA!B:B,"")=0,"",_xlfn.XLOOKUP(E2,LOJA!A:A,LOJA!B:B,""))'
    ws_saldo['G2'].font = fonte_negrito
    
    ws_saldo['Q2'] = "CIDADE:"
    ws_saldo['Q2'].font = fonte_negrito
    ws_saldo['R2'] = '=IF(_xlfn.XLOOKUP(E2,LOJA!A:A,LOJA!C:C,"")=0,"",_xlfn.XLOOKUP(E2,LOJA!A:A,LOJA!C:C,""))'
    ws_saldo['R2'].font = fonte_padrao
    
    headers_saldo = ["Codigo", "Descrição", "Estoque", "COMPTAR", "DIVERGÊNCIA"]
    for i in range(1, max_locais_encontrados + 1):
        headers_saldo.append(str(i) if i > 1 else "CONT. 1")
    for i in range(1, max_locais_encontrados + 1):
        headers_saldo.append(f"Local {i}" if i == 1 else str(i))
        
    ws_saldo.append([])
    ws_saldo.append(headers_saldo)
    
    for col_num in range(1, len(headers_saldo) + 1):
        cell = ws_saldo.cell(row=3, column=col_num)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    col_cod_est = next((c for c in df_estoque.columns if 'codigo' in c.lower()), 'Codigo')
    col_desc_est = next((c for c in df_estoque.columns if 'desc' in c.lower()), 'Descrição')
    col_est = next((c for c in df_estoque.columns if 'estoque' in c.lower()), 'Estoque')
    
    for idx, row in df_estoque.iterrows():
        r = idx + 4
        cod = str(row.get(col_cod_est, ''))
        desc = row.get(col_desc_est, '')
        est = row.get(col_est, 0)
        
        formula_comptar = f"=SUMIF(CompTar!B:B, A{r}, CompTar!D:D)"
        formula_divergencia = f"=D{r}-C{r}"
        
        linha_dados = [cod, desc, est, formula_comptar, formula_divergencia]
        locs_prod = locais_por_produto.get(cod, [])
        
        contagens_vals = [locs_prod[i][1] if i < len(locs_prod) else None for i in range(max_locais_encontrados)]
        linha_dados.extend(contagens_vals)
        
        locais_vals = [locs_prod[i][0] if i < len(locs_prod) else None for i in range(max_locais_encontrados)]
        linha_dados.extend(locais_vals)
        
        ws_saldo.append(linha_dados)
        
    for row in ws_saldo.iter_rows(min_row=4, max_row=ws_saldo.max_row, min_col=1, max_col=len(headers_saldo)):
        for cell in row:
            cell.font = fonte_padrao
            cell.border = borda_fina
            if cell.column in [1, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif cell.column == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for ws in [ws_loja, ws_comp, ws_saldo]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    val_str = str(cell.value)
                    if not val_str.startswith("="):
                        max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Resumo Executivo"
    ws1['A1'] = "⚡ GRUPO KMB // PAINEL DE GOVERNANÇA E AUDITORIA PATRIMONIAL"
    ws1.append([])
    ws1.append(["Ficheiro / Unidade", "Total SKUs", "Estoque Antes (R$)", "Estoque Depois (R$)", "Ajuste Líquido (R$)", "Impacto (%)", "Faltas", "Sobras", "Acurácia Vol.", "Status"])
    for r in resumos:
        ws1.append([r['nome'], r['total_skus'], r['soma_antes'], r['soma_depois'], r['val_ajuste_liquido'], f"{r['impacto_sobre_antes_pct']:.2f}%", r['skus_falta'], r['skus_sobra'], f"{r['acuracia_vol']:.2f}%", "CONFORME" if r['acuracia_vol'] >= 97 else "CRÍTICO"])
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    doc = docx.Document()
    doc.add_heading("⚡ GRUPO KMB // RELATÓRIO EXECUTIVO", level=1)
    for r in resumos:
        doc.add_paragraph(f"• Unidade: {r['nome']} | SKUs: {r['total_skus']:,} | Acurácia: {r['acuracia_vol']:.2f}%")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)
    buffer = io.BytesIO()
    doc_pdf = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [Paragraph("⚡ GRUPO KMB // RELATÓRIO EXECUTIVO C-LEVEL", styles['Heading1']), HRFlowable(width="100%", thickness=1, color=colors.HexColor('#3B82F6'))]
    for r in resumos:
        story.append(Paragraph(f"<b>Unidade:</b> {r['nome']} | <b>Acurácia:</b> {r['acuracia_vol']:.2f}%", styles['Normal']))
    doc_pdf.build(story)
    buffer.seek(0)
    return buffer


def gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(PptxInches(1), PptxInches(1), PptxInches(11), PptxInches(2)).text_frame.text = "⚡ GRUPO KMB // C-LEVEL"
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
        st.markdown("<style>.stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', monospace; }</style>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 1.5, 1])
        with col2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            if os.path.exists("brasao_sem_fundo.png"):
                st.image("brasao_sem_fundo.png", use_container_width=True)
            st.markdown("## ⚡ TIME DE INVENTÁRIO // AUTENTICAÇÃO")
            senha_digitada = st.text_input("Senha de Acesso", type="password", key="input_senha")
            if st.button("🔓 Entrar no Sistema", use_container_width=True):
                if senha_digitada == SENHA_MESTRE:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error("❌ Senha incorreta.")
        return

    with st.sidebar:
        if os.path.exists("brasao_sem_fundo.png"):
            st.image("brasao_sem_fundo.png", use_container_width=True)
        st.markdown("### **GRUPO KMB**")
        st.markdown("*Exatidão • Controle • Eficiência*")
        st.markdown("---")
        modo_operacao = st.radio("Modo de Operação:", ["Auditoria de Inventário (5 Etapas)", "Modo Planilha de Conferência"])
        st.markdown("---")
        if st.button("🔒 Encerrar Sessão", use_container_width=True):
            st.session_state.autenticado = False
            st.rerun()

    st.markdown("<div class='main-header'><div style='font-size:1.7rem; font-weight:800; color:#3b82f6;'>⚡ GRUPO KMB // DIRETORIA EXECUTIVA</div></div>", unsafe_allow_html=True)
    
    if modo_operacao == "Modo Planilha de Conferência":
        st.markdown("### 📋 MODO PLANILHA DE CONFERÊNCIA // GERAÇÃO AUTOMATIZADA")
        numero_loja_ui = st.number_input("Número da Loja (ex: 13016 ou 17315)", min_value=1, max_value=99999, value=13016, step=1)
        
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            arq_loc = st.file_uploader("1. Arquivo de Itens por Localização", type=['csv', 'xlsx', 'xls'], key="up_loc")
        with col_up2:
            arq_est = st.file_uploader("2. Arquivo de Saldo de Estoque", type=['csv', 'xlsx', 'xls'], key="up_est")
            
        if arq_loc and arq_est:
            df_loc_temp, msg1 = processar_arquivo_upload(arq_loc)
            df_est_temp, msg2 = processar_arquivo_upload(arq_est)
            
            if df_loc_temp is not None and df_est_temp is not None:
                st.success("✅ Arquivos carregados com sucesso!")
                if st.button("⚡ Gerar Planilha Oficial de Conferência", use_container_width=True):
                    buffer_planilha = gerar_planilha_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui)
                    st.download_button(
                        label="📥 Baixar Planilha de Conferência KMB (.xlsx)",
                        data=buffer_planilha,
                        file_name=f"Plano_de_Inventario_Loja_{numero_loja_ui}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        return

    arquivos_submetidos = st.file_uploader("Carregue o arquivo oficial de inventário", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)
    if arquivos_submetidos:
        lista_dfs = [processar_arquivo_upload(arq)[0] for arq in arquivos_submetidos if processar_arquivo_upload(arq)[0] is not None]
        nomes_arquivos = [arq.name for arq in arquivos_submetidos]
        if lista_dfs:
            resumo_unico = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)[0]
            st.markdown(f"### 📌 SKUs Auditados: {resumo_unico['total_skus']:,} | Acurácia: {resumo_unico['acuracia_vol']:.2f}%")
            col_e1, col_e2, col_e3, col_e4 = st.columns(4)
            with col_e1: st.download_button("📊 Excel", data=gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio.xlsx")
            with col_e2: st.download_button("📄 Word", data=gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio.docx")
            with col_e3: st.download_button("📑 PDF", data=gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio.pdf")
            with col_e4: st.download_button("📊 PPTX", data=gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio.pptx")

if __name__ == "__main__":
    main()
