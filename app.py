# ====================================================================================
# ⚡ GRUPO KMB // PROTOCOLO J.A.R.V.I.S. (VERSÃO OFICIAL UNIFICADA COMPLETA - CORRIGIDA)
# ====================================================================================
import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
# Dependências para relatórios executivos avançados C-Level e Planilhas de Conferência
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
import docx
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
# ReportLab para PDFs profissionais com rodapé gráfico personalizado
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfgen import canvas

def limpar_colunas_duplicadas(df):
    novas_colunas = []
    contagem = {}
    for col in df.columns:
        col_str = str(col).strip()
        if col_str in contagem:
            contagem[col_str] += 1
            novas_colunas.append(f"{col_str}_{contagem[col_str]}")
        else:
            contagem[col_str] = 0
            novas_colunas.append(col_str)
    df.columns = novas_colunas
    return df

@st.cache_data
def processar_arquivo_upload(conteudo_bytes, nome_arquivo):
    nome = nome_arquivo.lower()
    try:
        arquivo_carregado = io.BytesIO(conteudo_bytes)
        if nome.endswith('.csv'):
            for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
                try:
                    arquivo_carregado.seek(0)
                    df = pd.read_csv(arquivo_carregado, encoding=enc, sep=';', low_memory=False)
                    if df.shape[1] <= 1:
                        arquivo_carregado.seek(0)
                        df = pd.read_csv(arquivo_carregado, encoding=enc, sep=',', low_memory=False)
                    
                    df = limpar_colunas_duplicadas(df)
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
            df = limpar_colunas_duplicadas(df)
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

@st.cache_data
def calcular_metricas_gerenciais(lista_dfs_tupla, nomes_tupla):
    lista_dfs = list(lista_dfs_tupla)
    nomes_arquivos = list(nomes_tupla)
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

def limpar_nome_local(val):
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# ====================================================================================
# GERADOR DO EXCEL DE CONFERÊNCIA (COM LÓGICA CIRÚRGICA DE FILTROS)
# ====================================================================================
def gerar_planilha_filtrada_kmb(df_localizacao, df_estoque, numero_loja_input, tipo_filtro):
    wb = openpyxl.Workbook()
    
    fonte_padrao = Font(name="Arial", size=10)
    fonte_descricao_excel = Font(name="Arial", size=9)
    fonte_negrito = Font(name="Arial", size=10, bold=True)
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_cinza_claro = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    borda_fina = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )
    borda_destaque = Border(
        left=Side(style='medium', color='000000'), right=Side(style='medium', color='000000'),
        top=Side(style='medium', color='000000'), bottom=Side(style='medium', color='000000')
    )
    
    col_loc = next((c for c in df_localizacao.columns if 'local' in c.lower()), df_localizacao.columns[2])
    col_cod_loc = next((c for c in df_localizacao.columns if 'codigo' in c.lower() or 'cod' in c.lower()), df_localizacao.columns[3])
    col_desc_loc = next((c for c in df_localizacao.columns if 'desc' in c.lower()), df_localizacao.columns[4])
    col_qtd_loc = next((c for c in df_localizacao.columns if 'qtd' in c.lower() or 'quantidade' in c.lower()), df_localizacao.columns[5])
    
    locais_por_produto = {}
    soma_comptar_por_produto = {}
    descricoes_comptar = {}
    
    for _, row in df_localizacao.iterrows():
        loc = limpar_nome_local(row.get(col_loc, ''))
        cod = str(row.get(col_cod_loc, '')).strip()
        desc = str(row.get(col_desc_loc, '')).strip()
        try:
            qtd = int(float(row.get(col_qtd_loc, 0)))
        except Exception:
            qtd = 0
            
        if cod not in locais_por_produto:
            locais_por_produto[cod] = []
            soma_comptar_por_produto[cod] = 0
            descricoes_comptar[cod] = desc
            
        locais_por_produto[cod].append((loc, qtd))
        soma_comptar_por_produto[cod] += qtd
        
    col_cod_est = next((c for c in df_estoque.columns if 'codigo' in c.lower()), 'Codigo')
    col_desc_est = next((c for c in df_estoque.columns if 'desc' in c.lower()), 'Descrição')
    col_est = next((c for c in df_estoque.columns if 'estoque' in c.lower()), 'Estoque')
    
    codigos_estoque_set = set(df_estoque[col_cod_est].astype(str).str.strip().values)
    itens_filtrados_temp = []
    
    if tipo_filtro == "SEM_SALDO":
        for cod, locs in locais_por_produto.items():
            if cod not in codigos_estoque_set:
                desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
                soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
                divergencia_val = soma_comptar_val - 0
                itens_filtrados_temp.append((cod, desc_item, 0, True, soma_comptar_val, divergencia_val, locs))
    elif tipo_filtro == "ZERADOS":
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip()
            if cod not in locais_por_produto:
                desc = row.get(col_desc_est, '')
                try:
                    est = int(float(row.get(col_est, 0)))
                except Exception:
                    est = 0
                itens_filtrados_temp.append((cod, desc, est, False, 0, -est, []))
    else:
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip()
            desc = row.get(col_desc_est, '')
            try:
                est = int(float(row.get(col_est, 0)))
            except Exception:
                est = 0
            
            locs_prod = locais_por_produto.get(cod, [])
            tem_na_comptar = len(locs_prod) > 0
            soma_comptar_val = soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
            divergencia_val = soma_comptar_val - est
            
            incluir = False
            if tipo_filtro == "TODOS":
                incluir = True
            elif tipo_filtro == "FALTAS":
                if tem_na_comptar and divergencia_val < 0:
                    incluir = True
            elif tipo_filtro == "SOBRAS":
                if tem_na_comptar and divergencia_val > 0:
                    incluir = True
                    
            if incluir:
                itens_filtrados_temp.append((cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod))
                
    itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
    locais_com_contagem = [len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0]
    max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1
    
    nome_aba = "Itens Sem Saldo" if tipo_filtro == "SEM_SALDO" else ("Zerados" if tipo_filtro == "ZERADOS" else "Saldo Loja")
    ws_saldo = wb.active
    ws_saldo.title = nome_aba
    ws_saldo.views.sheetView[0].showGridLines = True
    
    ws_saldo.page_setup.orientation = ws_saldo.ORIENTATION_LANDSCAPE
    ws_saldo.page_setup.paperSize = ws_saldo.PAPERSIZE_A4
    ws_saldo.sheet_properties.pageSetUpPr.fitToPage = True
    ws_saldo.page_setup.fitToWidth = 1
    ws_saldo.page_setup.fitToHeight = 0
    
    ws_saldo.page_margins.left = 0.2
    ws_saldo.page_margins.right = 0.2
    ws_saldo.page_margins.top = 0.3
    ws_saldo.page_margins.bottom = 0.3
    ws_saldo.page_margins.header = 0.1
    ws_saldo.page_margins.footer = 0.1
    
    ws_saldo.oddFooter.left.text = f"Loja: {numero_loja_input} | Grupo KMB"
    ws_saldo.oddFooter.center.text = "Exatidão • Controle • Eficiência"
    ws_saldo.oddFooter.right.text = "Página &P de &N"
    
    caminho_brasao_timbre = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
    if caminho_brasao_timbre:
        try:
            img_timbre = OpenpyxlImage(caminho_brasao_timbre)
            img_timbre.width = 110
            img_timbre.height = 42
            ws_saldo.add_image(img_timbre, 'A1')
        except Exception:
            pass
            
    ws_saldo['D1'] = "DATA:"
    ws_saldo['D1'].font = fonte_negrito
    ws_saldo['D1'].alignment = Alignment(horizontal="center", vertical="center")
    ws_saldo['D1'].border = borda_destaque
    
    ws_saldo['E1'] = pd.Timestamp.now().strftime('%d/%m/%Y')
    ws_saldo['E1'].font = fonte_padrao
    ws_saldo['E1'].alignment = Alignment(horizontal="center", vertical="center")
    ws_saldo['E1'].border = borda_destaque
    
    ws_saldo['D2'] = "LOJA:"
    ws_saldo['D2'].font = fonte_negrito
    ws_saldo['D2'].alignment = Alignment(horizontal="center", vertical="center")
    ws_saldo['D2'].border = borda_destaque
    
    ws_saldo['E2'] = int(numero_loja_input)
    ws_saldo['E2'].font = fonte_negrito
    ws_saldo['E2'].alignment = Alignment(horizontal="center", vertical="center")
    ws_saldo['E2'].border = borda_destaque
    
    headers_saldo = ["Codigo", "Descrição", "Estoque", "COMPTAR", "DIVERGÊNCIA"]
    for i in range(1, max_locais_encontrados + 1):
        if i == 1:
            headers_saldo.append("CONT. 1")
        else:
            headers_saldo.append(str(int(i)))
            
    for i in range(1, max_locais_encontrados + 1):
        if i == 1:
            headers_saldo.append("Local 1")
        else:
            headers_saldo.append(str(int(i)))
        
    ws_saldo.row_dimensions[1].height = 24
    ws_saldo.row_dimensions[2].height = 20
    ws_saldo.row_dimensions[3].height = 10
    
    ws_saldo.row_dimensions[4].height = 22
    for col_idx, h_text in enumerate(headers_saldo, start=1):
        cell = ws_saldo.cell(row=4, column=col_idx, value=str(h_text).replace('.0', ''))
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina
    ws_saldo.print_title_rows = '4:4'
    ws_saldo.auto_filter.ref = f"A4:{get_column_letter(len(headers_saldo))}{ws_saldo.max_row}"
        
    for cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod in itens_filtrados_temp:
        cel_comptar = soma_comptar_val if tem_na_comptar else 0
        cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)
        
        linha_dados = [cod, desc, est, cel_comptar, cel_divergencia]
        
        if tem_na_comptar:
            contagens_vals = []
            for i in range(max_locais_encontrados):
                if i < len(locs_prod):
                    val_num = locs_prod[i][1]
                    try:
                        val_int = int(float(val_num))
                        contagens_vals.append(val_int)
                    except Exception:
                        contagens_vals.append(val_num)
                else:
                    contagens_vals.append("")
                    
            locais_vals = [limpar_nome_local(locs_prod[i][0]) if i < len(locs_prod) else "" for i in range(max_locais_encontrados)]
        else:
            if tipo_filtro == "ZERADOS":
                contagens_vals = ["0 - Comptar"] + [""] * (max_locais_encontrados - 1)
            else:
                contagens_vals = ["0"] + [""] * (max_locais_encontrados - 1)
            locais_vals = ["Sem local"] + [""] * (max_locais_encontrados - 1)
            
        linha_dados.extend(contagens_vals)
        linha_dados.extend(locais_vals)
        ws_saldo.append(linha_dados)
        
    for row in ws_saldo.iter_rows(min_row=5, max_row=ws_saldo.max_row, min_col=1, max_col=len(headers_saldo)):
        for cell in row:
            if cell.column == 2:
                cell.font = fonte_descricao_excel
            else:
                cell.font = fonte_padrao
                
            cell.border = borda_fina
            if cell.row >= 5 and cell.column >= 3:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0'
            
            if cell.column == 5:
                cell.fill = fill_cinza_claro
                cell.font = Font(name="Arial", size=10, bold=True)
            
            if cell.column in [1, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif cell.column == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
    for col in ws_saldo.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= 4 and cell.value is not None:
                val_str = str(cell.value)
                if not val_str.startswith("="):
                    max_len = max(max_len, len(val_str))
        ws_saldo.column_dimensions[col_letter].width = max(max_len + 2, 7)
        
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ====================================================================================
# GERADOR DE PDF DE CONFERÊNCIA OTIMIZADO (COM LÓGICA CIRÚRGICA DE FILTROS)
# ====================================================================================
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()
    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setFont("Helvetica", 7)
        self.setFillColor(colors.HexColor('#64748b'))
        
        largura_pagina, altura_pagina = self._pagesize
        margem = 10
        
        self.setStrokeColor(colors.HexColor('#cbd5e1'))
        self.setLineWidth(0.5)
        self.line(margem, 22, largura_pagina - margem, 22)
        
        caminho_brasao = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
        if caminho_brasao:
            try:
                self.drawImage(caminho_brasao, margem, 6, width=45, height=13, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
                
        texto_esquerda = "GRUPO KMB // Auditoria Patrimonial • Exatidão • Controle • Eficiência"
        self.drawString(margem + 50, 10, texto_esquerda)
        
        texto_pagina = f"Página {self._pageNumber} de {page_count}"
        self.drawRightString(largura_pagina - margem, 10, texto_pagina)
        self.restoreState()

def gerar_pdf_conferencia_kmb(df_localizacao, df_estoque, numero_loja_input, tipo_filtro):
    col_loc = next((c for c in df_localizacao.columns if 'local' in c.lower()), df_localizacao.columns[2])
    col_cod_loc = next((c for c in df_localizacao.columns if 'codigo' in c.lower() or 'cod' in c.lower()), df_localizacao.columns[3])
    col_desc_loc = next((c for c in df_localizacao.columns if 'desc' in c.lower()), df_localizacao.columns[4])
    col_qtd_loc = next((c for c in df_localizacao.columns if 'qtd' in c.lower() or 'quantidade' in c.lower()), df_localizacao.columns[5])
    
    locais_por_produto = {}
    soma_comptar_por_produto = {}
    descricoes_comptar = {}
    
    for _, row in df_localizacao.iterrows():
        loc = limpar_nome_local(row.get(col_loc, ''))
        cod = str(row.get(col_cod_loc, '')).strip()
        desc = str(row.get(col_desc_loc, '')).strip()
        try:
            qtd = int(float(row.get(col_qtd_loc, 0)))
        except Exception:
            qtd = 0
            
        if cod not in locais_por_produto:
            locais_por_produto[cod] = []
            soma_comptar_por_produto[cod] = 0
            descricoes_comptar[cod] = desc
        locais_por_produto[cod].append((loc, qtd))
        soma_comptar_por_produto[cod] += qtd
        
    col_cod_est = next((c for c in df_estoque.columns if 'codigo' in c.lower()), 'Codigo')
    col_desc_est = next((c for c in df_estoque.columns if 'desc' in c.lower()), 'Descrição')
    col_est = next((c for c in df_estoque.columns if 'estoque' in c.lower()), 'Estoque')
    codigos_estoque_set = set(df_estoque[col_cod_est].astype(str).str.strip().values)
    
    itens_filtrados_temp = []
    if tipo_filtro == "SEM_SALDO":
        for cod, locs in locais_por_produto.items():
            if cod not in codigos_estoque_set:
                desc_item = descricoes_comptar.get(cod, "Item Avulso / Sem Cadastro")
                soma_comptar_val = soma_comptar_por_produto.get(cod, 0)
                divergencia_val = soma_comptar_val - 0
                itens_filtrados_temp.append((cod, desc_item, 0, True, soma_comptar_val, divergencia_val, locs))
    elif tipo_filtro == "ZERADOS":
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip()
            if cod not in locais_por_produto:
                desc = str(row.get(col_desc_est, ''))
                try:
                    est = int(float(row.get(col_est, 0)))
                except Exception:
                    est = 0
                itens_filtrados_temp.append((cod, desc, est, False, 0, -est, []))
    else:
        for idx, row in df_estoque.iterrows():
            cod = str(row.get(col_cod_est, '')).strip()
            desc = str(row.get(col_desc_est, ''))
            try:
                est = int(float(row.get(col_est, 0)))
            except Exception:
                est = 0
            
            locs_prod = locais_por_produto.get(cod, [])
            tem_na_comptar = len(locs_prod) > 0
            soma_comptar_val = soma_comptar_por_produto.get(cod, 0) if tem_na_comptar else 0
            divergencia_val = soma_comptar_val - est
            
            incluir = False
            if tipo_filtro == "TODOS":
                incluir = True
            elif tipo_filtro == "FALTAS":
                if tem_na_comptar and divergencia_val < 0:
                    incluir = True
            elif tipo_filtro == "SOBRAS":
                if tem_na_comptar and divergencia_val > 0:
                    incluir = True
                    
            if incluir:
                itens_filtrados_temp.append((cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod))
                
    itens_filtrados_temp.sort(key=lambda x: str(x[1]).strip().upper())
    locais_com_contagem = [len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0]
    max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=22,
        leftMargin=22,
        topMargin=20,
        bottomMargin=30
    )
    
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#1F4E78'))
    
    estilo_th = ParagraphStyle('TH_Auto', fontName='Helvetica-Bold', fontSize=6.0, leading=7, textColor=colors.white, alignment=1)
    estilo_td_codigo = ParagraphStyle('TD_Codigo', fontName='Helvetica-Bold', fontSize=7.5, leading=9, textColor=colors.HexColor('#111827'), alignment=1)
    estilo_td_centro = ParagraphStyle('TD_Centro', fontName='Helvetica', fontSize=7.5, leading=9, textColor=colors.HexColor('#111827'), alignment=1)
    estilo_td_esq = ParagraphStyle('TD_Esq', fontName='Helvetica', fontSize=8.0, leading=10, textColor=colors.HexColor('#111827'), alignment=0)
    
    story = []
    story.append(Paragraph(f"<b>GRUPO KMB // RELATÓRIO DE CONFERÊNCIA ({tipo_filtro})</b> — LOJA: {numero_loja_input} | DATA: {pd.Timestamp.now().strftime('%d/%m/%Y')}", estilo_titulo))
    story.append(Spacer(1, 3))
    
    headers_pdf = ["Código", "Descrição", "EST.", "COMP.", "DIVERG."]
    for i in range(1, max_locais_encontrados + 1):
        headers_pdf.append("C.1" if i == 1 else str(int(i)))
    for i in range(1, max_locais_encontrados + 1):
        headers_pdf.append(f"Loc.{int(i)}" if i == 1 else str(int(i)))
        
    dados_tabela = [[Paragraph(f"<nobr>{str(h).strip().replace('.0', '').replace(' ', '&nbsp;')}</nobr>", estilo_th) for h in headers_pdf]]
    
    for cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod in itens_filtrados_temp:
        cel_comptar = soma_comptar_val if tem_na_comptar else 0
        cel_divergencia = divergencia_val if tem_na_comptar else (0 - est)
        
        cod_formatado = str(cod).strip().replace(" ", "&nbsp;")
        
        linha_celulas = [
            Paragraph(f"<nobr>{cod_formatado}</nobr>", estilo_td_codigo),
            Paragraph(str(desc), estilo_td_esq),
            Paragraph(f"<nobr>{est}</nobr>", estilo_td_centro),
            Paragraph(f"<nobr>{cel_comptar}</nobr>", estilo_td_centro),
            Paragraph(f"<nobr>{cel_divergencia}</nobr>", estilo_td_centro)
        ]
        
        if tem_na_comptar:
            contagens_vals = []
            for i in range(max_locais_encontrados):
                if i < len(locs_prod):
                    contagens_vals.append(Paragraph(f"<nobr>{int(float(locs_prod[i][1]))}</nobr>", estilo_td_centro))
                else:
                    contagens_vals.append(Paragraph("", estilo_td_centro))
            
            locais_vals = []
            for i in range(max_locais_encontrados):
                if i < len(locs_prod):
                    nome_loc_original = str(limpar_nome_local(locs_prod[i][0]))
                    nome_loc_seguro = nome_loc_original.replace('.', '&#46;').replace(' ', '&nbsp;')
                    locais_vals.append(Paragraph(f"<nobr>{nome_loc_seguro}</nobr>", estilo_td_centro))
                else:
                    locais_vals.append(Paragraph("", estilo_td_centro))
        else:
            if tipo_filtro == "ZERADOS":
                contagens_vals = [Paragraph("<nobr>0&nbsp;-&nbsp;Comptar</nobr>", estilo_td_centro)] + [Paragraph("", estilo_td_centro)] * (max_locais_encontrados - 1)
            else:
                contagens_vals = [Paragraph("<nobr>0</nobr>", estilo_td_centro)] + [Paragraph("", estilo_td_centro)] * (max_locais_encontrados - 1)
            locais_vals = [Paragraph("<nobr>Sem&nbsp;local</nobr>", estilo_td_centro)] + [Paragraph("", estilo_td_centro)] * (max_locais_encontrados - 1)
            
        linha_celulas.extend(contagens_vals)
        linha_celulas.extend(locais_vals)
        dados_tabela.append(linha_celulas)
        
    largura_util = 792 - 44
    largura_codigo = 58
    largura_desc = 120
    largura_est = 24
    largura_comp = 24
    largura_div = 28
    largura_fixa_inicial = largura_codigo + largura_desc + largura_est + largura_comp + largura_div
    
    num_colunas = len(headers_pdf)
    colunas_restantes = num_colunas - 5
    largura_outras = max((largura_util - largura_fixa_inicial) / colunas_restantes, 24) if colunas_restantes > 0 else 22
    
    col_widths = [largura_codigo, largura_desc, largura_est, largura_comp, largura_div] + [largura_outras] * colunas_restantes
    
    tabela_pdf = Table(dados_tabela, colWidths=col_widths, repeatRows=1)
    tabela_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('BACKGROUND', (4, 1), (4, -1), colors.HexColor('#D9D9D9')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
        ('TOPPADDING', (0, 0), (-1, -1), 1.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
    ]))
    
    story.append(tabela_pdf)
    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer

# ====================================================================================
# GERADOR DE PDF VISUAL DOS INDICADORES (MODELO C-LEVEL)
# ====================================================================================
def gerar_pdf_indicadores_tela(resumo_unico, nome_arquivo):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle('TituloPDF', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=12, leading=16, textColor=colors.HexColor('#1F4E78'))
    estilo_sub_topo = ParagraphStyle('SubTopoPDF', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=12, textColor=colors.HexColor('#0F172A'))
    estilo_meta_topo = ParagraphStyle('MetaTopoPDF', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11, textColor=colors.HexColor('#334155'))
    
    estilo_secao = ParagraphStyle('SecaoPDF', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=10, leading=13, textColor=colors.HexColor('#1F4E78'), spaceBefore=10, spaceAfter=4)
    
    estilo_th = ParagraphStyle('TH_PDF', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=0)
    estilo_th_center = ParagraphStyle('TH_PDF_C', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.white, alignment=1)
    
    estilo_td_esq = ParagraphStyle('TD_PDF_E', fontName='Helvetica', fontSize=8, leading=11, textColor=colors.HexColor('#111827'), alignment=0)
    estilo_td_esq_bold = ParagraphStyle('TD_PDF_EB', fontName='Helvetica-Bold', fontSize=8, leading=11, textColor=colors.HexColor('#111827'), alignment=0)
    estilo_td_centro = ParagraphStyle('TD_PDF_C', fontName='Helvetica', fontSize=8, leading=11, textColor=colors.HexColor('#111827'), alignment=1)
    
    estilo_balanco_tit = ParagraphStyle('Bal_Tit', fontName='Helvetica-Bold', fontSize=8, leading=10, textColor=colors.HexColor('#1F4E78'), alignment=0)
    estilo_balanco_val = ParagraphStyle('Bal_Val', fontName='Helvetica-Bold', fontSize=11, leading=14, textColor=colors.HexColor('#0F172A'), alignment=0)
    estilo_balanco_sub = ParagraphStyle('Bal_Sub', fontName='Helvetica', fontSize=7.5, leading=10, textColor=colors.HexColor('#475569'), alignment=0)
    
    estilo_parecer = ParagraphStyle('ParecerTxt', fontName='Helvetica', fontSize=8.5, leading=12, textColor=colors.HexColor('#334155'))
    
    story = []
    
    story.append(Paragraph("<b>GRUPO KMB DIRETORIA EXECUTIVA & AUDITORIA PATRIMONIAL</b>", estilo_titulo))
    story.append(Paragraph("Relatório Executivo de Indicadores", estilo_sub_topo))
    story.append(Paragraph(f"Unidade / Arquivo: {nome_arquivo} | Data de Emissão: {pd.Timestamp.now().strftime('%d/%m/%Y')}", estilo_meta_topo))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1F4E78'), spaceBefore=4, spaceAfter=8))
    
    story.append(Paragraph("1. Indicadores de SKUs e Volumetria", estilo_secao))
    
    dados_tabela_ind = [
        [Paragraph("Indicador Operacional", estilo_th), Paragraph("Resultado Apurado", estilo_th_center), Paragraph("Detalhamento / Referência", estilo_th)]
    ]
    
    linhas_indicadores = [
        ("Total SKUs Auditados", f"{resumo_unico['total_skus']:,} SKUs", f"Pré: {resumo_unico['qtd_antes']:,.0f} un | Pós: {resumo_unico['qtd_pos']:,.0f} un"),
        ("Produtos com Falta (Negativos)", f"{resumo_unico['skus_falta']:,} SKUs ({resumo_unico['perc_falta_itens']:.2f}%)", f"Vol: -{resumo_unico['vol_falta']:,.0f} un | R$ {resumo_unico['val_falta']:,.2f}"),
        ("Produtos com Sobra (Positivos)", f"{resumo_unico['skus_sobra']:,} SKUs ({resumo_unico['perc_sobra_itens']:.2f}%)", f"Vol: +{resumo_unico['vol_sobra']:,.0f} un | R$ +{resumo_unico['val_sobra']:,.2f}"),
        ("Total Falta + Sobra (SKUs)", f"{resumo_unico['total_div']:,} SKUs ({resumo_unico['perc_divergencia']:.2f}%)", "Total de Produtos Divergentes"),
        ("Volume Físico Total de Divergências", f"{resumo_unico['vol_div_total']:,.0f} unidades ({resumo_unico['perc_vol_div']:.2f}%)", f"Movimentação Absoluta: R$ {resumo_unico['val_div_absoluto']:,.2f}"),
        ("Acuracidade Volumétrica / Financeira", f"Vol: {resumo_unico['acuracia_vol']:.2f}% | Fin: {resumo_unico['acuracia_fin']:.2f}%", "Meta Corporativa: 97,00%")
    ]
    
    for label, res, det in linhas_indicadores:
        dados_tabela_ind.append([
            Paragraph(label, estilo_td_esq_bold),
            Paragraph(res, estilo_td_centro),
            Paragraph(det, estilo_td_esq)
        ])
        
    tabela_ind_pdf = Table(dados_tabela_ind, colWidths=[150, 140, 250])
    tabela_ind_pdf.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E78')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')])
    ]))
    
    story.append(tabela_ind_pdf)
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("2. Balanço Patrimonial & Impacto Geral", estilo_secao))
    
    cor_impacto = '#10b981' if resumo_unico['impacto_sobre_antes_pct'] >= 0 else '#ef4444'
    cor_ajuste = '#10b981' if resumo_unico['val_ajuste_liquido'] >= 0 else '#ef4444'
    
    dados_balanco = [
        [
            Paragraph("ESTOQUE ANTES (SISTEMA / TEÓRICO)", estilo_balanco_tit),
            Paragraph("ESTOQUE DEPOIS (FÍSICO AUDITADO)", estilo_balanco_tit)
        ],
        [
            Paragraph(f"<b>R$ {resumo_unico['soma_antes']:,.2f}</b>", estilo_balanco_val),
            Paragraph(f"<b>R$ {resumo_unico['soma_depois']:,.2f}</b>", estilo_balanco_val)
        ],
        [
            Paragraph(f"Quantidade Pré-Inventário: {resumo_unico['qtd_antes']:,.0f} unidades", estilo_balanco_sub),
            Paragraph(f"Quantidade Pós-Inventário: {resumo_unico['qtd_pos']:,.0f} unidades", estilo_balanco_sub)
        ],
        [
            Paragraph("IMPACTO GERAL SOBRE ESTOQUE TEÓRICO", estilo_balanco_tit),
            Paragraph("DIVERGÊNCIA / AJUSTE LÍQUIDO", estilo_balanco_tit)
        ],
        [
            Paragraph(f"<font color='{cor_impacto}'><b>{resumo_unico['impacto_sobre_antes_pct']:.2f}%</b></font>", estilo_balanco_val),
            Paragraph(f"<font color='{cor_ajuste}'><b>R$ {resumo_unico['val_ajuste_liquido']:,.2f}</b></font>", estilo_balanco_val)
        ],
        [
            Paragraph("Variação percentual líquida apurada", estilo_balanco_sub),
            Paragraph("Resultado financeiro consolidado (Sobras - Faltas)", estilo_balanco_sub)
        ]
    ]
    
    tabela_balanco = Table(dados_balanco, colWidths=[270, 270])
    tabela_balanco.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8FAFC')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    
    story.append(tabela_balanco)
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("3. Parecer Técnico de Governança", estilo_secao))
    
    status_acuracia = "atendendo à meta corporativa" if resumo_unico['acuracia_vol'] >= 97 else "abaixo da meta corporativa de 97,00%"
    texto_parecer = f"""
    A unidade apresentou acuracidade volumétrica de {resumo_unico['acuracia_vol']:.2f}% ({status_acuracia}) e 
    acuracidade financeira de {resumo_unico['acuracia_fin']:.2f}%, registrando ajuste líquido de R$ {resumo_unico['val_ajuste_liquido']:,.2f} 
    ({resumo_unico['impacto_sobre_antes_pct']:.2f}% sobre o saldo teórico inicial). Recomenda-se atuação imediata nas {resumo_unico['skus_falta']} 
    linhas com quebra apontadas na Etapa 5 (Plano de Ação Prioritário).
    """
    
    tabela_parecer = Table([[Paragraph(texto_parecer.strip(), estilo_parecer)]], colWidths=[540])
    tabela_parecer.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F1F5F9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(tabela_parecer)
    
    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer

# ====================================================================================
# FUNÇÕES DE EXPORTAÇÃO C-LEVEL
# ====================================================================================
def gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(tuple(lista_dfs), tuple(nomes_arquivos))
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
    resumos = calcular_metricas_gerenciais(tuple(lista_dfs), tuple(nomes_arquivos))
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
    
    doc.add_heading("2. Detalhamento de Divergências Críticas", level=2)
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
    p_assinatura.add_run("____________________________________________________\n").bold = True
    p_assinatura.add_run("Time de Inventário / Auditoria C-Level\n")
    p_assinatura.add_run("Grupo KMB - Exatidão • Controle • Eficiência")
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(tuple(lista_dfs), tuple(nomes_arquivos))
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
    estilo_titulo = ParagraphStyle('TituloExecutivo', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=colors.HexColor('#1F4E78'), spaceAfter=10)
    estilo_sub = ParagraphStyle('SubExecutivo', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, leading=16, textColor=colors.HexColor('#0F172A'), spaceBefore=12, spaceAfter=6)
    estilo_texto = ParagraphStyle('TextoExecutivo', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=13, textColor=colors.HexColor('#334155'), spaceAfter=4)
    
    story = []
    story.append(Paragraph("⚡ GRUPO KMB // RELATÓRIO DE AUDITORIA PATRIMONIAL C-LEVEL", estilo_titulo))
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
        
    story.append(Paragraph("2. Detalhamento de Divergências Críticas", estilo_sub))
    for idx, df in enumerate(lista_dfs):
        col_ajuste = 'Ajuste' if 'Ajuste' in df.columns else None
        if col_ajuste and col_ajuste in df.columns:
            total_faltas = len(df[df[col_ajuste] < 0])
            total_sobras = len(df[df[col_ajuste] > 0])
            total_zerados = len(df[(df[col_ajuste] == 0) & (df.get('Saldo em Estoque', 1) == 0)]) if 'Saldo em Estoque' in df.columns else len(df[df[col_ajuste] == 0])
            texto_detalhe = f"""
            <b>• Base / Unidade:</b> {nomes_arquivos[idx]}<br/>
            <b>- Itens em Falta (Quebras):</b> {total_faltas:,}<br/>
            <b>- Itens em Sobra:</b> {total_sobras:,}<br/>
            <b>- Itens com Estoque Congelado Zero:</b> {total_zerados:,}
            """
            story.append(Paragraph(texto_detalhe, estilo_texto))
            story.append(Spacer(1, 4))
            
    story.append(Paragraph("3. Validação e Assinatura C-Level", estilo_sub))
    story.append(Paragraph("Declaro que os dados acima foram auditados eletronicamente através do Protocolo J.A.R.V.I.S. / Grupo KMB e refletem com fidelidade a apuração patrimonial realizada.", estilo_texto))
    story.append(Spacer(1, 15))
    
    texto_assinatura = """
    <b>____________________________________________________</b><br/>
    <b>Time de Inventário / Auditoria C-Level</b><br/>
    Grupo KMB - Exatidão • Controle • Eficiência
    """
    story.append(Paragraph(texto_assinatura, estilo_texto))
        
    doc_pdf.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer

def gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos):
    resumos = calcular_metricas_gerenciais(tuple(lista_dfs), tuple(nomes_arquivos))
    prs = Presentation()
    slide_layout = prs.slide_layouts[6] # em branco
    
    # ----------------------------------------------------
    # SLIDE 1: CAPA EXECUTIVA MODERNA
    # ----------------------------------------------------
    slide1 = prs.slides.add_slide(slide_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = PptxRGBColor(11, 15, 25) # #0b0f19
    bg1.line.fill.background()
    
    tb1 = slide1.shapes.add_textbox(PptxInches(1), PptxInches(1.8), PptxInches(11.3), PptxInches(4))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    p1 = tf1.paragraphs[0]
    p1.text = "⚡ GRUPO KMB // DIRETORIA EXECUTIVA"
    p1.font.size = PptxPt(32)
    p1.font.bold = True
    p1.font.color.rgb = PptxRGBColor(0, 242, 254)
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf1.add_paragraph()
    p2.text = "Relatório Executivo de Auditoria Patrimonial & Indicadores"
    p2.font.size = PptxPt(20)
    p2.font.color.rgb = PptxRGBColor(243, 244, 246)
    p2.space_before = PptxPt(12)
    p2.alignment = PP_ALIGN.CENTER
    
    p3 = tf1.add_paragraph()
    p3.text = f"Data de Emissão: {pd.Timestamp.now().strftime('%d/%m/%Y')} | Protocolo J.A.R.V.I.S."
    p3.font.size = PptxPt(13)
    p3.font.color.rgb = PptxRGBColor(156, 163, 175)
    p3.space_before = PptxPt(25)
    p3.alignment = PP_ALIGN.CENTER

    for r in resumos:
        # ----------------------------------------------------
        # SLIDE 2: RESUMO EXECUTIVO CONSOLIDADO
        # ----------------------------------------------------
        slide2 = prs.slides.add_slide(slide_layout)
        tb_head2 = slide2.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.5), PptxInches(0.8))
        tf_head2 = tb_head2.text_frame
        p_h2 = tf_head2.paragraphs[0]
        p_h2.text = f"📋 Resumo Executivo // {r['nome']}"
        p_h2.font.size = PptxPt(22)
        p_h2.font.bold = True
        p_h2.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        card_res = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(11.2), PptxInches(5.0))
        card_res.fill.solid()
        card_res.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
        card_res.line.color.rgb = PptxRGBColor(203, 213, 225)
        
        tf_res = card_res.text_frame
        tf_res.word_wrap = True
        tf_res.margin_left = PptxInches(0.5)
        tf_res.margin_top = PptxInches(0.4)
        
        p_rt = tf_res.paragraphs[0]
        p_rt.text = "🎯 Visão Geral de Governança Patrimonial"
        p_rt.font.size = PptxPt(16)
        p_rt.font.bold = True
        p_rt.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        for l in [
            f"• Total de SKUs Auditados: {r['total_skus']:,} itens",
            f"• Acuracidade Volumétrica: {r['acuracia_vol']:.2f}% (Meta Corporativa: 97.00%)",
            f"• Acuracidade Financeira: {r['acuracia_fin']:.2f}%",
            f"• Valor Teórico Inicial (Antes): R$ {r['soma_antes']:,.2f}",
            f"• Valor Auditado Final (Depois): R$ {r['soma_depois']:,.2f}",
            f"• Ajuste Líquido Consolidado: R$ {r['val_ajuste_liquido']:,.2f} ({r['impacto_sobre_antes_pct']:.2f}%)"
        ]:
            pl = tf_res.add_paragraph()
            pl.text = l
            pl.font.size = PptxPt(14)
            pl.font.color.rgb = PptxRGBColor(51, 65, 85)
            pl.space_before = PptxPt(12)

        # ----------------------------------------------------
        # SLIDE 3: INDICADORES DE SKUS & VOLUMETRIA
        # ----------------------------------------------------
        slide3 = prs.slides.add_slide(slide_layout)
        tb_head3 = slide3.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.5), PptxInches(0.8))
        tf_head3 = tb_head3.text_frame
        p_h3 = tf_head3.paragraphs[0]
        p_h3.text = f"📊 Indicadores de SKUs & Volumetria // {r['nome']}"
        p_h3.font.size = PptxPt(22)
        p_h3.font.bold = True
        p_h3.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        card3_l = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(5.4), PptxInches(5.0))
        card3_l.fill.solid()
        card3_l.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
        card3_l.line.color.rgb = PptxRGBColor(203, 213, 225)
        
        tf3_l = card3_l.text_frame
        tf3_l.word_wrap = True
        tf3_l.margin_left = PptxInches(0.4)
        tf3_l.margin_top = PptxInches(0.4)
        
        p3lt = tf3_l.paragraphs[0]
        p3lt.text = "📦 Composição de SKUs"
        p3lt.font.size = PptxPt(15)
        p3lt.font.bold = True
        p3lt.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        for it in [
            f"• Total SKUs: {r['total_skus']:,}",
            f"• Produtos com Falta: {r['skus_falta']:,} ({r['perc_falta_itens']:.2f}%)",
            f"• Produtos com Sobra: {r['skus_sobra']:,} ({r['perc_sobra_itens']:.2f}%)",
            f"• SKUs Exatos/Conformes: {r['skus_exatos']:,}"
        ]:
            pt = tf3_l.add_paragraph()
            pt.text = it
            pt.font.size = PptxPt(13)
            pt.font.color.rgb = PptxRGBColor(51, 65, 85)
            pt.space_before = PptxPt(10)

        card3_r = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(6.6), PptxInches(1.5), PptxInches(5.4), PptxInches(5.0))
        card3_r.fill.solid()
        card3_r.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
        card3_r.line.color.rgb = PptxRGBColor(203, 213, 225)
        
        tf3_r = card3_r.text_frame
        tf3_r.word_wrap = True
        tf3_r.margin_left = PptxInches(0.4)
        tf3_r.margin_top = PptxInches(0.4)
        
        p3rt = tf3_r.paragraphs[0]
        p3rt.text = "⚖️ Volume Físico Absoluto"
        p3rt.font.size = PptxPt(15)
        p3rt.font.bold = True
        p3rt.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        for it in [
            f"• Quantidade Antes (Pré): {r['qtd_antes']:,.0f} un",
            f"• Quantidade Depois (Pós): {r['qtd_pos']:,.0f} un",
            f"• Volume Total Divergência: {r['vol_div_total']:,.0f} un",
            f"• Percentual de Volume Div.: {r['perc_vol_div']:.2f}%"
        ]:
            pt = tf3_r.add_paragraph()
            pt.text = it
            pt.font.size = PptxPt(13)
            pt.font.color.rgb = PptxRGBColor(51, 65, 85)
            pt.space_before = PptxPt(10)

        # ----------------------------------------------------
        # SLIDE 4: BALANÇO PATRIMONIAL & IMPACTO
        # ----------------------------------------------------
        slide4 = prs.slides.add_slide(slide_layout)
        tb_head4 = slide4.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.5), PptxInches(0.8))
        tf_head4 = tb_head4.text_frame
        p_h4 = tf_head4.paragraphs[0]
        p_h4.text = f"💰 Balanço Patrimonial & Impacto // {r['nome']}"
        p_h4.font.size = PptxPt(22)
        p_h4.font.bold = True
        p_h4.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        card4 = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(11.2), PptxInches(5.0))
        card4.fill.solid()
        card4.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
        card4.line.color.rgb = PptxRGBColor(203, 213, 225)
        
        tf4 = card4.text_frame
        tf4.word_wrap = True
        tf4.margin_left = PptxInches(0.5)
        tf4.margin_top = PptxInches(0.4)
        
        p4t = tf4.paragraphs[0]
        p4t.text = "📈 Demonstração Financeira Consolidada"
        p4t.font.size = PptxPt(16)
        p4t.font.bold = True
        p4t.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        for it in [
            f"• Estoque Teórico (Antes): R$ {r['soma_antes']:,.2f}",
            f"• Estoque Físico Auditado (Depois): R$ {r['soma_depois']:,.2f}",
            f"• Ajuste Líquido Financeiro: R$ {r['val_ajuste_liquido']:,.2f}",
            f"• Impacto Geral sobre o Saldo: {r['impacto_sobre_antes_pct']:.2f}%",
            f"• Total em Faltas (Quebras): R$ {r['val_falta']:,.2f}",
            f"• Total em Sobras: R$ +{r['val_sobra']:,.2f}"
        ]:
            pt = tf4.add_paragraph()
            pt.text = it
            pt.font.size = PptxPt(13.5)
            pt.font.color.rgb = PptxRGBColor(51, 65, 85)
            pt.space_before = PptxPt(10)

        # ----------------------------------------------------
        # SLIDE 5: ANÁLISE DE CRITICIDADE E QUEBRAS
        # ----------------------------------------------------
        slide5 = prs.slides.add_slide(slide_layout)
        tb_head5 = slide5.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.5), PptxInches(0.8))
        tf_head5 = tb_head5.text_frame
        p_h5 = tf_head5.paragraphs[0]
        p_h5.text = f"🚨 Análise de Criticidade & Desvios // {r['nome']}"
        p_h5.font.size = PptxPt(22)
        p_h5.font.bold = True
        p_h5.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        card5 = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(11.2), PptxInches(5.0))
        card5.fill.solid()
        card5.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
        card5.line.color.rgb = PptxRGBColor(203, 213, 225)
        
        tf5 = card5.text_frame
        tf5.word_wrap = True
        tf5.margin_left = PptxInches(0.5)
        tf5.margin_top = PptxInches(0.4)
        
        p5t = tf5.paragraphs[0]
        p5t.text = "⚠️ Diagnóstico Operacional C-Level"
        p5t.font.size = PptxPt(16)
        p5t.font.bold = True
        p5t.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        for it in [
            f"• SKUs com Quebra Crítica (Falta): {r['skus_falta']:,} produtos",
            f"• Volume Financeiro de Faltas: R$ {r['val_falta']:,.2f}",
            f"• SKUs com Sobra Operacional: {r['skus_sobra']:,} produtos",
            f"• Volume Financeiro de Sobras: R$ +{r['val_sobra']:,.2f}",
            "• Recomendação: Atuação imediata nas linhas de maior desvio apontadas na Etapa 5."
        ]:
            pt = tf5.add_paragraph()
            pt.text = it
            pt.font.size = PptxPt(13.5)
            pt.font.color.rgb = PptxRGBColor(51, 65, 85)
            pt.space_before = PptxPt(12)

        # ----------------------------------------------------
        # SLIDE 6: PLANO DE AÇÃO C-LEVEL & GOVERNANÇA
        # ----------------------------------------------------
        slide6 = prs.slides.add_slide(slide_layout)
        tb_head6 = slide6.shapes.add_textbox(PptxInches(0.8), PptxInches(0.5), PptxInches(11.5), PptxInches(0.8))
        tf_head6 = tb_head6.text_frame
        p_h6 = tf_head6.paragraphs[0]
        p_h6.text = f"🎯 Plano de Ação C-Level & Governança // {r['nome']}"
        p_h6.font.size = PptxPt(22)
        p_h6.font.bold = True
        p_h6.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        card6 = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(1.5), PptxInches(11.2), PptxInches(5.0))
        card6.fill.solid()
        card6.fill.fore_color.rgb = PptxRGBColor(248, 250, 252)
        card6.line.color.rgb = PptxRGBColor(203, 213, 225)
        
        tf6 = card6.text_frame
        tf6.word_wrap = True
        tf6.margin_left = PptxInches(0.5)
        tf6.margin_top = PptxInches(0.4)
        
        p6t = tf6.paragraphs[0]
        p6t.text = "🛡️ Validação & Diretrizes Estratégicas"
        p6t.font.size = PptxPt(16)
        p6t.font.bold = True
        p6t.font.color.rgb = PptxRGBColor(31, 78, 120)
        
        for it in [
            "1. Auditoria Eletrônica Validada via Protocolo J.A.R.V.I.S. / Grupo KMB.",
            "2. Homologação dos ajustes e baixa contábil das quebras críticas identificadas.",
            "3. Implantação de rondas preventivas nos setores com maior índice de divergência.",
            "____________________________________________________",
            "Time de Inventário / Auditoria C-Level • Exatidão • Controle • Eficiência"
        ]:
            pt = tf6.add_paragraph()
            pt.text = it
            pt.font.size = PptxPt(13)
            pt.font.color.rgb = PptxRGBColor(51, 65, 85)
            pt.space_before = PptxPt(10)

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

def main():
    st.set_page_config(page_title="GRUPO KMB // TIME DE INVENTÁRIO", page_icon="⚡", layout="wide")
    
    SENHA_MESTRE = st.secrets.get("SENHA_MESTRE", "novacore2026")
    
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False
        
    caminho_logo_tela = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
    if not st.session_state.autenticado:
        st.markdown("""
        <style>
        .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
        </style>
        """, unsafe_allow_html=True)
        
        col_l1, col_l2, col_l3 = st.columns([1, 1.5, 1])
        with col_l2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            if caminho_logo_tela:
                st.image(caminho_logo_tela, use_container_width=True)
            else:
                st.warning("⚠️ Brasão não encontrado na pasta.")
            
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
    .main-header { padding: 0px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px; }
    .jarvis-title {
        font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
        background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)
    
    with st.sidebar:
        if caminho_logo_tela:
            st.image(caminho_logo_tela, use_container_width=True)
        st.markdown("### **GRUPO KMB**")
        st.markdown("*Exatidão • Controle • Eficiência*")
        st.markdown("---")
        
        modo_operacao = st.radio(
            "Modo de Operação:",
            ["Auditoria de Inventário (5 Etapas)", "Modo Planilha de Conferência"],
            key="radio_modo_operacao"
        )
        
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
    
    if modo_operacao == "Modo Planilha de Conferência":
        st.markdown("### 📋 MODO PLANILHA DE CONFERÊNCIA // GERAÇÃO AUTOMATIZADA")
        st.markdown("Insira o número da loja e carregue os arquivos abaixo para gerar as listas com todas as colunas preservadas em linha única (ordem alfabética pela descrição):")
        
        numero_loja_ui = st.number_input("Número da Loja (ex: 13016 ou 17315)", min_value=1, max_value=99999, value=13016, step=1)
        
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            arq_loc = st.file_uploader("1. Arquivo de Itens por Localização (ex: Itens_localizacao.csv)", type=['csv', 'xlsx', 'xls'], key="up_loc")
        with col_up2:
            arq_est = st.file_uploader("2. Arquivo de Saldo de Estoque (ex: saldo-estoque.csv)", type=['csv', 'xlsx', 'xls'], key="up_est")
            
        if arq_loc and arq_est:
            df_loc_temp, msg1 = processar_arquivo_upload(arq_loc.getvalue(), arq_loc.name)
            df_est_temp, msg2 = processar_arquivo_upload(arq_est.getvalue(), arq_est.name)
            
            if df_loc_temp is not None and df_est_temp is not None:
                st.success("✅ Arquivos carregados e validados com sucesso!")
                
                st.markdown("---")
                st.markdown("### 👁️ Pré-visualização na Tela")
                
                buf_preview = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "TODOS")
                
                df_preview_exibir = pd.read_excel(buf_preview, sheet_name=0, skiprows=3, dtype=str).fillna("")
                
                cols_vistas = {}
                novas_colunas = []
                for c in df_preview_exibir.columns:
                    c_str = str(c).strip()
                    if '.' in c_str and c_str.split('.')[0].isdigit():
                        c_str = c_str.split('.')[0]
                    
                    if c_str in cols_vistas:
                        cols_vistas[c_str] += 1
                        novas_colunas.append(f"{c_str}_{cols_vistas[c_str]}")
                    else:
                        cols_vistas[c_str] = 0
                        novas_colunas.append(c_str)
                df_preview_exibir.columns = novas_colunas
                
                for col in df_preview_exibir.columns:
                    if "Descrição" not in col:
                        df_preview_exibir[col] = df_preview_exibir[col].apply(
                            lambda x: x[:-2] if isinstance(x, str) and x.endswith('.0') else x
                        )
                
                st.dataframe(df_preview_exibir, use_container_width=True, height=420)
                
                st.download_button(
                    label="📥 Baixar Pré-visualização em Excel (.xlsx)",
                    data=buf_preview,
                    file_name=f"Preview_SaldoLoja_Loja_{numero_loja_ui}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown("### 📥 1. Listas de Conferência em Excel (Ordem Alfabética)")
                c_imp1, c_imp2, c_imp3, c_imp4 = st.columns(4)
                
                with c_imp1:
                    buf_zerados = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS")
                    st.download_button("📥 Baixar Excel (Zerados)", data=buf_zerados, file_name=f"Excel_Zerados_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with c_imp2:
                    buf_faltas = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS")
                    st.download_button("📥 Baixar Excel (Faltas)", data=buf_faltas, file_name=f"Excel_Faltas_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with c_imp3:
                    buf_sobras = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS")
                    st.download_button("📥 Baixar Excel (Sobras)", data=buf_sobras, file_name=f"Excel_Sobras_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with c_imp4:
                    buf_sem_saldo = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO")
                    st.download_button("📥 Baixar Excel (Sem Saldo)", data=buf_sem_saldo, file_name=f"Excel_SemSaldo_Loja_{numero_loja_ui}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    
                data_atual_str = pd.Timestamp.now().strftime('%d.%m.%Y')
                st.markdown("---")
                st.markdown("### 📄 2. Listas de Conferência em PDF (Títulos Otimizados sem Quebras + Brasão)")
                c_pdf1, c_pdf2, c_pdf3, c_pdf4 = st.columns(4)
                
                with c_pdf1:
                    pdf_zerados = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS")
                    st.download_button("📑 Baixar PDF (Zerados)", data=pdf_zerados, file_name=f"{numero_loja_ui}.{data_atual_str}.zerados.pdf", mime="application/pdf", use_container_width=True)
                with c_pdf2:
                    pdf_faltas = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS")
                    st.download_button("📑 Baixar PDF (Faltas)", data=pdf_faltas, file_name=f"{numero_loja_ui}.{data_atual_str}.faltas.pdf", mime="application/pdf", use_container_width=True)
                with c_pdf3:
                    pdf_sobras = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS")
                    st.download_button("📑 Baixar PDF (Sobras)", data=pdf_sobras, file_name=f"{numero_loja_ui}.{data_atual_str}.sobras.pdf", mime="application/pdf", use_container_width=True)
                with c_pdf4:
                    pdf_sem_saldo = gerar_pdf_conferencia_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO")
                    st.download_button("📑 Baixar PDF (Sem Saldo)", data=pdf_sem_saldo, file_name=f"{numero_loja_ui}.{data_atual_str}.sem_saldo.pdf", mime="application/pdf", use_container_width=True)
            else:
                st.error("❌ Erro ao processar os arquivos.")
        return
    
    arquivos_submetidos = st.file_uploader("Carregue o arquivo oficial de inventário (ex: 13372.07.11.2026.csv) ou múltiplos arquivos", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)
    
    if arquivos_submetidos:
        lista_dfs = []
        nomes_arquivos = []
        for arq in arquivos_submetidos:
            df_temp, msg = processar_arquivo_upload(arq.getvalue(), arq.name)
            if df_temp is not None:
                lista_dfs.append(df_temp)
                nomes_arquivos.append(arq.name)
        
        if lista_dfs:
            if len(lista_dfs) == 1:
                df_filtrado = lista_dfs[0]
                num_cols = df_filtrado.select_dtypes(include=['number']).columns.tolist()
                cat_cols = df_filtrado.select_dtypes(include=['object']).columns.tolist()
                all_cols = df_filtrado.columns.tolist()
                
                resumo_unico = calcular_metricas_gerenciais(tuple(lista_dfs), tuple(nomes_arquivos))[0]
                
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
                
                buf_pdf_indicadores = gerar_pdf_indicadores_tela(resumo_unico, nomes_arquivos[0])
                st.download_button(
                    label="📑 Baixar Relatório de Indicadores em PDF",
                    data=buf_pdf_indicadores,
                    file_name=f"Relatorio_Indicadores_{nomes_arquivos[0].rsplit('.', 1)[0]}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                st.markdown("---")
                
                with st.expander("🎯 ETAPA 2: Concentração Financeira & Curva ABC (Pareto 80/20)", expanded=True):
                    if len(num_cols) >= 1 and len(all_cols) > 0:
                        col_sel1, col_sel2, col_sel3 = st.columns(3)
                        with col_sel1:
                            coluna_categoria_escolhida = st.selectbox("Dimensão (Eixo X):", options=all_cols, index=all_cols.index('Descrição do Produto') if 'Descrição do Produto' in all_cols else 0, key="abc_eixo_x")
                        with col_sel2:
                            idx_padrao_y = identificar_coluna_ideal(num_cols, ['valor de estoque', 'valor', 'saldo'])
                            eixo_y_escolhido = st.selectbox("Métrica Principal (Eixo Y):", options=num_cols, index=idx_padrao_y if idx_padrao_y < len(num_cols) else 0, key="abc_eixo_y")
                        with col_sel3:
                            filtro_classe_abc = st.selectbox("Filtro Curva ABC:", options=["Todas as Classes (Geral)", "Classe A (80%)", "Classe B (95%)", "Classe C (100%)"], index=0, key="abc_filtro_classe")
                        
                        df_abc = df_filtrado.sort_values(by=eixo_y_escolhido, ascending=False).copy()
                        df_abc['Soma_Acumulada'] = df_abc[eixo_y_escolhido].cumsum()
                        soma_total = df_abc[eixo_y_escolhido].sum()
                        
                        if soma_total > 0:
                            df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                            df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(lambda p: "Classe A (80%)" if p <= 80 else ("Classe B (95%)" if p <= 95 else "Classe C (100%)"))
                            
                            if filtro_classe_abc != "Todas as Classes (Geral)":
                                df_abc_filtrado = df_abc[df_abc['Classe_ABC'] == filtro_classe_abc].copy()
                            else:
                                df_abc_filtrado = df_abc.copy()
                            
                            df_plot = df_abc_filtrado.head(50).copy()
                            if not df_plot.empty:
                                fig_abc = go.Figure()
                                fig_abc.add_trace(go.Bar(x=df_plot[coluna_categoria_escolhida], y=df_plot[eixo_y_escolhido], name=eixo_y_escolhido, marker=dict(color=df_plot[eixo_y_escolhido], colorscale=[[0, '#1e3a8a'], [0.5, '#3b82f6'], [1, '#00f2fe']])))
                                fig_abc.add_trace(go.Scatter(x=df_plot[coluna_categoria_escolhida], y=df_plot['%_Acumulado'], name="% Acumulado (Pareto)", mode='lines+markers', yaxis='y2', line=dict(color='#f59e0b', width=3)))
                                fig_abc.update_layout(template="plotly_dark", paper_bgcolor='rgba(11, 15, 25, 0.9)', plot_bgcolor='rgba(11, 15, 25, 0.9)', height=480, yaxis=dict(title=eixo_y_escolhido), yaxis2=dict(title="% Acumulado", overlaying='y', side='right', range=[0, 105], showgrid=False))
                                st.plotly_chart(fig_abc, use_container_width=True)
                
                with st.expander("🚨 ETAPA 3: Matriz de Criticidade & Discrepância Financeira", expanded=False):
                    if len(num_cols) >= 2:
                        col_nome_prod = 'Descrição do Produto' if 'Descrição do Produto' in df_filtrado.columns else all_cols[0]
                        idx_val_div = identificar_coluna_ideal(num_cols, ['valor de ajuste', 'ajuste', 'valor', 'divergencia'])
                        col_val_div = num_cols[idx_val_div] if idx_val_div < len(num_cols) else num_cols[0]
                        
                        df_critico = df_filtrado.copy()
                        df_critico['Impacto_Absoluto'] = df_critico[col_val_div].abs()
                        df_critico = df_critico.sort_values(by='Impacto_Absoluto', ascending=False).head(50)
                        if not df_critico.empty:
                            fig_critico = px.bar(df_critico, x=col_nome_prod, y=col_val_div, title="Top 50 Maiores Desvios Financeiros (R$)", template="plotly_dark", color=col_val_div, color_continuous_scale=['#ef4444', '#f59e0b', '#3b82f6'])
                            fig_critico.update_layout(paper_bgcolor='rgba(11, 15, 25, 0.9)', plot_bgcolor='rgba(11, 15, 25, 0.9)', height=480)
                            st.plotly_chart(fig_critico, use_container_width=True)
                
                with st.expander("🛡️ ETAPA 4: Matriz de Risco & Acurácia Ponderada por Impacto", expanded=False):
                    if len(num_cols) >= 2 and len(all_cols) > 0:
                        df_risco = df_filtrado.copy()
                        col_val_aju_etapa4 = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df_risco.columns else next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), num_cols[0])
                        df_risco['Categoria_Prod'] = df_risco['Descrição do Produto'].astype(str).str.split().str[0]
                        df_agrupado_val = df_risco.groupby('Categoria_Prod')[col_val_aju_etapa4].sum().reset_index().sort_values(by=col_val_aju_etapa4, ascending=True).head(15)
                        fig_risco = px.bar(df_agrupado_val, x='Categoria_Prod', y=col_val_aju_etapa4, title="Balanço Líquido de Divergências por Categoria (R$)", template="plotly_dark", color=col_val_aju_etapa4, color_continuous_scale=['#ef4444', '#3b82f6', '#10b981'])
                        fig_risco.update_layout(paper_bgcolor='rgba(11, 15, 25, 0.9)', plot_bgcolor='rgba(11, 15, 25, 0.9)', height=420)
                        st.plotly_chart(fig_risco, use_container_width=True)
                
                with st.expander("⚡ ETAPA 5: CENTRO DE COMANDO C-LEVEL (PRESCRITIVO & OPERACIONAL)", expanded=True):
                    tab_e5_1, tab_e5_2, tab_e5_3, tab_e5_4 = st.tabs(["🔍 1. Matriz Causa Raiz", "📈 2. Quadrante 4D", "🎯 3. Simulador Meta 97%", "📋 4. Plano de Ação Prioritário"])
                    col_val_aju_e5 = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df_filtrado.columns else next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), num_cols[0])
                    col_ajuste_e5 = 'Ajuste' if 'Ajuste' in df_filtrado.columns else 'Ajuste'
                    
                    df_causa = df_filtrado.copy()
                    def classificar_causa(row):
                        aj = row.get(col_ajuste_e5, 0)
                        val = abs(row.get(col_val_aju_e5, 0))
                        saldo_sis = row.get('Saldo em Estoque', 1)
                        if aj < 0 and val > 200: return "🚨 Furto / Extravio Crítico (Alto Valor)"
                        elif aj == 0 and saldo_sis == 0: return "👻 Estoque Fantasma / Zumbi"
                        elif abs(aj) in [6, 12, 24]: return "📦 Erro de Recebimento / Caixa Fechada"
                        elif aj != 0: return "🔄 Divergência Operacional / Quebra Padrão"
                        else: return "✅ Conforme"
                    df_causa['Perfil_Causa'] = df_causa.apply(classificar_causa, axis=1)
                    with tab_e5_1:
                        resumo_causas = df_causa.groupby('Perfil_Causa').agg(SKUs=('Código do Produto', 'count'), Impacto_Financeiro=(col_val_aju_e5, 'sum')).reset_index()
                        st.dataframe(resumo_causas.style.format({'SKUs': '{:,}', 'Impacto_Financeiro': 'R$ {:,.2f}'}), use_container_width=True)
                    with tab_e5_2:
                        df_scatter = df_causa[df_causa[col_ajuste_e5] != 0].copy()
                        df_scatter['Tamanho_Bolha'] = df_scatter[col_val_aju_e5].abs().clip(lower=1)
                        fig_4d = px.scatter(df_scatter, x='Saldo em Estoque', y=col_val_aju_e5, size='Tamanho_Bolha', color='Perfil_Causa', hover_name='Descrição do Produto', template="plotly_dark", height=450)
                        st.plotly_chart(fig_4d, use_container_width=True)
                    with tab_e5_3:
                        c_sim1, c_sim2 = st.columns(2)
                        with c_sim1:
                            st.metric("Acurácia Volumétrica Atual", f"{resumo_unico['acuracia_vol']:.2f}%")
                        with c_sim2:
                            perc_recup = st.slider("Simular % de Recuperação de Faltas", 0, 100, 30, 5, key="slider_meta_e5")
                            quebras_totais = abs(df_causa[df_causa[col_ajuste_e5] < 0][col_val_aju_e5].sum())
                            st.success(f"💰 Ganho de Caixa: R$ {quebras_totais * (perc_recup / 100.0):,.2f}")
                    with tab_e5_4:
                        df_plano = df_causa[df_causa[col_ajuste_e5] < 0].sort_values(by=col_val_aju_e5, ascending=True).head(15).copy()
                        if not df_plano.empty:
                            st.dataframe(df_plano[['Código do Produto', 'Descrição do Produto', 'Saldo em Estoque', col_ajuste_e5, col_val_aju_e5, 'Perfil_Causa']].style.format({col_val_aju_e5: 'R$ {:,.2f}'}), use_container_width=True)
            else:
                st.markdown("## ⚡ AUDITORIA COMPARATIVA MULTI-INVENTÁRIO C-LEVEL")
                resumos = calcular_metricas_gerenciais(tuple(lista_dfs), tuple(nomes_arquivos))
                
                col_cards = st.columns(len(lista_dfs))
                for idx, r in enumerate(resumos):
                    with col_cards[idx]:
                        with st.container(border=True):
                            st.markdown(f"#### 📄 {r['nome']}")
                            st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{r['total_skus']:,}", f"Pré: {r['qtd_antes']:,.0f} un | Pós: {r['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{r['skus_falta']:,} SKUs ({r['perc_falta_itens']:.2f}%)", f"Vol: -{r['vol_falta']:,.0f} un | R$ {r['val_falta']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{r['skus_sobra']:,} SKUs ({r['perc_sobra_itens']:.2f}%)", f"Vol: +{r['vol_sobra']:,.0f} un | R$ +{r['val_sobra']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Total Falta + Sobra (SKUs)", f"{r['total_div']:,} SKUs ({r['perc_divergencia']:.2f}%)", "Total de Produtos Divergentes", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Volume Físico Total Divergências (Sobra + Falta)", f"{r['vol_div_total']:,.0f} unidades ({r['perc_vol_div']:.2f}%)", f"R$ {r['val_div_absoluto']:,.2f}", eh_negativo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Acuracidade Volumétrica / Financeira", f"Vol: {r['acuracia_vol']:.2f}% | Fin: {r['acuracia_fin']:.2f}%", "Meta Corporativa 97%", eh_negativo=(r['acuracia_vol'] < 97), eh_positivo=(r['acuracia_vol'] >= 97)), unsafe_allow_html=True)
                            
                            st.markdown("##### 💰 Balanço Patrimonial")
                            st.markdown(formatar_kpi_html("Estoque Antes (Sistema / Teórico)", f"R$ {r['soma_antes']:,.2f}", f"Qtd Pré: {r['qtd_antes']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Estoque Depois (Físico Auditado)", f"R$ {r['soma_depois']:,.2f}", f"Qtd Pós: {r['qtd_pos']:,.0f} un", eh_positivo=True), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Impacto Geral", f"{r['impacto_sobre_antes_pct']:.2f}%", f"Ajuste Líquido: R$ {r['val_ajuste_liquido']:,.2f}", eh_negativo=(r['impacto_sobre_antes_pct'] < 0), eh_positivo=(r['impacto_sobre_antes_pct'] >= 0)), unsafe_allow_html=True)
                            st.markdown(formatar_kpi_html("Divergência / Ajuste Líquido", f"R$ {r['val_ajuste_liquido']:,.2f}", "Variação Apurada", eh_negativo=(r['val_ajuste_liquido'] < 0), eh_positivo=(r['val_ajuste_liquido'] >= 0)), unsafe_allow_html=True)
                st.markdown("---")
                st.subheader("📊 Gráficos Consolidados")
                
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
            
            st.markdown("---")
            st.subheader("📥 Central de Relatórios Executivos C-Level")
            col_e1, col_e2, col_e3, col_e4 = st.columns(4)
            with col_e1: st.download_button("📊 Excel Gerencial", data=gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_e2: st.download_button("📄 Word Gerencial", data=gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_e3: st.download_button("📑 PDF Gerencial", data=gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.pdf", mime="application/pdf")
            with col_e4: st.download_button("📊 Apresentação PPTX", data=gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Apresentacao_GrupoKMB.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    main()
