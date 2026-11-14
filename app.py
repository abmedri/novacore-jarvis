# ====================================================================================
# ⚡ GRUPO KMB // PROTOCOLO J.A.R.V.I.S. (VERSÃO OFICIAL UNIFICADA COMPLETA)
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


def limpar_nome_local(val):
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


# ====================================================================================
# GERADOR DO BLOCO 2: PLANILHA DE CONFERÊNCIA COM LOGOTIPO NO TOPO E RODAPÉ DE IMPRESSÃO
# ====================================================================================
def gerar_planilha_filtrada_kmb(df_localizacao, df_estoque, numero_loja_input, tipo_filtro):
    wb = openpyxl.Workbook()
    
    fonte_padrao = Font(name="Arial", size=10)
    fonte_negrito = Font(name="Arial", size=10, bold=True)
    fonte_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3')
    )
    borda_destaque = Border(
        left=Side(style='medium', color='000000'), right=Side(style='medium', color='000000'),
        top=Side(style='medium', color='000000'), bottom=Side(style='medium', color='000000')
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
            
        ws_comp.append([loc, cod, desc, qtd])
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
            elif tipo_filtro == "ZERADOS":
                if not tem_na_comptar:
                    incluir = True
            elif tipo_filtro == "FALTAS":
                if tem_na_comptar and divergencia_val < 0:
                    incluir = True
            elif tipo_filtro == "SOBRAS":
                if tem_na_comptar and divergencia_val > 0:
                    incluir = True
                    
            if incluir:
                itens_filtrados_temp.append((cod, desc, est, tem_na_comptar, soma_comptar_val, divergencia_val, locs_prod))

    locais_com_contagem = [len(l[6]) for l in itens_filtrados_temp if len(l[6]) > 0]
    max_locais_encontrados = max(locais_com_contagem) if locais_com_contagem else 1

    # 3. Aba Saldo Loja (ou Itens Sem Saldo)
    nome_aba = "Itens Sem Saldo" if tipo_filtro == "SEM_SALDO" else "Saldo Loja"
    ws_saldo = wb.create_sheet(title=nome_aba)
    ws_saldo.views.sheetView[0].showGridLines = True
    
    ws_saldo.page_setup.orientation = ws_saldo.ORIENTATION_LANDSCAPE
    ws_saldo.page_setup.paperSize = ws_saldo.PAPERSIZE_A4
    ws_saldo.sheet_properties.pageSetUpPr.fitToPage = True
    ws_saldo.page_setup.fitToWidth = 1
    ws_saldo.page_setup.fitToHeight = 0
    
    # Configuração rígida do rodapé: Esquerda = Info da Loja | Centro = Vazio | Direita = Paginação + Imagem (&G)
    ws_saldo.oddFooter.left.text = f"Loja: {numero_loja_input} | Grupo KMB"
    ws_saldo.oddFooter.center.text = ""
    ws_saldo.oddFooter.right.text = "Página &P de &N  &G"
    
    caminho_brasao_rodape = "brasao_sem_fundo.png" if os.path.exists("brasao_sem_fundo.png") else ("1000352084.jpg" if os.path.exists("1000352084.jpg") else None)
    if caminho_brasao_rodape:
        try:
            ws_saldo.oddFooter.right.picture = caminho_brasao_rodape
        except Exception:
            pass

    # Logotipo no topo
    caminho_logo_topo = "1000352084.jpg" if os.path.exists("1000352084.jpg") else None
    if caminho_logo_topo:
        try:
            img_topo = OpenpyxlImage(caminho_logo_topo)
            img_topo.width = 200
            img_topo.height = 35
            ws_saldo.add_image(img_topo, 'B1')
        except Exception:
            pass

    # Posicionamento rigoroso e limpo dos dados do topo (Data e Loja) nas colunas D e E
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
    
    # Limpeza absoluta das células Q2 e R2 (evitando qualquer texto residual nas pontas)
    ws_saldo['Q2'] = ""
    ws_saldo['R2'] = ""
    
    headers_saldo = ["Codigo", "Descrição", "Estoque", "COMPTAR", "DIVERGÊNCIA"]
    for i in range(1, max_locais_encontrados + 1):
        headers_saldo.append(str(i) if i > 1 else "CONT. 1")
    for i in range(1, max_locais_encontrados + 1):
        headers_saldo.append(f"Local {i}" if i == 1 else str(i))
        
    ws_saldo.append(headers_saldo)
    ws_saldo.print_title_rows = '3:3'
    ws_saldo.auto_filter.ref = f"A3:{get_column_letter(len(headers_saldo))}{ws_saldo.max_row}"
    
    for col_num in range(1, len(headers_saldo) + 1):
        cell = ws_saldo.cell(row=3, column=col_num)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borda_fina
        
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
            contagens_vals = ["0 - Comptar"] + [""] * (max_locais_encontrados - 1)
            locais_vals = ["Sem local"] + [""] * (max_locais_encontrados - 1)
            
        linha_dados.extend(contagens_vals)
        linha_dados.extend(locais_vals)
        ws_saldo.append(linha_dados)
        
    for row in ws_saldo.iter_rows(min_row=4, max_row=ws_saldo.max_row, min_col=1, max_col=len(headers_saldo)):
        for cell in row:
            cell.font = fonte_padrao
            cell.border = borda_fina
            if cell.row >= 4 and cell.column >= 3:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '0'
            
            if cell.column in [1, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif cell.column == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Ajuste dinâmico e enxuto das larguras de coluna
    for col in ws_saldo.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= 3 and cell.value is not None:
                val_str = str(cell.value)
                if not val_str.startswith("="):
                    max_len = max(max_len, len(val_str))
        ws_saldo.column_dimensions[col_letter].width = max(max_len + 2, 7)

    for ws in [ws_loja, ws_comp]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    val_str = str(cell.value)
                    if not val_str.startswith("="):
                        max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


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
    p_assinatura.add_run("_____________________________________________________\n").bold = True
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
    <b>_____________________________________________________</b><br/>
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
    .main-header { padding: 20px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px; }
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
    
    # ====================================================================
    # MODO 2: PLANILHA DE CONFERÊNCIA COM PRÉ-VISUALIZAÇÃO LIMPA E RODAPÉ DE IMPRESSÃO
    # ====================================================================
    if modo_operacao == "Modo Planilha de Conferência":
        st.markdown("### 📋 MODO PLANILHA DE CONFERÊNCIA // GERAÇÃO AUTOMATIZADA")
        st.markdown("Insira o número da loja e carregue os arquivos abaixo para visualizar e gerar o plano de inventário unificado:")
        
        numero_loja_ui = st.number_input("Número da Loja (ex: 13016 ou 17315)", min_value=1, max_value=99999, value=13016, step=1)
        
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            arq_loc = st.file_uploader("1. Arquivo de Itens por Localização (ex: Itens_localizacao.csv)", type=['csv', 'xlsx', 'xls'], key="up_loc")
        with col_up2:
            arq_est = st.file_uploader("2. Arquivo de Saldo de Estoque (ex: saldo-estoque.csv)", type=['csv', 'xlsx', 'xls'], key="up_est")
            
        if arq_loc and arq_est:
            df_loc_temp, msg1 = processar_arquivo_upload(arq_loc)
            df_est_temp, msg2 = processar_arquivo_upload(arq_est)
            
            if df_loc_temp is not None and df_est_temp is not None:
                st.success("✅ Arquivos carregados e validados com sucesso!")
                
                # PRÉ-VISUALIZAÇÃO NA TELA COM CABEÇALHOS 100% LIMPOS
                st.markdown("---")
                st.markdown("### 👁️ Pré-visualização na Tela (Aba Saldo Loja)")
                st.caption("Abaixo está a prévia interativa com colunas limpas, sem pontos decimais nos cabeçalhos:")
                
                buf_preview = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "TODOS")
                df_preview_exibir = pd.read_excel(buf_preview, sheet_name="Saldo Loja", skiprows=2).fillna("")
                
                novas_cols_preview = []
                for col in df_preview_exibir.columns:
                    col_str = str(col).strip()
                    if col_str.endswith('.0'):
                        col_str = col_str[:-2]
                    novas_cols_preview.append(col_str)
                df_preview_exibir.columns = novas_cols_preview
                
                for col in df_preview_exibir.columns:
                    df_preview_exibir[col] = df_preview_exibir[col].apply(
                        lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else (str(x).replace('.0', '') if isinstance(x, float) else x)
                    )
                
                st.dataframe(df_preview_exibir, use_container_width=True, height=420)
                
                st.markdown("---")
                st.markdown("### 📥 Listas de Conferência para Impressão Cirúrgica")
                st.caption("Configuradas com A4 Paisagem, largura blindada, novo logotipo no topo e brasão clássico no rodapé de impressão.")
                
                c_imp1, c_imp2, c_imp3, c_imp4 = st.columns(4)
                
                with c_imp1:
                    st.markdown("##### 1️⃣ Zerados")
                    buf_zerados = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "ZERADOS")
                    st.download_button(
                        label="📥 Baixar (Zerados)",
                        data=buf_zerados,
                        file_name=f"Impressao_Zerados_Loja_{numero_loja_ui}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                with c_imp2:
                    st.markdown("##### 2️⃣ Faltas")
                    buf_faltas = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "FALTAS")
                    st.download_button(
                        label="📥 Baixar (Faltas)",
                        data=buf_faltas,
                        file_name=f"Impressao_Faltas_Loja_{numero_loja_ui}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                with c_imp3:
                    st.markdown("##### 3️⃣ Sobras")
                    buf_sobras = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SOBRAS")
                    st.download_button(
                        label="📥 Baixar (Sobras)",
                        data=buf_sobras,
                        file_name=f"Impressao_Sobras_Loja_{numero_loja_ui}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with c_imp4:
                    st.markdown("##### 4️⃣ Sem Saldo")
                    buf_sem_saldo = gerar_planilha_filtrada_kmb(df_loc_temp, df_est_temp, numero_loja_ui, "SEM_SALDO")
                    st.download_button(
                        label="📥 Baixar (Sem Saldo)",
                        data=buf_sem_saldo,
                        file_name=f"Impressao_ItensContadosSemSaldo_Loja_{numero_loja_ui}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                st.error("❌ Erro ao processar um ou ambos os arquivos. Verifique o formato.")
        return

    # ====================================================================
    # MODO 1: NOVA CORE ORIGINAL (5 ETAPAS COMPLETA)
    # ====================================================================
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
            if len(lista_dfs) == 1:
                df_filtrado = lista_dfs[0]
                num_cols = df_filtrado.select_dtypes(include=['number']).columns.tolist()
                cat_cols = df_filtrado.select_dtypes(include=['object']).columns.tolist()
                all_cols = df_filtrado.columns.tolist()
                
                resumo_unico = calcular_metricas_gerenciais(lista_dfs, nomes_arquivos)[0]
                
                # ETAPA 1
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
                
                # ETAPA 2
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

                # ETAPA 3
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

                # ETAPA 4
                with st.expander("🛡️ ETAPA 4: Matriz de Risco & Acurácia Ponderada por Impacto", expanded=False):
                    if len(num_cols) >= 2 and len(all_cols) > 0:
                        df_risco = df_filtrado.copy()
                        col_val_aju_etapa4 = 'Valor de Ajuste (R$)' if 'Valor de Ajuste (R$)' in df_risco.columns else next((c for c in num_cols if 'ajuste' in c.lower() and 'valor' in c.lower()), num_cols[0])
                        df_risco['Categoria_Prod'] = df_risco['Descrição do Produto'].astype(str).str.split().str[0]
                        df_agrupado_val = df_risco.groupby('Categoria_Prod')[col_val_aju_etapa4].sum().reset_index().sort_values(by=col_val_aju_etapa4, ascending=True).head(15)
                        fig_risco = px.bar(df_agrupado_val, x='Categoria_Prod', y=col_val_aju_etapa4, title="Balanço Líquido de Divergências por Categoria (R$)", template="plotly_dark", color=col_val_aju_etapa4, color_continuous_scale=['#ef4444', '#3b82f6', '#10b981'])
                        fig_risco.update_layout(paper_bgcolor='rgba(11, 15, 25, 0.9)', plot_bgcolor='rgba(11, 15, 25, 0.9)', height=420)
                        st.plotly_chart(fig_risco, use_container_width=True)

                # ETAPA 5
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

            # Central de Relatórios
            st.markdown("---")
            st.subheader("📥 Central de Relatórios Executivos C-Level")
            col_e1, col_e2, col_e3, col_e4 = st.columns(4)
            with col_e1: st.download_button("📊 Excel Gerencial", data=gerar_excel_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_e2: st.download_button("📄 Word Gerencial", data=gerar_docx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            with col_e3: st.download_button("📑 PDF Gerencial", data=gerar_pdf_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Relatorio_GrupoKMB.pdf", mime="application/pdf")
            with col_e4: st.download_button("📊 Apresentação PPTX", data=gerar_pptx_gerencial_avancado(lista_dfs, nomes_arquivos), file_name="Apresentacao_GrupoKMB.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if __name__ == "__main__":
    main()
