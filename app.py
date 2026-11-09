# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (VERSÃO 100% CORRIGIDA & MODULAR)
# ====================================================================================

import os
import io
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

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
                    return df, "Inventário processado com sucesso."
                except Exception:
                    continue
            return None, "Erro: Não foi possível decodificar o ficheiro CSV."
        elif nome.endswith(('.xls', '.xlsx')):
            return pd.read_excel(arquivo_carregado), "Excel processado com sucesso."
        else:
            return None, "Formato não suportado."
    except Exception as e:
        return None, "Erro ao processar ficheiro: " + str(e)

def limpar_serie_numerica(serie):
    if serie is None:
        return pd.Series(0.0, index=range(10))
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors='coerce').fillna(0.0)
    try:
        limpo = serie.astype(str).str.replace('R$', '', regex=False).str.replace('.', '', regex=False).str.replace(',', '.', regex=False).str.strip()
        return pd.to_numeric(limpo, errors='coerce').fillna(0.0)
    except Exception:
        return pd.to_numeric(serie, errors='coerce').fillna(0.0)

def auto_detectar_coluna_valor(colunas):
    colunas_str = [str(c) for c in colunas]
    for termo in ['valor de estoque', 'estoque atual', 'valor total', 'custo total', 'total', 'valor']:
        for c in colunas_str:
            if termo in c.lower() and 'ajuste' not in c.lower() and 'codigo' not in c.lower() and 'código' not in c.lower():
                return c
    return colunas_str[0]

def auto_detectar_coluna(colunas, palavras_chave):
    colunas_str = [str(c) for c in colunas]
    for palavra in palavras_chave:
        for c in colunas_str:
            if palavra in c.lower() and 'codigo' not in c.lower() and 'código' not in c.lower():
                return c
    return colunas_str[0]

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

def main():
    st.set_page_config(page_title="NOVA-CORE // J.A.R.V.I.S. EXECUTIVO", page_icon="⚡", layout="wide")

    st.markdown("""
    <style>
    .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
    .jarvis-title {
        font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
        background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .bento-card {
        background: rgba(17, 24, 39, 0.7); border: 1px solid rgba(31, 41, 55, 0.8);
        border-top: 3px solid #00f2fe; border-radius: 10px; padding: 18px; margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="padding: 20px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px;">
        <div class="jarvis-title">⚡ NOVA-CORE // DIRETORIA EXECUTIVA</div>
        <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">AUDITORIA DE ESTOQUE & ANÁLISE COMPARATIVA MULTI-INVENTÁRIO</div>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("### 🎛️ CONTROLES & NAVEGAÇÃO")
    arquivos_submetidos = st.sidebar.file_uploader("📂 Ingestão de Ficheiros", type=['csv', 'xlsx', 'xls'], accept_multiple_files=True)
    
    modo = st.sidebar.radio("Selecione o Modo:", ["Modo 1: Indicadores Executivos & Balanço", "Modo 2: Curva ABC & Concentração Financeira"])

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
            todas_colunas = df_filtrado.columns.tolist()

            col_val_sel = auto_detectar_coluna_valor(todas_colunas)
            col_qtd_sel = auto_detectar_coluna(todas_colunas, ['quantidade', 'fisico', 'estoque', 'atual', 'qtd'])
            col_aj_sel = auto_detectar_coluna(todas_colunas, ['ajuste', 'diferenca', 'dif', 'desvio'])
            col_custo_sel = auto_detectar_coluna(todas_colunas, ['custo', 'preco', 'preço', 'unitario'])

            total_pos = len(df_filtrado)
            serie_val = limpar_serie_numerica(df_filtrado[col_val_sel]) if col_val_sel in df_filtrado.columns else pd.Series(0.0, index=df_filtrado.index)
            serie_qtd = limpar_serie_numerica(df_filtrado[col_qtd_sel]) if col_qtd_sel in df_filtrado.columns else pd.Series(0.0, index=df_filtrado.index)
            serie_ajuste = limpar_serie_numerica(df_filtrado[col_aj_sel]) if col_aj_sel in df_filtrado.columns else pd.Series(0.0, index=df_filtrado.index)

            soma_depois = serie_val.sum()
            total_itens_pos = serie_qtd.sum()
            soma_ajuste_qtd = serie_ajuste.sum()
            total_itens_antes = total_itens_pos - soma_ajuste_qtd

            df_falta = df_filtrado[serie_ajuste < 0]
            df_sobra = df_filtrado[serie_ajuste > 0]
            
            qtd_falta = len(df_falta)
            qtd_sobra = len(df_sobra)
            vol_falta_un = abs(limpar_serie_numerica(df_falta[col_aj_sel])).sum() if not df_falta.empty else 0.0
            vol_sobra_un = limpar_serie_numerica(df_sobra[col_aj_sel]).sum() if not df_sobra.empty else 0.0

            if col_custo_sel in df_filtrado.columns and col_aj_sel in df_filtrado.columns:
                val_falta_fin = abs((limpar_serie_numerica(df_falta[col_aj_sel]) * limpar_serie_numerica(df_falta[col_custo_sel]))).sum() if not df_falta.empty else 0.0
                val_sobra_fin = (limpar_serie_numerica(df_sobra[col_aj_sel]) * limpar_serie_numerica(df_sobra[col_custo_sel])).sum() if not df_sobra.empty else 0.0
                diferenca_monetaria = val_sobra_fin - val_falta_fin
            else:
                val_falta_fin = abs(limpar_serie_numerica(df_falta[col_aj_sel])).sum() * (soma_depois / total_itens_pos if total_itens_pos > 0 else 1.0) if not df_falta.empty else 0.0
                val_sobra_fin = limpar_serie_numerica(df_sobra[col_aj_sel]).sum() * (soma_depois / total_itens_pos if total_itens_pos > 0 else 1.0) if not df_sobra.empty else 0.0
                diferenca_monetaria = val_sobra_fin - val_falta_fin

            soma_absoluta_val = val_falta_fin + val_sobra_fin
            soma_antes = soma_depois - diferenca_monetaria
            impacto_perc = (diferenca_monetaria / soma_antes) * 100 if soma_antes > 0 else 0.0
            acur_vol_real = ((total_pos - (qtd_falta + qtd_sobra)) / total_pos * 100) if total_pos > 0 else 0.0
            acur_fin_real = max(0.0, min(100.0, 100.0 - abs(impacto_perc) * 2.5))

            if modo == "Modo 1: Indicadores Executivos & Balanço":
                st.markdown("### 📌 MODO 1: Indicadores Executivos & Parecer Técnico")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{total_pos:,}".replace(',', '.'), f"Pré: {total_itens_antes:,.0f} un | Pós: {total_itens_pos:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{qtd_falta:,} SKUs", f"Vol: -{vol_falta_un:,.0f} un | R$ {val_falta_fin:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                with c2:
                    st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{qtd_sobra:,} SKUs", f"Vol: +{vol_sobra_un:,.0f} un | R$ +{val_sobra_fin:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Volumétrica / Financeira", f"Vol: {acur_vol_real:.2f}% | Fin: {acur_fin_real:.2f}%".replace('.', ','), "Meta Corporativa 95%", eh_positivo=(acur_vol_real >= 95)), unsafe_allow_html=True)

                st.markdown("##### 💰 Balanço Patrimonial & Divergência Monetária")
                bc1, bc2 = st.columns(2)
                with bc1:
                    st.markdown(formatar_kpi_html("Estoque Antes (Sistema / Teórico)", f"R$ {soma_antes:,.2f}", f"Qtd Pré: {total_itens_antes:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)
                with bc2:
                    st.markdown(formatar_kpi_html("Estoque Depois (Físico Auditado)", f"R$ {soma_depois:,.2f}", f"Qtd Pós: {total_itens_pos:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)

            elif modo == "Modo 2: Curva ABC & Concentração Financeira":
                st.markdown("### 📊 MODO 2: Concentração Financeira & Curva ABC (Pareto 80/20)")
                c_x2, c_y2 = st.columns(2)
                with c_x2:
                    eixo_x_abc = st.selectbox("Eixo X (Identificador / SKU / Produto):", todas_colunas, index=0)
                with c_y2:
                    eixo_y_abc = st.selectbox("Eixo Y (Valor Monetário / Quantidade):", todas_colunas, index=todas_colunas.index(col_val_sel) if col_val_sel in todas_colunas else 0)
                
                df_abc = df_filtrado.copy()
                df_abc['Val_Num_ABC'] = limpar_serie_numerica(df_abc[eixo_y_abc]) if eixo_y_abc in df_abc.columns else pd.Series(0.0, index=df_abc.index)
                df_abc = df_abc.sort_values(by='Val_Num_ABC', ascending=False).copy()
                df_abc['Soma_Acumulada'] = df_abc['Val_Num_ABC'].cumsum()
                soma_total = df_abc['Val_Num_ABC'].sum()
                
                if soma_total > 0:
                    df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                    df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(lambda p: "Classe A (Alto Valor / 80%)" if p <= 80 else ("Classe B (Médio Valor / 15%)" if p <= 95 else "Classe C (Baixo Valor / 5%)"))
                    
                    fig_abc = px.bar(
                        df_abc.head(100), 
                        x=eixo_x_abc, 
                        y='Val_Num_ABC', 
                        color='Classe_ABC', 
                        template="plotly_dark", 
                        title=f"Classificação ABC — Eixo X: {eixo_x_abc} | Eixo Y: {eixo_y_abc}"
                    )
                    fig_abc.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(type='category'))
                    st.plotly_chart(fig_abc, use_container_width=True)
                    
                st.dataframe(df_abc[[eixo_x_abc, eixo_y_abc, '%_Acumulado', 'Classe_ABC']], use_container_width=True)

if __name__ == "__main__":
    main()
