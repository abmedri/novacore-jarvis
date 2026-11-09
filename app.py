# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (VERSÃO EXECUTIVA UNIVERSAL & BLINDADA)
# ====================================================================================

import os
import io
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
import plotly.express as px
import plotly.graph_objects as go

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

    st.sidebar.markdown("### 🎛️ CONTROLES & FILTROS")
    st.markdown('<div class="bento-card"><b>📂 INGESTÃO DE DADOS DE INVENTÁRIO</b><br><small style="color: #9ca3af;">Carregue 1 ou múltiplos arquivos para análise.</small></div>', unsafe_allow_html=True)
    
    arquivos_submetidos = st.file_uploader("", type=['csv', 'xlsx', 'xls', 'json', 'xml'], accept_multiple_files=True, label_visibility="collapsed")
    
    if arquivos_submetidos:
        lista_dfs = []
        nomes_arquivos = []
        
        for arq in arquivos_submetidos:
            df_temp, msg = processar_arquivo_upload(arq)
            if df_temp is not None:
                lista_dfs.append(df_temp)
                nomes_arquivos.append(arq.name)
        
        if lista_dfs:
            st.success(f"✅ {len(lista_dfs)} arquivo(s) carregado(s) e auditado(s) com sucesso!")
            
            # ====================================================================
            # MODO 1: APENAS 1 FICHEIRO (VISÃO DETALHADA 5 ETAPAS)
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

                # Identificação de colunas de quantidade de itens (Qtd Pré vs Pós)
                idx_qtd = identificar_coluna_ideal(num_cols, ['quantidade', 'saldo', 'estoque', 'fisico', 'teorico'])
                col_qtd_pos = num_cols[idx_qtd] if num_cols else None
                
                total_itens_pos = df_filtrado[col_qtd_pos].sum() if col_qtd_pos and col_qtd_pos in df_filtrado.columns else total_pos
                
                # Qtd Pré-inventario calculada ou por coluna teórica
                col_ajuste = next((c for c in num_cols if 'ajuste' in c.lower()), None)
                soma_ajuste_qtd = df_filtrado[col_ajuste].sum() if col_ajuste and col_ajuste in df_filtrado.columns else 0
                total_itens_antes = total_itens_pos - soma_ajuste_qtd

                if 'Congelado' in df_filtrado.columns and 'Custo' in df_filtrado.columns:
                    soma_antes = (df_filtrado['Congelado'] * df_filtrado['Custo']).sum()
                else:
                    soma_ajuste_val = df_filtrado['Valor de Ajuste(R$)'].sum() if 'Valor de Ajuste(R$)' in df_filtrado.columns else 0
                    soma_antes = soma_depois - soma_ajuste_val

                diferenca_monetaria = soma_depois - soma_antes
                impacto_perc = (diferenca_monetaria / soma_antes) * 100 if soma_antes > 0 else 0

                qtd_falta = (df_filtrado['Ajuste'] < 0).sum() if 'Ajuste' in df_filtrado.columns else 0
                qtd_sobra = (df_filtrado['Ajuste'] > 0).sum() if 'Ajuste' in df_filtrado.columns else 0
                qtd_exatos = (df_filtrado['Ajuste'] == 0).sum() if 'Ajuste' in df_filtrado.columns else int(total_pos * 0.5)
                total_divergencias = qtd_falta + qtd_sobra
                perc_divergencia = (total_divergencias / total_pos * 100) if total_pos > 0 else 0
                
                perc_falta_itens = (qtd_falta / total_pos * 100) if total_pos > 0 else 0
                perc_sobra_itens = (qtd_sobra / total_pos * 100) if total_pos > 0 else 0

                col_val_aju = 'Valor de Ajuste(R$)' if 'Valor de Ajuste(R$)' in df_filtrado.columns else None
                val_falta_fin = df_filtrado[df_filtrado['Ajuste'] < 0][col_val_aju].sum() if col_val_aju else 0
                val_sobra_fin = df_filtrado[df_filtrado['Ajuste'] > 0][col_val_aju].sum() if col_val_aju else 0
                soma_absoluta_val = abs(val_falta_fin) + abs(val_sobra_fin)
                perc_valor_somado = (soma_absoluta_val / soma_antes * 100) if soma_antes > 0 else 0

                acur_vol_real = (qtd_exatos / total_pos * 100) if total_pos > 0 else 0
                acur_fin_real = max(0.0, min(100.0, 100.0 - abs(impacto_perc) * 2.5))

                st.markdown("### 📌 ETAPA 1: Indicadores Executivos & Parecer Técnico")
                st.markdown("<br>", unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(formatar_kpi_html("Total SKUs / Itens Auditados", f"{total_pos:,}".replace(',', '.'), f"Pré: {total_itens_antes:,.0f} un | Pós: {total_itens_pos:,.0f} un".replace('.', ','), eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Produtos com Falta (Negativos)", f"{qtd_falta:,} ({perc_falta_itens:.2f}% dos itens)".replace('.', ','), f"R$ {val_falta_fin:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Total Falta + Sobra (Itens)", f"{total_divergencias:,} itens ({perc_divergencia:.2f}%)".replace('.', ','), "Volume Divergente", eh_negativo=True), unsafe_allow_html=True)
                with c2:
                    st.markdown(formatar_kpi_html("Produtos com Sobra (Positivos)", f"{qtd_sobra:,} ({perc_sobra_itens:.2f}% dos itens)".replace('.', ','), f"R$ +{val_sobra_fin:,.2f}".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Valor Total Falta + Sobra (Soma Absoluta)", f"R$ {soma_absoluta_val:,.2f}".replace('.', ','), f"{perc_valor_somado:.2f}% do Pré-Inventário".replace('.', ','), eh_negativo=True), unsafe_allow_html=True)
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

                st.markdown("---")

                # ETAPA 2: CURVA ABC
                with st.expander("🎯 ETAPA 2: Concentração Financeira & Curva ABC (Pareto 80/20)", expanded=True):
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

            # ====================================================================
            # MODO 2: MÚLTIPLOS FICHEIROS (VISÃO COMPARATIVA EXECUTIVA)
            # ====================================================================
            else:
                st.markdown("## ⚡ AUDITORIA COMPARATIVA MULTI-INVENTÁRIO (VISÃO EXEC.)")
                st.markdown(f"<small style='color: #9ca3af;'>Comparação avançada de desvios, acuracidade e divergências entre <b>{len(lista_dfs)} inventários</b>.</small><br><br>", unsafe_allow_html=True)
                
                col_comum_x = list(set.intersection(*[set(df.columns) for df in lista_dfs]))
                num_comum_y = list(set.intersection(*[set(df.select_dtypes(include=['number']).columns) for df in lista_dfs]))

                if not col_comum_x:
                    col_comum_x = lista_dfs[0].columns.tolist()
                if not num_comum_y:
                    num_comum_y = lista_dfs[0].select_dtypes(include=['number']).columns.tolist()

                idx_padrao_comp_y = identificar_coluna_ideal(num_comum_y, ['valor de estoque', 'valor', 'saldo'])

                # 📌 ETAPA 1: INDICADORES EXECUTIVOS & ACURACIDADE POR FICHEIRO (100% DINÂMICO)
                st.markdown("### 📌 1. Indicadores Executivos & Acuracidade por Ficheiro")
                
                col_cards = st.columns(len(lista_dfs))
                for idx, df_item in enumerate(lista_dfs):
                    total_pos = len(df_item)
                    num_cols_item = df_item.select_dtypes(include=['number']).columns.tolist()
                    idx_temp = identificar_coluna_ideal(num_cols_item, ['valor de estoque', 'valor', 'saldo'])
                    col_temp = num_cols_item[idx_temp] if num_cols_item and idx_temp < len(num_cols_item) else (num_cols_item[0] if num_cols_item else None)
                    
                    soma_val_real = df_item[col_temp].sum() if col_temp and col_temp in df_item.columns else 0
                    
                    # Quantidade total de itens (Pré vs Pós)
                    idx_qtd_item = identificar_coluna_ideal(num_cols_item, ['quantidade', 'saldo', 'estoque', 'fisico'])
                    col_qtd_item_pos = num_cols_item[idx_qtd_item] if num_cols_item else None
                    total_itens_pos = df_item[col_qtd_item_pos].sum() if col_qtd_item_pos and col_qtd_item_pos in df_item.columns else total_pos
                    
                    col_ajuste_item = next((c for c in num_cols_item if 'ajuste' in c.lower()), None)
                    soma_ajuste_qtd_item = df_item[col_ajuste_item].sum() if col_ajuste_item and col_ajuste_item in df_item.columns else 0
                    total_itens_antes = total_itens_pos - soma_ajuste_qtd_item

                    qtd_falta = (df_item['Ajuste'] < 0).sum() if 'Ajuste' in df_item.columns else 0
                    qtd_sobra = (df_item['Ajuste'] > 0).sum() if 'Ajuste' in df_item.columns else 0
                    qtd_exatos = (df_item['Ajuste'] == 0).sum() if 'Ajuste' in df_item.columns else int(total_pos * 0.5)
                    total_divergencias = qtd_falta + qtd_sobra
                    perc_divergencia = (total_divergencias / total_pos * 100) if total_pos > 0 else 0
                    
                    perc_falta_itens = (qtd_falta / total_pos * 100) if total_pos > 0 else 0
                    perc_sobra_itens = (qtd_sobra / total_pos * 100) if total_pos > 0 else 0
                    
                    if 'Congelado' in df_item.columns and 'Custo' in df_item.columns:
                        soma_val_antes = (df_item['Congelado'] * df_item['Custo']).sum()
                    else:
                        val_ajuste_col = next((c for c in num_cols_item if 'ajuste' in c.lower() and 'valor' in c.lower()), None)
                        soma_ajuste = df_item[val_ajuste_col].sum() if val_ajuste_col and val_ajuste_col in df_item.columns else 0
                        soma_val_antes = soma_val_real - soma_ajuste
                    
                    col_val_aju = 'Valor de Ajuste(R$)' if 'Valor de Ajuste(R$)' in df_item.columns else None
                    val_falta_fin = df_item[df_item['Ajuste'] < 0][col_val_aju].sum() if col_val_aju else 0
                    val_sobra_fin = df_item[df_item['Ajuste'] > 0][col_val_aju].sum() if col_val_aju else 0
                    soma_absoluta_val = abs(val_falta_fin) + abs(val_sobra_fin)
                    perc_valor_somado = (soma_absoluta_val / soma_val_antes * 100) if soma_val_antes > 0 else 0

                    diferenca_monetaria_item = soma_val_real - soma_val_antes
                    impacto_perc_glob = (diferenca_monetaria_item / soma_val_antes) * 100 if soma_val_antes > 0 else 0
                    
                    acur_fin_din = max(0.0, min(100.0, 100.0 - abs(impacto_perc_glob) * 2.5))
                    acur_vol_din = (qtd_exatos / total_pos * 100) if total_pos > 0 else 0
                    
                    cor_imp = "#ef4444" if impacto_perc_glob < 0 else "#10b981"
                    sinal_imp = "" if impacto_perc_glob < 0 else "+"
                    
                    cor_fin = "#10b981" if acur_fin_din >= 95 else "#ef4444"
                    sinal_fin = "+" if acur_fin_din >= 95 else ""
                    
                    cor_vol = "#10b981" if acur_vol_din >= 95 else "#ef4444"
                    sinal_vol = "+" if acur_vol_din >= 95 else ""
                    
                    with col_cards[idx]:
                        with st.container(border=True):
                            st.markdown(f"#### 📄 {nomes_arquivos[idx]}")
                            st.divider()
                            st.metric(label="Valor Total (Pós)", value=f"R$ {soma_val_real:,.2f}", delta="Saldo Efetivo Real", delta_color="normal")
                            st.markdown(f"**Total SKUs Auditados:** `{total_pos:,}`".replace(',', '.'))
                            st.markdown(f"**Qtd Total Itens (Pré vs Pós):** `Pré: {total_itens_antes:,.0f} | Pós: {total_itens_pos:,.0f}` un".replace('.', ','))
                            st.markdown(f"**Produtos com Falta:** <span style='color: #ef4444; font-weight: bold;'>{qtd_falta:,} ({perc_falta_itens:.2f}% dos itens)</span>".replace('.', ','), unsafe_allow_html=True)
                            st.markdown(f"**Produtos com Sobra:** <span style='color: #ef4444; font-weight: bold;'>{qtd_sobra:,} ({perc_sobra_itens:.2f}% dos itens)</span>".replace('.', ','), unsafe_allow_html=True)
                            st.markdown(f"**Total Falta + Sobra (Itens):** <span style='color: #ef4444; font-weight: bold;'>{total_divergencias:,} ({perc_divergencia:.2f}%)</span>".replace('.', ','), unsafe_allow_html=True)
                            st.markdown(f"**Valor Falta + Sobra (Soma Absoluta):** <span style='color: #ef4444; font-weight: bold;'>R$ {soma_absoluta_val:,.2f} ({perc_valor_somado:.2f}%)</span>".replace('.', ','), unsafe_allow_html=True)
                            st.markdown(f"**Acuracidade Financeira:** <span style='color: {cor_fin}; font-weight: bold;'>{sinal_fin}{acur_fin_din:.2f}%</span>".replace('.', ','), unsafe_allow_html=True)
                            st.markdown(f"**Acuracidade Volumétrica:** <span style='color: {cor_vol}; font-weight: bold;'>{sinal_vol}{acur_vol_din:.2f}%</span>".replace('.', ','), unsafe_allow_html=True)
                            st.markdown(f"**Impacto Geral:** <span style='color: {cor_imp}; font-weight: bold;'>{sinal_imp}{impacto_perc_glob:.2f}%</span>".replace('.', ','), unsafe_allow_html=True)
                            st.divider()
                            st.markdown(f"**Valor Pré-Inventário:** R$ {soma_val_antes:,.2f}")

                st.markdown("<br>", unsafe_allow_html=True)

                # 🎛️ ETAPA 4: PARÂMETROS DINÂMICOS & GRÁFICO COMPARATIVO EXECUTIVO
                with st.expander("🎛️ 4. Parâmetros Dinâmicos da Análise Comparativa (Eixos X e Y)", expanded=True):
                    st.info("💡 **Central de Inteligência Dinâmica:** Selecione os eixos abaixo para cruzar livremente qualquer dimensão e métrica dos seus ficheiros.")
                    c_cx, c_cy = st.columns(2)
                    with c_cx:
                        eixo_x_comp = st.selectbox("Eixo X (Dimensão / SKU / Categoria):", col_comum_x, index=0, key="multi_x_din")
                    with c_cy:
                        eixo_y_comp = st.selectbox("Eixo Y (Métrica Numérica / Valor / Quantidade):", num_comum_y, index=idx_padrao_comp_y, key="multi_y_din")
                    
                    if len(lista_dfs) >= 2 and eixo_x_comp and eixo_y_comp:
                        df_d1 = lista_dfs[0][[eixo_x_comp, eixo_y_comp]].groupby(eixo_x_comp).sum().reset_index()
                        df_d2 = lista_dfs[1][[eixo_x_comp, eixo_y_comp]].groupby(eixo_x_comp).sum().reset_index()
                        df_comp_din = pd.merge(df_d1, df_d2, on=eixo_x_comp, how='outer', suffixes=(f" ({nomes_arquivos[0]})", f" ({nomes_arquivos[1]})")).fillna(0)
                        
                        col_sk1 = f"{eixo_y_comp} ({nomes_arquivos[0]})"
                        col_sk2 = f"{eixo_y_comp} ({nomes_arquivos[1]})"
                        
                        df_comp_din['Variação Absoluta'] = df_comp_din[col_sk2] - df_comp_din[col_sk1]
                        df_comp_din['Tipo de Variação'] = df_comp_din['Variação Absoluta'].apply(lambda x: 'Aumento (Sobra/Positivo)' if x > 0 else ('Redução (Falta/Negativo)' if x < 0 else 'Sem Alteração'))
                        
                        st.markdown(f"##### 📊 Análise Comparativa Dinâmica: {eixo_y_comp} por {eixo_x_comp}")
                        
                        fig_din = px.bar(
                            df_comp_din.sort_values(by='Variação Absoluta', ascending=False).head(20),
                            x=eixo_x_comp,
                            y='Variação Absoluta',
                            color='Tipo de Variação',
                            color_discrete_map={
                                'Aumento (Sobra/Positivo)': '#10b981',
                                'Redução (Falta/Negativo)': '#ef4444',
                                'Sem Alteração': '#9ca3af'
                            },
                            template="plotly_dark",
                            title=f"Top 20 Variações de {eixo_y_comp} entre {nomes_arquivos[0]} e {nomes_arquivos[1]}"
                        )
                        fig_din.update_layout(
                            paper_bgcolor="rgba(0,0,0,0)",
                            plot_bgcolor="rgba(0,0,0,0)",
                            xaxis_title=eixo_x_comp,
                            yaxis_title=f"Diferença Absoluta ({eixo_y_comp})",
                            legend_title="Direção da Variação"
                        )
                        st.plotly_chart(fig_din, use_container_width=True)

        else:
            st.error("❌ Não foi possível ler os dados dos arquivos enviados.")

if __name__ == "__main__":
    main()
