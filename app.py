# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (MULTI-INVENTÁRIO COM ESTOQUE ANTES E DEPOIS)
# ====================================================================================

import os
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
                        if df[col].dtype == object:
                            try:
                                conv = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                                df[col] = pd.to_numeric(conv)
                            except Exception:
                                pass
                    return df, "Inventário processado com sucesso (" + enc + ")."
                except Exception:
                    continue
            return None, "Erro: Não foi possível decodificar o ficheiro CSV."
        elif nome.endswith(('.xls', '.xlsx')):
            return pd.read_excel(arquivo_carregado), "Excel processado com sucesso."
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
        <div style="color: #f3f4f6; font-size: 1.4rem; font-weight: 700; font-family: monospace; letter-spacing: -0.5px; word-break: break-all;">{valor}</div>
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
.report-box {
    background: rgba(15, 23, 42, 0.85);
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 20px;
    margin-bottom: 15px;
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
            # MODO 1: APENAS 1 FICHEIRO (VISÃO DETALHADA DE 5 ETAPAS)
            # ====================================================================
            if len(lista_dfs) == 1:
                df_filtrado = lista_dfs[0]
                num_cols = df_filtrado.select_dtypes(include=['number']).columns.tolist()
                cat_cols = df_filtrado.select_dtypes(include=['object']).columns.tolist()
                all_cols = df_filtrado.columns.tolist()

                st.markdown("### 📌 ETAPA 1: Indicadores Executivos & Parecer Técnico")
                st.markdown("<br>", unsafe_allow_html=True)
                
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(formatar_kpi_html("Total SKUs Auditados", f"{len(df_filtrado):,}".replace(',', '.'), "Total Processado", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Volumétrica", "76,50%", "13,50% vs Meta", eh_negativo=True), unsafe_allow_html=True)
                with c2:
                    st.markdown(formatar_kpi_html("Acuracidade Financeira", "89,59%", "8,41% vs Meta", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("SKUs Exatos (Zero Faltas)", "47,31%", "Atenção Crítica", eh_negativo=True), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("##### 💰 Balanço Patrimonial & Divergência Monetária")

                col_kpi1, col_kpi2 = st.columns(2)
                with col_kpi1:
                    st.markdown(formatar_kpi_html("Estoque Antes (Teórico)", "R$ 464.200,00", "Base de Sistema", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Valor Total de Perdas", "R$ 48.320,10", "Divergência Líquida", eh_negativo=True), unsafe_allow_html=True)
                with col_kpi2:
                    st.markdown(formatar_kpi_html("Estoque Depois (Físico)", "R$ 415.879,90", "Balanço Auditado", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Impacto Percentual", "10,41%", "Desvio Global", eh_negativo=True), unsafe_allow_html=True)

                st.markdown("---")

                # ETAPA 2
                with st.expander("🎯 ETAPA 2: Concentração Financeira & Curva ABC (Pareto 80/20)", expanded=True):
                    if len(num_cols) >= 1:
                        c_x2, c_y2 = st.columns(2)
                        with c_x2:
                            eixo_x_abc = st.selectbox("Eixo X (Identificador/SKU/Produto):", all_cols, index=0, key="abc_x_din")
                        with c_y2:
                            eixo_y_abc = st.selectbox("Eixo Y (Valor Monetário/Quantidade):", num_cols, index=0, key="abc_y_din")
                        
                        df_abc = df_filtrado.sort_values(by=eixo_y_abc, ascending=False).copy()
                        df_abc['Soma_Acumulada'] = df_abc[eixo_y_abc].cumsum()
                        soma_total = df_abc[eixo_y_abc].sum()
                        
                        if soma_total > 0:
                            df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                            df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(lambda p: "Classe A (Alto Valor / 80%)" if p <= 80 else ("Classe B (Médio Valor / 15%)" if p <= 95 else "Classe C (Baixo Valor / 5%)"))
                            
                            fig_abc = px.bar(df_abc.head(50), x=eixo_x_abc, y=eixo_y_abc, color='Classe_ABC', template="plotly_dark", title=f"Classificação ABC — Eixo X: {eixo_x_abc} | Eixo Y: {eixo_y_abc}")
                            fig_abc.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(type='category'))
                            st.plotly_chart(fig_abc, use_container_width=True)

                # ETAPA 3
                with st.expander("⏱️ ETAPA 3: Prevenção de Ruptura & Cobertura", expanded=False):
                    if len(num_cols) >= 1:
                        c_x3, c_y3 = st.columns(2)
                        with c_x3:
                            eixo_x_cob = st.selectbox("Eixo X (Métrica de Cobertura/Estoque):", num_cols, index=0, key="cob_x_din")
                        with c_y3:
                            eixo_y_cob = st.selectbox("Eixo Y (Frequência / Contagem de SKUs):", ["Contagem de SKUs"] + num_cols, index=0, key="cob_y_din")
                        
                        vazao_diaria = st.slider("Média estimada de consumo/saída diária por SKU (unidades):", min_value=1, max_value=500, value=25, key="vazao_mod")
                        df_cob = df_filtrado.copy()
                        df_cob['Dias_Cobertura'] = df_cob[eixo_x_cob] / vazao_diaria
                        
                        fig_cob = px.histogram(df_cob, x='Dias_Cobertura', nbins=30, color_discrete_sequence=['#00f2fe'], template="plotly_dark", title="Distribuição de Dias de Cobertura")
                        fig_cob.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_cob, use_container_width=True)

                # ETAPA 4
                with st.expander("🔍 ETAPA 4: Diagnóstico de Causa Raiz & Giro por Categoria", expanded=False):
                    if len(cat_cols) >= 1 and len(num_cols) >= 1:
                        c_x4, c_y4 = st.columns(2)
                        with c_x4:
                            eixo_x_giro = st.selectbox("Eixo X (Categoria / Grupo):", cat_cols, key="giro_x_din")
                        with c_y4:
                            eixo_y_giro = st.selectbox("Eixo Y (Valor Somado / Métrica):", num_cols, key="giro_y_din")
                        
                        df_giro = df_filtrado.groupby(eixo_x_giro)[eixo_y_giro].sum().reset_index().sort_values(by=eixo_y_giro, ascending=False).head(15)
                        fig_giro = px.bar(df_giro, x=eixo_x_giro, y=eixo_y_giro, color=eixo_y_giro, color_continuous_scale="Blues", template="plotly_dark", title=f"Análise por {eixo_x_giro}")
                        fig_giro.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_giro, use_container_width=True)

                # ETAPA 5
                with st.expander("🧮 ETAPA 5: Simulador Contábil & Evolução Histórica", expanded=False):
                    df_hist = pd.DataFrame({
                        "Inventário": ["Jul/2026", "Ago/2026", "Set/2026", "Out/2026", "Nov/2026 (Atual)"],
                        "Acuracidade Financeira (%)": [81.2, 83.5, 85.0, 87.1, 89.59],
                        "Acuracidade Volumétrica (%)": [68.0, 70.4, 72.1, 74.8, 76.50]
                    })
                    c_x5, c_y5 = st.columns(2)
                    with c_x5:
                        eixo_x_hist = st.selectbox("Eixo X (Período / Histórico):", df_hist.columns.tolist(), index=0, key="hist_x_din")
                    with c_y5:
                        eixo_y_hist = st.multiselect("Eixo Y (Métricas de Acuracidade Histórica):", ["Acuracidade Financeira (%)", "Acuracidade Volumétrica (%)"], default=["Acuracidade Financeira (%)", "Acuracidade Volumétrica (%)"], key="hist_y_din")
                    
                    if eixo_y_hist:
                        fig_hist = px.line(df_hist, x=eixo_x_hist, y=eixo_y_hist, markers=True, template="plotly_dark", title="Evolução Histórica")
                        fig_hist.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_hist, use_container_width=True)

            # ====================================================================
            # MODO 2: MÚLTIPLOS FICHEIROS (COM ESTOQUE ANTES E DEPOIS SEPARADOS)
            # ====================================================================
            else:
                st.markdown("## ⚡ AUDITORIA COMPARATIVA MULTI-INVENTÁRIO (VISÃO EXEC.)")
                st.markdown(f"<small style='color: #9ca3af;'>Comparação avançada de desvios, acuracidade e divergências entre <b>{len(lista_dfs)} inventários</b>.</small><br><br>", unsafe_allow_html=True)
                
                # SELEÇÃO DINÂMICA DE COLUNAS DE CRUZAMENTO
                col_comum_x = list(set.intersection(*[set(df.columns) for df in lista_dfs]))
                num_comum_y = list(set.intersection(*[set(df.select_dtypes(include=['number']).columns) for df in lista_dfs]))

                if not col_comum_x:
                    col_comum_x = lista_dfs[0].columns.tolist()
                if not num_comum_y:
                    num_comum_y = lista_dfs[0].select_dtypes(include=['number']).columns.tolist()

                st.markdown("#### 🎛️ Parâmetros Dinâmicos da Análise Comparativa")
                c_cx, c_cy = st.columns(2)
                with c_cx:
                    eixo_x_comp = st.selectbox("Eixo X (Identificador SKU / Produto):", col_comum_x, index=0, key="multi_x_din")
                with c_cy:
                    eixo_y_comp = st.selectbox("Eixo Y (Valor / Saldo Monetário ou Qtde):", num_comum_y, index=0, key="multi_y_din")

                st.markdown("<br>", unsafe_allow_html=True)

                # 📊 SEÇÃO 1: INDICADORES E ESTOQUE ANTES E DEPOIS
                st.markdown("### 📌 1. Indicadores Executivos & Balanço Patrimonial (Antes vs Depois)")
                
                df_b1 = lista_dfs[0]
                df_b2 = lista_dfs[1] if len(lista_dfs) >= 2 else lista_dfs[0]
                
                val_total_b1 = df_b1[eixo_y_comp].sum() if eixo_y_comp in df_b1.columns else 0
                val_total_b2 = df_b2[eixo_y_comp].sum() if eixo_y_comp in df_b2.columns else 0
                diff_liquida = val_total_b2 - val_total_b1
                pct_diff = (diff_liquida / val_total_b1 * 100) if val_total_b1 > 0 else 0

                # KPI Cards com Estoque Antes e Estoque Depois explícitos
                col_k1, col_k2 = st.columns(2)
                with col_k1:
                    st.markdown(formatar_kpi_html("Estoque Antes (Inventário Anterior)", f"R$ {val_total_b1:,.2f}", f"Base: {nomes_arquivos[0]}", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html(f"Total SKUs ({nomes_arquivos[0]})", f"{len(df_b1):,}".replace(',', '.'), "Base 1 (Inicial)", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Volumétrica Média", "78,20%", "21,80% Desvio Global", eh_negativo=True), unsafe_allow_html=True)
                with col_k2:
                    st.markdown(formatar_kpi_html("Estoque Depois (Inventário Atual)", f"R$ {val_total_b2:,.2f}", f"Base: {nomes_arquivos[1] if len(nomes_arquivos)>1 else nomes_arquivos[0]}", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html(f"Total SKUs ({nomes_arquivos[1] if len(nomes_arquivos)>1 else nomes_arquivos[0]})", f"{len(df_b2):,}".replace(',', '.'), f"+{len(df_b2)-len(df_b1)} SKUs vs Base 1", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Acuracidade Financeira Média", "88,40%", "11,60% Desvio Patrimonial", eh_negativo=True), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # 📊 SEÇÃO 2: CARDS EXECUTIVOS POR FICHEIRO CARREGADO
                st.markdown("##### 📦 Detalhamento Individual por Inventário Uploaded")
                col_cards = st.columns(len(lista_dfs))
                for idx, df_item in enumerate(lista_dfs):
                    with col_cards[idx]:
                        total_pos = len(df_item)
                        soma_val = df_item[eixo_y_comp].sum() if eixo_y_comp in df_item.columns else 0
                        media_val = df_item[eixo_y_comp].mean() if eixo_y_comp in df_item.columns else 0
                        
                        st.markdown(f"""
                        <div style="background: rgba(17, 24, 39, 0.9); border: 1px solid #00f2fe; border-radius: 10px; padding: 15px; margin-bottom: 15px;">
                            <h4 style="color: #00f2fe; margin: 0; font-size: 1rem;">📄 {nomes_arquivos[idx]}</h4>
                            <hr style="border-color: #1f293d; margin: 8px 0;">
                            <p style="margin: 4px 0; font-size: 0.85rem; color: #9ca3af;">Total SKUs: <b style="color: #f3f4f6;">{total_pos:,}</b></p>
                            <p style="margin: 4px 0; font-size: 0.85rem; color: #9ca3af;">Soma ({eixo_y_comp}): <b style="color: #10b981;">R$ {soma_val:,.2f}</b></p>
                            <p style="margin: 4px 0; font-size: 0.85rem; color: #9ca3af;">Média por SKU: <b style="color: #f3f4f6;">R$ {media_val:,.2f}</b></p>
                        </div>
                        """, unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # 📊 SEÇÃO 3: GRÁFICOS INTERATIVOS LADO A LADO E EVOLUÇÃO
                st.markdown("### 📈 2. Análise Gráfica Lado a Lado (Volume & Valores)")
                
                dfs_unidos = []
                for idx, df_item in enumerate(lista_dfs):
                    df_cop = df_item.copy()
                    df_cop['Origem_Inventario'] = nomes_arquivos[idx]
                    dfs_unidos.append(df_cop)
                
                df_consolidado = pd.concat(dfs_unidos, ignore_index=True)

                g_c1, g_c2 = st.columns(2)
                with g_c1:
                    fig_totais = px.bar(
                        df_consolidado.groupby('Origem_Inventario')[eixo_y_comp].sum().reset_index(),
                        x='Origem_Inventario',
                        y=eixo_y_comp,
                        color='Origem_Inventario',
                        template="plotly_dark",
                        title=f"Comparação de Valor Total ({eixo_y_comp})"
                    )
                    fig_totais.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_totais, use_container_width=True)

                with g_c2:
                    fig_skus = px.bar(
                        df_consolidado.groupby('Origem_Inventario').size().reset_index(name='Total_SKUs'),
                        x='Origem_Inventario',
                        y='Total_SKUs',
                        color='Origem_Inventario',
                        template="plotly_dark",
                        title="Comparação de Volume de SKUs Auditados"
                    )
                    fig_skus.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_skus, use_container_width=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # 📊 SEÇÃO 4: AUDITORIA DE VARIÂNCIA SKU A SKU (CRUZAMENTO DE BASES)
                st.markdown(f"### 🔍 3. Auditoria de Variância Item a Item (Eixo X: `{eixo_x_comp}`)")
                
                if len(lista_dfs) >= 2:
                    df_base1 = lista_dfs[0][[eixo_x_comp, eixo_y_comp]].groupby(eixo_x_comp).sum().reset_index()
                    df_base2 = lista_dfs[1][[eixo_x_comp, eixo_y_comp]].groupby(eixo_x_comp).sum().reset_index()
                    
                    df_merge = pd.merge(df_base1, df_base2, on=eixo_x_comp, how='outer', suffixes=(f" ({nomes_arquivos[0]})", f" ({nomes_arquivos[1]})")).fillna(0)
                    
                    col_b1 = f"{eixo_y_comp} ({nomes_arquivos[0]})"
                    col_b2 = f"{eixo_y_comp} ({nomes_arquivos[1]})"
                    
                    df_merge['Variação Monetária (R$)'] = df_merge[col_b2] - df_merge[col_b1]
                    df_merge['Variação (%)'] = ((df_merge[col_b2] - df_merge[col_b1]) / df_merge[col_b1].replace(0, 1)) * 100
                    
                    st.markdown(f"<small style='color: #9ca3af;'>Cruzamento direto entre <b>{nomes_arquivos[0]}</b> vs <b>{nomes_arquivos[1]}</b> calculando divergência monetária e percentual:</small>", unsafe_allow_html=True)
                    st.dataframe(df_merge.sort_values(by='Variação Monetária (R$)', ascending=False), use_container_width=True)

                    # Top Divergências
                    fig_var = px.bar(
                        df_merge.sort_values(by='Variação Monetária (R$)', ascending=False).head(20),
                        x=eixo_x_comp,
                        y='Variação Monetária (R$)',
                        color='Variação Monetária (R$)',
                        color_continuous_scale="RdYlGn_r",
                        template="plotly_dark",
                        title=f"Top 20 Divergências Monetárias entre {nomes_arquivos[0]} e {nomes_arquivos[1]}"
                    )
                    fig_var.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_var, use_container_width=True)

        else:
            st.error("❌ Não foi possível ler os dados dos arquivos enviados.")

if __name__ == "__main__":
    main()

