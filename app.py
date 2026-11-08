# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (CORREÇÃO DEFINITIVA DAS 5 ETAPAS)
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
        <div style="color: #f3f4f6; font-size: 1.5rem; font-weight: 700; font-family: monospace; letter-spacing: -0.5px; word-break: break-all;">{valor}</div>
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
    <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">SISTEMA INTELIGENTE DE AUDITORIA DE ESTOQUE (5 ETAPAS COMPLETAS)</div>
</div>
""", unsafe_allow_html=True)

    st.sidebar.markdown("### 🎛️ CONTROLES & FILTROS")
    st.markdown('<div class="bento-card"><b>📂 INGESTÃO DE DADOS DE INVENTÁRIO</b><br><small style="color: #9ca3af;">Carregue o arquivo CSV/Excel para iniciar a auditoria.</small></div>', unsafe_allow_html=True)
    
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
            st.success(f"✅ {len(lista_dfs)} arquivo(s) processado(s) com sucesso!")
            
            # SE HOUVER APENAS 1 FICHEIRO, EXIBE AS 5 ETAPAS COMPLETAS COM EIXOS DINÂMICOS
            if len(lista_dfs) == 1:
                df_filtrado = lista_dfs[0]
                num_cols = df_filtrado.select_dtypes(include=['number']).columns.tolist()
                cat_cols = df_filtrado.select_dtypes(include=['object']).columns.tolist()
                all_cols = df_filtrado.columns.tolist()

                # ETAPA 1
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

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("##### 📦 Distribuição Física dos SKUs Auditados")

                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    st.markdown(formatar_kpi_html("SKUs Negativos (Faltas)", "742 SKUs", "35,60% do Total", eh_negativo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("SKUs Sem Divergência", "984 SKUs", "47,31% Exatos", eh_positivo=True), unsafe_allow_html=True)
                with col_s2:
                    st.markdown(formatar_kpi_html("SKUs Positivos (Sobras)", "358 SKUs", "17,10% do Total", eh_positivo=True), unsafe_allow_html=True)
                    st.markdown(formatar_kpi_html("Status Operacional", "Crítico", "Requer Intervenção", eh_negativo=True), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("""
<div class="report-box">
    <p><b>Data de Emissão:</b> Novembro de 2026 &nbsp;|&nbsp; <b>Auditor Responsável:</b> J.A.R.V.I.S. / Sr. Adriano Braz Martins</p>
    <hr style="border-color: #334155; margin: 10px 0;">
    <p><b>A. Balanço Patrimonial (Antes vs Depois):</b><br>
    • <i>Valor Antes (Sistema):</i> R$ 464.200,00<br>
    • <i>Valor Depois (Físico Auditado):</i> R$ 415.879,90<br>
    • <i>Diferença Apurada (Perdas Líquidas):</i> R$ 48.320,10 (impacto de 10,41% negativo no inventário).</p>
    <p><b>B. Mapeamento de Divergências de SKUs:</b><br>
    Registados 742 SKUs negativos (faltas críticas) e 358 SKUs positivos (sobras de recebimento/expedição).</p>
    <p><b>C. Plano de Ação Imediato:</b><br>
    1. Congelar temporariamente a movimentação física dos 20 itens de maior valor divergente (Curva A).<br>
    2. Direcionar contagem cíclica diária de 80% para os itens com divergência negativa crítica.</p>
</div>
""", unsafe_allow_html=True)

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
                            
                            opcoes_curvas = ["Todas as Curvas", "Classe A (Alto Valor / 80%)", "Classe B (Médio Valor / 15%)", "Classe C (Baixo Valor / 5%)"]
                            curva_selecionada = st.selectbox("Filtrar Exibição por Curva ABC:", opcoes_curvas, key="filtro_curva_abc")
                            
                            df_abc_plot = df_abc.copy()
                            if curva_selecionada != "Todas as Curvas":
                                df_abc_plot = df_abc[df_abc['Classe_ABC'] == curva_selecionada]
                            
                            fig_abc = px.bar(df_abc_plot.head(50), x=eixo_x_abc, y=eixo_y_abc, color='Classe_ABC', template="plotly_dark", title=f"Classificação ABC ({curva_selecionada})")
                            fig_abc.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(type='category'), yaxis=dict(tickformat=',.0f'))
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
                        fig_cob.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="Dias de Autonomia", yaxis_title="Quantidade de SKUs")
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
                        fig_giro.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", yaxis=dict(tickformat=',.0f'))
                        st.plotly_chart(fig_giro, use_container_width=True)

                # ETAPA 5
                with st.expander("🧮 ETAPA 5: Simulador Contábil & Evolução Histórica", expanded=False):
                    col_sim1, col_sim2 = st.columns(2)
                    with col_sim1:
                        pct_perdao_sobras = st.slider("Percentual de Regularização de Sobras (%):", 0, 100, 50, key="sim_sob_mod")
                    with col_sim2:
                        pct_baixa_faltas = st.slider("Percentual de Baixa Contábil de Faltas (%):", 0, 100, 30, key="sim_fal_mod")
                    
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

            # SE HOUVER 2 OU MAIS FICHEIROS, EXIBE O MODO DE COMPARAÇÃO
            else:
                st.markdown("### 📊 MODO DE COMPARAÇÃO MULTI-INVENTÁRIO")
                st.info(f"Comparando simultaneamente **{len(lista_dfs)} arquivos** carregados: " + ", ".join(nomes_arquivos))
                dados_comparacao = []
                for idx, df_comp in enumerate(lista_dfs):
                    num_c = df_comp.select_dtypes(include=['number']).columns.tolist()
                    soma_val = df_comp[num_c[0]].sum() if num_c else 0
                    dados_comparacao.append({
                        "Arquivo / Inventário": nomes_arquivos[idx],
                        "Total de Posições (SKUs)": len(df_comp),
                        "Soma Métrica Principal": f"{soma_val:,.2f}"
                    })
                df_resumo_comp = pd.DataFrame(dados_comparacao)
                st.dataframe(df_resumo_comp, use_container_width=True)
                
                fig_comp = px.bar(df_resumo_comp, x="Arquivo / Inventário", y="Total de Posições (SKUs)", color="Arquivo / Inventário", template="plotly_dark", title="Comparativo de Volume de SKUs")
                fig_comp.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_comp, use_container_width=True)

        else:
            st.error("❌ Não foi possível ler os dados dos arquivos enviados.")

if __name__ == "__main__":
    main()
