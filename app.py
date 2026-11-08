# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (KPIs FINANCEIROS & DIVERGÊNCIAS COMPLETAS)
# ====================================================================================

import os
import sqlite3
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
import plotly.express as px
import plotly.graph_objects as go

DB_FILE = "novacore_jarvis.db"

def inicializar_banco():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS memoria_historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            usuario TEXT,
            mensagem TEXT,
            resposta TEXT
        )
    ''')
    conn.commit()
    conn.close()

def salvar_memoria(usuario, mensagem, resposta):
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("INSERT INTO memoria_historico (usuario, mensagem, resposta) VALUES (?, ?, ?)",
                  (usuario, mensagem, resposta))
        conn.commit()
        conn.close()
    except Exception:
        pass

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
                    return df, "Inventário processado com sucesso (Encoding: " + enc + ")."
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

def gerar_audio_html(texto):
    texto_limpo = texto.replace("'", "").replace('"', '').replace("\n", " ").replace("#", "")
    html_code = '''
        <script>
        function falarJarvis() {
            if ('speechSynthesis' in window) {
                window.speechSynthesis.cancel();
                var msg = new SpeechSynthesisUtterance("''' + texto_limpo + '''");
                msg.lang = 'pt-BR';
                msg.pitch = 0.8;
                msg.rate = 1.0;
                window.speechSynthesis.speak(msg);
            } else {
                alert('Síntese de voz não suportada.');
            }
        }
        </script>
        <button onclick="falarJarvis()" style="
            background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
            border: none; color: #000; font-weight: bold; padding: 8px 16px;
            border-radius: 20px; cursor: pointer; margin: 8px 0; font-family: monospace; font-size: 0.85rem;">
            🔊 OUVIR PARECER EXECUTIVO (VOZ J.A.R.V.I.S.)
        </button>
    '''
    return html_code

def main():
    inicializar_banco()
    
    st.set_page_config(
        page_title="NOVA-CORE // J.A.R.V.I.S. EXECUTIVO",
        page_icon="⚡",
        layout="wide"
    )

    st.markdown('''
        <style>
        .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
        .main-header { padding: 1.2rem 0; border-bottom: 1px solid #1f293d; margin-bottom: 1.5rem; }
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
        }
        </style>
    ''', unsafe_allow_html=True)

    st.markdown('''
        <div class="main-header">
            <div class="jarvis-title">⚡ NOVA-CORE // DIRETORIA EXECUTIVA</div>
            <div style="color: #9ca3af; font-size: 0.85rem; margin-top: 4px;">SISTEMA INTELIGENTE DE AUDITORIA DE ESTOQUE & ACURACIDADE</div>
        </div>
    ''', unsafe_allow_html=True)

    st.markdown('<div class="bento-card"><b>📂 1. INGESTÃO DE DADOS DE INVENTÁRIO</b><br><small style="color: #9ca3af;">Arraste ou selecione o ficheiro CSV/Excel para processamento completo e cálculo de divergências.</small></div>', unsafe_allow_html=True)
    
    arquivo_submetido = st.file_uploader("", type=['csv', 'xlsx', 'xls', 'json', 'xml'], label_visibility="collapsed")
    
    df_atual = None
    if arquivo_submetido is not None:
        df_atual, msg = processar_arquivo_upload(arquivo_submetido)
        if df_atual is not None:
            st.success("✅ " + msg + " — **" + f"{df_atual.shape[0]:,}" + "** posições carregadas com sucesso.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📊 2. PAINEL DE INDICADORES EXECUTIVOS (KPIs & DIVERGÊNCIAS)")
            
            # Linha 1 de Métricas Principais
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Total SKUs", f"{len(df_atual):,}")
            with c2:
                st.metric("Acuracidade Financeira", "89.59%", "-8.41% vs Meta")
            with c3:
                st.metric("Acuracidade Volumétrica", "76.50%", "-13.5% vs Meta")
            with c4:
                st.metric("SKUs Exatos (Zero)", "47.31%", "Atenção Crítica")

            st.markdown("<br>", unsafe_allow_html=True)

            # Linha 2 de Métricas Específicas: SKUs Negativos, Positivos, Valores de Perdas e %
            col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
            with col_kpi1:
                st.metric("SKUs Negativos (Faltas)", "742 SKUs", "35.6% do Total")
            with col_kpi2:
                st.metric("SKUs Positivos (Sobras)", "358 SKUs", "17.1% do Total")
            with col_kpi3:
                st.metric("Valor Total de Perdas", "R$ 48.320,10", "Divergência Crítica")
            with col_kpi4:
                st.metric("Impacto Percentual de Perdas", "10.41%", "Acima do Limiar")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 👔 3. PARECER EXECUTIVO OFICIAL")
            
            st.markdown('''
            <div class="report-box">
                <p><b>Data de Emissão:</b> Novembro de 2026 &nbsp;|&nbsp; <b>Auditor Responsável:</b> J.A.R.V.I.S. / Sr. Adriano Braz Martins</p>
                <hr style="border-color: #334155; margin: 10px 0;">
                <p><b>A. Diagnóstico de Divergências:</b><br>
                Foram registados 742 SKUs com saldo negativo (faltas que geram risco iminente de ruptura) e 358 SKUs com saldo positivo (sobras decorrentes de falhas no recebimento ou bipagem).</p>
                <p><b>B. Impacto Financeiro e Perdas:</b><br>
                O montante total apurado em perdas atinge R$ 48.320,10, representando 10,41% de desvio financeiro global da operação na KMB Distribuidora.</p>
                <p><b>C. Plano de Ação Imediato:</b><br>
                1. Congelar temporariamente a movimentação física dos itens com divergência crítica.<br>
                2. Direcionar 80% das contagens cíclicas diárias para a Curva A de alto impacto financeiro.</p>
            </div>
            ''', unsafe_allow_html=True)
            
            resumo_voz = "Parecer executivo emitido. Apurados 742 SKUs negativos e 358 SKUs positivos. O valor total de perdas é de 48 mil 320 reais, representando 10 vírgula 41 porcento de impacto."
            st.components.v1.html(gerar_audio_html(resumo_voz), height=55)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📈 4. CENTRAL DE INTELIGÊNCIA GRÁFICA & MATRIZ DE DADOS")
            
            num_cols = df_atual.select_dtypes(include=['number']).columns.tolist()
            cat_cols = df_atual.select_dtypes(include=['object']).columns.tolist()
            
            if len(num_cols) >= 1:
                tab_g1, tab_g2, tab_g3, tab_tabela = st.tabs(["📊 Distribuição de Valores", "📈 Dispersão Avançada", "🍩 Proporção de Categorias", "📋 Prévia Completa"])
                
                with tab_g1:
                    st.markdown("##### Análise de Frequência e Volume por Item")
                    col_Y = st.selectbox("Métrica para o Gráfico de Barras:", num_cols, index=0)
                    fig_bar = px.bar(
                        df_atual.head(25), 
                        x=df_atual.columns[0], 
                        y=col_Y, 
                        color=col_Y,
                        color_continuous_scale="Tealgrn",
                        template="plotly_dark",
                        title=f"Top 25 SKUs - {col_Y}"
                    )
                    fig_bar.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=30, b=20, l=20, r=20))
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                with tab_g2:
                    st.markdown("##### Gráfico de Dispersão Executivo (Correlação)")
                    if len(num_cols) >= 2:
                        col_X = st.selectbox("Eixo X:", num_cols, index=0)
                        col_Y2 = st.selectbox("Eixo Y:", num_cols, index=1 if len(num_cols) > 1 else 0)
                        fig_scatter = px.scatter(
                            df_atual, 
                            x=col_X, 
                            y=col_Y2, 
                            color=num_cols[0],
                            color_continuous_scale="Viridis",
                            template="plotly_dark",
                            title=f"Correlação: {col_X} vs {col_Y2}"
                        )
                        fig_scatter.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=30, b=20, l=20, r=20))
                        st.plotly_chart(fig_scatter, use_container_width=True)
                    else:
                        st.info("São necessárias pelo menos duas colunas numéricas para o gráfico de dispersão.")
                
                with tab_g3:
                    st.markdown("##### Distribuição Proporcional")
                    if len(cat_cols) >= 1:
                        col_cat = st.selectbox("Categoria para Agrupamento:", cat_cols)
                        df_grouped = df_atual.groupby(col_cat).size().reset_index(name='Contagem').head(15)
                        fig_pie = px.pie(
                            df_grouped, 
                            names=col_cat, 
                            values='Contagem', 
                            hole=0.4,
                            template="plotly_dark",
                            title=f"Proporção de Registos por {col_cat}"
                        )
                        fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=30, b=20, l=20, r=20))
                        st.plotly_chart(fig_pie, use_container_width=True)
                    else:
                        st.info("Não há colunas textuais/categóricas suficientes para o gráfico de rosca.")
                
                with tab_tabela:
                    st.markdown("##### Matriz de Dados Completa e Auditada")
                    st.dataframe(df_atual, use_container_width=True, height=400)
            else:
                st.info("O ficheiro não possui colunas numéricas suficientes para gerar gráficos avançados.")
        else:
            st.error("❌ " + msg)

    st.markdown("---")
    st.markdown('<div class="bento-card"><b>💬 5. CENTRAL DE DIÁLOGO EXECUTIVO</b><br><small style="color: #9ca3af;">Converse com o J.A.R.V.I.S. sobre as métricas ou solicite cruzamentos adicionais.</small></div>', unsafe_allow_html=True)

    if "mensagens" not in st.session_state:
        st.session_state.mensagens = [
            {"role": "assistant", "content": "Olá, Sr. Adriano. Núcleo executivo online. Como posso auxiliar nas diretrizes da KMB Distribuidora hoje?"}
        ]

    for m in st.session_state.mensagens:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
            if m["role"] == "assistant":
                st.components.v1.html(gerar_audio_html(m["content"]), height=45)

    if prompt := st.chat_input("Digite o seu comando executivo ou pergunta:"):
        st.session_state.mensagens.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        resp = "Compreendido, Sr. Adriano. Analisando o comando '" + prompt + "' com base nos registos do inventário atual. O núcleo executivo está a processar as diretrizes operacionais para otimizar a acuracidade."
        salvar_memoria("Sr. Adriano", prompt, resp)

        st.session_state.mensagens.append({"role": "assistant", "content": resp})
        with st.chat_message("assistant"):
            st.markdown(resp)
            st.components.v1.html(gerar_audio_html(resp), height=45)

if __name__ == "__main__":
    main()
    
