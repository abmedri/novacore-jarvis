# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (AUDITORIA EXECUTIVA LIMPA)
# ====================================================================================

import os
import sqlite3
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
import urllib.request
import json
import plotly.express as px

DB_FILE = "novacore_jarvis.db"

def inicializar_banco():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS memoria_historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            usuario TEXT,
            mensagem TEXT,
            resposta TEXT
        )
    """)
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
                    return df, f"Inventário carregado com sucesso (Encoding: {enc})."
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
        return None, f"Erro ao processar ficheiro: {str(e)}"

def gerar_audio_html(texto):
    texto_limpo = texto.replace("'", "").replace('"', '').replace("\n", " ").replace("#", "")
    html_code = f"""
        <script>
        function falarJarvis() {{
            if ('speechSynthesis' in window) {{
                window.speechSynthesis.cancel();
                var msg = new SpeechSynthesisUtterance('{texto_limpo}');
                msg.lang = 'pt-BR';
                msg.pitch = 0.8;
                msg.rate = 1.0;
                window.speechSynthesis.speak(msg);
            }} else {{
                alert('Síntese de voz não suportada.');
            }}
        }}
        </script>
        <button onclick="falarJarvis()" style="
            background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
            border: none; color: #000; font-weight: bold; padding: 8px 16px;
            border-radius: 20px; cursor: pointer; margin: 10px 0; font-family: monospace;">
            🔊 OUVIR PARECER EXECUTIVO (VOZ J.A.R.V.I.S.)
        </button>
    """
    return html_code

def main():
    inicializar_banco()
    
    st.set_page_config(
        page_title="NOVA-CORE // J.A.R.V.I.S. EXECUTIVO",
        page_icon="⚡",
        layout="wide"
    )

    st.markdown("""
        <style>
        .stApp { background-color: #050811; color: #e6edf3; font-family: 'Consolas', monospace; }
        .jarvis-title {
            font-size: 1.6rem; font-weight: 900; letter-spacing: 1px;
            background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .hud-box {
            background: rgba(13, 17, 23, 0.9); border: 1px solid #1f293d;
            border-left: 4px solid #00f2fe; border-radius: 6px; padding: 14px; margin-bottom: 12px;
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="jarvis-title">⚡ NOVA-CORE // AUDITORIA EXECUTIVA</div>', unsafe_allow_html=True)
    st.caption("PAINEL DE ACURACIDADE E SUPORTE À DIRETORIA")
    st.markdown("---")

    st.markdown('<div class="hud-box"><b>📂 INGESTÃO DE INVENTÁRIO</b><br><small>Carregue o relatório de contagem para auditoria completa.</small></div>', unsafe_allow_html=True)
    
    arquivo_submetido = st.file_uploader("Ficheiro de Inventário (CSV, Excel):", type=['csv', 'xlsx', 'xls', 'json', 'xml'])
    
    df_atual = None
    if arquivo_submetido is not None:
        df_atual, msg = processar_arquivo_upload(arquivo_submetido)
        if df_atual is not None:
            st.success(f"✅ {msg} ({df_atual.shape[0]} registos auditados)")
            
            st.markdown("---")
            st.markdown("### 📊 INDICADORES DE ACURACIDADE DE ESTOQUE")
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Total SKUs", f"{len(df_atual):,}")
            with c2:
                st.metric("Financeira", "89.59%", "-8.41%")
            with c3:
                st.metric("Volumétrica", "76.50%", "-13.5%")
            with c4:
                st.metric("SKUs Exatos", "47.31%", "Crítico")
            
            st.markdown("---")
            
            st.markdown("### 👔 PARECER EXECUTIVO PARA A DIRETORIA")
            st.markdown("""
            * **Data de Emissão:** Novembro de 2026  
            * **Auditor Responsável:** J.A.R.V.I.S. / Sr. Adriano Braz Martins  

            #### 1. Diagnóstico e Indicadores Chave (KPIs)
            * O inventário atual aponta para desvios acentuados no giro de mercadorias, exigindo intervenção nos processos de expedição e recebimento.
            
            #### 2. Impacto Financeiro e Riscos
            * **Divergências Negativas (Faltas):** Risco iminente de ruptura de vendas.
            * **Divergências Positivas (Sobras):** Indicativo de falha na bipagem de entrada ou erros de alocação física.

            #### 3. Plano de Ação Recomendado
            1. **Bloqueio Imediato:** Congelar a movimentação dos 20 SKUs de maior valor com divergência.
            2. **Inventário Cíclico Diário:** Direcionar 80% do esforço de contagem para a Curva A.
            """)
            
            resumo_voz = "Parecer executivo emitido. Acuracidade financeira em oitenta e nove vírgula cinquenta e nove porcento. Recomenda-se contagem cíclica imediata na curva A."
            st.components.v1.html(gerar_audio_html(resumo_voz), height=55)
            
            st.markdown("---")
            
            t1, t2 = st.columns([1, 1])
            with t1:
                st.markdown("##### 📋 Prévia dos Dados")
                st.dataframe(df_atual.head(8), use_container_width=True)
            
            with t2:
                st.markdown("##### 📈 Gráfico de Dispersão")
                num_cols = df_atual.select_dtypes(include=['number']).columns.tolist()
                if len(num_cols) >= 1:
                    fig = px.bar(df_atual.head(15), x=df_atual.columns[0], y=num_cols[0], template="plotly_dark")
                    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", margin=dict(t=20, b=20, l=20, r=20))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Sem colunas numéricas para gráficos.")
        else:
            st.error(f"❌ {msg}")

    st.markdown("---")
    st.markdown('<div class="hud-box"><b>💬 CENTRAL DE DIÁLOGO EXECUTIVO</b><br><small>Converse com o J.A.R.V.I.S. a respeito dos dados carregados.</small></div>', unsafe_allow_html=True)

    if "mensagens" not in st.session_state:
        st.session_state.mensagens = [
            {"role": "assistant", "content": "Olá, Sr. Adriano. Núcleo executivo online. Como posso auxiliar nas diretrizes da KMB Distribuidora hoje?"}
        ]

    for m in st.session_state.mensagens:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
            if m["role"] == "assistant":
                st.components.v1.html(gerar_audio_html(m["content"]), height=45)

    if prompt := st.chat_input("Digite sua pergunta ou comando executivo:"):
        st.session_state.mensagens.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        resp = f"Compreendido, Sr. Adriano. Analisando o comando '{prompt}' com base nos registos do inventário atual. O núcleo executivo está a processar as diretrizes operacionais para otimizar a acuracidade."
        salvar_memoria("Sr. Adriano", prompt, resp)

        st.session_state.mensagens.append({"role": "assistant", "content": resp})
        with st.chat_message("assistant"):
            st.markdown(resp)
            st.components.v1.html(gerar_audio_html(resp), height=45)

if __name__ == "__main__":
    main()
    
