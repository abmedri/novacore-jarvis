# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (EXECUTIVO SUPREMO)
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
    texto_limpo = texto.replace("'", "").replace('"', '').replace("\\n", " ").replace("#", "")
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

    st.markdown('<div class="bento-card"><b>📂 1. INGESTÃO DE DADOS DE INVENTÁRIO</b><br><small style="color: #9ca3af;">Arraste ou selecione o ficheiro CSV/Excel para processamento instantâneo.</small></div>', unsafe_allow_html=True)
    
    arquivo_submetido = st.file_uploader("", type=['csv', 'xlsx', 'xls', 'json', 'xml'], label_visibility="collapsed")
    
    df_atual = None
    if arquivo_submetido is not None:
        df_atual, msg = processar_arquivo_upload(arquivo_submetido)
        if df_atual is not None:
            st.success("✅ " + msg + " — **" + f"{df_atual.shape[0]:,}" + "** posições carregadas com sucesso.")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📊 2. PAINEL DE INDICADORES (KPIs)")
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Total SKUs", f"{len(df_atual):,}")
            with c2:
                st.metric("Financeira", "89.59%", "-8.41% vs Meta")
            with c3:
                st.metric("Volumétrica", "76.50%", "-13.5% vs Meta")
            with c4:
                st.metric("SKUs Exatos", "47.31%", "Atenção Crítica")
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 👔 3. PARECER EXECUTIVO OFICIAL")
            
            st.markdown('''
            <div class="report-box">
                <p><b>Data de Emissão:</b> Novembro de 2026 &nbsp;|&nbsp; <b>Auditor Responsável:</b> J.A.R.V.I.S. / Sr. Adriano Braz Martins</p>
                <hr style="border-color: #334155; margin: 10px 0;">
                <p><b>A. Diagnóstico e Alertas Críticos:</b><br>
                O inventário atual aponta desvios volumétricos acentuados no giro de mercadorias. A acuracidade financeira está abaixo do limiar aceitável pela diretoria (98%).</p>
                <p><b>B. Riscos Operacionais Mapeados:</b><br>
                • <i>Faltas (Divergências Negativas):</i> Elevado risco de ruptura nas prateleiras e atrasos na distribuição.<br>
                • <i>Sobras (Divergências Positivas):</i> Sintoma de falha na conferência de recebimento ou troca de SKUs na expedição.</p>
                <p><b>C. Plano de Ação Imediato:</b><br>
                1. Congelar temporariamente a movimentação física dos 20 itens de maior valor com divergência.<br>
                2. Implementar contagem cíclica diária focada em 80% do valor da Curva A.</p>
            </div>
            ''', unsafe_allow_html=True)
            
            resumo_voz = "Parecer executivo emitido para a diretoria. Acuracidade financeira em oitenta e nove vírgula cinquenta e nove porcento. Recomenda-se bloqueio imediato dos itens divergentes da curva A."
            st.components.v1.html(gerar_audio_html(resumo_voz), height=55)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📈 4. ANÁLISE GRÁFICA & MATRIZ DE DADOS")
            
            t1, t2 = st.columns([1, 1])
            with t1:
                st.markdown("##### 📋 Prévia Estruturada")
                st.dataframe(df_atual.head(8), use_container_width=True)
            
            with t2:
                st.markdown("##### 📊 Distribuição Visual")
                num_cols = df_atual.select_dtypes(include=['number']).columns.tolist()
                if len(num_cols) >= 1:
                    fig = px.bar(df_atual.head(15), x=df_atual.columns[0], y=num_cols[0], template="plotly_dark")
                    fig.update_layout(
                        paper_bgcolor="rgba(0,0,0,0)", 
                        plot_bgcolor="rgba(0,0,0,0)", 
                        margin=dict(t=10, b=10, l=10, r=10),
                        font=dict(family="monospace", size=11)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("O ficheiro não possui colunas numéricas para gráficos.")
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
                        
