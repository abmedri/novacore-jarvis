# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. TOTAL (ROBUSTO PARA CSV COM SEPARADOR ';')
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
                    # Tenta ler com separador ';' e converte automaticamente colunas numéricas se possível
                    df = pd.read_csv(arquivo_carregado, encoding=enc, sep=';', low_memory=False)
                    # Se tiver apenas 1 coluna, tenta ler com vírgula comum
                    if df.shape[1] <= 1:
                        arquivo_carregado.seek(0)
                        df = pd.read_csv(arquivo_carregado, encoding=enc, sep=',', low_memory=False)
                    
                    # Tentar converter colunas que pareçam números
                    for col in df.columns:
                        if df[col].dtype == object:
                            try:
                                convertido = df[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                                df[col] = pd.to_numeric(convertido)
                            except Exception:
                                pass
                    return df, f"CSV processado com sucesso (encoding: {enc})."
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

def consultar_ia_natural(prompt, df=None):
    p = prompt.lower().strip()
    
    if df is not None:
        linhas, colunas = df.shape
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        if any(term in p for term in ["total", "soma", "somatorio", "quanto"]):
            if num_cols:
                col = num_cols[0]
                soma = df[col].sum()
                return f"📊 **Análise de Ficheiro:** O somatório da coluna numérica **'{col}'** é **{soma:,.2f}**."
        elif any(term in p for term in ["media", "média"]):
            if num_cols:
                col = num_cols[0]
                media = df[col].mean()
                return f"📊 **Análise de Ficheiro:** A média da coluna numérica **'{col}'** é **{media:,.2f}**."
        elif any(term in p for term in ["linhas", "colunas", "tamanho", "estrutura"]):
            return f"📂 **Estrutura de Dados:** O ficheiro contém **{linhas} linhas** e **{colunas} colunas**. Colunas numéricas detetadas: {num_cols}."
        return f"✅ Ficheiro ativo ({linhas}L x {colunas}C). Comando registado: '{prompt}'. Dashboard atualizado abaixo."

    try:
        url = "https://api.duckduckgo.com/?q=" + urllib.parse.quote(prompt) + "&format=json"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=4) as response:
            data = json.loads(response.read().decode())
            if data.get('AbstractText'):
                return f"🤖 **J.A.R.V.I.S.:** {data['AbstractText']}"
            elif data.get('RelatedTopics') and len(data['RelatedTopics']) > 0:
                for topico in data['RelatedTopics']:
                    if isinstance(topico, dict) and 'Text' in topico:
                        return f"🤖 **J.A.R.V.I.S.:** {topico['Text']}"
    except Exception:
        pass

    if any(term in p for term in ["ola", "olá", "oi", "bom dia", "boa tarde", "boa noite", "hey"]):
        return "Olá, Sr. Adriano! Núcleo J.A.R.V.I.S. totalmente operacional. Como posso ajudá-lo com as suas análises hoje?"
    elif any(term in p for term in ["quem e voce", "quem é você", "o que você faz", "funcao"]):
        return "Sou o J.A.R.V.I.S., seu assistente executivo autônomo na nuvem. Possuo ingestão universal, dashboards interativos, síntese de voz e memória persistente."
    elif any(term in p for term in ["brasil", "descobriu"]):
        return "O Brasil foi oficialmente descoberto em 22 de abril de 1500 pela frota comandada pelo navegador português Pedro Álvares Cabral."
    elif any(term in p for term in ["ajuda", "comandos", "recursos"]):
        return "Comandos disponíveis: pode fazer perguntas livremente, carregar ficheiros para ver dashboards estatísticos, ou tocar nos botões de áudio para me ouvir responder!"

    return f"Compreendido, Sr. Adriano. Processei a sua requisição ('{prompt}'). Estou à disposição para aprofundar este tema ou realizar análises no seu ficheiro."

def gerar_audio_html(texto):
    texto_limpo = texto.replace("'", "\\\\\'").replace('"', '\\\\"').replace("\n", " ")
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
                alert('Síntese de voz não suportada neste navegador.');
            }}
        }}
        </script>
        <button onclick="falarJarvis()" style="
            background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
            border: none; color: #000; font-weight: bold; padding: 6px 14px;
            border-radius: 20px; cursor: pointer; margin-top: 6px; font-family: monospace;">
            🔊 OUVIR RESPOSTA (VOZ J.A.R.V.I.S.)
        </button>
    """
    return html_code

def main():
    inicializar_banco()
    
    st.set_page_config(
        page_title="NOVA-CORE // J.A.R.V.I.S. TOTAL",
        page_icon="⚡",
        layout="wide"
    )

    st.markdown("""
        <style>
        .stApp { background-color: #050811; color: #e6edf3; font-family: 'Consolas', monospace; }
        .jarvis-title {
            font-size: 1.8rem; font-weight: 900; letter-spacing: 2px;
            background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .hud-card {
            background: rgba(13, 17, 23, 0.85); border: 1px solid #1f293d;
            border-left: 4px solid #00f2fe; border-radius: 6px; padding: 14px; margin-bottom: 12px;
        }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="jarvis-title">⚡ NOVA-CORE // J.A.R.V.I.S. TOTAL</div>', unsafe_allow_html=True)
    st.caption("SISTEMA EXECUTIVO SUPREMO — INGESTÃO, DASHBOARDS, MEMÓRIA SQLITE E SÍNTESE DE VOZ ATIVAS")
    st.markdown("---")

    st.markdown('<div class="hud-card"><b>📂 INGESTÃO UNIVERSAL & DASHBOARD AUTOMÁTICO</b><br><small>Carregue ficheiros para visualizar dados e gráficos dinâmicos instantaneamente.</small></div>', unsafe_allow_html=True)
    
    arquivo_submetido = st.file_uploader("Carregar Ficheiro (Excel, CSV, JSON, XML):", type=['csv', 'xlsx', 'xls', 'json', 'xml'])
    
    df_atual = None
    if arquivo_submetido is not None:
        df_atual, msg = processar_arquivo_upload(arquivo_submetido)
        if df_atual is not None:
            st.success(f"✅ {msg} ({df_atual.shape[0]} linhas x {df_atual.shape[1]} colunas)")
            
            col_tabela, col_grafico = st.columns([1, 1])
            with col_tabela:
                st.markdown("##### 📋 Tabela de Dados (Prévia)")
                st.dataframe(df_atual.head(8), use_container_width=True)
            
            with col_grafico:
                st.markdown("##### 📊 Dashboard Gráfico Dinâmico")
                num_cols = df_atual.select_dtypes(include=['number']).columns.tolist()
                if len(num_cols) >= 1:
                    col_y = num_cols[0]
                    col_x = df_atual.columns[0]
                    fig = px.bar(df_atual.head(15), x=col_x, y=col_y, title=f"Distribuição: {col_y} por {col_x}", template="plotly_dark")
                    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("O ficheiro carregado não possui colunas numéricas detetadas para montagem automática de gráficos.")
        else:
            st.error(f"❌ {msg}")

    st.markdown("---")

    st.markdown('<div class="hud-card"><b>💬 CENTRAL DE DIÁLOGO & REPRODUÇÃO DE VOZ J.A.R.V.I.S.</b><br><small>Converse livremente e clique no botão para ouvir o J.A.R.V.I.S. falar no seu telemóvel.</small></div>', unsafe_allow_html=True)

    if "mensagens" not in st.session_state:
        st.session_state.mensagens = [
            {"role": "assistant", "content": "Olá, Sr. Adriano. Núcleo J.A.R.V.I.S. Total ativado. Como posso auxiliar o senhor hoje?"}
        ]

    for m in st.session_state.mensagens:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
            if m["role"] == "assistant":
                st.components.v1.html(gerar_audio_html(m["content"]), height=45)

    if prompt := st.chat_input("Digite a sua pergunta ou comando para o J.A.R.V.I.S.:"):
        st.session_state.mensagens.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        resposta_jarvis = consultar_ia_natural(prompt, df_atual)
        salvar_memoria("Sr. Adriano", prompt, resposta_jarvis)

        st.session_state.mensagens.append({"role": "assistant", "content": resposta_jarvis})
        with st.chat_message("assistant"):
            st.markdown(resposta_jarvis)
            st.components.v1.html(gerar_audio_html(resposta_jarvis), height=45)

if __name__ == "__main__":
    main()
    
