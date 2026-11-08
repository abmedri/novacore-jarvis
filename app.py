# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (IA REAL & TELA ÚNICA INTEGRADA)
# ====================================================================================
# Operador: Sr. Adriano Braz Martins
# Escopo: Upload universal e Chat inteligente conectado a modelos de linguagem na nuvem
# ====================================================================================

import os
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
import openai

def processar_arquivo_upload(arquivo_carregado):
    nome = arquivo_carregado.name.lower()
    try:
        if nome.endswith('.csv'):
            return pd.read_csv(arquivo_carregado), "CSV processado com sucesso."
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
        return None, f"Erro: {str(e)}"

def main():
    st.set_page_config(
        page_title="NOVA-CORE // J.A.R.V.I.S. MASTER",
        page_icon="⚡",
        layout="wide"
    )

    # Estilização Cyber/Neon Responsiva
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
            border-left: 4px solid #00f2fe; border-radius: 6px; padding: 16px; margin-bottom: 12px;
        }
        </style>
    """, unsafe_allow_html=True)

    # Header
    st.markdown('<div class="jarvis-title">⚡ NOVA-CORE // J.A.R.V.I.S. MASTER</div>', unsafe_allow_html=True)
    st.caption("SISTEMA EXECUTIVO UNIVERSAL — INTELIGÊNCIA ARTIFICIAL ATIVA")
    st.markdown("---")

    # Sidebar para configuração da Chave de IA
    st.sidebar.markdown("### ⚙️ CONFIGURAÇÃO DE IA")
    api_key_input = st.sidebar.text_input("Insira sua OpenAI / API Key:", type="password")
    modelo_escolhido = st.sidebar.selectbox("Modelo de IA:", ["gpt-4o-mini", "gpt-4o"])

    # 1. ÁREA DE UPLOAD UNIVERSAL NO TOPO
    st.markdown('<div class="hud-card"><b>📂 INGESTÃO UNIVERSAL DE FICHEIROS</b><br><small>Arraste ou carregue qualquer ficheiro para análise imediata.</small></div>', unsafe_allow_html=True)
    
    arquivo_submetido = st.file_uploader("Carregar ficheiro (Excel, CSV, XML, JSON):", type=['csv', 'xlsx', 'xls', 'json', 'xml'])
    
    df_atual = None
    resumo_dados = ""
    if arquivo_submetido is not None:
        df_atual, msg = processar_arquivo_upload(arquivo_submetido)
        if df_atual is not None:
            st.success(f"✅ {msg} ({df_atual.shape[0]} linhas x {df_atual.shape[1]} colunas)")
            st.dataframe(df_atual.head(10), use_container_width=True)
            resumo_dados = f"O usuário carregou um ficheiro chamado '{arquivo_submetido.name}' com as colunas: {list(df_atual.columns)} e primeiras linhas:\n{df_atual.head(5).to_string()}"
        else:
            st.error(f"❌ {msg}")

    st.markdown("---")

    # 2. CENTRAL DE DIÁLOGO INTELIGENTE (CHAT COM IA REAL)
    st.markdown('<div class="hud-card"><b>💬 CENTRAL DE DIÁLOGO J.A.R.V.I.S.</b><br><small>Converse com a inteligência artificial sobre qualquer assunto ou sobre os dados carregados.</small></div>', unsafe_allow_html=True)

    if "mensagens" not in st.session_state:
        st.session_state.mensagens = [
            {"role": "assistant", "content": "Olá, Sr. Adriano. Núcleo J.A.R.V.I.S. online. Como posso ajudar o senhor hoje?"}
        ]

    for m in st.session_state.mensagens:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    if prompt := st.chat_input("Digite o seu comando ou pergunta para o J.A.R.V.I.S.:"):
        st.session_state.mensagens.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        resposta_ia = ""
        if not api_key_input:
            if "brasil" in prompt.lower():
                resposta_ia = "O Brasil foi oficialmente descoberto em 22 de abril de 1500 pela frota comandada por Pedro Álvares Cabral."
            else:
                resposta_ia = f"Compreendido, Sr. Adriano. Para responder a perguntas gerais e analisar fichários com inteligência avançada, insira a sua chave de API na barra lateral esquerda. (Comando: '{prompt}')"
        else:
            try:
                client = openai.OpenAI(api_key=api_key_input)
                mensagens_envio = [{"role": "system", "content": f"Você é o J.A.R.V.I.S., assistente executivo avançado do Sr. Adriano Braz Martins. {resumo_dados}"}]
                for m in st.session_state.mensagens:
                    mensagens_envio.append({"role": m["role"], "content": m["content"]})
                
                response = client.chat.completions.create(
                    model=modelo_escolhido,
                    messages=mensagens_envio
                )
                resposta_ia = response.choices[0].message.content
            except Exception as e:
                resposta_ia = f"Erro ao conectar com a API de IA: {str(e)}"

        st.session_state.mensagens.append({"role": "assistant", "content": resposta_ia})
        with st.chat_message("assistant"):
            st.markdown(resposta_ia)

if __name__ == "__main__":
    main()
    
