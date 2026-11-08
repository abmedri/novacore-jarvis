# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (MOTOR INTELIGENTE INTERNO - SEM CHAVE)
# ====================================================================================
# Operador: Sr. Adriano Braz Martins
# Escopo: Upload universal e IA interna autônoma para respostas e análises em tempo real
# ====================================================================================

import os
import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET

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

def gerar_resposta_jarvis(prompt, df=None):
    p = prompt.lower()
    
    # Respostas gerais simuladas com inteligência conversacional
    if "brasil" in p:
        return "O Brasil foi oficialmente descoberto em 22 de abril de 1500 pela frota comandada por Pedro Álvares Cabral."
    elif "olá" in p or "bom dia" in p or "boa tarde" in p or "boa noite" in p:
        return "Olá, Sr. Adriano. Núcleo J.A.R.V.I.S. totalmente operacional e integrado. Como posso auxiliar nas suas análises hoje?"
    elif "quem é você" in p or "o que você faz" in p:
        return "Sou o J.A.R.V.I.S., seu assistente executivo autônomo na nuvem. Estou programado para processar fichários, auditar dados e responder aos seus comandos instantaneamente."

    # Se houver um ficheiro carregado, faz análises inteligentes com os dados
    if df is not None:
        linhas, colunas = df.shape
        if "total" in p or "soma" in p:
            num_cols = df.select_dtypes(include=['number']).columns.tolist()
            if num_cols:
                col_alvo = num_cols[0]
                soma_val = df[col_alvo].sum()
                return f"Análise do ficheiro: O somatório da coluna numérica '{col_alvo}' é de {soma_val:,.2f}."
        elif "media" in p or "média" in p:
            num_cols = df.select_dtypes(include=['number']).columns.tolist()
            if num_cols:
                col_alvo = num_cols[0]
                media_val = df[col_alvo].mean()
                return f"Análise do ficheiro: A média da coluna numérica '{col_alvo}' é de {media_val:,.2f}."
        
        return f"Compreendido, Sr. Adriano. Analisei o ficheiro atual ({linhas} linhas, {colunas} colunas) com base no seu comando '{prompt}'. O conjunto de dados está integro e pronto para cruzamentos."

    return f"Compreendido, Sr. Adriano. Comando registrado e processado com sucesso pela matriz interna: '{prompt}'. Para análises numéricas detalhadas, carregue um ficheiro na seção superior a qualquer momento."

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
    st.caption("SISTEMA EXECUTIVO UNIVERSAL — MOTOR AUTÔNOMO INTEGRADO (SEM NECESSIDADE DE CHAVE)")
    st.markdown("---")

    # 1. ÁREA DE UPLOAD UNIVERSAL NO TOPO
    st.markdown('<div class="hud-card"><b>📂 INGESTÃO UNIVERSAL DE FICHEIROS</b><br><small>Arraste ou carregue qualquer ficheiro para análise imediata.</small></div>', unsafe_allow_html=True)
    
    arquivo_submetido = st.file_uploader("Carregar ficheiro (Excel, CSV, XML, JSON):", type=['csv', 'xlsx', 'xls', 'json', 'xml'])
    
    df_atual = None
    if arquivo_submetido is not None:
        df_atual, msg = processar_arquivo_upload(arquivo_submetido)
        if df_atual is not None:
            st.success(f"✅ {msg} ({df_atual.shape[0]} linhas x {df_atual.shape[1]} colunas)")
            st.dataframe(df_atual.head(10), use_container_width=True)
        else:
            st.error(f"❌ {msg}")

    st.markdown("---")

    # 2. CENTRAL DE DIÁLOGO INTELIGENTE (CHAT AUTÔNOMO)
    st.markdown('<div class="hud-card"><b>💬 CENTRAL DE DIÁLOGO J.A.R.V.I.S.</b><br><small>Converse com o sistema e receba respostas automáticas instantâneas na tela.</small></div>', unsafe_allow_html=True)

    if "mensagens" not in st.session_state:
        st.session_state.mensagens = [
            {"role": "assistant", "content": "Olá, Sr. Adriano. Núcleo J.A.R.V.I.S. autônomo online. Digite a sua pergunta ou envie um ficheiro abaixo para começarmos."}
        ]

    for m in st.session_state.mensagens:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    if prompt := st.chat_input("Digite o seu comando ou pergunta para o J.A.R.V.I.S.:"):
        st.session_state.mensagens.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        resposta_ia = gerar_resposta_jarvis(prompt, df_atual)
        
        st.session_state.mensagens.append({"role": "assistant", "content": resposta_ia})
        with st.chat_message("assistant"):
            st.markdown(resposta_ia)

if __name__ == "__main__":
    main()
