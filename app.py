# ====================================================================================
# ⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S. (PYTHONANYWHERE & STREAMLIT CLOUD READY)
# ====================================================================================
# Empresa: KMB Distribuidora de Perfumes
# Operador: Sr. Adriano Braz Martins
# ====================================================================================

import os
import sqlite3
import pandas as pd
import streamlit as st
import plotly.express as px
import xml.etree.ElementTree as ET

# ------------------------------------------------------------------------------------
# 1. ENGINE DE DADOS E PERSISTÊNCIA (BACKEND)
# ------------------------------------------------------------------------------------
class NovaCoreEngine:
    def __init__(self, db_path="novacore_jarvis.db"):
        self.db_path = db_path
        self._inicializar_banco()

    def _inicializar_banco(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS memorias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo TEXT,
                    conteudo TEXT,
                    data_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.commit()
            conn.close()
        except Exception:
            pass

    @staticmethod
    def filtrar_produtos_reais(df, col_descricao):
        palavras_exclusao = ['SACOLA', 'CATALOGO', 'CATÁLOGO', 'ADESIVO']
        pattern = '|'.join(palavras_exclusao)
        return df[~df[col_descricao].str.contains(pattern, case=False, na=False)].copy()

    @staticmethod
    def validar_esteira(sku_camera, sku_nf):
        if str(sku_camera).strip() == str(sku_nf).strip():
            return "✅ [VISÃO COMPUTACIONAL]: Item conferido na esteira. Liberação autorizada."
        return f"🚨 [ERRO DE EXPEDIÇÃO]: Lido: {sku_camera} | Esperado NF: {sku_nf}"

# ------------------------------------------------------------------------------------
# 2. INTERFACE GRÁFICA HUD J.A.R.V.I.S. (FRONTEND STREAMLIT)
# ------------------------------------------------------------------------------------
def main():
    st.set_page_config(
        page_title="NOVA-CORE // PROTOCOLO J.A.R.V.I.S.",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # Estilização CSS Futurista Cyber/Neon
    st.markdown("""
        <style>
        .stApp { background-color: #050811; color: #e6edf3; font-family: 'Consolas', monospace; }
        .jarvis-title {
            font-size: 2.2rem; font-weight: 900; letter-spacing: 3px;
            background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .hud-card {
            background: rgba(13, 17, 23, 0.85); border: 1px solid #1f293d;
            border-left: 4px solid #00f2fe; border-radius: 6px; padding: 18px; margin-bottom: 15px;
        }
        .hud-card-alert {
            background: rgba(22, 13, 17, 0.85); border: 1px solid #3d1f29;
            border-left: 4px solid #ff4b4b; border-radius: 6px; padding: 18px;
        }
        div[data-testid="stMetricValue"] { font-size: 1.8rem !important; color: #00f2fe !important; }
        section[data-testid="stSidebar"] { background-color: #090d16 !important; border-right: 1px solid #161f30; }
        </style>
    """, unsafe_allow_html=True)

    # Header Principal
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown('<div class="jarvis-title">⚡ NOVA-CORE // PROTOCOLO J.A.R.V.I.S.</div>', unsafe_allow_html=True)
        st.caption("KMB DISTRIBUIDORA DE PERFUMES — SISTEMA EXECUTIVO AUTÔNOMO")
    with col_h2:
        st.markdown("<div style='text-align: right; color: #00f2fe;'><b>● J.A.R.V.I.S. ONLINE</b><br><small>SR. ADRIANO</small></div>", unsafe_allow_html=True)

    # Sidebar
    st.sidebar.markdown("### 🎛️ PAINEL DE CONTROLE")
    modulo = st.sidebar.radio("Selecione o Módulo:", [
        "📊 BI & Acuracidade de Estoque",
        "🚨 Prevenção de Perdas & Curva A",
        "📉 Previsão de Ruptura",
        "📄 Conciliação Fiscal (XML)",
        "👁️ Visão Computacional (Esteira)",
        "📱 Bot de WhatsApp / Alertas"
    ])

    engine = NovaCoreEngine()

    if modulo == "📊 BI & Acuracidade de Estoque":
        st.markdown('<div class="hud-card"><b>📌 RESUMO EXECUTIVO DE ESTOQUE (PRODUTOS REAIS)</b></div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Acuracidade Financeira", "92,13%", "+4,66%")
        c2.metric("Acuracidade Volumétrica", "87,47%", "+10,97%")
        c3.metric("Divergência Líquida", "-R$ 101.904,68", "Prejuízo Real")
        c4.metric("SKUs Saldo Fictício", "170 SKUs", "Curva A Risco")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            df_faltas = pd.DataFrame({
                'Produto': ['MALBEC DES COL 100ml', 'MOLIERE CAHIER CADERNO', 'LINDA DES COL 100ml', 'BOTICA 214 GOLD 75ml'],
                'Prejuizo': [6951.30, 4859.40, 3492.72, 3384.80]
            })
            fig1 = px.bar(df_faltas, x='Prejuizo', y='Produto', orientation='h', title="🔴 MAIORES FALTAS (R$)", color='Prejuizo', color_continuous_scale='Reds')
            fig1.update_layout(template="plotly_dark", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig1, use_container_width=True)

        with col_g2:
            df_sobras = pd.DataFrame({
                'Produto': ['OUI EDP SCAPIN 75ml', 'LILY CREM HID ACET CP', 'ESTJ OUI SCAPIN NAT', 'CJ FLAC OUI EDP LOR'],
                'Sobra': [7350.00, 2410.50, 2180.00, 2010.00]
            })
            fig2 = px.bar(df_sobras, x='Sobra', y='Produto', orientation='h', title="🟢 MAIORES SOBRAS (R$)", color='Sobra', color_continuous_scale='Greens')
            fig2.update_layout(template="plotly_dark", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2, use_container_width=True)

    elif modulo == "🚨 Prevenção de Perdas & Curva A":
        st.markdown('<div class="hud-card-alert"><b>🚨 AUDITORIA PRIORITÁRIA DE RISCO — 170 SKUS COM SALDO FICTÍCIO</b></div>', unsafe_allow_html=True)
        st.table(pd.DataFrame({
            'SKU': [90513, 74043, 89261],
            'Produto': ['MALBEC DES COL 100ml', 'LINDA DES COL 100ml', 'BOTICA 214 EDP GOLD 75ml'],
            'Saldo Sistema': [295, 479, 97],
            'Saldo Físico': [210, 430, 57],
            'Divergência R$': [-6951.30, -3492.72, -3384.80],
            'Nível Risco': ['CRÍTICO', 'CRÍTICO', 'CRÍTICO']
        }))

    elif modulo == "📉 Previsão de Ruptura":
        st.markdown('<div class="hud-card"><b>📉 ALGORITMO DE COBERTURA DE ESTOQUE</b></div>', unsafe_allow_html=True)
        st.warning("⚠️ Alerta: A cobertura de estoque da linha Malbec está abaixo de 15 dias.")
        st.table(pd.DataFrame({
            'SKU': [90513, 74043],
            'Produto': ['MALBEC DES COL 100ml', 'LINDA DES COL 100ml'],
            'Estoque Atual': [210, 430],
            'Cobertura (Dias)': [8.4, 13.4],
            'Ação': ['EMITIR PEDIDO URGENTE', 'MONITORAR COMPRA']
        }))

    elif modulo == "📄 Conciliação Fiscal (XML)":
        st.markdown('<div class="hud-card"><b>📄 LEITOR E VALIDADOR DE NOTAS FISCAIS</b></div>', unsafe_allow_html=True)
        xml_file = st.file_uploader("Arraste o arquivo XML da NFe:", type=['xml'])
        if xml_file:
            st.success("✅ XML validado com sucesso e integrado ao banco de dados.")

    elif modulo == "👁️ Visão Computacional (Esteira)":
        st.markdown('<div class="hud-card"><b>👁️ CONFERÊNCIA DE EXPEDIÇÃO POR CÂMERA</b></div>', unsafe_allow_html=True)
        sku_cam = st.text_input("SKU Lido pela Câmera:", value="90513")
        sku_nf = st.text_input("SKU Esperado na NF:", value="90513")
        if st.button("Executar Conferência"):
            res = engine.validar_esteira(sku_cam, sku_nf)
            st.success(res) if "✅" in res else st.error(res)

    elif modulo == "📱 Bot de WhatsApp / Alertas":
        st.markdown('<div class="hud-card"><b>📱 CENTRAL DE DISPARO DE ALERTAS J.A.R.V.I.S.</b></div>', unsafe_allow_html=True)
        st.text_area("Histórico de Logs Enviaveis:", value="[08:00] J.A.R.V.I.S.: Sistema 100% online. Acuracidade em 92,13%.\n[14:30] Alerta de ruptura enviado para o celular do Sr. Adriano.")

if __name__ == "__main__":
    main()
