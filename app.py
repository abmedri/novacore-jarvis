# ====================================================================================
# ⚡ NOVA-CORE // J.A.R.V.I.S. (VERSÃO ESTÁVEL & BLINDADA)
# ====================================================================================

import pandas as pd
import streamlit as st
import plotly.express as px
import numpy as np

def main():
    st.set_page_config(page_title="NOVA-CORE // J.A.R.V.I.S.", page_icon="⚡", layout="wide")

    st.markdown("""
    <style>
    .stApp { background-color: #0b0f19; color: #f3f4f6; font-family: 'Inter', 'Segoe UI', monospace; }
    .jarvis-title {
        font-size: 1.7rem; font-weight: 800; letter-spacing: 1.5px;
        background: linear-gradient(90deg, #00f2fe 0%, #3b82f6 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div style="padding: 10px 0; border-bottom: 1px solid #1f293d; margin-bottom: 20px;">
        <div class="jarvis-title">⚡ NOVA-CORE // DIRETORIA EXECUTIVA</div>
        <div style="color: #9ca3af; font-size: 0.85rem;">AUDITORIA DE ESTOQUE & ANÁLISE COMPARATIVA</div>
    </div>
    """, unsafe_allow_html=True)

    # 📂 BARRA LATERAL (UPLOAD E MODOS)
    st.sidebar.markdown("### 🎛️ CONTROLES & NAVEGAÇÃO")
    
    # GARANTINDO QUE O UPLOADER APAREÇA SEMPRE
    uploaded_file = st.sidebar.file_uploader("📂 Carregar Ficheiro CSV", type=['csv', 'xlsx', 'xls'])
    
    modo = st.sidebar.radio("Selecione o Modo:", [
        "Modo 1: Indicadores Executivos & Balanço", 
        "Modo 2: Curva ABC & Concentração Financeira"
    ])

    if uploaded_file is not None:
        try:
            # Tenta ler o CSV com codificação robusta
            try:
                df = pd.read_csv(uploaded_file, encoding='latin1', sep=';')
            except Exception:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='utf-8', sep=',')

            # Limpeza das colunas numéricas
            cols_numericas = ['Congelado', 'Digitado', 'Ajuste', 'Saldo em Estoque', 'Custo', 'Valor de Ajuste(R$)', 'Valor de Estoque(R$)']
            for col in cols_numericas:
                if col in df.columns:
                    df[col] = (
                        df[col]
                        .astype(str)
                        .str.replace('.', '', regex=False)
                        .str.replace(',', '.', regex=False)
                        .astype(float)
                    )

            total_skus = len(df)
            sobras_df = df[df['Ajuste'] > 0] if 'Ajuste' in df.columns else pd.DataFrame()
            faltas_df = df[df['Ajuste'] < 0] if 'Ajuste' in df.columns else pd.DataFrame()
            
            skus_sobra = len(sobras_df)
            skus_falta = len(faltas_df)
            total_div = skus_sobra + skus_falta

            estoque_depois = df['Valor de Estoque(R$)'].sum() if 'Valor de Estoque(R$)' in df.columns else 0.0
            estoque_antes = (df['Congelado'] * df['Custo']).sum() if 'Congelado' in df.columns and 'Custo' in df.columns else estoque_depois
            ajuste_liq = (df['Ajuste'] * df['Custo']).sum() if 'Ajuste' in df.columns and 'Custo' in df.columns else 0.0

            # 📌 MODO 1
            if modo == "Modo 1: Indicadores Executivos & Balanço":
                st.markdown("### 📌 MODO 1: Indicadores Executivos & Parecer Técnico")
                
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.metric("Total SKUs Auditados", f"{total_skus:,}")
                with c2:
                    st.metric("Produtos com Sobra", f"{skus_sobra} SKUs")
                with c3:
                    st.metric("Produtos com Falta", f"{skus_falta} SKUs")

                st.markdown("##### 💰 Balanço Patrimonial & Divergência Monetária")
                bc1, bc2, bc3 = st.columns(3)
                with bc1:
                    st.metric("Estoque Antes (Teórico)", f"R$ {estoque_antes:,.2f}")
                with bc2:
                    st.metric("Estoque Depois (Físico)", f"R$ {estoque_depois:,.2f}")
                with bc3:
                    st.metric("Ajuste Líquido", f"R$ {ajuste_liq:,.2f}")

                st.markdown("---")
                st.dataframe(df.head(100), use_container_width=True)

            # 📊 MODO 2
            elif modo == "Modo 2: Curva ABC & Concentração Financeira":
                st.markdown("### 📊 MODO 2: Concentração Financeira & Curva ABC (Pareto 80/20)")
                
                todas_colunas = df.columns.tolist()
                c_x2, c_y2 = st.columns(2)
                with c_x2:
                    eixo_x = st.selectbox("Eixo X (Produto / SKU):", todas_colunas, index=0)
                with c_y2:
                    eixo_y = st.selectbox("Eixo Y (Valor / Quantidade):", todas_colunas, index=todas_colunas.index('Valor de Estoque(R$)') if 'Valor de Estoque(R$)' in todas_colunas else 0)
                
                df_abc = df.sort_values(by=eixo_y, ascending=False).copy()
                df_abc['Soma_Acumulada'] = df_abc[eixo_y].cumsum()
                soma_total = df_abc[eixo_y].sum()
                
                if soma_total > 0:
                    df_abc['%_Acumulado'] = (df_abc['Soma_Acumulada'] / soma_total) * 100
                    df_abc['Classe_ABC'] = df_abc['%_Acumulado'].apply(lambda p: "Classe A (80%)" if p <= 80 else ("Classe B (15%)" if p <= 95 else "Classe C (5%)"))
                    
                    fig = px.bar(
                        df_abc.head(100), 
                        x=eixo_x, 
                        y=eixo_y, 
                        color='Classe_ABC', 
                        template="plotly_dark", 
                        title=f"Curva ABC — {eixo_y} por {eixo_x}"
                    )
                    fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig, use_container_width=True)
                    
                st.dataframe(df_abc[[eixo_x, eixo_y, '%_Acumulado', 'Classe_ABC']], use_container_width=True)

        except Exception as e:
            st.error(f"Erro ao processar o ficheiro: {e}")
    else:
        st.info("👈 Utilize a barra lateral à esquerda para carregar o seu ficheiro CSV ou Excel de inventário.")

if __name__ == "__main__":
    main()
