import streamlit as st
import pandas as pd
from io import BytesIO
import re

# 1. Configuração da Página
st.set_page_config(page_title="Limpador de Excel Pro", layout="wide")

# --- FUNÇÕES DE UTILIDADE ---
def limpar_numero(tel):
    if pd.isna(tel): return ""
    return re.sub(r'\D', '', str(tel))

def normalizar_colunas(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df

def to_csv(df):
    return df.to_csv(index=False).encode('utf-8-sig') # utf-8-sig para abrir direto no Excel sem erro de acento

# --- INTERFACE PRINCIPAL ---
st.title("✂️ Ferramenta de Tratamento e Unificação de Dados")

tab1, tab2 = st.tabs(["🎯 1. Ajustador Inteligente", "🔗 2. Complementar Dados (VLOOKUP)"])

# --- ABA 1: AJUSTADOR INTELIGENTE ---
with tab1:
    st.header("Ajuste de Colunas e Formatação")
    arquivo_ajuste = st.file_uploader("Suba a planilha para formatar (.xlsx)", type=["xlsx"], key="ajuste_single")

    if arquivo_ajuste:
        df = pd.read_excel(arquivo_ajuste)
        df = normalizar_colunas(df)
        total_original = len(df)

        st.sidebar.header("⚙️ Configurações Aba 1")
        aj_duplicatas = st.sidebar.checkbox("Remover linhas 100% idênticas", key="dup_tab1")
        aj_socios = st.sidebar.checkbox("Sequenciar Sócios (Nome 2, 3...)", key="soc_tab1")
        unir_tels = st.sidebar.checkbox("Unir DDD + Números", key="tel_tab1")

        if unir_tels:
            pares = [('ddd_cel1', 'cel1', 'celular_1_completo'), ('ddd_cel2', 'cel2', 'celular_2_completo'),
                     ('ddd_cel3', 'cel3', 'celular_3_completo'), ('ddd_tel1', 'tel1', 'telefone_1_completo'),
                     ('ddd_tel2', 'tel2', 'telefone_2_completo')]
            col_remover = []
            for ddd, num, nova in pares:
                if ddd in df.columns and num in df.columns:
                    df[nova] = df.apply(lambda x: limpar_numero(x[ddd]) + limpar_numero(x[num]), axis=1)
                    col_remover.extend([ddd, num])
            df = df.drop(columns=[c for c in col_remover if c in df.columns])

        selecionadas = st.multiselect("Manter colunas:", df.columns.tolist(), default=df.columns.tolist())

        if selecionadas:
            df_final = df[selecionadas].copy()
            linhas_removidas = 0
            if aj_duplicatas:
                antes_dup = len(df_final)
                df_final = df_final.drop_duplicates()
                depois_dup = len(df_final)
                linhas_removidas = antes_dup - depois_dup

            if aj_socios and 'nome' in df_final.columns:
                df_final['count'] = df_final.groupby('nome').cumcount() + 1
                df_final['nome'] = df_final.apply(lambda x: f"{x['nome']} {x['count']}" if x['count'] > 1 else x['nome'], axis=1)
                df_final = df_final.drop(columns=['count'])

            st.markdown("---")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Original", f"{total_original} linhas")
            m2.metric("Total Final", f"{len(df_final)} linhas")
            if aj_duplicatas: m3.metric("Linhas Removidas", f"{linhas_removidas}", delta=f"-{linhas_removidas}", delta_color="inverse")
            
            st.dataframe(df_final.head(10))
            
            # BOTÕES DE DOWNLOAD LADO A LADO
            col_down1, col_down2 = st.columns(2)
            
            # Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            col_down1.download_button("🚀 Baixar em EXCEL", output.getvalue(), "planilha_ajustada.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            # CSV
            col_down2.download_button("📄 Baixar em CSV", to_csv(df_final), "planilha_ajustada.csv", "text/csv")

# --- ABA 2: UNIFICADOR / COMPLEMENTAR DADOS ---
with tab2:
    st.header("🔗 Cruzamento de Dados por CNPJ")
    c1, c2 = st.columns(2)
    with c1: arq_soc = st.file_uploader("Arquivo Base (Sócios)", type=["xlsx"], key="f_socios")
    with c2: arq_emp = st.file_uploader("Arquivo de Consulta (Empresas)", type=["xlsx"], key="f_empresas")

    if arq_soc and arq_emp:
        df_s = normalizar_colunas(pd.read_excel(arq_soc))
        df_e = normalizar_colunas(pd.read_excel(arq_emp))

        if 'cnpj' in df_s.columns and 'cnpj' in df_e.columns:
            cols_emp = [c for c in df_e.columns.tolist() if c != 'cnpj']
            col_extrair = st.selectbox("Qual coluna da planilha de EMPRESAS quer adicionar?", cols_emp)

            if st.button("Executar Cruzamento (Merge)"):
                df_e_clean = df_e[['cnpj', col_extrair]].drop_duplicates(subset=['cnpj'])
                df_resultado = pd.merge(df_s, df_e_clean, on='cnpj', how='left')
                
                st.success("Cruzamento realizado!")
                st.dataframe(df_resultado.head(15))

                col_u1, col_u2 = st.columns(2)
                # Excel
                output_uni = BytesIO()
                with pd.ExcelWriter(output_uni, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False)
                col_u1.download_button("🔗 Baixar EXCEL", output_uni.getvalue(), "unificado.xlsx")
                
                # CSV
                col_u2.download_button("📄 Baixar CSV", to_csv(df_resultado), "unificado.csv", "text/csv")
        else:
            st.error("⚠️ Ambas precisam ter a coluna 'cnpj'!")