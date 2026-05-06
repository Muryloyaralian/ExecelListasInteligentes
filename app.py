import streamlit as st
import pandas as pd
from io import BytesIO
import re

# 1. Configuração da Página (Deve ser a primeira linha de comando Streamlit)
st.set_page_config(page_title="Limpador de Excel Pro", layout="wide")

# --- FUNÇÕES DE UTILIDADE ---
def limpar_numero(tel):
    """Remove qualquer caractere que não seja número."""
    if pd.isna(tel): return ""
    return re.sub(r'\D', '', str(tel))

def normalizar_colunas(df):
    """Remove espaços e padroniza cabeçalhos para minúsculo."""
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df

# --- INTERFACE PRINCIPAL ---
st.title("✂️ Ferramenta de Tratamento e Unificação de Dados")
st.write("Siga os passos em cada aba para processar suas listas de leads e sócios.")

# Criação das Abas
tab1, tab2 = st.tabs(["🎯 1. Ajustador Inteligente", "🔗 2. Complementar Dados (VLOOKUP)"])

# --- ABA 1: AJUSTADOR INTELIGENTE ---
with tab1:
    st.header("Ajuste de Colunas e Formatação")
    st.write("Ideal para limpar um arquivo final, unir telefones e tratar duplicatas.")
    
    arquivo_ajuste = st.file_uploader("Suba a planilha para formatar (.xlsx)", type=["xlsx"], key="ajuste_single")

    if arquivo_ajuste:
        # Carregar e normalizar
        df = pd.read_excel(arquivo_ajuste)
        df = normalizar_colunas(df)
        
        # Guardamos o total original para o cálculo das métricas
        total_original = len(df)

        # Configurações na Barra Lateral
        st.sidebar.header("⚙️ Configurações Aba 1")
        aj_duplicatas = st.sidebar.checkbox("Remover linhas 100% idênticas", key="dup_tab1")
        aj_socios = st.sidebar.checkbox("Sequenciar Sócios (Nome 2, 3...)", key="soc_tab1")
        unir_tels = st.sidebar.checkbox("Unir DDD + Números", key="tel_tab1")

        # Lógica de Unir Telefones
        if unir_tels:
            pares = [
                ('ddd_cel1', 'cel1', 'celular_1_completo'),
                ('ddd_cel2', 'cel2', 'celular_2_completo'),
                ('ddd_cel3', 'cel3', 'celular_3_completo'),
                ('ddd_tel1', 'tel1', 'telefone_1_completo'),
                ('ddd_tel2', 'tel2', 'telefone_2_completo')
            ]
            col_remover = []
            for ddd, num, nova in pares:
                if ddd in df.columns and num in df.columns:
                    df[nova] = df.apply(lambda x: limpar_numero(x[ddd]) + limpar_numero(x[num]), axis=1)
                    col_remover.extend([ddd, num])
            
            # Remove as colunas originais para limpar a visão
            df = df.drop(columns=[c for c in col_remover if c in df.columns])
            st.sidebar.success("Telefones unidos!")

        # Seleção de Colunas
        st.subheader("Escolha as colunas para o resultado final")
        colunas_disponiveis = df.columns.tolist()
        selecionadas = st.multiselect("Manter colunas:", colunas_disponiveis, default=colunas_disponiveis)

        if selecionadas:
            df_final = df[selecionadas].copy()

            # Aplicar remoção de duplicatas e contar
            linhas_removidas = 0
            if aj_duplicatas:
                antes_dup = len(df_final)
                df_final = df_final.drop_duplicates()
                depois_dup = len(df_final)
                linhas_removidas = antes_dup - depois_dup

            # Aplicar sequenciamento de sócios
            if aj_socios and 'nome' in df_final.columns:
                df_final['count'] = df_final.groupby('nome').cumcount() + 1
                df_final['nome'] = df_final.apply(
                    lambda x: f"{x['nome']} {x['count']}" if x['count'] > 1 else x['nome'], 
                    axis=1
                )
                df_final = df_final.drop(columns=['count'])

            # --- EXIBIÇÃO DE MÉTRICAS ---
            st.markdown("---")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Original", f"{total_original} linhas")
            m2.metric("Total Final", f"{len(df_final)} linhas")
            
            if aj_duplicatas:
                m3.metric("Linhas Removidas", f"{linhas_removidas}", 
                          delta=f"-{linhas_removidas}" if linhas_removidas > 0 else 0, 
                          delta_color="inverse")
            else:
                m3.info("Filtro de duplicatas desativado.")
            st.markdown("---")

            st.write("### Prévia do Arquivo Ajustado:")
            st.dataframe(df_final.head(10))
            
            # Preparar Download
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            
            st.download_button(
                label="🚀 Baixar Arquivo Ajustado",
                data=output.getvalue(),
                file_name="planilha_ajustada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# --- ABA 2: UNIFICADOR / COMPLEMENTAR DADOS ---
with tab2:
    st.header("🔗 Cruzamento de Dados por CNPJ")
    st.write("Traga informações da planilha de Empresas para a sua planilha de Sócios.")
    
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Planilha de Sócios")
        arq_soc = st.file_uploader("Arquivo Base (Sócios)", type=["xlsx"], key="f_socios")
    with c2:
        st.subheader("Planilha de Empresas")
        arq_emp = st.file_uploader("Arquivo de Consulta (Empresas)", type=["xlsx"], key="f_empresas")

    if arq_soc and arq_emp:
        df_s = normalizar_colunas(pd.read_excel(arq_soc))
        df_e = normalizar_colunas(pd.read_excel(arq_emp))

        if 'cnpj' in df_s.columns and 'cnpj' in df_e.columns:
            st.info("CNPJ encontrado em ambas as planilhas!")
            
            cols_emp = [c for c in df_e.columns.tolist() if c != 'cnpj']
            col_extrair = st.selectbox("Qual coluna da planilha de EMPRESAS você quer adicionar aos SÓCIOS?", cols_emp)

            if st.button("Executar Cruzamento (Merge)"):
                df_e_clean = df_e[['cnpj', col_extrair]].drop_duplicates(subset=['cnpj'])
                df_resultado = pd.merge(df_s, df_e_clean, on='cnpj', how='left')
                
                st.success("Cruzamento realizado com sucesso!")
                st.dataframe(df_resultado.head(15))

                output_uni = BytesIO()
                with pd.ExcelWriter(output_uni, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False)
                
                st.download_button(
                    label="🔗 Baixar Lista Complementada",
                    data=output_uni.getvalue(),
                    file_name="socios_e_empresas_unificados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("⚠️ Ambas as planilhas precisam ter uma coluna chamada 'cnpj'!")