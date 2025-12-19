import streamlit as st
import pandas as pd
from datetime import datetime
import os
import time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Busca RMs | Deloitte", layout="wide")

# --- ESTILO VISUAL (CSS CORPORATIVO) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Open Sans', sans-serif; }
    
    [data-testid="stSidebar"] { background-color: #121212; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 { color: #86BC25 !important; }
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] p { color: #E0E0E0 !important; }
    
    h1, h2, h3 { color: #F0F0F0; }
    
    /* Destaque dos valores */
    [data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-size: 1.8em !important;
        font-weight: 800 !important;
        text-shadow: 0px 0px 10px rgba(134, 188, 37, 0.3);
    }
    [data-testid="stMetricLabel"] { color: #AAAAAA !important; font-size: 1em; }
    
    /* Botão de Atualizar */
    div.stButton > button {
        background-color: #333333;
        color: #86BC25;
        border: 1px solid #86BC25;
        width: 100%;
    }
    div.stButton > button:hover {
        background-color: #86BC25;
        color: #000000;
    }
    </style>
    """, unsafe_allow_html=True)

# --- CACHE INTELIGENTE (MULTI-ABAS) ---
@st.cache_data(ttl=3600) 
def carregar_dados_complexos(arquivo, usar_todas_abas, aba_especifica=None):
    # Carrega a estrutura do arquivo
    xls = pd.ExcelFile(arquivo)
    
    if usar_todas_abas:
        # Lê TODAS as abas e junta numa tabela só
        lista_dfs = []
        for nome_aba in xls.sheet_names:
            df_temp = xls.parse(nome_aba)
            # Converte TODAS as colunas para string para evitar erro de Data no cabeçalho
            df_temp.columns = df_temp.columns.astype(str)
            
            # Cria coluna para identificar a origem
            df_temp['[Origem da Aba]'] = nome_aba 
            lista_dfs.append(df_temp)
        
        # Concatena (Empilha)
        df_final = pd.concat(lista_dfs, ignore_index=True)
        return df_final, xls.sheet_names
    else:
        # Lê apenas uma aba específica
        df_final = xls.parse(aba_especifica)
        # Converte cabeçalho para texto também
        df_final.columns = df_final.columns.astype(str)
        return df_final, xls.sheet_names

# --- TÍTULO PRINCIPAL ---
st.markdown("<h1 style='margin-bottom: 10px;'><span style='color: #86BC25;'>Deloitte.</span> Buscador de RMs</h1>", unsafe_allow_html=True)

# --- BARRA LATERAL ---
st.sidebar.title("Controle de Dados")

if st.sidebar.button("Atualizar Base (Limpar Cache)"):
    st.cache_data.clear()
    st.rerun()

# 1. LOCALIZAR ARQUIVO
try:
    arquivos_na_pasta = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]
    arquivos_na_pasta.sort(key=lambda x: os.path.getmtime(x), reverse=True)
except:
    arquivos_na_pasta = []

arquivo_selecionado = None

aba_local, aba_upload = st.sidebar.tabs(["Pasta Local", "Upload Manual"])

with aba_local:
    if arquivos_na_pasta:
        st.caption("Arquivos encontrados:")
        sel_arquivo = st.selectbox("Arquivo:", arquivos_na_pasta, index=0)
        
        data_mod = datetime.fromtimestamp(os.path.getmtime(sel_arquivo)).strftime('%d/%m %H:%M')
        st.caption(f"Modificado em: {data_mod}")
        
        if st.checkbox("Carregar este arquivo", value=True):
            arquivo_selecionado = sel_arquivo
    else:
        st.warning("Nenhum arquivo Excel encontrado na pasta.")

with aba_upload:
    up = st.file_uploader("Arraste aqui:", type=["xlsx", "xls"])
    if up:
        arquivo_selecionado = up

# --- PROCESSAMENTO ---
if arquivo_selecionado:
    try:
        # Passo 1: Analisar as Abas antes de carregar
        excel_estrutura = pd.ExcelFile(arquivo_selecionado)
        nomes_abas = excel_estrutura.sheet_names
        
        df = None
        
        st.sidebar.markdown("---")
        st.sidebar.markdown("### Seleção de Abas")
        
        # SE TIVER MAIS DE UMA ABA, PERGUNTA
        if len(nomes_abas) > 1:
            modo = st.sidebar.radio("Opções de Leitura:", ["Ler TUDO (Juntar Abas)", "Escolher uma Aba"])
            
            if modo == "Ler TUDO (Juntar Abas)":
                with st.spinner("Fundindo todas as abas..."):
                    df, _ = carregar_dados_complexos(arquivo_selecionado, True)
                st.sidebar.success(f"Fundido: {len(nomes_abas)} abas.")
            else:
                aba_escolhida = st.sidebar.selectbox("Qual aba?", nomes_abas)
                df, _ = carregar_dados_complexos(arquivo_selecionado, False, aba_escolhida)
        else:
            # Se só tem 1 aba, carrega direto
            df, _ = carregar_dados_complexos(arquivo_selecionado, False, nomes_abas[0])

        # --- 2. VISUALIZAR PRÉVIA ---
        with st.expander(f"Visualizar Tabela Carregada ({len(df)} linhas)"):
            st.dataframe(df.head(), use_container_width=True)

        # --- 3. MAPEAMENTO ---
        colunas = df.columns.tolist()
        
        # --- A CORREÇÃO MÁGICA ESTÁ AQUI NA FUNÇÃO ACHAR ---
        def achar(padrao):
            for i, c in enumerate(colunas):
                # str(c) converte DATAS ou NÚMEROS do cabeçalho em TEXTO antes de comparar
                if padrao.lower() in str(c).lower(): return i
            return 0
        # ----------------------------------------------------

        st.sidebar.markdown("---")
        st.sidebar.markdown("### Mapeamento")
        
        col_chave = st.sidebar.selectbox("RMT (Código):", colunas, index=achar("RMT"))
        col_rev   = st.sidebar.selectbox("Revisão:", colunas, index=achar("Revis") if achar("Revis") else achar("R")) 
        if col_rev == colunas[0] and "R" in colunas: col_rev = "R"
            
        col_status= st.sidebar.selectbox("Status:", colunas, index=achar("Status"))
        col_prev  = st.sidebar.selectbox("Previsão:", colunas, index=achar("Previs"))
        col_desc  = st.sidebar.selectbox("Descrição:", colunas, index=achar("Descri"))

        # --- BUSCA ---
        st.divider()
        rm_input = st.text_input("Pesquisar RM:", placeholder="Digite o código (ex: ...-015)").strip()

        if rm_input:
            resultado = df[df[col_chave].astype(str).str.strip().str.contains(rm_input, case=False, na=False)]

            if not resultado.empty:
                dados = resultado.iloc[0]
                
                # Cabeçalho
                st.markdown(f"<h3 style='color:#86BC25; border-bottom: 1px solid #333;'>{dados[col_chave]}</h3>", unsafe_allow_html=True)
                
                # Se veio de fusão de abas, mostra a origem
                if '[Origem da Aba]' in dados:
                    st.caption(f"Encontrado na aba: {dados['[Origem da Aba]']}")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Status Atual", str(dados[col_status]))
                c2.metric("Previsão", str(dados[col_prev]))
                c3.metric("Revisão", str(dados[col_rev]))
                c4.metric("Fornecedor", str(dados.get('Fornecedor', '-')))

                st.markdown(f"**Descrição:** {dados[col_desc]}")

                # Cópia
                st.markdown("### Copiar (Ctrl+C)")
                data_hoje = datetime.now().strftime('%d/%m/%Y')
                v_rev = dados[col_rev] if pd.notna(dados[col_rev]) else ""
                v_stat = dados[col_status] if pd.notna(dados[col_status]) else "-"
                v_prev = dados[col_prev] if pd.notna(dados[col_prev]) else "-"

                texto_final = f"""{data_hoje}
{dados[col_chave]}
Revisão: {v_rev}
Status: {v_stat}
Previsão de entrega: {v_prev}"""
                st.text_area("Texto pronto:", value=texto_final, height=180)

            else:
                st.warning(f"RM '{rm_input}' não encontrada.")
                with st.expander("Gerar linha para cadastro manual"):
                    nov_desc = st.text_input("Descrição")
                    nov_forn = st.text_input("Fornecedor")
                    if st.button("Gerar Linha"):
                         linha = f"{rm_input}\t{nov_desc}\t0\tDiligenciamento\tVocê\tPendente\t{nov_forn}\tA definir"
                         st.code(linha, language="text")

    except Exception as e:
        st.error("Erro ao ler o arquivo.")
        st.write(e)
else:
    st.info("Nenhuma planilha carregada. Verifique a barra lateral.")
