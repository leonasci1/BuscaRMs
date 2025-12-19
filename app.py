import streamlit as st
import pandas as pd
from datetime import datetime
import os
import time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Busca RMs | Deloitte", layout="wide")

# --- ESTILO VISUAL (CSS) ---
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

# --- CACHE ---
@st.cache_data(ttl=3600) 
def carregar_excel(caminho_ou_arquivo):
    return pd.read_excel(caminho_ou_arquivo)

# --- TÍTULO PRINCIPAL ---
st.markdown("<h1 style='margin-bottom: 10px;'><span style='color: #86BC25;'>Deloitte.</span> Buscador de RMs</h1>", unsafe_allow_html=True)

# --- BARRA LATERAL (Fixa no Topo) ---
st.sidebar.title("📂 Controle de Dados")

# Botão de Reset (Sempre visível)
if st.sidebar.button("🔄 Atualizar Base (Limpar Cache)"):
    st.cache_data.clear()
    st.rerun()

# --- LÓGICA INTELIGENTE DE ARQUIVOS ---
# Procura qualquer arquivo Excel na pasta atual
arquivos_na_pasta = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.xls'))]

# Ordena pelo mais recente (Data de modificação)
arquivos_na_pasta.sort(key=lambda x: os.path.getmtime(x), reverse=True)

df = None
arquivo_escolhido = None

# Abas de seleção
aba_local, aba_upload = st.sidebar.tabs(["📁 Pasta Local", "📤 Upload Manual"])

with aba_local:
    if arquivos_na_pasta:
        st.caption("Arquivos encontrados na pasta (do mais novo para o antigo):")
        
        # O usuário pode escolher, mas já vem selecionado o primeiro (mais recente)
        arquivo_selecionado_nome = st.selectbox(
            "Selecione o arquivo:", 
            arquivos_na_pasta,
            index=0
        )
        
        # Mostra data do arquivo
        data_mod = datetime.fromtimestamp(os.path.getmtime(arquivo_selecionado_nome)).strftime('%d/%m/%Y %H:%M')
        st.caption(f"📅 Modificado em: {data_mod}")
        
        # Checkbox para confirmar carregamento
        if st.checkbox("Carregar este arquivo", value=True):
            try:
                arquivo_escolhido = arquivo_selecionado_nome
                df = carregar_excel(arquivo_escolhido)
            except Exception as e:
                st.error(f"Erro ao ler {arquivo_selecionado_nome}")
    else:
        st.warning("⚠️ Nenhum arquivo Excel (.xlsx) encontrado nesta pasta.")

with aba_upload:
    up = st.file_uploader("Ou arraste um arquivo aqui:", type=["xlsx", "xls"])
    if up:
        df = carregar_excel(up)
        arquivo_escolhido = "Upload Manual"

# --- PROCESSAMENTO DOS DADOS ---
if df is not None:
    try:
        colunas = df.columns.tolist()
        
        # Função para adivinhar colunas (Case Insensitive)
        def achar(padrao):
            for i, c in enumerate(colunas):
                if padrao.lower() in c.lower(): return i
            return 0

        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🎯 Mapeamento")
        
        # Selections com índices automáticos
        col_chave = st.sidebar.selectbox("RMT (Código):", colunas, index=achar("RMT"))
        col_rev   = st.sidebar.selectbox("Revisão:", colunas, index=achar("Revis") if achar("Revis") else achar("R")) 
        # Tenta achar 'R' sozinho se não achar 'Revisão'
        if col_rev == colunas[0] and "R" in colunas: 
            col_rev = "R"
            
        col_status= st.sidebar.selectbox("Status:", colunas, index=achar("Status"))
        col_prev  = st.sidebar.selectbox("Previsão:", colunas, index=achar("Previs"))
        col_desc  = st.sidebar.selectbox("Descrição:", colunas, index=achar("Descri"))

        # --- ÁREA DE BUSCA ---
        st.divider()
        rm_input = st.text_input("Pesquisar RM:", placeholder="Digite o código (ex: ...-015)").strip()

        if rm_input:
            resultado = df[df[col_chave].astype(str).str.strip().str.contains(rm_input, case=False, na=False)]

            if not resultado.empty:
                dados = resultado.iloc[0]
                
                # Cabeçalho do Resultado
                st.markdown(f"<h3 style='color:#86BC25; border-bottom: 1px solid #333;'>{dados[col_chave]}</h3>", unsafe_allow_html=True)
                
                # Métricas
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Status Atual", str(dados[col_status]))
                c2.metric("Previsão", str(dados[col_prev]))
                c3.metric("Revisão", str(dados[col_rev]))
                c4.metric("Fornecedor", str(dados.get('Fornecedor', '-')))

                st.markdown(f"**Descrição:** {dados[col_desc]}")

                # --- ÁREA DE CÓPIA (VERTICAL) ---
                st.markdown("### 📋 Copiar (Ctrl+C)")
                data_hoje = datetime.now().strftime('%d/%m/%Y')
                
                # Tratamento de valores nulos
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
                st.warning(f"⚠️ RM '{rm_input}' não encontrada em '{arquivo_escolhido}'.")
                with st.expander("Gerar linha para cadastro manual"):
                    nov_desc = st.text_input("Descrição")
                    nov_forn = st.text_input("Fornecedor")
                    if st.button("Gerar Linha"):
                         linha = f"{rm_input}\t{nov_desc}\t0\tDiligenciamento\tVocê\tPendente\t{nov_forn}\tA definir"
                         st.code(linha, language="text")

    except Exception as e:
        st.error("Erro ao ler o arquivo selecionado.")
        st.write(e)
else:
    st.info("👈 Nenhuma planilha carregada. Verifique a barra lateral.")