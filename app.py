import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE", 
    "CGB": "CUIABA", 
    "CWB": "CURITIBA", 
    "FLN": "FLORIANOPOLIS", 
    "GYN": "GOIANIA", 
    "MAO": "MANAUS", 
    "POA": "PORTO ALEGRE", 
    "PVH": "PORTO VELHO",
    "POA PRIME": "PRIME-RS PORTO ALEGRE",
    "FLN PRIME": "PRIME-SC FLORIANÓPOLIS"
}

st.set_page_config(page_title="New Post - Gerador Word Shippers", layout="wide")
st.title("📄 Gerador de Shippers New Post")
st.subheader("Cálculo Autônomo Oficial")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas (Ex: CGB, POA, POA PRIME):", value="").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Base)", type=["xlsm", "xlsx"])

def extrair_dados_coleta(df_raw, termo_busca):
    # Procura pelo cabeçalho dinâmico
    linha_cabecalho = None
    idx_destino, idx_qntde, idx_peso = None, None, None
    
    for idx, row in df_raw.iterrows():
        valores = [str(v).strip().upper() for v in row.values if pd.notnull(v)]
        if "DESTINO" in valores and ("QNTDE" in valores or "QNTD" in valores) and "PESO" in valores:
            linha_cabecalho = idx
            valores_linha_lista = [str(v).strip().upper() for v in row.values]
            idx_destino = valores_linha_lista.index("DESTINO")
            idx_qntde = valores_linha_lista.index("QNTDE") if "QNTDE" in valores_linha_lista else valores_linha_lista.index("QNTD")
            idx_peso = valores_linha_lista.index("PESO")
            break
            
    if linha_cabecalho is None: return None, None, None

    for idx in range(linha_cabecalho + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        val_destino = str(row.iloc[idx_destino]).strip().upper() if pd.notnull(row.iloc[idx_destino]) else ""
        
        # Lógica especial para o PRIME
        if "PRIME" in termo_busca:
            if termo_busca in val_destino:
                pass # Achou o destino específico
            else:
                continue
        else:
            if "TOTAL" in val_destino or val_destino == "" or val_destino.isdigit(): continue
            # Limpeza padrão para outros destinos
            destino_limpo = val_destino.replace("AGF", "").replace(" MT", "").replace(" MS", "").replace(" PR", "").replace(" SC", "").replace(" GO", "").replace(" AM", "").replace(" RS", "").replace(" RO", "").strip()
            if termo_busca not in destino_limpo and destino_limpo not in termo_busca: continue

        try:
            val_q = row.iloc[idx_qntde]
            qtd_volumes = int(float(str(val_q).replace(',', '.').strip()))
            val_p = row.iloc[idx_peso]
            peso_original = float(str(val_p).replace(',', '.').strip())
            return termo_busca, qtd_volumes, peso_original
        except: continue
    return None, None, None

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    st.markdown("### 3. Informe a quantidade de sacas:")
    cols = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        with cols[idx]:
            sacas_manuais[sigla] = st
