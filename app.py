import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES - Alinhado com a planilha de coleta
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE", 
    "CGB": "CUIABA", 
    "CWB": "CURITIBA", 
    "FLN": "FLORIANOPOLIS", 
    "GYN": "GOIANIA", 
    "MAO": "MANAUS", 
    "POA": "PORTO ALEGRE", 
    "PVH": "PORTO VELHO",
    "POA PRIME": "PRIME - RS PORTO ALEGRE",
    "FLN PRIME": "PRIME - SC FLORIANOPOLIS"
}

st.set_page_config(page_title="New Post - Gerador Word Shippers", layout="wide")
st.title("📄 Gerador de Shippers New Post")
st.subheader("Cálculo Autônomo Oficial")

# 1. ENTRADAS DE DADOS
st.info("💡 Separe cada destino por vírgula. Você pode usar espaços normalmente (Ex: CGB, POA PRIME, FLN PRIME)")
siglas_input = st.text_input("1. Digite os Destinos separados por vírgula:", value="").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Base)", type=["xlsm", "xlsx"])

def extrair_dados_coleta(df_raw, termo_busca):
    """
    Localiza a tabela dinamicamente na planilha, acha o cabeçalho correto,
    e captura os dados fazendo um cruzamento exato flexível.
    """
    linha_cabecalho = None
    idx_destino, idx_qntde, idx_peso = None, None, None
    
    for idx, row in df_raw.iterrows():
        valores = [str(v).strip().upper() for v in row.values if pd.notnull(v)]
        if "DESTINO" in valores and ("QNTDE" in valores or "QNTD" in valores) and "PESO" in valores:
            linha_cabecalho = idx
            valores_linha_lista = [str(v).strip().upper() for v in row.values]
            idx_destino = valores_linha_lista.index("DESTINO")
            
            if "QNTDE" in valores_linha_lista:
                idx_qntde = valores_linha_lista.index("QNTDE")
            elif "QNTD" in valores_linha_lista:
                idx_qntde = valores_linha_lista.index("QNTD")
                
            idx_peso = valores_linha_lista.index("PESO")
            break
            
    if linha_cabecalho is None:
        for idx, row in df_raw.iterrows():
            linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
            if termo_busca in linha_texto and "TOTAL" not in linha_texto:
                numeros = []
                for val in row.values:
                    if pd.notnull(val) and isinstance(val, (int, float)):
                        numeros.append(float(val))
                if len(numeros) >= 2:
                    return termo_busca, int(numeros[0]), float(numeros[1])
        return None, None, None

    for idx in range(linha_cabecalho + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        val_destino = str(row.iloc[idx_destino]).strip().upper() if pd.notnull(row.iloc[idx_destino]) else ""
        
        if "TOTAL" in val_destino or val_destino == "" or val_destino.isdigit():
            continue
            
        val_destino_norm = "".join(val_destino.replace("-", " ").split())
        termo_busca_norm = "".join(termo_busca.replace("-", " ").split())
        
        if termo_busca_norm in val_destino_norm or val_destino_norm in termo_busca_norm:
            try:
                val_q = row.iloc[idx_qntde]
                qtd_volumes = int(float(str(val_q).replace(',', '.').strip())) if not isinstance(val_q, (int, float)) else int(val_q)
                
                val_p = row.iloc[idx_peso]
                peso_original = float(str(val_p).replace(',', '.').strip()) if not isinstance(val_p, (int, float)) else float(val_p)
                
                return val_destino, qtd_volumes, peso_original
            except:
                continue
                
    return None, None, None

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    cols = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        # Correção interna: Remove os espaços apenas para criar a chave estável do Streamlit
        sigla_chave_limpa = "".join(sigla.split())
        with cols[idx]:
            sacas_manuais[sigla] = st.number_input(
                f"Sacas para {sigla}:", 
                min_value=1, 
                value=None, 
                step=1, 
                key=f"sacas_{sigla_chave_limpa}"
            )

    todas_sacas_preenchidas = all(saca is not None for saca in sacas_manuais.values())

    if file and todas_sacas_preenchidas:
        try:
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            st.markdown("---")
            if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []
                dados_conferencia = []

                mapa_normalizado = {"".join(k.split()): v for k, v in MAPA_DESTINOS.items()}

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        sigla_chave_limpa = "".join(sigla.split())
                        cidade_alvo = mapa_normalizado.get(sigla_chave_limpa, sigla)
                        
                        qtd_sacas_escolhida = sacas_manuais.get(sigla)
                        
                        # Proteção redundante contra recarregamento do Streamlit
                        if qtd_sacas_escolhida is None:
                            continue

                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            fracao_fib = float(q_volumes) / float(qtd_sacas_escolhida)
                            decimal_part = fracao_fib - math.floor(fracao_fib)
                            if decimal_part >= 0.50:
                                i_fibreboard = math.floor(fracao_fib) + 1
                            else:
                                i_fibreboard = math.floor(fracao_fib)
