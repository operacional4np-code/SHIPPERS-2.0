import streamlit as st
import pandas as pd
import io
import math
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

st.set_page_config(page_title="Gerador de Shippers", layout="wide")
st.title("📄 Gerador de Shippers New Post")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas (Ex: CGB, POA, POA PRIME):", value="").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Base)", type=["xlsm", "xlsx"])

def extrair_dados_coleta(df_raw, termo_busca):
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
        
        # Filtro especial para termos PRIME
        if "PRIME" in termo_busca:
            if termo_busca not in val_destino: continue
        else:
            if "TOTAL" in val_destino or val_destino == "" or val_destino.isdigit(): continue
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

# 2. SELETOR DE SACAS (Exibição forçada)
sacas_manuais = {}
lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]

if lista_siglas:
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    cols = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        with cols[idx]:
            sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=None, step=1, key=f"s_{sigla}")

    # 3. GERAÇÃO
    todas_sacas = all(s is not None for s in sacas_manuais.values())
    if file and todas_sacas:
        if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            zip_buffer = io.BytesIO()
            emitidos, erros = [], []

            with ZipFile(zip_buffer, "w") as zip_file:
                for sigla in lista_siglas:
                    cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                    _, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                    if p_original:
                        f_sacas = Decimal(str(sacas_manuais[sigla]))
                        d_peso_original = Decimal(str(p_original))
                        g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                        fracao_fib = float(q_volumes) / float(sacas_manuais[sigla])
                        i_fib = Decimal(str(max(1, math.floor(fracao_fib) + (1 if (fracao_fib - math.floor(fracao_fib)) >= 0.5 else 0))))
                        
                        # Cálculo J
                        base_j = float(g_peso_corrigido / f_sacas / i_fib)
                        j_inicio = Decimal(f"{max(0.01, math.floor(base_j * 100) / 100 - 0.50):.2f}")
                        perfeito_j = j_inicio
                        for a in range(1500):
                            j_teste = j_inicio + (Decimal(str(a)) * Decimal('0.01'))
                            if (j_teste * i_fib * f_sacas) - g_peso_corrigido >= 0:
                                perfeito_j = j_teste
                                break
                        
                        sigla_arq = sigla.replace(" ", "_")
                        contexto = {
                            'FIBREBOARD': str(int(i_fib)),
                            'PESO_G': "{:.2f}".format(perfeito_j).replace('.', ','),
                            'TOTAL_OVERPACK': "{:.2f}".format(perfeito_j * i_fib).replace('.', ','),
                            'MARCACAO': " ".join([f"#{i+1}" for i in range(int(sacas_manuais[sigla]))]),
                            'DATA': date.today().strftime('%d/%m/%Y'),
                            'QTD_OVERPACK': int(sacas_manuais[sigla])
                        }
                        
                        try:
                            doc = DocxTemplate(f"templates/{sigla_arq}-SHIPPER-t.docx")
                            doc.render(contexto)
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            zip_file.writestr(f"Shipper_{sigla_arq}.docx", doc_io.getvalue())
                            emitidos.append(sigla)
                        except: erros.append(f"{sigla} (Erro no Template)")
                    else: erros.append(f"{sigla} (Dados não achados)")

            if emitidos:
                st.success("✅ Sucesso!")
                st.download_button("📥 BAIXAR ZIP", data=zip_buffer.getvalue(), file_name="Shippers_Final.zip")
            for err in erros: st.warning(f"⚠️ {err}")
