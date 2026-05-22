import streamlit as st
import pandas as pd
import io
import math
import os
from datetime import date
from decimal import Decimal
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE", "CGB": "CUIABA", "CWB": "CURITIBA", 
    "FLN": "FLORIANOPOLIS", "GYN": "GOIANIA", "MAO": "MANAUS", 
    "POA": "PORTO ALEGRE", "PVH": "PORTO VELHO",
    "POA PRIME": "PRIME - RS PORTO ALEGRE",
    "FLN PRIME": "PRIME - SC FLORIANOPOLIS"
}

st.set_page_config(page_title="Gerador New Post", layout="wide")
st.title("📄 Gerador de Shippers")

siglas_input = st.text_input("Destinos (Ex: CGB, POA PRIME, FLN PRIME):", value="").upper().strip()
file = st.file_uploader("Carregue a Planilha de Coleta", type=["xlsm", "xlsx"])

# Função de busca mais robusta
def extrair_dados_coleta(df_raw, termo_busca):
    # Procura em todas as células por uma correspondência parcial
    for _, row in df_raw.iterrows():
        linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
        # Se achar o termo na linha, pega os dados
        if "".join(termo_busca.split()) in "".join(linha_texto.split()):
            # Tenta encontrar números na linha (Qtd e Peso)
            numeros = [float(v) for v in row.values if isinstance(v, (int, float))]
            if len(numeros) >= 2:
                return termo_busca, int(numeros[0]), float(numeros[1])
    return None, None, None

sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    for idx, sigla in enumerate(lista_siglas):
        sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=None, key=f"s_{idx}")

    if file and all(s is not None for s in sacas_manuais.values()):
        if st.button("🔢 CALCULAR E GERAR"):
            df_raw = pd.read_excel(file, header=None)
            zip_buffer = io.BytesIO()
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
            sucesso = False
            with ZipFile(zip_buffer, "w") as zip_file:
                for sigla in lista_siglas:
                    cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                    _, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)
                    
                    if p_original is None:
                        st.warning(f"⚠️ Não encontrei dados para: {sigla} (Buscando por: {cidade_alvo})")
                        continue
                    
                    # Se achou, gera o arquivo
                    sigla_segura = sigla.replace(" ", "_")
                    caminho_template = os.path.join(base_dir, "templates", f"{sigla_segura}-SHIPPER-t.docx")
                    
                    if os.path.exists(caminho_template):
                        doc = DocxTemplate(caminho_template)
                        # Preenche com dados básicos para testar
                        doc.render({'DATA': date.today().strftime('%d/%m/%Y')})
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"Shipper_{sigla_segura}.docx", doc_io.getvalue())
                        sucesso = True
                    else:
                        st.error(f"❌ Template não encontrado: {caminho_template}")

            if sucesso:
                st.download_button("📥 BAIXAR ZIP", data=zip_buffer.getvalue(), file_name="Shippers.zip")
            else:
                st.error("Nenhum arquivo pôde ser gerado. Verifique os avisos acima.")
