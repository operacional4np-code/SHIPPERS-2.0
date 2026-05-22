import streamlit as st
import pandas as pd
import io
import os
from datetime import date
from decimal import Decimal
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE DESTINOS (Nomes exatos como aparecem na sua planilha)
MAPA_DESTINOS = {
    "POA": "AGF PORTO ALEGRE",
    "FLN": "AGF FLORIANOPOLIS",
    "CGR": "AGF CAMPO GRANDE",
    "CGB": "AGF CUIABA",
    "POA PRIME": "PRIME - RS PORTO ALEGRE",
    "FLN PRIME": "PRIME - SC FLORIANOPOLIS"
}

st.title("📄 Gerador de Shippers")

siglas_input = st.text_input("Destinos (Ex: CGR, POA, POA PRIME):", value="").upper().strip()
file = st.file_uploader("Carregue a Planilha (CSV ou XLSX)", type=["csv", "xlsx"])

def extrair_dados_da_linha_simples(df_raw, nome_cidade):
    """
    Busca a linha que contém o nome da cidade e retorna os números encontrados na mesma linha.
    """
    for _, row in df_raw.iterrows():
        linha_str = " ".join([str(v) for v in row.values if pd.notnull(v)]).upper()
        if nome_cidade.upper() in linha_str:
            # Pega todos os números da linha (Qtd, Peso, etc)
            numeros = [float(v) for v in row.values if isinstance(v, (int, float))]
            if len(numeros) >= 2:
                # Retorna: nome achado, Qtd, Peso
                return nome_cidade, int(numeros[0]), float(numeros[1])
    return None, None, None

if siglas_input and file:
    lista_siglas = [s.strip() for s in siglas_input.split(",")]
    
    # Criar campos para sacas
    sacas = {}
    for sigla in lista_siglas:
        sacas[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=1)

    if st.button("GERAR ARQUIVOS"):
        # Detecta se é CSV ou Excel
        if file.name.endswith('.csv'):
            df = pd.read_csv(file, header=None)
        else:
            df = pd.read_excel(file, header=None)
            
        zip_buffer = io.BytesIO()
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        with ZipFile(zip_buffer, "w") as zip_file:
            for sigla in lista_siglas:
                nome_busca = MAPA_DESTINOS.get(sigla, sigla)
                _, qtd, peso = extrair_dados_da_linha_simples(df, nome_busca)
                
                if peso:
                    sigla_segura = sigla.replace(" ", "_")
                    caminho_template = os.path.join(base_dir, "templates", f"{sigla_segura}-SHIPPER-t.docx")
                    
                    if os.path.exists(caminho_template):
                        doc = DocxTemplate(caminho_template)
                        doc.render({'DATA': date.today().strftime('%d/%m/%Y'), 'QTD': qtd, 'PESO': peso})
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"Shipper_{sigla_segura}.docx", doc_io.getvalue())
                        st.success(f"Gerado: {sigla}")
                    else:
                        st.error(f"Template não existe: {caminho_template}")
                else:
                    st.warning(f"Não achei dados para: {nome_busca}")
        
        st.download_button("BAIXAR ZIP", data=zip_buffer.getvalue(), file_name="Shippers.zip")
