import streamlit as st
import pandas as pd
import io
import os
from datetime import date
from docxtpl import DocxTemplate
from zipfile import ZipFile

# Configuração da página
st.set_page_config(page_title="Gerador de Shippers", layout="wide")
st.title("📄 Gerador de Shippers")

# 1. ENTRADAS
siglas_input = st.text_input("Destinos (Ex: CGR, POA, MANAUS):", value="").upper().strip()
file = st.file_uploader("Carregue a Planilha (CSV ou XLSX)", type=["csv", "xlsx"])

# Função de busca inteligente
def encontrar_dados_na_planilha(df_raw, busca):
    busca_limpa = "".join(busca.upper().split())
    
    for _, row in df_raw.iterrows():
        # Converte toda a linha em string para busca
        linha_str = "".join([str(v).upper() for v in row.values if pd.notnull(v)])
        
        if busca_limpa in linha_str:
            # Captura os números da linha (Qtd, Peso)
            numeros = [float(v) for v in row.values if isinstance(v, (int, float))]
            if len(numeros) >= 2:
                return True, int(numeros[0]), float(numeros[1])
    return False, None, None

# 2. SELETOR DE SACAS
if siglas_input and file:
    lista_siglas = [s.strip() for s in siglas_input.split(",")]
    
    st.markdown("### 3. Informe a quantidade de sacas:")
    cols = st.columns(len(lista_siglas))
    sacas = {}

    for idx, sigla in enumerate(lista_siglas):
        with cols[idx]:
            # A key dinâmica com o nome do arquivo garante que o campo limpe ao trocar o arquivo
            key_campo = f"s_input_{sigla}_{file.name}"
            sacas[sigla] = st.number_input(
                f"Sacas para {sigla}:", 
                min_value=1, 
                value=None, 
                key=key_campo,
                placeholder="0"
            )

    # 3. GERAÇÃO
    # Verifica se todos os campos de sacas foram preenchidos
    if all(s is not None for s in sacas.values()):
        if st.button("🔢 GERAR ARQUIVOS"):
            df = pd.read_csv(file, header=None) if file.name.endswith('.csv') else pd.read_excel(file, header=None)
            
            zip_buffer = io.BytesIO()
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
            sucesso = False
            with ZipFile(zip_buffer, "w") as zip_file:
                for sigla in lista_siglas:
                    achou, qtd, peso = encontrar_dados_na_planilha(df, sigla)
                    
                    if achou:
                        sigla_segura = sigla.replace(" ", "_")
                        caminho_template = os.path.join(base_dir, "templates", f"{sigla_segura}-SHIPPER-t.docx")
                        
                        if os.path.exists(caminho_template):
                            doc = DocxTemplate(caminho_template)
                            doc.render({
                                'DATA': date.today().strftime('%d/%m/%Y'), 
                                'QTD': qtd, 
                                'PESO': peso,
                                'SACAS': sacas[sigla]
                            })
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            zip_file.writestr(f"Shipper_{sigla_
