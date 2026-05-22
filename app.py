import streamlit as st
import pandas as pd
import io
import os
from datetime import date
from docxtpl import DocxTemplate
from zipfile import ZipFile

st.set_page_config(page_title="Gerador de Shippers", layout="wide")
st.title("📄 Gerador de Shippers")

# 1. ENTRADAS
siglas_input = st.text_input("Destinos (Ex: CGR, POA, MANAUS):", value="").upper().strip()
file = st.file_uploader("Carregue a Planilha (CSV ou XLSX)", type=["csv", "xlsx"])

# Função de busca inteligente (ignora espaços e busca parcial)
def encontrar_dados_na_planilha(df_raw, busca):
    busca_limpa = "".join(busca.upper().split()) # Remove espaços da sua busca
    
    for _, row in df_raw.iterrows():
        # Transforma a linha toda em uma string sem espaços
        linha_str = "".join([str(v).upper() for v in row.values if pd.notnull(v)])
        
        # Se a palavra que você buscou estiver contida na linha, achamos!
        if busca_limpa in linha_str:
            numeros = [float(v) for v in row.values if isinstance(v, (int, float))]
            if len(numeros) >= 2:
                return True, int(numeros[0]), float(numeros[1])
    return False, None, None

if siglas_input and file:
    lista_siglas = [s.strip() for s in siglas_input.split(",")]
    
    sacas = {}
    for sigla in lista_siglas:
        sacas[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=1)

    if st.button("GERAR ARQUIVOS"):
        df = pd.read_csv(file, header=None) if file.name.endswith('.csv') else pd.read_excel(file, header=None)
            
        zip_buffer = io.BytesIO()
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
        with ZipFile(zip_buffer, "w") as zip_file:
            for sigla in lista_siglas:
                # Tenta achar usando o nome que você digitou
                achou, qtd, peso = encontrar_dados_na_planilha(df, sigla)
                
                if achou:
                    sigla_segura = sigla.replace(" ", "_")
                    caminho_template = os.path.join(base_dir, "templates", f"{sigla_segura}-SHIPPER-t.docx")
                    
                    if os.path.exists(caminho_template):
                        doc = DocxTemplate(caminho_template)
                        doc.render({'DATA': date.today().strftime('%d/%m/%Y'), 'QTD': qtd, 'PESO': peso})
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        zip_file.writestr(f"Shipper_{sigla_segura}.docx", doc_io.getvalue())
                        st.success(f"✅ Encontrado e gerado: {sigla}")
                    else:
                        st.error(f"❌ Template não encontrado: {caminho_template}")
                else:
                    st.warning(f"⚠️ Não achei dados para: '{sigla}'. Tente digitar apenas uma parte do nome (ex: apenas 'PORTO' ou apenas 'MANAUS')")
        
        st.download_button("📥 BAIXAR ZIP", data=zip_buffer.getvalue(), file_name="Shippers.zip")
