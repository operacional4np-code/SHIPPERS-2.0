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

# Função de busca segura usando os nomes das colunas
def encontrar_dados_na_planilha(df_raw, busca):
    busca = busca.upper().strip()
    # Itera sobre as linhas do DataFrame
    for _, row in df_raw.iterrows():
        try:
            # Acessa os dados pelo nome da coluna conforme sua planilha
            destino_planilha = str(row['DESTINO']).upper()
            if busca in destino_planilha:
                qtd = row['QNTDE']
                peso = row['PESO']
                return True, int(qtd), float(peso)
        except (KeyError, ValueError, TypeError):
            continue
    return False, None, None

if siglas_input and file:
    lista_siglas = [s.strip() for s in siglas_input.split(",")]
    
    st.markdown("### 3. Informe a quantidade de sacas:")
    cols = st.columns(len(lista_siglas))
    sacas = {}

    for idx, sigla in enumerate(lista_siglas):
        with cols[idx]:
            key_campo = f"s_input_{sigla}_{file.name}"
            sacas[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=None, key=key_campo, placeholder="0")

    if all(s is not None for s in sacas.values()):
        if st.button("🔢 GERAR ARQUIVOS"):
            # O parâmetro header=0 indica que a primeira linha contém os nomes das colunas
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, header=0)
            else:
                df = pd.read_excel(file, header=0)
            
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
                            doc.render({'DATA': date.today().strftime('%d/%m/%Y'), 'QTD': qtd, 'PESO': peso, 'SACAS': sacas[sigla]})
                            doc_io = io.BytesIO()
                            doc.save(doc_io)
                            zip_file.writestr(f"Shipper_{sigla_segura}.docx", doc_io.getvalue())
                            st.success(f"✅ Gerado: {sigla}")
                            sucesso = True
                        else:
                            st.error(f"❌ Template não encontrado: {caminho_template}")
                    else:
                        st.warning(f"⚠️ Não achei o destino '{sigla}' na coluna 'DESTINO' da planilha.")
            
            if sucesso:
                st.download_button("📥 BAIXAR ZIP", data=zip_buffer.getvalue(), file_name="Shippers.zip")
