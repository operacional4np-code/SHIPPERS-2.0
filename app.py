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

# Função de busca segura convertendo para lista
def encontrar_dados_na_planilha(df_raw, busca):
    busca = busca.upper().strip()
    
    # Iteramos sobre o dataframe
    for _, row in df_raw.iterrows():
        # Convertemos a linha para uma lista, assim row_lista[0] é sempre o primeiro item
        row_lista = row.tolist()
        
        # Garantimos que o primeiro item seja string antes de buscar
        destino_planilha = str(row_lista[0]).upper()
        
        if busca in destino_planilha:
            try:
                qtd = row_lista[1]
                peso = row_lista[2]
                return True, int(qtd), float(peso)
            except (ValueError, IndexError, TypeError):
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
            # Lemos o arquivo. Se header=0 estiver causando erro, removemos ele
            # para ler a planilha como uma matriz bruta
            try:
                if file.name.endswith('.csv'):
                    df = pd.read_csv(file, header=0)
                else:
                    df = pd.read_excel(file, header=0)
            except:
                # Caso o header=0 falhe, tenta ler sem cabeçalho
                df = pd.read_excel(file, header=None)
            
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
                        st.warning(f"⚠️ Não encontrei o destino '{sigla}' na coluna 0 da planilha.")
            
            if sucesso:
                st.download_button("📥 BAIXAR ZIP", data=zip_buffer.getvalue(), file_name="Shippers.zip")
