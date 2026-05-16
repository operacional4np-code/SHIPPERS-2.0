import streamlit as st
import pandas as pd
import io
import re
from datetime import date
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES para busca na planilha de dados
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE", 
    "CGB": "CUIABA", 
    "CWB": "CURITIBA", 
    "FLN": "FLORIANOPOLIS", 
    "GYN": "GOIANIA", 
    "MAO": "MANAUS", 
    "POA": "PORTO ALEGRE", 
    "PVH": "PORTO VELHO"
}

st.set_page_config(page_title="New Post - Gerador Word Shippers", layout="wide")
st.title("📄 Gerador de Shippers New Post — Padrão Word (.docx)")
st.subheader("Captura direta dos valores de Fibreboard, Kg G e Total Overpack da planilha")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Informações (.xlsx / .xlsm)", type=["xlsm", "xlsx"])

def formatar_valor_br(valor):
    """Garante a formatação com duas casas decimais e vírgula separando os centavos"""
    try:
        if pd.isna(valor) or valor == "":
            return "0,00"
        return "{:.2f}".format(float(valor)).replace('.', ',')
    except:
        return str(valor).replace('.', ',')

# 2. SELETOR DINÂMICO DE SACAS NA TELA
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    sacas_manuais = {}
    
    colunas_tela = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        with colunas_tela[idx]:
            default_val = 17 if sigla == "POA" else 7
            sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=default_val, step=1, key=f"sacas_{sigla}")

    if file:
        try:
            # Lê a planilha preservando a estrutura original de colunas
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            if st.button("🔢 GERAR DOCUMENTOS WORD (VALORES EXATOS)"):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        # Localizar a linha correta da cidade na planilha carregada
                        linha_dados = None
                        for index, row in df_raw.iterrows():
                            linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
                            if cidade_alvo in linha_texto and "TOTAL" not in linha_texto:
                                # Copia a linha inteira como uma lista simples para acessar os índices das colunas
                                linha_dados = list(row.values)
                                break

                        if linha_dados is not None:
                            try:
                                # CAPTURA DIRETA CONFORME A ESTRUTURA DA PLANILHA NEW POST:
                                # Coluna F (Índice 5) -> Qtd Sacas
                                # Coluna I (Índice 8) -> Fibreboard Boxes
                                # Coluna J (Índice 9) -> Kg G
                                # Coluna K (Índice 10) -> Total Overpack
                                
                                val_fibreboard = linha_dados[8]
                                val_kg_g       = linha_dados[9]
                                val_total_ovp  = linha_dados[10]

                                # Converte para o formato inteiro ou decimal com vírgula padrão Brasil
                                txt_fibreboard = str(int(float(val_fibreboard))) if pd.notnull(val_fibreboard) else "0"
                                txt_kg_g       = formatar_valor_br(val_kg_g)
                                txt_total_ovp  = formatar_valor_br(val_total_ovp)
                                
                                # Monta a string sequencial de marcações (#1 #2 #3...) baseado na saca digitada em tela
                                marcacao = " ".join([f"#{i+1}" for i in range(int(qtd_sacas_escolhida))])

                                # Montagem do dicionário com as tags que vão preencher o Word
                                contexto = {
                                    'FIBREBOARD': txt_fibreboard,
                                    'PESO_G': txt_kg_g,
                                    'TOTAL_OVERPACK': txt_total_ovp,
                                    'MARCACAO': marcacao,
                                    'DATA': date.today().strftime('%d/%m/%Y'),
                                    'QTD_OVERPACK': int(qtd_sacas_escolhida)
                                }

                                # Carrega o template Word correspondente à sigla
                                caminho_template = f"templates/{sigla}-SHIPPER-t.docx"
                                doc = DocxTemplate(caminho_template)
                                doc.render(contexto)
                                
                                # Guarda o documento na memória para empacotar no ZIP
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                zip_file.writestr(f"Shipper_{sigla}.docx", doc_io.getvalue())
                                emitidos.append(sigla)
                                
                            except Exception as e_calculo:
                                erros_cidades.append(f"{sigla} (Erro ao extrair colunas I, J, K: {e_calculo})")
                        else:
                            erros_cidades.append(f"{sigla} (Cidade não encontrada na planilha)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso! Shippers geradas com valores idênticos à referência para: {', '.join(emitidos)}")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name="Shippers_Corrigidas_NewPost.zip",
                        mime="application/zip"
                    )
        except Exception as e:
            st.error(f"Erro geral no processamento da planilha: {e}")
