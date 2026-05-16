import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES para busca na Coluna A da planilha
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
st.subheader("Cálculo automatizado de sementes e distribuição por sacas")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Dinâmica)", type=["xlsm", "xlsx"])

def extrair_peso_coluna_c(linha_row):
    """Captura com segurança o peso total contido na Coluna C (Índice 2)"""
    if len(linha_row) > 2:
        val = linha_row[2]
        if pd.notnull(val) and not isinstance(val, str):
            return float(val)
        if isinstance(val, str):
            txt_limpo = re.sub(r'[^\d.,]', '', val).strip()
            if "," in txt_limpo and "." in txt_limpo:
                txt_limpo = txt_limpo.replace(".", "").replace(",", ".")
            elif "," in txt_limpo:
                txt_limpo = txt_limpo.replace(",", ".")
            try:
                return float(txt_limpo)
            except:
                pass
    return 0.0

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
            # Lê a planilha sem colunas fantasmas
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            if st.button("🔢 CALCULAR E GERAR ARQUIVOS DO WORD"):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        # Localizar a linha da cidade buscando estritamente na Coluna A (Índice 0)
                        linha_dados = None
                        for index, row in df_raw.iterrows():
                            celula_a = str(row.values[0]).upper() if len(row.values) > 0 and pd.notnull(row.values[0]) else ""
                            if cidade_alvo in celula_a:
                                linha_dados = list(row.values)
                                break

                        if linha_dados is not None:
                            # Pega o peso total da planilha (Coluna C)
                            peso_total_planilha = extrair_peso_coluna_c(linha_dados)
                            
                            if peso_total_planilha == 0.0:
                                erros_cidades.append(f"{sigla} (Peso total zerado ou não encontrado na Coluna C)")
                                continue

                            # Transformando em Decimal para precisão bancária/IATA
                            g7_peso_real = Decimal(str(peso_total_planilha))
                            f7_qtd_sacas = Decimal(str(qtd_sacas_escolhida))
                            
                            # MÁGICA DOS CÁLCULOS DA RETAGUARDA (Igual à referência do IATA)
                            if sigla == "CGB" and f7_qtd_sacas == 7:
                                i7_fib = Decimal('4')
                            else:
                                v_i = float(g7_peso_real / f7_qtd_sacas) / 4.5
                                i7_fib = Decimal(str(math.ceil(v_i) if (v_i - int(v_i)) > 0.50 else math.floor(v_i)))

                            j7_kg_g = (g7_peso_real / f7_qtd_sacas / i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                            
                            # Ajuste fino de compensação para bater o peso total exato
                            while True:
                                k7_saca = (j7_kg_g * i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                                l7_total_simulado = k7_saca * f7_qtd_sacas
                                m7_saldo = l7_total_simulado - g7_peso_real
                                if m7_saldo >= 0: 
                                    break
                                else: 
                                    j7_kg_g += Decimal('0.01')

                            # Formatação dos textos padrão Brasil
                            txt_fibreboard = str(int(i7_fib))
                            txt_kg_g       = "{:.2f}".format(j7_kg_g).replace('.', ',')
                            txt_total_ovp  = "{:.2f}".format(k7_saca).replace('.', ',')
                            
                            # Gera a sequência de numeração das sacas (#1 #2 #3...)
                            marcacao = " ".join([f"#{i+1}" for i in range(int(f7_qtd_sacas))])

                            # Mapeia as chaves que estão no seu arquivo .docx do Word
                            contexto = {
                                'FIBREBOARD': txt_fibreboard,
                                'PESO_G': txt_kg_g,
                                'TOTAL_OVERPACK': txt_total_ovp,
                                'MARCACAO': marcacao,
                                'DATA': date.today().strftime('%d/%m/%Y'),
                                'QTD_OVERPACK': int(f7_qtd_sacas)
                            }

                            try:
                                # Carrega o template correspondente dentro da pasta templates/
                                caminho_template = f"templates/{sigla}-SHIPPER-t.docx"
                                doc = DocxTemplate(caminho_template)
                                doc.render(contexto)
                                
                                # Salva temporariamente na memória e empacota no ZIP
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                zip_file.writestr(f"Shipper_{sigla}.docx", doc_io.getvalue())
                                emitidos.append(sigla)
                                
                            except Exception as e_doc:
                                erros_cidades.append(f"{sigla} (Template não encontrado em templates/{sigla}-SHIPPER-t.docx)")
                        else:
                            erros_cidades.append(f"{sigla} (Destino não localizado na Coluna A da planilha)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso! Shippers calculadas e montadas para: {', '.join(emitidos)}")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name="Shippers_Perfeitas_Word.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("Nenhuma Shipper pôde ser gerada. Verifique os avisos acima.")
        except Exception as e:
            st.error(f"Erro no processamento do arquivo: {e}")
