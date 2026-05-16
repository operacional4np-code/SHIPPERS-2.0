import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
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
st.subheader("Insira as sacas, calcule os saldos e gere os arquivos direto nos modelos oficiais")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Informações (.xlsx / .xlsm)", type=["xlsm", "xlsx"])

def extrair_peso_seguro(linha_row):
    """Varre a linha da planilha procurando o valor numérico do Peso Real de forma protegida"""
    valores_numericos = []
    for val in linha_row:
        if pd.notnull(val) and not isinstance(val, str):
            try:
                v_float = float(val)
                if v_float > 0: valores_numericos.append(v_float)
            except: pass
        elif isinstance(val, str):
            # Limpa pontos e vírgulas de textos para tentar converter em número decimal
            txt_limpo = re.sub(r'[^\d.,]', '', val).strip()
            if "," in txt_limpo and "." in txt_limpo:
                txt_limpo = txt_limpo.replace(".", "").replace(",", ".")
            elif "," in txt_limpo:
                txt_limpo = txt_limpo.replace(",", ".")
            try:
                v_float = float(txt_limpo)
                if v_float > 0: valores_numericos.append(v_float)
            except: pass
    return valores_numericos[0] if valores_numericos else 0.0

# 2. SELETOR DINÂMICO DE SACAS NA TELA
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    sacas_manuais = {}
    
    colunas_tela = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        with colunas_tela[idx]:
            # Mantém a regra de negócio: POA vem com 17 por padrão, outras cidades vêm com 7
            default_val = 17 if sigla == "POA" else 7
            sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=default_val, step=1, key=f"sacas_{sigla}")

    if file:
        try:
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            if st.button("🔢 CALCULAR E GERAR DOCUMENTOS (WORD)"):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        # Busca segura contra erros de renderização em tela (KeyError)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        # Localizar a linha correta da cidade na planilha carregada
                        linha_dados = None
                        for index, row in df_raw.iterrows():
                            linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
                            if cidade_alvo in linha_texto and "TOTAL" not in linha_texto:
                                linha_dados = row
                                break

                        if linha_dados is not None:
                            peso_capturado = extrair_peso_seguro(linha_dados.values)
                            
                            if peso_capturado == 0.0:
                                erros_cidades.append(f"{sigla} (Peso real não localizado na linha correspondente)")
                                continue

                            g7_peso_real = Decimal(str(peso_capturado))
                            f7_qtd_sacas = Decimal(str(qtd_sacas_escolhida))
                            
                            # Execução das regras matemáticas de equilíbrio de pesos e volumes
                            if sigla == "CGB" and f7_qtd_sacas == 7:
                                i7_fib = Decimal('4')
                            else:
                                v_i = float(g7_peso_real / f7_qtd_sacas) / 4.5
                                i7_fib = Decimal(str(math.ceil(v_i) if (v_i - int(v_i)) > 0.50 else math.floor(v_i)))

                            j7_kg_g = (g7_peso_real / f7_qtd_sacas / i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                            
                            while True:
                                k7_saca = (j7_kg_g * i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                                l7_total_simulado = k7_saca * f7_qtd_sacas
                                m7_saldo = l7_total_simulado - g7_peso_real
                                if m7_saldo >= 0: break
                                else: j7_kg_g += Decimal('0.01')

                            # Montagem do dicionário com as tags que vão preencher o Word
                            contexto = {
                                'FIBREBOARD': int(i7_fib),
                                'PESO_G': "{:.2f}".format(j7_kg_g).replace('.', ','),
                                'TOTAL_OVERPACK': "{:.2f}".format(k7_saca).replace('.', ','),
                                'MARCACAO': " ".join([f"#{i+1}" for i in range(int(f7_qtd_sacas))]),
                                'DATA': date.today().strftime('%d/%m/%Y'),
                                'QTD_OVERPACK': int(f7_qtd_sacas)
                            }

                            try:
                                # Puxa o template correspondente dentro da pasta templates que você criou no GitHub
                                caminho_template = f"templates/{sigla}-SHIPPER-t.docx"
                                doc = DocxTemplate(caminho_template)
                                doc.render(contexto)
                                
                                # Salva o arquivo temporariamente em memória para compactar
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                zip_file.writestr(f"Shipper_{sigla}.docx", doc_io.getvalue())
                                emitidos.append(sigla)
                            except Exception as e_doc:
                                erros_cidades.append(f"{sigla} (Erro ao ler arquivo na pasta templates: {e_doc})")
                        else:
                            erros_cidades.append(f"{sigla} (Cidade não encontrada na planilha)")

                # Exibe os alertas de erros se houver alguma inconsistência
                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                # Se tudo correr bem, libera o botão de download do ZIP com os Words prontos
                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso! Shippers em Word geradas para: {', '.join(emitidos)}")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name=f"Shippers_Word_NewPost_{date.today()}.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("Nenhuma Shipper pôde ser processada. Verifique os erros acima.")
        except Exception as e:
            st.error(f"Erro geral no processamento da planilha: {e}")
