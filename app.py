import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from docxtpl import DocxTemplate, RichText
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
    "PVH": "PORTO VELHO"
}

st.set_page_config(
    page_title="New Post - Gerador Word Shippers", 
    layout="wide"
)
st.title("📄 Gerador de Shippers New Post")
st.subheader("Cálculo Autônomo")

# 1. ENTRADAS DE DADOS (Campos iniciam vazios)
siglas_input = st.text_input(
    "1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):", 
    value=""
).upper().strip()

file = st.file_uploader(
    "2. Carregue a Planilha de Coleta (Dinâmica/Base)", 
    type=["xlsm", "xlsx"]
)

def formatar_valor_br(valor):
    try:
        if pd.isna(valor) or valor == "":
            return "0,00"
        return "{:.2f}".format(float(valor)).replace('.', ',')
    except:
        return str(valor).replace('.', ',')

def extrair_dados_coleta(df_raw, termo_busca):
    for index, row in df_raw.iterrows():
        linha_texto = " ".join(
            [str(val).upper() for val in row.values if pd.notnull(val)]
        )
        if termo_busca in linha_texto and "TOTAL" not in linha_texto:
            valores = list(row.values)
            destino_txt = (
                str(valores[0]).upper() 
                if len(valores) > 0 
                else termo_busca
            )
            
            qtd_volumes = 1
            if len(valores) > 1 and pd.notnull(valores[1]):
                try:
                    qtd_volumes = int(
                        float(str(valores[1]).replace(',', '.'))
                    )
                except: pass
                
            peso_original = 0.0
            if len(valores) > 2 and pd.notnull(valores[2]):
                try:
                    txt_p = re.sub(r'[^\d.,]', '', str(valores[2])).strip()
                    if "," in txt_p and "." in txt_p:
                        txt_p = txt_p.replace(".", "").replace(",", ".")
                    elif "," in txt_p:
                        txt_p = txt_p.replace(",", ".")
                    peso_original = float(txt_p)
                except: pass
                
            return destino_txt, qtd_volumes, peso_original
    return None, None, None

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [
        s.strip() for s in siglas_input.split(",") if s.strip()
    ]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    for sigla in lista_siglas:
        sacas_manuais[sigla] = st.number_input(
            f"Sacas para {sigla}:", 
            min_value=1, 
            value=None, 
            step=1, 
            key=f"sacas_{sigla}"
        )

    if file:
        try:
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            st.markdown("---")
            
            if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
                valores_vazios = [
                    s for s, v in sacas_manuais.items() if v is None
                ]
                
                if valores_vazios:
                    st.error(
                        f"⚠️ Por favor, insira as sacas para: {', '.join(valores_vazios)}"
                    )
                else:
                    zip_buffer = io.BytesIO()
                    emitidos = []
                    erros_cidades = []

                    with ZipFile(zip_buffer, "w") as zip_file:
                        for sigla in lista_siglas:
                            cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                            qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                            
                            destino_completo, q_volumes, p_original = \
                                extrair_dados_coleta(df_raw, cidade_alvo)

                            if p_original is not None and p_original > 0:
                                f_sacas = Decimal(str(qtd_sacas_escolhida))
                                d_peso_original = Decimal(str(p_original))
                                
                                # Coluna G: Peso Corrigido
                                g_peso_corrigido = (
                                    (f_sacas * Decimal('3')) + d_peso_original
                                )
                                
                                # Coluna I: Fibreboard Boxes
                                fracao_fib = Decimal(str(q_volumes)) / f_sacas
                                i_fibreboard = int(
                                    fracao_fib.quantize(
                                        Decimal('1'), 
                                        rounding=ROUND_HALF_UP
                                    )
                                )
                                if i_fibreboard == 0: 
                                    i_fibreboard = 1
                                i_fib_dec = Decimal(str(i_fibreboard))
                                
                                # Coluna J: Peso Unitário Ideal
                                base_j_float = float(
                                    g_peso_corrigido / f_sacas / i_fib_dec
                                )
                                j_inicio_float = (
                                    math.floor(base_j_float * 100) / 100 - 0.50
                                )
                                if j_inicio_float < 0:
                                    j_inicio_float = 0.01
                                    
                                j_inicio = Decimal(f"{j_inicio_float:.2f}")
                                perfeito_j = None
                                menor_saldo = Decimal('inf')
                                
                                for acrescimo in range(2000): 
                                    j_teste = j_inicio + (
                                        Decimal(str(acrescimo)) * Decimal('0.01')
                                    )
                                    l_total_destino = (
                                        j_teste * i_fib_dec * f_sacas
                                    )
                                    m_conferencia = (
                                        l_total_destino - g_peso_corrigido
                                    )
                                    if m_conferencia >= 0:
                                        if m_conferencia < menor_saldo:
                                            menor_saldo = m_conferencia
                                            perfeito_j = j_teste
                                
                                if perfeito_j is None:
                                    perfeito_j = Decimal(f"{base_j_float:.2f}")

                                j7_kg_g = perfeito_j
                                k7_total_saca_final = j7_kg_g * i_fib_dec

                                # Formatação de Texto Segura
                                txt_fibreboard = str(int(i_fibreboard))
                                txt_kg_g = "{:.2f}".format(j7_kg_g).replace('.', ',')
                                txt_total_ovp = "{:.2f}".format(k7_total_saca_final).replace('.', ',')
                                
                                marcacao_crua = " ".join(
                                    [f"#{i+1}" for i in range(int(qtd_sacas_escolhida))]
                                )
                                
                                # RichText isolado e quebrado para evitar cortes
                                marcacao_rt = RichText()
                                marcacao_rt.add(
                                    marcacao_crua, 
                                    font='Arial Black', 
                                    size=16
                                )

                                contexto = {
                                    'FIBREBOARD': txt_fibreboard,
                                    'PESO_G': txt_kg_g,
                                    'TOTAL_OVERPACK': txt_total_ovp,
                                    'MARCACAO': marcacao_rt,
                                    'DATA': date.today().strftime('%d/%m/%Y'),
                                    'QTD_OVERPACK': int(qtd_sacas_escolhida)
                                }

                                try:
                                    caminho_template = (
                                        f"templates/{sigla}-SHIPPER-t.docx"
                                    )
                                    doc = DocxTemplate(caminho_template)
                                    doc.render(contexto)
                                    
                                    doc_io = io.BytesIO()
                                    doc.save(doc_io)
                                    zip_file.writestr(
                                        f"Shipper_{sigla}.docx", 
                                        doc_io.getvalue()
                                    )
                                    emitidos.append(sigla)
                                    
                                except Exception as e_doc:
                                    erros_cidades.append(
                                        f"{sigla} (Template indisponível)"
                                    )
                            else:
                                erros_cidades.append(
                                    f"{sigla} (Dados de coleta inválidos)"
                                )

                    if erros_cidades:
                        for err in erros_cidades:
                            st.warning(f"⚠️ {err}")

                    if emitidos:
                        zip_buffer.seek(0)
                        st.success("✅ Sucesso! Shippers geradas com sucesso.")
                        st.download_button(
                            label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                            data=zip_buffer,
                            file_name="Shippers_Final_NewPost.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                    else:
                        st.error("Nenhuma Shipper pôde ser gerada.")
        except Exception as e:
            st.error(f"Erro no processamento: {e}")
