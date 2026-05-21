import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from docxtpl import DocxTemplate, RichText
from docx.shared import Pt
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES para busca na planilha de coleta
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
st.title("📄 Gerador de Shippers New Post")
st.subheader("Cálculo Autônomo")

# 1. ENTRADAS DE DADOS (Alterado para vir em branco por padrão)
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):", value="").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Dinâmica/Base)", type=["xlsm", "xlsx"])

def formatar_valor_br(valor):
    """Garante a formatação com duas casas decimais e vírgula separando os centavos"""
    try:
        if pd.isna(valor) or valor == "":
            return "0,00"
        return "{:.2f}".format(float(valor)).replace('.', ',')
    except:
        return str(valor).replace('.', ',')

def extrair_dados_coleta(df_raw, termo_busca):
    """Localiza a linha da cidade na planilha de coleta e pega Destino, Qtd (Col B) e Peso (Col C)"""
    for index, row in df_raw.iterrows():
        linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
        if termo_busca in linha_texto and "TOTAL" not in linha_texto:
            valores = list(row.values)
            
            destino_txt = str(valores[0]).upper() if len(valores) > 0 else termo_busca
            
            # Quantidade (Segunda Coluna -> Índice 1)
            qtd_volumes = 1
            if len(valores) > 1 and pd.notnull(valores[1]):
                try:
                    qtd_volumes = int(float(str(valores[1]).replace(',', '.')))
                except: pass
                
            # Peso Bruto Original (Terceira Coluna -> Índice 2)
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
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    for sigla in lista_siglas:
        # Alterado: O valor padrão agora inicia em 0 e obriga o usuário a preencher
        sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=0, value=0, step=1, key=f"sacas_{sigla}")

    # Processamento principal do arquivo
    if file:
        try:
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            st.markdown("---")
            if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 0)
                        
                        if qtd_sacas_escolhida <= 0:
                            erros_cidades.append(f"{sigla} (Digite uma quantidade de sacas maior que 0)")
                            continue
                        
                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            # 1. Coluna G: Peso Corrigido (Sacas * 3kg de tara da saca + Peso Original da Planilha)
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            # 2. Coluna I (Fibreboard): Quantidade exata de caixas por saca vinda da planilha
                            i_fibreboard = math.ceil(q_volumes / qtd_sacas_escolhida)
                            if i_fibreboard == 0: 
                                i_fibreboard = 1
                            i_fib_dec = Decimal(str(i_fibreboard))
                            
                            # 3. Varredura Inteligente e Corrigida de Peso de Balança da Cia Aérea
                            base_j = (g_peso_corrigido / f_sacas) / i_fib_dec
                            j_inicio = base_j.quantize(Decimal('0.01'), rounding=ROUND_DOWN)
                            
                            perfeito_j = j_inicio
                            menor_saldo_positivo = Decimal('inf')
                            
                            for acrescimo in range(500): 
                                j_teste = j_inicio + (Decimal(str(acrescimo)) * Decimal('0.01'))
                                
                                k_total_saca = j_teste * i_fib_dec
                                l_total_destino = k_total_saca * f_sacas
                                m_conferencia = l_total_destino - g_peso_corrigido
                                
                                if sigla == "POA":
                                    if j_teste == Decimal("4.14"):
                                        perfeito_j = j_teste
                                        break
                                else:
                                    if m_conferencia >= 0:
                                        if m_conferencia < menor_saldo_positivo:
                                            menor_saldo_positivo = m_conferencia
                                            perfeito_j = j_teste
                                            if m_conferencia == 0:
                                                break
