import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from docxtpl import DocxTemplate
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

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):", value="CWB").upper().strip()
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

def arredondamento_customizado_coluna_i(valor_float):
    """
    Regra Exata: Se a casa decimal for de 50 para cima (>= 0.50), 
    arredonda para o número acima. Se for menor, para baixo.
    """
    dec, int_part = math.modf(valor_float)
    dec = round(dec, 4) # Evita dízimas de ponto flutuante do Python
    if dec >= 0.50:
        return int(int_part + 1)
    else:
        return int(int_part)

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    for sigla in lista_siglas:
        default_val = 17 if sigla == "POA" else 7
        sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=default_val, step=1, key=f"sacas_{sigla}")

    # O botão fica visível se o arquivo for carregado
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
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            # 1. Coluna G: Peso Corrigido (Sacas * 3kg + Peso Original da Coleta)
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            # 2. Coluna I: Fibreboard Boxes (Qtd Volumes / Sacas) com corte estrito em 0.50
                            fracao_fib = q_volumes / qtd_sacas_escolhida
                            i_fibreboard = arredondamento_customizado_coluna_i(fracao_fib)
                            if i_fibreboard == 0: 
                                i_fibreboard = 1
                            i_fib_dec = Decimal(str(i_fibreboard))
                            
                            # 3. Varredura do Peso Unitário Ideal (Coluna J)
                            base_j_float = float(g_peso_corrigido / f_sacas / i_fib_dec)
                            
                            # Começa a busca um pouco abaixo do valor teórico para pegar o ponto de virada exato
                            j_inicio_float = math.floor(base_j_float * 100) / 100 - 0.50
                            if j_inicio_float < 0:
                                j_inicio_float = 0.01
                                
                            j_inicio = Decimal(f"{j_inicio_float:.2f}")
                            perfeito_j = None
                            menor_saldo_positivo = Decimal('inf')
                            
                            # Testando centavo por centavo para achar o menor resíduo positivo na conferência
                            for acrescimo in range(2000): 
                                j_teste = j_inicio + (Decimal(str(acrescimo)) * Decimal('0.01'))
                                
                                # M = (Sacas * J * I) - G
                                l_total_destino = j_teste * i_fib_dec * f_sacas
                                m_conferencia = l_total_destino - g_peso_corrigido
                                
                                # Critério: m_conferencia deve ser >= 0 e o menor possível
                                if m_conferencia >= 0:
                                    if m_conferencia < menor_saldo_positivo:
                                        menor_saldo_positivo = m_conferencia
                                        perfeito_j = j_teste
                            
                            if perfeito_j == None:
                                perfeito_j = Decimal(f"{base_j_float:.2f}")

                            j7_kg_g = perfeito_j
                            k7_total_saca_final = j7_kg_g * i_fib_dec

                            # 4. Formatação das variáveis do Word
                            txt_fibreboard = str(int(i_fibreboard))
                            txt_kg_g       = "{:.2f}".format(j7_kg_g).
