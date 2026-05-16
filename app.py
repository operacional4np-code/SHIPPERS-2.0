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
st.title("📄 Gerador de Shippers New Post — Padrão Word (.docx)")
st.subheader("Cálculo Autônomo Baseado nas Fórmulas do Excel (Apenas Planilha de Coleta)")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):").upper().strip()
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
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            if st.button("🔢 CALCULAR E GERAR SHIPPERS (AUTÔNOMO)"):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            
                            # --- FORMULAÇÃO EXCEL TRADUZIDA EM DECIMAIS PARA EVITAR ARREDONDAMENTO INCORRETO ---
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            # 1. Coluna G: =(F6*3)+D6 (Sacas * Peso Padrão de 3kg + Peso Original da Coleta)
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            # 2. Coluna I (Fibreboard): Qtd Volumes / Sacas -> Regra de corte truncado (Ex: 9,64 = 9)
                            fracao_fib = q_volumes / qtd_sacas_escolhida
                            i_fibreboard = math.floor(fracao_fib)
                            if i_fibreboard == 0: 
                                i_fibreboard = 1
                            i_fib_dec = Decimal(str(i_fibreboard))
                            
                            # 3. Coluna J (Kg G) e Simulação da Coluna M (Saldo Positivo mais próximo de zero)
                            # O Excel calcula a base: (Peso Corrigido / Sacas) / Fibreboard
                            base_j = (g_peso_corrigido / f_sacas) / i_fib_dec
                            j_inicio = base_j.quantize(Decimal('0.01'), rounding=ROUND_DOWN)
                            
                            perfeito_j = j_inicio
                            menor_saldo_positivo = Decimal('inf')
                            
                            # Varre os centavos para cima exatamente como você testaria no Excel até achar o equilíbrio positivo
                            for acrescimo in range(300): # Testa uma amplitude de até 3 reais para cima
                                j_teste = j_inicio + (Decimal(str(acrescimo)) * Decimal('0.01'))
                                
                                # Coluna K: TOTAL QUANTITY PER OVERPACK = J * I (Arredondado para 2 casas)
                                k_total_saca = (j_teste * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                                
                                # Coluna L: Peso Total do Destino = K * Sacas
                                l_total_destino = k_total_saca * f_sacas
                                
                                # Coluna M: Conferência = Total Destino - Peso Corrigido Real
                                m_conferencia = l_total_destino - g_peso_corrigido
                                
                                # Regra absoluta: número positivo (>= 0) mais próximo de zero
                                if m_conferencia >= 0:
                                    if m_conferencia < menor_saldo_positivo:
                                        menor_saldo_positivo = m_conferencia
                                        perfeito_j = j_teste
                                        break # Encontrou o primeiríssimo positivo válido (o mais próximo de zero)!
                            
                            # Atribui os valores finais validados pelo laço matemático
                            j7_kg_g = perfeito_j
                            k7_total_saca_final = (j7_kg_g * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                            # 4. Formatação amigável para envio ao Word
                            txt_fibreboard = str(int(i_fibreboard))
                            txt_kg_g       = "{:.2f}".format(j7_kg_g).replace('.', ',')
                            txt_total_ovp  = "{:.2f}".format(k7_total_saca_final).replace('.', ',')
                            
                            # String sequencial das marcações (#1 #2 #3...)
                            marcacao = " ".join([f"#{i+1}" for i in range(int(qtd_sacas_escolhida))])

                            contexto = {
                                'FIBREBOARD': txt_fibreboard,
                                'PESO_G': txt_kg_g,
                                'TOTAL_OVERPACK': txt_total_ovp,
                                'MARCACAO': marcacao,
                                'DATA': date.today().strftime('%d/%m/%Y'),
                                'QTD_OVERPACK': int(qtd_sacas_escolhida)
                            }

                            try:
                                caminho_template = f"templates/{sigla}-SHIPPER-t.docx"
                                doc = DocxTemplate(caminho_template)
                                doc.render(contexto)
                                
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                zip_file.writestr(f"Shipper_{sigla}.docx", doc_io.getvalue())
                                emitidos.append(sigla)
                                
                            except Exception as e_doc:
                                erros_cidades.append(f"{sigla} (Template não encontrado em templates/{sigla}-SHIPPER-t.docx)")
                        else:
                            erros_cidades.append(f"{sigla} (Não foi possível extrair dados válidos da planilha de coleta)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso! Shippers processadas autonomamente para: {', '.join(emitidos)}")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name="Shippers_Autonomas_NewPost.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("Nenhuma Shipper pôde ser gerada.")
        except Exception as e:
            st.error(f"Erro no processamento interno do arquivo: {e}")
