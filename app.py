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

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    sacas_manuais[sigla] = st.number_input(f"Sacas para {siglas_input}:", min_value=1, step=1, key=f"sacas_{siglas_input}")

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
                            
                            # 2. Coluna I: Fibreboard Boxes (Qtd Volumes / Sacas)
                            # Usando ROUND_HALF_UP puramente com Decimais: se a dízima for >= 0.50, vai para cima.
                            fracao_fib = Decimal(str(q_volumes)) / f_sacas
                            i_fibreboard = int(fracao_fib.quantize(Decimal('1'), rounding=ROUND_HALF_UP))
                            
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
                            txt_kg_g       = "{:.2f}".format(j7_kg_g).replace('.', ',')
                            txt_total_ovp  = "{:.2f}".format(k7_total_saca_final).replace('.', ',')
                            
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
                    st.success(f"✅ Sucesso! Shippers geradas com a regra oficial aplicada para: {', '.join(emitidos)}")
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
            st.error(f"Erro no processamento interno do arquivo: {e}")
