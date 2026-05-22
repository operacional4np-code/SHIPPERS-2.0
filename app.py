import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
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
st.subheader("Cálculo Autônomo Alinhado com Planilha Modelo")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):", value="").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta", type=["xlsm", "xlsx"])

def formatar_valor_br(valor):
    """Garante a formatação com duas casas decimais e vírgula separando os centavos"""
    try:
        if pd.isna(valor) or valor == "":
            return "0,00"
        return "{:.2f}".format(float(valor)).replace('.', ',')
    except:
        return str(valor).replace('.', ',')

def extrair_dados_coleta(df_raw, termo_busca):
    """
    Busca o destino na Coluna A (Índice 0), 
    Quantidade na Coluna B (Índice 1) e Peso na Coluna C (Índice 2).
    """
    for index, row in df_raw.iterrows():
        # Verifica se a linha tem pelo menos as 3 colunas necessárias
        if len(row) >= 3:
            val_a = str(row.iloc[0]).upper().strip() if pd.notnull(row.iloc[0]) else ""
            
            # Procura o termo de busca na coluna A (ex: "CUIABA" ou "CGB")
            if termo_busca in val_a and "TOTAL" not in val_a:
                try:
                    # Lê quantidade (Coluna B)
                    val_b = row.iloc[1]
                    if isinstance(val_b, (int, float)):
                        qtd_volumes = int(val_b)
                    else:
                        qtd_volumes = int(float(str(val_b).replace(',', '.').strip()))
                    
                    # Lê peso (Coluna C)
                    val_c = row.iloc[2]
                    if isinstance(val_c, (int, float)):
                        peso_original = float(val_c)
                    else:
                        peso_original = float(str(val_c).replace(',', '.').strip())
                        
                    return termo_busca, qtd_volumes, peso_original
                except Exception as e:
                    # Se houver erro de conversão nessa linha específica, continua procurando
                    continue
                
    return None, None, None

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    for sigla in lista_siglas:
        sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=None, step=1, key=f"sacas_{sigla}")

    todas_sacas_preenchidas = all(saca is not None for saca in sacas_manuais.values())

    if file and todas_sacas_preenchidas:
        try:
            # Carrega a planilha sem pular cabeçalhos para mapear as colunas brutas por índice
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            st.markdown("---")
            if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla)
                        
                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            # Coluna G: Peso Corrigido = (Sacas * 3kg) + Peso Original
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            # Coluna I: Fibreboard Boxes = Qtd Volumes / Sacas (Com arredondamento padrão >= 0.5 sobe)
                            fracao_fib = float(q_volumes) / float(qtd_sacas_escolhida)
                            decimal_part = fracao_fib - math.floor(fracao_fib)
                            if decimal_part >= 0.50:
                                i_fibreboard = math.floor(fracao_fib) + 1
                            else:
                                i_fibreboard = math.floor(fracao_fib)
                                
                            if i_fibreboard == 0: 
                                i_fibreboard = 1

                            i_fib_dec = Decimal(str(i_fibreboard))
                            
                            # Varredura idêntica à lógica de menor resíduo positivo da coluna M:
                            # Formula da tabela: M = (J * I * Sacas) - G
                            # Queremos o menor J com duas casas decimais onde M >= 0
                            base_j_float = float(g_peso_corrigido / f_sacas / i_fib_dec)
                            j_inicio_float = max(0.01, math.floor(base_j_float * 100) / 100 - 0.50)
                            j_inicio = Decimal(f"{j_inicio_float:.2f}")
                            
                            perfeito_j = None
                            
                            for acrescimo in range(1000): 
                                j_teste = j_inicio + (Decimal(str(acrescimo)) * Decimal('0.01'))
                                
                                # Simulando a fórmula das colunas K e L do Excel:
                                # K (Peso saca) = J * I
                                # L (Peso Total) = K * Sacas
                                k_peso_saca = j_teste * i_fib_dec
                                l_total_destino = k_peso_saca * f_sacas
                                
                                # M (Conferência) = L - G
                                m_conferencia = l_total_destino - g_peso_corrigido
                                
                                # Procura o primeiro valor onde a conferência não seja negativa
                                if m_conferencia >= Decimal('0'):
                                    perfeito_j = j_teste
                                    break
                            
                            if perfeito_j is None:
                                perfeito_j = Decimal(f"{base_j_float:.2f}").quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                            j7_kg_g = perfeito_j
                            
                            # Coluna K: Total Quantity per Overpack (J * I)
                            k7_total_saca_final = j7_kg_g * i_fib_dec

                            # Formatação das variáveis para o Word
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
                            erros_cidades.append(f"{sigla} (Não foi possível extrair dados válidos para {sigla} nas colunas A, B e C)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso! Shippers geradas com os cálculos idênticos ao modelo oficial para: {', '.join(emitidos)}")
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
