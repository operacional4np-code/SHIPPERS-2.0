import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal
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
st.subheader("Cálculo Autônomo Alinhado")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):", value="").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Base)", type=["xlsm", "xlsx"])

def extrair_dados_coleta(df_raw, termo_busca):
    """
    Varre a planilha dinamicamente procurando a linha que contém o DESTINO.
    Identifica de forma inteligente a coluna de Quantidade e Peso daquela linha.
    """
    for index, row in df_raw.iterrows():
        # Transforma os valores da linha em texto limpo para localizar a cidade
        valores_linha = [str(val).strip().upper() for val in row.values if pd.notnull(val)]
        linha_texto = " ".join(valores_linha)
        
        if termo_busca in linha_texto and "TOTAL" not in linha_texto:
            numeros_linha = []
            # Captura todos os números válidos daquela linha da esquerda para a direita
            for val in row.values:
                if pd.notnull(val):
                    if isinstance(val, (int, float)):
                        numeros_linha.append(float(val))
                    else:
                        txt = str(val).strip().replace(',', '.')
                        if re.match(r'^-?\d+([.,]\d+)?$', txt):
                            numeros_linha.append(float(txt))
            
            # O primeiro número é sempre a Qtd de Volumes e o segundo é o Peso
            if len(numeros_linha) >= 2:
                qtd_volumes = int(numeros_linha[0])
                peso_original = float(numeros_linha[1])
                return termo_busca, qtd_volumes, peso_original
                
    return None, None, None

# 2. SELETOR DE SACAS
sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    cols = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        with cols[idx]:
            sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=None, step=1, key=f"sacas_{sigla}")

    todas_sacas_preenchidas = all(saca is not None for saca in sacas_manuais.values())

    if file and todas_sacas_preenchidas:
        try:
            # Carrega a planilha bruta ignorando cabeçalhos automáticos para fazermos a busca manual
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            st.markdown("---")
            if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                # Tabela de conferência visual para o usuário no Streamlit
                dados_conferencia = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla)
                        
                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            # Coluna G: Peso Corrigido = (Sacas * 3) + Peso da Coleta
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            # Coluna I: Fibreboard Boxes = Volumes / Sacas (Arredondamento oficial)
                            fracao_fib = float(q_volumes) / float(qtd_sacas_escolhida)
                            decimal_part = fracao_fib - math.floor(fracao_fib)
                            if decimal_part >= 0.50:
                                i_fibreboard = math.floor(fracao_fib) + 1
                            else:
                                i_fibreboard = math.floor(fracao_fib)
                                
                            if i_fibreboard == 0: 
                                i_fibreboard = 1

                            i_fib_dec = Decimal(str(i_fibreboard))
                            
                            # Varredura centavo por centavo baseada na Coluna M da sua planilha:
                            # M = (J * I * Sacas) - G. Buscando o menor J onde M >= 0
                            base_j_float = float(g_peso_corrigido / f_sacas / i_fib_dec)
                            j_inicio_float = max(0.01, math.floor(base_j_float * 100) / 100 - 0.50)
                            j_inicio = Decimal(f"{j_inicio_float:.2f}")
                            
                            perfeito_j = None
                            for acrescimo in range(1000): 
                                j_teste = j_inicio + (Decimal(str(acrescimo)) * Decimal('0.01'))
                                l_total_destino = j_teste * i_fib_dec * f_sacas
                                m_conferencia = l_total_destino - g_peso_corrigido
                                
                                if m_conferencia >= Decimal('0'):
                                    perfeito_j = j_teste
                                    break
                            
                            if perfeito_j is None:
                                perfeito_j = Decimal(f"{base_j_float:.2f}")

                            j7_kg_g = perfeito_j
                            
                            # Coluna K: Total Quantity per Overpack (J * I)
                            k7_total_saca_final = j7_kg_g * i_fib_dec

                            # Guarda os dados para exibir o relatório de conferência na tela
                            dados_conferencia.append({
                                "Destino": sigla,
                                "Qtd Volumes Encontrada": q_volumes,
                                "Peso Encontrado (Kg)": float(p_original),
                                "Fibreboard Boxes (I)": int(i_fibreboard),
                                "Peso Unitário Kg G (J)": float(j7_kg_g),
                                "Total Overpack (K)": float(k7_total_saca_final)
                            })

                            # Formatação final de texto para preencher as tags do Word
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
                            erros_cidades.append(f"{sigla} (Não foi possível localizar dados válidos na linha deste destino)")

                # Exibe a tabela de conferência na tela para você validar antes de baixar
                if dados_conferencia:
                    st.markdown("### 📊 Relatório de Conferência dos Cálculos Gerados:")
                    st.dataframe(pd.DataFrame(dados_conferencia), use_container_width=True)

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso! Shippers geradas com sucesso.")
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
