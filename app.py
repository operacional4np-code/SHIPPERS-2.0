import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP, ROUND_DOWN
from docxtpl import DocxTemplate
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES para busca na planilha carregada
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
st.subheader("Algoritmo de Aproximação de Saldo Positivo Próximo a Zero (Idêntico à Planilha)")

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA):").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Informações para Shippers (.xlsm / .xlsx)", type=["xlsm", "xlsx"])

def formatar_valor_br(valor):
    """Garante a formatação com duas casas decimais e vírgula separando os centavos"""
    try:
        if pd.isna(valor) or valor == "":
            return "0,00"
        return "{:.2f}".format(float(valor)).replace('.', ',')
    except:
        return str(valor).replace('.', ',')

def extrair_dados_linha_completa(df_raw, termo_busca):
    """Localiza a linha correspondente à cidade e extrai (Destino, Qtd, Peso Original)"""
    for index, row in df_raw.iterrows():
        # Transforma a linha inteira em texto para achar a cidade de forma flexível
        linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
        if termo_busca in linha_texto and "TOTAL" not in linha_texto:
            # Encontrou a linha! Agora extrai os dados base purificados das primeiras colunas
            valores = list(row.values)
            
            # Limpeza do destino (Coluna B original da planilha de coleta)
            destino_txt = str(valores[1]).upper() if len(valores) > 1 else termo_busca
            
            # Captura da Quantidade (Coluna C original da planilha de coleta)
            qtd_volumes = 1
            if len(valores) > 2 and pd.notnull(valores[2]):
                try:
                    qtd_volumes = int(float(str(valores[2]).replace(',', '.')))
                except: pass
                
            # Captura do Peso Original (Coluna D original da planilha de coleta)
            peso_original = 0.0
            if len(valores) > 3 and pd.notnull(valores[3]):
                try:
                    txt_p = re.sub(r'[^\d.,]', '', str(valores[3])).strip()
                    if "," in txt_p and "." in txt_p:
                        txt_p = txt_p.replace(".", "").replace(",", ".")
                    elif "," in txt_p:
                        txt_p = txt_p.replace(",", ".")
                    peso_original = float(txt_p)
                except: pass
                
            return destino_txt, qtd_volumes, peso_original
    return None, None, None

# 2. SELETOR DINÂMICO DE SACAS NA TELA DO SITE
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    sacas_manuais = {}
    
    colunas_tela = st.columns(len(lista_siglas))
    for idx, sigla in enumerate(lista_siglas):
        with colunas_tela[idx]:
            # Mantém a regra padrão operacional da New Post
            default_val = 17 if sigla == "POA" else 7
            sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=default_val, step=1, key=f"sacas_{sigla}")

    if file:
        try:
            # Carrega o arquivo sem assumir cabeçalhos fixos para evitar quebras de linhas superiores
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            if st.button("🔢 GERAR DOCUMENTOS WORD (CÁLCULO EXATO NEW POST)"):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        # Extrai as informações de origem da linha da planilha
                        destino_completo, q_volumes, p_original = extrair_dados_linha_completa(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            # --- EXECUÇÃO MATEMÁTICA CONFORME AS REGRAS DA PLANILHA ---
                            
                            # 1. Peso Corrigido (Coluna G): (Sacas * 3) + Peso Original
                            peso_corrigido_g = (qtd_sacas_escolhida * 3) + p_original
                            
                            # 2. Fibreboard (Coluna I): Qtd Volumes / Qtd Sacas (Arredondamento customizado baseado na fração)
                            fracao_fib = q_volumes / qtd_sacas_escolhida
                            resto_decimal = fracao_fib - int(fracao_fib)
                            if resto_decimal > 0.50:
                                i7_fib = math.ceil(fracao_fib)
                            else:
                                i7_fib = math.floor(fracao_fib)
                                
                            if i7_fib == 0:  # Garante que nunca divida por zero se a saca for maior que o volume
                                i7_fib = 1
                            
                            # Transforma em Decimal para rodar o laço de aproximação do saldo (Coluna M)
                            g_peso_ajustado_dec = Decimal(str(peso_corrigido_g))
                            f_sacas_dec = Decimal(str(qtd_sacas_escolhida))
                            i_fib_dec = Decimal(str(i7_fib))
                            
                            # 3. Determinação do Kg G por caixa (Coluna J) simulando a conferência do saldo positivo (Coluna M)
                            # Começa o teste com o valor matemático exato da divisão básica
                            j7_kg_g = (g_peso_ajustado_dec / f_sacas_dec / i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_DOWN)
                            
                            # Laço de repetição: vai subindo de 0.01 em 0.01 até achar o primeiro saldo Positivo (M >= 0)
                            while True:
                                k7_total_saca_simulada = (j7_kg_g * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                                l7_total_destino_simulado = k7_total_saca_simulada * f_sacas_dec
                                m7_saldo_conferencia = l7_total_destino_simulado - g_peso_ajustado_dec
                                
                                if m7_saldo_conferencia >= 0:
                                    break  # Paramos aqui pois encontramos o número positivo mais próximo de zero!
                                else:
                                    j7_kg_g += Decimal('0.01')

                            # 4. Formatação das variáveis do Word padrão Brasil
                            txt_fibreboard = str(int(i7_fib))
                            txt_kg_g       = "{:.2f}".format(j7_kg_g).replace('.', ',')
                            txt_total_ovp  = "{:.2f}".format(j7_kg_g * i_fib_dec).replace('.', ',')
                            
                            # Cria a string de marcações (#1 #2 #3...)
                            marcacao = " ".join([f"#{i+1}" for i in range(int(qtd_sacas_escolhida))])

                            # Prepara o contexto para colar nas chaves {{ }} do documento Word
                            contexto = {
                                'FIBREBOARD': txt_fibreboard,
                                'PESO_G': txt_kg_g,
                                'TOTAL_OVERPACK': txt_total_ovp,
                                'MARCACAO': marcacao,
                                'DATA': date.today().strftime('%d/%m/%Y'),
                                'QTD_OVERPACK': int(qtd_sacas_escolhida)
                            }

                            try:
                                # Procura o template original correspondente na pasta templates/
                                caminho_template = f"templates/{sigla}-SHIPPER-t.docx"
                                doc = DocxTemplate(caminho_template)
                                doc.render(contexto)
                                
                                # Exporta o arquivo para a memória e adiciona ao pacote final ZIP
                                doc_io = io.BytesIO()
                                doc.save(doc_io)
                                zip_file.writestr(f"Shipper_{sigla}.docx", doc_io.getvalue())
                                emitidos.append(sigla)
                                
                            except Exception as e_doc:
                                erros_cidades.append(f"{sigla} (Modelo não encontrado em templates/{sigla}-SHIPPER-t.docx)")
                        else:
                            erros_cidades.append(f"{sigla} (Não foi possível extrair dados de peso válidos para a cidade)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso total! Shippers calculadas com a regra oficial para: {', '.join(emitidos)}")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name="Shippers_Calculo_Oficial_NewPost.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("Nenhuma Shipper pôde ser gerada com o arquivo atual.")
        except Exception as e:
            st.error(f"Erro no processamento interno do arquivo: {e}")
