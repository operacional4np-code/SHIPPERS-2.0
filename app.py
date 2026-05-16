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
st.subheader("Algoritmo de Verificação de Saldo Positivo (Idêntico ao Ajuste Manual do Excel)")

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
        linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
        if termo_busca in linha_texto and "TOTAL" not in linha_texto:
            valores = list(row.values)
            
            # Mapeamento dinâmico conforme colunas B, C e D da planilha de coleta
            destino_txt = str(valores[1]).upper() if len(valores) > 1 else termo_busca
            
            qtd_volumes = 1
            if len(valores) > 2 and pd.notnull(valores[2]):
                try:
                    qtd_volumes = int(float(str(valores[2]).replace(',', '.')))
                except: pass
                
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
            
            if st.button("🔢 GERAR DOCUMENTOS WORD (CALIBRADO COM REFERÊNCIA)"):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []

                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        destino_completo, q_volumes, p_original = extrair_dados_linha_completa(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            # 1. Peso Corrigido (Coluna G): (Sacas * 3) + Peso Original
                            peso_corrigido_g = (qtd_sacas_escolhida * 3) + p_original
                            
                            # 2. Fibreboard (Coluna I): Truncamento/Corte matemático conforme exemplo (9,64 = 9)
                            fracao_fib = q_volumes / qtd_sacas_escolhida
                            i7_fib = math.floor(fracao_fib)
                            if i7_fib == 0: 
                                i7_fib = 1
                            
                            # Transforma para Decimal para precisão absoluta nas regras de saldo
                            g_peso_ajustado_dec = Decimal(str(peso_corrigido_g))
                            f_sacas_dec = Decimal(str(qtd_sacas_escolhida))
                            i_fib_dec = Decimal(str(i7_fib))
                            
                            # 3. Varredura Inteligente para Encontrar o Kg G Perfeito (Coluna J)
                            # Começa o teste partindo do valor básico e busca o ajuste ótimo da planilha
                            base_calculo = (g_peso_ajustado_dec / f_sacas_dec / i_fib_dec)
                            j7_kg_g = base_calculo.quantize(Decimal('0.01'), rounding=ROUND_DOWN)
                            
                            melhor_j = j7_kg_g
                            menor_saldo_positivo = Decimal('inf')
                            
                            # Varre uma faixa de opções de centavos para cima para encontrar o menor positivo igual ao seu ajuste manual
                            for teste_cents in range(100):
                                j_teste = j7_kg_g + (Decimal(str(teste_cents)) * Decimal('0.01'))
                                k7_total_saca_teste = (j_teste * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                                l7_total_destino_teste = k7_total_saca_teste * f_sacas_dec
                                m7_saldo_teste = l7_total_destino_teste - g_peso_ajustado_dec
                                
                                # Regra de ouro da New Post: Tem que ser positivo (>= 0) e o mais próximo possível de zero
                                if m7_saldo_teste >= 0:
                                    if m7_saldo_teste < menor_saldo_positivo:
                                        menor_saldo_positivo = m7_saldo_teste
                                        melhor_j = j_teste
                                        
                            # Aplica o melhor valor calibrado encontrado
                            j7_kg_g = melhor_j
                            k7_total_saca_final = (j7_kg_g * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                            # 4. Formatação das variáveis do Word padrão Brasil
                            txt_fibreboard = str(int(i7_fib))
                            txt_kg_g       = "{:.2f}".format(j7_kg_g).replace('.', ',')
                            txt_total_ovp  = "{:.2f}".format(k7_total_saca_final).replace('.', ',')
                            
                            # Cria a string de marcações (#1 #2 #3...)
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
                                erros_cidades.append(f"{sigla} (Modelo não encontrado em templates/{sigla}-SHIPPER-t.docx)")
                        else:
                            erros_cidades.append(f"{sigla} (Não foi possível extrair dados de peso válidos)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success(f"✅ Sucesso total! Shippers sincronizadas perfeitamente para: {', '.join(emitidos)}")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name="Shippers_Calibradas_NewPost.zip",
                        mime="application/zip"
                    )
                else:
                    st.error("Nenhuma Shipper pôde ser gerada com o arquivo atual.")
        except Exception as e:
            st.error(f"Erro no processamento interno do arquivo: {e}")
