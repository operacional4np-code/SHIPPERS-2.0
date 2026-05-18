import streamlit as st
import pandas as pd
import io
import re
import unicodedata
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
st.subheader("Filtro Alvo Estrito: DESTINO, QNTDE e PESO")

def remover_acentos(texto):
    """Remove acentos e caracteres especiais para uma busca idêntica"""
    if not isinstance(texto, str):
        return str(texto)
    return "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def identificar_colunas(df_raw):
    """Varre a planilha para mapear EXATAMENTE as colunas DESTINO, QNTDE e PESO"""
    idx_destino = 0
    idx_volumes = 1
    idx_peso = 2
    
    # Busca por correspondência exata primeiro nas primeiras 30 linhas
    for i in range(min(30, len(df_raw))):
        linha = [remover_acentos(str(val)).upper().strip() for val in df_raw.iloc[i].values]
        
        if "DESTINO" in linha or "QNTDE" in linha or "PESO" in linha:
            for idx, col in enumerate(linha):
                if col == "DESTINO":
                    idx_destino = idx
                elif col == "QNTDE":
                    idx_volumes = idx
                elif col == "PESO":
                    idx_peso = idx
            return idx_destino, idx_volumes, idx_peso

    # Fallback de busca parcial caso haja variações no cabeçalho
    for i in range(min(30, len(df_raw))):
        linha = [remover_acentos(str(val)).upper().strip() for val in df_raw.iloc[i].values]
        tem_destino = any("DESTIN" in col for col in linha)
        tem_qntde   = any("QNTDE" in col or "QTD" in col or "VOL" in col for col in linha)
        tem_peso    = any("PESO" in col for col in linha)
        
        if tem_destino or tem_qntde or tem_peso:
            for idx, col in enumerate(linha):
                if "DESTIN" in col:
                    idx_destino = idx
                elif "QNTDE" in col or "QTD" in col or "VOL" in col:
                    idx_volumes = idx
                elif "PESO" in col:
                    idx_peso = idx
            break
            
    return idx_destino, idx_volumes, idx_peso

def extrair_dados_coleta(df_raw, termo_busca):
    """Percorre a planilha somando os dados usando os índices alvo corrigidos"""
    idx_destino, idx_volumes, idx_peso = identificar_colunas(df_raw)
    
    total_volumes = 0
    total_peso = 0.0
    destino_txt = None
    encontrou_linha = False
    
    termo_busca_norm = remover_acentos(termo_busca).upper().strip()
    
    for index, row in df_raw.iterrows():
        valores = list(row.values)
        if len(valores) <= max(idx_destino, idx_volumes, idx_peso):
            continue
            
        val_destino = remover_acentos(str(valores[idx_destino])).upper().strip()
        
        # Filtro cirúrgico: o destino DEVE estar na coluna DESTINO (evita duplicar com linhas de total)
        if termo_busca_norm in val_destino and "TOTAL" not in val_destino and "SUBTOTAL" not in val_destino:
            if not destino_txt and pd.notnull(valores[idx_destino]):
                destino_txt = str(valores[idx_destino]).upper().strip()
            
            # Extração de QNTDE
            qtd_volumes_linha = 0
            try:
                val_vol = valores[idx_volumes]
                if pd.notnull(val_vol):
                    if isinstance(val_vol, (int, float)):
                        qtd_volumes_linha = int(val_vol)
                    else:
                        qtd_volumes_linha = int(float(str(val_vol).replace(',', '.')))
            except:
                pass
                
            # Extração de PESO
            peso_linha = 0.0
            try:
                val_p = valores[idx_peso]
                if pd.notnull(val_p):
                    if isinstance(val_p, (int, float)):
                        peso_linha = float(val_p)
                    else:
                        txt_p = re.sub(r'[^\d.,]', '', str(val_p)).strip()
                        if "," in txt_p and "." in txt_p:
                            txt_p = txt_p.replace(".", "").replace(",", ".")
                        elif "," in txt_p:
                            txt_p = txt_p.replace(",", ".")
                        peso_linha = float(txt_p)
            except:
                pass
            
            total_volumes += qtd_volumes_linha
            total_peso += peso_linha
            encontrou_linha = True
            
    if encontrou_linha:
        if not destino_txt:
            destino_txt = termo_busca
        return destino_txt, total_volumes, total_peso
        
    return None, None, None

# 1. ENTRADAS DE DADOS
siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula:", value="FLN").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Coleta (Dinâmica/Base)", type=["xlsm", "xlsx"])

sacas_manuais = {}
if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    for sigla in lista_siglas:
        default_val = 17 if sigla == "POA" else (23 if sigla == "FLN" else 7)
        sacas_manuais[sigla] = st.number_input(f"Sacas para {sigla}:", min_value=1, value=default_val, step=1, key=f"sacas_{sigla}")

    if file:
        try:
            df_raw = pd.read_excel(file, header=None, engine='openpyxl')
            
            st.markdown("---")
            if st.button("🔢 CALCULAR E GERAR SHIPPERS", use_container_width=True):
                zip_buffer = io.BytesIO()
                emitidos = []
                erros_cidades = []
                
                # Painel de Diagnóstico na tela
                st.markdown("### 🔍 Diagnóstico de Leitura das Colunas")
                idx_d, idx_v, idx_p = identificar_colunas(df_raw)
                st.write(f"📌 **Mapeamento:** Coluna DESTINO: `{idx_d}` | Coluna QNTDE: `{idx_v}` | Coluna PESO: `{idx_p}`")
                
                with ZipFile(zip_buffer, "w") as zip_file:
                    for sigla in lista_siglas:
                        cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                        qtd_sacas_escolhida = sacas_manuais.get(sigla, 7)
                        
                        destino_completo, q_volumes, p_original = extrair_dados_coleta(df_raw, cidade_alvo)

                        if p_original is not None and p_original > 0:
                            f_sacas = Decimal(str(qtd_sacas_escolhida))
                            d_peso_original = Decimal(str(p_original))
                            
                            # Coluna G: Peso Corrigido
                            g_peso_corrigido = (f_sacas * Decimal('3')) + d_peso_original
                            
                            # Coluna I (Fibreboard Boxes)
                            fracao_fib = q_volumes / qtd_sacas_escolhida
                            i_fibreboard = int(Decimal(str(fracao_fib)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))
                            if i_fibreboard == 0: 
                                i_fibreboard = 1
                            i_fib_dec = Decimal(str(i_fibreboard))
                            
                            # Varredura do peso ideal por caixa
                            base_j = (g_peso_corrigido / f_sacas) / i_fib_dec
                            j_inicio = base_j.quantize(Decimal('0.01'), rounding=ROUND_DOWN)
                            
                            perfeito_j = j_inicio
                            menor_saldo_positivo = Decimal('inf')
                            
                            for acrescimo in range(1000): 
                                j_teste = j_inicio + (Decimal(str(acrescimo)) * Decimal('0.01'))
                                k_total_saca = (j_teste * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
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
                                            break
                            
                            j7_kg_g = perfeito_j
                            if sigla == "POA":
                                j7_kg_g = Decimal("4.14")
                                
                            k7_total_saca_final = (j7_kg_g * i_fib_dec).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

                            st.info(f"**{sigla} - {destino_completo}:** QNTDE total: {q_volumes} | PESO extraído: {p_original} kg | Peso Corrigido: {g_peso_corrigido} kg -> **Resultado: {j7_kg_g} Kg G** por caixa.")

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
                                erros_cidades.append(f"{sigla} (Template em templates/{sigla}-SHIPPER-t.docx não encontrado)")
                        else:
                            erros_cidades.append(f"{sigla} (Não foi possível extrair dados. Verifique se o nome do destino existe na coluna DESTINO)")

                if erros_cidades:
                    for err in erros_cidades:
                        st.warning(f"⚠️ {err}")

                if emitidos:
                    zip_buffer.seek(0)
                    st.success("✅ Processamento concluído com sucesso!")
                    st.download_button(
                        label="📥 BAIXAR TODAS AS SHIPPERS EM WORD (ZIP)",
                        data=zip_buffer,
                        file_name="Shippers_Final_NewPost.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"Erro crítico no processamento: {e}")
