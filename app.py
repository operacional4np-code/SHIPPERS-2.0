import streamlit as st
import pandas as pd
import io
import math
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from zipfile import ZipFile
from fpdf import FPDF

# 1. CLASSE DO LAYOUT DO PDF
class ShipperPDF(FPDF):
    def layout_shipper(self, ctx):
        self.add_page()
        self.set_font("Arial", "B", 13)
        
        # Título Principal
        self.cell(0, 10, "DECLARACAO DO EXPEDIDOR PARA ARTIGOS PERIGOSOS", border=1, ln=1, align="C", fill=False)
        self.set_fill_color(240, 240, 240)
        
        # Bloco 1: Expedidor e AWB
        self.set_font("Arial", "", 9)
        x = self.get_x()
        y = self.get_y()
        
        self.multi_cell(95, 6, "Expedidor:\nNEW POST SOLUCOES EM LOGISTICAS LTDA\nCNPJ: 28.678.104/0001-79\nR UBALDO FAGGEANI, 355 - JARDIM LAS PALMAS\nPORTO FERREIRA/SP - CEP: 13667-262", border=1)
        
        self.set_xy(x + 95, y)
        self.multi_cell(95, 10, "Nº do Conhecimento Aereo:\n\nPagina 1 de 1 Paginas", border=1)
        
        # Bloco 2: Consignatário
        y_atual = self.get_y()
        self.set_xy(x, y_atual)
        self.multi_cell(95, 6, f"Consignatario:\nBRASIL POSTAL LTDA - ME\nAV FERNANDO CORREA DA COSTA, 3180\nSHANGRILA - {ctx['CIDADE']}\nCEP: 78070-200", border=1)
        
        self.set_xy(x + 95, y_atual)
        self.multi_cell(95, 15, "Nº de Referencia do Expedidor:\n(Opcional)", border=1)
        
        # Detalhes de Transporte
        self.set_font("Arial", "B", 9)
        self.cell(0, 6, "DETALHES DE TRANSPORTE", border=1, ln=1, fill=True)
        
        self.set_font("Arial", "", 9)
        y_transp = self.get_y()
        self.multi_cell(95, 6, "Este embarque esta dentro das limitacoes prescritas para:\nAERONAVE DE PASSAGEIROS E CARGA", border=1)
        self.set_xy(x + 95, y_transp)
        self.multi_cell(95, 6, f"Aeroporto de Origem: CONFINS\nAeroporto de Destino: {ctx['CIDADE']}", border=1)
        
        # Natureza e Quantidade
        self.set_font("Arial", "B", 9)
        self.cell(0, 6, "NATUREZA E QUANTIDADE DE ARTIGOS PERIGOSOS", border=1, ln=1, fill=True)
        
        # Tabela de Conteúdo
        self.set_font("Arial", "B", 8)
        self.cell(15, 6, "Nº UN", border=1, align="C")
        self.cell(50, 6, "Nome Apropriado", border=1, align="C")
        self.cell(15, 6, "Classe", border=1, align="C")
        self.cell(15, 6, "Grupo", border=1, align="C")
        self.cell(75, 6, "Quantidade e Tipo de Embalagem", border=1, align="C")
        self.cell(20, 6, "Inst. Emb.", border=1, ln=1, align="C")
        
        # Dados da Tabela
        self.set_font("Arial", "", 8.5)
        y_dados = self.get_y()
        
        self.cell(15, 30, "ID8000", border=1, align="C")
        self.cell(50, 30, "CONSUMER COMMODITY", border=1, align="C")
        self.cell(15, 30, "9", border=1, align="C")
        self.cell(15, 30, "-", border=1, align="C")
        
        self.set_xy(x + 95, y_dados)
        self.multi_cell(75, 5, f"\n{ctx['FIBREBOARD']} FIBREBOARD BOXES {ctx['PESO_G']} Kg G\nOVERPACK USED X {ctx['QTD_OVERPACK']}\n{ctx['MARCACAO']}\nTOTAL QUANTITY PER OVERPACK {ctx['TOTAL_OVERPACK']} Kg G", border=0)
        
        self.set_xy(x + 95, y_dados)
        self.cell(75, 30, "", border=1)
        self.cell(20, 30, "Y963", border=1, ln=1, align="C")
        
        # Termos e Assinatura
        self.ln(5)
        self.set_font("Arial", "B", 8)
        self.multi_cell(0, 4, "AVISO: A falha em cumprir em todos os aspectos com a regulamentacao aplicavel de artigos perigosos sera transgressao as leis em vigor e sujeita as penalidades legais.")
        self.ln(2)
        self.set_font("Arial", "", 8)
        self.multi_cell(0, 4, "Declaro que o conteudo desta remessa esta completa e precisamente descrito acima pelo nome apropriado para embarque, que esta classificado, embalado, marcado e etiquetado de acordo com as regulamentacoes aplicaveis.")
        
        self.ln(5)
        self.cell(0, 6, f"Data da Emissao: {ctx['DATA']}", ln=1)

# 2. MAPA DE TRADUÇÃO
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE", "CGB": "CUIABA", "CWB": "CURITIBA", 
    "FLN": "FLORIANOPOLIS", "GYN": "GOIANIA", "MAO": "MANAUS", 
    "POA": "PORTO ALEGRE", "PVH": "PORTO VELHO"
}

st.set_page_config(page_title="New Post - Shippers 2.0 PDF", layout="wide")
st.title("🚀 Gerador de Shippers New Post - V2.0")

siglas_input = st.text_input("Siglas dos Destinos (Ex: CGB, POA, MAO):").upper().strip()
file = st.file_uploader("Upload da Planilha de Coleta Base (.xlsm ou .xlsx)", type=["xlsm", "xlsx"])

if file and siglas_input:
    try:
        df = pd.read_excel(file, header=None, engine='openpyxl')
        lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]

        if st.button(f"🔢 CALCULAR E EMITIR {len(lista_siglas)} SHIPPERS EM PDF"):
            zip_buffer = io.BytesIO()
            emitidos = []

            with ZipFile(zip_buffer, "w") as zip_file:
                for sigla in lista_siglas:
                    cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                    
                    linha_dados = None
                    for index, row in df.iterrows():
                        linha_texto = " ".join([str(val).upper() for val in row.values if pd.notnull(val)])
                        if cidade_alvo in linha_texto and "TOTAL" not in linha_texto:
                            linha_dados = row
                            break

                    if linha_dados is not None:
                        v_sacas = linha_dados[5]
                        v_peso_real = linha_dados[6]

                        g7_peso_real = Decimal(str(pd.to_numeric(v_peso_real, errors='coerce')))
                        f7_qtd_sacas = Decimal(str(int(v_sacas) if pd.notnull(v_sacas) else 7))
                        
                        if sigla == "CGB" and f7_qtd_sacas == 7:
                            i7_fib = Decimal('4')
                        else:
                            v_i = float(g7_peso_real / f7_qtd_sacas) / 4.5
                            i7_fib = Decimal(str(math.ceil(v_i) if (v_i - int(v_i)) > 0.50 else math.floor(v_i)))

                        j7_kg_g = (g7_peso_real / f7_qtd_sacas / i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        
                        while True:
                            k7_saca = (j7_kg_g * i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                            l7_total_simulado = k7_saca * f7_qtd_sacas
                            m7_saldo = l7_total_simulado - g7_peso_real
                            if m7_saldo >= 0: break
                            else: j7_kg_g += Decimal('0.01')

                        contexto = {
                            'CIDADE': cidade_alvo,
                            'FIBREBOARD': int(i7_fib),
                            'PESO_G': "{:.2f}".format(j7_kg_g).replace('.', ','),
                            'TOTAL_OVERPACK': "{:.2f}".format(k7_saca).replace('.', ','),
                            'MARCACAO': " ".join([f"#{i+1}" for i in range(int(f7_qtd_sacas))]),
                            'DATA': date.today().strftime('%d/%m/%Y'),
                            'QTD_OVERPACK': int(f7_qtd_sacas)
                        }

                        # Geração do PDF isolada e segura para a memória
                        pdf = ShipperPDF()
                        pdf.layout_shipper(contexto)
                        
                        # Salvando o PDF diretamente em formato de bytes estável para o ZIP
                        pdf_bytes = bytes(pdf.output())
                        zip_file.writestr(f"Shipper_{sigla}.pdf", pdf_bytes)
                        emitidos.append(sigla)

            if emitidos:
                zip_buffer.seek(0)
                st.success(f"✅ Sucesso! PDFs gerados: {', '.join(emitidos)}")
                st.download_button(
                    label="📥 BAIXAR SHIPPERS EM PDF (ZIP)",
                    data=zip_buffer,
                    file_name=f"Shippers_PDF_{date.today()}.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
