import streamlit as st
import pandas as pd
import io
import math
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES
MAPA_DESTINOS = {
    "CGR": "CAMPO GRANDE", "CGB": "CUIABA", "CWB": "CURITIBA", 
    "FLN": "FLORIANOPOLIS", "GYN": "GOIANIA", "MAO": "MANAUS", 
    "POA": "PORTO ALEGRE", "PVH": "PORTO VELHO"
}

st.set_page_config(page_title="New Post - Shippers 2.0 HTML/PDF", layout="wide")
st.title("🚀 Gerador de Shippers New Post - V2.0 (Direct Web)")
st.subheader("Cálculo de Saldo Automático - Pronto para Imprimir")

siglas_input = st.text_input("Siglas do Destino (Ex: CGB, POA, MAO):").upper().strip()
file = st.file_uploader("Upload da Planilha de Coleta Base (.xlsm ou .xlsx)", type=["xlsm", "xlsx"])

# Função que desenha a folha oficial direto em formato Web imprimível
def gerar_html_shipper(ctx):
    html_content = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; color: #000; font-size: 12px; margin: 20px; }}
            .container {{ width: 700px; margin: 0 auto; border: 2px solid #000; padding: 10px; }}
            .title {{ text-align: center; font-weight: bold; font-size: 16px; padding: 10px; background-color: #e6e6e6; border-bottom: 2px solid #000; }}
            .table-block {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
            .table-block td, .table-block th {{ border: 1px solid #000; padding: 8px; vertical-align: top; }}
            .bg-gray {{ background-color: #f2f2f2; font-weight: bold; }}
            .footer-box {{ font-size: 10px; text-align: justify; margin-top: 15px; border-top: 1px solid #000; padding-top: 10px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="title">DECLARAÇÃO DO EXPEDIDOR PARA ARTIGOS PERIGOSOS</div>
            <table class="table-block">
                <tr>
                    <td style="width: 50%;"><strong>Expedidor:</strong><br>NEW POST SOLUÇÕES EM LOGISTICAS LTDA<br>CNPJ: 28.678.104/0001-79<br>R UBALDO FAGGEANI, 355 - JARDIM LAS PALMAS<br>PORTO FERREIRA/SP - CEP: 13667-262</td>
                    <td style="width: 50%;"><strong>Nº do Conhecimento Aéreo:</strong><br><br><strong>Página 1 de 1 Páginas</strong></td>
                </tr>
                <tr>
                    <td><strong>Consignatário:</strong><br>BRASIL POSTAL LTDA – ME<br>AV FERNANDO CORREA DA COSTA, 3180<br>SHANGRILA – {ctx['CIDADE']}<br>CEP: 78070-200</td>
                    <td><strong>Nº de Referência do Expedidor:</strong><br>(Opcional)</td>
                </tr>
            </table>

            <table class="table-block">
                <tr class="bg-gray"><td colspan="2">DETALHES DE TRANSPORTE</td></tr>
                <tr>
                    <td style="width: 50%;">Este embarque está dentro das limitações prescritas para:<br><strong>AERONAVE DE PASSAGEIROS E CARGA</strong></td>
                    <td style="width: 50%;">Aeroporto de Origem: <strong>CONFINS</strong><br>Aeroporto de Destino: <strong>{ctx['CIDADE']}</strong></td>
                </tr>
            </table>

            <table class="table-block">
                <tr class="bg-gray"><td colspan="6">NATUREZA E QUANTIDADE DE ARTIGOS PERIGOSOS</td></tr>
                <tr style="background-color: #fafafa; font-weight: bold; text-align: center;">
                    <td>Nº UN</td><td>Nome Apropriado</td><td>Classe</td><td>Grupo</td><td>Quantidade e Tipo de Embalagem</td><td>Inst.</td>
                </tr>
                <tr>
                    <td style="text-align: center;">ID8000</td>
                    <td>CONSUMER COMMODITY</td>
                    <td style="text-align: center;">9</td>
                    <td style="text-align: center;">-</td>
                    <td>
                        <strong>{ctx['FIBREBOARD']} FIBREBOARD BOXES {ctx['PESO_G']} Kg G</strong><br>
                        OVERPACK USED X {ctx['QTD_OVERPACK']}<br>
                        <small>{ctx['MARCACAO']}</small><br>
                        <strong>TOTAL QUANTITY PER OVERPACK {ctx['TOTAL_OVERPACK']} Kg G</strong>
                    </td>
                    <td style="text-align: center;">Y963</td>
                </tr>
            </table>

            <div class="footer-box">
                <strong>AVISO:</strong> A falha em cumprir em todos os aspectos com a regulamentação aplicável de artigos perigosos será transgressão às leis em vigor e sujeita às penalidades legais.<br><br>
                Declaro que o conteúdo desta remessa está completa e precisamente descrito acima pelo nome apropriado para embarque, que está classificado, embalado, marcado e etiquetado de acordo com as normas da ICA / IATA.<br><br>
                <strong>Data da Emissão:</strong> {ctx['DATA']}
            </div>
        </div>
        <script>window.print();</script>
    </body>
    </html>
    """
    return html_content

if file and siglas_input:
    try:
        df = pd.read_excel(file, header=None, engine='openpyxl')
        lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]

        if st.button(f"🔢 CALCULAR E EMITIR {len(lista_siglas)} SHIPPERS"):
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
                        # Pega os valores por posição física (Coluna F e G)
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
                        
                        # Loop de ajuste automático do saldo M >= 0
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

                        # Monta o arquivo HTML pronto
                        conteudo_html = gerar_html_shipper(contexto)
                        zip_file.writestr(f"Shipper_{sigla}.html", conteudo_html.encode('utf-8'))
                        emitidos.append(sigla)

            if emitidos:
                zip_buffer.seek(0)
                st.success(f"✅ Sucesso! Shippers calculadas para: {', '.join(emitidos)}")
                st.download_button(
                    label="📥 DOWNLOAD DAS SHIPPERS (ZIP)",
                    data=zip_buffer,
                    file_name=f"Shippers_NewPost_{date.today()}.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
