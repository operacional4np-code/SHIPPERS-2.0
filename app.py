import streamlit as st
import pandas as pd
import io
import math
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from zipfile import ZipFile

# Para a conversão do HTML preenchido em PDF estável
from weasyprint import HTML

# 1. CONFIGURAÇÃO DA VERSÃO 2.0
st.set_page_config(page_title="New Post - Shippers 2.0 PDF", layout="wide")
st.title("🚀 Gerador de Shippers New Post - V2.0 (PDF Direto)")
st.subheader("Cálculo Dinâmico, Ajuste de Saldo Automático e Emissão em Lote")

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

# 2. ENTRADA DE DADOS
col1, col2 = st.columns([1, 2])
with col1:
    siglas_input = st.text_input("Siglas dos Destinos (Ex: CGB, POA, MAO):").upper().strip()
with col2:
    file = st.file_uploader("Upload da Planilha de Coleta Base (.xlsm ou .xlsx)", type=["xlsm", "xlsx"])

# 3. TEMPLATE BASE EM HTML (Substitui o Word e gera o PDF idêntico)
def obter_html_template(dados_contexto):
    # Aqui fica o design visual completo do documento da Shipper
    html_content = f"""
    <html>
    <head>
        <style>
            @page {{ size: A4; margin: 15mm 12mm; }}
            body {{ font-family: 'Arial', sans-serif; color: #000; font-size: 10pt; margin: 0; padding: 0; }}
            .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
            .header-table td {{ border: 1px solid #000; padding: 6px; vertical-align: top; }}
            .title {{ text-align: center; font-weight: bold; font-size: 14pt; padding: 10px; background-color: #f2f2f2; }}
            .section-title {{ font-weight: bold; background-color: #e6e6e6; padding: 4px; font-size: 10pt; border: 1px solid #000; margin-top: 10px; }}
            .content-table {{ width: 100%; border-collapse: collapse; margin-top: -1px; }}
            .content-table th, .content-table td {{ border: 1px solid #000; padding: 6px; text-align: left; vertical-align: top; }}
            .footer-text {{ font-size: 8pt; text-align: justify; margin-top: 15px; }}
        </style>
    </head>
    <body>
        <div class="title">DECLARAÇÃO DO EXPEDIDOR PARA ARTIGOS PERIGOSOS</div>
        <table class="header-table">
            <tr>
                <td style="width: 50%;"><strong>Expedidor:</strong><br>NEW POST SOLUÇÕES EM LOGISTICAS LTDA<br>CNPJ: 28.678.104/0001-79<br>R UBALDO FAGGEANI, 355 - JARDIM LAS PALMAS<br>PORTO FERREIRA/SP - CEP: 13667-262</td>
                <td style="width: 50%;"><strong>Nº do Conhecimento Aéreo:</strong><br><br><strong>Página 1 de 1 Páginas</strong></td>
            </tr>
            <tr>
                <td><strong>Consignatário:</strong><br>BRASIL POSTAL LTDA – ME<br>AV FERNANDO CORREA DA COSTA, 3180<br>SHANGRILA – {dados_contexto['CIDADE']}<br>CEP: 78070-200</td>
                <td><strong>Nº de Referência do Expedidor:</strong><br>(Opcional)</td>
            </tr>
        </table>

        <div class="section-title">DETALHES DE TRANSPORTE</div>
        <table class="content-table">
            <tr>
                <td style="width: 50%;">Este embarque está dentro das limitações prescritas para:<br><strong>AERONAVE DE PASSAGEIROS E CARGA</strong></td>
                <td style="width: 50%;">Aeroporto de Origem: <strong>CONFINS</strong><br>Aeroporto de Destino: <strong>{dados_contexto['CIDADE']}</strong></td>
            </tr>
        </table>

        <div class="section-title">NATUREZA E QUANTIDADE DE ARTIGOS PERIGOSOS</div>
        <table class="content-table">
            <thead>
                <tr style="background-color: #f2f2f2;">
                    <th>Nº UN</th>
                    <th>Nome Apropriado</th>
                    <th>Classe</th>
                    <th>Grupo Emb.</th>
                    <th>Quantidade e Tipo de Embalagem</th>
                    <th>Inst. Emb.</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>ID8000</td>
                    <td>CONSUMER COMMODITY</td>
                    <td>9</td>
                    <td>-</td>
                    <td>
                        <strong>{dados_contexto['FIBREBOARD']} FIBREBOARD BOXES {dados_contexto['PESO_G']} Kg G</strong><br>
                        OVERPACK USED X {dados_contexto['QTD_OVERPACK']}<br>
                        <small>{dados_contexto['MARCACAO']}</small><br>
                        <strong>TOTAL QUANTITY PER OVERPACK {dados_contexto['TOTAL_OVERPACK']} Kg G</strong>
                    </td>
                    <td>Y963</td>
                </tr>
            </tbody>
        </table>

        <div class="footer-text">
            <strong>AVISO:</strong> A falha em cumprir em todos os aspectos com a regulamentação aplicável de artigos perigosos será transgressão às leis em vigor e sujeita às penalidades legais.<br><br>
            Declaro que o conteúdo desta remessa está completa e precisamente descrito acima pelo nome apropriado para embarque, que está classificado, embalado, marcado e etiquetado...
            <br><br>
            <strong>Data da Emissão:</strong> {dados_contexto['DATA']}
        </div>
    </body>
    </html>
    """
    return html_content

# 4. MOTOR DE CÁLCULO E LOGICA DE COMPENSAÇÃO (M >= 0)
if file and siglas_input:
    try:
        df_raw = pd.read_excel(file, header=None, engine='openpyxl')
        
        # Identifica a linha do cabeçalho da planilha de coleta dinamicamente
        header_row = 0
        for i, row in df_raw.iterrows():
            if "DESTINO" in [str(val).upper() for val in row.values]:
                header_row = i
                break
        
        df = pd.read_excel(file, header=header_row, engine='openpyxl')
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]

        if st.button(f"🔢 CALCULAR E EMITIR {len(lista_siglas)} SHIPPERS EM PDF"):
            zip_buffer = io.BytesIO()
            emitidos = []

            c_dest = next((c for c in df.columns if "DESTINO" in c), None)
            c_peso = next((c for c in df.columns if "PESO" in c), None)

            with ZipFile(zip_buffer, "w") as zip_file:
                for sigla in lista_siglas:
                    cidade_alvo = MAPA_DESTINOS.get(sigla, sigla)
                    
                    # Filtra a linha correspondente da planilha de coleta
                    df_f = df[df[c_dest].astype(str).str.contains(cidade_alvo, case=False, na=False)].copy()
                    df_f = df_f[~df_f[c_dest].astype(str).str.upper().str.contains("TOTAL", na=False)]

                    if not df_f.empty:
                        # --- EXECUÇÃO DOS CÁLCULOS AUTOMÁTICOS ---
                        # Coluna G (Peso Real coletado)
                        g7_peso_real = Decimal(str(pd.to_numeric(df_f[c_peso], errors='coerce').sum()))
                        
                        # Coluna F (Quantidade de Sacas extraída da Coluna F da linha encontrada)
                        c_sacas = df_f.iloc[0].get('SACAS', df_f.iloc[0].get('QTD', 7)) # fallback para 7 se não achar
                        f7_qtd_sacas = Decimal(str(int(c_sacas) if pd.notnull(c_sacas) else 7))
                        
                        # Coluna I: Determinação de Caixas (Fibreboard)
                        if sigla == "CGB" and f7_qtd_sacas == 7:
                            i7_fib = Decimal('4')
                        else:
                            v_i = float(g7_peso_real / f7_qtd_sacas) / 4.5
                            i7_fib = Decimal(str(math.ceil(v_i) if (v_i - int(v_i)) > 0.50 else math.floor(v_i)))

                        # Coluna J: Inicialização do Kg G
                        j7_kg_g = (g7_peso_real / f7_qtd_sacas / i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                        
                        # LOOP DE AJUSTE AUTOMÁTICO DE SALDO (Simulando as colunas L e M da planilha)
                        while True:
                            k7_saca = (j7_kg_g * i7_fib).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP) # Coluna K
                            l7_total_simulado = k7_saca * f7_qtd_sacas                                    # Coluna L
                            m7_saldo = l7_total_simulado - g7_peso_real                                   # Coluna M
                            
                            if m7_saldo >= 0:
                                break  # Meta batida: saldo zerado ou estritamente positivo para não faltar peso
                            else:
                                j7_kg_g += Decimal('0.01') # Incrementa centavo por centavo se M for negativo

                        # Formatação final para o documento
                        txt_kg_g = "{:.2f}".format(j7_kg_g).replace('.', ',')
                        txt_total_k = "{:.2f}".format(k7_saca).replace('.', ',')
                        marcacao = " ".join([f"#{i+1}" for i in range(int(f7_qtd_sacas))])

                        # Monta o contexto de substituição do PDF
                        contexto = {
                            'CIDADE': cidade_alvo,
                            'FIBREBOARD': int(i7_fib),
                            'PESO_G': txt_kg_g,
                            'TOTAL_OVERPACK': txt_total_k,
                            'MARCACAO': marcacao,
                            'DATA': date.today().strftime('%d/%m/%Y'),
                            'QTD_OVERPACK': int(f7_qtd_sacas)
                        }

                        # Renderiza o HTML e transforma em PDF binário na memória
                        html_renderizado = obter_html_template(contexto)
                        pdf_file = io.BytesIO()
                        HTML(string=html_renderizado).write_pdf(pdf_file)
                        
                        # Adiciona o arquivo .pdf pronto dentro do ZIP de saída
                        zip_file.writestr(f"Shipper_{sigla}.pdf", pdf_file.getvalue())
                        emitidos.append(sigla)
                    else:
                        st.warning(f"⚠️ Destino {sigla} ({cidade_alvo}) não localizado na planilha de coleta.")

            if emitidos:
                zip_buffer.seek(0)
                st.success(f"✅ Sucesso! Shippers Calculadas e Convertidas para PDF: {', '.join(emitidos)}")
                st.download_button(
                    label="📥 BAIXAR SHIPPERS EM PDF (ARQUIVO .ZIP)",
                    data=zip_buffer,
                    file_name=f"Shippers_PDF_NewPost_{date.today()}.zip",
                    mime="application/zip"
                )
    except Exception as e:
        st.error(f"Erro operacional no processamento dos PDFs: {e}")
