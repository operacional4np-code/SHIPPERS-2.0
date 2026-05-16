import streamlit as st
import pandas as pd
import io
import math
import re
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from zipfile import ZipFile

# MAPA DE TRADUÇÃO DAS CIDADES E DADOS DO CONSIGNATÁRIO DE REFERÊNCIA
MAPA_DESTINOS = {
    "CGR": {"nome": "CAMPO GRANDE", "endereco": "AV CORONEL ANTONINO, 1200 - CAMPO GRANDE/MS - CEP: 79013-000"},
    "CGB": {"nome": "CUIABA", "endereco": "AV FERNANDO CORREA DA COSTA, 3180 - SHANGRILA - CUIABA/MT - CEP: 78070-200"},
    "CWB": {"nome": "CURITIBA", "endereco": "RUA JOAO NEGRAO, 1250 - REBOUCAS - CURITIBA/PR - CEP: 80230-150"},
    "FLN": {"nome": "FLORIANOPOLIS", "endereco": "RUA PASCHOAL APOLONIO PASCOAL, 45 - CENTRO - FLORIANOPOLIS/SC - CEP: 88010-460"},
    "GYN": {"nome": "GOIANIA", "endereco": "AV TOCANTINS, 250 - SETOR CENTRAL - GOIANIA/GO - CEP: 74015-010"},
    "MAO": {"nome": "MANAUS", "endereco": "AV MAPA, 120 - DISTRITO INDUSTRIAL - MANAUS/AM - CEP: 69075-000"},
    "POA": {"nome": "PORTO ALEGRE", "endereco": "RUA DOS ANDRADAS, 1444 - CENTRO HISTORICO - PORTO ALEGRE/RS - CEP: 90020-010"},
    "PVH": {"nome": "PORTO VELHO", "endereco": "AV JORGE TEIXEIRA, 1500 - INDUSTRIAL - PORTO VELHO/RO - CEP: 76821-001"}
}

st.set_page_config(page_title="New Post - Shippers Oficiais IATA", layout="wide")
st.title("✈️ Emissor de Shippers Oficial New Post")
st.subheader("Layout Idêntico à Referência com Bordas Regulamentares")

siglas_input = st.text_input("1. Digite as Siglas dos Destinos separadas por vírgula (Ex: CGB, POA, MAO):").upper().strip()
file = st.file_uploader("2. Carregue a Planilha de Informações", type=["xlsm", "xlsx"])

def gerar_html_shipper_oficial(ctx):
    # Gera o fundo com as listras vermelhas nas laterais imitando o papel regulamentar da IATA
    html_content = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            @page {{ size: A4; margin: 0; }}
            body {{ 
                font-family: 'Arial Narrow', Arial, sans-serif; 
                color: #000; 
                margin: 0; 
                padding: 0;
                background-color: #fff;
            }}
            /* Bordas vermelhas regulamentares da IATA nas laterais */
            .page-border {{
                box-sizing: border-box;
                width: 210mm;
                height: 297mm;
                padding: 12mm 15mm;
                border-left: 8mm repeating-linear-gradient(-45deg, #d9534f, #d9534f 10px, #fff 10px, #fff 20px);
                border-right: 8mm repeating-linear-gradient(-45deg, #d9534f, #d9534f 10px, #fff 10px, #fff 20px);
                position: relative;
            }}
            .main-title {{
                text-align: center;
                font-size: 15pt;
                font-weight: bold;
                letter-spacing: 0.5px;
                margin-bottom: 5px;
            }}
            .sub-title {{
                text-align: center;
                font-size: 11pt;
                font-weight: bold;
                margin-bottom: 12px;
            }}
            .w-100 {{ width: 100%; }}
            .tbl {{
                width: 100%;
                border-collapse: collapse;
                margin-bottom: -1px;
            }}
            .tbl td {{
                border: 1px solid #000;
                padding: 5px 7px;
                vertical-align: top;
                font-size: 8.5pt;
                line-height: 1.2;
            }}
            .label {{
                font-size: 7.5pt;
                font-weight: bold;
                text-transform: uppercase;
                display: block;
                margin-bottom: 3px;
                color: #111;
            }}
            .content-text {{
                font-size: 9pt;
                font-weight: normal;
            }}
            .section-title {{
                background-color: #f2f2f2;
                font-weight: bold;
                text-align: left;
                font-size: 8.5pt;
                border: 1px solid #000;
                padding: 3px 7px;
                text-transform: uppercase;
            }}
            .center {{ text-align: center; }}
            .warning-text {{
                font-size: 7pt;
                text-align: justify;
                line-height: 1.2;
                margin-top: 8px;
            }}
            .decl-text {{
                font-size: 8pt;
                text-align: justify;
                line-height: 1.3;
                margin-top: 8px;
                font-weight: normal;
            }}
        </style>
    </head>
    <body>
        <div class="page-border">
            <div class="main-title">DECLARAÇÃO DO EXPEDIDOR PARA ARTIGOS PERIGOSOS</div>
            <div class="sub-title">(SHIPPER'S DECLARATION FOR DANGEROUS GOODS)</div>
            
            <table class="tbl">
                <tr>
                    <td style="width: 55%; height: 85px;">
                        <span class="label">Expedidor (Shipper)</span>
                        <div class="content-text">
                            <strong>NEW POST SOLUÇÕES EM LOGISTICAS LTDA</strong><br>
                            CNPJ.: 28.678.104/0001-79<br>
                            R UBALDO FAGGEANI, 355 - JARDIM LAS PALMAS<br>
                            PORTO FERREIRA/SP - CEP: 13667-262
                        </div>
                    </td>
                    <td style="width: 45%;">
                        <span class="label">Número do Conhecimento Aéreo (Air Waybill No.)</span>
                        <div class="content-text" style="height: 30px;"></div>
                        <hr style="border: 0; border-top: 1px solid #000; margin: 3px 0;">
                        <table style="width: 100%; border: 0;">
                            <tr>
                                <td style="border: 0; padding: 0;"><span class="label">Página (Page)</span><strong>1</strong></td>
                                <td style="border: 0; padding: 0;"><span class="label">de (of)</span><strong>1</strong></td>
                                <td style="border: 0; padding: 0;"><span class="label">Páginas (Pages)</span></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="height: 85px;">
                        <span class="label">Consignatário (Consignee)</span>
                        <div class="content-text">
                            <strong>AGENCIA DE CORREIOS FRANQUEADA GLOBO ANDRADAS LTDA</strong><br>
                            {ctx['ENDERECO_CONSIGNATARIO']}
                        </div>
                    </td>
                    <td>
                        <span class="label">Nº de Referência do Expedidor (Shipper's Reference Number)</span>
                        <div class="content-text">(opcional)</div>
                    </td>
                </tr>
            </table>

            <div class="warning-text" style="font-weight: bold; border-left: 1px solid #000; border-right: 1px solid #000; padding: 2px 7px; margin: 0;">
                Duas cópias preenchidas e assinadas desta declaração devem ser entregues ao operador aéreo.
            </div>

            <div class="section-title">Detalhes de Transporte (Transport Details)</div>
            <table class="tbl">
                <tr>
                    <td style="width: 55%;">
                        <span class="label">Este embarque está dentro das limitações prescritas para: (deletar o campo não aplicável)</span>
                        <table style="width: 100%; border: 0; font-size: 8pt; margin-top: 5px;">
                            <tr>
                                <td style="border: 0; padding: 0; width: 50%;"><s>AERONAVE DE PASSAGEIROS E CARGA</s></td>
                                <td style="border: 0; padding: 0; width: 50%;"><strong>SOMENTE AERONAVE DE CARGA</strong></td>
                            </tr>
                        </table>
                    </td>
                    <td style="width: 45%;">
                        <span class="label">Aeroporto de Origem (Airport of Departure):</span>
                        <div class="content-text"><strong>CONFINS</strong></div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span class="label">Tipo de expedição: (deletar o campo não aplicável)</span>
                        <div class="content-text" style="margin-top: 5px;">
                            <strong>NÃO RADIOATIVO</strong> / <s>RADIOATIVO</s>
                        </div>
                    </td>
                    <td>
                        <span class="label">Aeroporto de Destino (Airport of Destination):</span>
                        <div class="content-text"><strong>{ctx['CIDADE']}</strong></div>
                    </td>
                </tr>
            </table>

            <div class="section-title">Natureza e Quantidade de Artigos Perigosos (Nature and Quantity of Dangerous Goods)</div>
            <table class="tbl" style="text-align: center;">
                <tr style="background-color: #f9f9f9; font-weight: bold; font-size: 7.5pt;">
                    <td style="width: 10%;">N° UN<br>ou ID</td>
                    <td style="width: 32%;">Nome apropriado para embarque</td>
                    <td style="width: 12%;">Classe ou<br>Divisão</td>
                    <td style="width: 10%;">Grupo de<br>Embalagem</td>
                    <td style="width: 24%;">Quantidade e tipo de embalagem</td>
                    <td style="width: 12%;">Instrução de<br>Embalagem</td>
                </tr>
                <tr style="height: 180px; font-size: 9pt; text-align: left;">
                    <td class="center" style="vertical-align: top; padding-top: 10px;"><strong>ID 8000</strong></td>
                    <td style="vertical-align: top; padding-top: 10px;">CONSUMER COMMODITY</td>
                    <td class="center" style="vertical-align: top; padding-top: 10px;"><strong>9</strong></td>
                    <td class="center" style="vertical-align: top; padding-top: 10px;"></td>
                    <td style="vertical-align: top; padding-top: 10px; line-height: 1.4;">
                        <strong>{ctx['FIBREBOARD']} FIBREBOARD BOXES x {ctx['PESO_G']} Kg G</strong><br><br>
                        <strong>OVERPACK USED X {ctx['QTD_OVERPACK']}</strong><br>
                        <span style="font-size: 8pt; word-spacing: 2px;">{ctx['MARCACAO']}</span><br><br>
                        <strong>TOTAL QUANTITY PER OVERPACK {ctx['TOTAL_OVERPACK']} Kg G</strong>
                    </td>
                    <td class="center" style="vertical-align: top; padding-top: 10px;"><strong>Y963</strong></td>
                </tr>
            </table>

            <table class="tbl">
                <tr>
                    <td style="height: 55px;">
                        <span class="label">Informações Adicionais de Manuseio (Additional Handling Information)</span>
                        <div class="content-text" style="font-size: 9pt; margin-top: 3px;">
                            <strong>24-hour Number: +55 11 2898-9807</strong>
                        </div>
                    </td>
                </tr>
            </table>

            <div class="warning-text">
                <strong>AVISO:</strong> A falha em cumprir em todos os aspectos com a regulamentação aplicável de artigos perigosos será transgressão às leis em vigor e sujeita às penalidades legais.
            </div>

            <div class="decl-text">
                Declaro que o conteúdo desta remessa está completa e precisamente descrito acima pelo nome apropriado para embarque, e está classificado, embalado, marcado e etiquetado, e está em todas as condições para transporte de acordo com as regulamentações aplicáveis do Governo e da ICAO/IATA.
            </div>

            <table class="tbl" style="margin-top: 15px;">
                <tr>
                    <td style="width: 55%; height: 85px;">
                        <span class="label">Nome/Título do Signatário (Name/Title of Signatory)</span>
                        <div class="content-text" style="margin-top: 25px;">
                            <strong>NEW POST SOLUÇÕES EM LOGÍSTICA</strong>
                        </div>
                    </td>
                    <td style="width: 45%;">
                        <span class="label">Local e Data (Place and Date)</span>
                        <div class="content-text" style="margin-top: 25px;">
                            PORTO FERREIRA, {ctx['DATA']}
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="height: 70px;">
                        <span class="label">Assinatura (Signature) - (Não escrever nesta tarja regulamentar)</span>
                        <div style="border: 1px dashed #ccc; height: 40px; margin-top: 5px; text-align: center; color: #999; line-height: 40px; font-size: 8pt;">
                            Espaço destinado para assinatura manual obrigatória do expedidor
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <script>window.print();</script>
    </body>
    </html>
    """
    return html_content

def extrair_peso_seguro(linha_row):
    valores_numericos = []
    for val in linha_row:
        if pd.notnull(val) and not isinstance(val, str):
            try:
                v_float = float(val)
                if v_float > 0: valores_numericos.append(v_float)
            except: pass
        elif isinstance(val, str):
            txt_limpo = re.sub(r'[^\d.,]', '', val).strip()
            if "," in txt_limpo and "." in txt_limpo:
                txt_limpo = txt_limpo.replace(".", "").replace(",", ".")
            elif "," in txt_limpo:
                txt_limpo = txt_limpo.replace(",", ".")
            try:
                v_float = float(txt_limpo)
                if v_float > 0: valores_numericos.append(v_float)
            except: pass
    return valores_numericos[0] if valores_numericos else 0.0

if siglas_input:
    lista_siglas = [s.strip() for s in siglas_input.split(",") if s.strip()]
    
    st.markdown("### 3. Informe a quantidade de sacas para cada destino:")
    sacas_manuais = {}
    
    colunas_tela = st.columns(
