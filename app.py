import streamlit as st

st.set_page_config(page_title="Gerador de Shippers - Artigos Perigosos", layout="wide")
st.title("📦 Gerador de Declaração do Expedidor (Shipper)")

# --- BARRA LATERAL / ENTRADA DE DADOS ---
st.sidebar.header("1. Informações do Voo")
aeroporto_origem = st.sidebar.text_input("Aeroporto de Origem", value="CONFINS").upper().strip()
aeroporto_destino = st.sidebar.text_input("Aeroporto de Destino", value="GOIANIA").upper().strip()

st.sidebar.header("2. Dados da Carga")
total_overpacks = st.sidebar.number_input("Quantidade Total de Overpacks (Sacas)", min_value=1, value=13, step=1)
qtd_caixas_por_overpack = st.sidebar.number_input("Quantidade de Caixas por Overpack", min_value=1, value=4, step=1)

# Entrada do peso bruto total de toda a remessa (somando todas as sacas)
peso_bruto_total = st.sidebar.number_input(
    "Peso Bruto Total do Lote (Kg G) - Digite o total geral", 
    min_value=0.0, 
    value=240.76, 
    step=0.01,
    format="%.2f"
)

# --- PROCESSAMENTO DOS DADOS (Ajuste Matemático) ---
if total_overpacks > 0 and qtd_caixas_por_overpack > 0:
    
    # REGRA 1: Descobre o peso bruto individual de cada caixa e ARREDONDA para 2 casas decimais exatas
    # Isso evita dízimas periódicas ocultas no Python
    peso_caixa_individual = round(peso_bruto_total / (total_overpacks * qtd_caixas_por_overpack), 2)
    
    # REGRA 2: O peso do overpack passa a ser a multiplicação DIRETA do peso redondo da caixa
    # Dessa forma, a conta impressa (ex: 4 caixas x 4,63 = 18,52) será matematicamente perfeita para a Cia Aérea
    peso_total_do_overpack = round(qtd_caixas_por_overpack * peso_caixa_individual, 2)
    
    # Recalcula o total final real do documento para bater com a soma das sacas
    peso_final_declarado = round(total_overpacks * peso_total_do_overpack, 2)

    # --- MONTAGEM DO LAYOUT DE TEXTO (Ajuste das Vírgulas) ---
    id_un_id = "ID8000"  # Sem espaços internos
    nome_artigo = "CONSUMER COMMODITY"
    classe_divisao = "9"
    instrucao_embalagem = "Y963"
    
    # Cria a sequência automática de sacas (#1 #2 #3 ... #13)
    sequencia_sacas = " ".join([f"#{i}" for i in range(1, int(total_overpacks) + 1)])
    
    # Substituindo a concatenação por vírgulas consecutivas por quebras de linha '\n'
    # Isso mantém o bloco de texto unido dentro do campo correto da Shipper
    texto_quantidade_tipo = (
        f"{int(qtd_caixas_por_overpack)} FIBREBOARD BOXES x {peso_caixa_individual:.2f} Kg G\n"
        f"OVERPACK USED X {int(total_overpacks)}\n"
        f"{sequencia_sacas}\n"
        f"TOTAL QUANTITY PER\n"
        f"OVERPACK {peso_total_do_overpack:.2f} Kg G"
    )

    # --- INTERFACE VISUAL DO STREAMLIT ---
    st.subheader("📝 Pré-visualização dos Dados Calculados")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Peso Unitário por Caixa", f"{peso_caixa_individual:.2f} Kg G")
    with col2:
        st.metric("Peso Total por Overpack (Saca)", f"{peso_total_do_overpack:.2f} Kg G")
    with col3:
        st.metric("Peso Final da Remessa", f"{peso_final_declarado:.2f} Kg G")

    st.markdown("---")
    st.subheader("📋 Visualização de como o bloco será impresso na Shipper:")
    
    # Tabela simulando o preenchimento do formulário
    st.table({
        "Nº UN ou ID": [id_un_id],
        "Nome Apropriado para Embarque": [nome_artigo],
        "Classe": [classe_divisao],
        "Quantidade e Tipo de Embalagem (BLOCO CORRIGIDO)": [texto_quantidade_tipo.replace('\n', ' | ')],
        "Instrução": [instrucao_embalagem]
    })
    
    # Código para ver a quebra de linha real
    st.text_area("Texto bruto gerado internamente:", value=texto_quantidade_tipo, height=120)

    # --- ÁREA DE EXPORTAÇÃO (Simulada) ---
    st.markdown("---")
    if st.button("💾 Gerar Arquivo Final da Shipper"):
        # Aqui entra a sua função atual de geração do Word (.docx) ou PDF.
        # Você só precisa passar as variáveis corrigidas:
        # -> texto_quantidade_tipo (para a coluna de quantidade)
        # -> peso_final_declarado (para o peso total do documento)
        
        st.success(f"Documento estruturado com sucesso para {aeroporto_destino}! Pronto para exportação.")
        
else:
    st.warning("Por favor, preencha as quantidades na barra lateral para calcular.")
