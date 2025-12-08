# ============================== ETIQUETAS DE PACOTES - VERSÃO FINAL ==============================
# Autor: André Filipe Obenaus
# Objetivo: Extrai PDF → gera Excel com uma linha por caixa → nome perfeito
# Tudo em um único script limpo e confiável

import pdfplumber
import pandas as pd
import re
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------- CONFIGURAÇÕES --------------------------
BASE_QUANTIDADES_FILE = Path("base_quantities.xlsx")
DEFAULT_BOX_CAPACITY = 10
# ------------------------------------------------------------------

script_dir = Path(__file__).parent
pdf_files = list(script_dir.glob("*.pdf"))

if not pdf_files:
    raise FileNotFoundError("Nenhum PDF encontrado na pasta do script.")
input_pdf = pdf_files[0]

print(f"Processando: {input_pdf.name}")

# -------------------------- EXTRAÇÃO DO PDF --------------------------
def extrair_dados_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        lines = [line for page in pdf.pages for line in page.extract_text().split("\n")]

    client_name = pedido = None
    pedido_values = set()

    # Extrai cliente e pedido
    for line in lines:
        if match := re.search(r'Cliente:\s*(.+?)(?:\s*\(\d+\)|$)', line):
            client_name = match.group(1).strip()
        if match := re.search(r'(?i)pedido\s*nº:\s*(\d+)', line):
            pedido = match.group(1)

        # Fallback: pega número da primeira coluna (romaneio/orçamento)
        elif match := re.search(r'^\d+\s+\d+\s+(\d+)', line):
            pedido_values.add(match.group(1))
        elif match := re.search(r'^\d+\s+(\d+)', line):
            pedido_values.add(match.group(1))

    if not pedido and pedido_values:
        pedido = sorted(pedido_values)[0]

    # Extrai linhas de produtos
    data = []
    for line in lines:
        if match := re.search(r'^\d+\s+\d+\s+(\d+)\s+(.*?)\s+(\d+[\d,]*)$', line):
            produto, descricao, qtd = match.groups()
            qtd = float(qtd.replace(",", "."))
            data.append({"Produto": produto, "Descrição": descricao.strip(), "Qtd.": qtd})

    ordem_prod = pd.DataFrame(data)
    ordem_prod["Produto"] = ordem_prod["Produto"].astype(str).str.zfill(8)  # Garante 8 dígitos com zeros
    return ordem_prod, client_name or "CLIENTE NÃO IDENTIFICADO", pedido or "SEM PEDIDO"

# -------------------------- CARREGA BASE DE EMBALAGENS --------------------------
def carregar_base_embalagens():
    if not BASE_QUANTIDADES_FILE.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {BASE_QUANTIDADES_FILE.name}")
    
    ordem_prod = pd.read_excel(BASE_QUANTIDADES_FILE, dtype={"Produto": str})
    ordem_prod["Produto"] = ordem_prod["Produto"].str.zfill(8)
    return ordem_prod.set_index("Produto")["Qtd.Embalagem"].to_dict()

# -------------------------- GERA PACOTES --------------------------
def gerar_pacotes(df_produtos, capacidade_dict, client_name, pedido):
    pacotes = []
    for _, row in df_produtos.iterrows():
        codigo = row["Produto"]
        descricao = row["Descrição"]
        qtd_total = int(row["Qtd."])
        capacidade = int(capacidade_dict.get(codigo, DEFAULT_BOX_CAPACITY))

        caixas_cheias = qtd_total // capacidade
        resto = qtd_total % capacidade
        total_caixas = caixas_cheias + (1 if resto > 0 else 0)

        for i in range(1, total_caixas + 1):
            qtd_na_caixa = capacidade if i <= caixas_cheias else resto
            pacotes.append({
                "Cliente": client_name,
                "Pedido": pedido,
                "Produto": codigo,
                "Descrição": descricao,
                "Caixa": f"{i}/{total_caixas}",
                "Qtd. na Caixa": qtd_na_caixa,
                "Qtd. Total": qtd_total,
                "Capacidade": capacidade
            })
    return pd.DataFrame(pacotes)

# -------------------------- SALVA EXCEL FORMATADO --------------------------
def salvar_excel_formatado(ordem_prod, pedido):
    hora_minuto = datetime.now().strftime("%H%M")
    arquivo_saida = Path(f"Etiquetas Pedido {pedido} Data {hora_minuto}.xlsx")

    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        ordem_prod.to_excel(writer, index=False, sheet_name="Pacotes")
        ws = writer.sheets["Pacotes"]

        # Auto-ajustar largura
        for column in ws.columns:
            max_length = max(len(str(cell.value)) for cell in column if cell.value)
            ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 50)

        # Congelar e negrito cabeçalho
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

    print(f"\nArquivo gerado com sucesso!")
    print(f"   → {len(ordem_prod)} caixas")
    print(f"   → {arquivo_saida.name}")
    print(f"   → {arquivo_saida.resolve()}\n")

# ============================== EXECUÇÃO ==============================
def main():
    df_produtos, cliente, pedido = extrair_dados_pdf(input_pdf)
    capacidade_dict = carregar_base_embalagens()
    df_pacotes = gerar_pacotes(df_produtos, capacidade_dict, cliente, pedido)
    salvar_excel_formatado(df_pacotes, pedido)

if __name__ == "__main__":
    main()