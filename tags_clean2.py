import pdfplumber
import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl.styles import Alignment, Font

# ============================================================
# INITIAL SETUP
# ============================================================

script_dir = Path(__file__).parent
pdf_files = list(script_dir.glob("*.pdf"))

if not pdf_files:
    raise FileNotFoundError("Nenhum arquivo PDF encontrado na pasta do aplicativo.")

input_file = pdf_files[0]
romaneio = input_file.stem

BASE_FILE = Path("base_quantities.xlsx")
if not BASE_FILE.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {BASE_FILE.name}")

# ============================================================
# LOAD BASE QUANTITIES
# ============================================================

try:
    base_df = pd.read_excel(BASE_FILE, dtype={"Produto": str})
except Exception as e:
    raise ValueError(f"Erro ao ler {BASE_FILE.name}: {e}")

required_cols = ["Produto", "Qtd.Embalagem"]
missing_cols = [c for c in required_cols if c not in base_df.columns]
if missing_cols:
    raise ValueError(f"Colunas faltando em base_quantities.xlsx: {missing_cols}")

base_df["Produto"] = base_df["Produto"].str.zfill(8).str.strip()
capacidade_por_produto = base_df.set_index("Produto")["Qtd.Embalagem"].to_dict()

print(f"Base carregada: {len(capacidade_por_produto)} produtos definidos.")
print("Produtos sem capacidade definida usarão 10 peças por caixa.\n")

# ============================================================
# READ PDF
# ============================================================

all_text = ""
lines_by_page = []

with pdfplumber.open(input_file) as pdf:
    for page_num, page in enumerate(pdf.pages, 1):
        page_text = page.extract_text()
        if page_text:
            page_lines = page_text.split("\n")
            lines_by_page.append((page_num, page_lines))
            all_text += page_text + "\n"

lines = all_text.split("\n")

# ============================================================
# REGEX PATTERNS
# ============================================================

pad_romaneio = r'^\d+\s+\d+\s+(\d+)\s+(.*?)\s+(\d{1,3}(?:,\d{1,4})?)$'
pad_pedido = r'^\s*\d+\s+(\d+)\s+(.+?)\s+(?:PC|UN|CT|JG|KG|LT|PAR|MT)\s+([\d.,]+)'

# ============================================================
# EXTRACT CLIENT AND PEDIDO
# ============================================================

client_name = None
pedido = None
pedido_values = set()

for line in lines:

    # Cliente
    if m := re.search(r'Cliente:\s*(.+?)(?:\s*\(\d+\)|$)', line):
        client_name = m.group(1).strip()

    # Pedido explícito
    if m := re.search(r'pedido\s*nº:\s*(\d+)\s*data:', line, re.IGNORECASE):
        pedido = m.group(1)

    # Pedido implícito via tabela
    elif re.search(pad_romaneio, line) or re.search(pad_pedido, line):
        pedido_values.add(line.split()[0])

# Fallback
if not pedido:
    pedido = next(iter(pedido_values), "Unknown")
    print(f"Pedido não encontrado explicitamente, usando: {pedido}")

# ============================================================
# EXTRACT TABLE DATA
# ============================================================

data = []

for page_num, page_lines in lines_by_page:
    for line in page_lines:

        if m := re.search(pad_romaneio, line):
            produto, descricao, qtd = m.groups()

        elif m := re.search(pad_pedido, line):
            produto, descricao, qtd = m.groups()

        else:
            continue

        qtd_float = float(qtd.replace(",", "."))
        data.append([produto, descricao.strip(), qtd_float])

ordem_prod = pd.DataFrame(data, columns=["Produto", "Descrição", "Qtd."])

# ============================================================
# EXPAND INTO BOX ROWS
# ============================================================

print("Gerando linhas por caixa...\n")

pacotes = []

for _, row in ordem_prod.iterrows():

    codigo = row["Produto"]
    descricao = row["Descrição"]
    qtd_total = int(row["Qtd."])

    capacidade = int(capacidade_por_produto.get(codigo, 10))
    caixas_cheias = qtd_total // capacidade
    resto = qtd_total % capacidade
    total_caixas = caixas_cheias + (resto > 0)

    for i in range(1, total_caixas + 1):
        qtd_caixa = capacidade if i <= caixas_cheias else resto
        caixa_label = f"{i}/{total_caixas}"

        pacotes.append({
            "Cliente": client_name or "NÃO IDENTIFICADO",
            "Pedido": pedido,
            "Produto": codigo,
            "Descrição": descricao,
            "Caixa": caixa_label,
            "Qtd. na Caixa": qtd_caixa,
        })

df_final = pd.DataFrame(pacotes)

print(f"Concluído: {len(df_final)} caixas geradas.\n")

# ============================================================
# SAVE EXCEL OUTPUT
# ============================================================

timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_file = Path(f"Etiquetas Pedido {pedido} Data {timestamp}.xlsx")

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_final.to_excel(writer, index=False, sheet_name="Pacotes")

    ws = writer.sheets["Pacotes"]

    # Ajuste de largura
    for col in ws.columns:
        length = max(len(str(c.value)) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(length + 2, 50)

    # Congelar cabeçalho + bold
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

print("PRONTO!")
print(f"Arquivo gerado: {output_file.name}")
print(f"Local: {output_file.resolve()}\n")
