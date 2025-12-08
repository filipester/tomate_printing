import pdfplumber
import re
import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows

script_dir = Path(__file__).parent

pdf_files = list(script_dir.glob("*.pdf"))

if not pdf_files:
    raise FileNotFoundError("No PDF file found in the application folder.")


BASE_FILE = Path("base_quantities.xlsx")

if not BASE_FILE.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {BASE_FILE.name}")

try:
    base_df = pd.read_excel(BASE_FILE, dtype={"Produto": str})
except Exception as e:
    raise ValueError(f"Erro ao ler {BASE_FILE.name}: {e}")

base_df["Produto"] = base_df["Produto"].str.zfill(8)
# Verifica se as colunas necessárias existem
required_base_cols = ["Produto", "Qtd.Embalagem"]
missing_cols = [col for col in required_base_cols if col not in base_df.columns]
if missing_cols:
    raise ValueError(f"Colunas faltando em base_quantities.xlsx: {missing_cols}")

# Use the first one found, or loop if you expect more than one
input_file = pdf_files[0]
romaneio = input_file.stem  # e.g., "3647" without extension


input_file = script_dir / f"{romaneio}.pdf"
output_dir = script_dir / "OPs"

# Parse description
def parse_description(desc):
    pattern = r'(M\s?\d+|\d+/\d+"?)\s*[xX]\s*(\d+)\s*[xX]\s*(\d+)\s*([A-Z]{1,2})'
    try:
        if pd.isna(desc) or not isinstance(desc, str) or not desc.strip():
            print(f"Invalid description: '{desc}'")
            return pd.Series([None, None, None, None])
        match = re.search(pattern, desc)
        if match:
            bitola = match.group(1)
            abertura = match.group(2)
            comp = match.group(3)
            mod = match.group(4)
            return pd.Series([bitola, abertura, comp, mod])
        else:
            print(f"Failed to parse description: '{desc}'")
            return pd.Series([None, None, None, None])
    except Exception as e:
        print(f"Error parsing description '{desc}': {e}")
        return pd.Series([None, None, None, None])

# Define regex patterns
pad_romaneio = r'^\d+\s+\d+\s+(\d+)\s+(.*?)\s+(\d{1,3}(?:,\d{1,4})?)$'
pad_orcamento = r'^\d+\s+(\d+)\s+(GRAMPO.*?)\s+PC\s+(\d{1,3}(?:,\d{1,4})?)(?:\s+.*)?$'

# Extraindo o PDF
with pdfplumber.open(input_file) as pdf:
    all_text = ""
    lines_by_page = []
    for page_num, page in enumerate(pdf.pages, 1):
        page_text = page.extract_text()
        if page_text:
            page_lines = page_text.split('\n')
            lines_by_page.append((page_num, page_lines))
            all_text += page_text + "\n"

lines = all_text.split('\n')

# Extract client info and pedido
client_name = None
pedido = None
pedido_values = set()

# First pass: extract client info and pedido
for line in lines:
    if match := re.search(r'Cliente:\s*(.+?)(?:\s*\(\d+\)|$)', line):
        client_name = match.group(1).strip()
    if match := re.search(r'(?i)pedido\s*nº:\s*(\d+)\s*data:', line, re.IGNORECASE):
        print(f"Matched pedido line: '{line}' -> Pedido = {match.group(1)}")
        pedido = match.group(1)  # Explicit Pedido number
    elif match := re.search(pad_romaneio, line):
        pedido_values.add(line.split()[0])  # Pedido from first column
    elif match := re.search(pad_orcamento, line):
        pedido_values.add(line.split()[0])  # Item from first column

# Warn about multiple Pedido/Item values, but don't overwrite pedido
if len(pedido_values) > 1:
    print(f"Warning: Multiple Pedido/Item values found in table rows: {pedido_values}.")
if not pedido:
    pedido = list(pedido_values)[0] if pedido_values else "Unknown"
    print(f"No explicit Pedido found; using {pedido} from table rows.")


data = []

# Second pass: extract table data
for page_num, page_lines in lines_by_page:
    for line in page_lines:
        if match := re.search(pad_romaneio, line):
            produto = match.group(1)  # Produto
            descricao = match.group(2).strip()  # Descrição
            quantidade = match.group(3).replace(',', '.')  # Qtd.
            data.append([produto, descricao, float(quantidade)])
        elif match := re.search(pad_orcamento, line):
            produto = match.group(1)  # Produto
            descricao = match.group(2).strip()  # Descrição
            quantidade = match.group(3).replace(',', '.')  # Qtd.
            data.append([produto, descricao, float(quantidade)])
        else:
            print(f"Page {page_num}: Line did not match any pattern: '{line}'")

# Create DataFrame
ordem_prod = pd.DataFrame(data, columns=["Produto", "Descrição", "Qtd."])
# print(ordem_prod.head(20))
# Cria dicionário: código → capacidade por embalagem (com fallback para 10)
base_df["Produto"] = base_df["Produto"].astype(str).str.strip()
capacidade_por_produto = base_df.set_index("Produto")["Qtd.Embalagem"].to_dict()

# print(capacidade_por_produto)

print(f"Base de embalagens carregada: {len(capacidade_por_produto)} produtos definidos.")
print("Produtos sem capacidade definida usarão 10 peças por caixa.\n")

# ============================== EXPAND TO ONE ROW PER BOX ==============================
print("Gerando linhas por caixa...")

pacotes_linhas = []

for _, row in ordem_prod.iterrows():
    codigo = row["Produto"]
    descricao = row["Descrição"]
    qtd_total = int(float(row["Qtd."]))

    # Get box capacity from the base file (fallback to 10)
    capacidade = int(float(capacidade_por_produto.get(codigo, 10)))

    # Calculate how many boxes we need
    caixas_cheias = qtd_total // capacidade
    resto = qtd_total % capacidade
    total_caixas = caixas_cheias + (1 if resto > 0 else 0)

    # Create one row for each box
    for i in range(1, total_caixas + 1):
        qtd_na_caixa = capacidade if i <= caixas_cheias else resto
        caixa_label = f"{i}/{total_caixas}"

        pacotes_linhas.append({
            "Cliente": client_name or "NÃO IDENTIFICADO",
            "Pedido": pedido or "NÃO IDENTIFICADO",
            "Produto": codigo,
            "Descrição": descricao,
            "Caixa": caixa_label,
            "Qtd. na Caixa": qtd_na_caixa,
            "Qtd. Total": qtd_total,
            "Capacidade": capacidade
        })

# Convert to final DataFrame
df_pacotes = pd.DataFrame(pacotes_linhas)

print(f"Expansão concluída: {len(df_pacotes)} caixas geradas a partir de {len(ordem_prod)} produtos.\n")


# ============================== SAVE FINAL EXCEL ==============================
print("Salvando arquivo Excel final...")

# Reorder columns exactly how you want them
colunas_finais = [
    "Cliente",
    "Pedido",
    "Produto",
    "Descrição",
    "Caixa",
    "Qtd. na Caixa",
]

df_final = df_pacotes[colunas_finais]

# Generate timestamped filename
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_file = Path(f"Etiquetas Pedido {pedido} Data {timestamp}.xlsx")

# Save with formatting
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_final.to_excel(writer, index=False, sheet_name="Pacotes")
    
    # Get the workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets["Pacotes"]
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Freeze header row and bold it
    worksheet.freeze_panes = 'A2'
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

print(f"\nPRONTO!")
print(f"   → {len(df_final)} caixas geradas")
print(f"   → Arquivo salvo: {output_file.name}")
print(f"   → Local: {output_file.resolve()}\n")