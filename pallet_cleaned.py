import sys
from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path
import textwrap
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(filename='labels.log', level=logging.DEBUG, format='%(message)s')

DEFAULT_CONFIG = {
    "page_width": 100 * mm,
    "page_height": 150 * mm,
    "border_margin": 2 * mm,
    "border_width": 140 * mm,
    "border_height": 90 * mm,
    "start_y": 80 * mm,
    "line_spacing": 5 * mm,
    "large_spacing": 10 * mm,
    "font_title": ("Helvetica-Bold", 14),
    "font_body": ("Helvetica", 12),
    "text_widths": {"Cliente": 35, "Rua": 45, "Bairro": 50, "Cidade": 50},
    "border_thickness": 3  # Added for explicit control
}

def draw_wrapped_text(canvas, text, x, y, prefix, width, font, font_size, max_lines=3):
    """Draw wrapped text on the canvas, returning the new y position."""
    canvas.setFont(font, font_size)
    wrapped = textwrap.wrap(str(text).upper(), width=width)
    if len(wrapped) > max_lines:
        raise ValueError(f"Text too long for {prefix}: {text}")
    for i, line in enumerate(wrapped):
        canvas.drawString(x, y - i * 5 * mm, f"{prefix}: {line}" if i == 0 else line)
    return y - len(wrapped) * 5 * mm

def load_excel(excel_file):
    excel_path = Path(excel_file)

    # ==================== CREATE TEMPLATE IF FILE DOESN'T EXIST ====================
    if not excel_path.exists():
        print(f"⚠️  File not found: {excel_path.name}")
        print("Creating a boilerplate template for you...\n")

        # Create template with correct columns and examples
        template = pd.DataFrame(columns=[
            "Cliente", 
            "Rua", 
            "Bairro", 
            "Cidade", 
            "NF", 
            "Transportadora"
        ])

        template.to_excel(excel_path, index=False)

        print(f"✅ Arquivo criado com sucesso: {excel_path.name}")
        print("   Preencha o arquivo e rode o programa novamente.")
        
        return None   # Retorna nada para dar saída

    # Load Excel file
    try:
        df = pd.read_excel(excel_file)
        if df.empty:
            raise ValueError("Arquivo excel está vazio.")
        if df is None:
            return None # encerra limpo
        return df
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo excel não encontrado: {excel_file}")
    except Exception as e:
        raise ValueError(f"Leitura do arquivo excel falhou: {str(e)}")


def validate_data(df):
    # Check required columns
    required_cols = ["Cliente", "Rua", "Bairro", "Cidade", "NF", "Transportadora"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Colunas faltando no arquivo: {col}")
        if df[col].isna().any():
            raise ValueError(f"Coluna {col} está vazia.")

    # Log dataset size
    logging.debug(f"Processando {len(df)} etiquetas")
    if len(df) > 1000:
        logging.warning(f"Arquivos grandes ({len(df)} linhas) podem aumentar o tempo de processamento.")

    # Convert dataframe to list of dicts
    labels = df.to_dict(orient="records")

    return labels

def generate_pdf(labels, output_path=None, config=DEFAULT_CONFIG):
    # Setup PDF

    c = canvas.Canvas(str(output_path), pagesize=landscape((config["page_width"], config["page_height"])))
    
    for i, label in enumerate(labels, 1):
        # Save canvas state
        c.saveState()
        
        # Set border thickness for each page
        c.setLineWidth(config["border_thickness"])
        logging.debug(f"Label {i}: Setting line width to {config['border_thickness']}")
        
        # Draw border
        c.rect(config["border_margin"], config["border_margin"], config["border_width"], config["border_height"])
        y = config["start_y"]
        # Draw wrapped text fields
        y = draw_wrapped_text(c, label["Cliente"], 10 * mm, y, "Cliente", config["text_widths"]["Cliente"], *config["font_title"])
        y = draw_wrapped_text(c, label["Rua"], 10 * mm, y - config["line_spacing"], "Rua", config["text_widths"]["Rua"], *config["font_body"])
        y -= config["line_spacing"]
        c.setFont(*config["font_body"])
        c.drawString(10 * mm, y, f"Bairro: {str(label['Bairro']).upper()}")
        y -= config["large_spacing"]
        c.drawString(10 * mm, y, f"Cidade: {str(label['Cidade']).upper()}")
        y -= config["large_spacing"]
        c.drawString(10 * mm, y, f"NF: {str(label['NF'])}")
        y -= config["large_spacing"]
        c.drawString(10 * mm, y, f"Transp: {str(label['Transportadora']).upper()}")
        
        # Restore canvas state
        c.restoreState()
        c.showPage()
    c.save()

def generate_shipping_labels_from_excel(excel_file, output_file=None, config=DEFAULT_CONFIG):
    # Default output filename
    output_path = Path(output_file if output_file else f"Pallet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

    try:
        df = load_excel(excel_file)
        if df is None:
            input("\nAperte Enter para sair...")
            sys.exit(0)
        labels = validate_data(df)
        generate_pdf(labels, output_path, config)
        print(f"PDF gerado com sucesso: {output_path.resolve()}")
        logging.debug(f"Generated PDF: {output_path.resolve()}")
    except Exception as e:
        logging.error(f"Failed to generate PDF: {str(e)}")
        raise

if __name__ == "__main__":
    generate_shipping_labels_from_excel("Imprimir Etiqueta de Pallet.xlsx")