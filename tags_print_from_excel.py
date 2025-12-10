from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path
import textwrap
from datetime import datetime
import logging
import glob
import os

# Configure logging
logging.basicConfig(filename='labels.log', level=logging.DEBUG, format='%(message)s')

DEFAULT_CONFIG = {
    "page_width": 100 * mm,
    "page_height": 150 * mm,
    "start_y": 80 * mm,
    "line_spacing": 5 * mm,
    "large_spacing": 10 * mm,
    "font_title": ("Helvetica-Bold", 14),
    "font_body": ("Helvetica", 12),
    "text_widths": {"Cliente": 35, "Descrição": 50},

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

def generate_shipping_labels_from_excel(excel_file, output_file=None, config=None):
    """
    Generate shipping labels from an Excel file as a PDF.

    Args:
        excel_file (str): Path to the Excel file with label data.
        output_file (str, optional): Output PDF file path. Defaults to timestamped filename.
        config (dict, optional): Configuration for page size, fonts, and layout.

    Raises:
        FileNotFoundError: If the Excel file is not found.
        ValueError: If required columns are missing or data is invalid.
    """
    # Use default config if none provided
    config = config or DEFAULT_CONFIG

    # Load Excel file
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo excel não encontrado: {excel_file}")
    except Exception as e:
        raise ValueError(f"Leitura do arquivo excel falhou: {str(e)}")

    if df.empty:
        raise ValueError("Arquivo excel está vazio.")

    # Check required columns
    required_cols = ["Cliente", "Descrição", "Produto", "Produto", "Caixa", "Qtd. na Caixa"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Colunas faltando no arquivo: {col}")
        if df[col].isna().any():
            raise ValueError(f"Coluna {col} está vazia.")

    # Log dataset size
    logging.debug(f"Processing {len(df)} labels from {excel_file}")
    if len(df) > 1000:
        logging.warning(f"Large dataset ({len(df)} rows) may increase processing time.")

    # Convert dataframe to list of dicts
    labels = df.to_dict(orient="records")

    # Default output filename
    output_path = Path(output_file if output_file else f"etiquetas_pacotes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

    # Setup PDF
    try:
        c = canvas.Canvas(str(output_path), pagesize=landscape((config["page_width"], config["page_height"])))
        
        for i, label in enumerate(labels, 1):
            # Save canvas state
            c.saveState()
            
            y = config["start_y"]
            # Draw wrapped text fields
            y = draw_wrapped_text(c, label["Cliente"], 10 * mm, y, "Cliente", config["text_widths"]["Cliente"], *config["font_title"])
            y = draw_wrapped_text(c, label["Descrição"], 10 * mm, y - config["line_spacing"], "Descrição", config["text_widths"]["Descrição"], *config["font_body"])
            y -= config["line_spacing"]
            c.setFont(*config["font_body"])
            c.drawString(10 * mm, y, f"Produto: 0{str(label['Produto']).upper()}")
            y -= config["large_spacing"]
            c.drawString(10 * mm, y, f"Pedido: {str(label['Pedido']).upper()}")
            # c.drawString(10 * mm, y, f"Nota Fiscal: 2.397") # temporario, alterar
            y -= config["large_spacing"]
            c.drawString(10 * mm, y, f"Pacote: {str(label['Caixa'])}")
            y -= config["large_spacing"]
            c.drawString(10 * mm, y, f"Qtd: {str(label['Qtd. na Caixa']).upper()}")
            
            # Restore canvas state
            c.restoreState()
            c.showPage()

        c.save()
        print(f"PDF gerado com sucesso: {output_path.resolve()}")
        logging.debug(f"Generated PDF: {output_path.resolve()}")
    except Exception as e:
        logging.error(f"Failed to generate PDF: {str(e)}")
        raise

if __name__ == "__main__":
    # Pattern: files starting with "Etiquetas Pedido" and ending with .xlsx
    files = glob.glob("Etiquetas Pedido*.xlsx")

    # Filter out temporary Excel files (starting with ~)
    files = [f for f in files if not os.path.basename(f).startswith("~")]

    if not files:
        print("No matching Excel file found! (Expected: Etiquetas Pedido*.xlsx)")
    else:
        xlsx_file = files[0]  # First match
        print(f"Using Excel file: {xlsx_file}")
        generate_shipping_labels_from_excel(xlsx_file)