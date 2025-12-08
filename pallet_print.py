from reportlab.lib.pagesizes import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import pandas as pd
from pathlib import Path
import textwrap
from datetime import datetime

def generate_shipping_labels_from_excel(excel_file, output_file=None):
    # Load Excel file
    df = pd.read_excel(excel_file)

    # Check required columns
    required_cols = ["Cliente", "Rua", "Bairro", "Cidade", "NF", "Transportadora"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing column in Excel: {col}")

    # Convert dataframe to list of dicts
    labels = df.to_dict(orient="records")

    # Default output filename with timestamp
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"etiquetas_{timestamp}.pdf"

    # Setup PDF
    c = canvas.Canvas(output_file, pagesize=landscape((100*mm, 150*mm)))

    for label in labels:
        # Draw border
        c.setLineWidth(3)
        c.rect(2*mm, 2*mm, 140*mm, 90*mm)

        y = 80*mm
        c.setFont("Helvetica-Bold", 14)

        # Wrap Cliente name
        cliente_wrapped = textwrap.wrap(str(label["Cliente"]).upper(), width=35)

        for i, line in enumerate(cliente_wrapped):
            if i == 0:
                c.drawString(10*mm, y, f"Cliente: {line}")
            else:
                c.drawString(10*mm, y, line)
            y -= 5*mm
        y -= 5*mm
        c.setFont("Helvetica", 12)
        rua_wrapped = textwrap.wrap(str(label["Rua"]), width=45)

        for i, line in enumerate(rua_wrapped):
            if i == 0:
                c.drawString(10*mm, y, f"Rua: {line}")
            else:
                c.drawString(10*mm, y, line)
            y -= 5*mm
        y -= 5*mm
        c.drawString(10*mm, y, f"Bairro: {label['Bairro']}")
        y -= 10*mm
        c.drawString(10*mm, y, f"Cidade: {label['Cidade']}")
        y -= 10*mm
        c.drawString(10*mm, y, f"NF: {label['NF']}")
        y -= 10*mm
        c.drawString(10*mm, y, f"Transp: {label['Transportadora']}")

        # New page for next label
        c.showPage()

    c.save()
    print(f"PDF gerado com sucesso: {Path(output_file).resolve()}")

generate_shipping_labels_from_excel("Imprimir Etiqueta de Pallet.xlsx")