from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# === Page and Label Specs ===
page_width = 100 * mm
page_height = 150 * mm
label_width = 50 * mm
label_height = 15 * mm
labels_per_row = 2
rows_per_page = int(page_height // label_height)

# === Sample Data ===
products = [
    ("Product A", 10),
    ("Product B", 4),
    ("Product C", 25),
]

def generate_labels_pdf(filename="labels.pdf"):
    c = canvas.Canvas(filename, pagesize=(page_width, page_height))
    
    x_positions = [0, label_width]  # left column, right column
    y = page_height - label_height  # start from top
    
    col = 0  # current column index
    
    for product, qty in products:
        for i in range(qty):
            x = x_positions[col]

            # Draw border (for testing alignment – remove when done)
            c.rect(x, y, label_width, label_height)

            # Center text
            c.setFont("Helvetica", 8)
            text_x = x + label_width / 2
            text_y = y + label_height / 2
            c.drawCentredString(text_x, text_y, product)

            # Next column/row
            col += 1
            if col >= labels_per_row:  # finished row
                col = 0
                y -= label_height
                if y < 0:  # new page
                    c.showPage()
                    y = page_height - label_height

    c.save()
    print(f"✅ Labels PDF generated: {filename}")

if __name__ == "__main__":
    generate_labels_pdf("labels_test.pdf")