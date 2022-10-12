from docxtpl import DocxTemplate

# Load template
doc = DocxTemplate("template.docx")

invoice_list = [[5, "packs of paper", 0.5, 1.5],
                    [1, "box of pens", 2.0, 2.0],
                    [2, "bottles of water", 1.0, 2.0]]

# Name automation
doc.render({"name": "John Doe",
            "phone": "555-555-5555",
            "invoice_list": invoice_list,
            "subtotal": 5.50,
            "salestax": "8.25%",
            "total": 6.0})
            

doc.save("new_invoice.docx")

