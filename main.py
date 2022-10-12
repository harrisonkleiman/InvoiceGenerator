# Creating interface with tkinter
import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

#----------------------Functions----------------------#

def clear_item():
    qty_spinbox.delete(0, tkinter.END) # Clearing the spinbox
    qty_spinbox.insert(0, "1") # Resetting the spinbox to 1
    desc_entry.delete(0, tkinter.END) 
    price_spinbox.delete(0, tkinter.END) 
    price_spinbox.insert(0, "0.0")
    
invoice_list = []
def add_item():
    qty = qty_spinbox.get()
    desc = desc_entry.get()
    price = price_spinbox.get()
    total = float(qty) * float(price)
    invoice_item = [qty, desc, price, total]
    tree.insert("", 0, values=invoice_item)
    clear_item()
    
    invoice_list.append(invoice_item)

def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    contact_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()

def generate_invoice():
     # Load template
    doc = DocxTemplate("template.docx")
    name = first_name_entry.get() + " " + last_name_entry.get()
    contact = contact_entry.get()
    subtotal = sum(float(item[3]) for item in invoice_list)
    salestax = subtotal * 0.0825
    total = subtotal + salestax

    # Render invoice document
    doc.render({"name": name,
                "contact": contact,
                "invoice_list": invoice_list,
                "subtotal": subtotal,
                "salestax": salestax,
                "total": total})
                
    # Save invoice document
    # %%Y-%m-%d_%H-%M-%S is the format for the date and time
    doc_name = "new_invoice_" + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".docx" 
    doc.save(doc_name)
    
    messagebox.showinfo("Invoice Generated", "Invoice generated successfully!")
    
    new_invoice()
    

screen = tkinter.Tk()
screen.title("Invoice Generator")

frame = ttk.Frame(screen)
frame.pack(padx=20, pady=10) # Packing the frame allows it to be displayed


# Buttons Labels and Entries
first_name_label = tkinter.Label(frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

contact_label = tkinter.Label(frame, text="Contact")
contact_label.grid(row=0, column=2)
contact_entry = tkinter.Entry(frame)
contact_entry.grid(row=1, column=2)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

desc_label = tkinter.Label(frame, text="Descr. of Transaction")
desc_label.grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

price_label = tkinter.Label(frame, text="Unit Price")
price_label.grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=1000, increment=0.5)
price_spinbox.grid(row=3, column=2)

add_item_button = tkinter.Button(frame, text = "Add item", command = add_item)
add_item_button.grid(row=4, column=2, pady=5)

columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")

    
tree.grid(row=5, column=0, columnspan=3, padx=30, pady=20)


save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)


screen.mainloop() # This is the main loop that keeps the window open


    
    
    