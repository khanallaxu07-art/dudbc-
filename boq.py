import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.pdfgen import canvas

# ---------------- FILE PATH ----------------
def resource_path(filename):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, filename)

EXCEL_FILE = resource_path("norms.xlsx")
SHEET_NAME = "Rate Analysis"

if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

# ---------------- LOAD EXCEL ----------------
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=None, engine="openpyxl", dtype=str)
df = df.fillna("")

# ---------------- PARSE ITEMS ----------------
items = []
rows = len(df)
i = 0

while i < rows:
    col_a = df.iloc[i,0]

    if col_a.replace(".","",1).isdigit():

        code = col_a.strip()
        item_name = df.iloc[i,2].strip()

        resources = []
        bc_value = 1

        j = i + 1

        while j < rows:

            col_c = df.iloc[j,2]
            col_d = df.iloc[j,3]
            col_e = df.iloc[j,4]
            col_f = df.iloc[j,5]
            col_b = df.iloc[j,1]

            if col_b and not col_c:
                match = re.search(r"([\d\.]+)",col_b)
                if match:
                    bc_value=float(match.group(1))
                break

            if col_c and col_d and col_f:
                try:

                    qty=float(col_d)
                    rate=float(col_f)

                    resources.append({
                        "name":col_c.strip(),
                        "qty":qty,
                        "unit":col_e.strip(),
                        "rate":rate,
                        "vatable":False,
                        "vat":0,
                        "bc_value":bc_value
                    })

                except:
                    pass

            j+=1

        items.append({
            "code":code,
            "item":item_name,
            "resources":resources,
            "bc_value":bc_value
        })

        i=j

    else:
        i+=1


# ---------------- GUI ----------------
class DUDBCApp(tk.Tk):

    def __init__(self):

        super().__init__()

        self.title("DUDBC BOQ Resource Breakdown")
        self.geometry("1500x850")

        self.filtered_items = items
        self.current_breakdown = None
        self.boq_qty = 0

        # -------- Search --------
        tk.Label(self,text="Search Item").pack()

        self.search_var=tk.StringVar()
        self.search_entry=tk.Entry(self,textvariable=self.search_var,width=60)
        self.search_entry.pack()

        self.search_var.trace_add("write",lambda *args:self.update_search())

        # -------- Item List --------
        tk.Label(self,text="Select Items").pack()

        self.item_listbox=tk.Listbox(self,selectmode=tk.MULTIPLE,width=120,height=8)
        self.item_listbox.pack()

        self.refresh_items()

        # -------- BOQ Qty --------
        tk.Label(self,text="BOQ Quantity").pack()

        self.qty_entry=tk.Entry(self,width=20)
        self.qty_entry.pack()

        tk.Button(self,text="Calculator",command=self.calculator).pack()

        # -------- Buttons --------
        tk.Button(self,text="Calculate Breakdown",command=self.calculate).pack(pady=5)
        tk.Button(self,text="Clear Selection",command=self.clear_selection).pack(pady=5)
        tk.Button(self,text="Export Excel",command=self.export_excel).pack(pady=5)
        tk.Button(self,text="Export PDF",command=self.export_pdf).pack(pady=5)

        # -------- Selected Item Text --------
        self.items_desc_var=tk.StringVar(value="")
        tk.Label(self,textvariable=self.items_desc_var,font=("Arial",10,"bold"),fg="blue").pack()

        # -------- Table --------
        self.tree=ttk.Treeview(self,columns=("name","qty","unit","rate","vatable","vat","total"),show="headings")

        for col,w in zip(("name","qty","unit","rate","vatable","vat","total"),
                         (350,120,80,120,80,120,140)):
            self.tree.heading(col,text=col.upper())
            self.tree.column(col,width=w)

        self.tree.pack(fill="both",expand=True)

        self.tree.bind("<Double-1>",self.on_double_click)

        self.grand_total_var=tk.StringVar(value="Grand Total: 0")
        tk.Label(self,textvariable=self.grand_total_var,font=("Arial",12,"bold")).pack()

    # ---------------- SEARCH ----------------
    def update_search(self):

        text=self.search_var.get().upper()

        self.filtered_items=[i for i in items if text in i["item"].upper() or text in i["code"].upper()]

        self.refresh_items()

    def refresh_items(self):

        self.item_listbox.delete(0,tk.END)

        for i in self.filtered_items:
            self.item_listbox.insert(tk.END,f"{i['code']} - {i['item']}")

    # ---------------- CALCULATOR ----------------
    def calculator(self):

        expr=simpledialog.askstring("Calculator","Enter expression")

        if expr:
            try:
                result=eval(expr)
                self.qty_entry.delete(0,tk.END)
                self.qty_entry.insert(0,str(result))
            except:
                messagebox.showerror("Error","Invalid expression")

    # ---------------- CALCULATE ----------------
    def calculate(self):

        try:
            self.boq_qty=float(self.qty_entry.get())
        except:
            messagebox.showerror("Error","Enter BOQ quantity")
            return

        selected=self.item_listbox.curselection()

        if not selected:
            messagebox.showwarning("Select","Select item")
            return

        selected_items=[self.filtered_items[i] for i in selected]

        desc=" | ".join([f"{i+1}. {x['code']} - {x['item']}" for i,x in enumerate(selected_items)])

        self.items_desc_var.set("Selected: "+desc)

        groups={"Labour":[],"Materials":[],"Equipment":[]}

        labour_keywords=["LABOUR","HELPER","SKILLED"]
        equipment_keywords=["EXCAVATOR","ROLLER","GRADER","MIXER","GENERATOR"]

        for item in selected_items:

            for r in item["resources"]:

                name=r["name"].upper()

                if any(k in name for k in labour_keywords):
                    groups["Labour"].append(r)

                elif any(k in name for k in equipment_keywords):
                    groups["Equipment"].append(r)

                else:
                    groups["Materials"].append(r)

        self.current_breakdown=groups

        self.show_tree()

    # ---------------- SHOW TABLE ----------------
    def show_tree(self):

        self.tree.delete(*self.tree.get_children())

        grand_total=0

        for group,res_list in self.current_breakdown.items():

            if not res_list:
                continue

            self.tree.insert("",tk.END,values=(f"=== {group} ===","","","","","",""))

            group_total=0

            for r in res_list:

                adj_qty=(float(r["qty"])/float(r["bc_value"]))*self.boq_qty

                vat=r["rate"]*0.13 if r["vatable"] else 0

                total=(r["rate"]+vat)*adj_qty

                group_total+=total

                tick="☑" if r["vatable"] else "☐"

                self.tree.insert("",tk.END,
                                 values=(r["name"],f"{adj_qty:.3f}",r["unit"],
                                         r["rate"],tick,f"{vat:.2f}",f"{total:.2f}"))

            self.tree.insert("",tk.END,values=(f"Total {group}","","","","","",f"{group_total:.2f}"))

            grand_total+=group_total

        self.grand_total_var.set(f"Grand Total: {grand_total:.2f}")

    # ---------------- EDIT RATE / VAT ----------------
    def on_double_click(self,event):

        row=self.tree.identify_row(event.y)
        col=self.tree.identify_column(event.x)

        values=self.tree.item(row,"values")

        if not values or "===" in values[0]:
            return

        if col=="#4":

            x,y,w,h=self.tree.bbox(row,col)

            entry=tk.Entry(self.tree)
            entry.place(x=x,y=y,width=w,height=h)

            entry.insert(0,values[3])
            entry.focus()

            def save(e=None):

                try:
                    new=float(entry.get())
                    vals=list(values)
                    vals[3]=new
                    self.tree.item(row,values=vals)
                    entry.destroy()
                except:
                    messagebox.showerror("Error","Invalid number")

            entry.bind("<Return>",save)

        elif col=="#5":

            vals=list(values)

            if vals[4]=="☐":
                vals[4]="☑"
            else:
                vals[4]="☐"

            self.tree.item(row,values=vals)

    # ---------------- CLEAR ----------------
    def clear_selection(self):

        self.item_listbox.selection_clear(0,tk.END)
        self.tree.delete(*self.tree.get_children())
        self.items_desc_var.set("")
        self.current_breakdown=None
        self.grand_total_var.set("Grand Total: 0")

    # ---------------- EXPORT EXCEL ----------------
    def export_excel(self):

        if not self.current_breakdown:
            messagebox.showwarning("No Data","Calculate first")
            return

        path=filedialog.asksaveasfilename(defaultextension=".xlsx")

        if not path:
            return

        wb=Workbook()
        ws=wb.active

        for row in self.tree.get_children():
            ws.append(self.tree.item(row)["values"])

        wb.save(path)

        messagebox.showinfo("Saved","Excel exported")

    # ---------------- EXPORT PDF ----------------
    def export_pdf(self):

        if not self.current_breakdown:
            messagebox.showwarning("No Data","Calculate first")
            return

        path=filedialog.asksaveasfilename(defaultextension=".pdf")

        if not path:
            return

        c=canvas.Canvas(path)

        y=800

        for row in self.tree.get_children():

            vals=self.tree.item(row)["values"]

            text="  ".join([str(v) for v in vals])

            c.drawString(40,y,text)

            y-=20

        c.save()

        messagebox.showinfo("Saved","PDF exported")


# ---------------- RUN ----------------
if __name__=="__main__":

    app=DUDBCApp()
    app.mainloop()