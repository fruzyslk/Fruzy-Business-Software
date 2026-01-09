# purchase_entry.py
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd


class PurchaseEntryTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self._rate_trace_id = None
        self.create_widgets()

    def create_widgets(self):
        self._create_form()
        self._create_list()

    def _create_form(self):
        form_frame = ctk.CTkFrame(self.parent)
        form_frame.pack(fill='x', padx=15, pady=15)
        ctk.CTkLabel(
            form_frame,
            text="Add Purchase Entry",
            font=('Arial', 14, 'bold'),
            text_color=self.app.colors.get('text_dark', 'black')
        ).pack(anchor='w', padx=10, pady=(10, 5))

        # Row 1: Item, Quantity, Unit
        row1 = ctk.CTkFrame(form_frame, fg_color="transparent")
        row1.pack(fill='x', padx=10, pady=5)
        ctk.CTkLabel(row1, text="Item Name:", width=90, anchor='w').pack(side='left', padx=5)
        self.app.purchase_veg_var = tk.StringVar()
        self.app.purchase_veg_var.trace('w', self._on_vegetable_selected)
        ctk.CTkEntry(row1, textvariable=self.app.purchase_veg_var, width=220).pack(side='left', padx=5)
        ctk.CTkLabel(row1, text="Quantity:", width=80, anchor='w').pack(side='left', padx=10)
        self.app.purchase_qty_var = tk.StringVar()
        ctk.CTkEntry(row1, textvariable=self.app.purchase_qty_var, width=80).pack(side='left', padx=5)
        ctk.CTkLabel(row1, text="Unit:", width=60, anchor='w').pack(side='left', padx=10)
        self.app.purchase_unit_var = tk.StringVar(value='kg')
        self.unit_combo = ttk.Combobox(
            row1,
            textvariable=self.app.purchase_unit_var,
            values=['kg', 'piece', 'dozen', 'bundle'],
            font=('Arial', 10),
            width=12,
            state='readonly'
        )
        self.unit_combo.pack(side='left', padx=5)
        self.unit_label = ctk.CTkLabel(row1, text="", font=('Arial', 9), text_color='gray')
        self.unit_label.pack(side='left', padx=5)

        # Row 2: Rate, Vendor, Payment
        row2 = ctk.CTkFrame(form_frame, fg_color="transparent")
        row2.pack(fill='x', padx=10, pady=5)
        ctk.CTkLabel(row2, text="Rate (PKR):", width=90, anchor='w').pack(side='left', padx=5)
        self.app.purchase_rate_var = tk.StringVar()
        self.app.purchase_rate_var.trace('w', self.app.calculate_purchase_total)
        ctk.CTkEntry(row2, textvariable=self.app.purchase_rate_var, width=120).pack(side='left', padx=5)
        ctk.CTkLabel(row2, text="Vendor:", width=80, anchor='w').pack(side='left', padx=10)
        self.app.purchase_vendor_var = tk.StringVar(value='Main Vendor')
        ctk.CTkEntry(row2, textvariable=self.app.purchase_vendor_var, width=160).pack(side='left', padx=5)
        ctk.CTkLabel(row2, text="Payment:", width=80, anchor='w').pack(side='left', padx=10)
        self.app.purchase_payment_var = tk.StringVar(value='cash')
        payment_btn = ctk.CTkSegmentedButton(
            row2,
            values=["Cash", "Credit"],
            command=lambda v: self.app.purchase_payment_var.set("cash" if v == "Cash" else "credit"),
            width=180
        )
        payment_btn.set("Cash")
        payment_btn.pack(side='left', padx=5)

        # Row 2.5: Import Rates Button
        row2_5 = ctk.CTkFrame(form_frame, fg_color="transparent")
        row2_5.pack(fill='x', padx=10, pady=5)
        ctk.CTkButton(
            row2_5,
            text="📥 Import Purchase Rates from Excel",
            command=self.import_purchase_rates,
            width=220,
            font=('Arial', 12)
        ).pack(side='left', padx=5)

        # Row 3: Total + Add Button
        row3 = ctk.CTkFrame(form_frame, fg_color="transparent")
        row3.pack(fill='x', padx=10, pady=10)
        ctk.CTkLabel(
            row3, text="Total Amount:", font=('Arial', 12, 'bold'), width=100, anchor='w'
        ).pack(side='left', padx=5)
        self.app.purchase_total_var = tk.StringVar(value='0.00')
        ctk.CTkLabel(
            row3,
            textvariable=self.app.purchase_total_var,
            font=('Arial', 16, 'bold'),
            text_color=self.app.colors.get('primary', 'green')
        ).pack(side='left', padx=5)
        ctk.CTkButton(
            row3,
            text="➕ Add Purchase",
            command=self._validate_and_add_purchase,
            font=('Arial', 14, 'bold'),
            width=150
        ).pack(side='right', padx=10)

    def _validate_and_add_purchase(self):
        """Validate inputs and delegate to app.add_purchase with guaranteed string quantity."""
        veg = self.app.purchase_veg_var.get().strip()
        qty_str = self.app.purchase_qty_var.get().strip()
        unit = self.app.purchase_unit_var.get()
        rate_str = self.app.purchase_rate_var.get().strip()
        vendor = self.app.purchase_vendor_var.get().strip()
        payment = self.app.purchase_payment_var.get()

        if not veg:
            messagebox.showwarning("Missing Item", "Please enter the vegetable/item name.")
            return
        if not qty_str:
            messagebox.showwarning("Missing Quantity", "Please enter a quantity.")
            return
        if not rate_str:
            messagebox.showwarning("Missing Rate", "Please enter a rate.")
            return

        try:
            qty_val = float(qty_str)
            rate_val = float(rate_str)
        except ValueError:
            messagebox.showerror("Invalid Input", "Quantity and rate must be valid numbers.")
            return

        if qty_val <= 0:
            messagebox.showwarning("Invalid Quantity", "Quantity must be greater than zero.")
            return
        if rate_val < 0:
            messagebox.showwarning("Invalid Rate", "Rate cannot be negative.")
            return

        # ✅ CRITICAL: Format quantity as "value unit" string with 2 decimal places
        formatted_quantity = f"{qty_val:.2f} {unit}"

        # Call the main app's add_purchase with guaranteed correct format
        success = self.app.add_purchase(
            vegetable=veg,
            quantity=formatted_quantity,
            rate=rate_val,
            vendor=vendor,
            payment=payment
        )

        if success:
            # Clear form (except vendor, if desired)
            self.app.purchase_veg_var.set("")
            self.app.purchase_qty_var.set("")
            self.app.purchase_rate_var.set("")
            self.app.purchase_total_var.set("0.00")
            # Optionally keep vendor: self.app.purchase_vendor_var stays

    def _create_list(self):
        list_frame = ctk.CTkFrame(self.parent)
        list_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        ctk.CTkLabel(
            list_frame,
            text="Today's Purchases",
            font=('Arial', 14, 'bold'),
            text_color=self.app.colors.get('text_dark', 'black')
        ).pack(anchor='w', padx=10, pady=(10, 5))

        from utils import make_treeview
        self.app.purchase_tree = make_treeview(
            list_frame,
            columns=('Vegetable', 'Quantity', 'Rate', 'Total', 'Vendor', 'Payment'),
            headings=('Item Name', 'Quantity', 'Rate (PKR)', 'Total (PKR)', 'Vendor', 'Payment Type'),
            widths=(200, 120, 100, 120, 150, 100),
            height=12
        )
        self.app.purchase_tree.pack(fill='both', expand=True, padx=10, pady=10)

        self.app.purchase_tree.bind('<Delete>', lambda e: self.app.delete_purchase())
        self.app.purchase_tree.bind('<BackSpace>', lambda e: self.app.delete_purchase())

        self.reload_purchase_list()

        ctk.CTkButton(
            list_frame,
            text="🗑️ Delete Selected",
            command=self.app.delete_purchase,
            width=150,
            font=('Arial', 12, 'bold')
        ).pack(pady=10)

    def import_purchase_rates(self):
        file_path = filedialog.askopenfilename(
            title="Select Purchase Rate Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path, usecols=[0, 1], header=None, dtype=str)
            if df.shape[1] < 2:
                raise ValueError("Excel file must have at least two columns: Item Name and Rate.")

            df.columns = ['Item', 'Rate']
            df = df.dropna()

            rate_dict = {}
            for _, row in df.iterrows():
                item = str(row['Item']).strip()
                rate_val = row['Rate']
                if item and pd.notna(rate_val):
                    try:
                        rate_clean = str(rate_val).replace(',', '').strip()
                        rate_num = float(rate_clean)
                        rate_dict[item] = rate_num
                    except (ValueError, TypeError):
                        continue

            self.app.imported_purchase_rates = rate_dict
            self._setup_rate_autofill()
            messagebox.showinfo("Success", f"Successfully imported {len(rate_dict)} purchase rates!")

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import purchase rates:\n{str(e)}")

    def _setup_rate_autofill(self):
        if self._rate_trace_id:
            self.app.purchase_veg_var.trace_vdelete('w', self._rate_trace_id)
        self._rate_trace_id = self.app.purchase_veg_var.trace('w', self._on_item_name_change)

    def _on_item_name_change(self, *args):
        item = self.app.purchase_veg_var.get().strip()
        if hasattr(self.app, 'imported_purchase_rates') and item:
            rate = self.app.imported_purchase_rates.get(item)
            if rate is not None:
                self.app.purchase_rate_var.set(str(rate))

    def _on_vegetable_selected(self, *args):
        veg_name = self.app.purchase_veg_var.get().strip()
        if not veg_name:
            self.unit_combo.configure(state='readonly')
            self.unit_label.configure(text="")
            return
        
        invoice_unit = self._get_invoice_unit(veg_name)
        
        if invoice_unit:
            self.app.purchase_unit_var.set(invoice_unit)
            self.unit_combo.configure(state='disabled')
            self.unit_label.configure(text="(from invoices)", text_color='green')
        else:
            self.unit_combo.configure(state='readonly')
            self.unit_label.configure(text="")
    
    def _get_invoice_unit(self, veg_name):
        import re
        
        def get_english_name(veg_str):
            if not veg_str:
                return veg_str
            veg_str = str(veg_str).strip()
            matches = list(re.finditer(r'\([^()]*(?:\([^()]*\))?[^()]*\)', veg_str))
            for match in reversed(matches):
                content = match.group(0)[1:-1].strip()
                if content.lower() in ['big size', 'small size', 'large', 'small']:
                    continue
                if any(c.isalpha() for c in content):
                    content = re.sub(r'\s*\(\s*(big|small)\s*size\s*\)', '', content, flags=re.IGNORECASE).strip()
                    content = re.sub(r'\s*\(\s*\)', '', content).strip()
                    return content
            return veg_str
        
        if hasattr(self.app, 'sales') and self.app.sales:
            target_english = get_english_name(veg_name)
            
            for date_items in self.app.sales.values():
                for item in date_items:
                    if 'vegetable' in item:
                        invoice_veg = item['vegetable']
                        invoice_english = get_english_name(invoice_veg)
                        
                        if invoice_english.lower() == target_english.lower():
                            qty_str = item.get('quantity', '')
                            if isinstance(qty_str, str) and ' ' in qty_str:
                                unit = qty_str.split()[-1]
                                return unit
        
        return None

    def reload_purchase_list(self):
        try:
            if not (hasattr(self.app, 'purchase_tree') and self.app.purchase_tree):
                return
            tree = self.app.purchase_tree
            for iid in tree.get_children():
                tree.delete(iid)
            for p in getattr(self.app, 'purchases', []):
                tree.insert('', 'end', values=(
                    p.get('vegetable', ''),
                    p.get('quantity', ''),
                    p.get('rate', ''),
                    p.get('total', ''),
                    p.get('vendor', ''),
                    p.get('payment', '')
                ))
        except Exception as e:
            print(f"Error reloading purchase list: {e}")