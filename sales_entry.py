# sales_entry.py
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox

class SalesEntryTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.create_widgets()

    def create_widgets(self):
        # Info banner (styled like purchase tab)
        info_frame = ctk.CTkFrame(self.parent, fg_color=self.app.colors['secondary'], corner_radius=8)
        info_frame.pack(fill='x', padx=15, pady=(15, 10))
        ctk.CTkLabel(
            info_frame,
            text="ℹ️ Sales from invoices are automatically added here. You can also add manual entries.",
            font=('Arial', 10),
            text_color='black',
            wraplength=900
        ).pack(pady=8, padx=10)

        # Form frame
        form_frame = ctk.CTkFrame(self.parent)
        form_frame.pack(fill='x', padx=15, pady=10)
        ctk.CTkLabel(
            form_frame,
            text="Add Manual Sales Entry",
            font=('Arial', 14, 'bold'),
            text_color=self.app.colors.get('text_dark', 'black')
        ).pack(anchor='w', padx=10, pady=(10, 5))

        # Row 1: Item, Quantity, Unit
        row1 = ctk.CTkFrame(form_frame, fg_color="transparent")
        row1.pack(fill='x', padx=10, pady=5)
        ctk.CTkLabel(row1, text="Item Name:", width=90, anchor='w').pack(side='left', padx=5)
        self.app.sales_veg_var = ctk.StringVar()
        ctk.CTkEntry(row1, textvariable=self.app.sales_veg_var, width=220).pack(side='left', padx=5)
        ctk.CTkLabel(row1, text="Quantity:", width=80, anchor='w').pack(side='left', padx=10)
        self.app.sales_qty_var = ctk.StringVar()
        ctk.CTkEntry(row1, textvariable=self.app.sales_qty_var, width=80).pack(side='left', padx=5)
        ctk.CTkLabel(row1, text="Unit:", width=60, anchor='w').pack(side='left', padx=10)
        self.app.sales_unit_var = ctk.StringVar(value='kg')
        self.unit_combo = ttk.Combobox(
            row1,
            textvariable=self.app.sales_unit_var,
            values=['kg', 'piece', 'dozen', 'bundle'],
            font=('Arial', 10),
            width=12,
            state='readonly'
        )
        self.unit_combo.pack(side='left', padx=5)

        # Row 2: Rate, Total, Add Button
        row2 = ctk.CTkFrame(form_frame, fg_color="transparent")
        row2.pack(fill='x', padx=10, pady=5)
        ctk.CTkLabel(row2, text="Rate (PKR):", width=90, anchor='w').pack(side='left', padx=5)
        self.app.sales_rate_var = ctk.StringVar()
        self.app.sales_rate_var.trace('w', self.app.calculate_sales_total)
        ctk.CTkEntry(row2, textvariable=self.app.sales_rate_var, width=120).pack(side='left', padx=5)
        ctk.CTkLabel(row2, text="Total Amount:", font=('Arial', 12, 'bold'), width=100, anchor='w').pack(side='left', padx=20)
        self.app.sales_total_var = ctk.StringVar(value='0.00')
        ctk.CTkLabel(
            row2,
            textvariable=self.app.sales_total_var,
            font=('Arial', 16, 'bold'),
            text_color=self.app.colors.get('primary', 'green')
        ).pack(side='left', padx=5)
        ctk.CTkButton(
            row2,
            text="➕ Add Sale",
            command=self._validate_and_add_sale,
            font=('Arial', 14, 'bold'),
            width=150
        ).pack(side='right', padx=10)

        # Sales list
        list_frame = ctk.CTkFrame(self.parent)
        list_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        ctk.CTkLabel(
            list_frame,
            text="All Sales (Manual + From Invoices)",
            font=('Arial', 14, 'bold'),
            text_color=self.app.colors.get('text_dark', 'black')
        ).pack(anchor='w', padx=10, pady=(10, 5))

        from utils import make_treeview
        self.app.sales_tree = make_treeview(
            list_frame,
            columns=('Source', 'Vegetable', 'Quantity', 'Rate', 'Total'),
            headings=('Source', 'Item Name', 'Quantity', 'Rate (PKR)', 'Total (PKR)'),
            widths=(100, 220, 120, 120, 150),
            height=12
        )
        # 🔑 CRITICAL: Force extended selection & bind Cmd+A/Ctrl+A
        self.app.sales_tree.configure(selectmode='extended')
        self.app.sales_tree.bind('<Command-a>', lambda e: self._select_all_sales())
        self.app.sales_tree.bind('<Control-a>', lambda e: self._select_all_sales())

        self.app.sales_tree.pack(fill='both', expand=True, padx=10, pady=10)

        # Bind delete keys to app's delete method (supports multi-select)
        self.app.sales_tree.bind('<Delete>', lambda e: self.app.delete_sale())
        self.app.sales_tree.bind('<BackSpace>', lambda e: self.app.delete_sale())

        ctk.CTkButton(
            list_frame,
            text="🗑️ Delete Selected",
            command=self.app.delete_sale,
            width=150,
            font=('Arial', 12, 'bold')
        ).pack(pady=10)

    def _validate_and_add_sale(self):
        """Validate inputs and call app.add_sale with formatted quantity."""
        veg = self.app.sales_veg_var.get().strip()
        qty_str = self.app.sales_qty_var.get().strip()
        unit = self.app.sales_unit_var.get()
        rate_str = self.app.sales_rate_var.get().strip()

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

        # Format quantity as "value unit" string
        formatted_quantity = f"{qty_val:.2f} {unit}"

        # Call main app to add sale (assumes app handles data structure)
        success = self.app.add_sale()
        if success:
            self.app.sales_veg_var.set("")
            self.app.sales_qty_var.set("")
            self.app.sales_rate_var.set("")
            self.app.sales_total_var.set("0.00")

    def reload_sales_list(self):
        """Reload the sales list Treeview using index-based item IDs (iid = str(index))."""
        try:
            tree = self.app.sales_tree
            # Clear all existing items safely
            tree.delete(*tree.get_children())

            # Repopulate using current self.app.sales list
            sales_data = getattr(self.app, 'sales', [])
            for i, s in enumerate(sales_data):
                tree.insert('', 'end', iid=str(i), values=(
                    s.get('source', ''),
                    s.get('vegetable_display', ''),
                    s.get('quantity', ''),
                    s.get('rate', ''),
                    s.get('total', '')
                ))
        except Exception as e:
            print(f"Error reloading sales list: {e}")

    def _select_all_sales(self):
        """Select all items in the sales Treeview (for Cmd+A / Ctrl+A)."""
        try:
            children = self.app.sales_tree.get_children()
            if children:
                self.app.sales_tree.selection_set(children)
        except Exception as e:
            print(f"Error selecting all sales: {e}")