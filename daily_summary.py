# daily_summary.py
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
from utils import make_treeview

class DailySummaryTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.create_widgets()
        self.parent.after(100, self.refresh_all_data)

    def _parse_qty_and_unit(self, qty_str):
        """Parse '5.0 kg' → (5.0, 'kg'); default to (0.0, 'kg') on error"""
        if not qty_str or not isinstance(qty_str, str):
            return 0.0, 'kg'
        parts = qty_str.strip().split(maxsplit=1)
        try:
            qty = float(parts[0])
        except (ValueError, IndexError):
            qty = 0.0
        unit = parts[1] if len(parts) > 1 else 'kg'
        return qty, unit

    def create_widgets(self):
        cards_frame = ctk.CTkFrame(self.parent, fg_color="transparent")
        cards_frame.pack(fill='x', padx=15, pady=15)

        purchase_card = ctk.CTkFrame(cards_frame, corner_radius=10, border_width=2,
                                     border_color=self.app.colors.get('light', '#e0e0e0'))
        purchase_card.pack(side='left', fill='both', expand=True, padx=5)
        ctk.CTkLabel(purchase_card, text="Total Purchase", font=('Arial', 14, 'bold'),
                     text_color='#7f8c8d').pack(pady=(15, 5))
        self.app.total_purchase_label = ctk.CTkLabel(purchase_card, text="PKR 0.00",
                                                     font=('Arial', 24, 'bold'),
                                                     text_color=self.app.colors.get('red', 'red'))
        self.app.total_purchase_label.pack(pady=(0, 15))

        sales_card = ctk.CTkFrame(cards_frame, corner_radius=10, border_width=2,
                                  border_color=self.app.colors.get('light', '#e0e0e0'))
        sales_card.pack(side='left', fill='both', expand=True, padx=5)
        ctk.CTkLabel(sales_card, text="Total Sales", font=('Arial', 14, 'bold'),
                     text_color='#7f8c8d').pack(pady=(15, 5))
        self.app.total_sales_label = ctk.CTkLabel(sales_card, text="PKR 0.00",
                                                  font=('Arial', 24, 'bold'),
                                                  text_color=self.app.colors.get('primary', 'green'))
        self.app.total_sales_label.pack(pady=(0, 15))

        profit_card = ctk.CTkFrame(cards_frame, corner_radius=10, border_width=2,
                                   border_color=self.app.colors.get('light', '#e0e0e0'))
        profit_card.pack(side='left', fill='both', expand=True, padx=5)
        ctk.CTkLabel(profit_card, text="Profit/Loss", font=('Arial', 14, 'bold'),
                     text_color='#7f8c8d').pack(pady=(15, 5))
        self.app.profit_label = ctk.CTkLabel(profit_card, text="PKR 0.00",
                                             font=('Arial', 24, 'bold'),
                                             text_color=self.app.colors.get('primary', 'green'))
        self.app.profit_label.pack(pady=0)
        self.app.profit_percent_label = ctk.CTkLabel(profit_card, text="(0.00%)",
                                                     font=('Arial', 12),
                                                     text_color='#7f8c8d')
        self.app.profit_percent_label.pack(pady=(0, 15))

        bottom_frame = ctk.CTkFrame(self.parent, fg_color="transparent")
        bottom_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))

        left_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 7))

        profit_frame = ctk.CTkFrame(left_frame)
        profit_frame.pack(fill='x', pady=(0, 10))
        ctk.CTkLabel(profit_frame, text="🏆 Top 5 Most Profitable Items",
                     font=('Arial', 14, 'bold'),
                     text_color=self.app.colors.get('text_dark', 'black')).pack(pady=(10, 5), padx=10, anchor='w')
        self.app.profit_tree = make_treeview(
            profit_frame,
            columns=('Rank', 'Vegetable', 'Profit', 'Percent'),
            headings=('#', 'Item Name', 'Profit (PKR)', 'Profit %'),
            widths=(40, 200, 120, 100),
            height=6
        )
        for child in profit_frame.winfo_children():
            if isinstance(child, ttk.Treeview):
                child.pack(fill='x', padx=10, pady=(0, 10), expand=True)

        qty_frame = ctk.CTkFrame(left_frame)
        qty_frame.pack(fill='both', expand=True)
        
        qty_header_frame = ctk.CTkFrame(qty_frame, fg_color="transparent")
        qty_header_frame.pack(fill='x', padx=10, pady=(10, 5))
        ctk.CTkLabel(qty_header_frame, text="📦 Quantity Movement (Today)",
                     font=('Arial', 14, 'bold'),
                     text_color=self.app.colors.get('text_dark', 'black')).pack(side='left', anchor='w')
        ctk.CTkButton(qty_header_frame, text="📋 Copy Data", command=self._copy_qty_movement_data,
                      width=100, height=30, font=('Arial', 10)).pack(side='right', padx=5)
        
        self.app.qty_movement_tree = make_treeview(
            qty_frame,
            columns=('Item', 'Purchased', 'Sold', 'Remaining', 'Revenue'),
            headings=('Item Name', 'Purchased', 'Sold', 'Remaining', 'Revenue (PKR)'),
            widths=(220, 100, 80, 100, 120),
            height=10
        )
        try:
            tree = self.app.qty_movement_tree
            tree.bind("<Button-3>", self._on_qty_tree_right_click)
            tree.bind("<Button-2>", self._on_qty_tree_right_click)
            tree.bind("<Control-Button-1>", self._on_qty_tree_right_click)
        except Exception:
            pass

        right_frame = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        right_frame.pack(side='right', fill='both', expand=True, padx=(7, 0))

        purchase_summary = ctk.CTkFrame(right_frame)
        purchase_summary.pack(fill='x', pady=(0, 10))
        ctk.CTkLabel(purchase_summary, text="📦 Purchase Summary",
                     font=('Arial', 14, 'bold'),
                     text_color=self.app.colors.get('text_dark', 'black')).pack(pady=(10, 5), padx=15, anchor='w')
        self._create_summary_row(purchase_summary, "Total Items:", "purchase_items_label", is_currency=False)
        self._create_summary_row(purchase_summary, "Cash Purchases:", "cash_purchase_label")
        self._create_summary_row(purchase_summary, "Credit Purchases:", "credit_purchase_label")

        sales_summary = ctk.CTkFrame(right_frame)
        sales_summary.pack(fill='x')
        ctk.CTkLabel(sales_summary, text="💰 Sales Summary",
                     font=('Arial', 14, 'bold'),
                     text_color=self.app.colors.get('text_dark', 'black')).pack(pady=(10, 5), padx=15, anchor='w')
        self._create_summary_row(sales_summary, "Total Sales:", "sales_items_label", is_currency=False)
        self._create_summary_row(sales_summary, "From Invoices:", "invoice_sales_label", is_currency=False)
        self._create_summary_row(sales_summary, "Manual Entries:", "manual_sales_label", is_currency=False)
        self._create_summary_row(sales_summary, "Average Sale:", "avg_sale_label")

    def _create_summary_row(self, parent, label_text, attr_name, is_currency=True):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill='x', padx=15, pady=3)
        ctk.CTkLabel(row, text=label_text, font=('Arial', 12), width=140, anchor='w').pack(side='left')
        default_text = "PKR 0.00" if is_currency else "0"
        label_widget = ctk.CTkLabel(row, text=default_text, font=('Arial', 12, 'bold'), anchor='e')
        label_widget.pack(side='right', fill='x', expand=True)
        setattr(self.app, attr_name, label_widget)

    def _extract_english_name(self, veg_str):
        if not veg_str:
            return veg_str
        import re
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

    def update_qty_movement(self):
        if not hasattr(self.app, 'qty_movement_tree') or not self.app.qty_movement_tree:
            return
            
        for item in self.app.qty_movement_tree.get_children():
            self.app.qty_movement_tree.delete(item)

        # Build stats keyed by English name
        stats = {}
        for p in self.app.purchases:
            english = p['vegetable_english']
            if english not in stats:
                stats[english] = {
                    'display_name': p['vegetable_display'],
                    'purchased_qty': 0.0,
                    'sold_qty': 0.0,
                    'revenue': 0.0,
                    'unit': 'kg'
                }
            qty, unit = self._parse_qty_and_unit(p['quantity'])
            stats[english]['purchased_qty'] += qty
            stats[english]['unit'] = unit

        for s in self.app.sales:
            english = s['vegetable_english']
            if english not in stats:
                stats[english] = {
                    'display_name': s['vegetable_display'],
                    'purchased_qty': 0.0,
                    'sold_qty': 0.0,
                    'revenue': 0.0,
                    'unit': 'kg'
                }
            qty, unit = self._parse_qty_and_unit(s['quantity'])
            stats[english]['sold_qty'] += qty
            stats[english]['revenue'] += float(s.get('total', 0))
            stats[english]['unit'] = unit

        # Display in order of vegetable list
        for v in self.app.vegetables:
            english = v['english']
            if english in stats:
                data = stats[english]
                unit = data['unit']
                remaining = data['purchased_qty'] - data['sold_qty']
                try:
                    self.app.qty_movement_tree.insert('', 'end', values=(
                        data['display_name'],
                        f"{data['purchased_qty']:.2f} {unit}",
                        f"{data['sold_qty']:.2f} {unit}",
                        f"{remaining:.2f} {unit}",
                        f"{data['revenue']:,.2f}"
                    ))
                except Exception as e:
                    print(f"Error inserting qty movement row: {e}")

    def update_profit_items(self):
        if not hasattr(self.app, 'profit_tree') or not self.app.profit_tree:
            return
            
        for item in self.app.profit_tree.get_children():
            self.app.profit_tree.delete(item)

        profit_stats = {}
        for s in self.app.sales:
            english = s['vegetable_english']
            if english not in profit_stats:
                profit_stats[english] = {
                    'display_name': s['vegetable_display'],
                    'qty': 0,
                    'revenue': 0,
                    'cost': 0
                }
            qty, _ = self._parse_qty_and_unit(s['quantity'])
            revenue = float(s.get('total', 0))
            profit_stats[english]['qty'] += qty
            profit_stats[english]['revenue'] += revenue

        for p in self.app.purchases:
            english = p['vegetable_english']
            if english in profit_stats:
                cost = float(p.get('total', 0))
                profit_stats[english]['cost'] += cost

        profit_items = []
        for english, data in profit_stats.items():
            if data['revenue'] > 0:
                profit = data['revenue'] - data['cost']
                profit_pct = (profit / data['revenue'] * 100) if data['revenue'] > 0 else 0
                profit_items.append((data['display_name'], profit, profit_pct))

        profit_items.sort(key=lambda x: x[1], reverse=True)
        
        for rank, (disp_name, profit, profit_pct) in enumerate(profit_items[:5], 1):
            self.app.profit_tree.insert('', 'end', values=(
                rank,
                disp_name,
                f"{profit:,.2f}",
                f"{profit_pct:.2f}%"
            ))

    def update_summary_labels(self):
        total_purchases = len(self.app.purchases)
        cash_purchases = sum(1 for p in self.app.purchases if p.get('payment', '').lower() == 'cash')
        credit_purchases = total_purchases - cash_purchases
        cash_amount = sum(float(p.get('total', 0)) for p in self.app.purchases if p.get('payment', '').lower() == 'cash')
        credit_amount = sum(float(p.get('total', 0)) for p in self.app.purchases if p.get('payment', '').lower() == 'credit')

        total_sales = len(self.app.sales)
        invoice_sales = sum(1 for s in self.app.sales if 'invoice' in s.get('source', '').lower())
        manual_sales = total_sales - invoice_sales
        total_sales_amount = sum(float(s.get('total', 0)) for s in self.app.sales)
        avg_sale = (total_sales_amount / total_sales) if total_sales > 0 else 0

        if hasattr(self.app, 'purchase_items_label'):
            self.app.purchase_items_label.configure(text=str(total_purchases))
        if hasattr(self.app, 'cash_purchase_label'):
            self.app.cash_purchase_label.configure(text=f"PKR {cash_amount:,.2f}")
        if hasattr(self.app, 'credit_purchase_label'):
            self.app.credit_purchase_label.configure(text=f"PKR {credit_amount:,.2f}")

        if hasattr(self.app, 'sales_items_label'):
            self.app.sales_items_label.configure(text=str(total_sales))
        if hasattr(self.app, 'invoice_sales_label'):
            self.app.invoice_sales_label.configure(text=str(invoice_sales))
        if hasattr(self.app, 'manual_sales_label'):
            self.app.manual_sales_label.configure(text=str(manual_sales))
        if hasattr(self.app, 'avg_sale_label'):
            self.app.avg_sale_label.configure(text=f"PKR {avg_sale:,.2f}")

    def refresh_all_data(self):
        self.update_qty_movement()
        self.update_profit_items()
        self.update_summary_labels()
        try:
            self.parent.update_idletasks()
        except:
            pass

    def _copy_qty_movement_data(self):
        if not hasattr(self.app, 'qty_movement_tree') or not self.app.qty_movement_tree:
            messagebox.showwarning("No Data", "No quantity movement data to copy")
            return

        tree = self.app.qty_movement_tree
        items = tree.get_children()
        if not items:
            messagebox.showwarning("No Data", "No quantity movement data to copy")
            return

        headers = ['Item Name', 'Purchased', 'Sold', 'Remaining', 'Revenue (PKR)']
        rows = ['\t'.join(headers)]
        for item_id in items:
            values = tree.item(item_id, 'values')
            rows.append('\t'.join(str(v) for v in values))

        clipboard_data = '\n'.join(rows)
        try:
            self.parent.clipboard_clear()
            self.parent.clipboard_append(clipboard_data)
            self.parent.update()
            messagebox.showinfo("Success", "Quantity movement data copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy data: {str(e)}")

    # ===== Right-Click Edit =====
    def _on_qty_tree_right_click(self, event):
        tree = self.app.qty_movement_tree
        item_id = tree.identify_row(event.y)
        if not item_id:
            try:
                rel_y = int(getattr(event, 'y_root', 0) - tree.winfo_rooty())
                item_id = tree.identify_row(rel_y)
            except Exception:
                item_id = None
        if not item_id:
            try:
                pointer_rel_y = tree.winfo_pointery() - tree.winfo_rooty()
                item_id = tree.identify_row(int(pointer_rel_y))
            except Exception:
                item_id = None
        if not item_id:
            return

        tree.selection_set(item_id)
        tree.focus(item_id)

        values = tree.item(item_id, 'values')
        if not values or not values[0]:
            return

        veg_display_name = values[0]

        # Match using English name
        purchase_index = None
        purchase_entry = None
        for i, p in enumerate(self.app.purchases):
            if self._matches_veg_name(p['vegetable_english'], veg_display_name):
                purchase_index = i
                purchase_entry = p
                break

        if purchase_entry is None:
            if messagebox.askyesno("Add Purchase", "No purchase entry found for this item.\n\nWould you like to add a purchase entry now?"):
                self._open_purchase_create_dialog(veg_display_name)
            return

        self._open_purchase_edit_dialog(veg_display_name, purchase_index, purchase_entry)

    def _matches_veg_name(self, stored_english, display_name):
        """Match English name from purchase record to display name in tree."""
        extracted = self._extract_english_name(display_name)
        return stored_english.strip() == (extracted or "").strip()

    def _open_purchase_edit_dialog(self, veg_name, index, purchase):
        # Safely get quantity
        raw_qty = purchase.get('quantity', '0 kg')
        if raw_qty is None:
            raw_qty = '0 kg'
        elif not isinstance(raw_qty, str):
            raw_qty = str(raw_qty)
        if not raw_qty.strip():
            raw_qty = '0 kg'
        qty, current_unit = self._parse_qty_and_unit(raw_qty)

        # Safely get rate
        raw_rate = purchase.get('rate', '0')
        try:
            current_rate = float(raw_rate)
        except (ValueError, TypeError):
            current_rate = 0.0

        dialog = ctk.CTkToplevel(self.parent)
        dialog.title(f"Edit Purchase: {veg_name}")
        dialog.geometry("350x200")
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text=f"Editing: {veg_name}", font=('Arial', 12, 'bold')).pack(pady=(10, 15))

        qty_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        qty_frame.pack(pady=5, padx=20, fill='x')
        ctk.CTkLabel(qty_frame, text="Quantity:", width=80, anchor='w').pack(side='left')
        qty_var = tk.StringVar(value=f"{qty:.2f}")
        ctk.CTkEntry(qty_frame, textvariable=qty_var, width=80).pack(side='left', padx=5)
        unit_var = tk.StringVar(value=current_unit)
        unit_combo = ttk.Combobox(
            qty_frame,
            textvariable=unit_var,
            values=['kg', 'piece', 'dozen', 'bundle'],
            width=10,
            state='readonly'
        )
        unit_combo.pack(side='left', padx=5)

        rate_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        rate_frame.pack(pady=5, padx=20, fill='x')
        ctk.CTkLabel(rate_frame, text="Rate (PKR):", width=80, anchor='w').pack(side='left')
        rate_var = tk.StringVar(value=f"{current_rate:.2f}")
        ctk.CTkEntry(rate_frame, textvariable=rate_var, width=120).pack(side='left', padx=5)

        def save_changes():
            try:
                qty_val = float(qty_var.get())
                unit = unit_var.get()
                rate = float(rate_var.get())
            except ValueError:
                messagebox.showerror("Input Error", "Please enter valid numbers for quantity and rate.")
                return

            self.app.purchases[index]['quantity'] = f"{qty_val:.2f} {unit}"
            self.app.purchases[index]['rate'] = str(rate)
            self.app.purchases[index]['total'] = str(round(qty_val * rate, 2))

            if hasattr(self.app, 'imported_purchase_rates'):
                extracted = self._extract_english_name(veg_name)
                if extracted:
                    self.app.imported_purchase_rates[extracted] = rate

            self.app.refresh_daily_summary()
            dialog.destroy()

        ctk.CTkButton(dialog, text="Save Changes", command=save_changes, width=100).pack(pady=15)

    def _open_purchase_create_dialog(self, veg_name):
        dialog = ctk.CTkToplevel(self.parent)
        dialog.title(f"Add Purchase: {veg_name}")
        dialog.geometry("400x380")
        dialog.resizable(False, False)
        dialog.transient(self.parent)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text=f"Add Purchase for: {veg_name}", font=('Arial', 12, 'bold')).pack(pady=(15, 15))

        qty_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        qty_frame.pack(pady=8, padx=20, fill='x')
        ctk.CTkLabel(qty_frame, text="Quantity:", width=80, anchor='w').pack(side='left')
        qty_var = tk.StringVar(value="0.00")
        ctk.CTkEntry(qty_frame, textvariable=qty_var, width=80).pack(side='left', padx=5)
        unit_var = tk.StringVar(value='kg')
        unit_combo = ttk.Combobox(
            qty_frame,
            textvariable=unit_var,
            values=['kg', 'piece', 'dozen', 'bundle'],
            width=10,
            state='readonly'
        )
        unit_combo.pack(side='left', padx=5)

        rate_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        rate_frame.pack(pady=8, padx=20, fill='x')
        ctk.CTkLabel(rate_frame, text="Rate (PKR):", width=80, anchor='w').pack(side='left')
        rate_var = tk.StringVar(value='0.00')
        ctk.CTkEntry(rate_frame, textvariable=rate_var, width=120).pack(side='left', padx=5)

        vendor_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        vendor_frame.pack(pady=8, padx=20, fill='x')
        ctk.CTkLabel(vendor_frame, text="Vendor:", width=80, anchor='w').pack(side='left')
        vendor_var = tk.StringVar(value=getattr(self.app, 'purchase_vendor_var', tk.StringVar(value='Main Vendor')).get() if hasattr(self.app, 'purchase_vendor_var') else 'Main Vendor')
        ctk.CTkEntry(vendor_frame, textvariable=vendor_var, width=200).pack(side='left', padx=5, fill='x', expand=True)

        payment_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        payment_frame.pack(pady=8, padx=20, fill='x')
        ctk.CTkLabel(payment_frame, text="Payment:", width=80, anchor='w').pack(side='left')
        payment_var = tk.StringVar(value='cash')
        payment_btn = ctk.CTkSegmentedButton(
            payment_frame,
            values=["Cash", "Credit"],
            command=lambda v: payment_var.set("cash" if v == "Cash" else "credit"),
            width=150
        )
        payment_btn.set("Cash")
        payment_btn.pack(side='left', padx=5)

        def save_new_purchase():
            try:
                qty = float(qty_var.get())
                rate = float(rate_var.get())
            except ValueError:
                messagebox.showerror("Input Error", "Please enter valid numbers for quantity and rate.")
                return

            # Extract English name for storage
            extracted = self._extract_english_name(veg_name)
            if extracted:
                stored_veg = extracted
                urdu_part = veg_name.split('(')[0].strip()
            else:
                stored_veg = veg_name
                urdu_part = veg_name

            purchase = {
                'vegetable_urdu': urdu_part,
                'vegetable_english': stored_veg,
                'vegetable_display': veg_name,
                'quantity': f"{qty:.2f} {unit_var.get()}",
                'rate': f"{rate:.2f}",
                'total': f"{(qty*rate):.2f}",
                'vendor': vendor_var.get(),
                'payment': payment_var.get()
            }

            self.app.purchases.append(purchase)
            self.app.refresh_daily_summary()

            if hasattr(self.app, 'purchase_tab_instance') and hasattr(self.app.purchase_tab_instance, 'reload_purchase_list'):
                try:
                    self.app.purchase_tab_instance.reload_purchase_list()
                except Exception:
                    pass

            dialog.destroy()

        button_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        button_frame.pack(pady=20, padx=20, fill='x')
        ctk.CTkButton(button_frame, text="Add Purchase", command=save_new_purchase, width=120).pack(side='left', padx=5, expand=True)
        ctk.CTkButton(button_frame, text="Cancel", command=dialog.destroy, width=120).pack(side='left', padx=5, expand=True)