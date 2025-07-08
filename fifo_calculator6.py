import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from collections import deque
import os

# --- Core FIFO Logic (Improved Date Formatting) ---

def calculate_fifo_reports(df):
    """
    Calculates profit/loss and remaining holdings from transactions using the FIFO method.
    This version formats all output dates as DD/MM/YYYY.

    Args:
        df (pd.DataFrame): DataFrame with transaction data.
                           Must contain 'Date', 'Type', 'Code', 'Quantity', 'Total'.

    Returns:
        tuple[pd.DataFrame, pd.DataFrame]:
            - A DataFrame of detailed realized gains for each sale portion.
            - A DataFrame of remaining unsold holdings.
    """
    # 1. Data Preparation
    try:
        df['Date'] = pd.to_datetime(df['Date'])
    except Exception:
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)

    df.sort_values(by='Date', inplace=True)
    df.reset_index(inplace=True, drop=True)

    all_realized_gains = []
    all_holdings = {}  # Using a dict to store deques for each stock code

    for index, row in df.iterrows():
        code = row['Code']

        if code not in all_holdings:
            all_holdings[code] = deque()

        if row['Type'].lower() == 'buy':
            quantity = abs(row['Quantity'])
            total_cost = abs(row['Total'])
            
            if quantity == 0:
                continue

            cost_per_share = total_cost / quantity
            all_holdings[code].append({
                'quantity': quantity,
                'cost_per_share': cost_per_share,
                'buy_date': row['Date']
            })

        elif row['Type'].lower() == 'sell':
            quantity_to_sell = abs(row['Quantity'])
            sell_total = abs(row['Total'])
            sell_date = row['Date']
            
            if quantity_to_sell == 0:
                continue

            buy_queue = all_holdings.get(code)
            
            total_held_quantity = sum(holding['quantity'] for holding in buy_queue) if buy_queue else 0
            if quantity_to_sell > total_held_quantity:
                raise ValueError(
                    f"Oversell Error on {sell_date.date()} for stock '{code}'.\n\n"
                    f"Attempted to sell {quantity_to_sell} shares, but only {total_held_quantity} were held at the time."
                )

            proceeds_per_share = sell_total / quantity_to_sell

            while quantity_to_sell > 0 and buy_queue:
                oldest_buy = buy_queue[0]
                match_quantity = min(quantity_to_sell, oldest_buy['quantity'])

                proceeds = match_quantity * proceeds_per_share
                cost_basis = match_quantity * oldest_buy['cost_per_share']
                profit_loss = proceeds - cost_basis

                over_12_months = (sell_date - oldest_buy['buy_date']).days > 365

                cgd = 0
                if over_12_months and profit_loss > 0:
                    cgd = profit_loss * 0.5

                all_realized_gains.append({
                    'Sell Date': sell_date,
                    'Code': code,
                    'Quantity Sold': match_quantity,
                    'Proceeds per Share': round(proceeds_per_share, 4),
                    'Total Proceeds': round(proceeds, 2),
                    'Acquisition Date': oldest_buy['buy_date'],
                    'Cost Basis per Share': round(oldest_buy['cost_per_share'], 4),
                    'Total Cost Basis': round(cost_basis, 2),
                    'Profit/Loss': round(profit_loss, 2),
                    'Over 12 Months': over_12_months,
                    'CGD': round(cgd, 2)
                })

                quantity_to_sell -= match_quantity
                oldest_buy['quantity'] -= match_quantity

                if oldest_buy['quantity'] == 0:
                    buy_queue.popleft()

    # 2. Prepare Final Reports
    remaining_holdings_list = []
    for code, buy_queue in all_holdings.items():
        for holding in buy_queue:
            if holding['quantity'] > 0:
                total_cost = holding['quantity'] * holding['cost_per_share']
                remaining_holdings_list.append({
                    'Code': code,
                    'Remaining Quantity': holding['quantity'],
                    'Acquisition Date': holding['buy_date'],
                    'Cost Basis per Share': round(holding['cost_per_share'], 4),
                    'Total Cost': round(total_cost, 2)
                })

    # --- IMPROVEMENT: Format all dates to DD/MM/YYYY for final output ---
    date_format = '%d/%m/%Y'
    for gain in all_realized_gains:
        if isinstance(gain['Sell Date'], pd.Timestamp):
            gain['Sell Date'] = gain['Sell Date'].strftime(date_format)
        if isinstance(gain['Acquisition Date'], pd.Timestamp):
            gain['Acquisition Date'] = gain['Acquisition Date'].strftime(date_format)

    for holding in remaining_holdings_list:
        if isinstance(holding['Acquisition Date'], pd.Timestamp):
            holding['Acquisition Date'] = holding['Acquisition Date'].strftime(date_format)


    sales_df = pd.DataFrame(all_realized_gains)
    holdings_df = pd.DataFrame(remaining_holdings_list)

    return sales_df, holdings_df

# --- GUI Application Class (Upgraded) ---

class FifoCalculatorAppV6:
    def __init__(self, root):
        self.root = root
        self.root.title("FIFO Profit/Loss Calculator V6")
        self.root.geometry("950x700")
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X, pady=5)

        self.label = ttk.Label(self.top_frame, text="Import transactions (.csv/.xlsx). Required columns: 'Date', 'Type', 'Code', 'Quantity', 'Total'", font=("Helvetica", 12))
        self.label.pack(side=tk.LEFT, padx=5)

        self.import_button = ttk.Button(self.top_frame, text="Import and Process File", command=self.run_full_process)
        self.import_button.pack(side=tk.LEFT, padx=5)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(pady=10, expand=True, fill=tk.BOTH)

        self.sales_frame = ttk.Frame(self.notebook, padding=5)
        self.holdings_frame = ttk.Frame(self.notebook, padding=5)
        self.summary_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.sales_frame, text="Sales Report")
        self.notebook.add(self.holdings_frame, text="Holdings Report")
        self.notebook.add(self.summary_frame, text="Profit/Loss Summary")

        # --- Tab 1: Sales Report ---
        self.sales_tree = self.create_treeview(self.sales_frame)
        self.download_sales_button = ttk.Button(self.sales_frame, text="Download Sales CSV", command=lambda: self.download_csv(self.sales_df, "sales_report"))
        self.download_sales_button.pack(pady=5)

        # --- Tab 2: Holdings Report ---
        self.holdings_tree = self.create_treeview(self.holdings_frame)
        self.holdings_total_label = ttk.Label(self.holdings_frame, text="Total Cost of Holdings: $0.00", font=("Helvetica", 12, "bold"))
        self.holdings_total_label.pack(pady=10, side=tk.BOTTOM)
        self.download_holdings_button = ttk.Button(self.holdings_frame, text="Download Holdings CSV", command=lambda: self.download_csv(self.holdings_df, "holdings_report"))
        self.download_holdings_button.pack(pady=5)

        # --- Tab 3: Summary Report ---
        self.summary_totals_frame = ttk.Frame(self.summary_frame)
        self.summary_totals_frame.pack(fill=tk.X, pady=10)
        self.pl_total_label = ttk.Label(self.summary_totals_frame, text="Total Profit/Loss: $0.00", font=("Helvetica", 14, "bold"))
        self.pl_total_label.pack()
        self.cgd_total_label = ttk.Label(self.summary_totals_frame, text="Total Capital Gains Discount (CGD): $0.00", font=("Helvetica", 14, "bold"))
        self.cgd_total_label.pack()

        ttk.Separator(self.summary_frame, orient='horizontal').pack(fill='x', pady=10)
        ttk.Label(self.summary_frame, text="Breakdown by Stock Code", font=("Helvetica", 12, "italic")).pack()
        
        self.summary_tree = self.create_treeview(self.summary_frame)
        self.download_summary_button = ttk.Button(self.summary_frame, text="Download Summary CSV", command=self.download_summary_report)
        self.download_summary_button.pack(pady=5)


        # --- Status Bar ---
        self.status_var = tk.StringVar()
        self.status_var.set("Status: Waiting for file...")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, font=("Helvetica", 10, "italic"))
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.sales_df = pd.DataFrame()
        self.holdings_df = pd.DataFrame()
        self.summary_df = pd.DataFrame()

    def create_treeview(self, parent):
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(tree_frame, show="headings")
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=vsb.set)
        
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.configure(xscrollcommand=hsb.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        return tree

    def display_df_in_treeview(self, tree, df):
        for i in tree.get_children():
            tree.delete(i)
        
        if df.empty:
            tree["columns"] = []
            return

        tree["columns"] = list(df.columns)
        tree["displaycolumns"] = list(df.columns)

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor=tk.W)

        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def run_full_process(self):
        file_path = filedialog.askopenfilename(title="Select Transaction File", filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")))
        if not file_path:
            return

        try:
            self.status_var.set("Status: Reading file...")
            self.root.update_idletasks()
            df = self.read_transaction_file(file_path)

            self.status_var.set("Status: Processing FIFO calculations...")
            self.root.update_idletasks()
            self.sales_df, self.holdings_df = calculate_fifo_reports(df)

            self.status_var.set("Status: Calculation complete. Displaying reports...")
            # Display reports in their respective tabs
            self.display_df_in_treeview(self.sales_tree, self.sales_df)
            self.display_holdings_report()
            self.display_summary_report()
            self.status_var.set("Status: Done.")

        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred:\n\n{e}")
            self.status_var.set(f"Status: Error - {e}")

    def read_transaction_file(self, file_path):
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file type selected.")

        required_columns = {'Date', 'Type', 'Code', 'Quantity', 'Total'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"Input file is missing required columns: {', '.join(missing)}")

        self.status_var.set(f"Status: Loaded {os.path.basename(file_path)}")
        return df

    def display_holdings_report(self):
        self.display_df_in_treeview(self.holdings_tree, self.holdings_df)
        if not self.holdings_df.empty:
            total_cost = self.holdings_df['Total Cost'].sum()
            self.holdings_total_label.config(text=f"Total Cost of Holdings: ${total_cost:,.2f}")
        else:
            self.holdings_total_label.config(text="Total Cost of Holdings: $0.00")

    def display_summary_report(self):
        # Clear previous data
        self.pl_total_label.config(text="Total Profit/Loss: $0.00")
        self.cgd_total_label.config(text="Total Capital Gains Discount (CGD): $0.00")
        self.display_df_in_treeview(self.summary_tree, pd.DataFrame())

        if self.sales_df.empty:
            return

        # Recalculate totals from the dataframe in case of rounding differences
        total_pl = self.sales_df['Profit/Loss'].sum()
        total_cgd = self.sales_df['CGD'].sum()

        self.pl_total_label.config(text=f"Total Profit/Loss: ${total_pl:,.2f}")
        self.cgd_total_label.config(text=f"Total Capital Gains Discount (CGD): ${total_cgd:,.2f}")

        breakdown_df = self.sales_df.groupby('Code').agg(
            Total_Profit_Loss=('Profit/Loss', 'sum'),
            Total_CGD=('CGD', 'sum')
        ).reset_index()
        breakdown_df['Total_Profit_Loss'] = breakdown_df['Total_Profit_Loss'].round(2)
        breakdown_df['Total_CGD'] = breakdown_df['Total_CGD'].round(2)
        
        self.summary_df = breakdown_df
        self.display_df_in_treeview(self.summary_tree, self.summary_df)

    def download_csv(self, df, report_name):
        if df.empty:
            messagebox.showwarning("No Data", f"There is no data in the {report_name} to download.")
            return

        save_path = filedialog.asksaveasfilename(
            title=f"Save {report_name}",
            defaultextension=".csv",
            filetypes=(("CSV file", "*.csv"),),
            initialfile=f"{report_name}.csv"
        )
        if save_path:
            df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"{report_name} was saved successfully.")
            self.status_var.set(f"Status: {report_name} exported.")

    def download_summary_report(self):
        if self.sales_df.empty:
            messagebox.showwarning("No Data", "There is no summary data to download.")
            return
            
        total_pl = self.sales_df['Profit/Loss'].sum()
        total_cgd = self.sales_df['CGD'].sum()
        
        summary_data = {
            'Metric': [
                'Total Profit/Loss', 
                'Total Capital Gains Discount (CGD)',
                '---',
                'Code',
            ],
            'Value': [
                round(total_pl, 2), 
                round(total_cgd, 2),
                '---',
                'Total_Profit_Loss',
            ],
            'CGD_Breakdown': [
                '',
                '',
                '---',
                'Total_CGD'
            ]
        }
        
        header_df = pd.DataFrame(summary_data)
        
        breakdown_for_export = self.summary_df.rename(columns={
            'Code': 'Metric', 
            'Total_Profit_Loss': 'Value',
            'Total_CGD': 'CGD_Breakdown'
            })

        final_summary_df = pd.concat([header_df, breakdown_for_export], ignore_index=True)

        save_path = filedialog.asksaveasfilename(
            title="Save Summary Report",
            defaultextension=".csv",
            filetypes=(("CSV file", "*.csv"),),
            initialfile="profit_loss_summary.csv"
        )

        if save_path:
            final_summary_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Summary report was saved successfully.")
            self.status_var.set(f"Status: Summary report exported.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FifoCalculatorAppV6(root)
    root.mainloop()