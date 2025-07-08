import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from collections import deque
import os

# --- Core FIFO Logic (Improved) ---

def calculate_fifo_profit_loss(df):
    """
    Calculates profit/loss and remaining holdings from transactions using the FIFO method.

    Args:
        df (pd.DataFrame): DataFrame with transaction data.
                           Must contain 'Date', 'Type', 'Code', 'Quantity', 'Price', 'Fees'.

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
    df.reset_index(inplace=True) # Use original index as a unique ID for buys

    all_realized_gains = []
    all_holdings = {} # Using a dict to store deques for each stock code

    for index, row in df.iterrows():
        code = row['Code']

        # Initialize a queue for the stock code if not present
        if code not in all_holdings:
            all_holdings[code] = deque()

        if row['Type'].lower() == 'buy':
            # Add purchase lot to the specific code's queue
            cost_per_share = row['Price'] + (row['Fees'] / row['Quantity'])
            all_holdings[code].append({
                'original_index': index,
                'quantity': row['Quantity'],
                'cost_per_share': cost_per_share,
                'buy_date': row['Date']
            })

        elif row['Type'].lower() == 'sell':
            quantity_to_sell = row['Quantity']
            sell_price_per_share = row['Price']
            sell_fees = row['Fees']
            sell_date = row['Date']

            # Apportion sell fees across the shares being sold
            fee_per_share_sold = sell_fees / quantity_to_sell if quantity_to_sell else 0

            # Get the buy queue for the current stock
            buy_queue = all_holdings[code]

            while quantity_to_sell > 0 and buy_queue:
                oldest_buy = buy_queue[0]

                match_quantity = min(quantity_to_sell, oldest_buy['quantity'])

                # Calculate profit/loss for this specific portion
                proceeds = match_quantity * (sell_price_per_share - fee_per_share_sold)
                cost_basis = match_quantity * oldest_buy['cost_per_share']
                profit_loss = proceeds - cost_basis
                
                # Calculate if held for over 12 months
                over_12_months = (sell_date - oldest_buy['buy_date']).days > 365

                # Store the detailed result for this matched part of the sale
                all_realized_gains.append({
                    'Sell Date': sell_date,
                    'Code': code,
                    'Quantity Sold': match_quantity,
                    'Sell Price': sell_price_per_share,
                    'Proceeds': round(proceeds, 2),
                    'Acquisition Date': oldest_buy['buy_date'],
                    'Cost Basis per Share': round(oldest_buy['cost_per_share'], 2),
                    'Total Cost Basis': round(cost_basis, 2),
                    'Profit/Loss': round(profit_loss, 2),
                    'Over 12 Months': over_12_months
                })

                # Decrement quantities
                quantity_to_sell -= match_quantity
                oldest_buy['quantity'] -= match_quantity

                # If the entire oldest buy lot is sold, remove it from the queue
                if oldest_buy['quantity'] == 0:
                    buy_queue.popleft()

    # 2. Prepare Remaining Holdings Report
    remaining_holdings_list = []
    for code, buy_queue in all_holdings.items():
        for holding in buy_queue:
            if holding['quantity'] > 0: # Ensure we only report lots with shares left
                remaining_holdings_list.append({
                    'Code': code,
                    'Remaining Quantity': holding['quantity'],
                    'Acquisition Date': holding['buy_date'],
                    'Cost Basis per Share': round(holding['cost_per_share'], 2)
                })

    sales_df = pd.DataFrame(all_realized_gains)
    holdings_df = pd.DataFrame(remaining_holdings_list)

    return sales_df, holdings_df

# --- GUI Application Class (Modified for dual output) ---

class FifoCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FIFO Profit/Loss Calculator V2")
        self.root.geometry("800x600")
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")

        # Main frame
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Top frame for controls
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X, pady=5)

        self.label = ttk.Label(self.top_frame, text="Import your transactions file (.csv or .xlsx).", font=("Helvetica", 12))
        self.label.pack(side=tk.LEFT, padx=5)

        self.import_button = ttk.Button(self.top_frame, text="Import and Process File", command=self.run_full_process)
        self.import_button.pack(side=tk.LEFT, padx=5)

        # Results display area
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(pady=10, expand=True, fill=tk.BOTH)

        self.sales_frame = ttk.Frame(self.notebook)
        self.holdings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sales_frame, text="Sales Report")
        self.notebook.add(self.holdings_frame, text="Holdings Report")

        # Sales Report Tab
        self.sales_tree_frame = ttk.Frame(self.sales_frame)
        self.sales_tree_frame.pack(fill=tk.BOTH, expand=True)
        self.sales_tree = self.create_treeview(self.sales_tree_frame)
        
        self.sales_button_frame = ttk.Frame(self.sales_frame)
        self.sales_button_frame.pack(fill=tk.X)
        self.download_sales_button = ttk.Button(self.sales_button_frame, text="Download Sales CSV", command=lambda: self.download_csv(self.sales_df, "sales_report"))
        self.download_sales_button.pack(pady=5)

        # Holdings Report Tab
        self.holdings_tree_frame = ttk.Frame(self.holdings_frame)
        self.holdings_tree_frame.pack(fill=tk.BOTH, expand=True)
        self.holdings_tree = self.create_treeview(self.holdings_tree_frame)

        self.holdings_button_frame = ttk.Frame(self.holdings_frame)
        self.holdings_button_frame.pack(fill=tk.X)
        self.download_holdings_button = ttk.Button(self.holdings_button_frame, text="Download Holdings CSV", command=lambda: self.download_csv(self.holdings_df, "holdings_report"))
        self.download_holdings_button.pack(pady=5)

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Status: Waiting for file...")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, font=("Helvetica", 10, "italic"))
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.sales_df = pd.DataFrame()
        self.holdings_df = pd.DataFrame()

    def create_treeview(self, parent):
        tree = ttk.Treeview(parent, show="headings")
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.configure(xscrollcommand=hsb.set)

        return tree

    def display_df_in_treeview(self, tree, df):
        # Clear previous data
        for i in tree.get_children():
            tree.delete(i)

        # Set new columns
        tree["columns"] = list(df.columns)
        tree["displaycolumns"] = list(df.columns)

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100) # Adjust column width as needed

        # Add new data
        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def run_full_process(self):
        file_path = self.select_input_file()
        if not file_path:
            return

        try:
            self.status_var.set("Status: Reading file...")
            self.root.update_idletasks()
            df = self.read_transaction_file(file_path)

            self.status_var.set("Status: Processing FIFO calculations...")
            self.root.update_idletasks()
            self.sales_df, self.holdings_df = calculate_fifo_profit_loss(df)

            if self.sales_df.empty and self.holdings_df.empty:
                messagebox.showinfo("No Data", "The input file did not contain valid transaction data.")
                self.status_var.set("Status: Done. No data to process.")
                return

            self.status_var.set("Status: Calculation complete.")
            self.display_df_in_treeview(self.sales_tree, self.sales_df)
            self.display_df_in_treeview(self.holdings_tree, self.holdings_df)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_var.set(f"Status: Error - {e}")

    def select_input_file(self):
        return filedialog.askopenfilename(title="Select Transaction File", filetypes=(("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")))

    def read_transaction_file(self, file_path):
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file type.")

        required_columns = {'Date', 'Type', 'Code', 'Quantity', 'Price', 'Fees'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"Input file missing required columns: {', '.join(missing)}")

        self.status_var.set(f"Status: Loaded {os.path.basename(file_path)}")
        return df

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
            messagebox.showinfo("Success", f"{report_name} saved successfully to\n{save_path}")
            self.status_var.set(f"Status: {report_name} exported.")

if __name__ == "__main__":
    root = tk.Tk()
    app = FifoCalculatorApp(root)
    root.mainloop()