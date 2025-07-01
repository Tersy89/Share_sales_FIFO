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
                    'Profit/Loss': round(profit_loss, 2)
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
        self.root.geometry("500x250")
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.label = ttk.Label(self.main_frame, text="Import your transactions file (.csv or .xlsx).", font=("Helvetica", 12))
        self.label.pack(pady=10)
        self.import_button = ttk.Button(self.main_frame, text="Import and Process File", command=self.run_full_process)
        self.import_button.pack(pady=15, ipadx=10, ipady=5)
        self.status_var = tk.StringVar()
        self.status_var.set("Status: Waiting for file...")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, font=("Helvetica", 10, "italic"))
        self.status_label.pack(pady=20)

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
            sales_df, holdings_df = calculate_fifo_profit_loss(df)

            if sales_df.empty and holdings_df.empty:
                messagebox.showinfo("No Data", "The input file did not contain valid transaction data.")
                self.status_var.set("Status: Done. No data to process.")
                return

            self.status_var.set("Status: Calculation complete. Please save the results.")
            self.root.update_idletasks()
            self.save_output_files(sales_df, holdings_df)

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

    def save_output_files(self, sales_df, holdings_df):
        save_path = filedialog.asksaveasfilename(
            title="Choose a base name for your report files",
            defaultextension=".csv",
            filetypes=(("CSV file", "*.csv"),)
        )
        
        if not save_path:
            self.status_var.set("Status: Processing complete. Export was cancelled.")
            return

        # Create two distinct filenames from the base name
        base, ext = os.path.splitext(save_path)
        sales_path = f"{base}_sales_report.csv"
        holdings_path = f"{base}_holdings_report.csv"

        saved_files = []
        # Save the detailed sales report if it has data
        if not sales_df.empty:
            sales_df.to_csv(sales_path, index=False)
            saved_files.append(os.path.basename(sales_path))
        
        # Save the remaining holdings report if it has data
        if not holdings_df.empty:
            holdings_df.to_csv(holdings_path, index=False)
            saved_files.append(os.path.basename(holdings_path))

        if not saved_files:
            messagebox.showinfo("No Reports Generated", "No sales were made and no holdings remain based on the data.")
            self.status_var.set("Status: Done. No reports to save.")
            return

        messagebox.showinfo("Success", f"Reports successfully saved:\n\n" + "\n".join(saved_files))
        self.status_var.set("Status: Success! Reports exported.")


if __name__ == "__main__":
    root = tk.Tk()
    app = FifoCalculatorApp(root)
    root.mainloop()