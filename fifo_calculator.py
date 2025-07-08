import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from collections import deque
import os

# --- Core FIFO Logic ---

def calculate_fifo_profit_loss(df):
    """
    Calculates profit/loss and remaining holdings from transactions using the FIFO method.
    This function assumes it receives a cleaned and validated DataFrame.

    Args:
        df (pd.DataFrame): DataFrame with transaction data.
                           Must contain 'Date', 'Type', 'Code', 'Quantity', 'Price', 'Fees'.
                           All columns should have the correct data types.

    Returns:
        tuple[pd.DataFrame, pd.DataFrame]:
            - A DataFrame of detailed realized gains for each sale portion.
            - A DataFrame of remaining unsold holdings.
    """
    # Data is assumed to be pre-sorted by date from the reading function
    df.reset_index(inplace=True, drop=True) 

    all_realized_gains = []
    all_holdings = {}  # Using a dict to store deques for each stock code

    for index, row in df.iterrows():
        code = row['Code']
        
        # Initialize a queue for the stock code if not present
        if code not in all_holdings:
            all_holdings[code] = deque()

        if row['Type'].lower() == 'buy':
            # Add purchase lot to the specific code's queue
            # Fees are now assumed to be part of the cost basis calculation
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
            buy_queue = all_holdings.get(code) # Use .get() for safety

            if not buy_queue:
                # This case handles selling a stock that was never bought according to the sheet
                # You might want to log this or show a warning
                continue

            while quantity_to_sell > 0 and buy_queue:
                oldest_buy = buy_queue[0]
                
                match_quantity = min(quantity_to_sell, oldest_buy['quantity'])
                
                # Calculate profit/loss for this specific portion
                proceeds = match_quantity * sell_price_per_share - (match_quantity * fee_per_share_sold)
                cost_basis = match_quantity * oldest_buy['cost_per_share']
                profit_loss = proceeds - cost_basis
                
                # Store the detailed result for this matched part of the sale
                all_realized_gains.append({
                    'Sell Date': sell_date.strftime('%Y-%m-%d'),
                    'Code': code,
                    'Quantity Sold': match_quantity,
                    'Sell Price': sell_price_per_share,
                    'Proceeds': round(proceeds, 2),
                    'Acquisition Date': oldest_buy['buy_date'].strftime('%Y-%m-%d'),
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

    # Prepare Remaining Holdings Report
    remaining_holdings_list = []
    for code, buy_queue in all_holdings.items():
        for holding in buy_queue:
            if holding['quantity'] > 0:  # Ensure we only report lots with shares left
                remaining_holdings_list.append({
                    'Code': code,
                    'Remaining Quantity': holding['quantity'],
                    'Acquisition Date': holding['buy_date'].strftime('%Y-%m-%d'),
                    'Cost Basis per Share': round(holding['cost_per_share'], 2)
                })

    sales_df = pd.DataFrame(all_realized_gains)
    holdings_df = pd.DataFrame(remaining_holdings_list)

    return sales_df, holdings_df

# --- GUI Application Class ---

class FifoCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FIFO Profit/Loss Calculator V3")
        self.root.geometry("550x250")
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam")
        
        # Configure style for better appearance
        self.style.configure("TButton", padding=6, relief="flat", background="#0078D7", foreground="white")
        self.style.map("TButton", background=[('active', '#005A9E')])
        self.style.configure("TFrame", background="#F0F0F0")
        self.style.configure("TLabel", background="#F0F0F0")

        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(self.main_frame, text="Import your transactions file (.csv, .xlsx, .xls).", font=("Helvetica", 12))
        self.label.pack(pady=10)

        self.import_button = ttk.Button(self.main_frame, text="Import and Process File", command=self.run_full_process)
        self.import_button.pack(pady=15, ipadx=10, ipady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Status: Waiting for file...")
        self.status_label = ttk.Label(self.main_frame, textvariable=self.status_var, font=("Helvetica", 10, "italic"), wraplength=500)
        self.status_label.pack(pady=20)

    def run_full_process(self):
        file_path = self.select_input_file()
        if not file_path:
            return

        try:
            self.status_var.set("Status: Reading and cleaning file...")
            self.root.update_idletasks()
            df = self.read_and_clean_transaction_file(file_path)

            self.status_var.set("Status: Processing FIFO calculations...")
            self.root.update_idletasks()
            sales_df, holdings_df = calculate_fifo_profit_loss(df)

            if sales_df.empty and holdings_df.empty:
                messagebox.showinfo("No Data", "The input file did not contain valid transaction data or resulted in no reports.")
                self.status_var.set("Status: Done. No data to process.")
                return

            self.status_var.set("Status: Calculation complete. Please save the results.")
            self.root.update_idletasks()
            self.save_output_files(sales_df, holdings_df)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_var.set(f"Status: Error - {e}")

    def select_input_file(self):
        # Expanded file types to include older .xls format
        return filedialog.askopenfilename(
            title="Select Transaction File",
            filetypes=(
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            )
        )

    def read_and_clean_transaction_file(self, file_path):
        """Reads and performs robust cleaning and validation on the input file."""
        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path, thousands=',')
            elif file_path.lower().endswith(('.xlsx', '.xls')):
                # engine=None lets pandas auto-select the correct engine for xls/xlsx
                df = pd.read_excel(file_path, engine=None)
            else:
                raise ValueError(f"Unsupported file type: {os.path.basename(file_path)}")
        except Exception as e:
            raise ValueError(f"Failed to read file. It might be corrupt or in an unexpected format. Details: {e}")

        # --- Data Cleaning and Validation ---
        
        # 1. Normalize column names
        df.columns = df.columns.str.strip().str.title()
        
        # 2. Check for required columns
        required_columns = {'Date', 'Type', 'Code', 'Quantity', 'Price', 'Fees'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"Input file missing required columns: {', '.join(missing)}")
        
        # 3. Clean and convert data types with robust error handling
        # Clean string columns
        for col in ['Type', 'Code']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        # Clean and convert numeric columns
        for col in ['Quantity', 'Price', 'Fees']:
            if col in df.columns:
                # Convert to string to use string methods, then to numeric
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[$,]', '', regex=True), errors='coerce')

        # Handle missing values after conversion
        df['Fees'].fillna(0.0, inplace=True) # Assume no fee if missing
        
        # Check for critical missing data
        critical_cols = ['Date', 'Type', 'Code', 'Quantity', 'Price']
        nan_rows = df[df[critical_cols].isnull().any(axis=1)]
        if not nan_rows.empty:
            raise ValueError(f"File has rows with missing or invalid data in critical columns. Please check rows: {nan_rows.index.tolist()}")

        # 4. Parse dates with robust error handling
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        if df['Date'].isnull().any():
            # Try another common format if the first fails
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True)
            if df['Date'].isnull().any():
                invalid_dates_indices = df[df['Date'].isnull()].index.tolist()
                raise ValueError(f"Could not parse dates in some rows. Please check date format for rows: {invalid_dates_indices}")
        
        # 5. Sort by date to ensure correct FIFO order
        df.sort_values(by='Date', inplace=True)

        self.status_var.set(f"Status: Successfully loaded and cleaned {os.path.basename(file_path)}")
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