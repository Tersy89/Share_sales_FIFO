import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from collections import deque

# --- Core FIFO Logic ---

def calculate_fifo_profit_loss(df):
    """
    Calculates profit/loss from a dataframe of transactions using the FIFO method.

    Args:
        df (pd.DataFrame): DataFrame with transaction data. 
                           Must contain 'Date', 'Type', 'Code', 
                           'Quantity', 'Price', and 'Fees'.

    Returns:
        pd.DataFrame: A new DataFrame containing the realized gains/losses for each sale.
    """
    # 1. Data Preparation
    # Ensure 'Date' is a datetime object and sort by it
    try:
        df['Date'] = pd.to_datetime(df['Date'])
    except Exception:
        # Try a different date format if the first one fails
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
        
    df.sort_values(by='Date', inplace=True)
    
    # 2. Group by stock 'Code' and process each group
    all_realized_gains = []
    
    for code, group in df.groupby('Code'):
        buy_queue = deque() # Use a deque for efficient popping from the left (FIFO)
        
        for index, row in group.iterrows():
            if row['Type'].lower() == 'buy':
                # Add purchase lot to the queue
                # The cost basis includes the purchase price and fees
                cost_per_share = row['Price'] + (row['Fees'] / row['Quantity'])
                buy_queue.append({
                    'quantity': row['Quantity'],
                    'cost_per_share': cost_per_share,
                    'date': row['Date']
                })

            elif row['Type'].lower() == 'sell':
                quantity_to_sell = row['Quantity']
                proceeds_from_sale = (row['Price'] * quantity_to_sell) - row['Fees']
                total_cost_basis_of_sold_shares = 0
                
                while quantity_to_sell > 0 and buy_queue:
                    oldest_buy = buy_queue[0]
                    
                    # Determine how many shares to sell from the oldest buy lot
                    match_quantity = min(quantity_to_sell, oldest_buy['quantity'])
                    
                    # Add to the cost basis for this specific sale
                    total_cost_basis_of_sold_shares += match_quantity * oldest_buy['cost_per_share']
                    
                    # Decrement the quantities
                    quantity_to_sell -= match_quantity
                    oldest_buy['quantity'] -= match_quantity
                    
                    # If the entire oldest buy lot is sold, remove it from the queue
                    if oldest_buy['quantity'] == 0:
                        buy_queue.popleft()
                
                # Calculate profit/loss for this sale
                profit_loss = proceeds_from_sale - total_cost_basis_of_sold_shares
                
                # Store the result
                all_realized_gains.append({
                    'Sell Date': row['Date'],
                    'Code': code,
                    'Quantity Sold': row['Quantity'],
                    'Profit/Loss': round(profit_loss, 2)
                })

    return pd.DataFrame(all_realized_gains)

# --- GUI Application Class ---

class FifoCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FIFO Profit/Loss Calculator")
        self.root.geometry("500x250")

        # Style
        self.style = ttk.Style(self.root)
        self.style.theme_use("clam") # 'clam', 'alt', 'default', 'classic'

        # Main frame
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # Widgets
        self.label = ttk.Label(
            self.main_frame, 
            text="Please import your transactions file (.csv or .xlsx).",
            font=("Helvetica", 12)
        )
        self.label.pack(pady=10)

        self.import_button = ttk.Button(
            self.main_frame, 
            text="Import and Process File", 
            command=self.run_full_process
        )
        self.import_button.pack(pady=15, ipadx=10, ipady=5)

        self.status_var = tk.StringVar()
        self.status_var.set("Status: Waiting for file...")
        self.status_label = ttk.Label(
            self.main_frame, 
            textvariable=self.status_var, 
            font=("Helvetica", 10, "italic")
        )
        self.status_label.pack(pady=20)

    def run_full_process(self):
        """Handles the entire workflow from file selection to saving the result."""
        file_path = self.select_input_file()
        if not file_path:
            return # User cancelled the dialog

        try:
            self.status_var.set("Status: Reading file...")
            self.root.update_idletasks() # Force GUI update
            
            df = self.read_transaction_file(file_path)

            self.status_var.set("Status: Processing FIFO calculations...")
            self.root.update_idletasks()

            results_df = calculate_fifo_profit_loss(df)

            if results_df.empty:
                messagebox.showinfo("No Sales Found", "The calculation is complete, but no 'Sell' transactions were found to generate a profit/loss report.")
                self.status_var.set("Status: Done. No sales to report.")
                return

            self.status_var.set("Status: Calculation complete. Please save the results.")
            self.root.update_idletasks()

            self.save_output_file(results_df)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            self.status_var.set("Status: An error occurred. Please try again.")

    def select_input_file(self):
        """Opens a file dialog to select the input file."""
        file_path = filedialog.askopenfilename(
            title="Select Transaction File",
            filetypes=(
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            )
        )
        return file_path

    def read_transaction_file(self, file_path):
        """Reads the selected file into a pandas DataFrame."""
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            raise ValueError("Unsupported file type. Please use .csv or .xlsx.")
        
        # Validate required columns
        required_columns = {'Date', 'Type', 'Code', 'Quantity', 'Price', 'Fees'}
        if not required_columns.issubset(df.columns):
            missing = required_columns - set(df.columns)
            raise ValueError(f"Input file is missing required columns: {', '.join(missing)}")
            
        self.status_var.set(f"Status: Loaded {file_path.split('/')[-1]}")
        return df

    def save_output_file(self, results_df):
        """Opens a save dialog and saves the DataFrame to a CSV file."""
        save_path = filedialog.asksaveasfilename(
            title="Save Results As",
            defaultextension=".csv",
            filetypes=(("CSV file", "*.csv"),)
        )
        if save_path:
            results_df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"Results successfully saved to:\n{save_path}")
            self.status_var.set("Status: Success! Results exported.")
        else:
            self.status_var.set("Status: Processing complete. Export was cancelled.")


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = FifoCalculatorApp(root)
    root.mainloop()