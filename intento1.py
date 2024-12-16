import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox



def process_file():
    # Open a file dialog to select the Excel file
    filepath = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not filepath:
        return
    
    try:
        # Load the Excel file
        data = pd.read_excel(filepath)
        
        # Ensure columns `transacciones` and `ID` exist
        if 'transaccion' not in data.columns or 'ID' not in data.columns:
            messagebox.showerror("Error", "Required columns 'transacciones' or 'ID' not found.")
            return
        
        # Identify rows where `ID` is missing
        missing_id = data[data['ID'].isna()]
        
        # Open a dialog to save the output
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Output File"
        )
        if not output_path:
            return
        
        # Save results
        missing_id.to_excel(output_path, index=False)
        messagebox.showinfo("Success", "File processed and saved successfully!")
    
    except Exception as e:
        # Handle errors
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI setup
root = tk.Tk()
root.title("Excel Processor")
root.geometry("300x150")

# Add a button to trigger the process
process_button = tk.Button(root, text="Process Excel File", command=process_file)
process_button.pack(pady=40)

root.mainloop()
