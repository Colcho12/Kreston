import pandas as pd
from datetime import timedelta
import tkinter as tk
from tkinter import filedialog, messagebox

"""Prompt user to upload the Excel file and open sheets"""
def StartEstadoCuenta():
    # Ask the user to upload an Excel file
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file was selected. Exiting...")
        exit()

    try:
        # Load the specific sheets
        estado = pd.read_excel(file_path, sheet_name="ESTADO DE CUENTA")
        auxbanco = pd.read_excel(file_path, sheet_name="AUX BANCOS")
        return estado, auxbanco
    except Exception as e:
        messagebox.showerror("Error", f"Error opening Excel file: {e}")
        exit()

def Cleaning(estado, auxbanco):
    estado["FECHA"] = pd.to_datetime(estado["FECHA"], errors="coerce").dt.date
    auxbanco["Posting Date"] = pd.to_datetime(auxbanco["Posting Date"], errors="coerce").dt.date
    estado["Amount"] = estado["ABONOS"].fillna(0) - estado["CARGOS"].fillna(0)
    estado = estado[estado["FECHA"].notna()]
    return estado, auxbanco

def FirstSearch(estado, auxbanco):
    document_numbers = []
    for index, row in estado.iterrows():
        match = auxbanco[
            (auxbanco["Posting Date"] == row["FECHA"]) &
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (~auxbanco["Used"])
        ].head(1)
        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
        else:
            document_number = None
        document_numbers.append(document_number)
    estado["DOCUMENT NUMBER"] = document_numbers
    return estado

def SecondSearch(estado, auxbanco):
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    unique_amounts = auxbanco["Amount in doc. curr."].value_counts()
    unique_amounts = unique_amounts[unique_amounts == 1].index
    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (auxbanco["Amount in doc. curr."].isin(unique_amounts))
        ].head(1)
        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
    return estado

def ThirdSearch(estado, auxbanco):
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (auxbanco["Posting Date"] >= row["FECHA"] - timedelta(days=7)) &
            (auxbanco["Posting Date"] <= row["FECHA"] + timedelta(days=7)) &
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (~auxbanco["Used"])
        ].head(1)
        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
    return estado

def FourthSearch(estado, auxbanco):
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (auxbanco["Amount in doc. curr."] == -row["Amount"]) &
            (~auxbanco["Used"])
        ].head(1)
        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
    return estado

def FifthSearch(estado, auxbanco, max_days=10):
    def find_consecutive_sum(values, target):
        n = len(values)
        for start in range(n):
            current_sum = 0
            for end in range(start, n):
                current_sum += values[end]
                if current_sum == target:
                    return list(range(start, end + 1))
        return None

    unmatched_aux_rows = auxbanco[~auxbanco["Used"]]
    for aux_index, aux_row in unmatched_aux_rows.iterrows():
        target_amount = aux_row["Amount in doc. curr."]
        target_date = aux_row["Posting Date"]
        document_number = aux_row["Document Number"]
        candidates = estado[
            (estado["FECHA"] >= target_date - timedelta(days=max_days)) &
            (estado["FECHA"] <= target_date + timedelta(days=max_days)) &
            (estado["DOCUMENT NUMBER"].isna())
        ]
        candidate_values = candidates["Amount"].tolist()
        candidate_indices = candidates.index.tolist()
        subset_indices = find_consecutive_sum(candidate_values, target_amount)
        if subset_indices:
            matched_indices = [candidate_indices[i] for i in subset_indices]
            estado.loc[matched_indices, "DOCUMENT NUMBER"] = document_number
            auxbanco.at[aux_index, "Used"] = True
    return estado

def MatchFechasMontos(estado_og, auxbanco_og):
    estado, auxbanco = Cleaning(estado_og, auxbanco_og)
    auxbanco = auxbanco.copy()
    auxbanco["Used"] = False
    estadouno = FirstSearch(estado, auxbanco)
    estadodos = SecondSearch(estadouno, auxbanco)
    estadotres = ThirdSearch(estadodos, auxbanco)
    estadocuatro = FourthSearch(estadotres, auxbanco)
    estadofinal = FifthSearch(estadocuatro, auxbanco)
    return estadofinal

def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # File dialogs
    messagebox.showinfo("Info", "Select the Excel file to process.")
    estado, auxbanco = StartEstadoCuenta()
    updated_estado = MatchFechasMontos(estado, auxbanco)

    # Ask the user where to save the output
    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Processed Excel File"
    )

    if output_file:
        updated_estado.to_excel(output_file, index=False)
        messagebox.showinfo("Success", "File processed and saved successfully!")
    else:
        messagebox.showwarning("Warning", "No output file was selected. Exiting...")

if __name__ == "__main__":
    main()
