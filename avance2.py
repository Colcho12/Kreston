import pandas as pd
from datetime import timedelta
from itertools import combinations
import time  # To implement the timeout


"""Abre el archivo de Excel, y las hojas relacionadas con el Estado de Cuenta"""
def StartEstadoCuenta():
    file_path = "copiacolchi.xlsx"
    estado = pd.read_excel(file_path, sheet_name="ESTADO DE CUENTA")
    auxbanco = pd.read_excel(file_path, sheet_name="AUX BANCOS")
    return estado, auxbanco

def Cleaning(estado,auxbanco):
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
    estadofinal = estado.copy()
    estadofinal["DOCUMENT NUMBER"] = document_numbers
    return estadofinal

def SecondSearch(estado, auxbanco):
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    
    unique_amounts = auxbanco["Amount in doc. curr."].value_counts()
    
    unique_amounts = unique_amounts[unique_amounts == 1].index  # Only keep unique values
    for index, row in unmatched_rows.iterrows():
        # Search for a match where Amount matches Amount in doc. curr.
        match = auxbanco[
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (auxbanco["Amount in doc. curr."].isin(unique_amounts))  # Only consider unused matches
        ].head(1)

        if not match.empty:
            # Assign the Document Number and mark it as used
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
    return estado

def ThirdSearch(estado, auxbanco):

    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
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
        else:
            estado.loc[index, "DOCUMENT NUMBER"] = None

    return estado

def FourthSearch(estado, auxbanco):
    """
    Perform a search for unmatched rows where:
    - The absolute amount in "Estado de Cuenta" matches the auxiliary amount.
    - The sign is flipped (e.g., positive -> negative or vice versa).
    """
    # Identify rows where DOCUMENT NUMBER is still NA
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    print(f"Number of unmatched rows before FourthSearch: {unmatched_rows.shape[0]}")  # Debugging

    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (auxbanco["Amount in doc. curr."] == -row["Amount"]) &  # Flip the sign
            (~auxbanco["Used"])  # Ensure the row is not already used
        ].head(1)

        if not match.empty:
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
        else:
            # Leave as NA if no match is found
            estado.loc[index, "DOCUMENT NUMBER"] = None

    print(f"Number of unmatched rows after FourthSearch: {estado['DOCUMENT NUMBER'].isna().sum()}")
    return estado


def subset_sum_dp(values, target, max_time):
    """
    Dynamic programming approach to find a subset of values that sums to the target.
    Includes a timeout to prevent excessive runtimes.
    """
    start_time = time.time()
    n = len(values)
    dp = {0: []}  # Base case: sum 0 achieved with an empty subset

    for i, value in enumerate(values):
        # Check for timeout
        if time.time() - start_time > max_time:
            return None  # Timeout: skip this target

        new_dp = dp.copy()
        for s in dp:
            new_sum = s + value
            if abs(new_sum) > abs(target):  # Prune sums exceeding the target
                continue
            if new_sum == target:
                return dp[s] + [i]  # Return the indices of the subset
            new_dp[new_sum] = dp[s] + [i]
        dp = new_dp

    return None  # No valid subset found

def FifthSearch(estado, auxbanco, max_days=10, max_time=12):
    """
    Optimized search for large subsets where:
    - The sum of charges (over several days) matches the target amount.
    - Groups candidates by day to reduce search space.
    - Skips targets after exceeding the allowed time.
    """
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    print(f"Number of unmatched rows before FifthSearch: {unmatched_rows.shape[0]}")

    for index, row in unmatched_rows.iterrows():
        target_amount = row["Amount"]
        target_date = row["FECHA"]

        # Start timing for this target
        start_time = time.time()

        # Expand date range dynamically
        start_date = target_date - timedelta(days=max_days)
        end_date = target_date + timedelta(days=max_days)

        # Filter candidates in auxbanco within the date range
        candidates = auxbanco[
            (auxbanco["Posting Date"] >= start_date) &
            (auxbanco["Posting Date"] <= end_date) &
            (~auxbanco["Used"])
        ]

        if candidates.empty:
            continue

        # Group candidates by day
        grouped_candidates = candidates.groupby("Posting Date")["Amount in doc. curr."].apply(list)
        candidate_indices = candidates.index.tolist()

        # Flatten the candidate list while keeping indices
        values = []
        indices_map = []
        for date, amounts in grouped_candidates.items():
            for i, amount in enumerate(amounts):
                values.append(amount)
                indices_map.append(candidate_indices.pop(0))  # Map to original indices

        # Use DP to find the subset with timeout
        subset_indices = subset_sum_dp(values, target_amount, max_time)

        # Check for timeout or failure
        if subset_indices:
            matched_indices = [indices_map[i] for i in subset_indices]
            document_number = auxbanco.loc[matched_indices[0], "Document Number"]

            # Mark rows in auxbanco as used
            auxbanco.loc[matched_indices, "Used"] = True

            # Assign DOCUMENT NUMBER to all matched rows in estado
            for i in subset_indices:
                matched_value = values[i]
                estado.loc[
                    (estado["DOCUMENT NUMBER"].isna()) &
                    (estado["Amount"] == matched_value) &
                    (estado["FECHA"] >= start_date) &
                    (estado["FECHA"] <= end_date),
                    "DOCUMENT NUMBER"
                ] = document_number
        else:
            print(f"Skipping target {target_amount} due to timeout.")

        # Check if we've spent too long on this target
        if time.time() - start_time > max_time:
            print(f"Skipping target {target_amount} after {max_time} seconds.")
            continue  # Skip to the next row

    print(f"Number of unmatched rows after FifthSearch: {estado['DOCUMENT NUMBER'].isna().sum()}")
    return estado

def MatchFechasMontos(estado_og, auxbanco_og):

    # LIMPIEZA DE DATOS
    estado,auxbanco = Cleaning(estado_og,auxbanco_og)

    auxbanco = auxbanco.copy()
    auxbanco["Used"] = False 

    estadouno = FirstSearch(estado, auxbanco)
    estadodos = SecondSearch(estadouno, auxbanco)
    estadotres = ThirdSearch(estadodos, auxbanco)
    estadocuatro = FourthSearch(estadotres, auxbanco)
    # estadofinal = FifthSearch(estadocuatro, auxbanco)  # Sum search (final)

    return estadocuatro


def main():
    #1. SE ABRE EL EXCEL Y LAS HOJAS PARA ARMAR EL ESTADO DE CUENTA
    estadocrudo, auxbancocrudo = StartEstadoCuenta()

    #2. SE CONSTRUYE ESTADO DE CUENTA
    updated_estado = MatchFechasMontos(estadocrudo, auxbancocrudo)
    updated_estado.to_excel("updated_estado_de_cuenta2.xlsx", index=False, sheet_name="Updated Estado")
    print("File successfully saved as 'updated_estado_de_cuenta.xlsx'")
main()