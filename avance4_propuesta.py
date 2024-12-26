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

## ABSOLUTE VALUE
def FourthSearch(estado, auxbanco):
    """
    Verifica si arriba o abajo hay una fila con el mismo monto (signo opuesto) y misma fecha.
    Si se cumple, asigna el mismo DOCUMENT NUMBER a la fila actual.
    """
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    print(f"Number of unmatched rows before FourthSearch: {unmatched_rows.shape[0]}")

    for index, row in unmatched_rows.iterrows():
        # Verificar fila anterior (arriba)
        prev_index = index - 1
        if prev_index in estado.index:
            prev_row = estado.loc[prev_index]
            if (
                prev_row["FECHA"] == row["FECHA"] and  
                prev_row["Amount"] == -row["Amount"] and  # Signo opuesto
                pd.notna(prev_row["DOCUMENT NUMBER"])  
            ):
                estado.loc[index, "DOCUMENT NUMBER"] = prev_row["DOCUMENT NUMBER"]
                continue  # Pasar a la siguiente fila después de asignar

        # Verificar fila siguiente (abajo)
        next_index = index + 1
        if next_index in estado.index:
            next_row = estado.loc[next_index]
            if (
                next_row["FECHA"] == row["FECHA"] and 
                next_row["Amount"] == -row["Amount"] and  # Signo opuesto
                pd.notna(next_row["DOCUMENT NUMBER"])  
            ):
                estado.loc[index, "DOCUMENT NUMBER"] = next_row["DOCUMENT NUMBER"]

    print(f"Number of unmatched rows after FourthSearch: {estado['DOCUMENT NUMBER'].isna().sum()}")
    return estado


from datetime import timedelta

def find_consecutive_sum(values, target):
    """
    Find consecutive values (with at most one skip) that sum to the target.
    Returns the indices of the matching values if found, otherwise None.
    """
    n = len(values)
    for start in range(n):
        current_sum = 0
        skipped = False  # Allow one skip
        skip_index = -1  # Track the skipped index
        for end in range(start, n):
            current_sum += values[end]

            # Allow skipping one value if the sum overshoots
            if not skipped and abs(current_sum - target) > abs(target):
                skipped = True
                skip_index = end
                current_sum -= values[end]  # Remove the skipped value
                continue

            if current_sum == target:
                if skipped:
                    return list(range(start, skip_index)) + list(range(skip_index + 1, end + 1))
                return list(range(start, end + 1))
    return None


def FifthSearch(estado, auxbanco, max_days=10):
    """
    Search for rows in Auxiliar de Banco (targets) that match consecutive sums
    in Estado de Cuenta (candidates) within a date range.
    """

    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    print(f"Number of unmatched rows before FifthSearch: {unmatched_rows.shape[0]}")
    unmatched_aux_rows = auxbanco[~auxbanco["Used"]]
    print(f"Number of unmatched targets before FifthSearch: {unmatched_aux_rows.shape[0]}")

    for aux_index, aux_row in unmatched_aux_rows.iterrows():
        target_amount = aux_row["Amount in doc. curr."]
        target_date = aux_row["Posting Date"]
        document_number = aux_row["Document Number"]

        # Filter candidates in Estado de Cuenta within the date range
        candidates = estado[
            (estado["FECHA"] >= target_date - timedelta(days=max_days)) &
            (estado["FECHA"] <= target_date + timedelta(days=max_days)) &
            (estado["DOCUMENT NUMBER"].isna())  # Only unmatched rows
        ]

        if candidates.empty:
            continue

        # Extract values and indices
        candidate_values = candidates["Amount"].tolist()
        candidate_indices = candidates.index.tolist()

        # Find consecutive sums that match the target
        subset_indices = find_consecutive_sum(candidate_values, target_amount)
        if subset_indices:
            matched_indices = [candidate_indices[i] for i in subset_indices]

            # Assign the DOCUMENT NUMBER to all matched rows
            estado.loc[matched_indices, "DOCUMENT NUMBER"] = document_number

            # Mark the target row as used in Auxiliar de Banco
            auxbanco.at[aux_index, "Used"] = True

    print(f"Number of unmatched targets after FifthSearch: {auxbanco[~auxbanco['Used']].shape[0]}")
    print(f"Number of unmatched rows after FifthSearch: {estado['DOCUMENT NUMBER'].isna().sum()}")

    return estado


def SixthSearch(estado, auxbanco, tolerance=1, max_comb_size=20, max_iterations=1000):
    """
    1. Agrupa filas consecutivas con la misma DESCRIPCION y FECHA.
    2. Busca combinaciones de montos que sumen un valor objetivo con tolerancia.
    """
    unmatched_rows = estado[estado["DOCUMENT NUMBER"].isna()]
    aux_unmatched = auxbanco[(~auxbanco["Used"]) & (auxbanco["Amount in doc. curr."].abs() > 0.01)]

    print(f"Unmatched rows before SixthSearch: {len(unmatched_rows)}")
    print(f"Unmatched targets before SixthSearch: {len(aux_unmatched)}")

    # Agrupar por DESCRIPCION
    groups = ((unmatched_rows["DESCRIPCIÓN"] != unmatched_rows["DESCRIPCIÓN"].shift())).cumsum()
    grouped_rows = unmatched_rows.groupby(groups)

    iteration_count = 0 

    for group_id, group in grouped_rows:
        descripcion = group["DESCRIPCIÓN"].iloc[0]
        fecha = group["FECHA"].iloc[0]
        amounts = group["Amount"].tolist()
        indices = group.index.tolist()

        for aux_index, aux_row in aux_unmatched.iterrows():
            target_amount = aux_row["Amount in doc. curr."]
            document_number = aux_row["Document Number"]

            # Probar combinaciones de tamaño 1 hasta max_comb_size
            for r in range(1, min(len(amounts), max_comb_size) + 1):
                for subset in combinations(range(len(amounts)), r):
                    iteration_count += 1
                    if iteration_count >= max_iterations:
                        print(f"Search paused after {iteration_count} iterations.")
                        return estado  # Devuelve el estado hasta el punto actual

                    subset_sum = sum(amounts[i] for i in subset)
                    if abs(subset_sum - target_amount) <= tolerance:
                        matched_indices = [indices[i] for i in subset]

                        # Asignar el DOCUMENT NUMBER
                        estado.loc[matched_indices, "DOCUMENT NUMBER"] = document_number
                        auxbanco.loc[aux_index, "Used"] = True
                        print(f"Asignado {document_number} a filas {matched_indices} "
                              f"con suma {subset_sum} (target: {target_amount}, tol: {tolerance}).")

                        # Eliminar filas asignadas del grupo
                        for i in sorted(subset, reverse=True):
                            amounts.pop(i)
                            indices.pop(i)
                        break
                else:
                    continue
                break

    print(f"Unmatched rows after SixthSearch: {estado['DOCUMENT NUMBER'].isna().sum()}")
    return estado




def MatchFechasMontos(estado_og, auxbanco_og):
    # LIMPIEZA DE DATOS
    estado, auxbanco = Cleaning(estado_og, auxbanco_og)
    auxbanco = auxbanco.copy()
    auxbanco["Used"] = False 

    estadouno = FirstSearch(estado, auxbanco)
    estadodos = SecondSearch(estadouno, auxbanco)
    estadotres = ThirdSearch(estadodos, auxbanco)
    estadocuatro = FourthSearch(estadotres, auxbanco)
    estadocinco = FifthSearch(estadocuatro, auxbanco)
    estadofinal = SixthSearch(estadocinco, auxbanco)
    
    return estadofinal


def main():
    #1. SE ABRE EL EXCEL Y LAS HOJAS PARA ARMAR EL ESTADO DE CUENTA
    estadocrudo, auxbancocrudo = StartEstadoCuenta()

    #2. SE CONSTRUYE ESTADO DE CUENTA
    updated_estado = MatchFechasMontos(estadocrudo, auxbancocrudo)
    updated_estado.to_excel("updated_estado_de_cuenta_goat.xlsx", index=False, sheet_name="Updated Estado")
    print("File successfully saved as 'updated_estado_de_cuenta.xlsx'")
main()