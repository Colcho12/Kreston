import pandas as pd
from datetime import timedelta
from itertools import combinations


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

    for index, row in unmatched_rows.iterrows():
        match = auxbanco[
            (auxbanco["Posting Date"] >= row["FECHA"] - timedelta(days=1)) &
            (auxbanco["Posting Date"] <= row["FECHA"] + timedelta(days=1)) &
            (auxbanco["Amount in doc. curr."] == row["Amount"]) &
            (~auxbanco["Used"])
        ].head(1)

        if not match.empty:
            # Assign the Document Number and mark it as used
            document_number = match["Document Number"].iloc[0]
            auxbanco.loc[match.index, "Used"] = True
            estado.loc[index, "DOCUMENT NUMBER"] = document_number
        else:
            # No match found
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


def MatchFechasMontos(estado_og, auxbanco_og):

    # LIMPIEZA DE DATOS
    estado,auxbanco = Cleaning(estado_og,auxbanco_og)

    auxbanco = auxbanco.copy()
    auxbanco["Used"] = False 

    estadouno = FirstSearch(estado, auxbanco)
    estadodos = SecondSearch(estadouno, auxbanco)
    estadotres = ThirdSearch(estadodos, auxbanco)
    estadofinal = FourthSearch(estadotres, auxbanco)
    return estadofinal


def main():
    #1. SE ABRE EL EXCEL Y LAS HOJAS PARA ARMAR EL ESTADO DE CUENTA
    estadocrudo, auxbancocrudo = StartEstadoCuenta()

    #2. SE CONSTRUYE ESTADO DE CUENTA
    updated_estado = MatchFechasMontos(estadocrudo, auxbancocrudo)
    updated_estado.to_excel("updated_estado_de_cuenta2.xlsx", index=False, sheet_name="Updated Estado")
    print("File successfully saved as 'updated_estado_de_cuenta.xlsx'")
main()