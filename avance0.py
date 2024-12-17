import pandas as pd

"""Abre el archivo de Excel, y las hojas relacionadas con el Estado de Cuenta"""
def StartEstadoCuenta():
    file_path = "copiacolchi.xlsx"
    estado = pd.read_excel(file_path, sheet_name="ESTADO DE CUENTA")
    auxbanco = pd.read_excel(file_path, sheet_name="AUX BANCOS")
    return estado, auxbanco

"""Recupera las fechas y montos del estado de cuenta"""
def FechasMontosEstado(estado):
    ## SE HACE LIMPIEZA DE DATOS
    estado["FECHA"] = pd.to_datetime(estado["FECHA"], errors="coerce")
    estado = estado[estado["FECHA"].notna()]

    ## SE VALIDA EL NÚMERO DE FECHAS CORRECTAS
    validas = estado.shape[0]

    ## SE RECUPERAN LOS DATAFRAMES DE MONTOS Y FECHAS
    infoestado = estado[["FECHA","CARGOS", "ABONOS"]]
    return infoestado, estado

def FechasMontosAux(auxbanco):
    auxbanco["Posting Date"] = pd.to_datetime(auxbanco["Document Date"], errors="coerce")
    infoaux = auxbanco[["Document Number", "Posting Date", "Amount in local currency"]]
    return infoaux, auxbanco
  

def MatchFechasMontos(estado_de_cuenta, auxiliar_bancos):
    """
    Match FECHA and MONTOS between Estado de Cuenta and Auxiliar de Bancos 
    to retrieve the corresponding Document Number.
    """
    estado = estado_de_cuenta.copy()
    estado["FECHA"] = pd.to_datetime(estado["FECHA"], errors="coerce").dt.date
    estado = estado[estado["FECHA"].notna()]
    auxiliar_bancos["Posting Date"] = pd.to_datetime(auxiliar_bancos["Posting Date"], errors="coerce").dt.date

    # Step 2: Standardize and round amounts to 2 decimals
    estado = estado.copy()
    estado["Amount"] = estado["ABONOS"].fillna(0) - estado["CARGOS"].fillna(0)
    estado["Amount"] = estado["Amount"].round(2)
    auxiliar_bancos["Amount in local currency"] = auxiliar_bancos["Amount in local currency"].round(2)

    # Step 3: Perform the merge
    merged_data = pd.merge(
        estado,
        auxiliar_bancos,
        left_on=["FECHA", "Amount"],   # Match cleaned FECHA and Amount
        right_on=["Posting Date", "Amount in local currency"],  # Match Posting Date and Amount
        how="left",
        suffixes=("", "_bancos")  # Prevent column name clashes
    )

    print(merged_data["Document Number Aux"].head())
    # Step 4: Assign the matched Document Number
    # estadofinal = estado.copy
    num_rows_to_fill = len(estado_de_cuenta.loc[3:])
    estado_de_cuenta.loc[3:, "DOCUMENT NUMBER"] = merged_data["Document Number Aux"].iloc[:num_rows_to_fill].values
    return estado_de_cuenta


def main():
    #1. SE ABRE EL EXCEL Y LAS HOJAS PARA ARMAR EL ESTADO DE CUENTA
    estadocrudo, auxbancocrudo = StartEstadoCuenta()

    #2. SE GUARDAN LAS FECHAS Y MONTOS DEL ESTADO DE CUENTA
    infoestado, estado = FechasMontosEstado(estadocrudo)

    #3. SE GUARDAN LAS FECHAS Y MONTOS DEL AUXILIAR BANCARIO
    infoaux, auxbanco = FechasMontosAux(auxbancocrudo)

    #4. SE LLENAN LOS DOCUMENT NUMBERS 
    estadofinal = MatchFechasMontos(estado, auxbanco)

    estadofinal.to_excel("updated_estado_de_cuenta.xlsx", index=False, sheet_name="Updated Estado")
    print("File successfully saved as 'updated_estado_de_cuenta.xlsx'")
main()