import streamlit as st
import pandas as pd
import os
from datetime import date

# Function to list all Excel files in the current directory
def list_excel_files():
    return [x for x in os.listdir() if x.split('.')[-1] in ['xlsx', 'xlsb', 'xls']]

# Function to process Excel files and aggregate data
def process_files():
    lista_pedidos = list_excel_files()
    pedidos_dict = {}

    for p in lista_pedidos:
        xl = pd.ExcelFile(p)
        hojas_con_datos = [h for h in xl.sheet_names if (h not in ['INDICE', 'TOTAL']) and (not h.startswith('Hoja'))]

        datos = pd.read_excel(p, sheet_name=hojas_con_datos, skiprows=5,
                              names=['cod', 'producto', 'cod barra', 'detalles',
                                     'minimo compra', 'stock', 'novedad / Fecha ingreso',
                                     'precio', 'cant pedido', 'subtotal'])

        df_agregado = pd.DataFrame(columns=['cod', 'producto', 'cod barra', 'detalles',
                                            'minimo compra', 'stock', 'novedad / Fecha ingreso',
                                            'precio', 'cant pedido', 'subtotal'])

        all_pedidos = []
        for h in hojas_con_datos:
            df_list = pd.DataFrame(datos[h]).dropna(how='all')
            pedidos = df_list.dropna(subset=['producto', 'cod barra']).iloc[1:, :].copy()
            all_pedidos.append(pedidos[~pedidos['cant pedido'].isna()])

        df_agregado = pd.concat(all_pedidos, ignore_index=True)
        df_agregado = df_agregado[['cod', 'producto', 'cod barra', 'minimo compra', 'stock', 'precio', 'cant pedido', 'subtotal']]

        disponible = df_agregado[(df_agregado['stock'] == 'DISPONIBLE') | (df_agregado['stock'] == 'BAJA DISPONIBILIDAD')]
        sin_stock = df_agregado[df_agregado['stock'] == 'SIN STOCK']

        totales = pd.read_excel(p, sheet_name='TOTAL', usecols=[0, 1], nrows=6, header=None, names=['campo', 'valor'])
        cliente = totales[totales['campo'] == 'Cliente:']['valor'].values[0]

        if pd.isnull(cliente) or cliente == '':
            cliente = 'S/N- ' + p[:10] + '...'

        pedidos_dict[cliente] = {'disponible': disponible, 'sin_stock': sin_stock, 'totales': totales}

    # Correct invalid characters in dictionary keys
    pedidos_dict_corr = {}
    for k in pedidos_dict.keys():
        k_new = k.replace('/', '-').replace('*', '-').replace('?', '-').replace('\\', '-')
        pedidos_dict_corr[k_new] = pedidos_dict[k]

    return pedidos_dict_corr

# Streamlit app
def main():
    st.title("Sales Notes Aggregator")

    st.write("Upload your Excel files to aggregate sales data.")
    uploaded_files = st.file_uploader("Upload Excel Files", accept_multiple_files=True, type=['xlsx', 'xlsb', 'xls'])

    if uploaded_files:
        for uploaded_file in uploaded_files:
            with open(uploaded_file.name, "wb") as f:
                f.write(uploaded_file.getbuffer())
        
        st.success("Files uploaded successfully!")
        pedidos_dict = process_files()

        for cliente, data in pedidos_dict.items():
            st.write(f"### Client: {cliente}")
            st.write("#### Available Stock")
            st.dataframe(data['disponible'])
            st.write("#### Out of Stock")
            st.dataframe(data['sin_stock'])
            st.write("#### Totals")
            st.dataframe(data['totales'])

if __name__ == "__main__":
    main()