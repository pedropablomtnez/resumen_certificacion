import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

def atexto(valor):
    texto = "{:,.3f}".format(valor)
    try:
        nuevo_valor = texto.replace(",", "$").replace(".", ",").replace("$", ".")
    except:
        nuevo_valor = texto.replace(".", ",")
    return nuevo_valor

def main():
    st.title("Extraer resumen de certificaciones")
    archivo = st.file_uploader("Cargar certificación", type="xlsx")
    nombre_archivo = ""
    output = BytesIO()
    if archivo is not None:
        #with st.form(key='formulario'):
        colum1, colum2, colum3, colum4 = st.columns(4, vertical_alignment='bottom')
        with colum1:
            num_col = st.text_input("Numero de certificación:")
        with colum2:
            num_fila = st.text_input("Numero ultima fila:")
        with colum3:
            boton_enviar = st.button(label="Ejecutar",
                                     type="primary")

        if boton_enviar and num_col and num_fila:
            df = pd.read_excel(archivo, sheet_name='2 PO & Payment Details', skiprows=3, nrows=int(num_fila),
                               usecols=[1, 3, int(num_col) + 9, 38])
            df_filtrado = df.groupby(['Area', 'Trade']).sum()
            ndf = df_filtrado.style.format(precision=3, thousands=".", decimal=",")
            st.table(ndf)
            #st.data_editor(ndf)
            nombre_columna = df_filtrado.columns.values[0]
            suma = df_filtrado[nombre_columna].sum()
            suma = atexto(suma)
            col1, col2 = st.columns([0.7, 0.3], vertical_alignment="center")
            with col1:
                st.subheader('Total certificación:')
            with col2:
                st.subheader(f'{suma}')
            pa = archivo.name.split('_')[1]
            nombre_archivo = f'Resumen {pa}.xlsx'
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer)
            output.seek(0)
            st.download_button(label='Descargar en Excel', data=output.getvalue(),
                               file_name=nombre_archivo,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               type="primary")


if __name__ == '__main__':
    main()
