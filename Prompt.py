import os
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

def select_folder():
    folder_selected = st.text_input("Ingrese la ruta de la carpeta de datos:")
    return folder_selected

def process_data(folder_path, col_range, start_row):
    try:
        # Obtener la lista de archivos Excel en la carpeta
        excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and f.startswith('AvanceVentasINTI')]

        if not excel_files:
            st.error("No se encontraron archivos Excel en la carpeta seleccionada.")
            return

        all_data = []

        for i, file in enumerate(excel_files):
            file_path = os.path.join(folder_path, file)
            
            # Extraer año, mes y día del nombre del archivo
            date_str = file.split('.')[1:4]
            year, month, day = date_str

            # Leer el archivo Excel
            df = pd.read_excel(file_path, sheet_name='ITEM_O', usecols=col_range, skiprows=start_row)

            # Añadir columnas de año, mes y día
            df['ANIO'] = year
            df['MES'] = month
            df['DIA'] = day

            all_data.append(df)

            # Actualizar la barra de progreso
            st.progress((i + 1) / len(excel_files))

        # Consolidar todos los DataFrames
        final_df = pd.concat(all_data, ignore_index=True)

        # Exportar a Excel
        output_path = os.path.join(folder_path, 'Out.xlsx')
        final_df.to_excel(output_path, index=False)

        st.success(f"Proceso completado. Archivo guardado en {output_path}")

        # Mostrar el DataFrame final
        st.dataframe(final_df)

        # Calcular y mostrar gráficos
        show_plots(final_df)

    except Exception as e:
        st.error(f"Ocurrió un error durante el proceso: {str(e)}")

def show_plots(df):
    # Seleccionar solo las columnas numéricas
    numeric_df = df.select_dtypes(include=['number'])

    # Calcular promedios por columna
    column_averages = numeric_df.mean()

    # Ordenar promedios de mayor a menor y seleccionar los N más altos
    top_n = 10  # Número de promedios más altos a mostrar
    top_averages = column_averages.sort_values(ascending=False).head(top_n)

    # Crear gráficos
    fig, axs = plt.subplots(2, 1, figsize=(10, 8))

    # Gráfico de barras
    top_averages.plot(kind='bar', ax=axs[0])
    axs[0].set_title('Top Promedios por Columna')
    axs[0].set_xlabel('Columnas')
    axs[0].set_ylabel('Promedio')

    # Gráfico de torta
    top_averages.plot(kind='pie', ax=axs[1], autopct='%1.1f%%')
    axs[1].set_title('Distribución de Top Promedios')

    st.pyplot(fig)

# Interfaz de Streamlit
st.title("Proceso ETL")

# Selección de carpeta
folder_path = select_folder()

# Rango de columnas
col_range = st.text_input("Rango de columnas (ej. A:D):")

# Fila inicial
start_row = st.number_input("Fila inicial:", min_value=1, step=1) - 1

# Botón de proceso
if st.button("Procesar"):
    process_data(folder_path, col_range, start_row)
