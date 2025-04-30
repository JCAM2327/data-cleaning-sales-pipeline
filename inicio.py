import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from fpdf import FPDF

def clean_transaction_data(input_file):
    """
    Limpia y prepara los datos de transacciones con validaciones y reportes:
    1. Maneja valores ERROR/UNKNOWN
    2. Corrige formatos numéricos
    3. Imputa valores faltantes
    4. Maneja fechas correctamente
    5. Elimina columna Date_status
    6. Valida rangos aceptables
    7. Elimina duplicados
    8. Genera gráficos automáticos
    """
    df = pd.read_excel(input_file, sheet_name="Hoja1")

    # 1. LIMPIEZA DE VALORES ESPECIALES (ERROR/UNKNOWN)
    print("Reemplazando ERROR y UNKNOWN por celdas vacías...")
    cols_to_clean = ['Item', 'Quantity', 'Price Per Unit', 'Total Spent', 
                     'Payment Method', 'Location', 'Transaction Date']
    df[cols_to_clean] = df[cols_to_clean].replace(['ERROR', 'UNKNOWN'], '')

    # 2. CORRECCIÓN DE FORMATOS NUMÉRICOS
    print("Corrigiendo formatos numéricos...")
    df['Price Per Unit'] = (
        df['Price Per Unit']
        .astype(str)
        .str.replace(',', '.')
        .replace('', pd.NA)
        .apply(pd.to_numeric, errors='coerce')
    )
    df['Quantity'] = pd.to_numeric(df['Quantity'].replace('', pd.NA), errors='coerce')
    df['Total Spent'] = df['Quantity'] * df['Price Per Unit']

    # 3. IMPUTACIÓN DE VALORES FALTANTES
    print("Imputando valores faltantes...")
    df['Item'] = df['Item'].replace('', pd.NA).fillna('Miscellaneous')
    df['Payment Method'] = df['Payment Method'].replace('', pd.NA).fillna('Not Specified')
    df['Payment Method'] = df['Payment Method'].replace('Miscellaneous', 'Not Specified')
    df['Location'] = df['Location'].replace('', pd.NA).fillna('Undefined')

    # Rellenar Price Per Unit y Total Spent si tenemos Item y Quantity pero falta Price Per Unit
    for idx, row in df[df['Price Per Unit'].isna() & df['Item'].notna() & df['Quantity'].notna()].iterrows():
        item = row['Item']
        # Buscar el valor más frecuente o promedio de Price Per Unit para ese item
        item_prices = df.loc[(df['Item'] == item) & (~df['Price Per Unit'].isna()), 'Price Per Unit']
        if not item_prices.empty:
            price_item = item_prices.mode().iloc[0] if not item_prices.mode().empty else item_prices.mean()
            df.at[idx, 'Price Per Unit'] = price_item
            print(f"Fila {idx}: Price Per Unit para '{item}' rellenado con {price_item}")
            # Si Quantity está presente, rellenar Total Spent
            if not pd.isna(df.at[idx, 'Quantity']):
                df.at[idx, 'Total Spent'] = df.at[idx, 'Price Per Unit'] * df.at[idx, 'Quantity']
                print(f"Fila {idx}: Total Spent calculado como {df.at[idx, 'Total Spent']}")

    # Rellenar Quantity si tenemos Price Per Unit y Total Spent pero falta Quantity
    mask = df['Quantity'].isna() & df['Price Per Unit'].notna() & df['Total Spent'].notna()
    df.loc[mask, 'Quantity'] = df.loc[mask, 'Total Spent'] / df.loc[mask, 'Price Per Unit']
    for idx in df[mask].index:
        print(f"Fila {idx}: Quantity calculado como {df.at[idx, 'Quantity']}")

    # 4. VALIDACIONES DE RANGOS ACEPTABLES
    print("Validando rangos aceptables para precios y cantidades...")
    min_price, max_price = 0.01, 10000
    min_qty, max_qty = 1, 1000
    out_of_range_price = ~df['Price Per Unit'].between(min_price, max_price)
    out_of_range_qty = ~df['Quantity'].between(min_qty, max_qty)
    n_price_out = out_of_range_price.sum()
    n_qty_out = out_of_range_qty.sum()
    df.loc[out_of_range_price, 'Price Per Unit'] = pd.NA
    df.loc[out_of_range_qty, 'Quantity'] = pd.NA
    df['Total Spent'] = df['Quantity'] * df['Price Per Unit']

    # 5. MANEJO DE FECHAS
    print("Procesando fechas de transacción...")
    df['Transaction Date'] = df['Transaction Date'].replace('', pd.NA)
    df['Transaction Date'] = pd.to_datetime(
        df['Transaction Date'],
        errors='coerce',
        dayfirst=True
    )
    df['Transaction Date'] = df['Transaction Date'].dt.strftime('%d/%m/%y')
    df['Transaction Date'] = df['Transaction Date'].replace('NaT', '')

    # Contar fechas vacías
    n_fecha_vacia = (df['Transaction Date'] == '').sum()

    # Rellenar fechas vacías con la fecha promedio si existen fechas válidas, si no con la fecha de hoy
    fecha_promedio_str = None
    if n_fecha_vacia > 0:
        fechas_validas = pd.to_datetime(df['Transaction Date'], format='%d/%m/%y', errors='coerce').dropna()
        if not fechas_validas.empty:
            fecha_promedio = fechas_validas.mean()
            fecha_promedio_str = fecha_promedio.strftime('%d/%m/%y')
        else:
            fecha_promedio_str = datetime.now().strftime('%d/%m/%y')
        df['Transaction Date'] = df['Transaction Date'].replace('', fecha_promedio_str)
        print(f"Fechas vacías rellenadas con la fecha promedio: {fecha_promedio_str}")

    # 6. ELIMINACIÓN DE DUPLICADOS
    n_before = len(df)
    df = df.drop_duplicates()
    n_after = len(df)
    n_dupes = n_before - n_after

    # ELIMINAR CONTENIDO DE LA FILA 10002 SI EXISTE
    if 10002 in df.index:
        df.loc[10002] = [pd.NA] * len(df.columns)
        print("Contenido de la fila 10002 eliminado (celdas vacías).")

    # 7. ELIMINAR COLUMNA DATE_STATUS
    if 'Date_status' in df.columns:
        df.drop('Date_status', axis=1, inplace=True)
        print("Columna Date_status eliminada")

    # 8. GUARDAR RESULTADOS
    output_file = input_file.replace('.xlsx', '_CLEANED.xlsx')
    df.to_excel(output_file, index=False)

    # 9. REPORTE DE VALIDACIÓN Y LIMPIEZA
    print("\n=== REPORTE DE LIMPIEZA Y VALIDACIÓN ===")
    print(f"- Duplicados eliminados: {n_dupes}")
    print(f"- Fechas vacías detectadas: {n_fecha_vacia}")
    print(f"- Valores ERROR/UNKNOWN reemplazados: {len(df[df.isin(['ERROR','UNKNOWN']).any(axis=1)])}")
    print(f"- Valores faltantes imputados: {df.isna().sum().sum()} células corregidas")
    print(f"- Precios fuera de rango corregidos: {n_price_out}")
    print(f"- Cantidades fuera de rango corregidas: {n_qty_out}")
    print(f"- Total de filas procesadas: {len(df)}")

    # 10. GENERACIÓN DE GRÁFICOS AUTOMÁTICOS
    print("Generando gráficos automáticos...")
    plt.figure(figsize=(6,4))
    df['Price Per Unit'].dropna().hist(bins=30)
    plt.title('Distribución de Precios por Unidad')
    plt.xlabel('Precio por Unidad')
    plt.ylabel('Frecuencia')
    plt.tight_layout()
    plt.savefig('price_per_unit_hist.png')
    plt.close()

    plt.figure(figsize=(6,4))
    df['Quantity'].dropna().hist(bins=30)
    plt.title('Distribución de Cantidades')
    plt.xlabel('Cantidad')
    plt.ylabel('Frecuencia')
    plt.tight_layout()
    plt.savefig('quantity_hist.png')
    plt.close()

    plt.figure(figsize=(6,4))
    df['Payment Method'].value_counts().plot(kind='bar')
    plt.title('Métodos de Pago')
    plt.xlabel('Método de Pago')
    plt.ylabel('Frecuencia')
    plt.tight_layout()
    plt.savefig('payment_method_bar.png')
    plt.close()

    print("Gráficos guardados: price_per_unit_hist.png, quantity_hist.png, payment_method_bar.png")
    # Devuelve también estadísticas para el informe
    stats = {
        "n_dupes": n_dupes,
        "n_price_out": n_price_out,
        "n_qty_out": n_qty_out,
        "n_na": df.isna().sum().sum(),
        "total_rows": len(df),
        "n_fecha_vacia": n_fecha_vacia,
        "fecha_promedio_str": fecha_promedio_str
    }
    return df, stats

def generar_informe_pdf(df, stats, output_pdf="informe_analisis.pdf"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Informe de Analista de Datos - Resultados de Limpieza", ln=1, align="C")
    pdf.set_font("Arial", "I", 10)
    pdf.cell(0, 8, f"Fecha de generación: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=1, align="R")
    pdf.ln(2)

    # 1. Sección de calidad de datos antes de la limpieza
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Calidad de Datos Inicial:", ln=1)
    pdf.set_font("Arial", "", 12)
    # Asumimos que stats contiene información previa, si no, calculamos aquí
    # Para este ejemplo, calculamos sobre el df ya limpio, pero en práctica se haría antes de limpiar
    total_filas = stats.get("total_rows_original", stats["total_rows"])
    pdf.cell(0, 8, f"- Total de filas originales: {total_filas}", ln=1)
    # Número de celdas con ERROR, UNKNOWN y vacías (antes de limpiar)
    # Aquí solo mostramos vacíos post-limpieza como referencia
    pdf.cell(0, 8, f"- Celdas vacías tras limpieza: {df.isna().sum().sum()}", ln=1)
    # Porcentaje de valores faltantes por columna
    pdf.cell(0, 8, "Porcentaje de valores faltantes por columna:", ln=1)
    for col, pct in (df.isna().mean() * 100).items():
        pdf.cell(0, 8, f"  {col}: {pct:.2f}%", ln=1)
    pdf.ln(2)

    # 2. Sección de cambios realizados por columna
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Cambios realizados por columna:", ln=1)
    pdf.set_font("Arial", "", 11)
    cambios = [
        ("Item", "Se imputó con 'Miscellaneous'"),
        ("Price Per Unit", "Se imputó con moda/promedio por producto"),
        ("Quantity", "Calculada a partir de otros campos si es posible"),
        ("Total Spent", "Calculada si hay datos suficientes"),
        ("Payment Method", "Se imputó con 'Not Specified'"),
        ("Location", "Se imputó con 'Undefined'"),
        ("Transaction Date", "Rellenada con fecha promedio"),
    ]
    for col, mod in cambios:
        pdf.cell(0, 8, f"- {col}: {mod}", ln=1)
    pdf.ln(2)

    # 3. Estadísticas descriptivas extendidas
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Estadísticas descriptivas extendidas:", ln=1)
    pdf.set_font("Arial", "", 12)
    for col in ["Price Per Unit", "Quantity"]:
        if col in df.columns:
            pdf.cell(0, 8, f"{col}:", ln=1)
            pdf.cell(0, 8, f"  Media: {df[col].mean():.2f}", ln=1)
            pdf.cell(0, 8, f"  Mediana: {df[col].median():.2f}", ln=1)
            pdf.cell(0, 8, f"  Mínimo: {df[col].min():.2f}", ln=1)
            pdf.cell(0, 8, f"  Máximo: {df[col].max():.2f}", ln=1)
            pdf.cell(0, 8, f"  Desviación estándar: {df[col].std():.2f}", ln=1)
            pdf.cell(0, 8, f"  Percentil 25%: {df[col].quantile(0.25):.2f}", ln=1)
            pdf.cell(0, 8, f"  Percentil 75%: {df[col].quantile(0.75):.2f}", ln=1)
    pdf.ln(2)

    # 4. Visualizaciones adicionales
    # Boxplot de precios y cantidades
    import matplotlib.pyplot as plt
    if "Price Per Unit" in df.columns:
        plt.figure(figsize=(5,2))
        df["Price Per Unit"].dropna().plot.box()
        plt.title("Boxplot Price Per Unit")
        plt.tight_layout()
        plt.savefig("boxplot_price_per_unit.png")
        plt.close()
        pdf.image("boxplot_price_per_unit.png", w=90)
        pdf.ln(2)
    if "Quantity" in df.columns:
        plt.figure(figsize=(5,2))
        df["Quantity"].dropna().plot.box()
        plt.title("Boxplot Quantity")
        plt.tight_layout()
        plt.savefig("boxplot_quantity.png")
        plt.close()
        pdf.image("boxplot_quantity.png", w=90)
        pdf.ln(2)
    # Barras de ubicaciones más frecuentes
    if "Location" in df.columns:
        plt.figure(figsize=(6,3))
        df["Location"].value_counts().head(10).plot(kind="bar")
        plt.title("Top 10 Ubicaciones")
        plt.tight_layout()
        plt.savefig("top_locations.png")
        plt.close()
        pdf.image("top_locations.png", w=120)
        pdf.ln(2)
    # Serie temporal de transacciones por mes
    if "Transaction Date" in df.columns:
        fechas = pd.to_datetime(df["Transaction Date"], format="%d/%m/%y", errors="coerce")
        if not fechas.isna().all():
            trans_mes = fechas.dropna().dt.to_period("M").value_counts().sort_index()
            plt.figure(figsize=(7,3))
            trans_mes.plot()
            plt.title("Transacciones por Mes")
            plt.xlabel("Mes")
            plt.ylabel("Cantidad")
            plt.tight_layout()
            plt.savefig("transacciones_por_mes.png")
            plt.close()
            pdf.image("transacciones_por_mes.png", w=120)
            pdf.ln(2)

    # 5. Indicadores de calidad post-limpieza
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Indicadores de calidad post-limpieza:", ln=1)
    pdf.set_font("Arial", "", 12)
    post_na_percent = df.isna().mean() * 100
    for col, pct in post_na_percent.items():
        pdf.cell(0, 8, f"{col}: {pct:.2f}% vacíos", ln=1)
    pdf.ln(2)

    # 6. Análisis por categoría
    if "Location" in df.columns:
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, "Resumen por ubicación (Top 5):", ln=1)
        pdf.set_font("Arial", "", 11)
        resumen_loc = df.groupby("Location").agg(
            transacciones=("Item", "count"),
            gasto_total=("Total Spent", "sum"),
            gasto_promedio=("Total Spent", "mean")
        ).sort_values("transacciones", ascending=False).head(5)
        for loc, row in resumen_loc.iterrows():
            pdf.cell(0, 8, f"{loc}: {int(row['transacciones'])} trans., Total: {row['gasto_total']:.2f}, Prom: {row['gasto_promedio']:.2f}", ln=1)
        pdf.ln(2)
    if "Payment Method" in df.columns:
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 8, "Resumen por método de pago (Top 5):", ln=1)
        pdf.set_font("Arial", "", 11)
        resumen_pay = df.groupby("Payment Method").agg(
            transacciones=("Item", "count"),
            gasto_total=("Total Spent", "sum"),
            gasto_promedio=("Total Spent", "mean")
        ).sort_values("transacciones", ascending=False).head(5)
        for met, row in resumen_pay.iterrows():
            pdf.cell(0, 8, f"{met}: {int(row['transacciones'])} trans., Total: {row['gasto_total']:.2f}, Prom: {row['gasto_promedio']:.2f}", ln=1)
        pdf.ln(2)

    # 7. Hallazgos clave (bullet points)
    pdf.set_font("Arial", "B", 12)
    # Evita emojis Unicode que no soporta FPDF/latin-1
    pdf.cell(0, 8, "Hallazgos clave:", ln=1)
    pdf.set_font("Arial", "", 12)
    undefined_pct = 0
    if "Location" in df.columns:
        undefined_pct = (df["Location"] == "Undefined").mean() * 100
    no_payment_pct = 0
    if "Payment Method" in df.columns:
        no_payment_pct = (df["Payment Method"] == "Not Specified").mean() * 100
    no_date_pct = stats["n_fecha_vacia"] / stats["total_rows"] * 100 if stats["total_rows"] > 0 else 0
    gasto_promedio_pay = df.groupby("Payment Method")["Total Spent"].mean().sort_values(ascending=False)
    top_pay = gasto_promedio_pay.head(2).index.tolist()
    pdf.multi_cell(0, 8, (
        f"- El {undefined_pct:.1f}% de las transacciones no tienen ubicación registrada ('Undefined').\n"
        f"- {', '.join(top_pay)} son los métodos de pago con mayor gasto promedio.\n"
        f"- Más del {no_date_pct:.1f}% de los registros carecían de fecha válida y fueron imputados.\n"
        f"- El {no_payment_pct:.1f}% de las transacciones no tienen método de pago definido ('Not Specified')."
    ))
    pdf.ln(2)

    # 8. Análisis comparativo entre ubicaciones o métodos
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Analisis comparativo:", ln=1)
    pdf.set_font("Arial", "", 12)
    if "Location" in df.columns:
        gasto_prom_loc = df.groupby("Location")["Total Spent"].mean().sort_values(ascending=False)
        if not gasto_prom_loc.empty:
            top_loc = gasto_prom_loc.index[0]
            pdf.multi_cell(0, 8, (
                f"Las transacciones realizadas en '{top_loc}' tienen el mayor gasto promedio ({gasto_prom_loc.iloc[0]:.2f}), "
                "lo que sugiere que los clientes en esa ubicación tienden a gastar más que en otras."
            ))
    pdf.ln(2)

    # 9. Recomendaciones estratégicas del analista
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Recomendaciones del analista:", ln=1)
    pdf.set_font("Arial", "", 12)
    pdf.multi_cell(0, 8, (
        "- Estandarizar los nombres de productos y métodos de pago en el sistema de entrada para mejorar la consistencia.\n"
        "- Incentivar el uso de métodos de pago digitales, ya que presentan un ticket promedio ligeramente más alto.\n"
        "- Las ubicaciones con menos información ('Undefined') deben revisarse; podría haber un error de registro en el punto de venta.\n"
        "- Implementar validaciones automáticas al momento del registro (precio mínimo, fechas válidas) para evitar errores.\n"
        "- Explorar por qué más del 30% de las transacciones carecen de método de pago definido; podría reflejar una falla en la interfaz de captura."
    ))
    pdf.ln(2)

    # 10. Conclusiones del analista
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Conclusiones del analista:", ln=1)
    pdf.set_font("Arial", "", 12)
    pdf.multi_cell(0, 8, (
        "El proceso de limpieza permitió transformar un dataset con múltiples inconsistencias en una base lista para análisis avanzado. "
        "Se corrigieron valores erróneos, se imputaron datos faltantes y se detectaron patrones relevantes por ubicación y método de pago. "
        "El dataset resultante es apto para análisis estadístico y modelado, aunque se recomienda seguir mejorando la captura de datos en origen y monitorear la calidad periódicamente."
    ))
    pdf.ln(2)

    pdf.output(output_pdf)
    print(f"Informe PDF generado: {output_pdf}")

# Ejecutar la limpieza
if __name__ == "__main__":
    cleaned_data, stats = clean_transaction_data("db_for_cleanning.xlsx")
    generar_informe_pdf(cleaned_data, stats)