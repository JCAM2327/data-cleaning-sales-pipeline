# Data Cleaning and Reporting Pipeline for Sales Transactions
Automatización completa de limpieza y análisis de datos transaccionales en Excel con Python.

Este proyecto automatiza la limpieza, validación y análisis de una base de datos de transacciones en Excel, generando reportes y visualizaciones útiles para la toma de decisiones y la mejora continua de la calidad de datos.

## Características principales

- Limpieza automática de valores erróneos (`ERROR`, `UNKNOWN`).
- Corrección de formatos numéricos y de fechas.
- Imputación de valores faltantes.
- Eliminación y reporte de duplicados.
- Validación de rangos aceptables para precios y cantidades.
- Estandarización de métodos de pago y productos.
- Generación automática de gráficos (histogramas y barras).
- Exportación de resultados limpios a Excel y generación de informe PDF profesional.
- Sección de recomendaciones y conclusiones personalizadas en el informe.

## Instalación

Instala las dependencias necesarias ejecutando:

```bash
pip install pandas matplotlib fpdf
```

## Uso

1. Ubica tu archivo `.xlsx` con datos sucios en el directorio raíz del proyecto.
2. Ejecuta el script `inicio.py`.
3. Se generarán:
   - Un archivo Excel limpio (`*_CLEANED.xlsx`)
   - Gráficos en PNG
   - Un informe PDF con resumen, advertencias, recomendaciones y visualizaciones

## Ejemplo de uso

```python
from inicio import clean_transaction_data, generar_informe_pdf

df, stats = clean_transaction_data("db_for_cleanning.xlsx")
generar_informe_pdf(df, stats)
```
Esto generará el archivo limpio, los gráficos y el informe PDF en la misma carpeta.

Puedes ver un [ejemplo del informe PDF generado aquí](informe_analisis.pdf).

## Motivación

Este proyecto fue desarrollado para demostrar habilidades en limpieza y análisis de datos, automatización de reportes y buenas prácticas en la gestión de calidad de datos. Es ideal para incluir en tu CV o portafolio de GitHub como ejemplo de proyectos de Data Analytics aplicados a problemas reales.

## Requisitos

- Python 3.7 o superior
- Compatible con Windows, macOS y Linux
- Excel para visualizar los archivos generados

## Próximos pasos / Roadmap

- Añadir interfaz gráfica de usuario (GUI)
- Integrar validaciones automáticas al cargar datos
- Soporte para otros formatos de entrada (CSV, Google Sheets)
- Mejorar la personalización del informe PDF
- Añadir más visualizaciones y análisis exploratorio

## Consideraciones especiales

- Las fechas vacías en la columna `Transaction Date` se rellenan automáticamente con la **fecha promedio** de las transacciones existentes (o la fecha actual si no hay fechas válidas). Esto permite mantener la integridad del dato y facilita el análisis posterior (se sugiere futura revisión). El valor utilizado se reporta en el informe PDF.
- Si existen valores faltantes en `Price Per Unit`, `Quantity` o `Total Spent`, el sistema realiza una **imputación cruzada**:
  - Si se conoce el `Item` y la `Quantity`, pero falta `Price Per Unit`, se utiliza el valor más frecuente o promedio de ese producto.
  - Si se dispone de `Price Per Unit` y `Total Spent`, pero falta `Quantity`, se calcula como `Total Spent / Price Per Unit`.
  - Si se completan estos valores, también se recalcula `Total Spent` si corresponde.
  - Esta lógica reduce la cantidad de datos faltantes y mejora la calidad del archivo final.
- El informe PDF incluye:
  - Calidad de datos inicial y post-limpieza.
  - Cambios realizados por columna.
  - Estadísticas descriptivas extendidas (media, mediana, percentiles, std).
  - Visualizaciones adicionales (boxplots, barras, serie temporal).
  - Análisis por categoría (ubicación, método de pago).
  - **Hallazgos clave** y **análisis comparativo** entre ubicaciones y métodos de pago.
  - **Recomendaciones estratégicas** orientadas a negocio y calidad de datos.
  - Conclusiones del analista.

## Recomendaciones para la mejora continua

- Estandariza los valores desde el origen (por ejemplo, métodos de pago y productos).
- Implementa validaciones automáticas al momento de la captura de datos.
- Agrega controles de calidad periódicos y alertas para valores atípicos.
- Documenta el flujo de limpieza y análisis para facilitar la transferencia de conocimiento.
- Automatiza la ejecución periódica del script.
- Añade pruebas unitarias para asegurar la robustez del proceso.
- Integra control de versiones para el código y los datos limpios.

## Tecnologías utilizadas

- Python (pandas, matplotlib, fpdf)
- Excel

## Licencia

Este proyecto está bajo la licencia MIT. Puedes usarlo, modificarlo y compartirlo libremente.

---

Autor: José Mondragón  
Contacto: [linkedin.com/in/jose-mondragon-pylq](https://linkedin.com/in/jose-mondragon-pylq)

---

¿Tienes alguna sugerencia o recomendación para mejorar este proyecto?  
¡Estaré muy agradecido si me lo haces saber! Tu feedback es bienvenido y me ayuda a seguir creciendo profesionalmente.

---

🎯 *Este proyecto fue creado como muestra profesional de mis habilidades en Data Cleaning, Exploratory Data Analysis (EDA) y automatización de reportes con Python, ideal para roles como Data Analyst o Data Scientist.*
