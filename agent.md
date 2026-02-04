# Descripcion del programa

Este programa valida datos entre un archivo Excel de protocolos de pruebas
(`protocolo.xlsm`, hoja `DATOS`) y una hoja de datos en PDF (`datos.pdf`).
El objetivo es detectar errores de digitacion comparando campos clave del
transformador entre ambos documentos.

La validacion compara, entre otros:
- C.P. (Excel) vs Codigo del producto (PDF)
- Tipo (CRHV3 -> Seco, OTHV3 -> Aceite)
- Tensiones primaria y secundaria (kV)
- Potencia (kVA)
- Grupo de conexion
- Po, Io, Pcc y Uz garantizados

El resultado se muestra como OK/NO y un resumen con la cantidad de coincidencias.

# Como usarlo

1. Interfaz grafica (tkinter):
   - Ejecuta `python validar_datos.py`
   - Selecciona el Excel y el PDF
   - Presiona "Validar"
   - El resultado aparece en la ventana de logs

2. Linea de comandos:
   - `python validar_datos.py --excel protocolo.xlsm --pdf datos.pdf --sheet DATOS`

# Resumen del codigo (validar_datos.py)

El script esta organizado en tres partes:

1) Extraccion y normalizacion de datos
- `norm_text`: normaliza texto (mayusculas, sin tildes).
- `find_excel_value` y `find_excel_value_in_col_a`: buscan valores en Excel.
- `parse_pdf_text`: usa `pdfplumber` para extraer texto y tablas del PDF.
- `find_pdf_group`, `find_pdf_kv_pair`, `find_impedance_value`: extraen campos
  del PDF con regex y heuristicas.

2) Comparacion
- `compare_text`, `compare_ids`, `compare_group`, `compare_numeric`: comparan
  valores y calculan diferencias con tolerancias.
- `run_validation`: coordina toda la lectura y genera la lista de resultados.
- `format_results`: formatea el reporte de validacion.

3) Interfaces
- `main_cli`: modo linea de comandos.
- `launch_gui`: interfaz grafica con tkinter (selector de archivos y log).

El programa detecta diferencias y muestra exactamente que dato no coincide
entre Excel y PDF.

# Explicacion de cada funcion

- `norm_text(value)`: normaliza texto (mayusculas, sin tildes y espacios
  repetidos) para comparar sin errores por formato.
- `extract_numbers(text)`: extrae numeros (con punto o coma decimal) desde un
  texto.
- `find_excel_value(ws, label_substring, col_value=2)`: busca una etiqueta en
  la columna A y devuelve el valor de la columna indicada.
- `find_excel_value_in_col_a(ws, regex_pattern, col_value=2)`: busca una
  etiqueta en la columna A usando regex y devuelve el valor de la columna
  indicada.
- `excel_tipo_to_pdf(tipo)`: convierte el tipo del Excel (CRHV3/OTHV3) al texto
  esperado en el PDF (SECO/ACEITE).
- `parse_pdf_text(pdf_path)`: extrae texto y contenido de tablas del PDF para
  facilitar las busquedas con regex.
- `find_pdf_group(text_norm)`: obtiene el grupo de conexion desde el texto
  normalizado del PDF.
- `find_pdf_kv_pair(text)`: detecta el par de tensiones `kV` (primario/secundario)
  en el PDF.
- `find_value_in_lines(text, key_norm)`: busca un valor numerico en la misma
  linea donde aparece una clave.
- `find_code_producto(text_norm)`: obtiene el codigo del producto desde el PDF.
- `find_impedance_value(text_norm)`: extrae la impedancia @ 120 C en [%] desde
  el bloque correspondiente del PDF.
- `normalize_id(value)`: limpia IDs numericos para compararlos sin el `.0`.
- `to_kv(value)`: convierte tensiones de V a kV cuando el valor es mayor a 100.
- `normalize_group(value)`: normaliza el grupo de conexion para comparar.
- `compare_text(field, excel_value, pdf_value)`: compara textos exactos.
- `compare_ids(field, excel_value, pdf_value)`: compara IDs numericos normalizados.
- `compare_group(field, excel_value, pdf_value)`: compara grupos de conexion.
- `compare_numeric(field, excel_value, pdf_value, unit, rel_tol, abs_tol)`:
  compara numeros con tolerancias.
- `run_validation(excel_path, pdf_path, sheet)`: orquesta la lectura de Excel y
  PDF, y genera la lista de resultados.
- `format_results(results)`: genera el reporte en texto con OK/NO y resumen.
- `main_cli()`: ejecuta la validacion en modo linea de comandos.
- `launch_gui()`: inicia la interfaz grafica con tkinter.
