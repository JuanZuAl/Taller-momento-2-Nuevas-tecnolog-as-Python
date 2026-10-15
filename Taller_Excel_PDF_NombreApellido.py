from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from tabulate import tabulate
import openpyxl

## Cargar datos desde Excel 
wb = openpyxl.load_workbook("ventas_tecnologia_v2.xlsx", data_only=True)
hoja = wb.active  
print(f"Hoja cargada exitosamente: {hoja.title}")      
datos = []

for fila in hoja.iter_rows(min_row=2, values_only=True):
    if fila[3] > 100:  #  Reto 1: Filtro por precio unitario superior a $100
        datos.append(list(fila))

print (f"total de registros leidos en memoria: {len(datos)}")

## tabulate para mostrar datos en consola
encabezados = ["ID", "Producto", "Categoria", "Precio ($)", "Ventas (Unid)"]

alineacion =["Center", "Left", "Center", "Right", "Right"]

## Reto 2: Calcular métricas
total_unidades = sum(fila[4] for fila in datos)
precio_promedio = round(sum(fila[3] for fila in datos) / len(datos) if datos else 0, 2)
datos.append(["", "", "", "Total Unidades", f"{total_unidades:}"])
datos.append(["", "", "", f"Precio Promedio", f"${precio_promedio:}"])

print ("\n === REPORTE DE VENTAS EN CONSOLA (22 PRODUCTOS) ===")
print (tabulate(datos, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))


## Generar PDF con ReportLab
doc = SimpleDocTemplate("reporte_ejecutivo_ventas.pdf", pagesize=letter)
estilos = getSampleStyleSheet()
estilo_titulo = ParagraphStyle(
    "EstiloTitulo", parent = estilos["Title"],
    fontName = "Helvetica-Bold", fontSize = 18,
    textColor = colors.HexColor("#1A365D"), spaceAfter = 10
)
estilo_cuerpo = ParagraphStyle("EstiloCuerpo", parent = estilos["Normal"], fontSize = 9)
estilo_encabezado = ParagraphStyle("EstiloEncabezado", parent = estilos["Normal"],
                                   fontSize = 10, fontName = "Helvetica-Bold",
                                   textColor = colors.whitesmoke, alignment = 1)

story = [
    Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
    Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
    Spacer(1, 12)
]

encabezados_pdf =[Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]
filas_pdf = [
    [
        Paragraph(str(r[0]), ParagraphStyle("C", parent = estilo_cuerpo, alignment = 1)),
        Paragraph(str(r[1]), ParagraphStyle("L", parent = estilo_cuerpo, alignment = 0)),
        Paragraph(str(r[2]), ParagraphStyle("C", parent = estilo_cuerpo, alignment = 1)),
        Paragraph(str(r[3]), ParagraphStyle("R", parent = estilo_cuerpo, alignment = 2)),
        Paragraph(str(r[4]), ParagraphStyle("R", parent = estilo_cuerpo, alignment = 2))
    ]
    for r in datos
]

ancho_cols = [40, 160, 100, 80, 80]
tabla_pdf = Table([encabezados_pdf] + filas_pdf, colWidths = ancho_cols, repeatRows = 1)

tabla_pdf.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
    ("VALIGN" , (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
]))            

story.append(tabla_pdf)

doc.build(story)


# Reto 1 — Filtrado por Criterio de Precio: Modifica la lectura de datos para almacenar y presentar
# (tanto en la consola con tabulate como en el PDF) únicamente los productos cuyo precio unitario
# sea superior a $100.

# Reto 2 — Métricas y Fila de Resumen: Escribe el código necesario para calcular la suma total de
# unidades vendidas y el precio promedio de los productos. Muestra estas métricas al final de la
# consola y agrega una fila final resumida en la tabla del PDF.
# Reto 2: Calcular métricas

# Reto 3 — Búsqueda Dinámica por ID de Producto: Solicita al usuario un número de ID mediante la
# función input(). El programa debe buscar dicho ID en la matriz de datos y mostrar únicamente la fila
# correspondiente con tabulate. Si el ID no existe, debe mostrar un mensaje de advertencia.

def buscar_id():
    variable = int(input("Ingrese el ID del producto que desea buscar: "))
    resultado = [fila for fila in datos if fila[0] == variable]
    if resultado:
        print(tabulate(resultado, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
    else:
        print(f"El ID {variable} no se encuentra en la lista de productos.")

buscar_id()

# Reto 4 (Avanzado) — Menú Interactivo de Consola: Implementa un bucle while True con un menú
# interactivo que permita al usuario elegir entre: 1) Ver catálogo completo, 2) Filtrar por precio (>
# $100), 3) Buscar por ID, 4) Exportar informe PDF y 5) Salir del programa.


def mostrar_menu():
        while True:
                variable_2 = int(input("\nSeleccione una opción:\n1) Ver catálogo completo" \
                "\n2) Filtrar por precio (> $100)\n3) Buscar por ID\n4) Exportar informe PDF" \
                "\n5) Salir del programa\nIngrese el número de opción: ")) 
                datos1 = []
                for fila in hoja.iter_rows(min_row=2, values_only=True):
                        datos1.append(list(fila))
                encabezados = ["ID", "Producto", "Categoria", "Precio ($)", "Ventas (Unid)"]
                alineacion =["Center", "Left", "Center", "Right", "Right"]
                match variable_2:
                       
                        case 1:
                                print(tabulate(datos1, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
                        case 2:
                                datos_filtrados = [fila for fila in datos1 if fila[3] > 100]
                                print(tabulate(datos_filtrados, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
                        case 3:
                                variable_3 = int(input("Ingrese el ID del producto que desea buscar: "))
                                resultado_2 = [fila for fila in datos1 if fila[0] == variable_3]
                                if resultado_2:
                                    print(tabulate(resultado_2, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
                                else:
                                    print(f"El ID {variable_3} no se encuentra en la lista de productos.")
                        case 4:
                         # Calcular métricas
                            datos1 = [fila for fila in datos1 if fila[3] > 100]
                            total_unidades = sum(fila[4] for fila in datos1)
                            precio_promedio = round(sum(fila[3] for fila in datos1) / len(datos1) if datos1 else 0, 2)
                            datos1.append(["", "", "", "Total Unidades", f"{total_unidades:}"])
                            datos1.append(["", "", "", f"Precio Promedio", f"${precio_promedio:}"])

                                                
                            doc1 = SimpleDocTemplate("reporte_ejecutivo_ventas_mediante_menú.pdf", pagesize=letter)
                            estilos1 = getSampleStyleSheet()
                            estilo_titulo1 = ParagraphStyle(
                            "EstiloTitulo", parent = estilos1["Title"],
                            fontName = "Helvetica-Bold", fontSize = 18,
                            textColor = colors.HexColor("#1A365D"), spaceAfter = 10
                            )
                            estilo_cuerpo1 = ParagraphStyle("EstiloCuerpo", parent = estilos1["Normal"], fontSize = 9)
                            estilo_encabezado1 = ParagraphStyle("EstiloEncabezado", parent = estilos1["Normal"],
                            fontSize = 10, fontName = "Helvetica-Bold",
                            textColor = colors.whitesmoke, alignment = 1)

                            story1 = [
                            Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo1),
                            Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos1)}", estilo_cuerpo1),
                            Spacer(1, 12)
                            ]

                            encabezados_pdf1 =[Paragraph(f"<b>{h}</b>", estilo_encabezado1) for h in encabezados]
                            filas_pdf1 = [
                            [
                            Paragraph(str(r[0]), ParagraphStyle("C", parent = estilo_cuerpo1, alignment = 1)),
                            Paragraph(str(r[1]), ParagraphStyle("L", parent = estilo_cuerpo1, alignment = 0)),
                            Paragraph(str(r[2]), ParagraphStyle("C", parent = estilo_cuerpo1, alignment = 1)),
                            Paragraph(str(r[3]), ParagraphStyle("R", parent = estilo_cuerpo1, alignment = 2)),
                            Paragraph(str(r[4]), ParagraphStyle("R", parent = estilo_cuerpo1, alignment = 2))
                            ]
                            for r in datos1
                            ]

                            ancho_cols1 = [40, 160, 100, 80, 80]
                            tabla_pdf1 = Table([encabezados_pdf1] + filas_pdf1, colWidths = ancho_cols1, repeatRows = 1)

                            tabla_pdf1.setStyle(TableStyle([
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
                            ("VALIGN" , (0, 0), (-1, -1), "MIDDLE"),
                            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
                            ]))            

                            story1.append(tabla_pdf1)
                            doc1.build(story1)
                            print("Informe PDF exportado exitosamente.")
                        case 5:
                                print("Saliendo del programa...")
                                break
                        case _:
                                print("Opción inválida. Por favor, seleccione una opción válida.")
                        
mostrar_menu()