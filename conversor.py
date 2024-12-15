import os
from tkinter import Tk, Label, Button, Entry, StringVar, messagebox, ttk
from tkinter.ttk import Progressbar, Treeview
from tkinter.filedialog import askdirectory, askopenfilenames
from PIL import Image
from PyPDF2 import PdfMerger
from tkinter.filedialog import asksaveasfilename

# Clases base y derivadas
class Conversor:
    def __init__(self, ruta):
        self.ruta = ruta

    def procesar_archivos(self):
        raise NotImplementedError("Este método debe ser implementado en las clases derivadas.")

class ConversorImagen(Conversor):
    def __init__(self, ruta_imagen, formato_salida):
        super().__init__(ruta_imagen)
        self.formato_salida = formato_salida

    def procesar_archivos(self):
        try:
            archivos = [f for f in os.listdir(self.ruta) if f.lower().endswith(('jpg', 'jpeg', 'png', 'webp'))]
            archivos.sort()  # Orden natural
            total_archivos = len(archivos)

            # Lista para almacenar las imágenes convertidas a RGB
            imagenes_pdf = []

            for idx, archivo in enumerate(archivos):
                ruta_archivo = os.path.join(self.ruta, archivo)
                img = Image.open(ruta_archivo)
                imagenes_pdf.append(img.convert("RGB"))
                yield (idx + 1) / total_archivos

            # Obtener el nombre de la carpeta actual
            nombre_carpeta = os.path.basename(self.ruta)
            nombre_default = f"{nombre_carpeta}.pdf"

            # Cuadro de diálogo para guardar el archivo
            nombre_final = asksaveasfilename(
                title="Guardar PDF",
                defaultextension=".pdf",
                filetypes=[("Archivos PDF", "*.pdf")],
                initialfile=nombre_default
            )
            if not nombre_final:
                raise ValueError("No se seleccionó una ubicación para guardar el archivo.")

            # Guardar todas las imágenes en un único archivo PDF
            if imagenes_pdf:
                imagenes_pdf[0].save(nombre_final, save_all=True, append_images=imagenes_pdf[1:])
        except Exception as e:
            raise ValueError(f"Error al procesar imágenes: {str(e)}")


class CombinarPDFs(Conversor):
    def __init__(self, archivos):
        self.archivos = archivos

    def procesar_archivos(self):
        try:
            merger = PdfMerger()
            for archivo in self.archivos:
                merger.append(archivo)
            nombre_base = os.path.splitext(os.path.basename(self.archivos[0]))[0]
            ruta_salida = os.path.join(os.path.dirname(self.archivos[0]), f"{nombre_base}_combinado.pdf")
            merger.write(ruta_salida)
            merger.close()
        except Exception as e:
            raise ValueError(f"Error al combinar PDFs: {str(e)}")

# Funciones principales
def seleccionar_ruta():
    ruta_seleccionada = askdirectory()
    if ruta_seleccionada:
        var_ruta_imagen.set(ruta_seleccionada)

def seleccionar_archivos_pdf():
    archivos = askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
    if archivos:
        tabla_pdf.delete(*tabla_pdf.get_children())  # Limpiar tabla
        for idx, archivo in enumerate(archivos, start=1):
            nombre = os.path.basename(archivo)
            tipo = "PDF"
            tamano = f"{os.path.getsize(archivo) / 1024:.2f} KB"
            tabla_pdf.insert("", "end", values=(idx, nombre, tipo, tamano, archivo))

def ejecutar_conversion(conversor):
    progress['value'] = 0
    for progreso in conversor.procesar_archivos():
        progress['value'] = progreso * 100
        ventana.update_idletasks()
    abrir_carpeta_y_mostrar_mensaje(conversor.ruta if hasattr(conversor, 'ruta') else os.path.dirname(conversor.archivos[0]))

def abrir_carpeta_y_mostrar_mensaje(ruta):
    messagebox.showinfo("Proceso Completado", f"Operación completada. Los archivos procesados se encuentran en: {ruta}")

def iniciar_conversion():
    ruta_imagen = var_ruta_imagen.get()
    formato_salida = var_formato_salida.get()
    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    try:
        conversor = ConversorImagen(ruta_imagen, formato_salida)
        ejecutar_conversion(conversor)
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)})")

def iniciar_combinacion_imagenes():
    ruta_imagen = var_ruta_imagen.get()
    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    try:
        conversor = ConversorImagen(ruta_imagen, "pdf")
        ejecutar_conversion(conversor)
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)})")

def iniciar_combinacion_pdf():
    items = tabla_pdf.get_children()
    archivos = [tabla_pdf.item(item, "values")[4] for item in items]
    if not archivos:
        messagebox.showerror("Error", "Por favor, selecciona archivos PDF para combinar.")
        return

    try:
        merger = PdfMerger()
        for archivo in archivos:
            merger.append(archivo)
        
        # Obtener el nombre del archivo principal desde la tabla
        nombre_archivo = os.path.basename(archivos[0])
        cantidad_archivos = len(archivos)

        # Crear un nombre para el archivo PDF combinado con el formato "nombre ##-##"
        nombre_default = f"{nombre_archivo} {cantidad_archivos:02d}-{cantidad_archivos:02d}.pdf"

        # Cuadro de diálogo para guardar el archivo combinado
        nombre_final = asksaveasfilename(
            title="Guardar PDF Combinado",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            initialfile=nombre_default
        )
        if not nombre_final:
            raise ValueError("No se seleccionó una ubicación para guardar el archivo.")

        # Guardar el archivo PDF combinado
        merger.write(nombre_final)
        merger.close()
    except Exception as e:
        raise ValueError(f"Error al combinar PDFs: {str(e)}")

def limpiar_tabla():
    tabla_pdf.delete(*tabla_pdf.get_children())

# Interfaz gráfica
ventana = Tk()
ventana.title("Conversor de Imágenes y PDFs")
ventana.geometry("900x785")
ventana.resizable(False, False)

var_ruta_imagen = StringVar()
var_formato_salida = StringVar(value='png')

Label(ventana, text="Conversor de Imágenes y PDFs", font=("Helvetica", 16, "bold")).pack(pady=10)
progress = Progressbar(ventana, orient='horizontal', length=800, mode='determinate')
progress.pack(pady=20)

frame_imagen = ttk.Labelframe(ventana, text="Opciones de Imágenes", padding=(20, 10))
frame_imagen.pack(pady=10, fill="x", padx=10)

Label(frame_imagen, text="Seleccionar carpeta de archivos:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
Entry(frame_imagen, textvariable=var_ruta_imagen, width=50).grid(row=1, column=0, padx=5, pady=5)
Button(frame_imagen, text="Buscar", command=seleccionar_ruta).grid(row=1, column=1, padx=5, pady=5)

Label(frame_imagen, text="Formato de salida:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
menu_formato = ttk.Combobox(frame_imagen, textvariable=var_formato_salida, values=['jpg', 'png', 'webp'])
menu_formato.grid(row=3, column=0, padx=5, pady=5)

Button(frame_imagen, text="Convertir Imágenes", command=iniciar_conversion).grid(row=4, column=0, padx=5, pady=5, sticky="w")
Button(frame_imagen, text="Combinar Imágenes", command=iniciar_combinacion_imagenes).grid(row=4, column=1, padx=5, pady=5, sticky="ew")

frame_pdf = ttk.Labelframe(ventana, text="Opciones de PDFs", padding=(20, 10))
frame_pdf.pack(pady=10, fill="both", expand=True, padx=10)

Button(frame_pdf, text="Seleccionar Archivos PDF", command=seleccionar_archivos_pdf).pack(pady=5)
Button(frame_pdf, text="Resetear Tabla", command=limpiar_tabla).pack(pady=5, anchor="ne")

columnas = ("#", "Nombre del Archivo", "Tipo", "Espacio")
tabla_pdf = Treeview(frame_pdf, columns=columnas, show="headings", height=10)
tabla_pdf.heading("#", text="N°")
tabla_pdf.heading("Nombre del Archivo", text="Nombre del Archivo")
tabla_pdf.heading("Tipo", text="Tipo")
tabla_pdf.heading("Espacio", text="Espacio")


for col in columnas[:-1]:
    tabla_pdf.column(col, width=150, anchor="center")

def mover_item(event):
    item = tabla_pdf.selection()[0]
    tabla_pdf.move(item, '', tabla_pdf.index(item) - 1 if event.keysym == 'Up' else tabla_pdf.index(item) + 1)

ventana.bind('<Up>', mover_item)
ventana.bind('<Down>', mover_item)

tabla_pdf.pack(pady=5, fill="x")

Button(frame_pdf, text="Combinar PDFs", command=iniciar_combinacion_pdf).pack(pady=10)

ventana.mainloop()
