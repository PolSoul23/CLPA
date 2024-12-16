import os
from tkinter import Tk, Label, Button, Entry, StringVar, messagebox, ttk
from tkinter import filedialog
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
        """Método abstracto para ser implementado en clases derivadas."""
        raise NotImplementedError("Este método debe ser implementado en las clases derivadas.")

class ConversorImagen(Conversor):
    def __init__(self, ruta_imagen, formato_salida):
        super().__init__(ruta_imagen)
        self.formato_salida = formato_salida

    def procesar_archivos(self):
        """Convierte las imágenes en la ruta especificada al formato de salida deseado."""
        try:
            archivos = [f for f in os.listdir(self.ruta) if f.lower().endswith(('jpg', 'jpeg', 'png', 'webp'))]
            archivos.sort()  # Orden natural
            total_archivos = len(archivos)

            # Convertir y guardar imágenes en el formato deseado
            for idx, archivo in enumerate(archivos):
                try:
                    ruta_archivo = os.path.join(self.ruta, archivo)
                    img = Image.open(ruta_archivo)

                    # Convertir a RGB si la imagen está en un modo incompatible con JPEG
                    if img.mode not in ["RGB", "L"]:
                        img = img.convert("RGB")

                    # Validar formato de salida
                    if self.formato_salida.lower() not in ["jpg", "jpeg", "png", "webp"]:
                        raise ValueError(f"Formato de salida '{self.formato_salida}' no soportado.")    

                    # Definir la ruta del archivo de salida
                    extension_salida = self.formato_salida.lower() if self.formato_salida.lower() != "jpg" else "jpeg"
                    ruta_guardado = os.path.join(self.ruta, f"{os.path.splitext(archivo)[0]}.{extension_salida}")

                    # Guardar la imagen en el formato deseado
                    img.save(ruta_guardado, "JPEG" if self.formato_salida.lower() == "jpg" else self.formato_salida.upper())
                    yield (idx + 1) / total_archivos
                except Exception as e:
                    print(f"Error al procesar la imagen {archivo}: {str(e)}")
                    raise ValueError(f"Error al procesar la imagen {archivo}: {str(e)}")
        except Exception as e:
            raise ValueError(f"Error al procesar imágenes: {str(e)}")

# Función para convertir imágenes a PDF
def convertir_imagen_a_pdf(imagen_path, pdf_path):
    """Convierte una imagen a PDF."""
    try:
        img = Image.open(imagen_path)
        img = img.convert("RGB")  # Convertir la imagen al formato RGB si es necesario
        img.save(pdf_path, "PDF")
    except Exception as e:
        print(f"Error al convertir la imagen a PDF: {e}")
        raise ValueError(f"Error al convertir la imagen a PDF: {str(e)}")

# Función para combinar PDFs en uno solo
def combinar_pdfs(pdf_paths, output_pdf):
    """Combina múltiples archivos PDF en uno solo."""
    try:
        merger = PdfMerger()
        for pdf in pdf_paths:
            merger.append(pdf)
        merger.write(output_pdf)
        merger.close()
    except Exception as e:
        print(f"Error al combinar PDFs: {e}")
        raise ValueError(f"Error al combinar PDFs: {str(e)}")

# Funciones auxiliares
def seleccionar_ruta():
    """Permite al usuario seleccionar una carpeta."""
    ruta_seleccionada = askdirectory()
    if ruta_seleccionada:
        var_ruta_imagen.set(ruta_seleccionada)

def seleccionar_archivos_pdf():
    """Permite al usuario seleccionar archivos PDF y los muestra en una tabla."""
    archivos = askopenfilenames(filetypes=[("Archivos PDF", "*.pdf")])
    if archivos:
        tabla_pdf.delete(*tabla_pdf.get_children())  # Limpiar tabla
        for idx, archivo in enumerate(archivos, start=1):
            nombre = os.path.basename(archivo)
            tipo = "PDF"
            tamano = f"{os.path.getsize(archivo) / 1024:.2f} KB"
            tabla_pdf.insert("", "end", values=(idx, nombre, tipo, tamano, archivo))

def ejecutar_conversion(conversor):
    """Ejecuta la conversión de archivos y actualiza el progreso."""
    progress['value'] = 0
    for progreso in conversor.procesar_archivos():
        progress['value'] = progreso * 100
        ventana.update_idletasks()
    abrir_carpeta_y_mostrar_mensaje(conversor.ruta)

def abrir_carpeta_y_mostrar_mensaje(ruta):
    """Muestra un mensaje cuando la operación ha finalizado."""
    messagebox.showinfo("Proceso Completado", f"Operación completada. Los archivos procesados se encuentran en: {ruta}")

# Funciones para verificar tipos de imágenes en una carpeta
def verificar_imagenes(ruta, extension):
    """Verifica si existen imágenes con una extensión específica en la carpeta."""
    return any(f.lower().endswith(extension) for f in os.listdir(ruta))

def iniciar_conversion():
    """Inicia el proceso de conversión de imágenes según el formato seleccionado."""
    ruta_imagen = var_ruta_imagen.get()
    formato_salida = var_formato_salida.get()

    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    # Verificación de formato
    if (formato_salida == "png" and verificar_imagenes(ruta_imagen, 'png')) or \
       (formato_salida in ["jpg", "jpeg"] and verificar_imagenes(ruta_imagen, ('jpg', 'jpeg'))) or \
       (formato_salida == "webp" and verificar_imagenes(ruta_imagen, 'webp')):
        messagebox.showwarning("Advertencia", f"Ya existen imágenes {formato_salida.upper()} en la carpeta. No puedes convertir a {formato_salida.upper()}.")
        return

    try:
        conversor = ConversorImagen(ruta_imagen, formato_salida)
        ejecutar_conversion(conversor)
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)}")

# Iniciar la conversión de imágenes a PDF y luego combinarlas
def iniciar_combinacion_imagenes():
    """Convierte imágenes a PDF y luego las combina en un solo archivo PDF."""
    ruta_imagen = var_ruta_imagen.get()
    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    try:
        archivos_imagen = [f for f in os.listdir(ruta_imagen) if f.lower().endswith(('jpg', 'jpeg', 'png', 'webp'))]
        pdfs_generados = []
        for archivo in archivos_imagen:
            imagen_path = os.path.join(ruta_imagen, archivo)
            pdf_name = os.path.splitext(archivo)[0] + ".pdf"
            pdf_path = os.path.join(ruta_imagen, pdf_name)
            convertir_imagen_a_pdf(imagen_path, pdf_path)
            pdfs_generados.append(pdf_path)

        output_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Archivos PDF", "*.pdf")])
        if output_pdf:
            combinar_pdfs(pdfs_generados, output_pdf)
            messagebox.showinfo("Proceso Completado", f"Las imágenes fueron combinadas exitosamente en: {output_pdf}")
        else:
            raise ValueError("No se seleccionó una ubicación para guardar el PDF combinado.")
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)}")

def limpiar_tabla():
    """Limpia la tabla de archivos PDF seleccionados."""
    tabla_pdf.delete(*tabla_pdf.get_children())

# Interfaz gráfica
ventana = Tk()
ventana.title("Conversor CLPA")
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
    """Permite mover los elementos en la tabla."""
    item = tabla_pdf.selection()[0]
    tabla_pdf.move(item, '', tabla_pdf.index(item) - 1 if event.keysym == 'Up' else tabla_pdf.index(item) + 1)

ventana.bind('<Up>', mover_item)
ventana.bind('<Down>', mover_item)

tabla_pdf.pack(fill="both", expand=True, padx=10, pady=10)

ventana.mainloop()