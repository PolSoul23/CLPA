import os
import re
import threading
import comtypes.client
from tkinter import BooleanVar, Checkbutton, Tk, Label, Button, Entry, StringVar, messagebox, ttk
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
    def __init__(self, ruta_imagen, formato_salida, combinar_pdf=False):
        super().__init__(ruta_imagen)
        self.formato_salida = formato_salida.lower()
        self.combinar_pdf = combinar_pdf
        
    class ConversorWord(Conversor):
     def __init__(self, archivos, combinar_pdf=False):
        super().__init__(None)
        self.archivos = archivos
        self.combinar_pdf = combinar_pdf

    def procesar_archivos(self):
        
      try:
 
        # Extensiones permitidas para imágenes y documentos Word
        extensiones_imagenes = ('jpg', 'jpeg', 'png', 'webp', 'bmp', 'tiff', 'tif', 'gif', 'ico', 'heic', 'heif')
        extensiones_word = ('.doc', '.docx')

        archivos_imagenes = [f for f in os.listdir(self.ruta) if f.lower().endswith(extensiones_imagenes)]
        archivos_word = [f for f in os.listdir(self.ruta) if f.lower().endswith(extensiones_word)]

        archivos_imagenes.sort()  # Orden natural
        total_archivos = len(archivos_imagenes) + len(archivos_word)

        if not archivos_imagenes and not archivos_word:
            raise ValueError("No se encontraron archivos válidos en la ruta seleccionada.")

        if archivos_imagenes:
            if self.combinar_pdf:
                # Lista para almacenar imágenes convertidas a RGB
                imagenes_pdf = []

                for idx, archivo in enumerate(archivos_imagenes):
                    ruta_archivo = os.path.join(self.ruta, archivo)
                    img = Image.open(ruta_archivo)
                    img = img.convert("RGB")  # Asegurar compatibilidad con PDF
                    imagenes_pdf.append(img)
                    yield (idx + 1) / total_archivos

                # Guardar todas las imágenes como PDF combinado
                nombre_carpeta = os.path.basename(self.ruta.rstrip(os.sep))
                nombre_default = f"{nombre_carpeta}.pdf"

                nombre_final = asksaveasfilename(
                    title="Guardar PDF",
                    defaultextension=".pdf",
                    filetypes=[("Archivos PDF", "*.pdf")],
                    initialfile=nombre_default
                )
                if not nombre_final:
                    raise ValueError("No se seleccionó una ubicación para guardar el archivo.")

                if imagenes_pdf:
                    imagenes_pdf[0].save(nombre_final, save_all=True, append_images=imagenes_pdf[1:])
            else:
                # Convertir cada imagen individualmente
                for idx, archivo in enumerate(archivos_imagenes):
                    ruta_archivo = os.path.join(self.ruta, archivo)
                    img = Image.open(ruta_archivo)
                    # Comprobar compatibilidad para salida
                    if self.formato_salida in ["jpg", "jpeg", "png", "bmp", "webp"]:
                        img = img.convert("RGB")  # Convertir a RGB para formatos sin transparencia
                    elif self.formato_salida in ["tiff", "tif"]:
                        img = img.convert("RGBA")  # Soporte para transparencia en TIFF

                    nombre_archivo_salida = os.path.splitext(archivo)[0] + f".{self.formato_salida}"
                    ruta_archivo_salida = os.path.join(self.ruta, nombre_archivo_salida)
                    img.save(ruta_archivo_salida, format=self.formato_salida.upper())
                    yield (idx + 1) / total_archivos

        if archivos_word:
            documentos_pdf = []

            for idx, archivo in enumerate(archivos_word):
                ruta_archivo = os.path.join(self.ruta, archivo)
                word = comtypes.client.CreateObject('Word.Application')
                doc = word.Documents.Open(ruta_archivo)
                archivo_pdf = ruta_archivo.replace('.docx', '.pdf').replace('.doc', '.pdf')
                doc.SaveAs(archivo_pdf, FileFormat=17)  # 17: Formato PDF
                doc.Close()
                word.Quit()

                documentos_pdf.append(archivo_pdf)
                yield (len(archivos_imagenes) + idx + 1) / total_archivos

            if self.combinar_pdf:
                merger = PdfMerger()
                for pdf in documentos_pdf:
                    merger.append(pdf)

                nombre_base = os.path.basename(self.ruta.rstrip(os.sep))
                nombre_default = f"{nombre_base}_word_combinado.pdf"

                nombre_final = asksaveasfilename(
                    title="Guardar PDF Combinado",
                    defaultextension=".pdf",
                    filetypes=[("Archivos PDF", "*.pdf")],
                    initialfile=nombre_default
                )
                if not nombre_final:
                    raise ValueError("No se seleccionó una ubicación para guardar el archivo combinado.")

                merger.write(nombre_final)
                merger.close()
      except Exception as e:
        raise ValueError(f"Error al procesar archivos: {str(e)}")

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

def ejecutar_conversion_en_hilo(conversor):
    def proceso_conversion():
        progress['value'] = 0
        try:
            # Ejecutar el proceso
            for progreso in conversor.procesar_archivos():
                progress['value'] = progreso * 100
                ventana.update_idletasks()

            # Mensaje de éxito al finalizar el proceso
            messagebox.showinfo("Proceso Completado", "Operación completada exitosamente.")

            # Abrir la carpeta de destino
            abrir_carpeta_y_mostrar_mensaje(conversor.ruta if hasattr(conversor, 'ruta') else os.path.dirname(conversor.archivos[0]))

        except Exception as e:
            # Si hay un error, mostrar mensaje de error
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)}")

    # Crear y ejecutar el hilo
    hilo_conversion = threading.Thread(target=proceso_conversion)
    hilo_conversion.start()


def abrir_carpeta_y_mostrar_mensaje(ruta):
    messagebox.showinfo("Proceso Completado", f"Operación completada. Los archivos procesados se encuentran en: {ruta}")
    try:
        # Intentar abrir la ruta en el explorador de archivos
        os.startfile(ruta)  # Solo funciona en sistemas Windows
    except Exception as e:
        # Si no se puede abrir la carpeta, mostrar un mensaje
        messagebox.showerror("Error", f"No se pudo abrir la carpeta: {str(e)}")
        
def iniciar_conversion():
    ruta_imagen = var_ruta_imagen.get()
    formato_salida = var_formato_salida.get()
    combinar_pdf = var_combinar_pdf.get()  # Supongamos que usas un checkbutton para esta opción

    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    try:
        conversor = ConversorImagen(ruta_imagen, formato_salida, combinar_pdf=combinar_pdf)
        # Ejecutar la conversión con confirmación
        ejecutar_conversion_en_hilo(conversor)

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
        conversor = ConversorImagen(ruta_imagen, "pdf", combinar_pdf=True)
        # Ejecutar la conversión con confirmación
        ejecutar_conversion_en_hilo(conversor)

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
        conversor = CombinarPDFs(archivos)
        # Ejecutar la combinación de PDFs con confirmación
        ejecutar_conversion_en_hilo(conversor)
    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)})")
        
    try:
        merger = PdfMerger()
        for archivo in archivos:
            merger.append(archivo)

        # Obtener las numeraciones desde los nombres de los archivos
        numeros = []
        for archivo in archivos:
            nombre = os.path.basename(archivo)
            # Extraer el número de capítulo/volumen del nombre
            match = re.search(r'\d+(\.\d+)?', nombre)
            if match:
                numeros.append(float(match.group()))
        
        if numeros:
            numeros.sort()  # Ordenar las numeraciones
            inicio = numeros[0]
            fin = numeros[-1]
        else:
            inicio, fin = 1, len(archivos)  # Valores por defecto si no se encuentran numeraciones

        # Crear un nombre para el archivo PDF combinado con el rango de numeración detectado
        nombre_default = f"{os.path.splitext(os.path.basename(archivos[0]))[0]} {inicio:.1f}-{fin:.1f}.pdf"

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

#Funciones de Word
def seleccionar_archivos_word():
      archivos = askopenfilenames(filetypes=[("Documentos Word", "*.doc;*.docx")])
      if archivos:
        tabla_pdf.delete(*tabla_pdf.get_children())  # Reutilizar la tabla para mostrar los archivos
        for idx, archivo in enumerate(archivos, start=1):
            nombre = os.path.basename(archivo)
            tipo = "Word"
            tamano = f"{os.path.getsize(archivo) / 1024:.2f} KB"
            tabla_pdf.insert("", "end", values=(idx, nombre, tipo, tamano, archivo))
def iniciar_conversion_word():
    items = tabla_pdf.get_children()
    archivos = [tabla_pdf.item(item, "values")[4] for item in items]
    combinar_pdf = var_combinar_pdf.get()  # Supongamos que usas un checkbutton para esta opción

    if not archivos:
        messagebox.showerror("Error", "Por favor, selecciona documentos Word para procesar.")
        return

    try:
        conversor = ConversorWord(archivos, combinar_pdf=combinar_pdf)
        # Ejecutar la conversión con confirmación
        ejecutar_conversion_en_hilo(conversor)

    except ValueError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado: {str(e)}")

# Ventana principal
ventana = Tk()
ventana.title("Conversor Luna")
ventana.geometry("900x800")
ventana.resizable(False, False)

var_ruta_imagen = StringVar()
var_formato_salida = StringVar(value='png')

# Título principal
Label(ventana, text="Conversor de Imágenes y PDFs", font=("Helvetica", 16, "bold")).pack(pady=10)

# Barra de progreso
progress = Progressbar(ventana, orient='horizontal', length=800, mode='determinate')
progress.pack(pady=20)

# Frame de opciones de imágenes
frame_imagen = ttk.Labelframe(ventana, text="Opciones de Imágenes", padding=(20, 10))
frame_imagen.pack(pady=10, fill="x", padx=10)

frame_imagen.columnconfigure(0, weight=3)  # Espacio para el Entry
frame_imagen.columnconfigure(1, weight=1)  # Espacio para el botón

# Campo de selección de ruta
Label(frame_imagen, text="Seleccionar carpeta de archivos:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_ruta = Entry(frame_imagen, textvariable=var_ruta_imagen, width=38, state="readonly")
entry_ruta.grid(row=1, column=0, padx=(1, 1), pady=5, sticky="ew")
entry_ruta.configure({"readonlybackground": "white"})

Button(frame_imagen, text="Buscar", command=seleccionar_ruta).grid(row=1, column=1, padx=(1, 1), pady=5)


# Configurar la columna para que se ajuste al tamaño del contenido
frame_imagen.columnconfigure(0, weight=1, minsize=50)  # Ajusta la columna para que sea más pequeña si es necesario
# Formato de salida 
Label(frame_imagen, text="Formato de salida:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
menu_formato = ttk.Combobox(frame_imagen, textvariable=var_formato_salida, state="readonly", width=15)
menu_formato['values'] = ["JPG", "PNG", "WEBP", "BMP", "TIFF", "GIF", "ICO", "HEIC"]
menu_formato.current(0)
menu_formato.grid(row=3, column=0, padx=1, pady=5, sticky="ew")

# Botones de acción para imágenes
Button(frame_imagen, text="Convertir Imágenes", command=iniciar_conversion).grid(row=4, column=0, padx=5, pady=5, sticky="w")
Button(frame_imagen, text="Combinar Imágenes", command=iniciar_combinacion_imagenes).grid(row=4, column=1, padx=5, pady=5, sticky="ew")

btn_seleccionar_word = Button(ventana, text="Seleccionar Word", command=seleccionar_archivos_word)
btn_seleccionar_word.pack(pady=5)
btn_iniciar_word = Button(ventana, text="Convertir Word a PDF", command=iniciar_conversion_word)
btn_iniciar_word.pack(pady=5)

# Checkbox para combinar imágenes en PDF
var_combinar_pdf = BooleanVar(value=False)
checkbox_combinar_pdf = Checkbutton(ventana, text="Combinar imágenes en PDF", variable=var_combinar_pdf)
checkbox_combinar_pdf.pack(pady=5)
checkbox_combinar_pdf.pack_forget()  # Hace que el checkbox sea invisible


# Frame de opciones de PDFs
frame_pdf = ttk.Labelframe(ventana, text="Opciones de PDFs", padding=(20, 10))
frame_pdf.pack(pady=10, fill="both", expand=True, padx=10)

# Botones para PDFs
Button(frame_pdf, text="Seleccionar Archivos PDF", command=seleccionar_archivos_pdf).pack(pady=5, anchor="w")
Button(frame_pdf, text="Resetear Tabla", command=limpiar_tabla).pack(pady=5, anchor="ne")

# Tabla de PDFs
columnas = ("#", "Nombre del Archivo", "Tipo", "Espacio")
tabla_pdf = Treeview(frame_pdf, columns=columnas, show="headings", height=10)
tabla_pdf.heading("#", text="N°")
tabla_pdf.heading("Nombre del Archivo", text="Nombre del Archivo")
tabla_pdf.heading("Tipo", text="Tipo")
tabla_pdf.heading("Espacio", text="Espacio")

for col in columnas[:-1]:
    tabla_pdf.column(col, width=150, anchor="center")
tabla_pdf.pack(pady=5, fill="x")

# Botón de combinación de PDFs
Button(frame_pdf, text="Combinar PDFs", command=iniciar_combinacion_pdf).pack(pady=10)

# Funciones de arrastre y orden automático
def iniciar_arrastre(event):
    global dragging_item, drag_start_index
    dragging_item = tabla_pdf.identify_row(event.y)
    drag_start_index = tabla_pdf.index(dragging_item) if dragging_item else None

def realizar_arrastre(event):
    if dragging_item:
        destino = tabla_pdf.identify_row(event.y)
        if destino and destino != dragging_item:
            tabla_pdf.move(dragging_item, '', tabla_pdf.index(destino))

def finalizar_arrastre(event):
    global dragging_item, drag_start_index
    if dragging_item and drag_start_index is not None:
        nuevo_index = tabla_pdf.index(dragging_item)
        if nuevo_index != drag_start_index:
            print(f"El elemento se movió de {drag_start_index} a {nuevo_index}.")
            
# Función de orden automático
def ordenar_automaticamente():
    elementos = [(tabla_pdf.set(item, "Nombre del Archivo"), item) for item in tabla_pdf.get_children()]
    elementos.sort(key=lambda x: x[0])  # Ordenar por nombre del archivo
    for index, (_, item) in enumerate(elementos):
        tabla_pdf.move(item, '', index)

# Asociar eventos al Treeview para drag and drop
tabla_pdf.bind('<ButtonPress-1>', iniciar_arrastre)
tabla_pdf.bind('<B1-Motion>', realizar_arrastre)
tabla_pdf.bind('<ButtonRelease-1>', finalizar_arrastre)

# Iniciar la ventana principal
ventana.mainloop()