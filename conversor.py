import threading
import os
import subprocess
import re
from tkinter import Tk, Label, Button, Entry, StringVar, filedialog, messagebox, ttk
from tkinter.ttk import Progressbar
from PIL import Image
from PyPDF2 import PdfMerger

# Función para seleccionar la carpeta de imágenes
def seleccionar_ruta():
    ruta = filedialog.askdirectory()
    if ruta:
        var_ruta_imagen.set(ruta)

# Función para seleccionar archivos PDF
def seleccionar_pdfs():
    archivos = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if archivos:
        for archivo in archivos:
            tabla_pdfs.insert('', 'end', values=(archivo, os.path.basename(archivo)))

# Función para la clave de orden natural (numeración)
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

# Función para convertir imágenes a otro formato
def convertir_imagen(ruta_imagen, formato_salida, progress):
    archivos = [f for f in os.listdir(ruta_imagen) if f.lower().endswith(('jpg', 'jpeg', 'png', 'webp'))]
    total_archivos = len(archivos)
    for idx, archivo in enumerate(archivos):
        try:
            ruta_archivo = os.path.join(ruta_imagen, archivo)
            img = Image.open(ruta_archivo)
            nombre_salida = os.path.splitext(archivo)[0] + f'.{formato_salida}'
            ruta_salida = os.path.join(ruta_imagen, nombre_salida)
            img.save(ruta_salida, formato_salida.upper())
            progress['value'] = ((idx + 1) / total_archivos) * 100
            ventana.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar {archivo}: {str(e)}")
    abrir_carpeta_y_mostrar_mensaje(ruta_imagen)

# Función para iniciar la conversión en segundo plano
def convertir():
    ruta_imagen = var_ruta_imagen.get()
    formato_salida = var_formato_salida.get()
    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    archivos = [f for f in os.listdir(ruta_imagen) if f.lower().endswith(('jpg', 'jpeg', 'png', 'webp'))]
    total_archivos = len(archivos)
    if total_archivos == 0:
        messagebox.showinfo("Información", "No se encontraron imágenes en la carpeta seleccionada.")
        return

    progress['value'] = 0
    thread = threading.Thread(target=convertir_imagen, args=(ruta_imagen, formato_salida, progress))
    thread.start()

def iniciar_combinacion_pdf():
    thread = threading.Thread(target=combinar_imagenes_a_pdf, args=(progress,))
    thread.start()

def combinar_imagenes_a_pdf(progress):
    ruta_imagen = var_ruta_imagen.get()
    if not os.path.isdir(ruta_imagen):
        messagebox.showerror("Error", "Por favor, selecciona una ruta válida.")
        return

    archivos = sorted([f for f in os.listdir(ruta_imagen) if f.lower().endswith(('jpg', 'jpeg', 'png', 'webp'))],
                      key=natural_sort_key)
    total_archivos = len(archivos)
    if total_archivos == 0:
        messagebox.showinfo("Información", "No se encontraron imágenes JPG o PNG en la carpeta seleccionada.")
        return

    imagenes = []
    for idx, archivo in enumerate(archivos):
        try:
            ruta_archivo = os.path.join(ruta_imagen, archivo)
            img = Image.open(ruta_archivo).convert('RGB')
            imagenes.append(img)
            progress['value'] = ((idx + 1) / total_archivos) * 100
            ventana.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Error al abrir {archivo}: {str(e)}")

    if imagenes:
        ruta_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if ruta_pdf:
            imagenes[0].save(ruta_pdf, save_all=True, append_images=imagenes[1:])
            abrir_carpeta_y_mostrar_mensaje(os.path.dirname(ruta_pdf))

# Función para combinar los archivos PDF seleccionados
def combinar_pdfs():
    if not tabla_pdfs.get_children():
        messagebox.showerror("Error", "No se han seleccionado PDFs para combinar.")
        return

    # Selecciona la ruta de guardado
    ruta_pdf_final = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if not ruta_pdf_final:
        return

    merger = PdfMerger()

    # Añadir los PDFs seleccionados al merger
    for item in tabla_pdfs.get_children():
        ruta_pdf = tabla_pdfs.item(item, "values")[0]
        merger.append(ruta_pdf)

    try:
        # Guardar el archivo combinado
        merger.write(ruta_pdf_final)
        merger.close()
        abrir_carpeta_y_mostrar_mensaje(os.path.dirname(ruta_pdf_final))
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo combinar los PDFs: {str(e)}")

# Función para abrir carpeta y mostrar mensaje
def abrir_carpeta_y_mostrar_mensaje(carpeta):
    if os.name == 'nt':
        os.startfile(carpeta)
    elif os.name == 'posix':
        subprocess.call(['open', carpeta])
    messagebox.showinfo("Finalizado", "El proceso ha finalizado correctamente.")

# Interfaz gráfica
ventana = Tk()
ventana.title("Conversor de Imágenes y PDFs")
ventana.geometry("600x700")
ventana.resizable(False, False)

var_ruta_imagen = StringVar()
var_formato_salida = StringVar(value='png')

# Título
titulo = Label(ventana, text="Conversor de Imágenes y PDFs", font=("Helvetica", 16, "bold"))
titulo.pack(pady=10)

# Barra de progreso
progress = Progressbar(ventana, orient='horizontal', length=500, mode='determinate')
progress.pack(pady=20)

# Sección para conversión de imágenes
frame_imagen = ttk.Labelframe(ventana, text="Conversión de Imágenes", padding=(20, 10))
frame_imagen.pack(pady=10, fill="x", padx=10)

# Ruta de imagen
etiqueta_imagen = Label(frame_imagen, text="Seleccionar carpeta de imágenes:")
etiqueta_imagen.grid(row=0, column=0, padx=5, pady=5, sticky="w")
entrada_imagen = Entry(frame_imagen, textvariable=var_ruta_imagen, width=40)
entrada_imagen.grid(row=1, column=0, padx=5, pady=5)
boton_seleccionar = Button(frame_imagen, text="Buscar", command=seleccionar_ruta)
boton_seleccionar.grid(row=1, column=1, padx=5, pady=5)

# Formato de salida
etiqueta_formato = Label(frame_imagen, text="Formato de salida:")
etiqueta_formato.grid(row=2, column=0, padx=5, pady=5, sticky="w")
menu_formato = ttk.Combobox(frame_imagen, textvariable=var_formato_salida, values=['jpg', 'png', 'webp'])
menu_formato.grid(row=3, column=0, padx=5, pady=5)

# Botones de conversión
boton_convertir = Button(frame_imagen, text="Convertir Imagen(es)", command=convertir)
boton_convertir.grid(row=4, column=0, padx=5, pady=10, sticky="ew")
boton_combinar_pdf = Button(frame_imagen, text="Combinar Imágenes en PDF", command=iniciar_combinacion_pdf)
boton_combinar_pdf.grid(row=5, column=0, padx=5, pady=10, sticky="ew")

# Sección para combinar PDFs
frame_pdf = ttk.Labelframe(ventana, text="Combinar PDFs", padding=(20, 10))
frame_pdf.pack(pady=10, fill="x", padx=10)

# Tabla para seleccionar PDFs
etiqueta_pdfs = Label(frame_pdf, text="Seleccionar PDFs para combinar:")
etiqueta_pdfs.grid(row=0, column=0, padx=5, pady=5, sticky="w")
tabla_pdfs = ttk.Treeview(frame_pdf, columns=("Ruta", "Nombre"), show="headings", selectmode="extended", height=5)
tabla_pdfs.heading("Ruta", text="Ruta Completa")
tabla_pdfs.heading("Nombre", text="Nombre del Archivo")
tabla_pdfs.grid(row=1, column=0, padx=5, pady=5, columnspan=2, sticky="ew")

# Botón para agregar PDFs
boton_agregar_pdf = Button(frame_pdf, text="Agregar PDFs", command=seleccionar_pdfs)
boton_agregar_pdf.grid(row=2, column=0, padx=5, pady=10, sticky="ew")

# Botón para combinar PDFs
boton_combinar_pdfs = Button(frame_pdf, text="Combinar PDFs", command=combinar_pdfs)
boton_combinar_pdfs.grid(row=2, column=1, padx=5, pady=10, sticky="ew")

ventana.mainloop()
