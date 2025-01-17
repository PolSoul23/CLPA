import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
import os
import threading

def seleccionar_archivo_word():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Word", "*.docx")])
    if archivo:
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)

def seleccionar_archivo_pdf():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
    if archivo:
        entry_archivo.delete(0, tk.END)
        entry_archivo.insert(0, archivo)

def convertir_a_pdf():
    archivo_word = entry_archivo.get()
    if archivo_word == "":
        messagebox.showerror("Error", "Por favor, selecciona un archivo Word.")
        return

    # Verificar si el archivo ya es un PDF
    if archivo_word.lower().endswith(".pdf"):
        messagebox.showerror("Error", "El archivo ya está en formato PDF.")
        return

    # Deshabilitar el botón y mostrar un mensaje de progreso
    boton_convertir_pdf.config(state=tk.DISABLED)
    mensaje_progreso.config(text="Convirtiendo archivo a PDF...", fg="blue")

    # Crear un hilo para ejecutar la conversión en segundo plano
    hilo_conversion_pdf = threading.Thread(target=convertir_archivo_a_pdf_en_segundo_plano, args=(archivo_word,))
    hilo_conversion_pdf.start()

def convertir_a_word():
    archivo_pdf = entry_archivo.get()
    if archivo_pdf == "":
        messagebox.showerror("Error", "Por favor, selecciona un archivo PDF.")
        return

    # Verificar si el archivo ya es un archivo Word
    if archivo_pdf.lower().endswith(".docx"):
        messagebox.showerror("Error", "El archivo ya está en formato Word.")
        return

    # Deshabilitar el botón y mostrar un mensaje de progreso
    boton_convertir_word.config(state=tk.DISABLED)
    mensaje_progreso.config(text="Convirtiendo archivo a Word...", fg="blue")

    # Crear un hilo para ejecutar la conversión en segundo plano
    hilo_conversion_word = threading.Thread(target=convertir_archivo_a_word_en_segundo_plano, args=(archivo_pdf,))
    hilo_conversion_word.start()

def convertir_archivo_a_pdf_en_segundo_plano(archivo_word):
    try:
        # Verificar que el archivo existe
        if not os.path.exists(archivo_word):
            messagebox.showerror("Error", "El archivo no existe.")
            return

        # Iniciar la aplicación Word de forma oculta
        word = win32.Dispatch('Word.Application')
        word.visible = False  # No mostrar la ventana de Word

        # Asegurarse de que se usa una ruta absoluta
        archivo_word_absoluto = os.path.abspath(archivo_word)

        # Abrir el archivo Word
        doc = word.Documents.Open(archivo_word_absoluto)

        # Establecer el nombre del archivo PDF de salida
        archivo_pdf = archivo_word.replace(".docx", ".pdf")
        archivo_pdf_absoluto = os.path.abspath(archivo_pdf)

        # Guardar el documento como PDF
        doc.SaveAs(archivo_pdf_absoluto, FileFormat=17)  # 17 es el formato PDF

        # Cerrar el documento y la aplicación de Word
        doc.Close()
        word.Quit()

        # Mostrar el mensaje de éxito
        messagebox.showinfo("Éxito", f"Archivo convertido a PDF: {archivo_pdf_absoluto}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")
    finally:
        # Borrar la ruta del archivo en la entrada
        entry_archivo.delete(0, tk.END)
        # Rehabilitar el botón y limpiar el mensaje de progreso
        boton_convertir_pdf.config(state=tk.NORMAL)
        mensaje_progreso.config(text="")

def convertir_archivo_a_word_en_segundo_plano(archivo_pdf):
    try:
        # Verificar que el archivo existe
        if not os.path.exists(archivo_pdf):
            messagebox.showerror("Error", "El archivo no existe.")
            return

        # Iniciar la aplicación Word de forma oculta
        word = win32.Dispatch('Word.Application')
        word.visible = False  # No mostrar la ventana de Word

        # Asegurarse de que se usa una ruta absoluta
        archivo_pdf_absoluto = os.path.abspath(archivo_pdf)

        # Establecer el nombre del archivo Word de salida
        archivo_word = archivo_pdf.replace(".pdf", ".docx")
        archivo_word_absoluto = os.path.abspath(archivo_word)

        # Abrir el archivo PDF en Word
        doc = word.Documents.Open(archivo_pdf_absoluto)

        # Guardar el documento como Word
        doc.SaveAs(archivo_word_absoluto, FileFormat=16)  # 16 es el formato DOCX

        # Cerrar el documento y la aplicación de Word
        doc.Close()
        word.Quit()

        # Mostrar el mensaje de éxito
        messagebox.showinfo("Éxito", f"Archivo convertido a Word: {archivo_word_absoluto}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")
    finally:
        # Borrar la ruta del archivo en la entrada
        entry_archivo.delete(0, tk.END)
        # Rehabilitar el botón y limpiar el mensaje de progreso
        boton_convertir_word.config(state=tk.NORMAL)
        mensaje_progreso.config(text="")

# Configuración de la interfaz gráfica con Tkinter
ventana = tk.Tk()
ventana.title("Conversor de Word a PDF / PDF a Word")

# Etiqueta para mostrar el archivo seleccionado
label = tk.Label(ventana, text="Selecciona el archivo para convertir:")
label.pack(pady=10)

# Entrada para mostrar el archivo seleccionado
entry_archivo = tk.Entry(ventana, width=50)
entry_archivo.pack(pady=10)

# Botón para seleccionar archivo Word
boton_seleccionar_word = tk.Button(ventana, text="Seleccionar archivo Word", command=seleccionar_archivo_word)
boton_seleccionar_word.pack(pady=5)

# Botón para seleccionar archivo PDF
boton_seleccionar_pdf = tk.Button(ventana, text="Seleccionar archivo PDF", command=seleccionar_archivo_pdf)
boton_seleccionar_pdf.pack(pady=5)

# Botón para convertir a PDF
boton_convertir_pdf = tk.Button(ventana, text="Convertir a PDF", command=convertir_a_pdf)
boton_convertir_pdf.pack(pady=5)

# Botón para convertir a Word
boton_convertir_word = tk.Button(ventana, text="Convertir a Word", command=convertir_a_word)
boton_convertir_word.pack(pady=5)

# Mensaje de progreso
mensaje_progreso = tk.Label(ventana, text="")
mensaje_progreso.pack(pady=10)

# Iniciar la interfaz gráfica
ventana.mainloop()

