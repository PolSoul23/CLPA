import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
import os
import threading
from tkinter import ttk
import time

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

    if archivo_word.lower().endswith(".pdf"):
        messagebox.showerror("Error", "El archivo ya está en formato PDF.")
        return

    boton_convertir_pdf.config(state=tk.DISABLED)

    ventana_progreso = tk.Toplevel()
    ventana_progreso.title("Progreso de la Conversión")
    ventana_progreso.geometry("300x150")
    ventana_progreso.transient(ventana)
    ventana_progreso.grab_set()
    ventana_progreso.focus_set()

    barra_progreso = ttk.Progressbar(ventana_progreso, orient="horizontal", length=250, mode="determinate")
    barra_progreso.pack(pady=20)
    barra_progreso["maximum"] = 100

    porcentaje_label = tk.Label(ventana_progreso, text="Cargando... 0%")
    porcentaje_label.pack()

    nombre_archivo_label = tk.Label(ventana_progreso, text=f"Cargando {os.path.basename(archivo_word)} a PDF...")
    nombre_archivo_label.pack()

    hilo_conversion_pdf = threading.Thread(
        target=convertir_archivo_a_pdf_en_segundo_plano,
        args=(archivo_word, barra_progreso, porcentaje_label, ventana_progreso, nombre_archivo_label)
    )
    hilo_conversion_pdf.start()

def convertir_a_word():
    archivo_pdf = entry_archivo.get()
    if archivo_pdf == "":
        messagebox.showerror("Error", "Por favor, selecciona un archivo PDF.")
        return

    if archivo_pdf.lower().endswith(".docx"):
        messagebox.showerror("Error", "El archivo ya está en formato Word.")
        return

    boton_convertir_word.config(state=tk.DISABLED)

    ventana_progreso = tk.Toplevel()
    ventana_progreso.title("Progreso de la Conversión")
    ventana_progreso.geometry("300x150")
    ventana_progreso.transient(ventana)
    ventana_progreso.grab_set()
    ventana_progreso.focus_set()

    barra_progreso = ttk.Progressbar(ventana_progreso, orient="horizontal", length=250, mode="determinate")
    barra_progreso.pack(pady=20)
    barra_progreso["maximum"] = 100

    porcentaje_label = tk.Label(ventana_progreso, text="Cargando... 0%")
    porcentaje_label.pack()

    nombre_archivo_label = tk.Label(ventana_progreso, text=f"Cargando {os.path.basename(archivo_pdf)} a Word...")
    nombre_archivo_label.pack()

    hilo_conversion_word = threading.Thread(
        target=convertir_archivo_a_word_en_segundo_plano,
        args=(archivo_pdf, barra_progreso, porcentaje_label, ventana_progreso, nombre_archivo_label)
    )
    hilo_conversion_word.start()

def convertir_archivo_a_pdf_en_segundo_plano(archivo_word, barra_progreso, porcentaje_label, ventana_progreso, nombre_archivo_label):
    try:
        if not os.path.exists(archivo_word):
            messagebox.showerror("Error", "El archivo no existe.")
            return

        word = win32.Dispatch('Word.Application')
        word.visible = False
        archivo_word_absoluto = os.path.abspath(archivo_word)
        doc = word.Documents.Open(archivo_word_absoluto)

        # Obtener el nombre del archivo original sin la extensión
        nombre_sugerido = os.path.splitext(os.path.basename(archivo_word))[0]

        # Usar asksaveasfilename para obtener la ruta sin codificar espacios
        archivo_pdf_absoluto = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            initialfile=nombre_sugerido
        )

        # Si el usuario cancela, no continuar
        if not archivo_pdf_absoluto:
            messagebox.showerror("Cancelado", "Operación cancelada por el usuario.")
            return

        # Asegurarse de que la ruta no tenga caracteres especiales no válidos para el sistema operativo
        if not archivo_pdf_absoluto.lower().endswith(".pdf"):
            archivo_pdf_absoluto += ".pdf"

        # Guardar el archivo en formato PDF sin codificar los espacios
        archivo_pdf_absoluto = os.path.normpath(archivo_pdf_absoluto)

        # Guardar el archivo en formato PDF
        doc.SaveAs(archivo_pdf_absoluto, FileFormat=17)

        for i in range(1, 101):
            barra_progreso["value"] = i
            porcentaje_label.config(text=f"Cargando... {i}%")
            nombre_archivo_label.config(text=f"Convirtiendo {os.path.basename(archivo_word)} a PDF... ({i}%)")
            ventana_progreso.update_idletasks()
            time.sleep(0.05)

        doc.Close()
        word.Quit()

        messagebox.showinfo("Éxito", f"Archivo convertido a PDF: {archivo_pdf_absoluto}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")
    finally:
        entry_archivo.delete(0, tk.END)
        boton_convertir_pdf.config(state=tk.NORMAL)
        ventana_progreso.destroy()

def convertir_archivo_a_word_en_segundo_plano(archivo_pdf, barra_progreso, porcentaje_label, ventana_progreso, nombre_archivo_label):
    try:
        if not os.path.exists(archivo_pdf):
            messagebox.showerror("Error", "El archivo no existe.")
            return

        word = win32.Dispatch('Word.Application')
        word.visible = False
        archivo_pdf_absoluto = os.path.abspath(archivo_pdf)

        # Obtener el nombre del archivo original sin la extensión
        nombre_sugerido = os.path.splitext(os.path.basename(archivo_pdf))[0]

        # Usar asksaveasfilename para obtener la ruta sin codificar espacios
        archivo_word_absoluto = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Archivos Word", "*.docx")],
            initialfile=nombre_sugerido
        )

        # Si el usuario cancela, no continuar
        if not archivo_word_absoluto:
            messagebox.showerror("Cancelado", "Operación cancelada por el usuario.")
            return

        # Asegurarse de que la ruta no tenga caracteres especiales no válidos para el sistema operativo
        if not archivo_word_absoluto.lower().endswith(".docx"):
            archivo_word_absoluto += ".docx"

        doc = word.Documents.Open(archivo_pdf_absoluto)
        doc.SaveAs(archivo_word_absoluto, FileFormat=16)

        for i in range(1, 101):
            barra_progreso["value"] = i
            porcentaje_label.config(text=f"Cargando... {i}%")
            nombre_archivo_label.config(text=f"Convirtiendo {os.path.basename(archivo_pdf)} a Word... ({i}%)")
            ventana_progreso.update_idletasks()
            time.sleep(0.05)

        doc.Close()
        word.Quit()

        messagebox.showinfo("Éxito", f"Archivo convertido a Word: {archivo_word_absoluto}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")
    finally:
        entry_archivo.delete(0, tk.END)
        boton_convertir_word.config(state=tk.NORMAL)
        ventana_progreso.destroy()

# Configuración de la interfaz gráfica con Tkinter
ventana = tk.Tk()
ventana.title("Conversor de Word a PDF / PDF a Word")

label = tk.Label(ventana, text="Selecciona el archivo para convertir:")
label.pack(pady=10)

entry_archivo = tk.Entry(ventana, width=50)
entry_archivo.pack(pady=10)

boton_seleccionar_word = tk.Button(ventana, text="Seleccionar archivo Word", command=seleccionar_archivo_word)
boton_seleccionar_word.pack(pady=5)

boton_seleccionar_pdf = tk.Button(ventana, text="Seleccionar archivo PDF", command=seleccionar_archivo_pdf)
boton_seleccionar_pdf.pack(pady=5)

boton_convertir_pdf = tk.Button(ventana, text="Convertir a PDF", command=convertir_a_pdf)
boton_convertir_pdf.pack(pady=5)

boton_convertir_word = tk.Button(ventana, text="Convertir a Word", command=convertir_a_word)
boton_convertir_word.pack(pady=5)

ventana.mainloop()
