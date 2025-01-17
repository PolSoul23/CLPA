import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32
import os

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

    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False  # Deshabilita alertas
        archivo_word_absoluto = os.path.abspath(archivo_word)
        doc = word.Documents.Open(archivo_word_absoluto, ReadOnly=True)

        archivo_pdf_absoluto = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf")],
            initialfile=os.path.splitext(os.path.basename(archivo_word))[0]
        )

        if not archivo_pdf_absoluto:
            messagebox.showerror("Cancelado", "Operación cancelada por el usuario.")
            return

        doc.SaveAs(archivo_pdf_absoluto, FileFormat=17)
        doc.Close()
        word.Quit()

        messagebox.showinfo("Éxito", f"Archivo convertido a PDF: {archivo_pdf_absoluto}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")

def convertir_a_word():
    archivo_pdf = entry_archivo.get()
    if archivo_pdf == "":
        messagebox.showerror("Error", "Por favor, selecciona un archivo PDF.")
        return

    try:
        word = win32.Dispatch('Word.Application')
        word.Visible = False
        archivo_pdf_absoluto = os.path.abspath(archivo_pdf)
        doc = word.Documents.Open(archivo_pdf_absoluto)

        archivo_word_absoluto = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Archivos Word", "*.docx")],
            initialfile=os.path.splitext(os.path.basename(archivo_pdf))[0]
        )

        if not archivo_word_absoluto:
            messagebox.showerror("Cancelado", "Operación cancelada por el usuario.")
            return

        doc.SaveAs(archivo_word_absoluto, FileFormat=16)
        doc.Close()
        word.Quit()

        messagebox.showinfo("Éxito", f"Archivo convertido a Word: {archivo_word_absoluto}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al convertir el archivo: {e}")

# Configuración de la interfaz gráfica con Tkinter
ventana = tk.Tk()
ventana.title("Conversor de Word a PDF / PDF a Word")

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
