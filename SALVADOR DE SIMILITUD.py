import pathlib
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client as win32

def procesar_con_com(doc_path, paso=70):
    wdColorWhite = 16777215
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        ruta = str(pathlib.Path(doc_path).resolve())
        doc = word.Documents.Open(ruta)

        for para in doc.Paragraphs:
            # Verificar si el párrafo está en la primera página
            page_number = para.Range.Information(3)  # wdActiveEndPageNumber = 3
            if page_number == 1:
                continue  # Saltar si está en la primera página

            texto = para.Range.Text.rstrip('\r\x07')
            L = len(texto)
            if L < paso:
                continue

            objetivos = list(range(paso, L, paso))
            for obj in objetivos:
                izq = texto.rfind(" ", 0, obj + 1)
                der = texto.find(" ", obj)
                if izq == -1 and der == -1:
                    continue
                idx = izq if (der == -1 or (obj - izq) <= (der - obj)) else der

                start = para.Range.Start + idx
                rng = doc.Range(start, start + 1)
                if rng.Text == " ":
                    rng.Text = "0"
                    rng.Font.Color = wdColorWhite

        # Crear carpeta MODIFICADO
        original_path = pathlib.Path(doc_path)
        carpeta_modificado = original_path.parent / "MODIFICADO"
        carpeta_modificado.mkdir(exist_ok=True)

        nuevo_nombre = original_path.stem + "_modificado.docx"
        nuevo_path = str(carpeta_modificado / nuevo_nombre)

        doc.SaveAs2(nuevo_path, FileFormat=16)
        doc.Close(False)

        # Abrir automáticamente el documento modificado
        word.Visible = True
        word.Documents.Open(nuevo_path)

        return nuevo_path

    except Exception:
        traceback.print_exc()
        raise

    finally:
        if word:
            word.Visible = True

def seleccionar():
    path = filedialog.askopenfilename(filetypes=[("Word", "*.docx;*.docm")])
    if not path:
        return
    try:
        salida = procesar_con_com(path, paso=70)
        messagebox.showinfo("Listo", f"Documento guardado y abierto:\n{salida}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Insertar 0 blanco desde página 2")
    root.geometry("420x160")
    tk.Button(root, text="Seleccionar archivo Word", command=seleccionar, font=("Arial", 12)).pack(pady=40)
    root.mainloop()
