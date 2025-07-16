import os
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import win32com.client


def procesar_pptx(path):
    powerpoint = None
    try:
        original_path = Path(path)

        if not original_path.exists():
            raise FileNotFoundError(f"Archivo no encontrado:\n{original_path}")

        carpeta_modificado = original_path.parent / "modificado"
        carpeta_modificado.mkdir(exist_ok=True)

        nuevo_nombre = original_path.stem + "_modificado.pptx"
        nuevo_path = str(carpeta_modificado / nuevo_nombre)

        print(f"Intentando abrir archivo:\n{str(original_path)}")

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        presentation = powerpoint.Presentations.Open(str(original_path), WithWindow=False)

        # 👉 Recorremos cada diapositiva y cada forma
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    if shape.TextFrame.HasText:
                        # Insertar "0" blanco al final del texto
                        text_range = shape.TextFrame.TextRange
                        text_range.Text += "0"

                        # Hacer el "0" blanco o transparente (último carácter)
                        last_char = text_range.Characters(text_range.Length, 1)
                        last_char.Font.Color.RGB = 16777215  # Blanco (RGB = 255,255,255)

        # Guardar y cerrar original
        presentation.SaveAs(nuevo_path)
        presentation.Close()

        # Abrir el nuevo archivo modificado
        powerpoint.Presentations.Open(nuevo_path, WithWindow=True)

        return nuevo_path

    except Exception:
        traceback.print_exc()
        raise

    finally:
        if powerpoint:
            powerpoint.Visible = True


def seleccionar():
    path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx;*.pptm")])
    if not path:
        return
    try:
        salida = procesar_pptx(path)
        messagebox.showinfo("Listo", f"Presentación guardada y abierta:\n{salida}")
    except Exception as e:
        messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Agregar ceros invisibles a PowerPoint")
    root.geometry("420x160")
    tk.Button(root, text="Seleccionar archivo PowerPoint", command=seleccionar, font=("Arial", 12)).pack(pady=40)
    root.mainloop()
