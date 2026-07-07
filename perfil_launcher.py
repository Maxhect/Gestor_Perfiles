"""
Asistente de Perfiles - Lanzador de aplicaciones y archivos
=============================================================
Permite crear perfiles (ej: "Trabajo", "Diseño", "Streaming") que agrupan
programas y archivos. Con un clic, se abren todos los elementos de un perfil.

Requisitos: Python 3.8+ (ya incluye tkinter en Windows).
Ejecutar con:  python perfil_launcher.py
(o hacer doble clic si asocias .py con pythonw.exe para que no muestre consola)

Los datos se guardan en:  %APPDATA%\\PerfilLauncher\\perfiles.json
"""

import json
import os
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# ---------------------------------------------------------------------------
# Almacenamiento
# ---------------------------------------------------------------------------

APPDATA = os.environ.get("APPDATA", os.path.expanduser("~"))
CONFIG_DIR = os.path.join(APPDATA, "PerfilLauncher")
CONFIG_FILE = os.path.join(CONFIG_DIR, "perfiles.json")


def cargar_perfiles():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def guardar_perfiles(perfiles):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(perfiles, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Lanzamiento de elementos
# ---------------------------------------------------------------------------

def abrir_item(ruta, argumentos=""):
    """Abre un archivo o programa usando la asociación predeterminada de Windows."""
    try:
        if argumentos.strip():
            subprocess.Popen(f'"{ruta}" {argumentos}', shell=True)
        else:
            os.startfile(ruta)  # noqa: usa la app/asociación predeterminada de Windows
    except Exception as e:
        messagebox.showerror("Error al abrir", f"No se pudo abrir:\n{ruta}\n\n{e}")


# ---------------------------------------------------------------------------
# Interfaz gráfica
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Asistente de Perfiles")
        self.geometry("780x480")
        self.minsize(650, 400)

        self.perfiles = cargar_perfiles()
        self.perfil_actual = None

        self._construir_ui()
        self._refrescar_lista_perfiles()

    # ---------------- UI ----------------
    def _construir_ui(self):
        # Panel izquierdo: perfiles
        izq = ttk.Frame(self, padding=10)
        izq.pack(side="left", fill="y")

        ttk.Label(izq, text="Perfiles", font=("Segoe UI", 12, "bold")).pack(anchor="w")

        self.lista_perfiles = tk.Listbox(izq, width=24, height=20, exportselection=False)
        self.lista_perfiles.pack(fill="y", expand=True, pady=(5, 5))
        self.lista_perfiles.bind("<<ListboxSelect>>", self._on_seleccionar_perfil)

        btns_perfil = ttk.Frame(izq)
        btns_perfil.pack(fill="x", pady=(5, 0))
        ttk.Button(btns_perfil, text="+ Nuevo", command=self._nuevo_perfil).pack(side="left", expand=True, fill="x")
        ttk.Button(btns_perfil, text="Renombrar", command=self._renombrar_perfil).pack(side="left", expand=True, fill="x")
        ttk.Button(btns_perfil, text="Eliminar", command=self._eliminar_perfil).pack(side="left", expand=True, fill="x")

        # Panel derecho: contenido del perfil
        der = ttk.Frame(self, padding=10)
        der.pack(side="left", fill="both", expand=True)

        cabecera = ttk.Frame(der)
        cabecera.pack(fill="x")
        self.lbl_perfil = ttk.Label(cabecera, text="Selecciona o crea un perfil", font=("Segoe UI", 13, "bold"))
        self.lbl_perfil.pack(side="left")

        self.btn_lanzar = ttk.Button(cabecera, text="▶ Iniciar perfil", command=self._lanzar_perfil, state="disabled")
        self.btn_lanzar.pack(side="right")

        ttk.Label(der, text="Programas y archivos incluidos:").pack(anchor="w", pady=(15, 3))

        cols = ("nombre", "ruta", "args")
        self.tabla = ttk.Treeview(der, columns=cols, show="headings", height=14)
        self.tabla.heading("nombre", text="Nombre")
        self.tabla.heading("ruta", text="Ruta")
        self.tabla.heading("args", text="Argumentos (opcional)")
        self.tabla.column("nombre", width=140)
        self.tabla.column("ruta", width=380)
        self.tabla.column("args", width=140)
        self.tabla.pack(fill="both", expand=True)

        btns_items = ttk.Frame(der)
        btns_items.pack(fill="x", pady=(8, 0))
        ttk.Button(btns_items, text="+ Agregar archivo/programa", command=self._agregar_item).pack(side="left")
        ttk.Button(btns_items, text="Quitar seleccionado", command=self._quitar_item).pack(side="left", padx=6)

    # ---------------- Lógica de perfiles ----------------
    def _refrescar_lista_perfiles(self):
        self.lista_perfiles.delete(0, tk.END)
        for nombre in sorted(self.perfiles.keys()):
            self.lista_perfiles.insert(tk.END, nombre)

    def _on_seleccionar_perfil(self, event=None):
        sel = self.lista_perfiles.curselection()
        if not sel:
            return
        nombre = self.lista_perfiles.get(sel[0])
        self.perfil_actual = nombre
        self.lbl_perfil.config(text=nombre)
        self.btn_lanzar.config(state="normal")
        self._refrescar_tabla()

    def _nuevo_perfil(self):
        nombre = simpledialog.askstring("Nuevo perfil", "Nombre del perfil (ej: Trabajo):", parent=self)
        if not nombre:
            return
        nombre = nombre.strip()
        if not nombre:
            return
        if nombre in self.perfiles:
            messagebox.showwarning("Ya existe", "Ya existe un perfil con ese nombre.")
            return
        self.perfiles[nombre] = []
        guardar_perfiles(self.perfiles)
        self._refrescar_lista_perfiles()

    def _renombrar_perfil(self):
        if not self.perfil_actual:
            return
        nuevo = simpledialog.askstring("Renombrar perfil", "Nuevo nombre:", initialvalue=self.perfil_actual, parent=self)
        if not nuevo or nuevo == self.perfil_actual:
            return
        if nuevo in self.perfiles:
            messagebox.showwarning("Ya existe", "Ya existe un perfil con ese nombre.")
            return
        self.perfiles[nuevo] = self.perfiles.pop(self.perfil_actual)
        self.perfil_actual = nuevo
        guardar_perfiles(self.perfiles)
        self._refrescar_lista_perfiles()
        self.lbl_perfil.config(text=nuevo)

    def _eliminar_perfil(self):
        if not self.perfil_actual:
            return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar el perfil '{self.perfil_actual}'?"):
            return
        del self.perfiles[self.perfil_actual]
        guardar_perfiles(self.perfiles)
        self.perfil_actual = None
        self.lbl_perfil.config(text="Selecciona o crea un perfil")
        self.btn_lanzar.config(state="disabled")
        self._refrescar_lista_perfiles()
        self._refrescar_tabla()

    # ---------------- Lógica de items ----------------
    def _refrescar_tabla(self):
        self.tabla.delete(*self.tabla.get_children())
        if not self.perfil_actual:
            return
        for item in self.perfiles.get(self.perfil_actual, []):
            self.tabla.insert("", tk.END, values=(item["nombre"], item["ruta"], item.get("args", "")))

    def _agregar_item(self):
        if not self.perfil_actual:
            messagebox.showinfo("Selecciona un perfil", "Primero selecciona o crea un perfil.")
            return
        ruta = filedialog.askopenfilename(title="Selecciona un programa o archivo")
        if not ruta:
            return
        nombre_sugerido = os.path.splitext(os.path.basename(ruta))[0]
        nombre = simpledialog.askstring("Nombre para mostrar", "Nombre para este elemento:", initialvalue=nombre_sugerido, parent=self)
        if not nombre:
            nombre = nombre_sugerido
        args = simpledialog.askstring("Argumentos (opcional)", "Argumentos de línea de comandos (dejar vacío si no aplica):", parent=self)
        self.perfiles[self.perfil_actual].append({"nombre": nombre, "ruta": ruta, "args": args or ""})
        guardar_perfiles(self.perfiles)
        self._refrescar_tabla()

    def _quitar_item(self):
        sel = self.tabla.selection()
        if not sel or not self.perfil_actual:
            return
        idx = self.tabla.index(sel[0])
        del self.perfiles[self.perfil_actual][idx]
        guardar_perfiles(self.perfiles)
        self._refrescar_tabla()

    # ---------------- Lanzar perfil ----------------
    def _lanzar_perfil(self):
        if not self.perfil_actual:
            return
        items = self.perfiles.get(self.perfil_actual, [])
        if not items:
            messagebox.showinfo("Perfil vacío", "Este perfil no tiene archivos ni programas agregados.")
            return
        for item in items:
            abrir_item(item["ruta"], item.get("args", ""))


if __name__ == "__main__":
    if sys.platform != "win32":
        print("Este script está pensado para Windows (usa os.startfile).")
    app = App()
    app.mainloop()
